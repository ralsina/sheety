require "termisu"

module Sheety
  # TUI spreadsheet viewer using Termisu
  class TUI
    @termisu : Termisu
    @sheets : Array(String)
    @sheet_data : Hash(String, Array(NamedTuple(cell: String, formula: String, value: String)))
    @current_sheet_idx : Int32
    @active_row : Int32
    @active_col : Int32
    @row_offset : Int32
    @col_offset : Int32
    @grid_width : Int32
    @grid_height : Int32
    @cell_width : Int32
    @row_num_width : Int32
    @header_height : Int32
    @status_height : Int32

    # Grid dimensions
    @max_row : Int32 = 0
    @max_col : Int32 = 0
    @grid : Array(Array(String)) = [] of Array(String)

    # Color scheme
    @fg_default : Termisu::Color
    @bg_default : Termisu::Color
    @fg_header : Termisu::Color
    @bg_header : Termisu::Color
    @fg_active : Termisu::Color
    @bg_active : Termisu::Color
    @fg_status : Termisu::Color
    @bg_status : Termisu::Color
    @fg_formula : Termisu::Color
    @fg_value : Termisu::Color

    # Edit mode
    @edit_mode : Bool = false
    @edit_buffer : String = ""
    @formula_bar_height : Int32 = 1

    # Callback for updating cell values
    @value_update_callback : Proc(String, String, String, Nil)?
    @refresh_callback : Proc(Nil)?
    @value_getter_callback : Proc(String, String, String)?

    def initialize(sheets : Array(String), sheet_data : Hash(String, Array(NamedTuple(cell: String, formula: String, value: String))), update_callback : Proc(String, String, String, Nil) = ->(_s : String, _c : String, _v : String) {})
      @termisu = Termisu.new
      @sheets = sheets.sort
      @sheet_data = sheet_data
      @current_sheet_idx = 0
      @active_row = 0
      @active_col = 0
      @row_offset = 0
      @col_offset = 0

      # Layout constants
      @cell_width = 12
      @row_num_width = 5
      @header_height = 1
      @status_height = 2
      @formula_bar_height = 1

      # Get terminal size
      term_size = @termisu.size
      @grid_width = term_size[0]
      @grid_height = term_size[1] - @header_height - @status_height - @formula_bar_height

      # Colors (Lotus 1-2-3 inspired)
      @fg_default = Termisu::Color.white
      @bg_default = Termisu::Color.black
      @fg_header = Termisu::Color.cyan
      @bg_header = Termisu::Color.black
      @fg_active = Termisu::Color.white
      @bg_active = Termisu::Color.blue
      @fg_status = Termisu::Color.cyan
      @bg_status = Termisu::Color.black
      @fg_formula = Termisu::Color.green
      @fg_value = Termisu::Color.white

      # Initialize callbacks
      @value_update_callback = update_callback
      @refresh_callback = nil
      @value_getter_callback = ->(_sheet : String, _cell : String) { "" }

      # Initialize grid for current sheet
      initialize_grid
    end

    def self.new(sheets : Array(String), sheet_data : Hash(String, Array(NamedTuple(cell: String, formula: String, value: String))), &block : Proc(String, String, String, Nil))
      new(sheets, sheet_data, block)
    end

    def run : Nil
      @termisu.hide_cursor
      render

      loop do
        if event = @termisu.poll_event(100)
          case event
          when Termisu::Event::Key
            handle_key_event(event)
            render if should_render?(event)
          when Termisu::Event::Resize
            handle_resize(event)
            render
          end
        end
      end
    ensure
      @termisu.show_cursor
      @termisu.close
    end

    private def initialize_grid : Nil
      sheet_name = @sheets[@current_sheet_idx]
      data = @sheet_data[sheet_name]?

      if data.nil? || data.empty?
        @max_row = 0
        @max_col = 0
        @grid = [] of Array(String)
        return
      end

      # Find grid dimensions
      @max_row = 0
      @max_col = 0

      data.each do |cell|
        if match = cell[:cell].match(/^([A-Z]+)(\d+)$/)
          col = match[1]
          row = match[2].to_i

          # Convert column to number
          col_num = col_to_num(col)

          @max_col = col_num if col_num > @max_col
          @max_row = row if row > @max_row
        end
      end

      # Create grid
      @grid = Array.new(@max_row) { Array.new(@max_col, "") }

      # Fill grid with values
      data.each do |cell|
        if match = cell[:cell].match(/^([A-Z]+)(\d+)$/)
          col = match[1]
          row = match[2].to_i - 1 # Convert to 0-indexed
          col_num = col_to_num(col) - 1

          # Get value: try callback first, fall back to stored value
          value = if callback = @value_getter_callback
                    fetched = callback.call(sheet_name, cell[:cell])
                    fetched.empty? ? cell[:value] : fetched
                  else
                    cell[:value]
                  end

          # Show calculated result for formulas, or just value
          display = value

          @grid[row][col_num] = display if row < @max_row && col_num < @max_col
        end
      end
    end

    private def col_to_num(col : String) : Int32
      result = 0
      col.each_char do |char|
        result = result * 26 + (char.ord - 'A'.ord + 1)
      end
      result
    end

    private def num_to_col(num : Int32) : String
      result = ""
      n = num
      while n > 0
        n -= 1
        result = ('A' + (n % 26)).to_s + result
        n //= 26
      end
      result
    end

    private def handle_key_event(event : Termisu::Event::Key) : Nil
      if @edit_mode
        handle_edit_key_event(event)
      else
        handle_normal_key_event(event)
      end
    end

    private def handle_normal_key_event(event : Termisu::Event::Key) : Nil
      case event.key
      when .q?
        @termisu.close
        exit 0
      when .escape?
        @termisu.close
        exit 0
      when .enter?
        # Enter edit mode
        enter_edit_mode
      when .up?
        move_active(-1, 0)
      when .down?
        move_active(1, 0)
      when .left?
        move_active(0, -1)
      when .right?
        move_active(0, 1)
      when .home?
        @active_col = 0
        @col_offset = 0
      when .end?
        @active_col = @max_col - 1
        adjust_col_offset
      when .page_up?
        @active_row = {@active_row - @grid_height, 0}.max
        adjust_row_offset
      when .page_down?
        @active_row = {@active_row + @grid_height, @max_row - 1}.min
        adjust_row_offset
      when .tab?
        # Next sheet (Ctrl+Tab would be better, but Tab works for now)
        switch_sheet(1)
      when .back_tab?
        # Previous sheet
        switch_sheet(-1)
      end
    end

    private def handle_edit_key_event(event : Termisu::Event::Key) : Nil
      case event.key
      when .escape?
        # Cancel edit mode
        @edit_mode = false
        @edit_buffer = ""
      when .enter?
        # Save and exit edit mode
        save_edit_value
        @edit_mode = false
        @edit_buffer = ""
      when .backspace?
        # Delete last character
        @edit_buffer = @edit_buffer[0...-1]
      else
        # Check if this is a printable character
        if char = event.char
          # Don't include control characters
          if char.printable?
            @edit_buffer += char
          end
        end
      end
    end

    private def move_active(d_row : Int32, d_col : Int32) : Nil
      @active_row = ({@active_row + d_row, 0}.max).clamp(0, @max_row - 1) if @max_row > 0
      @active_col = ({@active_col + d_col, 0}.max).clamp(0, @max_col - 1) if @max_col > 0

      adjust_row_offset
      adjust_col_offset
    end

    private def adjust_row_offset : Nil
      return if @max_row == 0

      # Keep active cell visible
      if @active_row < @row_offset
        @row_offset = @active_row
      elsif @active_row >= @row_offset + @grid_height
        @row_offset = @active_row - @grid_height + 1
      end
    end

    private def adjust_col_offset : Nil
      return if @max_col == 0

      # Calculate how many columns fit
      cols_per_screen = (@grid_width - @row_num_width) // (@cell_width + 1)

      # Keep active cell visible
      if @active_col < @col_offset
        @col_offset = @active_col
      elsif @active_col >= @col_offset + cols_per_screen
        @col_offset = @active_col - cols_per_screen + 1
      end
    end

    private def switch_sheet(direction : Int32) : Nil
      @current_sheet_idx = (@current_sheet_idx + direction) % @sheets.size
      @current_sheet_idx = (@current_sheet_idx + @sheets.size) % @sheets.size if @current_sheet_idx < 0

      # Reset position
      @active_row = 0
      @active_col = 0
      @row_offset = 0
      @col_offset = 0

      # Reinitialize grid for new sheet
      initialize_grid
    end

    private def current_cell_ref : String
      num_to_col(@active_col + 1) + (@active_row + 1).to_s
    end

    private def current_cell_formula : String
      sheet_name = @sheets[@current_sheet_idx]
      data = @sheet_data[sheet_name]?
      return "" if data.nil?

      cell_ref = current_cell_ref
      data.each do |cell|
        if cell[:cell] == cell_ref
          return cell[:formula]
        end
      end
      ""
    end

    private def current_cell_value : String
      return @grid[@active_row][@active_col] if @active_row < @max_row && @active_col < @max_col
      ""
    end

    private def formula_cell? : Bool
      !current_cell_formula.empty?
    end

    private def enter_edit_mode : Nil
      # Check if current cell is a formula cell (locked)
      if formula_cell?
        # Formula cells are locked, don't enter edit mode
        return
      end

      @edit_mode = true
      # Initialize edit buffer with current value
      @edit_buffer = current_cell_value
    end

    private def save_edit_value : Nil
      return if formula_cell?

      sheet_name = @sheets[@current_sheet_idx]
      cell_ref = current_cell_ref

      # Call the update callback if provided
      if callback = @value_update_callback
        callback.call(sheet_name, cell_ref, @edit_buffer)
      end

      # Call the refresh callback if provided to update the grid from store
      if refresh_cb = @refresh_callback
        refresh_cb.call
      else
        # Fallback: just update the local grid
        if @active_row < @max_row && @active_col < @max_col
          @grid[@active_row][@active_col] = @edit_buffer
        end
      end
    end

    def set_refresh_callback(&callback : Proc(Nil))
      @refresh_callback = callback
    end

    def set_value_getter(&callback : Proc(String, String, String))
      @value_getter_callback = callback
    end

    def refresh_current_sheet : Nil
      initialize_grid
    end

    private def handle_resize(event : Termisu::Event::Resize) : Nil
      @grid_width = event.width
      @grid_height = event.height - @header_height - @status_height - @formula_bar_height
    end

    private def should_render?(event : Termisu::Event::Key) : Bool
      # Render on all key events for now
      true
    end

    private def render : Nil
      @termisu.clear
      render_header
      render_grid
      render_formula_bar
      render_status
      @termisu.render
    end

    private def render_header : Nil
      sheet_name = @sheets[@current_sheet_idx]
      header_text = "Sheet: #{sheet_name}"

      # Center the sheet name
      x = (@grid_width - header_text.size) // 2

      header_text.each_char_with_index do |char, i|
        @termisu.set_cell(x + i, 0, char, fg: @fg_header, bg: @bg_header, attr: Termisu::Attribute::Bold)
      end

      # Fill rest of header line
      (0...@grid_width).each do |i|
        if i < x || i >= x + header_text.size
          @termisu.set_cell(i, 0, ' ', fg: @fg_header, bg: @bg_header)
        end
      end
    end

    private def render_grid : Nil
      return if @max_row == 0 || @max_col == 0

      # Calculate visible columns
      cols_per_screen = (@grid_width - @row_num_width) // (@cell_width + 1)
      end_col = {@col_offset + cols_per_screen, @max_col}.min

      # Render column headers (A, B, C, ...)
      col_x = @row_num_width
      (@col_offset...end_col).each do |col_idx|
        col_label = num_to_col(col_idx + 1)
        draw_text_centered(col_x, @header_height, @cell_width, col_label, @fg_header, @bg_header, Termisu::Attribute::Bold)
        col_x += @cell_width + 1
      end

      # Render grid cells
      (0...@grid_height).each do |screen_row|
        grid_row = @row_offset + screen_row
        break if grid_row >= @max_row

        # Row number
        row_label = (grid_row + 1).to_s
        row_x = 0
        row_label.each_char_with_index do |char, i|
          @termisu.set_cell(row_x + i, @header_height + 1 + screen_row, char, fg: @fg_header, bg: @bg_header, attr: Termisu::Attribute::Bold)
        end

        # Fill rest of row number area
        (row_label.size...@row_num_width).each do |i|
          @termisu.set_cell(i, @header_height + 1 + screen_row, ' ', fg: @fg_header, bg: @bg_header)
        end

        # Cells
        col_x = @row_num_width
        (@col_offset...end_col).each do |col_idx|
          break if col_idx >= @max_col

          cell_value = @grid[grid_row][col_idx]
          is_active = grid_row == @active_row && col_idx == @active_col

          # Truncate value to fit
          display_value = truncate_value(cell_value, @cell_width)

          # Determine color
          fg = @fg_default
          bg = @bg_default
          attr = Termisu::Attribute::None

          if is_active
            fg = @fg_active
            bg = @bg_active
            attr = Termisu::Attribute::Bold
          end

          # Draw cell
          (0...@cell_width).each do |i|
            char = i < display_value.size ? display_value[i] : ' '
            @termisu.set_cell(col_x + i, @header_height + 1 + screen_row, char, fg: fg, bg: bg, attr: attr)
          end

          # Separator
          if col_idx < end_col - 1
            @termisu.set_cell(col_x + @cell_width, @header_height + 1 + screen_row, ' ', fg: @fg_default, bg: @bg_default)
          end

          col_x += @cell_width + 1
        end
      end
    end

    private def render_formula_bar : Nil
      formula_y = @header_height + 1 + @grid_height

      # Draw formula bar background
      (0...@grid_width).each do |i|
        @termisu.set_cell(i, formula_y, ' ', fg: @fg_default, bg: @bg_default)
      end

      # Get current cell info
      cell_ref = current_cell_ref
      formula = current_cell_formula

      if @edit_mode
        # Edit mode: show the edit buffer
        label = "Editing #{cell_ref}: "
        label.each_char_with_index do |char, i|
          break if i >= @grid_width
          @termisu.set_cell(i, formula_y, char, fg: @fg_header, bg: @bg_default, attr: Termisu::Attribute::Bold)
        end

        # Show edit buffer
        display_buffer = truncate_value(@edit_buffer, @grid_width - label.size)
        display_buffer.each_char_with_index do |char, i|
          x = label.size + i
          break if x >= @grid_width
          @termisu.set_cell(x, formula_y, char, fg: @fg_active, bg: @bg_default, attr: Termisu::Attribute::Bold)
        end
      else
        # Normal mode: show formula or value
        if formula.empty?
          # Value cell - show the value
          label = "#{cell_ref}: "
          value = current_cell_value

          label.each_char_with_index do |char, i|
            break if i >= @grid_width
            @termisu.set_cell(i, formula_y, char, fg: @fg_header, bg: @bg_default, attr: Termisu::Attribute::Bold)
          end

          display_value = truncate_value(value, @grid_width - label.size)
          display_value.each_char_with_index do |char, i|
            x = label.size + i
            break if x >= @grid_width
            @termisu.set_cell(x, formula_y, char, fg: @fg_value, bg: @bg_default)
          end
        else
          # Formula cell - show formula and value
          label = "#{cell_ref} [Formula]: "
          formula_display = truncate_value(formula, @grid_width - label.size)

          label.each_char_with_index do |char, i|
            break if i >= @grid_width
            @termisu.set_cell(i, formula_y, char, fg: @fg_formula, bg: @bg_default, attr: Termisu::Attribute::Bold)
          end

          formula_display.each_char_with_index do |char, i|
            x = label.size + i
            break if x >= @grid_width
            @termisu.set_cell(x, formula_y, char, fg: @fg_formula, bg: @bg_default)
          end
        end
      end
    end

    private def render_status : Nil
      status_y = @header_height + 1 + @grid_height + @formula_bar_height

      sheet_name = @sheets[@current_sheet_idx]
      cell_ref = num_to_col(@active_col + 1) + (@active_row + 1).to_s

      if @max_row == 0 || @max_col == 0
        status_text = "Empty sheet"
      else
        status_text = "#{sheet_name}!#{cell_ref}"
      end

      # Draw status bar background
      (0...@grid_width).each do |i|
        @termisu.set_cell(i, status_y, ' ', fg: @fg_status, bg: @bg_status)
      end

      # Draw status text
      status_text.each_char_with_index do |char, i|
        break if i >= @grid_width
        @termisu.set_cell(i, status_y, char, fg: @fg_status, bg: @bg_status, attr: Termisu::Attribute::Bold)
      end

      # Draw help hints
      help_text = @edit_mode ? "ENTER:Save | ESC:Cancel" : "Arrows:Move | ENTER:Edit | Tab:Sheet | Q:Quit"
      help_x = {@grid_width - help_text.size, 0}.max

      help_text.each_char_with_index do |char, i|
        if help_x + i < @grid_width
          @termisu.set_cell(help_x + i, status_y, char, fg: @fg_status, bg: @bg_status, attr: Termisu::Attribute::Dim)
        end
      end
    end

    private def draw_text_centered(x : Int32, y : Int32, width : Int32, text : String, fg : Termisu::Color, bg : Termisu::Color, attr : Termisu::Attribute = Termisu::Attribute::None) : Nil
      truncated = truncate_value(text, width)
      padding = (width - truncated.size) // 2

      (0...width).each do |i|
        char = ' '
        if i >= padding && i < padding + truncated.size
          char = truncated[i - padding]
        end
        @termisu.set_cell(x + i, y, char, fg: fg, bg: bg, attr: attr)
      end
    end

    private def truncate_value(value : String, max_width : Int32) : String
      return "" if max_width <= 0
      return value if value.size <= max_width
      value[0...max_width]
    end
  end
end
