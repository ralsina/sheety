require "termisu"
require "yaml"

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

    # Grid dimensions (fixed size)
    @max_row : Int32 = 1000
    @max_col : Int32 = 1000
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
    @edit_cursor : Int32 = 0
    @formula_bar_height : Int32 = 1
    @notification_message : String = ""
    @notification_timeout : Int32 = 0

    # Callback for updating cell values
    @value_update_callback : Proc(String, String, String, Nil)?
    @refresh_callback : Proc(Nil)?
    @value_getter_callback : Proc(String, String, String)?
    @save_callback : Proc(Nil)?
    @source_file : String?

    def initialize(sheets : Array(String), sheet_data : Hash(String, Array(NamedTuple(cell: String, formula: String, value: String))), update_callback : Proc(String, String, String, Nil) = ->(_s : String, _c : String, _v : String) {}, source_file : String? = nil)
      @termisu = Termisu.new
      @sheets = sheets.sort
      @sheet_data = sheet_data
      @source_file = source_file
      @current_sheet_idx = 0
      @active_row = 0
      @active_col = 0
      @row_offset = 0
      @col_offset = 0

      # Layout constants
      @cell_width = 12
      @row_num_width = 4
      @spacer_width = 1
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

    def set_initial_position(sheet_name : String, cell_ref : String) : Nil
      # Find the sheet index
      if sheet_idx = @sheets.index(sheet_name)
        @current_sheet_idx = sheet_idx

        # Parse the cell reference to get row and column
        if match = cell_ref.match(/^([A-Z]+)(\d+)$/)
          col_str = match[1]
          row_str = match[2]

          # Convert column string to column index
          @active_col = col_to_num(col_str) - 1
          # Convert row string to row index (1-based to 0-based)
          @active_row = row_str.to_i - 1

          # Clamp to valid ranges
          @active_col = @active_col.clamp(0, @max_col - 1)
          @active_row = @active_row.clamp(0, @max_row - 1)

          # Reset offsets
          @row_offset = 0
          @col_offset = 0

          # Reinitialize grid for the correct sheet
          initialize_grid
        end
      end
    end

    def self.new(sheets : Array(String), sheet_data : Hash(String, Array(NamedTuple(cell: String, formula: String, value: String))), &block : Proc(String, String, String, Nil))
      new(sheets, sheet_data, block)
    end

    def run : Nil
      @termisu.hide_cursor
      render

      loop do
        # Decrease notification timeout
        if @notification_timeout > 0
          @notification_timeout -= 1
          if @notification_timeout == 0
            @notification_message = ""
            render
          end
        end

        if event = @termisu.poll_event(100)
          case event
          when Termisu::Event::Key
            handle_key_event(event)
            # Clear notification on any key press
            if @notification_timeout > 0
              clear_notification
            end
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

      # Create fixed 1000x1000 grid filled with empty strings
      @grid = Array.new(@max_row) { Array.new(@max_col, "") }

      # If no data and no callback, we're done
      return if data.nil? || data.empty?

      # Build a map of cells with their formulas for quick lookup
      cell_map = {} of String => NamedTuple(formula: String, value: String)
      data.each do |cell|
        cell_map[cell[:cell]] = {formula: cell[:formula], value: cell[:value]}
      end

      # Fill grid by checking all cells that have either data or might have values in Croupier
      (0...@max_row).each do |row|
        (0...@max_col).each do |col|
          # Convert to cell reference
          cell_ref = num_to_col(col + 1) + (row + 1).to_s

          # Try to get value from callback first (for edited cells), then from data
          value = if callback = @value_getter_callback
                    fetched = callback.call(sheet_name, cell_ref)
                    if fetched.empty?
                      # Not in store, check original data
                      cell_data = cell_map[cell_ref]?
                      cell_data ? cell_data[:value] : ""
                    else
                      fetched
                    end
                  else
                    # No callback, use original data
                    cell_data = cell_map[cell_ref]?
                    cell_data ? cell_data[:value] : ""
                  end

          @grid[row][col] = value unless value.empty?
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
      when .s?
        # Save to YAML
        save_to_yaml
      end
    end

    private def handle_edit_key_event(event : Termisu::Event::Key) : Nil
      case event.key
      when .escape?
        # Cancel edit mode
        @edit_mode = false
        @edit_buffer = ""
        @edit_cursor = 0
      when .enter?
        # Save and exit edit mode
        save_edit_value
        @edit_mode = false
        @edit_buffer = ""
        @edit_cursor = 0
      when .backspace?
        # Delete character before cursor
        if @edit_cursor > 0
          @edit_buffer = @edit_buffer[0...@edit_cursor - 1] + @edit_buffer[@edit_cursor..]
          @edit_cursor -= 1
        end
      when .delete?
        # Delete character at cursor
        if @edit_cursor < @edit_buffer.size
          @edit_buffer = @edit_buffer[0...@edit_cursor] + @edit_buffer[@edit_cursor + 1..]
        end
      when .left?
        # Move cursor left
        @edit_cursor = {@edit_cursor - 1, 0}.max
      when .right?
        # Move cursor right
        @edit_cursor = {@edit_cursor + 1, @edit_buffer.size}.min
      when .home?
        # Move cursor to start
        @edit_cursor = 0
      when .end?
        # Move cursor to end
        @edit_cursor = @edit_buffer.size
      else
        # Check if this is a printable character
        if char = event.char
          # Don't include control characters
          if char.printable?
            @edit_buffer = @edit_buffer[0...@edit_cursor] + char + @edit_buffer[@edit_cursor..]
            @edit_cursor += 1
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
      # Keep active cell visible
      if @active_row < @row_offset
        @row_offset = @active_row
      elsif @active_row >= @row_offset + @grid_height
        @row_offset = @active_row - @grid_height + 1
      end
    end

    private def adjust_col_offset : Nil
      # Calculate how many columns fit
      cols_per_screen = (@grid_width - @row_num_width - @spacer_width) // (@cell_width + 1)

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
      @edit_mode = true

      # For formula cells, edit the formula itself
      # For value cells, edit the current value
      if formula_cell?
        @edit_buffer = current_cell_formula
      else
        @edit_buffer = current_cell_value
      end

      # Set cursor to end of buffer
      @edit_cursor = @edit_buffer.size
    end

    private def show_notification(message : String, duration : Int32) : Nil
      @notification_message = message
      @notification_timeout = duration
    end

    private def clear_notification : Nil
      @notification_message = ""
      @notification_timeout = 0
    end

    private def save_edit_value : Nil
      sheet_name = @sheets[@current_sheet_idx]
      cell_ref = current_cell_ref

      # If editing a formula cell, we need to recompile
      if formula_cell?
        # Validate that it's still a formula (starts with =)
        unless @edit_buffer.starts_with?("=")
          show_notification("Formula must start with =", 60)
          @edit_mode = false
          @edit_buffer = ""
          @edit_cursor = 0
          return
        end

        # Save the new formula to YAML
        update_formula_in_yaml(sheet_name, cell_ref, @edit_buffer)

        # Trigger recompile and restart
        show_notification("Recompiling...", 30)
        @termisu.render

        # Exit and let the wrapper script recompile
        @termisu.close
        exit 42 # Special exit code to signal recompile needed
      else
        # Regular value cell - just update via callback
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
    end

    def set_refresh_callback(&callback : Proc(Nil))
      @refresh_callback = callback
    end

    def set_value_getter(&callback : Proc(String, String, String))
      @value_getter_callback = callback
    end

    def set_save_callback(&callback : Proc(Nil))
      @save_callback = callback
    end

    def set_source_file(source_file : String) : Nil
      @source_file = source_file
    end

    def refresh_current_sheet : Nil
      initialize_grid
    end

    def save_to_yaml : Nil
      source_file = @source_file
      return if source_file.nil? || source_file.empty?

      if callback = @save_callback
        callback.call
      else
        # Default save implementation
        perform_save
      end
    end

    private def perform_save : Nil
      source_file = @source_file
      return if source_file.nil? || source_file.empty?

      # Build YAML structure
      yaml_structure = {} of String => Hash(String, Hash(String, String | Float64))

      # Build a map of original formulas by sheet and cell
      formula_map = {} of String => Hash(String, String)
      @sheet_data.each do |sheet, cells|
        formula_map[sheet] = {} of String => String
        cells.each do |cell|
          unless cell[:formula].empty?
            formula_map[sheet][cell[:cell]] = cell[:formula]
          end
        end
      end

      @sheets.each do |sheet|
        sheet_data = {} of String => Hash(String, String | Float64)
        formulas = formula_map[sheet]?

        # Scan the entire grid for this sheet
        (0...@max_row).each do |row|
          (0...@max_col).each do |col|
            # Convert to cell reference
            cell_ref = num_to_col(col + 1) + (row + 1).to_s

            # Get current value from grid or callback
            current_value = if getter = @value_getter_callback
                              getter.call(sheet, cell_ref)
                            else
                              @grid[row][col]
                            end

            # Skip empty cells
            next if current_value.empty?

            cell_info = {} of String => String | Float64

            # Check if this cell has a formula
            if formulas && formulas.has_key?(cell_ref)
              cell_info["formula"] = formulas[cell_ref]
            end

            # Try to parse as number, otherwise keep as string
            parsed = current_value.to_f?
            if parsed && current_value == parsed.to_s
              cell_info["value"] = parsed
            else
              cell_info["value"] = current_value
            end

            sheet_data[cell_ref] = cell_info
          end
        end

        # Only add sheet if it has data
        yaml_structure[sheet] = sheet_data unless sheet_data.empty?
      end

      # Convert to YAML::Any structure to add UI metadata
      yaml_any_structure = {} of YAML::Any => YAML::Any
      yaml_structure.each do |key, value|
        # Convert the inner hash (cell_data) to YAML::Any
        cell_data_any = {} of YAML::Any => YAML::Any
        value.each do |cell_ref, cell_info|
          # Convert cell_info to YAML::Any
          cell_info_any = {} of YAML::Any => YAML::Any
          cell_info.each do |info_key, info_value|
            cell_info_any[YAML::Any.new(info_key)] = YAML::Any.new(info_value)
          end
          cell_data_any[YAML::Any.new(cell_ref)] = YAML::Any.new(cell_info_any)
        end
        yaml_any_structure[YAML::Any.new(key)] = YAML::Any.new(cell_data_any)
      end

      # Add UI state metadata
      current_sheet = @sheets[@current_sheet_idx]
      ui_metadata = {} of YAML::Any => YAML::Any
      ui_metadata[YAML::Any.new("active_sheet")] = YAML::Any.new(current_sheet)
      ui_metadata[YAML::Any.new("active_cell")] = YAML::Any.new(num_to_col(@active_col + 1) + (@active_row + 1).to_s)
      yaml_any_structure[YAML::Any.new("_ui_state")] = YAML::Any.new(ui_metadata)

      # Write to YAML file
      File.open(source_file, "w") do |file|
        file.print(yaml_any_structure.to_yaml)
        file.flush
        file.fsync
      end
      show_notification("Saved to #{source_file}", 30)
      @termisu.render
    end

    private def update_formula_in_yaml(sheet_name : String, cell_ref : String, new_formula : String) : Nil
      source_file = @source_file
      return if source_file.nil? || source_file.empty?

      # Read existing YAML
      yaml_content = File.read(source_file)
      data = YAML.parse(yaml_content)

      # Update the formula - iterate through the hash to find the right cell
      data.as_h.each do |sheet_key, sheet_value|
        if sheet_key.as_s == sheet_name
          sheet_value.as_h.each do |cell_key, cell_value|
            if cell_key.as_s == cell_ref
              # Found the cell - update the formula
              cell_value.as_h[YAML::Any.new("formula")] = YAML::Any.new(new_formula)

              # Update UI state to current position
              current_sheet = @sheets[@current_sheet_idx]
              ui_metadata = {} of YAML::Any => YAML::Any
              ui_metadata[YAML::Any.new("active_sheet")] = YAML::Any.new(current_sheet)
              ui_metadata[YAML::Any.new("active_cell")] = YAML::Any.new(num_to_col(@active_col + 1) + (@active_row + 1).to_s)
              data.as_h[YAML::Any.new("_ui_state")] = YAML::Any.new(ui_metadata)

              # Write back with explicit flush
              new_yaml = data.to_yaml
              File.open(source_file, "w") do |file|
                file.print(new_yaml)
                file.flush
                # Also call fsync to ensure OS writes to disk
                file.fsync
              end
              return
            end
          end
        end
      end

      # If we get here, the cell wasn't found
      show_notification("Error: Could not update formula in YAML", 60)
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

      # Show cursor only in edit mode, hide otherwise
      if @edit_mode
        # Cursor already positioned in render_formula_bar
      else
        @termisu.hide_cursor
      end

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
      # Calculate visible columns
      cols_per_screen = (@grid_width - @row_num_width - @spacer_width) // (@cell_width + 1)
      end_col = {@col_offset + cols_per_screen, @max_col}.min

      # Render column headers (A, B, C, ...)
      col_x = @row_num_width + @spacer_width
      (@col_offset...end_col).each do |col_idx|
        col_label = num_to_col(col_idx + 1)
        draw_text_centered(col_x, @header_height, @cell_width, col_label, @fg_header, @bg_header, Termisu::Attribute::Bold | Termisu::Attribute::Reverse)
        col_x += @cell_width + 1
      end

      # Render grid cells
      (0...@grid_height).each do |screen_row|
        grid_row = @row_offset + screen_row
        break if grid_row >= @max_row

        # Row number (right-aligned in 4 characters)
        row_label = (grid_row + 1).to_s.rjust(@row_num_width)
        row_label.each_char_with_index do |char, i|
          @termisu.set_cell(i, @header_height + 1 + screen_row, char, fg: @fg_header, bg: @bg_header, attr: Termisu::Attribute::Bold | Termisu::Attribute::Reverse)
        end

        # Empty spacer column
        @termisu.set_cell(@row_num_width, @header_height + 1 + screen_row, ' ', fg: @fg_default, bg: @bg_default)

        # Cells
        col_x = @row_num_width + @spacer_width
        (@col_offset...end_col).each do |col_idx|
          break if col_idx >= @max_col

          cell_value = @grid[grid_row][col_idx]
          is_active = grid_row == @active_row && col_idx == @active_col

          # Check if value is numeric for right-alignment
          is_numeric = cell_value.match(/^\s*-?\d+\.?\d*\s*$/)

          # Truncate value to fit
          if is_numeric
            # Right-align numeric values
            truncated = truncate_value(cell_value, @cell_width)
            display_value = truncated.rjust(@cell_width)
          else
            # Left-align text values
            display_value = truncate_value(cell_value, @cell_width)
          end

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
          display_value.each_char_with_index do |char, i|
            break if i >= @cell_width
            @termisu.set_cell(col_x + i, @header_height + 1 + screen_row, char, fg: fg, bg: bg, attr: attr)
          end

          # Fill remaining cell width with spaces
          if display_value.size < @cell_width
            (display_value.size...@cell_width).each do |i|
              @termisu.set_cell(col_x + i, @header_height + 1 + screen_row, ' ', fg: fg, bg: bg, attr: attr)
            end
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

        # Position actual cursor at edit position
        cursor_x = label.size + @edit_cursor
        cursor_x = {@grid_width - 1, cursor_x}.min  # Clamp to screen width
        @termisu.set_cursor(cursor_x, formula_y)
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

      # Show notification or normal status text
      if @notification_timeout > 0
        status_text = @notification_message
      else
        status_text = "#{sheet_name}!#{cell_ref}"
      end

      # Draw status bar background
      (0...@grid_width).each do |i|
        @termisu.set_cell(i, status_y, ' ', fg: @fg_status, bg: @bg_status)
      end

      # Draw status text (use different color for notifications)
      status_color = @notification_timeout > 0 ? Termisu::Color.yellow : @fg_status
      status_text.each_char_with_index do |char, i|
        break if i >= @grid_width
        @termisu.set_cell(i, status_y, char, fg: status_color, bg: @bg_status, attr: Termisu::Attribute::Bold)
      end

      # Draw help hints
      help_text = @edit_mode ? "ENTER:Save | ESC:Cancel" : "Arrows:Move | ENTER:Edit | Tab:Sheet | S:Save | Q:Quit"
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
