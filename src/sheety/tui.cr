require "big"
require "termisu"
require "yaml"
require "./importers/excel_exporter"
{% unless flag?(:light_mode) %}
  require "./rebuilder"
{% end %}
require "./croupier_generator"
require "./data_dir"

module Sheety
  # Notification with level and timestamp
  struct Notification
    enum Level
      Info
      Warning
      Error

      def color : Termisu::Color
        case self
        in Info    then Termisu::Color.cyan
        in Warning then Termisu::Color.yellow
        in Error   then Termisu::Color.red
        end
      end
    end

    getter text : String
    getter level : Level
    getter timestamp : Time::Instant

    def initialize(@text : String, @level : Level = Level::Info)
      @timestamp = Time.instant
    end

    def age : Time::Span
      Time.instant - @timestamp
    end

    def expired?(timeout : Time::Span = 5.seconds) : Bool
      age > timeout
    end
  end

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
    @notification : Notification?

    # Filename edit mode for save prompt
    @filename_edit_mode : Bool = false
    @filename_edit_buffer : String = ""
    @filename_edit_cursor : Int32 = 0

    # Callback for updating cell values
    @value_update_callback : Proc(String, String, String, Nil)?
    @refresh_callback : Proc(Nil)?
    @value_getter_callback : Proc(String, String, String)?
    @save_callback : Proc(Nil)?
    @source_file : String?
    @intermediate_file : String?
    @original_source_file : String? # Track the original file for saves
    @rebuilding : Bool = false
    @pending_exec : String? = nil

    def initialize(sheets : Array(String), sheet_data : Hash(String, Array(NamedTuple(cell: String, formula: String, value: String))), update_callback : Proc(String, String, String, Nil) = ->(_s : String, _c : String, _v : String) { }, source_file : String? = nil)
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

      # Enable mouse support
      @termisu.enable_mouse

      # Mouse click tracking
      @last_click_time = 0_i64
      @last_click_col = -1
      @last_click_row = -1
      @click_count = 0

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
        # Check if we need to exec to new binary
        if pending = @pending_exec
          @termisu.close
          # Exec the new binary directly, not sheety
          Process.exec(pending, [] of String)
        end

        # Update notification state (auto-clear expired)
        had_notification = @notification != nil
        update_notification
        notification_changed = (@notification != nil) != had_notification

        # Always render while rebuilding to keep notification visible
        if @rebuilding
          render
        end

        if event = @termisu.poll_event(100)
          case event
          when Termisu::Event::Key
            handle_key_event(event)
            # Clear notification on any key press
            clear_notification
            render if should_render?(event)
          when Termisu::Event::Mouse
            handle_mouse_event(event)
            render
          when Termisu::Event::Resize
            handle_resize(event)
            render
          end
        elsif notification_changed
          # Re-render if notification state changed (expired/cleared)
          render
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

      # If no data, we're done
      return if data.nil? || data.empty?

      # Build a map of cells with their formulas for quick lookup
      cell_map = {} of String => NamedTuple(formula: String, value: String)
      data.each do |cell|
        cell_map[cell[:cell]] = {formula: cell[:formula], value: cell[:value]}
      end

      # Only fill cells that we know have data in cell_map
      # This avoids calling the callback for all 1 million cells
      cell_map.each do |cell_ref, cell_data|
        # Parse cell reference to get row and col
        if match = cell_ref.match(/^([A-Za-z]+)(\d+)$/)
          col_str = match[1]
          row_str = match[2]

          col = col_to_num(col_str) - 1
          row = row_str.to_i - 1

          # Check bounds
          next if row < 0 || row >= @max_row || col < 0 || col >= @max_col

          # Try to get value from callback first (for edited cells), then from data
          value = if callback = @value_getter_callback
                    fetched = callback.call(sheet_name, cell_ref)
                    if fetched.empty?
                      # Not in store, use original data
                      cell_data[:value]
                    else
                      fetched
                    end
                  else
                    # No callback, use original data
                    cell_data[:value]
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
      if @filename_edit_mode
        handle_filename_edit_key_event(event)
      elsif @edit_mode
        handle_edit_key_event(event)
      else
        handle_normal_key_event(event)
      end
    end

    private def handle_normal_key_event(event : Termisu::Event::Key) : Nil
      # Ignore input while rebuilding
      return if @rebuilding

      # Check for character-based keys first (for uppercase S)
      if char = event.char
        case char
        when 's', 'S'
          save_to_yaml
          return
        end
      end

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
      when .q?
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

    private def handle_filename_edit_key_event(event : Termisu::Event::Key) : Nil
      case event.key
      when .escape?
        # Cancel filename edit
        @filename_edit_mode = false
        @filename_edit_buffer = ""
        @filename_edit_cursor = 0
      when .enter?
        # Save with the entered filename
        save_with_filename(@filename_edit_buffer)
        @filename_edit_mode = false
        @filename_edit_buffer = ""
        @filename_edit_cursor = 0
      when .backspace?
        # Delete character before cursor
        if @filename_edit_cursor > 0
          @filename_edit_buffer = @filename_edit_buffer[0...@filename_edit_cursor - 1] + @filename_edit_buffer[@filename_edit_cursor..]
          @filename_edit_cursor -= 1
        end
      when .delete?
        # Delete character at cursor
        if @filename_edit_cursor < @filename_edit_buffer.size
          @filename_edit_buffer = @filename_edit_buffer[0...@filename_edit_cursor] + @filename_edit_buffer[@filename_edit_cursor + 1..]
        end
      when .left?
        # Move cursor left
        @filename_edit_cursor = {@filename_edit_cursor - 1, 0}.max
      when .right?
        # Move cursor right
        @filename_edit_cursor = {@filename_edit_cursor + 1, @filename_edit_buffer.size}.min
      when .home?
        # Move cursor to start
        @filename_edit_cursor = 0
      when .end?
        # Move cursor to end
        @filename_edit_cursor = @filename_edit_buffer.size
      else
        # Check if this is a printable character
        if char = event.char
          # Don't include control characters
          if char.printable?
            @filename_edit_buffer = @filename_edit_buffer[0...@filename_edit_cursor] + char + @filename_edit_buffer[@filename_edit_cursor..]
            @filename_edit_cursor += 1
          end
        end
      end
    end

    private def handle_mouse_event(event : Termisu::Event::Mouse) : Nil
      x = event.x
      y = event.y

      # Handle wheel events for scrolling
      if event.wheel?
        handle_mouse_wheel(event.button)
        return
      end

      # Only handle clicks in the grid area
      grid_start_y = @header_height + 1
      grid_end_y = grid_start_y + @grid_height

      return if y < grid_start_y || y >= grid_end_y

      if event.press?
        handle_mouse_click(x, y, event.button)
      end
    end

    private def handle_mouse_wheel(button : Termisu::Event::Mouse::Button) : Nil
      # Scroll up or down
      case button
      when Termisu::Event::Mouse::Button::WheelUp
        move_active(-3, 0) # Scroll up 3 rows
      when Termisu::Event::Mouse::Button::WheelDown
        move_active(3, 0) # Scroll down 3 rows
      else
        # Other wheel events (left/right) could scroll horizontally
      end
    end

    private def handle_mouse_click(x : Int32, y : Int32, button : Termisu::Event::Mouse::Button) : Nil
      # Calculate which cell was clicked
      grid_y = y - @header_height - 1

      # Calculate column position (accounting for row labels and spacer)
      cols_per_screen = (@grid_width - @row_num_width - @spacer_width) // (@cell_width + 1)

      # Find which column was clicked
      clicked_col = -1
      col_x = @row_num_width + @spacer_width
      (@col_offset...@col_offset + cols_per_screen).each do |col_idx|
        break if col_idx >= @max_col
        if x >= col_x && x < col_x + @cell_width + 1
          clicked_col = col_idx
          break
        end
        col_x += @cell_width + 1
      end

      return if clicked_col == -1

      # Calculate row (accounting for offset)
      clicked_row = @row_offset + grid_y - 1

      return if clicked_row >= @max_row || clicked_row < 0

      # Detect double-click (within 500ms and same position)
      current_time = Time.utc.to_unix_ms
      is_double_click = false

      if @last_click_col == clicked_col && @last_click_row == clicked_row
        time_diff = current_time - @last_click_time
        if time_diff < 500
          is_double_click = true
          @click_count = 0
        end
      end

      @last_click_time = current_time
      @last_click_col = clicked_col
      @last_click_row = clicked_row

      # Update active cell
      @active_col = clicked_col
      @active_row = clicked_row

      adjust_row_offset
      adjust_col_offset

      # Double-click enters edit mode
      if is_double_click
        enter_edit_mode
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
      # In light mode, don't allow editing formula cells
      {% if flag?(:light_mode) %}
        if formula_cell?
          show_notification("Formulas are read-only in light mode", Notification::Level::Warning)
          return
        end
      {% end %}

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

    private def show_notification(message : String, level : Notification::Level = Notification::Level::Info) : Nil
      @notification = Notification.new(message, level)
    end

    private def clear_notification : Nil
      @notification = nil
    end

    private def update_notification : Nil
      # Auto-clear expired notifications, but not while rebuilding
      return if @rebuilding
      if notif = @notification
        @notification = nil if notif.expired?
      end
    end

    private def save_edit_value : Nil
      sheet_name = @sheets[@current_sheet_idx]
      cell_ref = current_cell_ref

      # Check if the edit buffer contains a formula (starts with =)
      is_formula = @edit_buffer.starts_with?("=")

      if is_formula
        # Ensure formula starts with = (add it if missing)
        unless @edit_buffer.starts_with?("=")
          @edit_buffer = "=" + @edit_buffer
        end

        # Update or add the formula in sheet_data
        if cells = @sheet_data[sheet_name]?
          # Find the cell in sheet_data and update its formula
          found = false
          cells.each_with_index do |cell|
            if cell[:cell] == cell_ref
              # Update the formula in the array
              @sheet_data[sheet_name] = cells.map do |c|
                if c[:cell] == cell_ref
                  {cell: c[:cell], formula: @edit_buffer, value: c[:value]}
                else
                  c
                end
              end
              found = true
              break
            end
          end

          # If cell wasn't found in sheet_data, add it as a new formula cell
          unless found
            @sheet_data[sheet_name] = cells + [{cell: cell_ref, formula: @edit_buffer, value: ""}]
          end
        else
          # No cells for this sheet yet, create the array with this formula cell
          @sheet_data[sheet_name] = [{cell: cell_ref, formula: @edit_buffer, value: ""}]
        end

        # Ensure we have an intermediate file
        rebuild_file = @intermediate_file || @source_file
        return unless rebuild_file && !rebuild_file.empty?

        # Save current state to intermediate file as YAML
        # Build YAML structure directly from current state
        save_to_yaml_file(rebuild_file)

        if rebuild_file && !rebuild_file.empty?
          {% unless flag?(:light_mode) %}
            # Check if crystal is available for rebuilding
            unless Process.find_executable("crystal")
              show_notification("Cannot rebuild: crystal not found", Notification::Level::Error)
              @edit_mode = false
              @edit_buffer = ""
              @edit_cursor = 0
              render
              return
            end

            # Show notification
            show_notification("Rebuilding...", Notification::Level::Info)
            @edit_mode = false
            @edit_buffer = ""
            @edit_cursor = 0
            render

            # Build the new binary in background
            # After build completes, we'll exec to replace this process
            @rebuilding = true
            spawn do
              # Use Rebuilder to rebuild in-process
              # Use the original source file for tracking, but rebuild from intermediate file
              original_for_rebuild = @original_source_file || rebuild_file
              rebuilder = Sheety::Rebuilder.new(original_for_rebuild)
              rebuilder.set_intermediate_file(rebuild_file)

              # Set UUID if we have it (from _ui_state in YAML)
              begin
                yaml_content = File.read(rebuild_file)
                data = YAML.parse(yaml_content)
                if data.as_h? && data["_ui_state"]? && data["_ui_state"]["spreadsheet_uuid"]?
                  rebuilder.set_spreadsheet_uuid(data["_ui_state"]["spreadsheet_uuid"].as_s)
                end
              rescue
                # Ignore errors reading UUID
              end

              binary_path = rebuilder.rebuild

              @rebuilding = false

              if binary_path && File.exists?(binary_path)
                @pending_exec = binary_path
              else
                show_notification("Rebuild failed", Notification::Level::Error)
              end
            end
          {% else %}
            # Light mode: formulas are read-only, just show a message
            show_notification("Formulas are read-only in light mode", Notification::Level::Warning)
          {% end %}
        else
          # Fallback: exit with code 42 if no source file known
          exit 42
        end
      else
        # Regular value cell - just update via callback
        if callback = @value_update_callback
          callback.call(sheet_name, cell_ref, @edit_buffer)
        end

        # Update @sheet_data so the new cell is included in future initialize_grid calls
        if data = @sheet_data[sheet_name]?
          # Find existing cell or add new one
          existing_idx = data.index { |c| c[:cell] == cell_ref }
          if existing_idx
            # Update existing cell's value
            data[existing_idx] = {cell: cell_ref, formula: data[existing_idx][:formula], value: @edit_buffer}
          else
            # Add new cell to data
            data << {cell: cell_ref, formula: "", value: @edit_buffer}
          end
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

    def set_intermediate_file(intermediate_file : String) : Nil
      @intermediate_file = intermediate_file
    end

    def set_original_source_file(original_file : String) : Nil
      @original_source_file = original_file
    end

    def refresh_current_sheet : Nil
      initialize_grid
    end

    def save_to_yaml : Nil
      # Use original source file for saves, not intermediate file
      source_file = @original_source_file || @source_file
      return if source_file.nil? || source_file.empty?

      # Enter filename edit mode with current filename as default
      @filename_edit_mode = true
      @filename_edit_buffer = source_file
      @filename_edit_cursor = source_file.size
    end

    private def save_with_filename(filename : String) : Nil
      # Save to the specified filename
      do_save(filename)

      # Update original_source_file to remember this choice for future saves
      @original_source_file = filename

      # Also save to the intermediate file if it exists and is different
      if intermediate = @intermediate_file
        do_save(intermediate) if intermediate != filename
      end
    end

    private def do_save(filename : String) : Nil
      # Temporarily override source_file for this save
      original_source_file = @source_file
      @source_file = filename

      if callback = @save_callback
        callback.call
      else
        # Default save implementation
        perform_save
      end

      # Restore original source file
      @source_file = original_source_file
    end

    private def perform_save : Nil
      source_file = @source_file
      return if source_file.nil? || source_file.empty?

      # Build internal format structure
      internal_format = {} of String => Hash(String, Hash(String, Sheety::Functions::CellValue))

      @sheets.each do |sheet|
        sheet_data = {} of String => Hash(String, Sheety::Functions::CellValue)

        # Only iterate over cells that have original data in @sheet_data
        # This avoids scanning all 1 million cells
        original_cells = @sheet_data[sheet]?
        next unless original_cells

        original_cells.each do |cell|
          cell_ref = cell[:cell]

          # Get current value from grid or callback
          current_value = if getter = @value_getter_callback
                            getter.call(sheet, cell_ref)
                          else
                            # Parse cell reference to get grid position
                            if match = cell_ref.match(/^([A-Za-z]+)(\d+)$/)
                              col = col_to_num(match[1]) - 1
                              row = match[2].to_i - 1
                              if row >= 0 && row < @max_row && col >= 0 && col < @max_col
                                @grid[row][col]
                              else
                                cell[:value]
                              end
                            else
                              cell[:value]
                            end
                          end

          # Skip empty cells
          next if current_value.empty?

          cell_info = {} of String => Sheety::Functions::CellValue

          # Check if this cell has a formula
          unless cell[:formula].empty?
            cell_info["formula"] = cell[:formula]
          end

          # Try to parse as number, otherwise keep as string
          parsed = current_value.to_f?
          if parsed && current_value == parsed.to_s
            cell_info["value"] = BigFloat.new(parsed, precision: 64)
          else
            cell_info["value"] = current_value
          end

          sheet_data[cell_ref] = cell_info
        end

        # Only add sheet if it has data
        internal_format[sheet] = sheet_data unless sheet_data.empty?
      end

      # Determine output format based on file extension
      ext = File.extname(source_file).downcase

      if ext == ".sheety"
        # Generate and compile a standalone binary
        show_notification("Compiling binary...", Notification::Level::Info)
        @termisu.render

        # First, generate Crystal source code
        generator = CroupierGenerator.new
        initial_values = Hash(String, BigFloat | String | Bool).new

        # Populate formulas and initial values from internal_format
        # This is the single source of truth for all cell data
        internal_format.each do |sheet, cells|
          cells.each do |cell_ref, cell_data|
            key = sheet.empty? ? cell_ref : "#{sheet}!#{cell_ref}"

            # Add formula if present
            if cell_data.has_key?("formula")
              generator.add_formula(cell_ref, cell_data["formula"].as(String), sheet)
            end

            # Add value as initial value (for all cells, with or without formulas)
            if cell_data.has_key?("value")
              value = cell_data["value"]
              # Convert ErrorValue and Nil to appropriate types
              case value
              when Sheety::Functions::ErrorValue
                initial_values[key] = value.to_s
              when Nil
                initial_values[key] = ""
              else
                initial_values[key] = value
              end
            end
          end
        end

        # Generate the source code (interactive for TUI binary)
        source_code = generator.generate_source(initial_values, true, source_file, nil)

        if source_code.empty?
          show_notification("Failed to generate source code", Notification::Level::Error)
          @termisu.render
          return
        end

        # Create a temporary source file
        temp_source = File.join(DataDir.path, "tmp", "#{File.basename(source_file, ext)}.cr")
        File.write(temp_source, source_code)

        # Compile the binary with appropriate flags
        # Keep the .sheety extension for the binary
        binary_name = source_file

        compile_result = Process.run("crystal", ["build", "-Dpreview_mt", "-Dno_embedded_files", temp_source, "-o", binary_name],
          output: Process::Redirect::Inherit, error: Process::Redirect::Inherit)

        if compile_result.success?
          show_notification("Compiled binary: #{binary_name}", Notification::Level::Info)
        else
          show_notification("Compilation failed", Notification::Level::Error)
        end
      elsif ext == ".cr"
        # Generate Crystal source code
        generator = CroupierGenerator.new
        initial_values = Hash(String, BigFloat | String | Bool).new

        # Populate formulas and initial values from internal_format
        internal_format.each do |sheet, cells|
          cells.each do |cell_ref, cell_data|
            key = sheet.empty? ? cell_ref : "#{sheet}!#{cell_ref}"

            # Add formula if present
            if cell_data.has_key?("formula")
              generator.add_formula(cell_ref, cell_data["formula"].as(String), sheet)
            end

            # Add value as initial value
            if cell_data.has_key?("value")
              value = cell_data["value"]
              case value
              when Sheety::Functions::ErrorValue
                initial_values[key] = value.to_s
              when Nil
                initial_values[key] = ""
              else
                initial_values[key] = value
              end
            end
          end
        end

        # Generate the source code (non-interactive for standalone code generation)
        source_code = generator.generate_source(initial_values, true, source_file, nil)

        if source_code.empty?
          show_notification("Failed to generate source code", Notification::Level::Error)
        else
          File.write(source_file, source_code)
          show_notification("Generated Crystal code: #{source_file}", Notification::Level::Info)
        end
      elsif ext == ".xlsx"
        # Save as Excel
        begin
          Sheety::ExcelExporter.export_to_xlsx(internal_format, source_file)
          show_notification("Saved to #{source_file}", Notification::Level::Info)
        rescue ex : Exception
          show_notification("Excel export failed: #{ex.message}", Notification::Level::Error)
        end
      else
        # Save as YAML (default for .yaml, .yml, or unknown extensions)
        yaml_structure = internal_format_to_yaml_structure(internal_format)
        File.open(source_file, "w") do |file|
          file.print(yaml_structure.to_yaml)
          file.flush
          file.fsync
        end
        show_notification("Saved to #{source_file}", Notification::Level::Info)
      end

      @termisu.render
    end

    private def internal_format_to_yaml_structure(internal_format : Hash(String, Hash(String, Hash(String, Sheety::Functions::CellValue)))) : Hash(YAML::Any, YAML::Any)
      yaml_any_structure = {} of YAML::Any => YAML::Any
      internal_format.each do |sheet, cells|
        cell_data_any = {} of YAML::Any => YAML::Any
        cells.each do |cell_ref, cell_info|
          cell_info_any = {} of YAML::Any => YAML::Any
          cell_info.each do |info_key, info_value|
            # Convert CellValue to YAML-compatible types
            yaml_value = convert_cell_value_to_yaml(info_value)
            cell_info_any[YAML::Any.new(info_key)] = YAML::Any.new(yaml_value)
          end
          cell_data_any[YAML::Any.new(cell_ref)] = YAML::Any.new(cell_info_any)
        end
        yaml_any_structure[YAML::Any.new(sheet)] = YAML::Any.new(cell_data_any)
      end

      # Add UI state metadata
      current_sheet = @sheets[@current_sheet_idx]
      ui_metadata = {} of YAML::Any => YAML::Any
      ui_metadata[YAML::Any.new("active_sheet")] = YAML::Any.new(current_sheet)
      ui_metadata[YAML::Any.new("active_cell")] = YAML::Any.new(num_to_col(@active_col + 1) + (@active_row + 1).to_s)
      yaml_any_structure[YAML::Any.new("_ui_state")] = YAML::Any.new(ui_metadata)

      yaml_any_structure
    end

    private def convert_cell_value_to_yaml(value : Sheety::Functions::CellValue) : String | BigFloat | Bool | Nil
      case value
      when String, BigFloat, Bool, Nil
        value
      when Sheety::Functions::ErrorValue
        value.to_s
      else
        value.to_s
      end
    end

    private def save_to_yaml_file(filename : String) : Nil
      # Build YAML structure from current state
      yaml_structure = {} of YAML::Any => YAML::Any

      @sheets.each do |sheet|
        sheet_data_any = {} of YAML::Any => YAML::Any

        # Only iterate over cells that have data
        original_cells = @sheet_data[sheet]?
        next unless original_cells

        original_cells.each do |cell|
          cell_ref = cell[:cell]

          # Get current value from Croupier store
          current_value = if getter = @value_getter_callback
                            getter.call(sheet, cell_ref)
                          else
                            cell[:value]
                          end

          # Skip empty cells UNLESS they have a formula
          next if current_value.empty? && cell[:formula].empty?

          cell_info_any = {} of YAML::Any => YAML::Any

          # Add formula if present
          unless cell[:formula].empty?
            cell_info_any[YAML::Any.new("formula")] = YAML::Any.new(cell[:formula])
          end

          # Add value
          parsed = current_value.to_f?
          if parsed && current_value == parsed.to_s
            cell_info_any[YAML::Any.new("value")] = YAML::Any.new(parsed)
          else
            cell_info_any[YAML::Any.new("value")] = YAML::Any.new(current_value)
          end

          sheet_data_any[YAML::Any.new(cell_ref)] = YAML::Any.new(cell_info_any)
        end

        yaml_structure[YAML::Any.new(sheet)] = YAML::Any.new(sheet_data_any) unless sheet_data_any.empty?
      end

      # Add UI state
      ui_metadata = {} of YAML::Any => YAML::Any
      ui_metadata[YAML::Any.new("active_sheet")] = YAML::Any.new(@sheets[@current_sheet_idx])
      ui_metadata[YAML::Any.new("active_cell")] = YAML::Any.new(num_to_col(@active_col + 1) + (@active_row + 1).to_s)

      # Preserve spreadsheet_uuid if it exists
      begin
        if File.exists?(filename)
          existing_content = File.read(filename)
          existing_data = YAML.parse(existing_content)
          if existing_data.as_h? && existing_data["_ui_state"]? && existing_data["_ui_state"]["spreadsheet_uuid"]?
            ui_metadata[YAML::Any.new("spreadsheet_uuid")] = existing_data["_ui_state"]["spreadsheet_uuid"]
          end
        end
      rescue
        # Ignore errors
      end

      yaml_structure[YAML::Any.new("_ui_state")] = YAML::Any.new(ui_metadata)

      # Write to file
      File.open(filename, "w") do |file|
        file.print(yaml_structure.to_yaml)
        file.flush
        file.fsync
      end
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
      if @edit_mode || @filename_edit_mode
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

      # Get current cell info (needed for edit/normal modes)
      cell_ref = current_cell_ref
      formula = current_cell_formula

      if @filename_edit_mode
        # Filename edit mode: show the save filename prompt
        label = "Save as: "
        label.each_char_with_index do |char, i|
          break if i >= @grid_width
          @termisu.set_cell(i, formula_y, char, fg: @fg_header, bg: @bg_default, attr: Termisu::Attribute::Bold)
        end

        # Show filename edit buffer
        display_buffer = truncate_value(@filename_edit_buffer, @grid_width - label.size)
        display_buffer.each_char_with_index do |char, i|
          x = label.size + i
          break if x >= @grid_width
          @termisu.set_cell(x, formula_y, char, fg: @fg_active, bg: @bg_default, attr: Termisu::Attribute::Bold)
        end

        # Position actual cursor at edit position
        cursor_x = label.size + @filename_edit_cursor
        cursor_x = {@grid_width - 1, cursor_x}.min # Clamp to screen width
        @termisu.set_cursor(cursor_x, formula_y)
      elsif @edit_mode
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
        cursor_x = {@grid_width - 1, cursor_x}.min # Clamp to screen width
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
      if notif = @notification
        status_text = notif.text
        status_color = notif.level.color
      else
        status_text = "#{sheet_name}!#{cell_ref}"
        status_color = @fg_status
      end

      # Draw status bar background
      (0...@grid_width).each do |i|
        @termisu.set_cell(i, status_y, ' ', fg: @fg_status, bg: @bg_status)
      end

      # Draw status text (use notification color if active)
      status_text.each_char_with_index do |char, i|
        break if i >= @grid_width
        @termisu.set_cell(i, status_y, char, fg: status_color, bg: @bg_status, attr: Termisu::Attribute::Bold)
      end

      # Draw notification or help hints in the right side of status bar
      if notif = @notification
        # Show notification on the right side with its color
        notif_x = {@grid_width - notif.text.size, 0}.max
        notif.text.each_char_with_index do |char, i|
          if notif_x + i < @grid_width
            @termisu.set_cell(notif_x + i, status_y, char, fg: notif.level.color, bg: @bg_status, attr: Termisu::Attribute::Bold)
          end
        end
      else
        # Show help hints when no notification
        help_text = if @filename_edit_mode
                      "ENTER:Save | ESC:Cancel"
                    elsif @edit_mode
                      "ENTER:Save | ESC:Cancel"
                    else
                      "Arrows:Move | ENTER:Edit | Click:Select | DblClick:Edit | Tab:Sheet | S:Save | Q:Quit"
                    end
        help_x = {@grid_width - help_text.size, 0}.max

        help_text.each_char_with_index do |char, i|
          if help_x + i < @grid_width
            @termisu.set_cell(help_x + i, status_y, char, fg: @fg_status, bg: @bg_status, attr: Termisu::Attribute::Dim)
          end
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
