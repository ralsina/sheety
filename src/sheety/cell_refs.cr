module Sheety
  # Shared A1-style column letter <-> number conversion.
  #
  # Excel columns use bijective base-26: A=1, Z=26, AA=27, and so on. This
  # conversion previously existed as seven copy-pasted variants across the
  # TUI, the dependency extractor, the runtime helpers baked into generated
  # binaries, and the Excel exporter; they are all defined here now.
  module CellRefs
    # Highest row/column sheety's fixed 1000x1000 grid addresses. Whole-column
    # ranges (A:B) clamp their rows to this.
    GRID_MAX = 1000

    # Normalized range bounds: start/end column letters and 1-based rows.
    record RangeBounds, start_col : String, start_row : Int32, end_col : String, end_row : Int32

    # Parse an A1-style range into normalized bounds, or return nil when the
    # range is unsupported.
    #
    # Accepted forms (any case, with or without $ anchors):
    # - "A1:B5"      concrete range; reversed bounds (e.g. "B2:A1") are
    #                swapped, matching Excel's normalization
    # - "A:B"        whole-column range, clamped to rows 1..GRID_MAX
    #
    # Whole-row ranges ("1:10") and anything else return nil; callers treat
    # that as a loud per-formula failure rather than silently computing
    # over an empty set of cells.
    def self.parse_range(range : String) : RangeBounds?
      cleaned = range.upcase.delete('$').strip

      if match = cleaned.match(/\A([A-Z]+)(\d+):([A-Z]+)(\d+)\z/)
        bounds = RangeBounds.new(match[1], match[2].to_i, match[3], match[4].to_i)
        normalize_bounds(bounds)
      elsif match = cleaned.match(/\A([A-Z]+):([A-Z]+)\z/)
        bounds = RangeBounds.new(match[1], 1, match[2], GRID_MAX)
        normalize_bounds(bounds)
      end
    end

    # Parse a single cell reference ("A1", "$B$2") into 1-based column/row,
    # or nil if it doesn't have that shape.
    def self.parse_ref(ref : String) : {col: Int32, row: Int32}?
      if match = ref.upcase.delete('$').strip.match(/\A([A-Z]+)(\d+)\z/)
        {col: col_to_num(match[1]), row: match[2].to_i}
      end
    end

    # Swap range endpoints so start <= end on both axes, as Excel does for
    # reversed references like B2:A1.
    private def self.normalize_bounds(bounds : RangeBounds) : RangeBounds
      start_col, end_col = bounds.start_col, bounds.end_col
      if col_to_num(start_col) > col_to_num(end_col)
        start_col, end_col = end_col, start_col
      end
      start_row = Math.min(bounds.start_row, bounds.end_row)
      end_row = Math.max(bounds.start_row, bounds.end_row)
      RangeBounds.new(start_col, start_row, end_col, end_row)
    end

    # Convert column letter(s) to a 1-based number ("A" -> 1, "AA" -> 27).
    # Input is upcased, so callers may pass either case.
    def self.col_to_num(col : String) : Int32
      num = 0
      col.upcase.each_char do |char|
        num = num * 26 + (char.ord - 'A'.ord + 1)
      end
      num
    end

    # Convert a 1-based number to column letter(s) (1 -> "A", 27 -> "AA").
    # Returns "" for 0, matching the behavior of every previous copy.
    def self.num_to_col(num : Int32) : String
      result = ""
      while num > 0
        num -= 1
        result = ('A' + (num % 26)).to_s + result
        num //= 26
      end
      result
    end
  end
end
