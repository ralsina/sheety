module Sheety
  # Shared A1-style column letter <-> number conversion.
  #
  # Excel columns use bijective base-26: A=1, Z=26, AA=27, and so on. This
  # conversion previously existed as seven copy-pasted variants across the
  # TUI, the dependency extractor, the runtime helpers baked into generated
  # binaries, and the Excel exporter; they are all defined here now.
  module CellRefs
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
