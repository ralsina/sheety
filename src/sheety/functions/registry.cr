module Sheety
  module Functions
    # Type alias for Excel cell values
    alias CellValue = Float64 | String | Bool | ErrorValue | Nil

    # Excel error value representation
    struct ErrorValue
      getter error : String

      def initialize(@error : String)
      end

      def to_s(io : IO) : Nil
        io << @error
      end

      def inspect(io : IO) : Nil
        io << "<Error #{@error}>"
      end
    end

    # Create common Excel errors
    def self.div0 : ErrorValue
      ErrorValue.new("#DIV/0!")
    end

    def self.value : ErrorValue
      ErrorValue.new("#VALUE!")
    end

    def self.ref : ErrorValue
      ErrorValue.new("#REF!")
    end

    def self.name : ErrorValue
      ErrorValue.new("#NAME?")
    end

    def self.num : ErrorValue
      ErrorValue.new("#NUM!")
    end

    def self.na : ErrorValue
      ErrorValue.new("#N/A")
    end

    # Helper to extract numeric values from cell values, ignoring errors and non-numeric values
    private def self.extract_numbers(values : Array(CellValue)) : Array(Float64)
      result = [] of Float64
      values.each do |v|
        case v
        when Float64
          result << v
        when String
          # Try to convert string to number
          begin
            result << v.to_f
          rescue
            # Skip non-numeric strings
          end
        end
      end
      result
    end

    # Helper to extract all values excluding errors
    private def self.extract_values(values : Array(CellValue)) : Array(CellValue)
      values.reject(ErrorValue)
    end

    # Helper to convert cell value to string
    def self.to_string(value : CellValue) : String
      case value
      when String     then value
      when Float64    then value.to_s
      when Bool       then value ? "TRUE" : "FALSE"
      when ErrorValue then value.to_s
      when Nil        then ""
      else
        ""
      end
    end

    # Helper to convert cell value to float
    def self.to_float(value : CellValue) : Float64?
      case value
      when Float64 then value
      when String
        begin
          value.to_f
        rescue
          nil
        end
      when Bool then value ? 1.0 : 0.0
      else           nil
      end
    end

    # Math functions

    # SUM: Adds all numbers in a range
    def self.sum(values : Array(CellValue)) : CellValue
      numbers = extract_numbers(values)
      numbers.sum
    end

    # AVERAGE: Returns the arithmetic mean of arguments
    def self.average(values : Array(CellValue)) : CellValue
      numbers = extract_numbers(values)
      return div0 if numbers.empty?
      numbers.sum / numbers.size
    end

    # MIN: Returns the minimum value in a range
    def self.min(values : Array(CellValue)) : CellValue
      numbers = extract_numbers(values)
      return num if numbers.empty?
      numbers.min
    end

    # MAX: Returns the maximum value in a range
    def self.max(values : Array(CellValue)) : CellValue
      numbers = extract_numbers(values)
      return num if numbers.empty?
      numbers.max
    end

    # COUNT: Counts how many numbers are in the list of arguments
    def self.count(values : Array(CellValue)) : CellValue
      extract_numbers(values).size.to_f
    end

    # ROUND: Rounds a number to a specified number of digits
    def self.round(value : CellValue, digits : CellValue = 0.0) : CellValue
      num = to_float(value)
      d = to_float(digits) || 0.0
      return value if num.nil?
      num.round(d.to_i).to_f
    end

    # ABS: Returns the absolute value of a number
    def self.abs(value : CellValue) : CellValue
      num = to_float(value)
      return value if num.nil?
      num.abs
    end

    # POWER: Returns the result of a number raised to a power
    def self.power(base : CellValue, exponent : CellValue) : CellValue
      b = to_float(base)
      e = to_float(exponent)
      return value if b.nil? || e.nil?
      b ** e
    end

    # SQRT: Returns the square root of a number
    def self.sqrt(value : CellValue) : CellValue
      num = to_float(value)
      return value if num.nil?
      return num if num < 0
      Math.sqrt(num)
    end

    # MOD: Returns the remainder after division
    def self.mod(number : CellValue, divisor : CellValue) : CellValue
      n = to_float(number)
      d = to_float(divisor)
      return value if n.nil? || d.nil?
      return div0 if d == 0
      n % d
    end

    # INT: Rounds a number down to the nearest integer
    def self.int(value : CellValue) : CellValue
      num = to_float(value)
      return value if num.nil?
      num.floor.to_f
    end

    # Logical functions

    # IF: Specifies a logical test to perform
    def self.if(condition : CellValue, true_value : CellValue, false_value : CellValue) : CellValue
      cond = to_bool(condition)
      return true_value if cond == true
      false_value
    end

    # Helper to convert cell value to boolean
    private def self.to_bool(value : CellValue) : Bool?
      case value
      when Bool    then value
      when Float64 then value != 0.0
      when String  then !value.empty?
      else              nil
      end
    end

    # AND: Returns TRUE if all arguments are TRUE
    def self.and(values : Array(CellValue)) : CellValue
      return false if values.empty?
      values.all? do |v|
        case v
        when Bool    then v
        when Float64 then v != 0.0
        when String  then !v.empty?
        else              false
        end
      end
    end

    # OR: Returns TRUE if any argument is TRUE
    def self.or(values : Array(CellValue)) : CellValue
      return false if values.empty?
      values.any? do |v|
        case v
        when Bool    then v
        when Float64 then v != 0.0
        when String  then !v.empty?
        else              false
        end
      end
    end

    # NOT: Reverses the logic of its argument
    def self.not(value : CellValue) : CellValue
      result = to_bool(value)
      return false if result.nil?
      !result
    end

    # Text functions

    # CONCAT: Joins several text strings into one text string
    def self.concat(values : Array(CellValue)) : CellValue
      values.map { |v| to_string(v) }.join
    end

    # LEFT: Returns the specified number of characters from the start of a text string
    def self.left(text : CellValue, num_chars : CellValue = 1.0) : CellValue
      str = to_string(text)
      n = to_float(num_chars) || 1.0
      str[0...(n.to_i)]
    end

    # RIGHT: Returns the specified number of characters from the end of a text string
    def self.right(text : CellValue, num_chars : CellValue = 1.0) : CellValue
      str = to_string(text)
      n = to_float(num_chars) || 1.0
      return "" if n > str.size
      str[-(n.to_i)..]
    end

    # MID: Returns a specific number of characters from a text string starting at a specified position
    def self.mid(text : CellValue, start_num : CellValue, num_chars : CellValue) : CellValue
      str = to_string(text)
      start = to_float(start_num) || 1.0
      n = to_float(num_chars) || 0.0
      return "" if start < 1
      start_idx = (start.to_i - 1)
      return "" if start_idx >= str.size
      str[start_idx...(start_idx + n.to_i)]
    end

    # LEN: Returns the number of characters in a text string
    def self.len(text : CellValue) : CellValue
      to_string(text).size.to_f
    end

    # UPPER: Converts text to uppercase
    def self.upper(text : CellValue) : CellValue
      to_string(text).upcase
    end

    # LOWER: Converts text to lowercase
    def self.lower(text : CellValue) : CellValue
      to_string(text).downcase
    end

    # TRIM: Removes spaces from text except for single spaces between words
    def self.trim(text : CellValue) : CellValue
      to_string(text).gsub(/\s+/, " ").strip
    end

    # Comparison functions

    # Equality test
    def self.eq(left : CellValue, right : CellValue) : CellValue
      result = compare_values(left, right)
      result == 0
    end

    # Inequality test
    def self.ne(left : CellValue, right : CellValue) : CellValue
      result = compare_values(left, right)
      result.nil? ? false : result != 0
    end

    # Less than
    def self.lt(left : CellValue, right : CellValue) : CellValue
      result = compare_values(left, right)
      result.nil? ? false : result < 0
    end

    # Less than or equal
    def self.le(left : CellValue, right : CellValue) : CellValue
      result = compare_values(left, right)
      result.nil? ? false : result <= 0
    end

    # Greater than
    def self.gt(left : CellValue, right : CellValue) : CellValue
      result = compare_values(left, right)
      result.nil? ? false : result > 0
    end

    # Greater than or equal
    def self.ge(left : CellValue, right : CellValue) : CellValue
      result = compare_values(left, right)
      result.nil? ? false : result >= 0
    end

    # Helper to compare two cell values
    # Returns -1 if left < right, 0 if equal, 1 if left > right
    private def self.compare_values(left : CellValue, right : CellValue) : Int32?
      # Handle errors
      return 0 if left.is_a?(ErrorValue) || right.is_a?(ErrorValue)

      case {left, right}
      when {Float64, Float64}
        left <=> right
      when {Float64, String}
        -1 # Numbers are always less than strings in Excel
      when {String, Float64}
        1
      when {String, String}
        left <=> right
      when {Bool, Bool}
        (left ? 1 : 0) <=> (right ? 1 : 0)
      else
        0
      end
    end
  end
end
