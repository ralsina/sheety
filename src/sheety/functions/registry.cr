require "big"

module Sheety
  module Functions
    # Type alias for Excel cell values
    # Using BigFloat for arbitrary precision arithmetic
    alias CellValue = BigFloat | String | Bool | ErrorValue | Nil

    # Default precision for BigFloat operations (enough for most Excel use cases)
    DEFAULT_PRECISION = 64

    # Helper to convert to BigFloat with default precision
    private def self.to_big_float(value : CellValue, precision : Int = DEFAULT_PRECISION) : BigFloat?
      case value
      when BigFloat
        value
      when Int, Int32
        BigFloat.new(value.to_f, precision: precision)
      when String
        begin
          BigFloat.new(value, precision: precision)
        rescue
          nil
        end
      when Bool
        value ? BigFloat.new(1.0, precision: precision) : BigFloat.new(0.0, precision: precision)
      else
        nil
      end
    end

    # Helper methods for date operations
    private struct TimeHelpers
      # Helper to add years to a Time
      def self.add_years(time : Time, years : Int32) : Time
        year = time.year + years
        month = time.month
        day = time.day

        # Handle Feb 29 on non-leap years
        if month == 2 && day == 29 && !leap_year?(year)
          day = 28
        end

        Time.utc(year, month, day, 0, 0, 0)
      end

      # Helper to add months to a Time
      def self.add_months(time : Time, months : Int32) : Time
        total_months = time.year * 12 + (time.month - 1) + months
        new_year = total_months // 12
        new_month = (total_months % 12) + 1

        day = time.day
        max_day = utc_days_in_month(new_year, new_month)
        day = max_day if day > max_day

        Time.utc(new_year, new_month, day, 0, 0, 0)
      end

      # Check if year is a leap year
      def self.leap_year?(year : Int32) : Bool
        (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
      end

      # Get days in month
      def self.utc_days_in_month(year : Int32, month : Int32) : Int32
        case month
        when 1, 3, 5, 7, 8, 10, 12
          31
        when 4, 6, 9, 11
          30
        when 2
          leap_year?(year) ? 29 : 28
        else
          30
        end
      end
    end

    def self.utc_days_in_month(year : Int32, month : Int32) : Int32
      TimeHelpers.utc_days_in_month(year, month)
    end

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
    private def self.extract_numbers(values : Array(CellValue)) : Array(BigFloat)
      result = [] of BigFloat
      values.each do |v|
        case v
        when BigFloat
          result << v
        when String
          # Try to convert string to number
          begin
            result << BigFloat.new(v, precision: DEFAULT_PRECISION)
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
      when BigFloat   then value.to_s
      when Bool       then value ? "TRUE" : "FALSE"
      when ErrorValue then value.to_s
      when Nil        then ""
      else
        ""
      end
    end

    # Helper to convert cell value to BigFloat
    def self.to_float(value : CellValue) : BigFloat?
      case value
      when BigFloat
        value
      when String
        begin
          BigFloat.new(value, precision: DEFAULT_PRECISION)
        rescue
          nil
        end
      when Bool
        value ? BigFloat.new(1.0, precision: DEFAULT_PRECISION) : BigFloat.new(0.0, precision: DEFAULT_PRECISION)
      else
        nil
      end
    end

    # Math functions

    # SUM: Adds all numbers - generic version that handles mixed types
    private def self.sum_impl(*args : CellValue | Array(CellValue)) : BigFloat
      all_numbers = Array(BigFloat).new
      args.each do |arg|
        case arg
        when Array
          all_numbers.concat(extract_numbers(arg))
        else
          if num = to_float(arg)
            all_numbers << num
          end
        end
      end
      all_numbers.sum
    end

    # SUM: Multiple arguments (arrays and/or scalars mixed)
    def self.sum(*values : CellValue | Array(CellValue)) : CellValue
      sum_impl(*values)
    end

    # SUM: Multiple arrays only
    def self.sum(*values : Array(CellValue)) : CellValue
      all_numbers = Array(BigFloat).new
      values.each do |arr|
        all_numbers.concat(extract_numbers(arr))
      end
      all_numbers.sum
    end

    # SUM: Single array overload for backward compatibility
    def self.sum(values : Array(CellValue)) : CellValue
      numbers = extract_numbers(values)
      numbers.sum
    end

    # AVERAGE: Generic version that handles mixed types
    private def self.average_impl(*args : CellValue | Array(CellValue)) : CellValue
      all_numbers = Array(BigFloat).new
      args.each do |arg|
        case arg
        when Array
          all_numbers.concat(extract_numbers(arg))
        else
          if num = to_float(arg)
            all_numbers << num
          end
        end
      end
      return div0 if all_numbers.empty?
      all_numbers.sum / all_numbers.size
    end

    # AVERAGE: Multiple arguments (arrays and/or scalars mixed)
    def self.average(*values : CellValue | Array(CellValue)) : CellValue
      average_impl(*values)
    end

    # AVERAGE: Multiple arrays only
    def self.average(*values : Array(CellValue)) : CellValue
      all_numbers = Array(BigFloat).new
      values.each do |arr|
        all_numbers.concat(extract_numbers(arr))
      end
      return div0 if all_numbers.empty?
      all_numbers.sum / all_numbers.size
    end

    # AVERAGE: Single array overload for backward compatibility
    def self.average(values : Array(CellValue)) : CellValue
      numbers = extract_numbers(values)
      return div0 if numbers.empty?
      numbers.sum / numbers.size
    end

    # MIN: Generic version that handles mixed types
    private def self.min_impl(*args : CellValue | Array(CellValue)) : CellValue
      all_numbers = Array(BigFloat).new
      args.each do |arg|
        case arg
        when Array
          all_numbers.concat(extract_numbers(arg))
        else
          if num = to_float(arg)
            all_numbers << num
          end
        end
      end
      return num if all_numbers.empty?
      all_numbers.min
    end

    # MIN: Multiple arguments (arrays and/or scalars mixed)
    def self.min(*values : CellValue | Array(CellValue)) : CellValue
      min_impl(*values)
    end

    # MIN: Multiple arrays only
    def self.min(*values : Array(CellValue)) : CellValue
      all_numbers = Array(BigFloat).new
      values.each do |arr|
        all_numbers.concat(extract_numbers(arr))
      end
      return num if all_numbers.empty?
      all_numbers.min
    end

    # MIN: Single array overload for backward compatibility
    def self.min(values : Array(CellValue)) : CellValue
      numbers = extract_numbers(values)
      return num if numbers.empty?
      numbers.min
    end

    # MAX: Generic version that handles mixed types
    private def self.max_impl(*args : CellValue | Array(CellValue)) : CellValue
      all_numbers = Array(BigFloat).new
      args.each do |arg|
        case arg
        when Array
          all_numbers.concat(extract_numbers(arg))
        else
          if num = to_float(arg)
            all_numbers << num
          end
        end
      end
      return num if all_numbers.empty?
      all_numbers.max
    end

    # MAX: Multiple arguments (arrays and/or scalars mixed)
    def self.max(*values : CellValue | Array(CellValue)) : CellValue
      max_impl(*values)
    end

    # MAX: Multiple arrays only
    def self.max(*values : Array(CellValue)) : CellValue
      all_numbers = Array(BigFloat).new
      values.each do |arr|
        all_numbers.concat(extract_numbers(arr))
      end
      return num if all_numbers.empty?
      all_numbers.max
    end

    # MAX: Single array overload for backward compatibility
    def self.max(values : Array(CellValue)) : CellValue
      numbers = extract_numbers(values)
      return num if numbers.empty?
      numbers.max
    end

    # COUNT: Generic version that handles mixed types
    private def self.count_impl(*args : CellValue | Array(CellValue)) : BigFloat
      count = 0
      args.each do |arg|
        case arg
        when Array
          count += extract_numbers(arg).size
        else
          if to_float(arg)
            count += 1
          end
        end
      end
      count.to_f
    end

    # COUNT: Multiple arguments (arrays and/or scalars mixed)
    def self.count(*values : CellValue | Array(CellValue)) : CellValue
      count_impl(*values)
    end

    # COUNT: Multiple arrays only
    def self.count(*values : Array(CellValue)) : CellValue
      count = 0
      values.each do |arr|
        count += extract_numbers(arr).size
      end
      count.to_f
    end

    # COUNT: Single array overload for backward compatibility
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
      when BigFloat then value != BigFloat.new(0.0, precision: DEFAULT_PRECISION)
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
        when BigFloat then v != BigFloat.new(0.0, precision: DEFAULT_PRECISION)
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
        when BigFloat then v != BigFloat.new(0.0, precision: DEFAULT_PRECISION)
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

    # CONCAT: Joins several text strings into one text string (supports multiple arrays)
    def self.concat(*values : Array(CellValue)) : CellValue
      values.map { |arr| arr.map { |v| to_string(v) } }.flatten.join
    end

    # CONCAT: Single array overload for backward compatibility
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
      when {BigFloat, BigFloat}
        left <=> right
      when {BigFloat, String}
        -1 # Numbers are always less than strings in Excel
      when {String, BigFloat}
        1
      when {String, String}
        left <=> right
      when {Bool, Bool}
        (left ? 1 : 0) <=> (right ? 1 : 0)
      else
        0
      end
    end

    # Additional statistical functions

    # COUNTA: Counts how many values are in the list of arguments (non-empty)
    def self.counta(values : Array(CellValue)) : CellValue
      values.reject do |v|
        v.nil? || (v.is_a?(String) && v.empty?)
      end.size.to_f
    end

    # MEDIAN: Returns the median of the given numbers
    def self.median(values : Array(CellValue)) : CellValue
      numbers = extract_numbers(values)
      return num if numbers.empty?
      return numbers[0] if numbers.size == 1

      sorted = numbers.sort
      mid = sorted.size // 2

      if sorted.size.odd?
        sorted[mid]
      else
        (sorted[mid - 1] + sorted[mid]) / 2.0
      end
    end

    # STDEV.S: Estimates standard deviation based on a sample
    def self.stdev(values : Array(CellValue)) : CellValue
      numbers = extract_numbers(values)
      return div0 if numbers.size < 2
      return 0.0 if numbers.size == 1

      mean = numbers.sum / numbers.size
      variance = numbers.sum { |n| (n - mean) ** 2 } / (numbers.size - 1)
      Math.sqrt(variance)
    end

    # STDEV.P: Calculates standard deviation based on entire population
    def self.stdev_p(values : Array(CellValue)) : CellValue
      numbers = extract_numbers(values)
      return div0 if numbers.empty?

      mean = numbers.sum / numbers.size
      variance = numbers.sum { |n| (n - mean) ** 2 } / numbers.size
      Math.sqrt(variance)
    end

    # VAR.S: Estimates variance based on a sample
    def self.var_s(values : Array(CellValue)) : CellValue
      numbers = extract_numbers(values)
      return div0 if numbers.size < 2

      mean = numbers.sum / numbers.size
      numbers.sum { |n| (n - mean) ** 2 } / (numbers.size - 1)
    end

    # VAR.P: Calculates variance based on entire population
    def self.var_p(values : Array(CellValue)) : CellValue
      numbers = extract_numbers(values)
      return div0 if numbers.empty?

      mean = numbers.sum / numbers.size
      numbers.sum { |n| (n - mean) ** 2 } / numbers.size
    end

    # Additional math functions

    # CEILING: Rounds number up to nearest multiple of significance
    def self.ceiling(number : CellValue, significance : CellValue = 1.0) : CellValue
      num = to_float(number)
      sig = to_float(significance) || 1.0
      return number if num.nil?
      return div0 if sig == 0
      return num if sig < 0

      (num / sig).ceil * sig
    end

    # FLOOR: Rounds number down to nearest multiple of significance
    def self.floor(number : CellValue, significance : CellValue = 1.0) : CellValue
      num = to_float(number)
      sig = to_float(significance) || 1.0
      return number if num.nil?
      return div0 if sig == 0
      return num if sig < 0

      (num / sig).floor.to_f * sig
    end

    # ROUNDUP: Rounds number up away from zero
    def self.roundup(number : CellValue, digits : CellValue = 0.0) : CellValue
      num = to_float(number)
      d = to_float(digits) || 0.0
      return number if num.nil?

      multiplier = 10.0 ** d
      if num >= 0
        (num * multiplier).ceil.to_f / multiplier
      else
        (num * multiplier).floor.to_f / multiplier
      end
    end

    # ROUNDDOWN: Rounds number down toward zero
    def self.rounddown(number : CellValue, digits : CellValue = 0.0) : CellValue
      num = to_float(number)
      d = to_float(digits) || 0.0
      return number if num.nil?

      multiplier = 10.0 ** d
      if num >= 0
        (num * multiplier).floor.to_f / multiplier
      else
        (num * multiplier).ceil.to_f / multiplier
      end
    end

    # RAND: Returns a random number between 0 and 1
    def self.rand : CellValue
      ::Random.rand.to_f
    end

    # RANDBETWEEN: Returns random integer between two numbers
    def self.randbetween(bottom : CellValue, top : CellValue) : CellValue
      b = to_float(bottom)
      t = to_float(top)
      return num if b.nil? || t.nil?
      return num if b > t

      ::Random.rand(b.to_i..t.to_i).to_f
    end

    # Additional text functions

    # FIND: Returns starting position of one text string within another (case-sensitive)
    def self.find(find_text : CellValue, within_text : CellValue, start_num : CellValue = 1.0) : CellValue
      find = to_string(find_text)
      within = to_string(within_text)
      start = to_float(start_num) || 1.0

      return value if find.empty?
      return num if start < 1 || start > within.size

      pos = within[(start.to_i - 1)..].index(find)
      return num if pos.nil?

      (pos + start.to_i).to_f
    end

    # SEARCH: Returns position of one text string within another (case-insensitive)
    def self.search(find_text : CellValue, within_text : CellValue, start_num : CellValue = 1.0) : CellValue
      find = to_string(find_text).downcase
      within = to_string(within_text).downcase
      start = to_float(start_num) || 1.0

      return value if find.empty?
      return num if start < 1 || start > within.size

      pos = within[(start.to_i - 1)..].index(find)
      return num if pos.nil?

      (pos + start.to_i).to_f
    end

    # SUBSTITUTE: Replaces existing text with new text
    def self.substitute(text : CellValue, old_text : CellValue, new_text : CellValue, instance_num : CellValue? = nil) : CellValue
      txt = to_string(text)
      old = to_string(old_text)
      new = to_string(new_text)

      return txt if old.empty?

      if instance_num
        inst = to_float(instance_num)
        return txt if inst.nil? || inst < 1

        # Replace only nth instance
        count = 0
        txt.gsub(old) do |match|
          count += 1
          count == inst.to_i ? new : match
        end
      else
        # Replace all instances
        txt.gsub(old, new)
      end
    end

    # TEXT: Formats a number and converts to text
    def self.text_func(value : CellValue, format_text : CellValue) : CellValue
      num = to_float(value)
      fmt = to_string(format_text)

      return to_string(value) if num.nil?

      # Basic format codes support
      case fmt
      when /^0+$/, /^0+\.0+$/
        # Decimal format
        decimal_places = fmt.count('0') - (fmt.index('.') || fmt.size)
        num.round(decimal_places).to_s
      when /#,##0/
        # Thousands separator
        num.to_s.reverse.gsub(/(\d{3})(?=\d)/, "\\1,").reverse
      when /^%$/
        "#{(num * 100).round(0).to_i}%"
      else
        # Default to number string
        num.to_s
      end
    end

    # VALUE: Converts text to number
    def self.value_func(text : CellValue) : CellValue
      txt = to_string(text).strip

      # Try to convert to float
      begin
        txt.to_f
      rescue
        value
      end
    end

    # PROPER: Capitalizes first letter of each word
    def self.proper(text : CellValue) : CellValue
      txt = to_string(text)
      return "" if txt.empty?

      txt.split(' ').map do |word|
        next word if word.empty?
        word[0].upcase + word[1..].downcase
      end.join(' ')
    end

    # CLEAN: Removes non-printable characters
    def self.clean(text : CellValue) : CellValue
      txt = to_string(text)
      # Remove ASCII control characters (0-31, except 9, 10, 13)
      txt.gsub(/[\x00-\x08\x0B\x0C\x0E-\x1F]/, "")
    end

    # EXACT: Compares two text strings (case-sensitive)
    def self.exact(text1 : CellValue, text2 : CellValue) : CellValue
      to_string(text1) == to_string(text2)
    end

    # REPT: Repeats text given number of times
    def self.rept(text : CellValue, number_times : CellValue) : CellValue
      txt = to_string(text)
      num = to_float(number_times)
      return num if num.nil?
      return value if num < 0

      txt * num.to_i
    end

    # Date and time functions

    # TODAY: Returns current date as serial number
    def self.today : CellValue
      # Excel date system: days since January 1, 1900
      epoch = Time.utc(1900, 1, 1)
      seconds = (Time.utc - epoch).total_seconds.to_i64
      days = (seconds // 86400).to_i
      (days + 2).to_f # Excel's 1900 date system has a bug treating 1900 as leap year
    end

    # NOW: Returns current date and time as serial number
    def self.now : CellValue
      # Excel datetime: fractional part represents time
      epoch = Time.utc(1900, 1, 1)
      seconds = (Time.utc - epoch).total_seconds.to_i64
      days = (seconds // 86400).to_i
      # Add fractional time component
      seconds_today = Time.utc.hour * 3600 + Time.utc.minute * 60 + Time.utc.second
      fractional = seconds_today / 86400.0

      (days + 2).to_f + fractional
    end

    # YEAR: Extracts year from date serial number
    def self.year(serial_number : CellValue) : CellValue
      num = to_float(serial_number)
      return value if num.nil?

      epoch = Time.utc(1900, 1, 1)
      date = epoch + Time::Span.new(seconds: (num.to_i64 - 2) * 86400)
      date.year.to_f
    end

    # MONTH: Extracts month from date serial number
    def self.month(serial_number : CellValue) : CellValue
      num = to_float(serial_number)
      return value if num.nil?

      epoch = Time.utc(1900, 1, 1)
      date = epoch + Time::Span.new(seconds: (num.to_i64 - 2) * 86400)
      date.month.to_f
    end

    # DAY: Extracts day from date serial number
    def self.day(serial_number : CellValue) : CellValue
      num = to_float(serial_number)
      return value if num.nil?

      epoch = Time.utc(1900, 1, 1)
      date = epoch + Time::Span.new(seconds: (num.to_i64 - 2) * 86400)
      date.day.to_f
    end

    # DATEDIF: Calculates difference between two dates
    def self.datedif(start_date : CellValue, end_date : CellValue, unit : CellValue) : CellValue
      start = to_float(start_date)
      end_d = to_float(end_date)
      u = to_string(unit)

      return value if start.nil? || end_d.nil?

      start_epoch = Time.utc(1900, 1, 1) + Time::Span.new(seconds: (start.to_i64 - 2) * 86400)
      end_epoch = Time.utc(1900, 1, 1) + Time::Span.new(seconds: (end_d.to_i64 - 2) * 86400)

      case u.upcase
      when "Y"
        # Years
        years = end_epoch.year - start_epoch.year
        years -= 1 if end_epoch < TimeHelpers.add_years(start_epoch, years.to_i)
        years.to_f
      when "M"
        # Months
        months = (end_epoch.year - start_epoch.year) * 12 + (end_epoch.month - start_epoch.month)
        months -= 1 if end_epoch < TimeHelpers.add_months(start_epoch, months.to_i)
        months.to_f
      when "D"
        # Days
        (end_d - start).abs.to_f
      when "MD"
        # Days ignoring months and years
        day_diff = end_epoch.day - start_epoch.day
        day_diff += 30 if day_diff < 0
        day_diff.to_f
      when "YM"
        # Months ignoring years
        month_diff = end_epoch.month - start_epoch.month
        month_diff += 12 if month_diff < 0
        month_diff.to_f
      when "YD"
        # Days ignoring years
        days_diff = (end_epoch.day_of_year - start_epoch.day_of_year)
        days_diff += 365 if days_diff < 0
        days_diff.to_f
      else
        value
      end
    end

    # EOMONTH: Returns last day of month offset from date
    def self.eomonth(start_date : CellValue, months : CellValue = 0.0) : CellValue
      start = to_float(start_date)
      m = to_float(months) || 0.0
      return value if start.nil?

      epoch = Time.utc(1900, 1, 1) + Time::Span.new(seconds: (start.to_i64 - 2) * 86400)
      target = TimeHelpers.add_months(epoch, m.to_i)

      # Find last day of target month
      last_day = utc_days_in_month(target.year, target.month)
      end_of_month = Time.utc(target.year, target.month, last_day)

      # Convert back to serial number
      seconds = (end_of_month - Time.utc(1900, 1, 1)).total_seconds.to_i64
      days = (seconds // 86400).to_i
      (days + 2).to_f
    end

    # Conditional functions

    # IFS: Evaluates multiple conditions and returns value for first true condition
    def self.ifs(pairs : Array(CellValue)) : CellValue
      # Pairs should be [condition1, value1, condition2, value2, ...]
      return value if pairs.size < 2
      return value if pairs.size.odd?

      pairs.each_slice(2) do |slice|
        condition = to_bool(slice[0])
        next if condition.nil?

        if condition
          return slice[1]
        end
      end

      na # No condition was true
    end

    # SWITCH: Evaluates value against list and returns matching result
    def self.switch_func(expression : CellValue, pairs : Array(CellValue), default : CellValue? = nil) : CellValue
      # Pairs should be [value1, result1, value2, result2, ...]
      return expression if pairs.size < 2
      return expression if pairs.size.odd?

      pairs.each_slice(2) do |slice|
        if compare_values(expression, slice[0]) == 0
          return slice[1]
        end
      end

      default || na
    end

    # Conditional aggregation functions

    # COUNTIF: Counts cells meeting condition
    def self.countif(range : Array(CellValue), criteria : CellValue) : CellValue
      count = 0
      criterion = to_string(criteria)

      range.each do |value|
        if matches_criteria?(value, criterion)
          count += 1
        end
      end

      count.to_f
    end

    # SUMIF: Sums cells meeting condition
    def self.sumif(range : Array(CellValue), criteria : CellValue, sum_range : Array(CellValue)? = nil) : CellValue
      sum_range ||= range
      total = 0.0
      criterion = to_string(criteria)

      range.each_with_index do |value, idx|
        if matches_criteria?(value, criterion) && idx < sum_range.size
          num = to_float(sum_range[idx])
          total += num if num
        end
      end

      total
    end

    # Helper to check if value matches Excel criteria
    private def self.matches_criteria?(value : CellValue, criterion : String) : Bool
      # Handle operators
      if criterion =~ /^([<>]=?|>=|<=|=|<>)(.*)$/
        operator = $1
        crit_value = $2

        case value
        when BigFloat, String
          # Try to convert both to numbers for comparison
          val_num = to_float(value)
          crit_num = to_float(crit_value)

          if val_num && crit_num
            case operator
            when ">"  then val_num > crit_num
            when "<"  then val_num < crit_num
            when ">=" then val_num >= crit_num
            when "<=" then val_num <= crit_num
            when "="  then val_num == crit_num
            when "<>" then val_num != crit_num
            else           false
            end
          else
            # String comparison
            val_str = to_string(value)
            case operator
            when "="  then val_str == crit_value
            when "<>" then val_str != crit_value
            else           false
            end
          end
        else
          false
        end
      else
        # Simple equality or wildcard match
        val_str = to_string(value)

        if criterion.includes?('*') || criterion.includes?('?')
          # Wildcard match
          regex_str = "^#{Regex.escape(criterion).gsub("\\*", ".*").gsub("\\?", ".")}$"
          !(val_str =~ Regex.new(regex_str, Regex::Options::IGNORE_CASE)).nil?
        else
          # Case-insensitive equality
          val_str.downcase == criterion.downcase
        end
      end
    end

    # Lookup functions (basic implementation for single-column lookups)

    # VLOOKUP: Vertical lookup
    def self.vlookup(lookup_value : CellValue, table_array : Array(Array(CellValue)), col_index_num : CellValue, range_lookup : CellValue? = true) : CellValue
      col_idx = to_float(col_index_num)
      return value if col_idx.nil? || col_idx < 1

      exact_match = range_lookup == false || to_bool(range_lookup) == false

      table_array.each do |row|
        next if row.empty?

        compare_result = compare_values(lookup_value, row[0])
        next if compare_result.nil?

        if exact_match
          if compare_result == 0 && (col_idx.to_i - 1) < row.size
            return row[col_idx.to_i - 1]
          end
        else
          # Approximate match (table must be sorted)
          if compare_result <= 0 && (col_idx.to_i - 1) < row.size
            return row[col_idx.to_i - 1]
          end
        end
      end

      if exact_match
        na
      else
        value
      end
    end

    # HLOOKUP: Horizontal lookup
    def self.hlookup(lookup_value : CellValue, table_array : Array(Array(CellValue)), row_index_num : CellValue, range_lookup : CellValue? = true) : CellValue
      row_idx = to_float(row_index_num)
      return value if row_idx.nil? || row_idx < 1

      exact_match = range_lookup == false || to_bool(range_lookup) == false

      # Find matching column in first row
      match_col = -1

      if table_array.empty? || table_array[0].empty?
        return value
      end

      first_row = table_array[0]

      first_row.each_with_index do |cell, col|
        compare_result = compare_values(lookup_value, cell)
        next if compare_result.nil?

        if exact_match
          if compare_result == 0
            match_col = col
            break
          end
        elsif compare_result <= 0
          match_col = col
          break
        end
      end

      return na if exact_match && match_col == -1
      return value if match_col == -1

      # Return value from specified row
      target_row = row_idx.to_i - 1
      return ref if target_row >= table_array.size

      table_array[target_row][match_col]?
    end

    # INDEX: Returns value from array at given position
    def self.index_func(array : Array(Array(CellValue)), row_num : CellValue, column_num : CellValue) : CellValue
      row = to_float(row_num)
      col = to_float(column_num)
      return value if row.nil? || col.nil? || row < 1 || col < 1

      target_row = row.to_i - 1
      target_col = col.to_i - 1

      return ref if target_row >= array.size
      return ref if target_col >= array[target_row].size

      array[target_row][target_col]? || ref
    end
  end
end
