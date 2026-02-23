module Sheety
  # Base error class for all formula errors
  class FormulaError < Exception
    def initialize(message : String? = nil)
      msg = message || "Not a valid formula"
      super(msg)
    end
  end

  # Error raised when an invalid token is encountered
  class TokenError < FormulaError
    def initialize(message = "Invalid string")
      super(message)
    end
  end

  # Error raised when parentheses are mismatched
  class ParenthesesError < FormulaError
    def initialize(message = "Mismatched or misplaced parentheses!")
      super(message)
    end
  end

  # Error raised when a function is not implemented
  class FunctionError < FormulaError
    def initialize(message = "Function not implemented!")
      super(message)
    end
  end

  # Error raised when a range is invalid
  class InvalidRangeError < FormulaError
    def initialize(range_name : String)
      super("Invalid range #{range_name}!")
    end
  end

  # Error raised when a range has no value
  class RangeValueError < FormulaError
    def initialize(range_name : String)
      super("Range #{range_name} has no value!")
    end
  end
end
