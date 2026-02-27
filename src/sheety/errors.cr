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
end
