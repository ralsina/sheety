require "./spec_helper"
require "../src/sheety"

describe Sheety::FormulaExtractor do
  describe ".translate_shared" do
    it "offsets relative references by the cell delta" do
      # Master C1 holds =B1*2; slave D2 is one right, one down.
      Sheety::FormulaExtractor.translate_shared("B1*2", "C1", "D2").should eq("C2*2")
    end

    it "leaves $-anchored parts alone" do
      Sheety::FormulaExtractor.translate_shared("$B$1+B2", "C1", "D2").should eq("$B$1+C3")
      Sheety::FormulaExtractor.translate_shared("$A1+B$1", "A1", "B1").should eq("$A1+C$1")
    end

    it "translates both endpoints of a range" do
      Sheety::FormulaExtractor.translate_shared("SUM(A1:B2)", "A1", "B1").should eq("SUM(B1:C2)")
    end

    it "does not touch text inside string literals" do
      Sheety::FormulaExtractor.translate_shared("CONCAT(\"A1\", A1)", "A1", "B1").should eq("CONCAT(\"A1\", B1)")
    end

    it "does not mistake function names for references" do
      Sheety::FormulaExtractor.translate_shared("LOG10(A1)+1", "A1", "B2").should eq("LOG10(B2)+1")
    end

    it "keeps sheet names intact while shifting the reference after the bang" do
      Sheety::FormulaExtractor.translate_shared("Sheet2!A1", "A1", "B1").should eq("Sheet2!B1")
    end

    it "emits #REF! for references shifted off the sheet" do
      Sheety::FormulaExtractor.translate_shared("A1+1", "B1", "A1").should eq("#REF!+1")
      Sheety::FormulaExtractor.translate_shared("A1", "A2", "A1").should eq("#REF!")
    end

    it "returns the master unchanged for zero delta" do
      Sheety::FormulaExtractor.translate_shared("B1*2", "C1", "C1").should eq("B1*2")
    end
  end
end
