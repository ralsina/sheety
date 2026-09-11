require "spec"
require "croupier"
require "../src/sheety/croupier_helpers"

# croupier_helpers.cr is only ever compiled into generated binaries (nothing
# in src/ requires it), so without this file the Crystal compiler would never
# even type-check it as part of `crystal spec`.
describe Sheety::CroupierHelpers do
  before_each do
    {"Sheet1!A1", "Sheet1!B1", "Sheet1!A2", "Sheet1!B2"}.each do |key|
      Croupier::TaskManager.set(key, "")
    end
  end

  describe "#fetch_cell_range" do
    it "fetches a flat row-major array of values" do
      Croupier::TaskManager.set("Sheet1!A1", "1")
      Croupier::TaskManager.set("Sheet1!B1", "2")
      Croupier::TaskManager.set("Sheet1!A2", "3")
      Croupier::TaskManager.set("Sheet1!B2", "4")

      fetch_cell_range("Sheet1", "A", 1, "B", 2).should eq(["1", "2", "3", "4"])
    end

    it "defaults missing cells to empty strings" do
      fetch_cell_range("Sheet1", "A", 1, "A", 2).should eq(["", ""])
    end
  end

  describe "#fetch_cell_range_2d" do
    it "fetches one array of values per row" do
      Croupier::TaskManager.set("Sheet1!A1", "1")
      Croupier::TaskManager.set("Sheet1!B1", "2")
      Croupier::TaskManager.set("Sheet1!A2", "3")
      Croupier::TaskManager.set("Sheet1!B2", "4")

      fetch_cell_range_2d("Sheet1", "A", 1, "B", 2).should eq([["1", "2"], ["3", "4"]])
    end

    it "defaults missing cells to empty strings" do
      fetch_cell_range_2d("Sheet1", "A", 1, "B", 1).should eq([["", ""]])
    end
  end

  describe "#range_inputs" do
    it "builds the kv input keys for every cell in the range" do
      range_inputs("Sheet1", "A", 1, "B", 2).should eq(
        ["kv://Sheet1!A1", "kv://Sheet1!B1", "kv://Sheet1!A2", "kv://Sheet1!B2"])
    end
  end

  describe "#bin_div" do
    it "divides numeric values" do
      bin_div("6", "4").should eq("1.5")
    end

    it "returns #DIV/0! instead of raising on a zero divisor" do
      bin_div("6", "0").should eq("#DIV/0!")
    end

    it "returns nil for non-numeric operands" do
      bin_div("banana", "2").should be_nil
    end
  end

  describe "#bin_pow" do
    it "raises zero to a negative power as #DIV/0!, like Excel" do
      bin_pow("0", "-1").should eq("#DIV/0!")
    end

    it "computes integer and fractional powers" do
      bin_pow("2", "10").should eq("1024")
      bin_pow("9", "0.5").should eq("3")
    end
  end
end
