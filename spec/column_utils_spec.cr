require "spec"
require "../src/sheety/cell_refs"

describe Sheety::CellRefs do
  describe ".col_to_num" do
    it "converts single-letter columns" do
      Sheety::CellRefs.col_to_num("A").should eq(1)
      Sheety::CellRefs.col_to_num("B").should eq(2)
      Sheety::CellRefs.col_to_num("Z").should eq(26)
    end

    it "converts multi-letter columns" do
      Sheety::CellRefs.col_to_num("AA").should eq(27)
      Sheety::CellRefs.col_to_num("AB").should eq(28)
      # XFD is Excel's last column
      Sheety::CellRefs.col_to_num("XFD").should eq(16384)
    end

    it "upcases its input" do
      Sheety::CellRefs.col_to_num("a").should eq(1)
      Sheety::CellRefs.col_to_num("ab").should eq(28)
    end
  end

  describe ".num_to_col" do
    it "converts numbers to letters" do
      Sheety::CellRefs.num_to_col(1).should eq("A")
      Sheety::CellRefs.num_to_col(26).should eq("Z")
      Sheety::CellRefs.num_to_col(27).should eq("AA")
      Sheety::CellRefs.num_to_col(16384).should eq("XFD")
    end

    it "round-trips with col_to_num" do
      (1..703).each do |num|
        Sheety::CellRefs.col_to_num(Sheety::CellRefs.num_to_col(num)).should eq(num)
      end
    end

    it "returns an empty string for zero" do
      Sheety::CellRefs.num_to_col(0).should eq("")
    end
  end
end
