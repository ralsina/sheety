require "./spec_helper"
require "../src/sheety"
require "docopt-config"

describe Sheety::CLI do
  describe ".parse_args" do
    it "parses a bare filename" do
      options = Sheety::CLI.parse_args(["sheet.yaml"])
      options["<file>"].should eq("sheet.yaml")
      options["--save-to"]?.should be_nil
    end

    it "accepts --save-to in any position" do
      after = Sheety::CLI.parse_args(["sheet.yaml", "--save-to=out.xlsx"])
      before = Sheety::CLI.parse_args(["--save-to=out.xlsx", "sheet.yaml"])

      after["<file>"].should eq("sheet.yaml")
      after["--save-to"]?.should eq("out.xlsx")
      before["<file>"].should eq("sheet.yaml")
      before["--save-to"]?.should eq("out.xlsx")
    end

    it "raises a usage error for unknown flags" do
      expect_raises(Docopt::DocoptExit) do
        Sheety::CLI.parse_args(["--frobnicate", "sheet.yaml"])
      end
    end

    it "raises a usage error for a missing file argument" do
      expect_raises(Docopt::DocoptExit) do
        Sheety::CLI.parse_args(["--save-to=out.xlsx"])
      end
    end

    it "raises ConfigExit for --help and --version" do
      expect_raises(Docopt::ConfigExit) do
        Sheety::CLI.parse_args(["--help"])
      end
      expect_raises(Docopt::ConfigExit) do
        Sheety::CLI.parse_args(["-h"])
      end
      expect_raises(Docopt::ConfigExit) do
        Sheety::CLI.parse_args(["--version"])
      end
    end
  end
end
