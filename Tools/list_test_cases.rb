=begin
#------------------------------------------------------------------------------#
This script will recursively return a list of test cases from Testlog

The specific DIRECTORY path is determined in "dir = []. All test cases in that
directory and sub-directories below will be returned along with the Title of the
test case "test_title"

The Testlog ".tlg" files are .xml format.  This script can be modified to
extract any information desired by specifying the desired xml xpath

Example:
    puts xml.xpath("//test_title").text
#------------------------------------------------------------------------------#
=end

require 'find'
require 'nokogiri'


dirs = ['I:\lmg test engineering\TestLog\Test Cases\700. Configure']
excludes = []

for dir in dirs
  Find.find(dir) do |fname|     # fname =
    if FileTest.directory?(fname)
      if excludes.include?(File.basename(fname))
        Find.prune       # Don't look any further into this directory.
      else
        next
      end
    else
      if fname =~ /\.tlg/
        print File.basename(fname).sub('.tlg','') + ','
        xml = Nokogiri::XML(File.open(fname))
        puts xml.xpath("//test_title").text # test case title
      end
    end
  end
end