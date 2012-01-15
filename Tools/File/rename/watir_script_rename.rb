# Rename the ruby scripts / spreadsheets to include the Testlog test case number
#
# A ref csv file needs to be created: (create a spreadsheet and save as csv)
#  Col A shall contain the test case number
#  Col B shall contain the original ruby script name
#  
#  Note - watir_script_rename.csv does not need to include the watir spreadsheet 
#  names
#
# Each iteration of the do loop will:
#  1) Change the ruby script name
#  2) Change the script spreadsheet name
#  3) Print old names and new names to console
#
# For simplicity, copy this file and the ref csv file into the target directory
# Remove this script and the ref csv file from the target directory when done
#
# chomp is used below to remove the "\n" newline character from the strings
#

Dir.chdir(File.dirname(__FILE__))     # change to the directory of this file

open('watir_script_rename.csv').each do |line|
  old_rb = line.split(',')[1].chomp   # old file name is the second element
  new_rb = line.gsub(',','-').chomp   # join Testlog no. and script with hyphen
  File.rename(old_rb,new_rb)          # rename ruby file
  old_xls = old_rb.gsub('rb','xls')   # substitute .rb with .xls
  new_xls = new_rb.gsub('rb','xls')   # substitute .rb with .xls
  File.rename(old_xls,new_xls)        # rename spreadsheet

  puts old_rb + " | " + new_rb
  puts old_xls + " | " + new_xls
end

puts "\n Watir script and spreadsheet rename finished\n\n"
puts " To see the renamed files, refresh the project view in NetBeans:"
puts " Select 'Source' in the menu bar and click 'Scan for External Changes'"