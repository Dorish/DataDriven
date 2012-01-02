# Rename the ruby scripts / spreadsheets to include the Testlog test case number
#
# A ref csv file needs to be created: (create a spreadsheet and save as csv)
#  Col A shall contain the test case number
#  Col B shall contain the original script name
#
# Each iteration of the do loop will:
#  1) Change the script name
#  2) Print old name and new name to console
#
# For simplicity, copy this file and the ref csv file into the target directory
# Remove this script and the ref csv file from the target directory when done
#
# chomp is used below to remove the "\n" newline character from the strings
#

Dir.chdir(File.dirname(__FILE__))     # change to the directory of this file

open('telnet_script_rename.csv').each do |line|
  old = line.split(',')[1].chomp   # old file name is the second element
  new = line.gsub(',','-').chomp   # join Testlog no. and script with hyphen
  File.rename(old,new)             # rename file
  puts old + " | " + new
end

puts "\n Telnet script rename finished\n\n"
puts " To see the renamed files, refresh the project view in NetBeans:"
puts " Select 'Source' in the menu bar and click 'Scan for External Changes'"