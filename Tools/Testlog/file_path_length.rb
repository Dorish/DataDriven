=begin

Recursive directory listing of files that have an absolute file length
greater than the number specified in "if f.length > n"

show the contents of the current directory and all its subdirectories
    require 'find'
    Find.find('./')do |f| p f end

Or can use
    Find.find(',/') do |file|
      print file
    end
************************************
This script script was written to show the file path lengths in Testlog.
Expectation is that the script will be ran in the Testlog folder.

=end


require 'find'
Find.find('./')do |f| 
  if f.length > 200  # adjust to display file lengths > n
    print f    # file order is from top to bottom
    puts "(#{f.length})" # print the actual path length in parenthesis
  end
end


