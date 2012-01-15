=begin
  Write ruby script filenames to a .csv file
  The output here is intended to be used for the script rename utilities
=end

Dir.chdir(File.dirname(__FILE__))         # change to directory of this file
File.open('telnet_script_rename.csv', 'w+') do |f|      # open new csv file
  Dir.open(Dir.pwd).each do |fname|       # open directory and start iterating
  f.puts ',' + fname if fname =~ /tn.xls/ # get telnet script files only
  end                                     # Insert comma to create Col A
end

puts "telnet_script_rename.csv - file created \n\n"
puts " To see the new .csv file, refresh the project view in NetBeans as follows:"
puts " Select 'Source' in the menu bar and click 'Scan for External Changes'\n\n"
puts " Edit telnet_script_rename.csv by adding the Testlog case number to Col A"
puts " Now, run telnet_script_rename.rb to complete the file name conversion "
