=begin
  Write ruby script filenames to a .csv file
  The output here is intended to be used for the script rename utilities

**Copy this file to the target directory for execution**

Note: The output file "watir_script_renemae_.csv" does not appear in the
      NetbBeans file structure quickly.  Use windows file system to access
=end

Dir.chdir(File.dirname(__FILE__))         # change to directory of this file
File.open('watir_script_rename.csv', 'w+') do |f|  # open new csv file
  Dir.open(Dir.pwd).each do |fname|       # open directory and start iterating
  f.puts "," + fname if fname =~ /web.rb/ # get ruby web script files only
  end                                     # Insert comma to create Col A
end

puts "\n Watir_script_rename.csv - file created"
puts " It may take some time for the csv file to be visible in NetBeans"
puts " Delete the tools form the folder"