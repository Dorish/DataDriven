
=begin

This script parses log files and provides output in csv format.  It uses two 
input files:
1) parse_key.txt - contains a list of parameters that are used to parse the log 
2) (example)  log file is: 10.100.24.19_08-23_06-51-34.txt  contains many log 
   records

The output file produced is:
1) (example) csv file is: 10.100.24.19_08-23_06-51-34.csv  has a csv line for
   each log record

Note: All files in the directory that meet the log file naming format "10...txt"
are processed to provide a like named .csv output file

=end

Dir.chdir(File.dirname(__FILE__))      # change to directory of this file
parse_key = 'parse_unity_key.txt'            # file that contains items to parse from log
keyword = []
open(parse_key).each {|line| keyword.push(line.chop)}   # get keywords from each file line
                                                      # strip "\n" with chop before pushing to array
puts heading = keyword.join(',')       # headings for the csv csv_line file
csv_line  = []                         # initialize csv_line array(csv_line is one written to csv file)

Dir.open(Dir.pwd).each do |fname|      # read log file directory
  if fname =~ /^10.*unity.txt/               # parse each txt file that starts with "10"
    infile = fname
    puts "processing - #{infile}"
    outfile = infile.gsub('.txt', '.csv')  # create csv file name based on log txt file name

    File.open(outfile, 'w') do |f|      # open csv file
      f.puts heading                    # write heading to csv file
      open(infile).each do|line|        # read input file lines
        key_index = 0
        keyword.each do |key|           # start checking input file line for a match from match.txt file
          regexp = /\d{4}[-\d{2}]+T[\d{2}:]+.\d{3}\+[\d{2}:]+\s*#{key}\s*-*[\w]+\(\d{5}+\)[\s*:]+/
          if line =~ regexp          # get the lines in the log that match the one of the parse keys
            values = line.split(regexp)    # grab the value from the paramter:value pair that matched the parse key (after the ":")
            value = values[1].gsub(",", ".")     # logging information contains semicolon. It conflicts with csv format.
            for i in 0..keyword.length - 1
               if i == key_index
                 csv_line.push value.to_s.chomp  #add parsed value to the suitable position in a csv line.
               else
                 csv_line.push " "  # put empty information for unmatch keyword in a csv line.
               end
            end
            f.puts csv_line.join(',')  # build the csv line for values parsed from log record
            csv_line = []
          end
          key_index = key_index + 1
        end
      end
    end
  end
end

puts "----  Done  ----"