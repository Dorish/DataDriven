=begin

(EVENTS col, sheet 1)
'A' GDD ID
'L' GDD Label
'B' LIFE ID
'C' Life Event
'F' EmerDelay
'G' InReport
'H' Emergen
'I' Othercall
'J' AlmCond

The LIFE Event Test Case format is <space><hyphen><space> separated as follows:
LF - GDD ID - GDD label - LIFE ID - LIFE Event - EmerDelay - InReport - Emergen - OtherCall - AlmCond
Example: LF - 4168 - Battery Discharging - 308 - Battery discharging -  - Yes - Normal - Normal - event

--------------------------------------------------------------------------------

(MEASURES col, sheet 2)
'A' GDD ID
'N' GDD Label
'B' LIFE ID
'C' Life Measure
'J' Unit of Measure

Measures test case format is <space><hyphen><space> separated as follows:
LF - GDD ID - GDD label - LIFE ID - LIFE Measure - Unit of Measure
Example: LF - 4096 - System Input RMS A-N - 87 - Mains L1-N voltage  - VAC

=end

puts ''
#TODO Update the framework xls lib to support opening multiple work sheets in one session

#add framework library to load path. Needed for the 'require' statement below
$:.unshift File.expand_path("../../../lib", __FILE__)

require 'xls'
include Xls


# Select the LIFE mapping spreadsheet to read from list
def select_file_from_list(file_type)
  puts'Working directory = ' + Dir.pwd
  file_list = [nil]            # pre-load the array so file_counter can start at 1
  file_counter = 1             # initialize file counter

  Dir.glob('*.' + file_type).each do |f| # get the base filename  by file type
    puts "     #{file_counter}" + ' - ' + f
    file_list.push(File.expand_path(f)) # build absolute file path
    file_counter += 1
  end

  # Select the test file to display - show its absolute path
  puts "     Type the number of the desired file above followed by <enter>: "
  file_number = gets.to_i
  return file_list[file_number]  # 'file_list' is an array of available files
                                 # 'file_number' is the element for file selected
end


# Array containing event and measure test case prefixes
def test_case_prefixes
  puts "     Type the test case number prefix for events followed by <enter>:"
  puts "     example: 816.600.30 .....then first testcase will be 816.600.30.10"
  events_prefix = gets.to_s.chomp

  puts "     Type the test case number prefix for measures followed by <enter>:"
  puts "     example: 816.600.10 .....then first testcase will be 816.600.10.10"
  measures_prefix = gets.to_s.chomp

  return t_c_prefixes = [events_prefix,measures_prefix]
end


# Setup filename for output .csv file
def csv_output(input_file,type)
  input_file.gsub('.xlsx',"_#{type}.csv")
end


# Write the test case number and title to a .csv file
def file_write(file,test_suite)
  File.open(file, 'w+') do |f|
    test_suite.each{|x|f.puts x.to_s}  #.to_s converts each nested array to a string
  end
end


# Build a test suite by reading a tab from LIFE mapping spreadsheet
# The test suite is returned by test_suite
def build_test_cases(spreadsheet,columns,tc_prefix) 
  row = 2
  tc_suffix = 10
  ws = spreadsheet[2]
  test_suite = Array.new
  while(ws.Range("B#{row}")['Value'] != nil) # check if row is empty
    title = Array.new
    title.push tc_prefix + "." + tc_suffix.to_s + ","
    title.push "LF - "
    columns.each do |col|
      cell=(ws.Range("#{col}#{row}")['Value'])

      # If Float, convert to integer to remove decimal
      cell = cell.to_i.to_s if cell.class == Float

      # If cell is empty (nil), convert to empty string
      cell = "" if cell.nil?

      # Add cell content to test case title array
      title = title.push cell

      # Add hyphen delimeter unless column = "J" (last column)
      title = title.push(" - ") unless col == "J"

      # Fill GDD ID cell with 'none if spreadsheet cell is empty
      title[2] = 'none' if (ws.Range("A#{row}")['Value']).nil?
    end
    
    test_suite.push title         # push each test case to test suite   
    row += 1                      # increment row
    tc_suffix += 10               # increment test case number
  end
  test_suite.each{|x|puts x.to_s} # print test suite to console for debug
 end


begin
  event_col = ["A","L","B","C","F","G","H","I","J"] # Event data columns from sheet 1
  meas_col = ["A","N","B","C","J"]                  # Measure data columns from sheet 2
  
  Dir.chdir(File.expand_path("../../temp_files", __FILE__)) # Set working directory
  input_file = select_file_from_list('xlsx')        # Select the input file by type

  tc_prefix = test_case_prefixes()            # Array containing test case number prefixes

  # Create _events.csv file
  spreadsheet = new_xls(input_file,1)         # Open spreadsheet, sheet 1
  events_suite = build_test_cases(spreadsheet,event_col,tc_prefix[0])
  spreadsheet[1].close                        # Close the workbook
  events = csv_output(input_file,'_events' )  # Setup filename for events output
  file_write(events,events_suite)             # Write test cases to events output file

  # Create _measures.csv file
  spreadsheet = new_xls(input_file,2)             # Open spreadsheet, sheet 2
  measures_suite = build_test_cases(spreadsheet,meas_col,tc_prefix[1])
  spreadsheet[1].close                            # Close the workbook
  measures = csv_output(input_file,'_measures' )  # Setup filename for measures output
  file_write(measures,measures_suite)             # Write test cases to measures output file
end