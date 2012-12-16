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

#add framework library to load path. Needed for the 'require' statement below
$:.unshift File.expand_path("../../../lib", __FILE__)

require 'xls'
include Xls

def select_file_from_list(file_type)
  puts'Directory = ' + Dir.pwd
  file_list = [nil]            # pre-load the array so file_counter can start at 1
  file_counter = 1             # initialize file counter

  Dir.glob(file_type).each do |f| # get the base file names (without the path)
    puts "\n     #{file_counter}" + ' - ' + f
    file_list.push(File.expand_path(f)) # build absolute file path
    file_counter += 1
  end

  # Select the test file to display - show its absolute path
  puts "     Type the number of the desired file above followed by <enter>: "
  file_number = gets.to_i
  return file_list[file_number]
end


def file_write(file,line)
  File.open(file, 'a') do |f|
    f.puts line
  end
end


def build_test_case_title(spreadsheet,columns,fname) #build test case title
  row = 2
  ws = spreadsheet[2]
  
  while(ws.Range("B#{row}")['Value'] != nil) # check if row is empty
    title = Array.new
    title.push "LF - "
    columns.each do |col|
      cell=(ws.Range("#{col}#{row}")['Value'])

      # If Float, convert to integer to remove decimal
      cell = cell.to_i.to_s if cell.class == Float

      # If cell is empty (nil), convert to empty string
      cell = "" if cell.nil?

      # Add hyphen delimeter unless column = "J" (last column)
      cell = cell + " - " unless col == "J"

      title = title.push cell
    end
    file_write(fname,title.to_s)
    
    print title                  #test case title to console
    row = row + 1
    puts "\n"
  end
 end


begin
  event_col = ["A","L","B","C","F","G","H","I","J"] #Event data columns from sheet 1
  meas_col = ["A","N","B","C","J"]                  #Measure data columns from sheet 2

  Dir.chdir(File.expand_path("../../temp_files", __FILE__)) #Change working directory to temp_file

  desired_file = select_file_from_list('*.xlsx')

  events = desired_file.gsub('.xlsx','_events.csv')
  spreadsheet = new_xls(desired_file,1)   # Open sheet 1
  build_test_case_title(spreadsheet,event_col,events)
  spreadsheet[1].close                    # Close the workbook

  puts"*****************\n"

  measures = desired_file.gsub('.xlsx','_measures.csv')
  spreadsheet = new_xls(desired_file,2)   # Open sheet 2
  build_test_case_title(spreadsheet,meas_col,measures)
  spreadsheet[1].close                    # Close the workbook
end