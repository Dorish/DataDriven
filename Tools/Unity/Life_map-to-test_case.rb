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

#add private lib to load path for the 'require' statement below
$:.unshift File.expand_path("../../../lib", __FILE__) 

require 'xls'
include Xls

def tn_close(telnet,navigate_str)
  esc = "\x1b"
  navigation = navigate_str.split(',')
  esc_num = navigation.length + 2 #need to press esc 2 times more from main menu to exit telnet session
  esc_num.times do
    telnet.write(esc) {|c| print c}
  end
end


def get_test_title_data(spreadsheet,columns) #build test case title
  row = 2
  ws = spreadsheet[2]
  while(ws.Range("B#{row}")['Value'] != nil) # check if row is empty
    title = Array.new
    title.push "LF - "
    columns.each do |col|
      cell=(ws.Range("#{col}#{row}")['Value'])

      if cell.class == Float
        cell = cell.to_i.to_s   # If Float, convert to integer to loose decimal
      elsif cell.nil?
        cell = ""               # If cell is empty (nil), convert to string
      end

      cell = cell + " - " unless col == "J" # Add hyphen delimeter unless last column
        title = title.push cell
    end

    print title                  #test case title
    row = row + 1
    puts "\n"
  end
 end

begin
  event_col = ["A","L","B","C","F","G","H","I","J"]#Event data columns from sheet 1
  meas_col = ["A","N","B","C","J"]#Measure data columns from sheet 2
  
  excel_name = File.dirname(__FILE__) + '\\' + 'NXr_866-events-meas_11-30-2012.xlsx' # use this path when we run the script in ruby environment
  
  #excel_name = Dir.pwd + '/' + 'NXr_866-events-meas_11-30-2012.xlsx' # use this path when we create the executable file use Exerb
  spreadsheet = new_xls(excel_name,1)# Open sheet 1
  get_test_title_data(spreadsheet,event_col)
  spreadsheet[1].close  # Close the workbook

  puts"*****************\n"

  spreadsheet = new_xls(excel_name,2)  # Open sheet 2
  get_test_title_data(spreadsheet,meas_col)
end