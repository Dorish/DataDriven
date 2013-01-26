=begin
Below is an example changing datapoint val_sys_ep_inpRmsPhsAB mmdx 0 new value 456 for the NXr
4156,0,456
    where:
      datapoint = 4156
      mmdx = 0  (multi-module index)
      value = 456

sock.send("4156,0,456", 0, "127.0.0.1", port)


====================================================
Read spreadsheet parameter columns ["B","D","F","H"]

B = IP address of simulator (127.0.0.1 if local)
D = Email address
F = SMTP server fro email
H = Enable / Disable email


====================================================
Read spreadsheet data columns ["A","H","G","D","F","I","J","K"]

A = PointId (GGD ID)
H = Multi-Module Index
G = Value to be written to simulator
D = Life Event label
F = Emergency Delay
I = Optional step delay
J = Use Emergency Delay time - Yes or No (if yes, add to optional delay)
K = Run this step - Yes or No

Build an array that looks like this:
[["4122", "0", "19","Primary mains power failure","10", "Immediate", "y", "y"],
["4143", "0", "19","Bypass static switch fault / meas. defective", "0", nil, "n", "n"],
["4146", "0", "19","Primary mains wrong phase rotation","20", "1 minute", "y", "y"],
["4162", "0", "19","Battery Shutdown Imminent", "5", "10 minutes", "n", "y"]]


*Instructions*
Simple instructions:
1)	With the spreadsheet and script in the same folder, launch the script.
2)	The script will query user to select spreadsheet(s) to run.
3)	Enter how many time to repeat the scenario.

*Features*
If you want to run a quick test, either save as and delete some rows or mark the run flag as ‘n’. 

Has the following features implemented:
1)	Writes to V4 simulator via UDP  (remote control checkbox must be enabled on the simulator).
2)	Creates a select list for available spreadsheets in the same directory that the script is setting in. 
3)	The spreadsheet has an entry to configure the IP of the simulator (typically local host or a remote PC)
4)	Utilizes the existing EmerDelay in the LIFE mapping spreadsheets.
5)	Enable or disable EmerDelay for each data point.
6)	Optional (step) delay entry for each data point available . Current default is 10 seconds for each point.
7)	Enable or disable any data point (exclude from execution).
8)	Console log contains time stamp for each event start and end.
9)	The script will query the user for “how many loops” how many times to continuously repeat the scenario.

Features not implemented
1)	Email when test is complete or aborted (partially implemented but disabled).
2)	Log console to a date/time stamped .csv file.
3)	Rescue mechanism to ensure that email is sent.

=end

require 'win32ole'
require 'socket'
require 'net/smtp'
#
# - select the input spreadsheet to use
def select_file_from_list(file_type)
  puts"Scenario File Menu:"
  file_list = [nil]            # pre-load the array so file_counter can start at 1
  file_counter = 1             # initialize file counter

  Dir.glob('*.' + file_type).each do |f| # get the base filename  by file type
    puts "   #{file_counter}" + ' - ' + f
    file_list.push(File.expand_path(f)) # build absolute file path
    file_counter += 1
  end

  # Select the test file to display - show its absolute path
  print "\nPlease select file number then press <enter>: "
  file_number = gets.to_i
  puts "Executing: #{file_number} - " + File.basename(file_list[file_number]) + "\n\n"
  return file_list[file_number]  # 'file_list' is an array of available files
                                 # 'file_number' is the element for file selected
end


#
#  - create and return new instance of excel
def new_xls(s_s,num) #wb name and sheet number
  ss = WIN32OLE::new('excel.Application')
  wb = ss.Workbooks.Open(s_s)
  ws = wb.Worksheets(num)
  ss.visible = true # For debug
  xls = [ss,wb,ws]
end


#
# - translate emergency delay minutes and hours into seconds
def emergency_delay(value)
  case value
  when /Immediate/ then  x  = 0
  when nil then x = "next call"
  when /([0-9]+)\s*seconds/ then x = $1
  when /([0-9]+)\s*minute/ then x = $1.to_i * 60
  when  /([0-9]+)\s*hour/ then x= $1.to_i * 3600
  end
  return x
  
end


#
# - get simulator IP, email address, smtp server, email enable on first row
def script_parameters(spreadsheet,columns)
  row = 1                                           # read row 1
  parameters = []                                   
  ws = spreadsheet[2]
  columns.each do |col|
      cell = (ws.Range("#{col}#{row}")['Value'])

      # Add cell content parameters array
      parameters = parameters.push cell
    end
  return parameters
end


#
# - build a scenario (array) of simulator command strings
def build_sim_inputs(spreadsheet,columns)
  row = 3                                     # start reading on row 3
  ws = spreadsheet[2]
  scenario = []
  while(ws.Range("A#{row}")['Value'] != nil)  # stop if column "A" is empty
    command = []
    columns.each do |col|
      cell = (ws.Range("#{col}#{row}")['Value'])

      # If Float, convert to integer to remove decimal
      cell = cell.to_i.to_s if cell.class == Float

      # Add cell content to test case command array
      command = command.push cell
      #command = command.push
    end
      
    scenario.push command                 # push each command array onto scenario array
    row += 1                              # increment spreadsheet row
  end
  # scenario.each{|x|p x} # debug - print array of simulator command strings

  spreadsheet[1].save
  spreadsheet[0].quit
  return scenario
 end


#
# - time stamp file by inserting month-year_hour-minute-second before the dot
def time_stamped_file(file)
  file.gsub(/\./,"_" + Time.now.strftime("%m-%d_%H-%M-%S")+ '.')
end


#
# - write data to csv file
def write_csv(outfile,log)
  File.open(outfile, 'w') do |f|        # open csv file
    f.puts log                          # write heading to csv file
  end
end


#
# - time stamp in 'month-day_hour-minute-second' format
def t_stamp
Time.now.strftime("%m-%d_%H:%M:%S")
end


#
# - send email
def send_mail(from,to)
  message = <<MESSAGE_END
  From: Darryl Brown <darryl.brown@emerson.com>
  To: d-l-brown@roadrunner.com
  Subject: LIFE Scenario test

  The LIFE station has exploded.
MESSAGE_END

Net::SMTP.start('inetmail.emrsn.net') do |smtp|
  smtp.send_message message, "#{from}","#{to}"
  smtp.finish
  end
end



#TODO clean up send_email method, parameterize message body
#TODO add rescue mechanism to send email for successful completion or failure
#TODO log the console out to a time stamped .csv file based on spreadsheet name
#TODO cleanup write_csv method
#TODO the command array is mutable, don't mutate it.
#TODO 'map(&:dup)' method is not happy on 1.8.6
#TODO add life event label the the console log and output log "D"


s = Time.now
port = 47809                                # destination port2
sock = UDPSocket.new
script_log = []
loop_count = 1

from = "darryl.brown@emerson.com"
to =   "d-l-brown@roadrunner.com"
message = ""

sock.bind("", 47123)                        # source udp port


columns_1 = ["B","D","F","H"]               # script data, see header for details
columns_2 = ["A","H","G","D","I","F","J","K"]   # script data, see header for details

# set current location of this file as the working directory
Dir.chdir(File.dirname(__FILE__)) 

# show list of available spreadsheet for selection
puts driver_file = select_file_from_list('xlsx')

# Get number of times to loop. Default is 1.  Press enter to accept default
print "\nPlease enter the number of times to run the scenario <enter>: "
loop = gets.to_i + 1                        # loop count is initialized


# open spreadsheet
input_file = new_xls(driver_file,1)

parameters = script_parameters(input_file,columns_1) # array of script parameters in spreadsheet row 1
                                                     # ip address for simulator is in first array element

commands = build_sim_inputs(input_file,columns_2)    # array of all simulator commands from spreadsheet
ip = parameters.first


while loop_count < loop                              # get loop from console entry  

_commands = commands.map{|x| x.dup}                 # preserve the original command array by duplicating (deep copy)

  puts "Start loop #{loop_count}"
# write array to V4 Simulator
  _commands.each do |command|               # commands is an array of command arrays
    step_log = []
    # puts "\n\n"                             # each command array is a spreadsheet row
    print command.inspect
    step_log.push(command).flatten          # command array to log
    run_flag = command.pop                  # pop run flag ("K") off the 'command' array
    if run_flag != "y"                      # do not execute row if run flag != y
    puts "      **Run flag = 'n' - do not run this step**"
    else
      emr_dly_flag = command.pop            # pop emergency delay flag ("J") off the 'command' array
      emr_dly = emergency_delay(command.pop)# pop emergency delay ("F") and convert to time in seconds
      emr_dly = 0 if emr_dly_flag != "y"    # override: emergency delay = 0 seconds when flag is 'no'
      step_dly = command.pop.to_i           # pop step delay ("I") off the 'command' array and convert to integer
      life_lbl = command.pop                # remove the life label from the array to keep off of the simulator command 
      cmd = command.join(",")               # create a comma separated string for simulator
      print "Start: " + "#{t_stamp} | "
       step_log.push(life_lbl)              # Add life point label to log string
      step_log.push(t_stamp)                # start time to log
      t1 = Time.now

      sock.send("#{cmd}", 0, ip, port)      # "A","H","G" to simulator
      sleep (emr_dly + step_dly)            # sleep for emergency + step delay in seconds
      print "Stop: " + "#{t_stamp}"
      step_log.push(t_stamp)
      t2 = Time.now
      step_t = t2-t1
      puts " | Duration = #{step_t}  "
      step_log.push(step_t)
      step_log.flatten
    end
    script_log.push(step_log)
  end
  loop_count += 1
end

sock.close
f = Time.now

puts"\n ** log console to csv ** \n\n"

csv_out = script_log.each {|x| puts x.join(',')}
# write_csv(csv_out,log)
puts csv_out.inspect
# send_mail(from,to)

File.open('csv_out', 'w+') do |f|

  f.puts 'test started at ' + Time.now.to_s
  (csv_out.each{|x|(x.to_s).puts})
  f.puts
 
  # f.puts ....   write something
  sleep 5
  f.puts 'test ended at ' + Time.now.to_s
end




f = Time.now
print "\n\n Total elapsed time = #{f - s} sec"  #script run time
