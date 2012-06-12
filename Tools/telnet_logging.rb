# To change this template, choose Tools | Templates
# and open the template in the editor.

$:.unshift File.dirname(__FILE__).sub('Tools','lib') #add lib to load path
require 'generic'
$desired_location = '4,3'  # the guide of '4,3' can navigate to system operation info manu item,

def get_input_parameters(work_sheet)
    input_parameters = Hash.new()
    input_parameters["ip_address"] = work_sheet.Range("A2")['Value']
    input_parameters["user_name"] = work_sheet.Range("B2")['Value']
    input_parameters["password"] = work_sheet.Range("C2")['Value']
    input_parameters["interval_time"] = work_sheet.Range("D2")['Value']
    input_parameters["loop_times"] = work_sheet.Range("E2")['Value']
    return input_parameters
 end

def is_card_available(ip)
    reply_from = Regexp.new('Reply from')
       results = `ping #{ip}`
      if (reply_from.match(results)) then
        return true
      end
      return false
  end

def navigate_to_location(telnet, navigate_str,logging_file)
  navigation = navigate_str.split(',')
      navigation.each do |s| 
          telnet.write(s) {|c| print c}
          log_info = telnet.waitfor(/\?>|press any key to exit|Hit/) {|c| print c}
          if s == '3'
              # delete the first character( it is the navigate number).
              log_info = log_info[1,log_info.length]
              logging_file.write(log_info)
          end  
      end
end

def add_timestamp(logging_file)
  logging_file.puts '-------------------------------------------------'
  logging_file.puts Time.now
  logging_file.puts '-------------------------------------------------'
end

begin
  g = Generic.new
  
  excel_name = File.dirname(__FILE__) + '\\' + 'Telnet_diagnostic_template.xls'
  setup = g.new_xls(excel_name,1)
  spreat_sheet = setup[0]
  work_book = setup[1]
  work_sheet = setup[2]

  # get all input parameters
  input_parameters = get_input_parameters(work_sheet)
  ip_address = input_parameters["ip_address"]
  interval_time = input_parameters["interval_time"]
  loop_times = input_parameters["loop_times"]
  
  # open the output file,
  output_file_name = File.dirname(__FILE__) + '\\' + ip_address + '_' + Time.now.strftime("%m-%d_%H-%M-%S") + '.txt'
  logging_file = File.new(output_file_name, "a+")
  logging_file.sync = true
  logging_file.binmode

  while  ( loop_times.to_i > 0)
    # record the card's state, available or not?
    is_available = is_card_available(ip_address)

    # add the time stamp at the beginning of log file.
    add_timestamp(logging_file)

    if is_available
        # connect to card from telnet
        telnet = g.telnet_connect(ip_address, input_parameters["user_name"], input_parameters["password"])

        # navigate to desired location
        navigate_to_location(telnet, $desired_location,logging_file)

        # shut up telnet
        telnet.close
        
      else
      logging_file.puts("#{ip_address} is not available, please check the connection...")
    end

    sleep (interval_time.to_i)
    loop_times = loop_times.to_i - 1
  end

rescue Exception => e
   puts "Telnet logging failed: #{e}\n\n"
   puts $@.to_s
ensure
  work_book.close
  spreat_sheet.quit
  logging_file.close
end
