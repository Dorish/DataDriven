# To change this template, choose Tools | Templates
# and open the template in the editor.

$:.unshift File.dirname(__FILE__).sub('Tools','lib') #add lib to load path
require 'generic'
$desired_location = '4,3'  # the guide of '4,3' can navigate to system operation info menu item,

def get_input_parameters(work_sheet) 
  index = 2
  input_parameter_list = Array.new()
  while(work_sheet.Range("A#{index}")['Value'] != nil)
    input_parameter = Hash.new()
    input_parameter["ip_address"] = work_sheet.Range("A#{index}")['Value']
    
    # ignore the repeat items
    if have_repeat_item?(input_parameter_list, input_parameter["ip_address"]) == true
      index = index + 1
      next
    end
    
    input_parameter["user_name"] = work_sheet.Range("B#{index}")['Value']
    input_parameter["password"] = work_sheet.Range("C#{index}")['Value']
    input_parameter["interval_time"] = work_sheet.Range("D#{index}")['Value']
    input_parameter["loop_times"] = work_sheet.Range("E#{index}")['Value']

    # get instance of logging file for each IP address
    output_file_name = File.dirname(__FILE__) + '\\' + input_parameter["ip_address"] + '_' + Time.now.strftime("%m-%d_%H-%M-%S") + '.txt'
    logging_file = File.new(output_file_name, "a+")
    logging_file.sync = true
    logging_file.binmode
    input_parameter["logging_file"] = logging_file

    input_parameter_list.push(input_parameter)
    index = index + 1
  end
  return input_parameter_list
 end

def have_repeat_item?(input_parameter_list,ip_address)
  input_parameter_list.each { |item|
      if item["ip_address"] == ip_address
        return true
      end
   }
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
  logging_file.puts "\n\n-------------------------------------------------"
  logging_file.puts Time.now
  logging_file.puts '-------------------------------------------------'
end

begin
  g = Generic.new

  # open spreadsheet and hide after 3 seconds
  excel_name = File.dirname(__FILE__) + '\\' + 'telnet_logging.xls' # use this path when we run the script in ruby environment
  #excel_name = Dir.pwd + '/' + 'telnet_logging.xls' # use this path when we create the executable file use Exerb
  setup = g.new_xls(excel_name,1)
  spread_sheet = setup[0]
  work_book = setup[1]
  work_sheet = setup[2]
  sleep 3
  spread_sheet.visible = false

  # get all input parameters
  input_parameters_list = get_input_parameters(work_sheet)
  loop_times = input_parameters_list[0]["loop_times"]
  interval_time = input_parameters_list[0]["interval_time"]

  while(loop_times.to_i > 0)
    input_parameters_list.each{|input_parameters|
      ip_address = input_parameters["ip_address"]
      logging_file = input_parameters["logging_file"]

      # record the card's state, available or not?
      is_available = is_card_available(ip_address)

      # add the time stamp at the beginning of log file.
      add_timestamp(logging_file)

      if is_available
        # connect to card from telnet
        telnet = g.telnet_connect(ip_address, input_parameters["user_name"], input_parameters["password"])
        # navigate to desired location
        navigate_to_location(telnet, $desired_location,logging_file)
        # terminate telnet
        telnet.close
        else
        logging_file.puts("#{ip_address} did not respond to the ping request")
      end

      if loop_times.to_i == 1
        logging_file.close
      end
    }
    sleep (interval_time.to_i)*60
    loop_times = loop_times.to_i - 1
end

rescue Exception => e
  puts "Telnet logging failed: #{e}\n\n"
  puts $@.to_s
ensure
  work_book.close
  spread_sheet.quit
end
