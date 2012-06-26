# To change this template, choose Tools | Templates
# and open the template in the editor.

$:.unshift File.dirname(__FILE__).sub('Tools\\Telnet','lib') #add lib to load path
require 'generic'

$desired_location = '4,3'  # the guide of '4,3' can navigate to system operation info menu item,

def tn_close(telnet,navigate_str)
  esc = "\x1b"
  navigation = navigate_str.split(',')
  esc_num = navigation.length + 2 #need to press esc 2 times more from main menu to exit telnet session
  esc_num.times do
    telnet.write(esc) {|c| print c}
    telnet.waitfor(/\?>|<Esc> Ends Session|[Bb]ye/) {|c| print c}
  end
end

def get_input_parameters(csv_file)
  input_parameter_list = Array.new()
  open(csv_file).map do |line|
  if line !~ /IP Address/     # ignore column header row
    input_parameter = Hash.new()
    line =  line.split(/,/).to_a.push  # convert each line to an array
    input_parameter["ip_address"] = line[0]
    input_parameter["user_name"] = line[1]
    input_parameter["password"] = line[2]
    input_parameter["interval_time"] = line[3]
    input_parameter["loop_times"] = line[4]

    # get instance of logging file for each IP address
    output_file_name = File.dirname(__FILE__) + '\\' + input_parameter["ip_address"] + '_' + Time.now.strftime("%m-%d_%H-%M-%S") + '.txt'
    logging_file = File.new(output_file_name, "a+")
    logging_file.sync = true
    logging_file.binmode
    input_parameter["logging_file"] = logging_file

    input_parameter_list.push(input_parameter)
    end
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

  csv_file = File.dirname(__FILE__) + '\\' + 'telnet_logging.csv' # use this path when we run the script in ruby environment
  input_parameter_list = Array.new()
  input_parameter_list = get_input_parameters(csv_file)

  # For serial processing, they have the same loop times and interval time for all IP addresses.
  loop_times = input_parameter_list[0]["loop_times"]
  interval_time = input_parameter_list[0]["interval_time"]
  while(loop_times.to_i > 0)
    input_parameter_list.each{|input_parameters|
      ip_address = input_parameters["ip_address"]
      logging_file = input_parameters["logging_file"]

      # record the card's state, available or not?
      is_available = is_card_available(ip_address)

      # add the time stamp at the beginning of log file.
      add_timestamp(logging_file)
      if is_available
        tries = 0
        begin
          # connect to card from telnet
          telnet = g.telnet_connect(ip_address, input_parameters["user_name"], input_parameters["password"])
          # navigate to desired location
          navigate_to_location(telnet, $desired_location,logging_file)
          # close telnet session from card side
          tn_close(telnet,$desired_location)
          # terminate telnet
          telnet.close
        rescue Exception=>e
          sleep 1
          tries += 1
          if telnet == nil # build telnet connection failed.
            puts "telnet to #{ip_address} failed"
            logging_file.puts("telnet to #{ip_address} failed")
          else
           # tn_close(telnet,$desired_location) # if this sentence error, no handle script for it.
            telnet.close
          end
          if is_card_available(ip_address)==true # card is still available - retry
            puts "retry #{tries} times"
            retry if tries <= 9
            puts "retry limit reached!"
          else # card is not available - log that 'card did not respond ' and continue running
            logging_file.puts("#{ip_address} did not respond to the ping request")
          end
        end
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
end
