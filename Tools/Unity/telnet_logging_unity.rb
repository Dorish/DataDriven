# To change this template, choose Tools | Templates
# and open the template in the editor.

require 'net/telnet'

$enter = "\x0D" # Ascii Hex code for Carriage Return is 0D
$ctrl_c = "\x03" # Ascii Hex code for Ctrl+C is 03
$unity_exit_7 = "\x07" # Ascii Hex code for 7 is 07
$desired_location = "4.2.2"

 # Recreate the following method to adapting the unity card and webAdapt card.
def telnet_connect(ip,login,pswd)
    prompt_type=/[ftpusr:~>]*\z/n
    puts"ip,=#{ip},login=#{login},password=#{pswd}"
    telnet = Net::Telnet.new('Host' => ip,'Prompt' =>prompt_type ,"Output_log"   => "dump_log.txt"  )

    #The prompt should be the real prompt while you entered your system
    telnet.cmd(''){|c| print c}
    telnet.waitfor(/[Ll]ogin[: ]*\z/n) {|c| print c}
    telnet.cmd(login) {|c| print c}
    telnet.waitfor(/Password[: ]*\z/n) {|c| print c}
    telnet.cmd(pswd) {|c| print c}

   # the following sentence can wrok for unity and webAdapt.
    telnet.waitfor(/[>]|Enter selection\:/n) {|c| print c}

    sleep 5
    return telnet
  end

# close the connection for unity
def tn_unity_close(telnet)
  navigation = $desired_location.split('.')
  telnet.write($ctrl_c) {|c| print c} # Press Ctrl+C to stop DLOG file watcher
  telnet.waitfor(/Press <Enter> to return to menu/) {|c| print c}

  # Press 'enter' several times come back to the first menu
  navigation.length.times do
    telnet.write($enter) {|c| print c}
    telnet.waitfor(/Enter selection\:/) {|c| print c}
  end
  # Press the menu of '7' to exit the connect.
  telnet.write($unity_exit_7) {|c| print c}
  telnet.write($enter) {|c| print c}
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

def navigate_to_location(telnet,logging_file)
  navigation = $desired_location.split('.')
  for i in 0..navigation.length-1 do
    telnet.write(navigation[i]) {|c| print c}
    telnet.write($enter){|c| print c}
    #^-*\s*\/[\w\/]+\/dlog[0-9]{5}\.log+[\s*\w]+\s*-*$ to match "---- \/var\/log\/dlog00001.log starting ----"
    log_info = telnet.waitfor(/\?>|Enter selection\:|-*\s*\/[\w\/]+\/dlog[0-9]{5}\.log+[\s*\w]+\s*-*/) {|c| print c}
    # output the needed information from last index from.
    if i == navigation.length - 1
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

  csv_file = File.dirname(__FILE__) + '\\' + 'telnet_logging_unity.csv' # use this path when we run the script in ruby environment
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
          telnet = telnet_connect(ip_address, input_parameters["user_name"], input_parameters["password"])
          # navigate to desired location
          navigate_to_location(telnet,logging_file)
          # close telnet session from card side
          tn_unity_close(telnet)

          # terminate telnet
          telnet.close
        rescue Exception=>e
          sleep 1
          tries += 1
          if telnet == nil # build telnet connection failed.
            puts "telnet to #{ip_address} failed"
            logging_file.puts("telnet to #{ip_address} failed")
          else
           # tn_close(telnet,desired_location) # if this sentence error, no handle script for it.
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

