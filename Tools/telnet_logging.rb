# To change this template, choose Tools | Templates
# and open the template in the editor.

$:.unshift File.dirname(__FILE__).sub('Tools','lib') #add lib to load path
require 'generic'

def get_input_parameters()
  input_parameters = Hash.new()
  if ARGV.length == 6
    input_parameters = { "ip_address"=>ARGV[0], "user_name"=>ARGV[1], "password"=>ARGV[2], "interval_time"=>ARGV[3], "desired_location"=>ARGV[4],"output_file"=>ARGV[5]}
  else
    puts "Please type the IP Address which you want to telnet: "
    input_parameters["ip_address"] = gets
    puts "Please type the card's username which you want to telnet: "
    input_parameters["user_name"] = gets
    puts "Please type the card's password which you want to telnet: "
    input_parameters["password"] = gets
    puts "Please type the Interval Time: "
    input_parameters["interval_time"] = gets.chomp.to_i
    puts "Please type the desired location which you want to navigate, following the format as <n,n>, n should be integer number: "
    input_parameters["desired_location"] = gets
    puts "Please type the output File which you want to store the log: "
    input_parameters["output_file"] = gets
  end
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

def navigate_to_location(telnet, navigate_str)
  navigation = navigate_str.split(',')
      navigation.each do |s| 
      #telnet.cmd(s) {|c| print c}
      telnet.write(s) {|c| print c}
      telnet.waitfor(/\?>|press any key to exit|Hit/) {|c| print c}
      end
end

# create a new method to connect the telnet, because the /lib/telnet.rb have a 5s sleep.
def telnet_connect(ip,login,pswd,output_log)
    prompt_type=/[ftpusr:~>]*\z/n
    puts"ip,=#{ip},login=#{login},password=#{pswd}"
    telnet = Net::Telnet.new('Host' => ip,'Prompt' =>prompt_type ,"Output_log"   => output_log  )

    #The prompt should be the real prompt while you entered your system
    telnet.cmd(''){|c| print c}
    telnet.waitfor(/[Ll]ogin[: ]*\z/n) {|c| print c}
    telnet.cmd(login) {|c| print c}
    telnet.waitfor(/Password[: ]*\z/n) {|c| print c}
    telnet.cmd(pswd) {|c| print c}

    telnet.waitfor(/[>]/n) {|c| print c}
    return telnet
  end

begin
  g = Generic.new
  input_parameters = get_input_parameters()
  logging_file = File.new(input_parameters["output_file"], "a+")
  logging_file.sync = true
  logging_file.binmode

  ip_address = input_parameters["ip_address"]
  interval_time = input_parameters["interval_time"]

  while   1
    # record the card's state, available or not?
    is_available = is_card_available(ip_address)
    if is_available
        logging_file.puts("\n#{ip_address} is available, please wait for connect... \n\n")

        # connect to card from telnet
        #telnet = g.telnet_connect(input_parameters["ip_address"], input_parameters["user_name"], input_parameters["password"],input_parameters["output_file"])
        telnet = telnet_connect(input_parameters["ip_address"], input_parameters["user_name"], input_parameters["password"],input_parameters["output_file"])

        # navigate to desired location
        navigate_to_location(telnet, input_parameters["desired_location"])

        # shut up telnet
        telnet.close

      else
      logging_file.puts("#{ip_address} is not available, please try again...")
    end

    # wait for the interval
    logging_file.puts("\n\nPlease wait for  #{interval_time} seconds to start another logging...\n\n")
    logging_file.puts("\n\n==========================================\n\n")
    sleep (interval_time.to_i)
  end

rescue Exception => e
   puts "Telnet logging failed: #{e}\n\n"
   puts $@.to_s
ensure
  logging_file.close
end
