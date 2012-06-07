# To change this template, choose Tools | Templates
# and open the template in the editor.

$:.unshift File.dirname(__FILE__).sub('Tools','lib') #add lib to load path
#require  File.dirname(__FILE__).sub('Tools','lib')+ 'generic'
require 'generic'

def get_input_parameters()
  input_parameters = Hash.new()
  if ARGV.length == 6
    input_parameters = { "ip_address"=>ARGV[0], "user_name"=>ARGV[1], "password"=>ARGV[2], "interval_time"=>ARGV[3], "script_file"=>ARGV[4],"output_file"=>ARGV[5]}
  else
    puts "Please type the IP Address which you want to telnet: "
    input_parameters["ip_address"] = gets
    puts "Please type the card's username which you want to telnet: "
    input_parameters["user_name"] = gets
    puts "Please type the card's password which you want to telnet: "
    input_parameters["password"] = gets
    puts "Please type the Interval Time: "
    input_parameters["interval_time"] = gets.chomp.to_i
    puts "Please type the Script File which record the test process: "
    input_parameters["script_file"] = gets
    puts "Please type the output File which you want to store the log: "
    input_parameters["output_file"] = gets
  end
  return input_parameters
end

begin
  g = Generic.new
  input_parameters = get_input_parameters()
  logging_file = File.new(input_parameters["output_file"], "w")
  
  # connect to card from telnet
  telnet = g.telnet_connect(input_parameters["ip_address"], input_parameters["user_name"], input_parameters["password"])
  logging_file.puts("#{telnet}")

  # read the script file
  running_location = nil

  # wait for the interval
  sleep input_parameters["interval_time"]

  # output the log information
  logging_file.puts("#{running_location}")

rescue Exception => e
   puts "Telnet logging failed: #{e}\n\n"
   puts $@.to_s
ensure
logging_file.close
end
