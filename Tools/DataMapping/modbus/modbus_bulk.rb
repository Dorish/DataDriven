# == Synopsis
#
# modbus_bulk: performs a bulk set of modus queries from the supplied
# spreadsheet and creates a time stamped copy of that spreadsheet with results
#
#input file - A modbus spreadsheet - can get from the ipbml card web page->data/logs->modbus elements
#         put it in the same folder with this script.
#run - type argument - ip address spreadsheet name. For example -  126.4.202.121 pa-dse.xls
#         
# == Usage
#
# modbus_bulk ip_address|comport path_to_spreadsheet
#
$:.unshift File.dirname(__FILE__)
require 'rdoc/usage'
require 'modbus_poller'

# Verify arguments
if ARGV.length != 2
  RDoc::usage
end

# Handle arguments
interface = ARGV.shift
path_to_spreadsheet = ARGV.shift

begin
  if interface =~ /\d+\.\d+\.\d+\.\d+/ #Not a perfect regex, but good enough
    modbus_bulk = ModbusPollerTCP.new(1, interface, path_to_spreadsheet, ModbusPoller::PATH_TO_MODPOLL,'')
    modbus_bulk.run
  elsif interface =~ /com\d+/i
    modbus_bulk = ModbusPoller485.new(1, interface, path_to_spreadsheet, ModbusPoller::PATH_TO_MODPOLL)
    modbus_bulk.run
  else
    RDoc::usage
  end

rescue Exception => e
  puts"\n\n #{e} \n\n"
end

