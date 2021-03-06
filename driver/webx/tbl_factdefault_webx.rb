=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  factdefault_table_info

*Test_Case_Number*
  pending

*Description*
  Validate Factory Defaults Information table

*Variable_Definitions*
    s = test start time
    f = test finish time
    e = test elapsed time
    roe = row number in controller spreadsheet
    excel = nested array that contains an instance of excel and driver parameters
    ss = spreadsheet
    wb = workbook
    ws = worksheet
    dvr_ss = driver spreadsheet
    rows = number of rows in spreadsheet to execute
    site = url/ip address of card being tested
    name = user name for login
    pswd = password for login

=end

$:.unshift File.dirname(__FILE__).chomp('driver/webx')<<'lib' # add library to path
s = Time.now
require 'generic'

begin
  puts" \n Executing: #{(__FILE__)}\n\n" # show current filename
  g = Generic.new
  roe = ARGV[1].to_i
  excel = g.setup(__FILE__)
  wb,ws = excel[0][1,2]
  
  g.config.click
  g.logn_chk(g.factdef,excel[1])
  
  g.table_info(1,2,2,2,ws)

rescue Exception => e
  puts" \n\n **********\n\n #{$@ } \n\n #{e} \n\n ***"
  error_present=$@.to_s
ensure #this section is executed even if script goes in error
  f = Time.now
  # If roe > 0, script is called from controller
  # If roe = 0, script is being ran independently
  #Close and save the spreadsheet and thes web browser.
  g.tear_down_d(excel[0],s,f,roe,error_present)
  if roe == 0
    $ie.close
  end
end

 






