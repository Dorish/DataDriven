=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  managementprotocol_table_info

*Test_Case_Number*
  pending

*Description*
  Validate Management Protocol Information table

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

$:.unshift File.dirname(__FILE__).chomp('driver/web')<<'lib' # add library to path
s = Time.now
require 'generic'

# - used for configure information table scripts
def hidden_table_info(g,start,_end,col,idx,ws)
  # iterate through all rows and columns
  pag_row = start # pag_row is the row number of web page table
  while (start <= _end)
    j = 1
    while (j <= col)
      case j
      when 1
        while g.param_descr(idx,pag_row,j).visible? ==false # skip the rows if invisible
          pag_row += 1
        end
        parameter = g.param_descr(idx,pag_row,j).text
        ws.Range("bc#{start+1}")['Value'] =  parameter
      when 2
        description = g.param_descr(idx,pag_row,j).text
        ws.Range("bd#{start+1}")['Value'] = description
      end
      j += 1
    end
    start += 1
    pag_row += 1
  end
end

begin
  puts" \n Executing: #{(__FILE__)}\n\n" # show current filename
  g = Generic.new
  roe = ARGV[1].to_i
  excel = g.setup(__FILE__)
  wb,ws = excel[0][1,2]
  
  g.config.click
  g.logn_chk(g.mgtprot,excel[1])
  
  hidden_table_info(g,1,3,2,2,ws)
  
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

 






