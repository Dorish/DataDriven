=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  cfg_firmware_Iswebcard

*Test_Case_Number*
  700.030.20.110

*Description*
  Validate the Firmware Update Web Configuration
     - Positive
     - Negative
    
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

#Launch the Firmware Update Web ruby script
#Add library file to the path
$:.unshift File.dirname(__FILE__).chomp('driver/web')<<'lib' # add library to path
s = Time.now
require 'generic'
require 'watir/process' 

def read_ip_addr
  host = 'C:/WINDOWS/system32/drivers/etc/hosts'
  testsite = /^(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\s+(test_site)/
  
  File.open(host).each do|line|
  if line =~ testsite
    ip = $1  # $1 is the group in the valid_ip regex
    puts "Existing test_site IP address is: ",ip
    return ip
  end
end
end

begin
  puts" \n Executing: #{(__FILE__)}\n\n" # print current filename
  g = Generic.new
  roe = ARGV[1].to_i
  #Open up a new excel instance and save it with the timestamp.
  #Open the IE browser
  #Collect the support page information and save it in the time stamped spreadsheet.
  excel = g.setup(__FILE__)
  wb,ws = excel[0][1,2]
  rows = excel[1][1]
  #temp is used to remember test_site, it will be used in the ensure section.
  temp=excel[1][2]
  
  $ie.speed = :zippy
  #Navigate to the 'Configure' tab
  g.config.click
  $ie.maximize
    #Login if not called from controller
  g.logn_chk(g.equipinfo,excel[1])

  # Below scripts complete using IP address to open the page,
  # and need to login again no matter run indivadually or run from controller.
  excel[1][2] = read_ip_addr
  $ie.goto(excel[1][2])
  site,name,pswd = excel[1][2..4]
  g.config.click
  g.login(site,name,pswd)
  g.equipinfo.click
  #Click the Configure Firmware Update Web on the left side of window
  g.updtweb.click
  
  row = 1
  while(row <= rows)
    puts "Test step #{row}"
    row +=1 # add 1 to row as execution starts at drvr_ss row 2
    #sleep 5
    
    # Write the Firmware Update Web fields
    filepath=(File.dirname(__FILE__).chomp('driver/web')<<'TestFiles').gsub('/', '\\')
    if (ws.Range("k#{row}")['Value'].to_s != "")
      g.web_file.click_no_wait
      g.web_file.set(filepath+ws.Range("k#{row}")['Value'].to_s)
    else
      #Empty input
      $ie.refresh
    end
    #Is there a popup expected? 
    pop = ws.Range("af#{row}")['Value'].to_s 
    puts "  pop_up value = #{pop}" unless pop == 'no'
    #sleep 1

    #If popup, handle with reset OK or reset Cancel to continue
    if (pop == "msg")
      g.web_updt.click_no_wait
      popup_txt  = g.jsClick('OK')
      puts "Pop-Up text is #{popup_txt}"
      ws.Range("bk#{row}")['Value'] = popup_txt
    end
    if (pop == "no")
      g.web_updt.click
      $ie.refresh# If page becomes no display, it can go back to the select file page, so next step can run.
    end
    wb.Save
  end

  f = Time.now  #finish time
#Capture error if any in the script  
rescue Exception => e
  f = Time.now  #finish time 
  puts" \n\n **********\n\n #{$@ } \n\n #{e} \n\n ***"
  error_present=$@.to_s

ensure #this section is executed even if script goes in error
    #Go back to the page using 'test_site' and back to the configure->equipinfo page,so next script can run.
    excel[1][2]=temp
    $ie.goto(excel[1][2])
    g.config.click
    g.equipinfo.click_no_wait
    g.jsClick('OK')
    # If roe > 0, script is called from controller
    # If roe = 0, script is being ran independently
    #Close and save the spreadsheet and thes web browser.
    g.tear_down_d(excel[0],s,f,roe,error_present)
    if roe == 0
      $ie.close
    end
end