=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  manag_protocol_info for GXT/NX devices

*Test_Case_Number*
  700.090.20.110

*Description*
  Validate the Management Protocol Configuration Information
     - Enable / Disable SNMP Agent Check Box
     - Popup handle with Reset Ok button
     - Popup handle with Reset Cancel button

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

#Launch the Management Protocol ruby script
#Add library file to the path

#$:.unshift File.dirname(__FILE__).chomp('Tools\SNMP\snmp_config')<<'lib' # add library to path
$:.unshift File.dirname(__FILE__).chomp('Tools/SNMP/snmp_config')<<'lib' # add library to path

s = Time.now
require 'watir'
require 'generic'

#define web node
def uav; $ie.frame(:index,4);end
def tav; $ie.frame(:index,5);end

#tab link and butten
def protocols; uav.link(:text, 'Protocols'); end
def snmp; uav.link(:text, 'SNMP');end
def snmpaccess_link; uav.link(:text, 'SNMPv1/v2c Access Settings (20)');end
def snmpv1v2access(val); uav.link(:text, "SNMPv1/v2c Access Settings [#{val}]");end
def edit; tav.button(:id, 'editButton');end
def submit; tav.button(:id,'submitButton');end
def cancel; tav.button(:id,'cancelButton');end

#Snmpv1v2 trap setting
def accaddress; tav.text_field(:id, 'str44');end
def acctype; tav.select_list(:id, 'enum22');end
def community; tav.text_field(:id, 'str129');end

#login
def login(site,user,pswd)
  lang = `systeminfo`
  if lang =~ /en-us*/
    $titl          = "Connect to "
  elsif lang =~ /zh-cn*/
    $titl           = "Á¬½Óµ½ "
  end

  conn_to = $titl + site
  Thread.new{
    thread_cnt = Thread.list.size
    sleep 1 #This sleep is critical, timing may need to be adjusted
    Watir.autoit.WinWait(conn_to)
    Watir.autoit.WinActivate(conn_to)
    Watir.autoit.Send(user)
    Watir.autoit.Send('{TAB}')
    Watir.autoit.Send(pswd)
    Watir.autoit.Send('{ENTER}')
  }
end

begin

  puts" \n Executing: #{(__FILE__)}\n\n" # print current filename
  g = Generic.new

  #Open up a new excel instance and save it with the timestamp.
  #Open the IE browser
  #Collect the support page information and save it in the time stamped spreadsheet.

  #excel = g.setup(__FILE__)
  base_xl = (__FILE__).gsub('/','\\').chomp('rb')<<'xls'
  excel = g.xls_timestamp(base_xl,'ind',nil)
  
  g.open_ie(excel[1][2])
  wb,ws = excel[0][1,2]
  rows = excel[1][1]

  $ie.speed = :zippy
  $ie.maximize
  #Click the Management Protocl link on the left side of window
  #Login if not called from controller
  sleep 1
  protocols.click
  snmp.click
  login(excel[1][2],excel[1][3],excel[1][4])
  edit.click
  sleep 5
  submit.click
  snmpaccess_link.click

  number = 1 #every user number

  for row in 1..rows
    row +=1 # add 1, execution starts at drvr_ss row 2
    puts " Executing -  Test step #{ws.Range("g#{row}")['Value'].to_i}"

    #Every 20 row start from user 1
    if number > 20
      number = 1
    end

    #Open access number page
    snmpv1v2access(number).click

    ################# write  settings info############
    sleep 1
    edit.click
    sleep 0.5

    #1 SNMP Trap Target Address
    puts "1 SNMP Trap Target Address"
    accaddress.set(ws.Range("k#{row}")['Value'].to_s)

    #2 SNMP Trap Port
    puts "2 SNMP Trap Port"
    acctype.select_value(ws.Range("L#{row}")['Value'].to_s)

    #3 SNMP Trap Community String
    puts "3 SNMP Trap Community String"
    community.set(ws.Range("M#{row}")['Value'].to_s)

    sleep 0.5
    submit.click
    number +=1
  end

  ## read and resum all trap setting
  for i in 1..rows

    snmpv1v2access(i).click
    sleep 1
    edit.click
    sleep 0.5

    ######### # read settings setting and write result to ss

    #1 read SNMP Trap Target Address
    ws.Range("t#{i+1}")['Value'] = accaddress.value.to_s

    #2 read SNMP Trap Port
    ws.Range("u#{i+1}")['Value'] = acctype.value.to_s

    #3 read SNMP Trap Port
    ws.Range("v#{i+1}")['Value'] = community.value.to_s

    #1 resum SNMP Trap Target Address
    puts "1 resum SNMP Trap Target Address"
    accaddress.set('')

    #2 resum SNMP Trap Port
    puts "2 resum SNMP Trap Port"
    acctype.select_value("1")

    #3 resum SNMP Trap Community String
    puts "3 resum SNMP Trap Community String"
    community.set('')

    sleep 0.5
    submit.click
    i += 1

  end

  f = Time.now  #finish time
  #Capture error if any in the script


  $ie.close
  wb.save
  wb.close
  puts "Test finished at #{f}"

end