=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  cfg_read.rb

*Test_Case_Number*


*Description*
  Read all the configration page Information configuration
     Select
     - Deselect
     - Reset Ok and Reset Cancel

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

#Launch the Factory Default ruby script
#Add library file to the path
dir = File.dirname(__FILE__)
$:.unshift(dir.gsub('Tools','lib'))# add lib to load path
s = Time.now
require 'watir/ie'
require 'win32ole'
require 'generic'
require 'watir/process'

begin
  puts" \n Executing: #{(__FILE__)}\n\n" # print current filename
  g = Generic.new
  roe = ARGV[1].to_i
  row = 5

  #Open up excel
  excel  = (__FILE__.chomp(".rb")<<(".xls"))
  ss = WIN32OLE::new('excel.Application')
  ss.DisplayAlerts = false #Stops excel from displaying alerts
  ss.Visible = true # For debug
  wb = ss.Workbooks.Open(excel)
  ws = wb.Worksheets(1)

  site = ws.Range("B2")['Value'].to_s #Define Test Site
  name = ws.Range("C2")['Value'].to_s #Define Login Name
  pswd = ws.Range("D2")['Value'].to_s #Define Password

  #Open the IE browser
  $ie = Watir::IE.new
  $ie.goto( site)
  $ie.speed = :zippy

  #Navigate to the Configure tab
  g.systemos
  g.config.click
  $ie.maximize
  g.login(site,name,pswd)

  #Read Agent Information
  puts 'Read Agent Information'
  g.equipinfo.click
  ws.Range("B#{row}")['Value'] = g.name.value
  ws.Range("C#{row}")['Value'] = g.cont.value
  ws.Range("D#{row}")['Value'] = g.loc.value
  ws.Range("E#{row}")['Value'] = g.desc.value
  row += 2

  #Read Network Settings
  puts 'Read Network Settings'
  g.netset.click
  ws.Range("B#{row}")['Value'] = g.net_speed.value
  if (g.net_bootmode(0).checked? == true)
    ws.Range("C#{row}")['Value'] = "0"
    ws.Range("D#{row}")['Value'] = g.net_ipaddr.value
    ws.Range("E#{row}")['Value'] = g.net_subnet.value
    ws.Range("F#{row}")['Value'] = g.net_gateway.value
  elsif (g.net_bootmode(1).checked? == true)
    ws.Range("C#{row}")['Value'] = "1"
  else
    ws.Range("C#{row}")['Value'] = "2"
  end  
  row += 2
  
  #Read  DNS
  puts 'Read DNS'
  g.dns.click
  if (g.dns_mode(0).checked? == true)
    ws.Range("B#{row}")['Value'] = "0"
  else
    ws.Range("B#{row}")['Value'] = "1"
	end
  ws.Range("C#{row}")['Value'] = g.dns_addr1.value
  ws.Range("D#{row}")['Value'] = g.dns_addr2.value
  ws.Range("E#{row}")['Value'] = g.dns_int.value
  ws.Range("F#{row}")['Value'] = g.dns_suf.value
  row += 2
  
  #Read  Time (SNTP)
  puts 'Read Time (SNTP)'
  g.time.click
  ws.Range("B#{row}")['Value'] = g.timesrvr.value
  if (g.timesync('0').checked? == true)
    ws.Range("C#{row}")['Value'] = "0"
  else
    ws.Range("C#{row}")['Value'] = "1"
  end
  ws.Range("D#{row}")['Value'] = g.timezone.value
  row += 2
  
  #Read  Management Protocol
  puts 'Read Management Protocol '
  g.mgtprot.click
  ws.Range("B#{row}")['Value'] = g.snmp_v1v2.value
  ws.Range("C#{row}")['Value'] = g.snmp_v3.value
  row += 2
  
  #Read  SNMP
  puts 'Read SNMP'
  g.snmp.click
  ws.Range("B#{row}")['Value'] = g.checkbox(g.snmp_auth)
  ws.Range("C#{row}")['Value'] = g.snmp_hb.value
  ws.Range("D#{row}")['Value'] = g.checkbox(g.upsmib)
  ws.Range("E#{row}")['Value'] = g.checkbox(g.upstraps)
  ws.Range("F#{row}")['Value'] = g.checkbox(g.lgpmib)
  ws.Range("G#{row}")['Value'] = g.checkbox(g.lgptraps)
  ws.Range("H#{row}")['Value'] = g.checkbox(g.sysnotify)
  row += 2
  
  #Read V1 Access
  puts 'Read V1 Access'
  g.v1access.click
  for i in 1..20
    ws.Range("C#{row}")['Value'] = g.access_addr(i).value
    if g.access_type(i,'1').checked? == true
      ws.Range("D#{row}")['Value'] = "1"
    elsif g.access_type(i,'0').checked? == true
      ws.Range("D#{row}")['Value'] = "0"
    end
    ws.Range("E#{row}")['Value'] = g.access_com(i).value
    row += 1
  end
  row += 1
  
  #Read V1 Traps
  puts 'Read V1 Traps'
  g.v1traps.click
  for i in 1..20
    ws.Range("C#{row}")['Value'] = g.trap_addr(i).value
    ws.Range("D#{row}")['Value'] = g.trap_port(i).value
    ws.Range("E#{row}")['Value'] = g.trap_com(i).value
    ws.Range("F#{row}")['Value'] = g.checkbox(g.trap_hb(i))
    row += 1
  end
  row += 1
  
  #Read V3 Settings
  puts 'Read V3 Settings'
  g.snmpv3.click
  for i in 1..20
    ws.Range("C#{row}")['Value'] = g.checkbox(g.v3_enable(i))
    ws.Range("D#{row}")['Value'] = g.v3_user(i).value
    if g.v3_auth(i,'0').checked? == true
      ws.Range("E#{row}")['Value'] = "0"
    elsif g.v3_auth(i,'1').checked? == true
      ws.Range("E#{row}")['Value'] = "1"
    elsif g.v3_auth(i,'2').checked? == true
      ws.Range("E#{row}")['Value'] = "2"
    end
    ws.Range("F#{row}")['Value'] = g.v3_auth_secret(i).value
    if g.v3_privacy(i,'0').checked? == true
      ws.Range("G#{row}")['Value'] = "0"
    elsif g.v3_privacy(i,'1').checked? == true
      ws.Range("G#{row}")['Value'] = "1"
    end
    ws.Range("H#{row}")['Value'] = g.v3_privacy_secret(i).value
    ws.Range("I#{row}")['Value'] = g.checkbox(g.v3_acc_read(i))
    ws.Range("J#{row}")['Value'] = g.checkbox(g.v3_acc_write(i))
    ws.Range("K#{row}")['Value'] = g.v3_sources(i).value
    ws.Range("L#{row}")['Value'] = g.checkbox(g.v3_notify(i))
    ws.Range("M#{row}")['Value'] = g.v3_destinations(i).value
    ws.Range("N#{row}")['Value'] = g.checkbox(g.v3_heartbeat(i))
    ws.Range("O#{row}")['Value'] = g.v3_port(i).value
    row += 1
  end
  row += 1
  
  #Read Messaging
  puts 'Read Messaging'
  g.msging.click
  ws.Range("B#{row}")['Value'] =  g.checkbox(g.email_msg)
  ws.Range("C#{row}")['Value'] = g.checkbox(g.sms_msg)
  row += 2
  
  #Read Email
  puts 'Read Email'
  g.email.click
  ws.Range("B#{row}")['Value'] = g.email_from.value
  ws.Range("C#{row}")['Value'] = g.email_to.value
  if g.email_subjectttype(0).checked? == true                 # Subject Type
    ws.Range("D#{row}")['Value'] = "0"
  else
    ws.Range("D#{row}")['Value'] = "1"
    ws.Range("E#{row}")['Value'] = g.email_custsubj.value    # Custom Subject
  end
  ws.Range("F#{row}")['Value'] = g.email_srvr.value
  ws.Range("G#{row}")['Value'] = g.email_port.value
  row += 2
  
  #Read Sms
  puts 'Read Sms'
  g.sms.click
  ws.Range("B#{row}")['Value'] = g.sms_from.value
  ws.Range("C#{row}")['Value'] = g.sms_to.value
  if g.sms_subjecttype(0).checked? == true                 # Subject Type
    ws.Range("D#{row}")['Value'] = "0"
  else
    ws.Range("D#{row}")['Value'] = "1"
    ws.Range("E#{row}")['Value'] = g.sms_custsubj.value    # Custom Subject
  end
  ws.Range("F#{row}")['Value'] = g.sms_srvr.value
  ws.Range("G#{row}")['Value'] = g.sms_port.value
  row += 2
  
  #Read Customize Message
  puts 'Read Customize Message'
  g.custmsg.click
  ws.Range("B#{row}")['Value'] = g.checkbox(g.email_addr)  #Email
  ws.Range("C#{row}")['Value'] = g.checkbox(g.email_evtdesc)
  ws.Range("D#{row}")['Value'] = g.checkbox(g.email_name)
  ws.Range("E#{row}")['Value'] = g.checkbox(g.email_cont)
  ws.Range("F#{row}")['Value'] = g.checkbox(g.email_loc)
  ws.Range("G#{row}")['Value'] = g.checkbox(g.email_desc)
  ws.Range("H#{row}")['Value'] = g.checkbox(g.email_weblnk)
  ws.Range("I#{row}")['Value'] = g.checkbox(g.email_consol)
  ws.Range("J#{row}")['Value'] = g.email_consoltime.value
  ws.Range("K#{row}")['Value'] = g.email_consolevt.value
  
  ws.Range("L#{row}")['Value'] = g.checkbox(g.sms_addr)  #SMS
  ws.Range("M#{row}")['Value'] = g.checkbox(g.sms_evtdesc)
  ws.Range("N#{row}")['Value'] = g.checkbox(g.sms_name)
  ws.Range("O#{row}")['Value'] = g.checkbox(g.sms_cont)
  ws.Range("P#{row}")['Value'] = g.checkbox(g.sms_loc)
  ws.Range("Q#{row}")['Value'] = g.checkbox(g.sms_desc)
  ws.Range("R#{row}")['Value'] = g.checkbox(g.sms_weblnk)
  ws.Range("S#{row}")['Value'] = g.checkbox(g.sms_consol)
  ws.Range("T#{row}")['Value'] = g.sms_consoltime.value
  ws.Range("U#{row}")['Value'] = g.sms_consolevt.value
  row += 2
  
  #Read Telnet
  puts 'Read Telnet'
  g.telnet.click
  ws.Range("B#{row}")['Value'] = g.checkbox(g.telnet1)
  row += 2
  
  #Read Users
  puts 'Read TelnUserset'
  g.users.click
  ws.Range("B#{row}")['Value'] = g.admin_name.value
  ws.Range("C#{row}")['Value'] = g.admin_pswd.value
  ws.Range("D#{row}")['Value'] = g.admin_pswd2.value
  ws.Range("E#{row}")['Value'] = g.user_name.value
  ws.Range("F#{row}")['Value'] = g.user_pswd.value
  ws.Range("G#{row}")['Value'] = g.user_pswd2.value
  row += 2
  
  #Read Web
  puts 'Read Web'
  g.cfgweb.click
  ws.Range("B#{row}")['Value'] = g.websrvr.value
  ws.Range("C#{row}")['Value'] = g.httpport.value
  ws.Range("D#{row}")['Value'] = g.httpsport.value
  ws.Range("E#{row}")['Value'] = g.checkbox(g.pswdprtct)
  ws.Range("F#{row}")['Value'] = g.checkbox(g.cfgctrl)
  ws.Range("G#{row}")['Value'] = g.refresh.value
  row += 2
  
  $ie.close
  wb.save
  wb.close
  puts 'Finished'
end
