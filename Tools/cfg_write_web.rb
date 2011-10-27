=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Test_Script_Name*
  cfg_Write.rb

*Test_Case_Number*


*Description*
  Write all the configration page Information configuration
     Select
     - Deselect
     - Reset Ok and Reset Cancel

*Variable_Definitions*
    s = test start time
    f = test finish time
    e = test elapsed time
    roe = row number in controller spWritesheet
    excel = nested array that contains an instance of excel and driver parameters
    ss = spWritesheet
    wb = workbook
    ws = worksheet
    dvr_ss = driver spWritesheet
    rows = number of rows in spWritesheet to execute
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

  #Write Agent Information
  puts 'Write Agent Information'
  g.equipinfo.click
  g.edit.click
  g.name.set(ws.Range("B#{row}")['Value'] )
  g.cont.set(ws.Range("C#{row}")['Value'])
  g.loc.set(ws.Range("D#{row}")['Value'])
  g.desc.set(ws.Range("E#{row}")['Value'])
  g.save.click
  row += 2
  
  #Write Network Settings
  puts 'Write Network Settings'
  g.netset.click
  g.edit.click
  g.net_speed.select_value( ((ws.Range("B#{row}")['Value']).to_i).to_s)
  g.net_bootmode( ((ws.Range("C#{row}")['Value']).to_i).to_s).set

  if (g.net_bootmode(0).checked? == true)
    g.net_ipaddr.set(ws.Range("D#{row}")['Value'].to_s)
    g.net_subnet.set(ws.Range("E#{row}")['Value'].to_s)
    g.net_gateway.set(ws.Range("F#{row}")['Value'].to_s)
  end
  
  g.save.click_no_wait
  g.jsClick('OK')
  row += 2
  
  
  #Write  DNS
  puts 'Write DNS'
  g.dns.click
  g.edit.click
  g.dns_mode(((ws.Range("B#{row}")['Value']).to_i).to_s).set

  if (ws.Range("B#{row}")['Value']).to_i == 1
    g.dns_addr1.set(ws.Range("C#{row}")['Value'].to_s)
    g.dns_addr2.set(ws.Range("D#{row}")['Value'].to_s)
  end

  g.dns_int.select_value(((ws.Range("E#{row}")['Value']).to_i).to_s)
  g.dns_suf.set(ws.Range("F#{row}")['Value'].to_s)
  g.save.click
  row += 2
  
  #Write  Time (SNTP)
  puts 'Write Time (SNTP)'
  g.time.click
  g.edit.click
  g.timesrvr.set(ws.Range("B#{row}")['Value'].to_s)

  if (ws.Range("C#{row}")['Value']).to_i == 0
    g.timesync('0').set
  else
    g.timesync('1').set
  end

  g.timezone.select_value((ws.Range("D#{row}")['Value']).to_s)
  g.save.click_no_wait
  g.jsClick('OK')
  row += 2
  
  #Write  Management Protocol
  puts 'Write Management Protocol '
  g.mgtprot.click
  g.edit.click
  if ws.Range("B#{row}")['Value'].to_s  == 'set' then g.snmp_v1v2.set else g.snmp_v1v2.clear end
  if ws.Range("C#{row}")['Value'].to_s  == 'set' then g.snmp_v3.set else g.snmp_v3.clear end
  snmp_v1v2 = ws.Range("B#{row}")['Value'].to_s
  snmp_v3 =  ws.Range("C#{row}")['Value'].to_s
  g.save.click_no_wait
  g.jsClick('OK')
  row += 2
  
  #Write  SNMP
  puts 'Write SNMP'
  g.snmp.click
  g.edit.click
  if snmp_v1v2 == 'set' || snmp_v3 == 'set'
    g.wrt_checkbox(g.snmp_auth, ws.Range("B#{row}")['Value'])
    g.snmp_hb.select_value((ws.Range("C#{row}")['Value'].to_i).to_s)
    g.wrt_checkbox(g.upsmib, ws.Range("D#{row}")['Value'])

    if g.checkbox(g.upsmib) == 'set'
      g.wrt_checkbox(g.upstraps, ws.Range("E#{row}")['Value'])
    end

    g.wrt_checkbox(g.lgpmib, ws.Range("F#{row}")['Value'])

    if g.checkbox(g.lgpmib) =='set'
      g.wrt_checkbox(g.lgptraps, ws.Range("G#{row}")['Value'])
      g.wrt_checkbox(g.sysnotify, ws.Range("H#{row}")['Value'])
    end

  end
  g.save.click_no_wait
  g.jsClick('OK')
  row += 2
  
  #Write V1 Access
  puts 'Write V1 Access'
  g.v1access.click
  g.edit.click
  for i in 1..20
    g.access_addr(i).set(ws.Range("C#{row}")['Value'].to_s)

    if  ws.Range("D#{row}")['Value'].to_i == 0
      g.access_type((i),'0').set
    elsif ws.Range("D#{row}")['Value'].to_i == 1
      g.access_type((i),'1').set
    end
    g.access_com(i).set(ws.Range("E#{row}")['Value'].to_s)
    row += 1
  end

  g.save.click_no_wait
  g.jsClick('OK')
  row += 1

  #Write V1 Traps
  puts 'Write V1 Traps'
  g.v1traps.click
  g.edit.click
  for i in 1..20
    g.trap_addr(i).set(ws.Range("C#{row}")['Value'].to_s)
    g.trap_port(i).set(ws.Range("D#{row}")['Value'].to_i.to_s)
    g.trap_com(i).set(ws.Range("E#{row}")['Value'].to_s)
    if ws.Range("F#{row}")['Value'].to_s  == 'set' then g.trap_hb(i).set else g.trap_hb(i).clear end
    row += 1
  end
  g.save.click_no_wait
  g.jsClick('OK')
  row += 1
  
  #Write V3 Settings
  puts 'Write V3 Settings'
  g.snmpv3.click
  if snmp_v3 == 'set'
    g.edit.click
    for i in 1..20
      g.wrt_checkbox(g.v3_enable(i), ws.Range("C#{row}")['Value'])
      g.v3_user(i).set(ws.Range("D#{row}")['Value'].to_s)

      if (ws.Range("E#{row}")['Value']).to_i == 0
        g.v3_auth(i,'0').set
      elsif  (ws.Range("E#{row}")['Value']).to_i == 1
        g.v3_auth(i,'1').set
      elsif  (ws.Range("E#{row}")['Value']).to_i == 2
        g.v3_auth(i,'2').set
      end
      g.v3_auth_secret(i).set( ws.Range("F#{row}")['Value'].to_s)
      if (ws.Range("G#{row}")['Value']).to_i == 0
        g.v3_privacy(i,'0').set
      elsif  (ws.Range("G#{row}")['Value']).to_i == 1
        g.v3_privacy(i,'1').set
      end

      g.v3_privacy_secret(i).set( ws.Range("H#{row}")['Value'].to_s)
      g.wrt_checkbox(g.v3_acc_read(i), ws.Range("I#{row}")['Value'])
      g.wrt_checkbox(g.v3_acc_write(i), ws.Range("J#{row}")['Value'])
      g.v3_sources(i).set( ws.Range("K#{row}")['Value'].to_s)
      g.wrt_checkbox(g.v3_notify(i), ws.Range("L#{row}")['Value'])

      if ws.Range("L#{row}")['Value'].to_s == 'set'
        g.v3_destinations(i).set( ws.Range("M#{row}")['Value'].to_s)
        g.wrt_checkbox(g.v3_heartbeat(i), ws.Range("N#{row}")['Value'])
        g.v3_port(i).set( ws.Range("O#{row}")['Value'].to_i.to_s)
      end
      
      row += 1
    end
    g.save.click_no_wait
    g.jsClick('OK')
  else
    row += 20
  end
  row += 1

  #Write Messaging
  puts 'Write Messaging'
  g.msging.click
  g.edit.click
  if ws.Range("B#{row}")['Value'] == 'set' then g.email_msg.set else g.email_msg.clear end
  if ws.Range("C#{row}")['Value'] == 'set' then g.sms_msg.set else g.sms_msg.clear end
  email = ws.Range("B#{row}")['Value'].to_s
  sms = ws.Range("C#{row}")['Value'].to_s
  g.save.click_no_wait
  g.jsClick('OK')
  row += 2

  #Write Email
  puts 'Write Email'
  g.email.click
  g.edit.click
  if email == 'set'
    g.email_from.set(ws.Range("B#{row}")['Value'].to_s)
    g.email_to.set(ws.Range("C#{row}")['Value'].to_s)
    if ws.Range("D#{row}")['Value'] == "0"                 # Subject Type

      g.email_subjectttype(0).set
    else
      g.email_subjectttype(1).set
      g.email_custsubj.set(ws.Range("E#{row}")['Value'].to_s)  # Custom Subject
    end

    g.email_srvr.set(ws.Range("F#{row}")['Value'].to_s)
    g.email_port.set(ws.Range("G#{row}")['Value'].to_i.to_s)
    g.save.click
  end
  row += 2

  #Write Sms
  puts 'Write Sms'
  g.sms.click
  g.edit.click
  if sms == 'set'
    g.sms_from.set(ws.Range("B#{row}")['Value'].to_s)
    g.sms_to.set(ws.Range("C#{row}")['Value'].to_s)

    if ws.Range("D#{row}")['Value'] == "0"                # Subject Type
      g.sms_subjecttype(0).set
    else
      g.sms_subjecttype(1).set
      g.sms_custsubj.set(ws.Range("E#{row}")['Value'].to_s)   # Custom Subject
    end

    g.sms_srvr.set(ws.Range("F#{row}")['Value'].to_s)
    g.sms_port.set(ws.Range("G#{row}")['Value'].to_i.to_s)
    g.save.click
  end
  row += 2
 
  #Write Customize Message
  puts 'Write Customize Message'
  g.custmsg.click
  g.edit.click
  if ws.Range("B#{row}")['Value'] == 'set' then g.email_addr.set else g.email_addr.clear end   #Email
  if ws.Range("C#{row}")['Value'] == 'set' then g.email_evtdesc.set else g.email_evtdesc.clear end
  if ws.Range("D#{row}")['Value'] == 'set' then g.email_name.set else g.email_name.clear end
  if ws.Range("E#{row}")['Value'] == 'set' then g.email_cont.set else g.email_cont.clear end
  if ws.Range("F#{row}")['Value'] == 'set' then g.email_loc.set else g.email_loc.clear end
  if ws.Range("G#{row}")['Value'] == 'set' then g.email_desc.set else g.email_desc.clear end
  if ws.Range("H#{row}")['Value'] == 'set' then g.email_weblnk.set else g.email_weblnk.clear end

  if ws.Range("I#{row}")['Value'] == 'set'
    g.email_consol.set
    g.email_consoltime.set(ws.Range("J#{row}")['Value'].to_i.to_s)
    g.email_consolevt.set(ws.Range("K#{row}")['Value'].to_i.to_s)
  else
    g.email_consol.clear
  end

  if ws.Range("L#{row}")['Value'] == 'set' then g.sms_addr.set else g.sms_addr.clear end  #SMS
  if ws.Range("M#{row}")['Value'] == 'set' then g.sms_evtdesc.set else g.sms_evtdesc.clear end
  if ws.Range("N#{row}")['Value'] == 'set' then g.sms_name.set else g.sms_name.clear end
  if ws.Range("O#{row}")['Value'] == 'set' then g.sms_cont.set else g.sms_cont.clear end
  if ws.Range("P#{row}")['Value'] == 'set' then g.sms_loc.set else g.sms_loc.clear end
  if ws.Range("Q#{row}")['Value'] == 'set' then g.sms_desc.set else g.sms_desc.clear end
  if ws.Range("R#{row}")['Value'] == 'set' then g.sms_weblnk.set else g.sms_weblnk.clear end

  if ws.Range("S#{row}")['Value'] == 'set'
    g.sms_consol.set
    g.sms_consoltime.set(ws.Range("T#{row}")['Value'].to_i.to_s)
    g.sms_consolevt.set(ws.Range("U#{row}")['Value'].to_i.to_s)
  else
    g.sms_consol.clear
  end

  g.save.click_no_wait
  g.jsClick('OK')
  row += 2

  #Write Telnet
  puts 'Write Telnet'
  g.telnet.click
  g.edit.click
  if ws.Range("B#{row}")['Value'] == 'set' then g.telnet1.set else g.telnet1.clear end
  g.save.click_no_wait
  g.jsClick('OK')
  row += 2

  #Write Users
  puts 'Write TelnUserset'
  g.users.click
  g.edit.click
  g.admin_name.set((ws.Range("B#{row}")['Value']).to_s)
  g.admin_pswd.set((ws.Range("C#{row}")['Value']).to_s)
  g.admin_pswd2.set((ws.Range("D#{row}")['Value']).to_s)
  g.user_name.set((ws.Range("E#{row}")['Value']).to_s)
  g.user_pswd.set((ws.Range("F#{row}")['Value']).to_s)
  g.user_pswd2.set((ws.Range("G#{row}")['Value']).to_s)
  g.save.click_no_wait
  g.jsClick('OK')
  row += 2

  #Write Web
  puts 'Write Web'
  g.cfgweb.click
  g.edit.click
  g.websrvr.select_value((ws.Range("B#{row}")['Value'].to_i).to_s)
  
  if ws.Range("B#{row}")['Value'].to_i == 2
    g.httpport.set((ws.Range("C#{row}")['Value']).to_i.to_s)
  elsif ws.Range("B#{row}")['Value'].to_i == 3
    g.httpsport.set((ws.Range("D#{row}")['Value']).to_i.to_s)
  end

  unless (ws.Range("B#{row}")['Value'].to_i).to_s == "1"
    g.wrt_checkbox(g.pswdprtct, ws.Range("E#{row}")['Value'])
    g.wrt_checkbox(g.cfgctrl, ws.Range("F#{row}")['Value'])
    g.refresh.set((ws.Range("G#{row}")['Value']).to_i.to_s)
  else
    puts "The Web Server is Disabled"
  end
  g.save.click_no_wait
  g.jsClick('OK')

 
  $ie.close
  wb.save
  wb.close
  puts 'Finished'
end
