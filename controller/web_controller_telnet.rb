require 'net/telnet'
require 'win32ole'

$:.unshift File.dirname(__FILE__).sub('controller','lib') #add lib to load path
require 'generic'
#begin this flags used to decrease the amount of reset --add by Ryan
$_flag_rd = 0
$_flag_wr = 0
$_flag_cfg = 0
#end
begin

 #Get the path of telnet spreedsheet files from controller_telnet.xls --add by Ryan
def getpath()
  g = Generic.new

  setup = g.setup(__FILE__)
  xl = setup[0]
  ss = xl[0]
  wb = xl[1]
  ws = xl[2] # spreadsheet
  ctrl_ss,rows,site,name,pswd = setup[1]

  drvr_path = Array.new
  row = 1
  while (row <= rows)
  row += 1

  if ws.Range("e#{row}")['Value'] == true
    path = File.dirname(__FILE__).sub('controller','driver')
    drvr = path << (ws.Range("j#{row}")['Value'].to_s)
    drvr_path << drvr
  end
end
wb.save
wb.close
ss.quit
 return drvr_path
 end

def main_menu(telnet)
  esc = "\x1b"
  flag = true
  main_menu_regex = 'Main Menu'
  Thread.new {
    while flag == true
      telnet.write(esc){|c| print c}
      sleep(1)
    end
  }
  while flag == true
    telnet.waitfor('String' => main_menu_regex, 'waittime' => 1){|c| print c}
    flag = false
  end
end

#wait until the device has been discovered?
def wait_for_reboot(ip)
  puts "\n\n"
  flag = true
  reply_from = Regexp.new('Reply from')
  while flag == true
    puts "Waiting for device to reboot"
    sleep(20)
    results = `ping #{ip}`
    puts results
    if (reply_from.match(results)) then
      puts "Device is booting..."
      sleep(30)
      flag = false
    end
  end
  puts "\n\n"
end

#Reset the card to factory default
def reset_factory_defaults(ip,prompt,login_name, login_password)
   puts "Start to reset ......"
   telnet = Net::Telnet.new('Host' => ip, 'Prompt' => prompt, "Output_log" => "data_input_log.txt")

      login(telnet, login_name, login_password)
      telnet.waitfor(/\?>/) {|c| print c}

      default_s ="4,2"
      navigation = default_s.split(',')
      navigation.each do |s|
      telnet.write(s) {|c| print c}
      telnet.waitfor(/\?>/) {|c| print c}
       end
       telnet.write('y') {|c| print c}
       wait_for_reboot(ip)
 end
# Method abstraction for logging into a ENP Intellislot card
def login(telnet, login_name, login_password)

#Send carriage return to bring up login in prompt
telnet.write($cr){|c| print c}

#Wait for login prompt
telnet.waitfor(/[Ll]ogin[: ]*\z/) {|c| print c}

#Send username
telnet.cmd(login_name) {|c| print c} #Use cmd when you want to send CR, use write when you dont

#Wait for password prompt
telnet.waitfor(/Password[: ]*\z/) {|c| print c}

#Send password
telnet.cmd(login_password) {|c| print c}

end

#
def input_data(telnet, ws, start_row, end_row)
	while (start_row < end_row) do
	  if ws.Range("H#{start_row}")['Value'] == 'IGNORE'
		start_row += 1
		next
	  end
	  navigation_string = ws.Range("F#{start_row}")['Value']
	  reg_ex_cr = Regexp.new('\{enter\}')
	  navigation = navigation_string.split(',')
	  navigation.each do |s|
		if reg_ex_cr.match(s) then
		  telnet.write($cr)
		  telnet.waitfor(/\?>/) {|c| print c}
		  next
		end
		telnet.write(s) {|c| print c}
		##Don't wait for a new prompt if telnet is waiting for a carriage return - forgive the disgusting syntax...
		if (reg_ex_cr.match(navigation[navigation.rindex(s)+1]) or reg_ex_cr.match(navigation[navigation.size-2])) then next; end;
		telnet.waitfor(/\?>|Hit/) {|c| print c} # The Hit is for the "Hit any key prompts..."
	  end

	  command_string = ws.Range("G#{start_row}")['Value'].to_s
	  command = command_string.split(',')
	  if command_string =~ /enter/ then
		command.each do |s|
		  puts "\nCommand String is: #{s}\n"
		  if reg_ex_cr.match(s) then
			puts 'We got an enter up in here!';
			telnet.write($cr)
			telnet.waitfor(/\?>|Hit/) {|c| print c}
			next
		  end
		  telnet.write(s.to_s) {|c| print c}
		end
	  else
		command.each do |s|
		  puts "\nCommand String is: #{s}\n"
		  telnet.write(s) {|c| print c}
		end
	  end

	  main_menu(telnet) #Go back to the main menu
	  start_row += 1
	end
end

#
def verify_data(telnet,ws, start_row, end_row)
	while (start_row < end_row) do
	  if ws.Range("H#{start_row}")['Value'] == nil or ws.Range("H#{start_row}")['Value'] == 'IGNORE' then
		puts "\nExpected a value at H#{start_row} for test #{ws.Range("E#{start_row}")['Value']} skipping this verification..."
		start_row += 1
		next
	  end

	  navigation_string = ws.Range("F#{start_row}")['Value']
	  reg_ex_data_label = Regexp.new(ws.Range("H#{start_row}")['Value'])
	  reg_ex_cr = Regexp.new('\{enter\}')
	  navigation = navigation_string.split(',')

	  navigation.each do |s|
		buffer = ''
		#if ws.Range("J#{start_row}")['Value'] != nil then puts "Broken!"; break; end; #Not sure why this is here.
		if reg_ex_cr.match(s) then
		  telnet.write($cr)
		  telnet.waitfor(/\?>/) {|c| print c; buffer += c;}
		  next
		end
		telnet.write(s) {|c| print c}
    #Used to slow down the navigation -- add by Ryan
      sleep(1)
		##Don't wait for a new prompt if telnet is waiting for a carriage return - forgive the disgusting syntax...
		#if (reg_ex_cr.match(navigation[navigation.rindex(s)+1]) or reg_ex_cr.match(navigation[navigation.size-2])) then next; end;
		telnet.waitfor(/\?>/) {|c| print c; buffer += c;}
		buffer.each_line do |line|
		  if (reg_ex_data_label.match(line)) then
			actual_value = ''
			line = line.sub!(reg_ex_data_label, '')
			line = line.split(" ")
			0.upto(line.size-1) { |i| actual_value << line[i] << " " }
			ws.Range("J#{start_row}")['Value'] = actual_value.chomp(" ")
		  end
		end
	  end

	  #Go back to the main menu
	  main_menu(telnet)
	  start_row += 1
	end
end

# Returns an array containing the row numbers in which a reboot must occur
def identify_reboots(ws, column='F', reboot=/REBOOT/i)
	xldown = -4121 #Constant used by excel
	total_rows = ws.Range("#{column}:#{column}").End(xldown).row
	reboots = Array.new
	for row in 2..total_rows
		if ws.Range("#{column}#{row}").Value =~ reboot
			reboots << row
		end
	end
	reboots
end

# Attempts to reboot (by sending an 'x') at the current menu
def reboot(telnet)
	telnet.cmd('x') {|c| print c}
end


# Execute telnet script test
def singletelnet(telname)
      #base_ss = File.expand_path(File.dirname(__FILE__)) + '/' + telname
      base_ss = telname
      new_ss = (base_ss.chomp(".xls")<<'_'<<Time.now.strftime("%m-%d_%H-%M-%S")<<(".xls")).gsub('driver/telnet','result')
      ss = WIN32OLE::new('excel.Application')
      ss.DisplayAlerts = false #Stops excel from displaying alerts
      ss.Visible = true # For debug
      wb = ss.Workbooks.Open(base_ss)
      new_ss.gsub!('/','\\')
      wb.SaveAs(new_ss)
      ws = wb.Worksheets(1)

      ip = ws.Range("B3")['Value'].to_s #Define Test Site
      login_name = ws.Range("B4")['Value'].to_s #Define Login Name
      login_password = ws.Range("B5")['Value'].to_s #Define Password
      $cr = "\x0D" # Ascii Hex code for Carriage Return is 0D
      #begin regex to match files
      reg_ex_rd = Regexp.new('def\_rd')
      reg_ex_wr = Regexp.new('def\_wrt')
      reg_ex_cfg = Regexp.new('cfg\_')
      #end
     puts ""
     puts"Open telnet connection to device"

      #regex to match prompt
      prompt=/[:>]*/n

  #begin used to decrease the amount of reset -- add by Ryan
     if reg_ex_rd.match(base_ss)
      if $_flag_rd == 0
        reset_factory_defaults(ip,prompt,login_name, login_password)
        $_flag_rd = 1
        puts "flag_rd changed #{$_flag_rd }"
      else
        puts "flag_rd already done #{$_flag_rd }"

      end
    end
    if reg_ex_wr.match(base_ss)
      if $_flag_wr == 0
        reset_factory_defaults(ip,prompt,login_name, login_password)
        $_flag_wr = 1
        puts "flag_wr changed #{$_flag_wr }"
      else
        puts "flag_wr already done #{$_flag_wr }"
      end
   end
   if reg_ex_cfg.match(base_ss)
      if $_flag_cfg == 0
         reset_factory_defaults(ip,prompt,login_name, login_password)
        $_flag_cfg= 1
        puts "flag_cfg changed #{$_flag_cfg }"
      else
        puts "flag_cfg already done #{$_flag_cfg }"

      end
    end
  #end
      puts"Repen telnet connection to device"

      telnet = Net::Telnet.new('Host' => ip, 'Prompt' => prompt, "Output_log" => "data_input_log.txt")

      login(telnet, login_name, login_password)
      telnet.waitfor(/\?>/) {|c| print c}

      # Read in the number of rows to execute
      rows = ws.Range("B2")['Value']

      # Start at row 2
      current_row = 2
      total_rows = current_row + rows.to_i

      # Identify reboots in 'Navigation' column
      reboots = identify_reboots(ws, 'F')

      #Input data and verify after each reboot
      reboots.each do |reboot_row|
        input_data(telnet, ws, current_row, reboot_row)
        reboot(telnet)
        wait_for_reboot(ip)
      #begin telnet retry method --add by Ryan
      tries = 0
      begin # retry to connect to the card via telnet --- add by Ryan
          telnet = Net::Telnet.new('Host' => ip,'Prompt' => prompt, "Output_log" => "data_verification_log.txt")
      rescue
          tries += 1
          puts "we have tried #{tries}"
         sleep(10)
      retry if (tries <= 9)
         puts "retry limit reached!"
      end
      #end
        login(telnet, login_name, login_password)
        verify_data(telnet, ws, current_row, reboot_row)
        current_row = reboot_row + 1
      end

      if reboots.size == 0
        input_data(telnet, ws, current_row, total_rows)
        reboot(telnet)
        wait_for_reboot(ip)
        telnet = Net::Telnet.new('Host' => ip,'Prompt' => prompt, "Output_log" => "data_verification_log.txt")
        login(telnet, login_name, login_password)
        verify_data(telnet, ws, current_row, total_rows)
      end
    
      puts "\n\nTest complete!\nStatus is #{ws.Range("B16")['Value']}"
      statusflag= ws.Range("B16")['value'] #Track the status-----Gary
      rescue Exception => e
      puts "Script failed on row: #{current_row}"
      puts $@.to_s #Array of backtrace
      puts $! #Exception Information
      ensure
      wb.save
      wb.close
      ss.quit

  return statusflag #return the status of tn.xls file , Pass or Fail-----Gary

 end

 telarray=Array.new
 telarray= getpath()
 telstatus=Hash.new
#Execute the telnet script by order

#for j in 0...110
 for i in 0...telarray.length
  telstatuss=telarray[i].chomp(".xls")
  telstatus[telstatuss]=singletelnet(telarray[i])
 end
 #puts "now we are in"
 #puts j
#end

end

puts telstatus.inspect # Output tn xls files Name and Status, Pass or Fail-----Gary


