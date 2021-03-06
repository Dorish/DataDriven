=begin rdoc
*Revisions*
  | Change                                               | Name        | Date  |

*Module_Name*
  Setup

*Description*
  Test script setup methods

*Variables*
    s = test start time
    f = test finish time
    roe = row number in controller spreadsheet if called from controller
    excel = nested array containing an instance of excel and script parameters
      excel[0] = excel instance
        ss = spreadsheet (excel[0][0])
        wb = workbook (excel[0][1])
        ws = worksheet (excel[0][2])
      excel[1] = parameters
        dvr_ss = driver spreadsheet (excel[1][0])
        rows = number of rows in spreadsheet to execute (excel[1][1])
        site = url/ip address of card being tested (excel[1][2])
        name = user name for login (excel[1][3])
        pswd = password for login (excel[1][4])
    ARGV[1] = command line parameter passed from the controller for 'roe'
    (ARGV.length !=0) script called from controller else running independently

=end

module  Setup

  #
  # - time stamp in 'month-day_hour-minute-second' format
  def t_stamp
    Time.now.strftime("%m-%d_%H-%M-%S")
  end


  #    - create time stamped controller spreadsheet
  #    - open IE or attach to existing IE session
  #    - populate the spreadsheet with web support page info
  def setup(file,rs_name = nil)
    systemos      #Determine whether the OS is Chinese or English
    base_xl = (file).gsub('/','\\').chomp('rb')<<'xls'
    if(ARGV.length != 0)          # called from controller
      excel = xls_timestamp(base_xl,nil,ARGV[2]) # called ,connect to existing excel instance. # ARGV[2] is the result sub-folder name.
      attach_ie(excel[1][2])  # test site ip
    else
      excel = xls_timestamp(base_xl,'ind',rs_name) # independent, start new excel instance
      open_ie(excel[1][2])
      support(excel[0])
      ver_info(excel[0])
   
    end
    return excel
  end


  #    - create time stamped spreadsheet using base name
  #    - connect to excel and open base workbook
  #    - create instance of excel (xl)
  #    - returns a nested array containing spreadsheet and script parameters
  def xls_timestamp(s_s,type=nil,rs_name=nil)
    new_ss = (s_s.chomp(".xls")<<'_'<<t_stamp<<(".xls"))
    # common statement assigned with one variable
    if new_ss.include? "controller"
      new_ss1 = new_ss.sub('controller',"result\\#{rs_name}")
      xl = new_xls(s_s,1) #open base controller ss with new excel session
    elsif new_ss.include? "driver"
      new_ss1 = new_ss.sub(/driver\\.+\\/,"result\\#{rs_name}\\")
      if (type == 'ind') # driver was launched independently
        xl = new_xls(s_s,1) #open base driver ss with new excel session
      else # driver was launched by controller
        xl = conn_open_xls(s_s,1) #connect to existing excel session and create new workbook for L2
      end
    elsif new_ss.include? "Tools"
      new_ss1 = new_ss.sub(/Tools\\.+\\/,"result\\#{rs_name}\\")
      xl = new_xls(s_s,1)
    end

    ws = xl[2] # worksheet

    param = Array.new
    param[0] = new_ss1
    param[1] = ws.Range("b2")['Value'].to_i        # rows
    param[2] = ws.Range("b3")['Value']             # Test_site
    param[3] = ws.Range("b4")['Value']             # username
    param[4] = ws.Range("b5")['Value']             # password

    # This is a nested array that contains the instance of excel
    # and the spreadsheet parameters listed directly above
    excel = [xl,param]

    # save spreadsheet as timestamped name.
    save_as_xls(xl,new_ss1)
    return excel
  end


  #
  #    - create new IE instance and navigate to the test site
  def open_ie(site)
    puts "\n    **Open IE **\n"
    $ie = Watir::IE.new
    $ie.goto(site)
  end


  #
  #     - attach to existing IE instance and navigate to the test site
  def attach_ie(site)
    puts "\n    **Attach to IE **\n"
    site = 'http://'<<site<<'/'
    $ie = Watir::IE.attach(:url, site)
  end


  #
  #    - navigate with 'click'; script is called from controller; no login req'd
  #    - navigate with 'click_no_wait'; script called directly; login req'd
  def logn_chk(nav,logn)
    site,name,pswd = logn[2,4] #logn is an array
    login(site,name,pswd) if (ARGV.length == 0)
    nav.click
  end


  def systemos
    lang = `systeminfo`
    if lang =~ /en-us*/
      @@os          = "English"
      @@titl          = "Connect to "
      @@ok       ="OK"
      @@cancel    = "Cancel"
    elsif lang =~ /zh-cn*/
      @@os           = "Chinese"
      @@titl           = "连接到 "
      @@ok        ="确定"
      @@cancel      = "取消"
    end
    puts "This OS is #{@@os}"
  end

  
  #    - collects support page table:
 
  def support(xl)
    puts "  Collect Support page info"
    supp.click
    sleep 1
    wb,ws = xl[1,2]
    row = 1
    while ws.Range("A#{row}")['Value'] != nil # find the row to start recording support info.
      row = row + 1
    end
    supprt.each do|key|
      if !key[0].nil?
        c = ws.range("A#{(row)}:B#{(row)}")
        c.value = key
        c.Interior['ColorIndex'] = 19   # change background color
        c.Borders.ColorIndex = 1        # add border
        #ws.range("A#{row}:B#{row}").Columns.Autofit
        row+=1
      end
    end
    os = ws.range("A#{(row)}:B#{(row)}") #add system os info to ss
    os.value = ["Operating System Language","#{@@os}"]
    os.Interior['ColorIndex'] = 43  # change background color
    os.Borders.ColorIndex = 1        # add border
    ws.range("A:B").ColumnWidth = 255 #255 is the maximum column width
    ws.range("A:B").Rows.Autofit
    ws.range("A:B").Columns.Autofit
    wb.Save
  end

  def ver_info(xl)
    version =[ ] #Array used to put the version we collect
    wb,ws = xl[1,2]
    row = 1
    while ws.Range("A#{row}")['Value'] != nil # find the row to start recording version info.
      row = row + 1
    end

    #collect the Ruby version
    puts "  Collect the Ruby version"
    r_version = `ruby -v`.chomp #get the Ruby version
    puts "  This Ruby version is #{r_version}"
    version << ["Ruby version","#{r_version}"]

    #collect the Watir version
    puts "  Collect the Watir version"
    w_version = `ruby -e 'require "watir";puts Watir::VERSION'`.chomp
    puts "  This Watir version is #{w_version}"
    version << ["Watir version","#{w_version}"]

    #collect the IE version
    puts "  Collect the IE version"
    temp =  `reg query \"HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Internet Explorer\" \/v version"`
    v_version = /version\s*[A-Z_]+\s*([0-9.]*)/.match(temp)
    puts "  This IE version is #{v_version[1]} "
    version << ["IE version","#{v_version[1]}"]
    
    #collent the git SHA-1 hash
    puts "  Collect the git SHA-1 hash"
    g_log = `git log -1` # Need to make sure 'C:\Program Files\Git\cmd' is in the environment variable and REBOOT PC.
    g_version = /[0-9a-z]{40}/.match(g_log).to_s
    puts "  This current git SHA-1 hash is #{g_version[0..6]}"
    version << ["git SHA-1 hash","#{g_version[0..6]}"]

    #write the version in ss
    version.each{|x|
      ver = ws.range("A#{(row)}:B#{(row)}") #add version info to ss
      ver.value = x
      row += 1
      ver.Interior['ColorIndex'] = 43  # change background color
      ver.Borders.ColorIndex = 1        # add border
    
    }
    ws.range("A:B").ColumnWidth = 255 #255 is the maximum column width
    ws.range("A:B").Rows.Autofit
    ws.range("A:B").Columns.Autofit
    wb.Save
  end

  def snmp_setup(wb)
    ws = wb.Worksheets('SupportInfo')
    @test_site = ws.range('B3').value
    @community_string = ws.range('B5').value # Use the password field for the community string
  end

  #  - Add a time stamp to anything that is passed in as 'variable'
  #  - Need to accommodate two input conditions -
  #   1) ‘variable' => 'variable’_01-05_14-58-45'
  #   2) ‘variable.ext’=> 'variable_01-05_14-58-45.ext', where 'ext’can be any extension.
  def timeStamp(vari)
    ext = /\.\w*$/.match(vari).to_s # match extension from end of the string
    if ext
      vari.chomp(ext)+'_'+t_stamp+ext
    else
      vari+'_'+t_stamp
    end
  end

end
