=begin
*Script_Name*
  Controller
    The controller is used to run suites of driver scripts.

*Suite_Number*
  pending

*Description*
  Validate Agent Table Information

*Variables*
    s = test start time
    f = test finish time
    excel = nested array that contains an instance of excel and script parameters
      excel[0] = excel instance
        ss = spreadsheet[0]
        wb = workbook[1]
        ws = worksheet[2]
      excel[1] = parameters
        ctrl_ss = controller spreadsheet[0]
        rows = number of rows in spreadsheet to execute[1]
        site = url/ip address of card being tested[2]
        name = user name for login[3]
        pswd = password for login[4]
    row = incremented each cycle, indicates the row that is being executed
    rows = number of spreadsheet rows to iterate ; cell 'B2'
    r_ow = Sequential number of script located in the controller spreadsheet
    run_flag = determines if script in row will run based on check in column 'D'
    
*Methods*
TODO move the method info to the appropriate lib files
    ss_info - location: setup module
      - create time stamped controller spreadsheet
      - open IE or attache to existing IE session
      - populate the spreadsheet with web support page info

    conn_act_xls - location: xls module
      - connect to the active instance of excel

    tear_down_c - location: teardown module
      - writes the finished and elapsed time to the controller spreadsheet
      - save and close the controller spreadsheet
=end

$:.unshift File.expand_path(File.dirname(__FILE__)).sub('controller','lib') #add lib to load path
$:.unshift '.'

require 'generic'
s = Time.now

RUN_COLUMN = 'E'
SCRIPT_COLUMN = 'J'

begin
  #Check for existing instances of IE and kill them if present
  if is_process_running?('iexplore.exe')
    kill_processes(get_process_pids('iexplore.exe'))
    puts "Found and terminated existing instances of IE"
  end
  puts" \n Executing: #{(__FILE__)}\n\n" #current file
  g = Generic.new
  
  excel = g.setup(__FILE__)
  ws = excel[0][2]
  ctrl_ss,rows,site,name,pswd = excel[1]

  g.config.click    # login now so drivers won't have to
  g.login(site,name,pswd)
 
  row = 2
  while (row <= rows)
    run_flag = ws.Range("#{RUN_COLUMN}#{row}")['Value']
    if run_flag == true
      print" Executing Driver script #{row-1} -- "
      path = File.dirname(__FILE__).sub('controller','driver/') # driver path 
      drvr = path << (ws.Range("#{SCRIPT_COLUMN}#{row}")['Value'].to_s) #concat path and driver
      drvr << '.rb' if !(drvr =~ /\.rb$/) #Add .rb to script name if necessary.
      t = Time.now.to_a.reverse[5..9].to_s
      log = (drvr.gsub('.rb',"-#{t}.log" )).sub('driver','result')
      run_drvr = system "ruby #{drvr} #{ctrl_ss} #{row}"# > {log}" # run driver 
      puts "Driver Status = #{run_drvr}"
      #g.conn_act_xls # connect to controller spreadsheet
    end
    row += 1
  end
  f = Time.now  
  g.tear_down_c(excel[0],s,f)
  
rescue Exception => e
  puts" \n\n **********\n\n #{$@ } \n\n #{e} \n\n ***"
ensure 
end

