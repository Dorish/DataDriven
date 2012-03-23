=begin

*Description*
  The controller is used to run suites of driver scripts.

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

TODO driver log files are currently disabled
=end

$:.unshift File.dirname(__FILE__).sub('controller','lib') #add lib to load path
require 'generic'
s = Time.now

# User select one from the existing test suites to execute.
def select_test_suite(path)
  # search the directory and create the hash with index and spreadsheet names.
  fl_list = Dir.entries(path).delete_if{ |e| e=~ /^\..*/|| e=~/^.*\.rb/}
  while 1
    index = 0
    puts "The following test suites are available for execution:"
    fl_list.each do |i|
      print index+1,' - ',i.chomp(".xls"),"\n"
      index = index+1
    end
    puts "Please type the number of the desired suite followed by <Enter>"
    input_str = gets.chomp.to_i
    if fl_list.include?(fl_list[input_str-1])
      break
    end
  end
  return fl_list[input_str-1] # return the spreadsheet name
end


begin
  g = Generic.new
  # select test suite
  contr_dir = File.dirname(__FILE__)
  test_suite = select_test_suite(contr_dir)
  exec_path = contr_dir + '/' + test_suite
  puts "Executing: #{exec_path} now"
  
  # create time stamped result folder
  rs_folder = g.timeStamp(test_suite.chomp('.xls'))
  original_dir = Dir.pwd
  Dir.chdir(contr_dir.gsub("controller","result")) # Change DIR to result folder
  Dir.mkdir(rs_folder)
  Dir.chdir(original_dir)# Change DIR back to the original
  
  setup = g.setup(exec_path.chomp('xls'),rs_folder)# chomp('xls')-avoid duplicate of 'xls' in setup method.
  xl = setup[0]
  ws = xl[2] # spreadsheet
  ctrl_ss,rows,site,name,pswd = setup[1]

  # login now so drivers won't have to
  # This web login is not necessary to telnet run
  g.config.click    
  g.login(site,name,pswd)
  g.equipinfo.click
 
  row  = 1
  while (row <= rows)
    row += 1
    # run driver if the 'run' box is checked
    if ws.Range("e#{row}")['Value'] == true
      print" Run driver script #{row - 1} -- "
      path = File.expand_path('driver')
      if test_suite.include?'telnet'
        drvr = path << '/telnet/telnet_prototype.rb' << ' ' << ws.Range("j#{row}")['Value'].to_s# Two arguments, telnet agent path and driver name
      else
        drvr = path << (ws.Range("j#{row}")['Value'].to_s) # watir driver path
      end
      log = (drvr.gsub('.rb',"-#{g.t_stamp}.log" )).sub('driver','result')
      system "ruby #{drvr} #{ctrl_ss} #{row} #{rs_folder}"# > {log}" # run driver # add rs_folder as an ARGV to pass to drivers.
      g.conn_act_xls # reconnect to controller spreadsheet
    end
  end
  f = Time.now  
  g.tear_down_c(xl,s,f)
  
rescue Exception => e
  puts" \n\n **********\n\n #{$@ } \n\n #{e} \n\n ***"
end