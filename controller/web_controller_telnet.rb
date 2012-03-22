$:.unshift File.dirname(__FILE__).sub('controller','lib') #add lib to load path
require 'generic'
s = Time.now
  
begin
  g = Generic.new
  # create time stamped result folder
  contr_dir = File.dirname(__FILE__)
  test_suite = (__FILE__).sub(/#{contr_dir}\//, '') # Get the exactly filename
  rs_folder = g.timeStamp(test_suite.chomp('.rb')) # time stamped result folder name
  original_dir = Dir.pwd
  Dir.chdir(contr_dir.gsub("controller","result")) # Change DIR to result folder
  Dir.mkdir(rs_folder)
  Dir.chdir(original_dir)# Change DIR back to the original

  setup = g.setup(__FILE__,rs_folder)# chomp('xls')-avoid duplicate of 'xls' in setup method.
  xl = setup[0]
  ws = xl[2] # spreadsheet
  ctrl_ss,rows = setup[1]

  row = 1
  while (row <= rows)
    row += 1
    # run driver if the 'run' box is checked
    if ws.Range("e#{row}")['Value'] == true
      print" Run driver script #{row - 1} -- "
      path = File.expand_path('driver')
      drvr = path + '/telnet/telnet_prototype.rb' # driver path
      sprt = ws.Range("j#{row}")['Value'].to_s
      log = (drvr.gsub('.rb',"-#{g.t_stamp}.log" )).sub('driver','result')
      system "ruby #{drvr} #{sprt} #{ctrl_ss} #{row} #{rs_folder}"# > {log}" # run driver # add rs_folder as an ARGV to pass to drivers.
      g.conn_act_xls # reconnect to controller spreadsheet
    end
  end
  f = Time.now
  g.tear_down_c(xl,s,f)
end