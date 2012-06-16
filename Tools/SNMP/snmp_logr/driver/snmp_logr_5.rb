=begin rdoc

This version
Collect all oids into an array and fill the 'alldata' array in one go.
All of the snmp data points to an array and then the array
is written to the spreadsheet in one line. There is no timing improvement
over writing each data point to the spreadsheet one at a time.

mib         = MIB
com_str     = community string
rows        = number OIDs to GET
ip          = IP address
iter        = number of iterations
dly         = delay between iterations in seconds
row         = oid rows
rows        = total number of oids to execute
data_row    = snmp data rows
data_col    = snmp data columns
data        = an array that contains all of the snmp data

a mapping of the columns to oids is contained in spreadsheet column 'M'


=end


##TODO Change to send result the standard test result folder.
##TODO rename the script name to snmp_logging to be consistent with other logging


s = Time.now

require 'win32ole'

#    - create time stamped controller spreadsheet
#    - open IE or attach to existing IE session
#    - populate the spreadsheet with web support page info
def setup(file)
  base_xl = (file).gsub('/','\\').chomp('rb')<<'xls'
  excel = xls_timestamp(base_xl) # timestamped instance of excel
 end

#    - create time stamped spreadsheet using base name
#    - connect to excel and open base workbook
#    - create instance of excel (xl)
#    - returns a nested array containing spreadsheet and script parameters
def xls_timestamp(s_s)
  new_ss = (s_s.chomp(".xls")<<'_'<<Time.now.to_a.reverse[5..9].to_s<<(".xls"))
  new_ss1 = new_ss.sub('driver','result')
  xl = new_xls(s_s,1) #open base driver ss with new excel session
  ws = xl[2] # worksheet
  param = Array.new # contains no elements. just used as a place holder here
  excel = [xl,param]

  # save spreadsheet as timestamped name.
  save_as_xls(xl,new_ss1)
  return excel
end

#
#  - createand return new instance of excel
def new_xls(s_s,num) #wb name and sheet number
  ss = WIN32OLE::new('excel.Application')
  wb = ss.Workbooks.Open(s_s)
  ws = wb.Worksheets(num)
  ss.visible = true # For debug
  xls = [ss,wb,ws]
end

#
#  - save an existing workbook as another file name
def save_as_xls(s_s,save_as)
  sleep 1
  s_s[2].saveas(save_as)
end

#
#  - snmpget using Net-Snmp
def snmpget(com,ip,mib,oid)
  command = 'snmpget -v2c -c '<< com << ' ' << ip << ' ' << mib << '::' << oid
  puts"   #{command}" ;  $stdout.flush
  return `#{command}`.split(/ /)[3] # value is in 4th element
end
#
#  - teardown driver - this function will update driver spreadsheet.
def tear_down_d(s_s,s,f)
  # The variable 's_s' is an array that holds the current spreadsheet instance
  ss,wb,ws = s_s
  #Save Summary and elapsed time to current ss
  ws.Range("b8")['Value'] = s.to_s
  ws.Range("b9")['Value'] = f.to_s
  run_time = elapsed(f,s)
  ws.Range("b10")['Value'] = run_time.to_s
  wb.save
  wb.close #Close Driver spreadsheet
end

#
#  - calculates difference between start and finish time
def elapsed(finish,start)
  time = (finish-start).to_i
  hours = time/3600.to_i
  minutes = (time/60 - hours * 60).to_i
  seconds = (time - (hours * 3600 + minutes * 60)).to_i
  test_time  = minutes.to_s << 'min'<<seconds.to_s<<'sec'
  puts "\n\nTest Start  = "<<start.strftime("%H:%M:%S")
  puts "Test Finish = "<<finish.strftime("%H:%M:%S")
  puts "Test Time   = #{minutes}min#{seconds}sec"
  return test_time
end


puts" \n Executing: #{(__FILE__)}\n\n" # show current filename
excel = setup(__FILE__)
wb,ws = excel[0][1,2]

ip,mib,com      = ws.Range("b2:b4")['Value'].map{|x|x.to_s}
rows,iter,dly   = ws.Range("b5:b7")['Value'].map{|x|x.to_s.to_i}

data_row = 2 # first data row
while data_row <= iter +1
  row = 2 # first oid row
  data_col = "ae" #  is first data column -1
  data = [] # (re)initialize array
  # Collect all of the snmp data into a single array
  ws.range("k#{row}:k#{rows +1}")['Value'].each do|oid|
    data = data.to_a + snmpget(com,ip,mib,oid.to_s).split(/ /).push
    data_col = data_col.next
  end
  # write all of the snmp data to a range of cells
  ws.Range("af#{data_row}:#{data_col}#{data_row}")['Value'] = data
  ws.Columns("af:#{data_col}").Autofit
  wb.save
  data_row +=1
  sleep (dly)if data_row <= iter +1 # sleep each iteration except last
end
f = Time.now
tear_down_d(excel[0],s,f)


