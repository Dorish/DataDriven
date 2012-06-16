=begin rdoc


mib         = MIB; par1[0]
com_str     = community string; par1[1]
rows        = number OIDs to GET; par1[2]
ip          = IP address; par2[0]
iter        = number of iterations; par2[1]
dly         = delay between iterations in seconds; par2[2]

par1 = array contains string parameters
par2 = array containing integer parameters


a mapping of the columns to oids is contained in spreadsheet column 'M'

=end




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
`#{command}`.split(/ /)[3] # value is in 4th element
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
puts excel
puts "element 0 #{excel[0]}"
puts "element 0 #{excel[1]}"
puts "element 0 #{excel[2]}"
wb,ws = excel[0][1,2]

ip,mib,com      = ws.Range("b2:b4")['Value'].map{|x|x.to_s}
rows,iter,dly   = ws.Range("b5:b7")['Value'].map{|x|x.to_s.to_i}

data_row = 2 # first data row, increment after data columns are filled
while data_row <= iter +1
  row = 2 # first oid row
  data_col = "af" #  'af' is first data column
  while (row <= rows +1)
    oid  = ws.range("k#{row}")['Value']
    data = snmpget(com,ip,mib,oid)
    ws.Range("#{data_col}#{data_row}")['Value'] = data
    data_col = data_col.next
    row += 1
  end
  ws.Columns("af:#{data_col}").Autofit
  wb.save
  data_row +=1
  sleep (dly)if data_row <= iter +1 # sleep each iteration except last
end
f = Time.now
tear_down_d(excel[0],s,f)


