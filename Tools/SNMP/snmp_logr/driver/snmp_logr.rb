s = Time.now
require 'win32ole'
#require 'generic'
 #    - create time stamped controller spreadsheet
  #    - open IE or attach to existing IE session
  #    - populate the spreadsheet with web support page info
  def setup(file)
    base_xl = (file).gsub('/','\\').chomp('rb')<<'xls'
    excel = xls_timestamp(base_xl,'ind') # independent, start new excel instance
  end


  #    - create time stamped spreadsheet using base name
  #    - connect to excel and open base workbook
  #    - create instance of excel (xl)
  #    - returns a nested array containing spreadsheet and script parameters
  def xls_timestamp(s_s,*type)
    new_ss = (s_s.chomp(".xls")<<'_'<<Time.now.to_a.reverse[5..9].to_s<<(".xls"))
    new_ss1 = new_ss.sub('driver','result')
    xl = new_xls(s_s,1) #open base driver ss with new excel session  
    ws = xl[2] # worksheet

    param = Array.new # contains no elements. just used as a place holder here

    # This is a nested array that contains the instance of excel
    # and the spreadsheet parameters listed directly above
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
  #  - teardown driver - this function will update driver spreadsheet.
  def tear_down_d(s_s,s,f,roe)
    # The variable 's_s' is an array that holds the current spreadsheet instance
    ss,wb,ws = s_s
    #Save Summary and elapsed time to current ss
    ws.Range("b8")['Value'] = s.to_s
    ws.Range("b9")['Value'] = f.to_s
    run_time = elapsed(f,s)
    ws.Range("b10")['Value'] = run_time.to_s
    status = ws.Range("b16")['Value'].to_s # Pass / Fail from Driver.xls
    puts "Status      = #{status}"
    wb.save
    wb.close #Close Driver spreadsheet
  end


  #
  #  - calculates difference between start and finish time
  def elapsed(finish,start,*row)
    time = (finish-start).to_i
    hours = time/3600.to_i
    minutes = (time/60 - hours * 60).to_i
    seconds = (time - (hours * 3600 + minutes * 60)).to_i
    if(row != 0 ) # If driver script - use min and sec
      test_time  = minutes.to_s << 'min'<<seconds.to_s<<'sec'
      puts "\n\nTest Start  = "<<start.strftime("%H:%M:%S")
      puts "Test Finish = "<<finish.strftime("%H:%M:%S")
      puts "Test Time   = #{minutes}min#{seconds}sec"
    else # Default output (Controller script) use hrs, min, and sec
      test_time  = hours.to_s << 'hr' <<minutes.to_s << 'min'<<seconds.to_s<<'sec'
      puts "\n\nTest Start  = "<<start.strftime("%H:%M:%S")
      puts "Test Finish = "<<finish.strftime("%H:%M:%S")
      puts "Test Time   = #{hours}hr#{minutes}min#{seconds}sec"
    end
    return test_time
  end






begin
  puts" \n Executing: #{(__FILE__)}\n\n" # show current filename
  #g = Generic.new
  roe = ARGV[1].to_i
  excel = setup(__FILE__)
  wb,ws = excel[0][1,2]
  
  ip          = ws.Range("b3")['Value']
  com_str     = ws.Range("b5")['Value']
  mib         = ws.Range("b4")['Value']
  rows        = ws.Range("b2")['Value'] # Number of row to execute
  iterations  = ws.Range("b6")['Value']
  delay       = ws.Range("b7")['Value']

  colm = "ah"
  col_cnt = iterations
  col = 1

  while col <= col_cnt
  sleep (delay)
  row = 1
  while (row <= rows)
    row += 1 # start at row 2
    _oid  = ws.Range("k#{row}")['Value']
    ndxa   = (ws.Range("l#{row}")['Value']).to_i.to_s
    ndxb  = (ws.Range("m#{row}")['Value']).to_i.to_s
    type  = ws.Range("n#{row}")['Value']
    val   = ws.Range("o#{row}")['Value'].to_i.to_s

    get   = _oid + "." + (ndxa) + "." + (ndxb)

    command = 'snmpget -v2c -c '<< com_str << ' ' << ip << ' ' << mib << '::' << get
    #puts   %x(#{command})
    puts"   #{command}" ;  $stdout.flush
    # GET return
    ret = `#{command}`.to_s
    # Convert return oid to array and extract VALUE from 4th element[3]
    snmp_data = `#{command}`.to_s.split(/ /)[3]
    ws.Range("#{colm}#{row}")['Value'] = snmp_data
  end
  wb.save

  colm = colm.next
  col +=1
  end
  f = Time.now
  tear_down_d(excel[0],s,f,roe)
    
 
end

