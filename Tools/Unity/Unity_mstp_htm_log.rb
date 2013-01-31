=begin
This script will collect data from /devMem.htm data pages on the IS-WEBx cards.
The device ip is currently hard coded eg. collect_log("126.4.10.243")
The number of times to poll for data is controlled by collect_log / cnt
The polling interval is controlled by collect_log / dly
=end





#require 'lib\Generic'
require 'watir/ie'
require 'win32ole'


def creat_ss(topic,title)
  #creat  ss
  ss = WIN32OLE::new('excel.Application')
  wb = ss.Workbooks.add
  ss.visible = true
  s_s = []

  # add worksheets if it is necessary
  wb.Worksheets.Add
  ws = [ ]

  #change the worksheets name
  t = 1
  topic.each{|x|
    wb.Worksheets(t).name = x[1].sub(/\//,"_")
    ws << wb.Worksheets(t)
    t += 1
  }

  #write the tital in the ss
  n = 0
  ws.each{|x|
    x.Range("A1")['Value'] = topic[n][0]  #write page  ip
    x.Range("A2")['Value'] = title[n].shift  #write title in the ss
    x.Range("A3")['Value'] = 'Date / Time'   #write time

    number =  (66  + title[n].length - 1).chr #66 is the ASCII "B"
    x.Range("B3:#{number}3")['Value'] = title[n]
    n += 1
    x.range("A:Z").Columns.Autofit
  }
  s_s = [wb,ws]
  return s_s
end

def write_ss(info,s_s)
  #write the data in the ss
  ws = s_s[1]
  n = 0
  ws.each{|x|
    x.Range("B2")['Value'] = info[n].shift  #write the column name
    number =  (65 + info[n].length).chr #65 is the ASCII "A"
    x.Range("A#{$row}:#{number}#{$row}")['Value'] = info[n].unshift $s.strftime("%m.%d %H:%M%::%S")
    n += 1
    x.range("A:Z").Columns.Autofit
  }
  $row += 1
  return s_s
end

def collect_log(ip,cnt,dly)

  log_name = (File.dirname(__FILE__).sub('Tools/Unity','result')<<'/'<<"mstp_log_"<<Time.now.to_a.reverse[5..9].to_s<<(".xls")).gsub('/','\\')

  p log_name
  $ip = ip
  $ie = Watir::IE.new

  site = "http://" + $ip +"/" +"mstp.htm?devId=0"
  p site
  #      puts site
  $ie.goto(site)

  #Puts  the link info save in array
  puts "Puts the link info"
  topic_add =[[site,"mstp" ]]

  

  #collection every tital info save in array
  puts "collection every tital info"
  title_info = []
  temp3 = []
  $ie.table(:index,1).to_a.map{|x| temp3 << x[0]}
  title_info << temp3
 

  # creat a ss template
  puts "creat ss"
  s_s = creat_ss(topic_add,title_info)
  s_s[0].saveas(log_name)
 
  #collection info from page in array
  puts "collection info from page"
  $row = 4
    page_info = []
    cnt.times{|t|
    puts "loop time == #{t}"
    $s = Time.new
     temp4 = []
      $ie.goto(site)
      $ie.table(:index,1).to_a.map{|x| temp4 << x[1] }
       page_info << temp4
     
    #write info to ss
    write_ss(page_info,s_s)
    s_s[0].save

    #wait and collection
    sleep(dly)
 }
    #save the wb info
    s_s[0].close
    $ie.close
  
    puts "Finish"
  end


  # Initialize variables
  host = 'C:/WINDOWS/system32/drivers/etc/hosts'
  valid_ip = /^(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})/
  testsite = /^(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\s+(test_site)/

  ip1 = nil
  ip2 = nil

  File.open(host).each do|line|
    if line =~ testsite
      ip1 = $1  # $1 is the group in the valid_ip regex
      print "Existing test_site IP address is: ",ip1
    end
  end

  ip = ip1    # ip = test_site

  puts "\n\nTo keep the existing IP address, press <Enter>"
  print "-OR- Type new IP address followed by <Enter>: "

  ip2 = gets.chomp
  if ip2 != ''
    while ip2 !~ valid_ip
      print "\nPlease type a valid IP address followed by <Enter>: "
      ip2 = gets.chomp
    end
    ip = ip2    # ip = user entry
  end

  # how many times to log
  puts "\nType number of time to log data followed by <Enter>"
  cnt = gets.chomp.to_i

  # log interval time in seconds
  puts "\nType the log interval(in seconds) followed by <Enter>"
  dly = gets.chomp.to_i

  #puts cnt.class
  #puts dly.class

  collect_log(ip,cnt,dly)
