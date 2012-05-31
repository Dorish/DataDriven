=begin rdoc
*Revisions*
  | Initial File | Scott Shanks | 06/17/10 |

*Test_Script_Name*
  mon_gather_data_points

*Test_Case_Number*
  multiple

*Description*
  Gathers all datapoints from the monitor tab and write to a spreadsheet.
This script uses the populate_links_array method to build an array of all text
based links on the navigation pane.

*Variable_Definitions*

=end

$:.unshift File.dirname(__FILE__).chomp('Tools/DataMapping')<<'lib' # add library to path
s = Time.now
require 'generic'

#    - create time stamped controller spreadsheet
#    - open IE or attach to existing IE session
#    - populate the spreadsheet with web support page info
def setup(g,file,rs_name = nil)
  systemos      #Determine whether the OS is Chinese or English
  base_xl = (file).gsub('/','\\').chomp('rb')<<'xls'
  if(ARGV.length != 0)          # called from controller
    excel = xls_timestamp(g,base_xl,nil,ARGV[2]) # called ,connect to existing excel instance. # ARGV[2] is the result sub-folder name.
    g.attach_ie(excel[1][2])  # test site ip
  else
    excel = xls_timestamp(g,base_xl,'ind',rs_name) # independent, start new excel instance
    g.open_ie(excel[1][2])
    support(g,excel[0])
    g.ver_info(excel[0])

  end
  return excel
end

#    - create time stamped spreadsheet using base name
#    - connect to excel and open base workbook
#    - create instance of excel (xl)
#    - returns a nested array containing spreadsheet and script parameters
def xls_timestamp(g,s_s,type=nil,rs_name=nil)
  new_ss = (s_s.chomp(".xls")<<'_'<<g.t_stamp<<(".xls"))
  new_ss1 = new_ss.sub(/Tools\\.+\\/,"result\\#{rs_name}\\")
  if (type == 'ind') # driver was launched independently
    xl = g.new_xls(s_s,1) #open base driver ss with new excel session
  else # driver was launched by controller
    xl = g.conn_open_xls(s_s,1) #connect to existing excel session and create new workbook for L2
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
  g.save_as_xls(xl,new_ss1)
  return excel
end

def support(g,xl)
  puts "  Collect Support page info"
  g.supp.click
  sleep 1
  wb,ws = xl[1,2]
  row = 11
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
  os.value = ["Operating System Language","#{$os}"]
  os.Interior['ColorIndex'] = 43  # change background color
  os.Borders.ColorIndex = 1        # add border
  ws.range("A:B").ColumnWidth = 255 #255 is the maximum column width
  ws.range("A:B").Rows.Autofit
  ws.range("A:B").Columns.Autofit
  wb.Save
end


def systemos
  lang = `systeminfo`
  if lang =~ /en-us*/
    $os          = "English"
    $titl          = "Connect to "
    $ok       ="OK"
    $cancel    = "Cancel"
  elsif lang =~ /zh-cn*/
    $os           = "Chinese"
    $titl           = "连接到 "
    $ok        ="确定"
    $cancel      = "取消"
  end
  puts "This OS is #{$os}"
end

# - Support table
#def supprt; det.table(:index, 2).to_a .compact; end
def supprt
  $ie.frame(:index, 5).table(:index, 2).to_a .compact
end

begin
  puts" \n Executing: #{(__FILE__)}\n\n" # print current filename
  g = Generic.new
  roe = ARGV[1].to_i
  excel = setup(g,__FILE__)
  wb,ws = excel[0][1,2]
  rows = excel[1][1]

  $ie.speed = :zippy
  ws = wb.Worksheets('Data')

  g.count_frames
  g.monitor.click
  sleep(5)
  while g.count_images('folderplus.gif') > 0
    g.click_all('folderplus.gif')
  end

  g.populate_links_array

  #Clean up the links array for the navigation frame

  for i in 0..g.links_array[1].size-1 do
    begin
      #We don't want to click links with parenthesis for this test case
      unless g.links_array[1][i].text =~ /\(\d*\)/ then
        puts "Trying link: #{g.links_array[1][i].text}"
        $ie.frame(:index, 2).link(:id, g.links_array[1][i].id).click
        sleep(2) #Wait for the table to finish populating
        if g.links_array[1][i-1].text =~ /\[\d*\]/ then
          g.table_to_ss(3,ws,g.links_array[1][i].text + ' ' + $&)
        else g.table_to_ss(3,ws,g.links_array[1][i].text)
        end
      end
    rescue => e
      if e.to_s =~ /unknown property or method/ then
        next #Ignore this error and continue the loop - I think it raises an
        #exception because it is trying to access the text of an image...
      else
        puts e.to_s
      end
    end
  end

  f = Time.now  #finish time
rescue Exception => e
  f = Time.now  #finish time
  puts" \n\n **********\n\n #{$@ } \n\n #{e} \n\n ***"
  error_present=$@.to_s

ensure #this section is executed even if script goes in error
  if(error_present == nil)
    # If roe > 0, script is called from controller
    # If roe = 0, script is being ran independently
    g.tear_down_d(excel[0],s,f,roe)
    if roe == 0
      $ie.close
    end
  else
    puts" There were errors in the script"
    status = "script in error"
    wb.save
    wb.close
  end
end
