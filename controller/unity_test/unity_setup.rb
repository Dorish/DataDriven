# To change this template, choose Tools | Templates
# and open the template in the editor.

class Unity_SetUp

  #    - create new IE instance and navigate to the test site
  def open_ie(site)
    puts "\n    **Open IE **\n"
    $ie = Watir::IE.new
    $ie.goto(site)
  end

  #     - attach to existing IE instance and navigate to the test site
  def attach_ie(site)
    puts "\n    **Attach to IE **\n"
    site = 'http://'<<site<<'/'
    $ie = Watir::IE.attach(:url, site)
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

  # parse an excel file and pick up needed information.
  def parse_case(excel_file,result_folder)
    spread_sheet = new_xls(excel_file)
    work_sheet = spread_sheet[2] # worksheet
    
    file_time_stamp = (excel_file.chomp(".xls")<<'_'<<Time.now.strftime("%m-%d_%H-%M-%S")<<(".xls"))
    new_spread_sheet = file_time_stamp.sub(File.dirname(file_time_stamp),"#{result_folder}")
    
    parameters = Hash.new()
    parameters["rows"] = work_sheet.Range("b2")['Value'].to_i
    parameters["test_site"] = work_sheet.Range("b3")['Value']
    parameters["username"] = work_sheet.Range("b4")['Value']
    parameters["password"] = work_sheet.Range("b5")['Value']
    parameters["new_spread_sheet"] = new_spread_sheet
    parameters["spread_sheet"] = spread_sheet
    parameters["work_sheet"] = work_sheet

    sleep 1
    work_sheet.saveas(new_spread_sheet)
    return parameters
  end

  def connect_to_unity(excel_file,result_folder)
    systemos      #Determine whether the OS is Chinese or English
    parameters = Hash.new()
    parameters = parse_case(excel_file,result_folder)

   open_ie(parameters["test_site"])
   return parameters
  end

  #  - create and return new instance of excel
def new_xls(s_s) #wb name and sheet number
  ss = WIN32OLE::new('excel.Application')
  wb = ss.Workbooks.Open(s_s)
  ws = wb.Worksheets(1)
  ss.visible = true # For debug
  xls = [ss,wb,ws]
end

end
