require 'win32ole'

def setup(file,folder)
  driver_path = (file).gsub(/Tools.*/,'driver/'+folder+'/') #get the driver path
  contr_path = (file).gsub(/Tools.*/,'controller/') #get the controller path
  adriver_name = adriver_name= Array.new
  adriver_name = Dir.entries(driver_path ).delete_if{ |e| e=~ /^\..*/|| e=~/^.*\.xls/|| e=~/backup/i} #read the driver folder and write the rb file name to array

  temp = Dir.entries(contr_path ).delete_if{ |e| e=~ /^\..*/|| e=~/^.*\.rb/} #read the controller  folder and write the ss file name to array
  ind = file_search(contr_path){|i| i=~/#{folder}/} #get the target spreadsheet index of the array of the controller spreadsheets
  contr_name = contr_path + temp[ind] # controller path and name

  param = Array.new 
  param[0..2] = driver_path, adriver_name,contr_name
  return param
end

#  - createand return new instance of excel
def new_xls(s_s) #wb name and sheet number
  ss = WIN32OLE::new('excel.Application')
  wb = ss.Workbooks.Open(s_s)
  ws = wb.Worksheets(1)
  ss.visible = true # For debug
  xls = [ss,wb,ws]
end

def change_con(file,folder)
  param = setup(file,folder) #param is array include driver_path,adriver_name,contr_name

  ss =new_xls(param[2])
  wb = ss[1]
  ws = ss[2]

  ws.Range("j2:j101").Columns.Delete
  row =2

  param[1].each{|x|   #add every driver name to ss
    x="/#{folder}/"+x
    ws.Range("j#{row}")['Value'] =x
    puts x
    row =row+1
  }
  puts " \n** #{param[1].length} script names were added to #{param[2]} **\n\n"
  wb.save
  ss[0].quit
end

#Search a target file in the given path, return the index of the array of the filenames if exist, otherwise return -1
def file_search(path)
  fl_list = Dir.entries(path).delete_if{ |e| e=~ /^\..*/|| e=~/^.*\.rb/}
  fl_list.each { |i|
    if yield(i)# Use yield to do different things when is called
      return fl_list.index(i)
    end
  }
  return -1
end


#script running
p "Start Running"
driver_path = (__FILE__).gsub(/Tools.*/,'driver/') #get the driver path
while 1
  print "\nPlease type the target folder name inside driver folder followed by <Enter>: "
  foldername = gets.chomp
  if file_search(driver_path){|i| i==foldername} != -1 
    break
  end
end

change_con(__FILE__,foldername)
p "End Running"
