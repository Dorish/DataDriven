=begin
Normalization the rough velocity data and write it into the data spreadsheet of other protocols.
This script running requires the another protocol data spreadsheet(webl_mon_data_map.xls) ready.

input -
1)'savedDevice.xml' - velocity walk from device browser. This is the rough velocity data that the script
will convert with.
2)'iCOM_CR 468.xml' - fdm(change based on different device) of the device. if local data points, script will lookup in fdm.
3)'enp2dd.xml' - gdd of the device. if global data points, script will lookup in gdd file.
4)'webl_mon_data_map.xls' - spreadsheet(change based on different protocol) that contains values of another protocol.

output -
Time stamped spreadsheet of the input spreadsheet. Add report name, velocity data, scaling, resolution
offset and normalization value into this spreadsheet.
=end

require 'win32ole'
require 'time'

#  - createand return new instance of excel
def new_xls(s_s,num) #wb name and sheet number
  ss = WIN32OLE::new('excel.Application')
  wb = ss.Workbooks.Open(s_s)
  ws = wb.Worksheets(num)
  ss.visible = true # For debug
  xls = [ss,wb,ws]
end
#  - save an existing workbook as another file name
def save_as_xls(s_s,save_as)
  sleep 1
  s_s[2].saveas(save_as)
end
# - time stamp in 'month-day_hour-minute-second' format
def t_stamp
  Time.now.strftime("%m-%d_%H-%M-%S")
end

def timeStamp(vari)
  ext = /\.\w*$/.match(vari).to_s # match extension from end of the string
  if ext
    vari.chomp(ext)+'_'+t_stamp+ext
  else
    vari+'_'+t_stamp
  end
end

#Search the label from a specific row of a spreadsheet.
#Use regular expression to match. Because the label of other protocols(web,modbus,etc)may not be equal.
def excel_search(ws, label, row)
  i = row
  cellstr = "This is a initialization string"
  while cellstr != label
    i = i + 1 # search from the next row of 'row'
    cellstr = ws.Range("B#{i}")['Value'] #TODO Colum index may change based on different protocol spreadsheet.
    if cellstr == nil
      break
    else
      cellstr = cellstr.lstrip
      if cellstr.include?"("
        if cellstr == label
          cellstr
        else
          cellstr = cellstr.sub(/ \(.+\)$/, "")
        end
      end
    end
  end
  return i
end
def excel_max_rows(ws)
  i = 1
  while ws.Range("A#{i}")['Value'] != nil
    i = i + 1
  end
  i = i - 1
end
# process the round and decimal point.
def value_round(resolution,value)
  y=1
  resolution.times {|n| y=y*10}
  if resolution == 0
    value = (value.round)/y
  else
    value = (value.round)/(y*1.0)
  end
end

def digital_value(line,value)
  name = /name="(\w|\s)*"/.match(line).to_s
  resolution = /resolution="\d{1,}"/.match(line).to_s.delete('resolution="')
  scale = /scale="0{0,1}\.{0,1}\d{1,}"/.match(line).to_s.delete('scale="') # 1/0.1
  int_value = value.to_i
  if int_value >= 32768 || int_value <= -32768
    d_value = 'Unavailable'
  else
    temp_resolution = resolution.to_i
    temp_scale = scale.to_f
    if temp_scale >=1.0
      temp_scale = temp_scale.to_i
    end
    d_value = int_value * temp_scale
    d_value = value_round(temp_resolution,d_value)
  end
  return scale,resolution,d_value
end
def time_value(value)
  t_value = value.to_i
  t_value = "'" + Time.at(t_value).strftime("%m/%d/%Y %H:%M")
  return t_value
end
def event_value(value)
  eve_num = value.to_i
  case eve_num
  when 4  #when 4 ||12 ||20 ||28 ||16384 does not work?
    e_value = 'Normal'
  when 12
    e_value = 'Normal'
  when 20
    e_value = 'Normal'
  when 28
    e_value = 'Normal'
  when 16384
    e_value = 'Normal'
  when 3 #when 3 ||11 ||19 ||27 does not work
    e_value = 'Active'
  when 11
    e_value = 'Active'
  when 19
    e_value = 'Active'
  when 27
    e_value = 'Active'
  end
  return e_value
end

def enum_value(f_fdm,f_gdd,value,gddid)
  f_fdm.rewind # every data point need to search from the beginning.
  f_gdd.rewind
  temp = String.new("")
  flag = 0
  #lookup the fdm to find the specific data point definition
  f_fdm.each do |line|
    if line =~ /<dataPoint type="DataEnum" id="#{gddid}">/
      flag = 1
    end
    if flag == 1
      temp = temp + line
    end
    if flag == 1 && line =~ /<\/dataPoint>/
      break
    end
  end
  # get the enumstate id
  enumstr = /<EnumStateDefnID>\d{1,}<\/EnumStateDefnID>/.match(temp).to_s.delete('<EnumStateDefnID>/')
  if enumstr.to_i<=99 # distinguish local enum definition or gdd enum definition
    f_fdm.rewind
    f = f_fdm
  else
    f = f_gdd
  end
  #get the enumstate string ID
  flag = 0
  stringID = ""# used to contain the stiring ID of the enumvalue
  normalization_value = "velocity value #{value} is not defined in GDD file" # used to contain the final value. This cannot be removed, because it no enumvalue found equal to value, then stringID will be "", then normalization_value will be the init value now.
  f.each do |line|
    if line =~/<EnumStateDefn .*Id="#{enumstr}"/
      flag = 1
    end
    if flag == 1
      enumvalue = /Value="\d{1,}"/.match(line).to_s.delete('Value="') #  nil.to_i = 0
      if enumvalue == value
        stringID = />\d{1,}</.match(line).to_s.delete('><') #get the stiring ID of the enumvalue
      end
    end
    if flag == 1 && line =~/<\/EnumStateDefn>/
      break
    end
  end
  #get the final value
  f.rewind
  f.each do |line|
    if line =~/<String Id="#{stringID}">/
      normalization_value = />.+</.match(line).to_s.delete('><')#get the final string value
    end
  end
  return normalization_value
end

mapfile = (File.dirname(__FILE__)+'\\')+'webl_mon_data_07-06_14-40-46.xls'
newss = timeStamp(mapfile)
xl = new_xls(mapfile,1) #open base driver ss with new excel session
wb,ws=xl[1,2]
fdm = ws.Range("B#{12}")['Value']
gdd = ws.Range("B#{13}")['Value']
savedDevice = ws.Range("B#{14}")['Value']
Dir.chdir(File.dirname(__FILE__).sub('velocity', 'InputFiles')) # change to directory of this file
f_fdm = File.open(fdm)
f_gdd = File.open(gdd)
f_dev = File.open(savedDevice)

ws = wb.Worksheets(3)
maxrows = excel_max_rows(ws)

f_dev.each do|line|
  gddid = /id="\d{1,}"/.match(line).to_s.delete('id="')# get gdd id
  value = />.*</.match(line).to_s.delete('><')# positive number, negative number and string
  temp_gddlabel = /name="(\w|\s|-|\(|\)|\/|')+"/.match(line).to_s.delete('"')
  gddlabel = temp_gddlabel[5,temp_gddlabel.length - 1] # get gdd label.
  if line=~/<datapoint id="\d{1,}"/
    row = 1
    row = excel_search(ws, gddlabel, row) # return the row of the same gdd label in spreadsheet
    if row > maxrows
      puts gddid +"-"+ gddlabel.to_s + " is not in spreadsheet ..............................."
    else

      # search until find the row and cell is empty, because some velocity points are multi module, label are similar
      while ws.Range("K#{row}")['Value'] != nil # don't remove, for situation that name are similar for example sensor
        row = excel_search(ws, gddlabel, row)
      end
      # Digital values
      if line =~ /type="DescriptorUint16" | type="DescriptorInt16"/
        scale,resolution,normalization_value = digital_value(line,value)
        ws.Range("M#{row}")['Value']= scale
        ws.Range("N#{row}")['Value']= resolution
        print "ditital value ------- ", "gddid=#{gddid}"," - ","veloticy value=#{value}"," - ", "resolution=#{resolution}"," - ","scale=#{scale}"," - ", normalization_value, "\n"
      end

      #Text values
      if line =~ /type="DescriptorText"/
        normalization_value = value
        print "text value ------- ", "gddid=#{gddid}"," - ","velocity value=#{value}" ," - ", normalization_value, "\n"
      end

      #Time values
      if line =~ /type="DescriptorTime32"/
        normalization_value = time_value(value)
        print "time value ------- ", "gddid=#{gddid}"," - ","velocity value=#{value}" ," - ", normalization_value, "\n"
      end

      #Event values
      if line =~ /type="DescriptorEvent16"/
        normalization_value = event_value(value)
        print "event value ------- ", "gddid=#{gddid}"," - ","velocity value=#{value}" ," - ", normalization_value, "\n"
      end

      #Enum values
      if line =~ /type="DescriptorEnum"/
        normalization_value = enum_value(f_fdm,f_gdd,value,gddid)
        print "enum value ------- ","gddid=#{gddid}"," - ", "velocity value=#{value}", " - ",normalization_value,"\n"
      end
      ws.Range("K#{row}")['Value']= gddlabel
      ws.Range("L#{row}")['Value']= value
      ws.Range("P#{row}")['Value']= normalization_value
    end
  else
    #not velocity data points
  end
end
save_as_xls(xl,newss)