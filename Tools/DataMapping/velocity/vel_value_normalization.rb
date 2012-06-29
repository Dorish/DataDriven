=begin
Normal the rough velocity data, output value type, gddid, velocity value, resolution, scale and final value.

input -
1)'savedDevice.xml' - velocity walk from device browser. This is the rough velocity data that the script
will convert with.
2)'iCOM_CR 468.xml' - fdm(change based on different device) of the device. if local data points, script will lookup in fdm.
3)'enp2dd.xml' - gdd of the device. if global data points, script will lookup in gdd file.

output -
Output each savedDevice.xml data point - report name, velocity data, scaling, resolution
offset and normalization value.
=end

require 'win32ole'

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

Dir.chdir(File.dirname(__FILE__).sub('Velocity','InputFiles')) # change to directory of this file
savedDevice = 'crv-savedDevice.xml'
fdm = 'iCOM_CR 468.xml'
gdd = 'enp2dd.xml'

f_dev = File.open(savedDevice)
f_fdm = File.open(fdm)
f_gdd = File.open(gdd)

f_dev.each do|line|
  gddid = /id="\d{1,}"/.match(line).to_s.delete('id="') # get gdd id
  value = />.*</.match(line).to_s.delete('><')# positive number, negative number and string
  temp_gddlabel = /name="(\w|\s|-|\(|\)|\/)+"/.match(line).to_s.delete('"')
  gddlabel = temp_gddlabel[5,temp_gddlabel.length - 1] # get gdd label.

  # Start to calculate the values. Five types of value in total
  # Digital values,Text values,Time values,Event values and Enum values.
  
  # Digital values
  if line =~ /type="DescriptorUint16" | type="DescriptorInt16"/
    scale,resolution,normalization_value = digital_value(line,value)
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
end
