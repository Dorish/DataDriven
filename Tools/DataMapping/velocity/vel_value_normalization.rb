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

def datamapping_setup
  puts "----- Follow below steps to setup the test -----\n\n"
  Dir.chdir(File.dirname(__FILE__).sub('velocity','InputFiles'))
  while File.exist?('enp2dd.xml') == false
    puts "Please move gdd xml latest version into InputFiles folder"
    puts "Get gdd xml from here http://126.4.1.113/twiki/bin/view/LmgEmbedded/MAT_FDMs"
    puts "Press Enter after done - "
    gets
  end
  dev=[[],[]]
  num = 0
  devind = 0
  if File.exist?('device.xml')
    f_device = File.open('device.xml','r')
    f_device.each do |line|
      if line =~ /<DeviceName>.+<\/DeviceName>/
        dev[num][0]= />.+</.match(line).to_s.delete('><')
      end
      if line =~ /<FDM>.+<\/FDM>/
        dev[num][1]= />.+</.match(line).to_s.delete('><')
        num = num + 1
      end
    end
    f_device.close
    puts "Select the device to test on - "
    for i in 0...num
      puts (i+1).to_s + ' - ' + dev[i][0]
    end
    puts "If the device is not in the list, press Enter "
    devinput = gets
    while devinput != "\n" && (devinput.chomp.to_i>num||devinput.chomp.to_i<1)
      puts "Select again - "
      devinput = gets
    end
    if devinput == "\n"
      puts "Type in the Device Name - "
      test_device = gets.chomp
      dev[num][0] = test_device
      dev[num][1] = 'unknownfdm'
      devind = num
      num = num + 1
    else
      test_device = dev[devinput.chomp.to_i - 1][0]
      devind = devinput.chomp.to_i - 1
    end
  else
    puts "Type in the Device Name - "
    test_device = gets.chomp
    dev[num][0] = test_device
    dev[num][1] = 'unknownfdm'
    devind = num
    num = num + 1
  end
  puts "Test device is - #{test_device}"
  if File.directory?(test_device)
    test_fdm = dev[devind][1]
    if File.exist?("#{test_device}/#{test_fdm}")
      puts "The last used fdm file for this device is #{test_fdm}"
      puts "Is this still the latest version ? (Y/N)"
      judg = gets.chomp
      while judg!='Y'&&judg!='N'&&judg!='y'&&judg!='n'
        puts "Type in again - "
        judg = gets.chomp
      end
      if judg =='Y'||judg =='y'
        newfdm = test_fdm
      else
        puts "Please move #{test_device} FDM xml latest version into #{test_device} folder"
        puts "Get FDM xml from here http://126.4.1.113/twiki/bin/view/LmgEmbedded/MAT_FDMs"
        puts "Type in the FDM xml file name - "
        newfdm = gets.chomp
        while newfdm == ''||!File.exist?("#{test_device}/#{newfdm}")
          puts "Please check the fdm name and type in again - "
          newfdm = gets.chomp
        end
      end
    else
      puts "Please move #{test_device} FDM xml latest version into #{test_device} folder"
      puts "Get FDM xml from here http://126.4.1.113/twiki/bin/view/LmgEmbedded/MAT_FDMs"
      puts "Type in the FDM xml file name - "
      newfdm = gets.chomp
      while newfdm == ''||!File.exist?("#{test_device}/#{newfdm}")
        puts "Please check the fdm name and type in again - "
        newfdm = gets.chomp
      end
    end
    while !File.exist?("#{test_device}/savedDevice.xml")
      puts "Please move #{test_device} savedDevice.xml file into this folder"
      puts "Press Enter after done - "
      gets
    end
  else
    Dir.mkdir(test_device)
    puts "#{test_device} folder is created under InputFiles folder"
    puts "Please move #{test_device} FDM xml latest version into #{test_device} folder"
    puts "Get FDM xml from here http://126.4.1.113/twiki/bin/view/LmgEmbedded/MAT_FDMs"
    puts "Type in the FDM xml file name - "
    newfdm = gets.chomp
    while newfdm == ''||!File.exist?("#{test_device}/#{newfdm}")
      puts "Please check the fdm name and type in again - "
      newfdm = gets.chomp
    end
    puts "Please move #{test_device} savedDevice.xml file into #{test_device} folder"
    puts "Press Enter after done - "
    gets
    while !File.exist?("#{test_device}/savedDevice.xml")
      puts "Please move #{test_device} savedDevice.xml file into this folder"
      puts "Press Enter after done - "
      gets
    end
  end
  dev[devind][1] = newfdm # update the device.xml
  f_device = File.open('device.xml','w')
  f_device.puts '<?xml version="1.0"?>'
  f_device.puts '<records>'
  for j in 0...num
    f_device.puts " <DeviceName>#{dev[j][0]}</DeviceName>"
    f_device.puts " <FDM>#{dev[j][1]}</FDM>"
  end
  f_device.puts '</records>'
  f_device.close
  return test_device,newfdm
end

dev,fdm = datamapping_setup
puts "----- Test Setup Done -----\n\n"
sleep (1)
Dir.chdir(File.dirname(__FILE__).sub('velocity', "InputFiles/#{dev}")) # change to directory of this file
f_fdm = File.open(fdm)
f_dev = File.open('savedDevice.xml')
Dir.chdir(File.dirname(__FILE__).sub('velocity', 'InputFiles')) # change to directory of this file
f_gdd = File.open('enp2dd.xml')

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
