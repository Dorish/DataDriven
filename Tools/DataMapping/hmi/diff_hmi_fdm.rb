=begin
This script is written for hmi data mapping test case generation to check if the velocity data points in
the hmi.xml exist in the fdm. The expectation should be yes.

input -  hmi file and fdm file, put them in the DataMapping/InputFiles folder. Also, need to put their
name into the scripts parameters 'hmi' and 'fdm'.

output - generate a time stamped spreadsheet contains the velocity ID of hmi and fdm in two different
colums. And, the script will output the data point difference of these two files.
=end

require 'win32ole'

#  - create and return new instance of excel
def creat_xls(num)
  ss = WIN32OLE::new('excel.Application')
  wb = ss.Workbooks.add
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

Dir.chdir(File.dirname(__FILE__).sub('hmi','InputFiles'))
hmi = 'hmi.xml' # This need to change when file name changed
fdm = 'iCOM_PA 204.xml' # This need to change when file name changed
newss = timeStamp((__FILE__).sub('.rb','.xls'))

f_hmi = File.open(hmi)
f_fdm = File.open(fdm)


xl = creat_xls(1)
ws = xl[2]
hmi_array = []
fdm_array = []
hmi_row = 1
fdm_row = 1
ws.Range("A#{hmi_row}")['Value'] = 'HMI'
ws.Range("B#{fdm_row}")['Value'] = 'FDM'

# Read the velocity IDs from hmi.xml
f_hmi.each do |line|
  if line =~ /<DataIdentifier>\d{2,}<\/DataIdentifier>/
    hmi_array[hmi_row-1] = /\d{2,}/.match(line).to_s
    hmi_row = hmi_row + 1
    ws.Range("A#{hmi_row}")['Value'] = /\d{2,}/.match(line).to_s
  end
end

# Read the velocity IDs from fdm
f_fdm.each do |line|
  if line =~ /<dataPoint type="\w+" id="\d+">/
    fdm_array[fdm_row-1] = /id="\d+"/.match(line).to_s.delete('id="')
    fdm_row = fdm_row + 1
    ws.Range("B#{fdm_row}")['Value'] = /id="\d+"/.match(line).to_s.delete('id="')
  end
end

save_as_xls(xl,newss)
puts "#{hmi_row-1} gdd in #{hmi}"
puts "#{fdm_row-1} gdd in #{fdm}"

puts "#{hmi}-#{fdm} are,"
puts hmi_array - fdm_array # in hmi while not in fdm.
puts "#{fdm}-#{hmi} are,"
puts fdm_array - hmi_array # in fdm while not in hmi.
