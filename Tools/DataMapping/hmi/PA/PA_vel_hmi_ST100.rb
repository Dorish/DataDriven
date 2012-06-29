#This script is used to get the HMI and Velocity data mapping information from a HMI-Velocity spreadsheet (for example PA_Echo_ST100_Parameters.xls).
#Input - spreadsheet like PA_Echo_ST100_Parameters.xls 
#Function - Script abstract Velocity ID, Velocity Label, HMI register, HMI Label and HMI page information.
#Output - a txt file contains those above information, with ; as separator.
# situation handled - 
# 1 - Velocity ID existed and at least 2 digits
# 2 - HMI register can existed or '-'
# 3 - One HMI register to multiple velocity IDs
# 4 - One velocity ID to multiple HMI registers.

require 'win32ole'

# - time stamp in 'month-day_hour-minute-second' format
def t_stamp
  Time.now.strftime("%m-%d_%H-%M-%S")
end

#  - createand return new instance of excel
def new_xls(s_s,num) #wb name and sheet number
  ss = WIN32OLE::new('excel.Application')
  wb = ss.Workbooks.Open(s_s)
  ws = wb.Worksheets(num)
  ss.visible = true # For debug
  xls = [ss,wb,ws]
end

def timeStamp(vari)
  ext = /\.\w*$/.match(vari).to_s # match extension from end of the string
  if ext
    vari.chomp(ext)+'_'+t_stamp+ext
  else
    vari+'_'+t_stamp
  end
end

#return whether the current sheet is finished or not
#decide by checking each cell of the row nil or not, it's not end unless each cell is nil
def to_end(ws, row)
  j = 0;
  while j <26
    number =  (65 + j).chr #65 is the ASCII "A"
    if ws.range(("#{number}#{row}")).value != nil
      return 0
    end
    j = j + 1
  end
  return 1
end

sor_ss = File.expand_path(File.dirname(__FILE__)) + '/' + 'example_PA_Echo_ST100_Parameters.xls'
i = 1
ss,wb,ws = new_xls(sor_ss,i)

num = 3 # sheets number - need to change based on the amount of sheets it contains
newfile = timeStamp(sor_ss.chomp(".xls")<<(".txt"))
o_file = File.new(newfile, "w")

while i<=num
  ws = wb.Worksheets(i)
  puts "on sheet #{i}........................."
  j = 0;
  while j<26
    number =  (65 + j).chr #65 is the ASCII "A"
    if ws.range("#{number}3").value =~/V4 Data Label/
      k = 5
      while to_end(ws,k)==0
        dup = 0
        gdd_temp = ws.range("#{number}#{k}").value
        hmiregister_temp = ws.range("b#{k}").value
        hmilabel_temp = ws.range("c#{k}").value
        hmipage_temp = ws.name
        if gdd_temp == nil && hmiregister_temp == nil
          dup = 1
        end
        if gdd_temp !=nil # one hmi mutiple gdd. And for page 1 of 5/hmiregister not nil but hmilabel nil
          gdd = gdd_temp.delete("\n")
          if hmiregister_temp !=nil
            hmiregister = hmiregister_temp.delete("\n")
          end
          if hmilabel_temp != nil # For hmiregister_temp read no nil but hmilabel_temp read nil
            hmilabel = hmilabel_temp.delete("\n")
          end
        else # one gdd mutiple hmi
          if hmiregister_temp !=nil
            hmiregister = hmiregister_temp.delete("\n")
            if hmilabel_temp != nil # For hmiregister_temp read no nil but hmilabel_temp read nil
              hmilabel = hmilabel_temp.delete("\n")
            end
          end
        end
        #separate the velocity id and velocity label to print out
        if dup == 0
          if gdd =~/\[\d{2,}\]/
            gddid = (/\[\d{2,}\]/).match(gdd).to_s.delete("[").delete("]")
            if gdd =~ /\s*\[\d{2,}\]$/
              gddlabel = gdd.sub((/\s*\[\d{2,}\]/), "")# Ext Air Sensor A Over Temp Threshold [5337]
            else
              gddlabel = gdd.sub((/\s*\[\d{2,}\]/), "-") # Condenser Fan Power [5538]PRIVATE
            end
            print gddid,";",gddlabel,";", hmiregister, ";", hmilabel, ";", hmipage_temp, "\n"
            o_file.print gddid,";",gddlabel,";", hmiregister, ";", hmilabel, ";", hmipage_temp, "\n"
#          elsif gdd=='-' # GDD ID is '-' but hmi register existed.
#            if hmiregister != '-'
#              print gdd,";",'-',";", hmiregister, ";", hmilabel, ";", hmipage_temp, "\n"
#              o_file.print gdd,";",'-',";", hmiregister, ";", hmilabel, ";", hmipage_temp,"\n"
#            end
          end
        end
        k = k + 1
      end
    end
    j = j + 1
  end
  i = i + 1
end






