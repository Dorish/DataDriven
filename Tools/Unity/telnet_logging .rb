# To change this template, choose Tools | Templates
# and open the template in the editor.
require 'win32ole'
require 'net/telnet'

$enter = "\x0D" # Ascii Hex code for Carriage Return is 0D
$ctrl_c = "\x03" # Ascii Hex code for Ctrl+C is 03
$unity_exit_7 = "\x07" # Ascii Hex code for 7 is 07
#get input parameters from input file
def get_input_parameters(csv_file)
  input_parameter_list = Array.new()
  open(csv_file).map do |line|
    # ignore column header row
    if line !~ /IP Address/     
      input_parameter = Hash.new()

      # convert each line to an array
      line =  line.split(/,/).to_a.push  
      input_parameter["ip_address"] = line[0]
      input_parameter["user_name"] = line[1]
      input_parameter["password"] = line[2]
      input_parameter["interval_time"] = line[3]
      input_parameter["loop_times"] = line[4]
      input_parameter["desired_location"] = line[5]
      input_parameter["info_type"] = line[6]
      input_parameter_list.push(input_parameter)
    end
  end
  return input_parameter_list
end
#create spreatsheet
def create_sheet(spread_sheet,ip_address,info_type)
  output_file_name = (File.dirname(__FILE__) + '/' + ip_address + '.csv').gsub('/','\\')

  # judge the workbook with a given name exist or not.
  # if exist and opened, active it
  # if exist but unopened. open it
  #if not exist, create a new one
  if !File.exists?(output_file_name)
    work_book = spread_sheet.Workbooks.add
    work_book.saveas(output_file_name)
  else
    work_book = spread_sheet.ActiveWorkbook
    if work_book == nil
      work_book = spread_sheet.Workbooks.open(output_file_name)
    else
       spread_sheet.Workbooks.each{|book|
        if work_book.name !=  ip_address + '.csv'
          work_book = spread_sheet.Workbooks.open(output_file_name)
        end
      }
    end
  end

  #get the exist worksheet
  work_sheet = nil
  sheet_exist = false
  work_book.worksheets.each { |item|
    if item.name == info_type
      work_sheet = item
      sheet_exist = true
      break
    end
  }

  #if sheet with a given name doesn't exist, create it
  if !sheet_exist
    work_book.worksheets.add
    work_book.worksheets(1).name = info_type
    work_sheet = work_book.worksheets(1)  #first sheet is always the new added one
  end

  #return the necessary information
  xls = [work_book,work_sheet]
  return xls
end
# Is card available?
def is_card_available(ip)
  reply_from = Regexp.new('Reply from')
  results = `ping #{ip}`
  if (reply_from.match(results)) then
    return true
  end
  return false
end
# Recreate the following method to adapting the unity card and webAdapt card.
def telnet_connect(ip,login,pswd)
  prompt_type=/[ftpusr:~>]*\z/n
  puts"ip,=#{ip},login=#{login},password=#{pswd}"
  telnet = Net::Telnet.new('Host' => ip,'Prompt' =>prompt_type ,"Output_log"   => "dump_log.txt"  )

  #The prompt should be the real prompt while you entered your system
  telnet.cmd(''){|c| print c}
  telnet.waitfor(/[Ll]ogin[: ]*\z/n) {|c| print c}
  telnet.cmd(login) {|c| print c}
  telnet.waitfor(/Password[: ]*\z/n) {|c| print c}
  telnet.cmd(pswd) {|c| print c}

  # the following sentence can wrok for unity and webAdapt.
  telnet.waitfor(/[>]|Enter selection\:/n) {|c| print c}

  sleep 5
  return telnet
end
# read desired information from telnet
def get_desired_info(telnet,desired_location)
  navigation = desired_location.split('.')
  log_info = " "

  for i in 0..navigation.length-1 do
    telnet.write(navigation[i]) {|c| print c}
    telnet.write($enter){|c| print c}

    #wait for the end flag, record the necessary information into string
    #^-*\s*\/[\w\/]+\/dlog[0-9]\.log+[\s*\w]+\s*-*$ to match "---- \/var\/log\/dlog00001.log starting ----"
    log_info = telnet.waitfor(/\?>|Enter selection\:|-*\s*\/[\w\/]+\/dlog[0-9]{0,}\.log+[\s*\w]+\s*-*|Press <Enter> to return to menu/) {|c| print c}

    # output the needed information from last index from.
    if i == navigation.length - 1
      # delete the first character( it is the navigate number).
      log_info = log_info[1,log_info.length]
    end

  end
  return log_info
end
# close the connection for unity
def tn_unity_close(telnet,desired_location,info_type)
  navigation = desired_location.split('.')
  
  #Capture_Dlog need press ctrl+C to return main menu
  case info_type
  when "Capture_Dlog"
    telnet.write($ctrl_c) {|c| print c} # Press Ctrl+C to stop DLOG file watcher
    telnet.waitfor(/Press <Enter> to return to menu/) {|c| print c}
  else
    telnet.write($enter) {|c| print c}
  end

  # Press 'enter' several times come back to the first menu
  navigation.length.times do
    telnet.write($enter) {|c| print c}
    telnet.waitfor(/Enter selection\:/) {|c| print c}
  end

  # Press the menu of '7' to exit the connect.
  telnet.write($unity_exit_7) {|c| print c}
  telnet.write($enter) {|c| print c}
end
# write the Filesystem information to output file
def write_filesystem(logging_file,desired_info)
  str_array = desired_info.split("\n")
  row_count = get_csv_rows(logging_file)
  str_array_size = str_array.size - 1

  write_index = 1
  for i in 0..str_array_size
    # ignore the last sentence and title
    if str_array[i] =~ /Press <Enter> to return to menu|Disk Usage/
      next
    end
    #ignore the empty line
    if str_array[i] == ""
      next
    end

    line_info =  str_array[i].split(" ")
    line_info_size = line_info.size - 1

    #write the information into cells
    for j in 0..line_info_size
      column =  (65 + j).chr #65 is the ASCII "A"
      row_index = row_count + write_index
      logging_file.Range("#{column}#{row_index}").value = line_info[j]
    end
    write_index = write_index + 1
  end
end
#write the Memory information to output file
def  write_memory(logging_file,desired_info)
  first_line = true
  row_count = get_csv_rows(logging_file)
  str_array = desired_info.split("\n")

  write_index = 1
  str_array_size = str_array.size - 1
  for i in 0..str_array_size
    # ignore the last sentence and title
    if str_array[i] =~ /Press <Enter> to return to menu|Memory Usage/
      next
    end
    #ignore the empty line
    if str_array[i] == ""
      next
    end

    line_info = Array.new    
    #first line miss a empty value to fill the information into cells, so there is a flag to hlep deal with it. 
    if first_line
      line_info.push(" ")
      str_array[i].split(" ").each{|element| line_info.push(element) }
      first_line = false
    else
      line_info =  str_array[i].split(" ")
    end
    
    #write the information into cells
    line_info_size = line_info.size - 1
    for j in 0..line_info_size
      column =  (65 + j).chr #65 is the ASCII "A"
      row_index = row_count + write_index
      logging_file.Range("#{column}#{row_index}").value = line_info[j]
    end
    write_index = write_index + 1
  end
end
#write the Realtime System Summary information to output file
def  write_output(logging_file,desired_info)
  first_line = true
  row_count = get_csv_rows(logging_file)
  str_array = desired_info.split("\n")
  str_array_size = str_array.size - 1

  write_index = 1
  for i in 0..str_array_size
    # ignore the last sentence
    if str_array[i] =~ /Press <Enter> to return to menu/
      next
    end

    # ignore the last sentence and title
    if str_array[i] =~ /Press <Enter> to return to menu|Realtime System Summary/
      next
    end
    #ignore the empty line
    if str_array[i] == ""
      next
    end
    
    line_info = Array.new    
    #first line is a whole sentence, need to fill in one cell 
    if first_line
      logging_file.Range("A#{i}").value = str_array[i]
      first_line = false
    else     
      # these four sentence need to process separately
      if str_array[i] =~ /^Tasks:*|Cpu\(s\):*|Mem:*|Swap:*/
        para1 = str_array[i].split(":")
        line_info.push para1[0]
        para = para1[1].split(",")
        para.each{|item|
          line_info.push item
        }
      else
        line_info =  str_array[i].split(" ")
      end
      
      #write the information into cells
      line_info_size = line_info.size - 1
      for j in 0..line_info_size
        column =  (65 + j).chr #65 is the ASCII "A"
        row_index = row_count + write_index
        logging_file.Range("#{column}#{row_index}").value = line_info[j]
      end
      write_index = write_index + 1
    end
   
  end

end
#write the Current DLOG File Output information to output file
def write_capturelog(logging_file,desired_info)
  row_count = get_csv_rows(logging_file)
  keyword = ["<warning>","<notice>","<err>"]
  str_array = desired_info.split("\n")
 
  # write the header for this sheet
  for col in 0..2
    column =  (65 + col).chr #65 is the ASCII "A"
    row_index = row_count + 1
    logging_file.Range("#{column}#{row_index}").value = keyword[col]
  end

  write_index = 1
  row_count = row_count + 1
  str_array_size = str_array.size - 1
  for line in 0..str_array_size
    key_index = 0

    # start checking keyword for a match
    keyword.each do |key|           
      regexp = /\.*#{key}/
      if str_array[line] =~ regexp
        line_info = Array.new
       
        # get the lines in the log that match the one of the parse keys
        for i in 0..keyword.length - 1
          if i == key_index
            line_info.push str_array[line].to_s.chomp  #add parsed value to the suitable position in a csv line.
          else
            line_info.push " "  # put empty information for unmatch keyword in a csv line.
          end
        end

        #write the information into cells
        line_info_size = line_info.size - 1
        for j in 0..line_info_size
          column =  (65 + j).chr #65 is the ASCII "A"
          row_index = row_count + write_index
          logging_file.Range("#{column}#{row_index}").value = line_info[j]
        end
        write_index = write_index + 1
      end

      key_index = key_index + 1
    end
  end
end
#get exist rows in csv file
def get_csv_rows(logging_file)
  row = 1
  while  logging_file.Range("A#{row}").value != nil || logging_file.Range("B#{row}").value != nil || logging_file.Range("C#{row}").value != nil
    row = row + 1
  end
  return row - 1
end
#add timestamp to output file
def add_timestamp(logging_file)
  row = get_csv_rows(logging_file)
  new_row_num = row + 1
  logging_file.Range("A#{new_row_num}").value = "------------------------#{Time.now}------------------------"
end

begin
  csv_file = File.dirname(__FILE__) + '\\' + 'telnet_logging.csv' # use this path when we run the script in ruby environment
  spread_sheet = WIN32OLE::new('excel.Application')
  spread_sheet.visible = true
  input_parameter_list = Array.new()
  input_parameter_list = get_input_parameters(csv_file)

  # For serial processing, they have the same loop times and interval time for all IP addresses.
  loop_times = input_parameter_list[0]["loop_times"]
  interval_time = input_parameter_list[0]["interval_time"]

  while(loop_times.to_i > 0)
    input_parameter_list.each{|input_parameters|
      ip_address = input_parameters["ip_address"]
      desired_location = input_parameters["desired_location"]
      info_type = input_parameters["info_type"]

      # create its own spreadsheet
      excel = create_sheet(spread_sheet,ip_address,info_type)
      work_book = excel[0]
      logging_file = excel[1]

      # record the card's state, available or not?
      is_available = is_card_available(ip_address)

      # add the time stamp at the beginning of log file.
      add_timestamp(logging_file)
      if is_available
        tries = 0
        begin
          # connect to card from telnet
          telnet = telnet_connect(ip_address, input_parameters["user_name"], input_parameters["password"])

          # navigate to desired location
          desired_info = get_desired_info(telnet,desired_location)

          # choose the appropriate method to write the record information into  output file
          info_type = info_type.gsub("\n", "")
          case info_type.to_s
          when "Filesystem"
            write_filesystem(logging_file,desired_info)
          when "Memory"
            write_memory(logging_file,desired_info)
          when "Output_Form_Top"
            write_output(logging_file,desired_info)
          when "Capture_Dlog"
            write_capturelog(logging_file,desired_info)
          else
            puts "Not Define yet"
          end

          # close telnet session from card side
          tn_unity_close(telnet,desired_location,info_type)

          # terminate telnet
          telnet.close

        rescue Exception=>e
          sleep 1
          tries += 1
          if telnet == nil # build telnet connection failed.
            puts "telnet to #{ip_address} failed"
            logging_file.puts("telnet to #{ip_address} failed")
          else
            # tn_close(telnet,desired_location) # if this sentence error, no handle script for it.
            telnet.close
          end
          if is_card_available(ip_address)==true # card is still available - retry
            puts "retry #{tries} times"
            retry if tries <= 9
            puts "retry limit reached!"
          else # card is not available - log that 'card did not respond ' and continue running
            logging_file.puts("#{ip_address} did not respond to the ping request")
          end
        end
      else
        logging_file.puts("#{ip_address} did not respond to the ping request")
      end

      if loop_times.to_i == 1
        work_book.save
        work_book.close
      end
    }
    sleep interval_time.to_i
    loop_times = loop_times.to_i - 1
  end

rescue Exception => e
  puts "Telnet logging failed: #{e}\n\n"
  puts $@.to_s
ensure
  spread_sheet.quit
end