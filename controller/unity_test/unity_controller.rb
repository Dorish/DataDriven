# To change this template, choose Tools | Templates
# and open the template in the editor.

$:.unshift File.dirname(__FILE__)
require 'unity_actions.rb'
require 'unity_setup.rb'

include Unity_actions
include Unity_SetUp


def parse_test_case(case_parameters)
  work_sheet_allcases = case_parameters["work_sheet"]
  ss = case_parameters["work_book"]
  $ie.speed = :zippy
  $ie.maximize
  row = 1
  while(work_sheet_allcases.Range("C#{row}")['Value'] != nil)
    row +=1
    type = work_sheet_allcases.Range("D#{row}")['Value']
    case type
    when 'Function'

    when 'Item'
      item_process(case_parameters,row)
    when 'Pause'
      item_process(case_parameters,row)
    when 'Comment'
      
    else

    end
  end
  ss.save
end

def item_process(case_parameters,row)
  test_site = case_parameters["test_site"]
  work_sheet_allcases = case_parameters["work_sheet"]

  componet = work_sheet_allcases.Range("G#{row}")['Value']
  arguments1 = work_sheet_allcases.Range("I#{row}")['Value']
  arguments2 = work_sheet_allcases.Range("J#{row}")['Value']
  action = work_sheet_allcases.Range("H#{row}")['Value']
  puts "------------Start to #{componet}------------------"
  case action
  when 'Navigate'
    navigate_to(componet)

  when 'Set_TextBox'
    set_textbox(componet,arguments1)

  when 'Click'
    clickbtn(componet)
    sleep 1

  when 'Login'
    login(test_site,arguments1,arguments2)

  when 'Set_CheckBox'
    set_checkbox(componet,arguments1)

  when 'Select_ComboBox'
    select_combobox(componet,arguments1)

  when 'WaitSave'
    waitsave(arguments1)

  when 'Jsclick'
    jsClick('OK')

  when 'Verify_Result'
    verify_result(componet,arguments1,work_sheet_allcases,row)
  else
    puts "Not Define yet"
  end
end


# initialize the connection.
def initialize_connect(execl_path, result_folder)

  parameters = connect_to_unity(execl_path, result_folder)

  # navigate to unity configure page, tab4 is unity configuration tab id.
  unity_config("tab4")
  return parameters
end

begin
  execl_path = __FILE__.gsub(".rb",".xls")
  result_folder = File.dirname(__FILE__)
  parameters = initialize_connect(execl_path, result_folder)

  rows = parameters["rows"]
  work_sheet_allcases = parameters["work_sheet"]
  spread_sheet_allcases = parameters["work_book"]

  row  = 1
  while (row <= rows)
    row += 1
    # run test case if the 'run' box is checked
    if work_sheet_allcases.Range("e#{row}")['Value'] == true
      puts "strat to run the cases......"

      #Parse one test case file
      case_file = File.dirname(__FILE__)<<work_sheet_allcases.Range("j#{row}")['Value']
      case_paras = parse_case(case_file,result_folder)
      #work_sheet_singlecases = case_paras["work_sheet"]
      #spread_sheet_singlecases = case_paras["spread_sheet"]

      #Execute lines one by one
      parse_test_case(case_paras)

      # Close the active test case for one case.

      # reconnect to controller spreadsheet
       
    end
  end
rescue Exception => e
  puts "Executing failed: #{e}\n\n"
  puts $@.to_s
ensure
  # Close the active test case before exit.
  #spread_sheet_allcases.save
end