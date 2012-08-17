# To change this template, choose Tools | Templates
# and open the template in the editor.

$:.unshift File.dirname(__FILE__)
require 'unity_navigate.rb'
require 'unity_setup.rb'
$:.unshift File.dirname(__FILE__).sub('controller\\unity_test','lib')
require 'generic'

def parse_test_case(navigate, case_parameters)
   rows = case_parameters["rows"]
   work_sheet_allcases = case_parameters["work_sheet"]
   test_site = case_parameters["test_site"]
   username = case_parameters["username"]
   password = case_parameters["password"]

  $ie.speed = :zippy
  $ie.maximize
  row = 1
  while(row <= rows)
   row +=1 
   type = work_sheet_allcases.Range("D#{row}")['Value']
    case type
      when 'Function'
        navigate.login(test_site,username,password)
      when 'Item'
        item_process(navigate,case_parameters,row)
      when 'Parse'
        sleep 1
      when 'Comment'
      
    else

    end
  end
end

def item_process(navigate,case_parameters,row)
   test_site = case_parameters["test_site"]
   username = case_parameters["username"]
   password = case_parameters["password"]
   work_sheet_allcases = case_parameters["work_sheet"]

   compenoent = work_sheet_allcases.Range("G#{row}")['Value']
   form_name = work_sheet_allcases.Range("F#{row}")['Value']
   arguments = work_sheet_allcases.Range("I#{row}")['Value']
   action = work_sheet_allcases.Range("H#{row}")['Value']
  case action
     when 'Navigate'
          navigate.navigate_node(compenoent).click
     when 'Set_TextValue'
          navigate.set_text_value(form_name, compenoent).set(arguments)
     when 'Click'
          navigate.click(compenoent).click
          navigate.login(test_site,username,password)
     when 'Set'
          navigate.set_check_value(form_name,compenoent).set
      when 'Clear'
          navigate.set_check_value(form_name,compenoent).clear
     when 'Set_FileField'
          navigate.set_filefield(form_name, compenoent).click_no_wait
          navigate.set_filefield(form_name, compenoent).set(arguments)
     when 'Select_Combo'
          navigate.select_combo(form_name, compenoent).select_value(arguments)
    else
   end
end


# initialize the connection.
def initialize_connect(navigate, set_up, execl_path, result_folder)
  parameters = set_up.connect_to_unity(execl_path, result_folder)

  # identity confirmation when the mode of web access is password protected site
  test_site = parameters["test_site"]
  username = parameters["username"]
  username = parameters["password"]
  #navigate.login(test_site,username,username)

  # navigate to unity configure page, tab4 is unity configuration tab id.
  navigate.unity_config("tab4")


  return parameters
end

begin
  generic = Generic.new
  navigate = Unity_Navigate.new
  set_up = Unity_SetUp.new
  execl_path = __FILE__.gsub(".rb",".xls")
  result_folder = File.dirname(__FILE__)
  parameters = initialize_connect(navigate,set_up,execl_path, result_folder)

  rows = parameters["rows"]
  work_sheet_allcases = parameters["work_sheet"]
  spread_sheet_allcases = parameters["spread_sheet"]

  row  = 1
  while (row <= rows)
    row += 1
    # run test case if the 'run' box is checked
    if work_sheet_allcases.Range("e#{row}")['Value'] == true
      puts "strat to run the cases......"
      #############################################
      #1. Parse one test case file
      #
      #2. Execute lines one by one
      #
      #3. Record the result
      #
      #4. Close this test case file
      #############################################

      #Parse one test case file
      case_file = File.dirname(__FILE__)<<work_sheet_allcases.Range("j#{row}")['Value']
      case_paras = set_up.parse_case(case_file,result_folder)
      #work_sheet_singlecases = case_paras["work_sheet"]
      spread_sheet_singlecases = case_paras["spread_sheet"]

      #Execute lines one by one
      parse_test_case(navigate, case_paras)

      # Close the active test case for one case.

      # reconnect to controller spreadsheet
      generic.conn_act_xls 
    end
  end
  rescue Exception => e
    puts "Executing failed: #{e}\n\n"
    puts $@.to_s
ensure
    # Close the active test case before exit.
end