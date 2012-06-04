# To change this template, choose Tools | Templates
# and open the template in the editor.

$:.unshift File.dirname(__FILE__).sub('Tools','lib') #add lib to load path
require 'generic'
require 'find'
require 'fileutils'
require "rexml/document"
include REXML

#$dir_test_case = ['D:\enpc_work\TestLog\Test Cases']
#$dir_project = ['D:\enpc_work\TestLog\Project']
$dir_test_case = ['I:\lmg test engineering\TestLog\Test Cases']
$dir_project = ['I:\LMG TEST ENGINEERING\TestLog\Project\project_test\Project Test Cases']
$template_case_name = "test_case_template.xml"

# User select the card type for create its test case.
def select_card_type(cardtypes_list)
  puts "The following card types are available to be selected:"
  index = 0
  while(index < cardtypes_list.length )
    print index + 1,' - ',cardtypes_list[index],"\n"
    index = index + 1
  end
  puts "Please type the number of the desired type followed by <Enter>"
  input_str = gets.chomp.to_i
  return cardtypes_list[input_str - 1]
end

#Pickup all card types from spreadsheet
def pickup_all_cards_type(work_sheet)
    cardtypes_list = Array.new(4)
    cardtypes_list[0] = work_sheet.Range("H1")['Value']
    cardtypes_list[1] = work_sheet.Range("I1")['Value']
    cardtypes_list[2] = work_sheet.Range("J1")['Value']
    cardtypes_list[3] = work_sheet.Range("K1")['Value']
    return cardtypes_list
end

# select the column number for selected card type
def select_type_column(card_type,work_sheet)
  columns = ['H','I','J','K']
  index = nil
  columns.each { |i|
    if(card_type == work_sheet.Range(i + '1')['Value'])
      index = i
    end
  }
  return index
end

# pick up all test cases which user selected in the spreadsheet
def pickup_selected_cases(card_type_column,work_sheet)
  index = 2
  cases_list = Array.new
    while(work_sheet.Range("A#{index}")['Value'] != nil)
    if work_sheet.Range(card_type_column + "#{index}")['Value'] == 'X' && work_sheet.Range('A' +"#{index}")['Value'] == 'case'
      cases_list.push(index)
    end
    index = index + 1
  end
  return cases_list
end

# create the execute test cases project
def create_executable_project(cases_list,work_sheet)
  cases_list = cases_list.sort
  cases_list.each { |i|
    cases_name = work_sheet.Range("B#{i}")['Value']
    target_path = "#{$dir_project}\\#{cases_name}"
    source_path = "#{$dir_test_case}\\#{cases_name}"

     if !File.exists?(target_path)
       create_folder(cases_name)
     end

    case_id =  work_sheet.Range("C#{i}")['Value'] + '.tlg'
    source_file = source_path +'\\' + case_id
    convert_to_executable_cases(case_id, target_path,source_file)
  }
end

# convert the cases to executable ones
def convert_to_executable_cases(case_id, target_path,source_file)
  template_case = File.dirname(__FILE__) <<"\\"<<$template_case_name
  FileUtils.copy template_case, target_path
  Dir.chdir(target_path)
  File.rename($template_case_name, case_id)
  new_case = target_path<<"\\"<<case_id
 
  doc_source = Document.new(File.open(source_file))
  doc_target = Document.new(File.open(new_case))

  id, test_title, test_duration, test_type, test_phase,create_date, create_time, update_time = nil
  update_date, create_user_id, update_user_id, prerequisites,test_description,notes2,priority = nil
  doc_source.elements.each("test_case"){  |elem|
        id = elem.elements["id"].text
        test_title = elem.elements["test_title"].text
        test_duration = elem.elements["test_duration"].text
        test_type = elem.elements["test_type"].text
        test_phase = elem.elements["test_phase"].text
        create_date = elem.elements["create_date"].text
        create_time = elem.elements["create_time"].text
        update_date = elem.elements["update_date"].text
        update_time = elem.elements["update_time"].text
        create_user_id = elem.elements["create_user_id"].text
        update_user_id = elem.elements["update_user_id"].text
        prerequisites = elem.elements["prerequisites"].text
        test_description = elem.elements["test_description"].text
        notes2 = elem.elements["notes2"].text
        priority = elem.elements["priority"].text
  }
  time = Time.now.strftime("/%H/%M/%S")
  date = Time.now.strftime("/%d/%m/%Y")

  doc_target.elements.each("project_test_case"){  |elem|
        elem.elements["id"].add_text( "#{id}")
        elem.elements["test_title"].add_text("#{test_title}")
        elem.elements["test_duration"].add_text( "#{test_duration}")
        elem.elements["test_type"].add_text( "#{test_type}")
        elem.elements["test_phase"].add_text( "#{test_phase}")
        elem.elements["create_date"].add_text( "#{create_date}")
        elem.elements["create_time"].add_text( "#{create_time}")
        elem.elements["update_date"].add_text( "#{update_date}")
        elem.elements["update_time"].add_text( "#{update_time}")
        elem.elements["create_user_id"].add_text( "#{create_user_id}")
        elem.elements["update_user_id"].add_text( "#{update_user_id}")
        elem.elements["prerequisites"].add_text( "#{prerequisites}")
        elem.elements["test_description"].add_text( "#{test_description}")
        elem.elements["notes2"].add_text( "#{notes2}")
        elem.elements["priority"].add_text( "#{priority}")
        elem.elements.each("history_entry"){|e|
          e.elements["date"].add_text( "#{date}")
          e.elements["time"].add_text( "#{time}")
        }
  }
  file_target = File.open(new_case, "w")
  file_target.write(doc_target)
  file_target.close
end

# create the project folder and their tgp files
def create_folder(case_name)
    path = nil
    split_list = case_name.split("\\")
    split_list.each{ |substr|
    if path == nil
      path = substr
    else
      path = path +"\\" + substr
    end

    target_path = "#{$dir_project}\\#{path}"
    source_path = "#{$dir_test_case}\\#{path}"

    Dir.mkdir target_path unless File.exist? target_path
    father_dir =  File.dirname(target_path)
    tgp_file = "#{source_path}.tgp"
    if File.exists?(tgp_file) && !File.exists?("#{target_path}.tgp")
        FileUtils.copy tgp_file,father_dir
    end
  }
end

begin
  g = Generic.new
  excel_name = __FILE__.sub('.rb','.xls')
  setup = g.new_xls(excel_name,1)
  spreat_sheet = setup[0]
  work_book = setup[1]
  work_sheet = setup[2]

 cardtypes_list = pickup_all_cards_type(work_sheet)
 card_type = select_card_type(cardtypes_list)
 card_type_column = select_type_column(card_type,work_sheet)

 cases_list = pickup_selected_cases(card_type_column,work_sheet)
 create_executable_project(cases_list,work_sheet)

rescue Exception => e
   puts "Create the project failed: #{e}\n\n"
   puts $@.to_s
ensure
  work_book.close
  spreat_sheet.quit
end