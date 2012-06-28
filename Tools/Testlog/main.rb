# == Synopsis
#
# This class represents test cases in test log for bulk manipulation.
# 


require 'find'
require 'rexml/document'

# The match method extension to the find module was "borrowed"
# from the O'reilly Ruby Cookbook Chapter 6 'Files and Directories' Page 234
module Find
  def match(*paths)
    matched = []
    find(*paths) { |path| matched << path if yield path }
    return matched
  end
  module_function :match
end

# A Class to represent Test Log for manipulation in Ruby scripts
class Test_log
  # BACKUP_EXT file name extension used when a file is replaced using the replace_file method.
  BACKUP_EXT = '.frb'

  attr_reader :test_cases, :test_suites, :root_directory
  attr_accessor :current_directory
  
  # This class variable represents the xml element ids that constitute a Testlog test case.
  @@valid_id_array = Array['id',	'test_title',	'test_duration',	'test_type',
    'test_phase',	'rec_test_configs',	'test_resources',	'test_status',
    'test_attempts',	'last_attempt_date',	'last_attempt_time',
    'actual_duration',	'testers',	'fault_report_id',	'change_req_id',
    'results_obtained',	'version_tested',	'build_tested',	'external_link2',
    'result_notes',	'requirement',	'create_date',	'create_time',
    'update_date',	'update_time',	'create_user_id',	'update_user_id',
    'prerequisites',	'test_description',	'test_result',	'notes1',	'notes2',
    'external_link',	'priority',	'history_entry']

  # Creates a new Testlog object.  The object contains all test cases and test suites.
  # - _root_directory_ is a folder in a TestLog database (i.e. file structure)
  def initialize(root_directory)
    @root_directory = root_directory
    @current_directory = root_directory
    @test_cases = Find.match(root_directory) { |p| ext = p[-4...p.size]; ext && ext.downcase == ".tlg" }
    @test_suites = Find.match(root_directory) { |p| ext = p[-4...p.size]; ext && ext.downcase == ".tgp" }
  end

  # Replaces a test case file and creates a backup of the previous file with extension of BACKUP_EXT
  def replace_file(to_be_replaced, replacement)
    counter = 1
    while (line = to_be_replaced.gets)
      line.strip!
      File.copy(line, line + BACKUP_EXT)
      update_test_case(line, 'prerequisites', retrieve_test_case_value(replacement,'prerequisites'))
      update_test_case(line, 'test_description', retrieve_test_case_value(replacement,'test_description'))
      update_test_case(line, 'requirement', retrieve_test_case_value(replacement,'requirement'))
      update_test_case(line, 'notes2', retrieve_test_case_value(replacement,'notes2')) #Revision History
      update_test_case(line, 'external_link', retrieve_test_case_value(replacement,'external_link'))
      counter = counter + 1
    end
    to_be_replaced.close
    puts "#{counter-1} file(s) replaced"
  end

  # Finds instances of a file (test case) and stores a listing of those files in a file called found_files
  def find_file(file,found_files)
    full_file_path = file.gsub('\\','/')
    file = File.basename(file)

    i=0
    Find.find(@root_directory) do |path|
      if FileTest.directory?(path)
        if File.basename(path)[0] == ?. and File.basename(path) != '.'
          Find.prune
        else
          next
        end
      else if File.fnmatch?(file, File.basename(path)) and !File.fnmatch?(full_file_path, path.gsub('\\','/'))
          found_files.puts(path)
          #results.puts(File.basename(path) + " found in " + File.dirname(path))
          #FileUtils.mv(path, temp_dir + File.basename(path) + i.to_s)
          i += 1
        end
      end
    end

    puts "Found #{i} instances of #{file}.  A listing of the file(s) are located in #{found_files.path}"
  end

  # Updates a specific test case file attribute _id_ to _value_. 
  # - _test_case_file_ is the *full path*
  def update_test_case(test_case_file, id, value)

  index = @@valid_id_array.index(id)

  case index
  when nil then puts "Invalid update statement: #{id} is not a valid test case field."; exit;
  when 15 then cdata = true
  when 19 then cdata = true
  when 27..31 then cdata = true
  else cdata = false
  end

  if cdata == true and value != nil
    if File.exists?(value)
      temp = ''
      File.open(value) do |file|
        file.each do |line|
        temp += line.chomp
        end
      end
    value = REXML::CData.new(temp)
    else
    value = REXML::CData.new(value)
    end
  end


  if File.exists?(test_case_file)
    File.open(test_case_file) do |config_file|
      # Open the document and edit the value associated with id
      config = REXML::Document.new(config_file)
      config.root.elements[id].text = value

      # Write the result to a new file.
      formatter = REXML::Formatters::Default.new
      File.open(test_case_file, 'w') do |result|
      formatter.write(config, result)
      end
    end
  else
    puts "Invalid file name: \"#{test_case_file}\" is not a valid test case file."
    exit
  end
  end

  # Updates all test cases with attribute _id_ to value
  # - Example below will update all test cases to an actual duration of 1 minute
  #  project = Test_log.new('I:\LMG TEST ENGINEERING\TestLog\Projects\159252 - V4C-II Alpha\')
  #  project.update_test_case_all('actual_duration','0:01')
  def update_test_case_all(id, value)
    @test_cases.each do |test_case|
      self.update_test_case(test_case,id,value)
      puts test_case
    end
  end

  # Retrieves the value of a specific test case field
  def retrieve_test_case_value(test_case_file, id)

  index = @@valid_id_array.index(id)

  case index
  when nil then puts "Invalid update statement: #{id} is not a valid test case field."; exit;
  when 15 then cdata = true
  when 19 then cdata = true
  when 27..31 then cdata = true
  else cdata = false
  end

    if File.exists?(test_case_file)
      File.open(test_case_file) do |config_file|
        if cdata == true # This conditional and flag may be unneccessary
          config = REXML::Document.new(config_file)
          value = config.root.elements[id].text
          return value
        else
          config = REXML::Document.new(config_file)
          value = config.root.elements[id].text
          return value
        end
      end
    end

    puts "Invalid test case file specified"
    exit

  end

end

v4_cooling = Test_log.new('I:\lmg test engineering\TestLog\Projects\159626 - V4 Cooling IV(PA Foxtrot)\Project Test Cases\IS-WEBL\APM\SNMP\1_Monitor\1_Input')
v4_cooling.update_test_case_all('actual_duration','00:05')



