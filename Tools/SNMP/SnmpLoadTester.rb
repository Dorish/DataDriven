require 'thread'

# Runs a given +testScript+ for +numIterations+ using +numUsers+ with a delay of +delayBetweenUsers+
# +testScript+ should be the full path of a test script.
# Results are logged in the specified +logFileName+
#
#NOTE:  Watir scipts executed using this loadRunner method need to run in seperate IE processes.  
#To do this be sure to use the following to create a new IE process in your testScript:
#   require 'watir/contrib/ie-new-process'  #This helpfull method is available as of watir development build 1.5.1.1100 see http://wiki.openqa.org/display/WTR/Development+Builds
#   ie = Watir::IE.new_process

def loadTestRunner(testScript,numItterations=1,numUsers=1,delayBetweenUsers=0, logFileName = Dir.pwd + '/myLoadTest.log')
  total_passing = 0
  total_failing = 0
  log = File.open(logFileName,"a")
  
  log.puts("\n\n*********************************\n**Load Test Configuration: numItterations=#{numItterations}, numUsers=#{numUsers}, delayBetweenUsers=#{delayBetweenUsers} seconds.")

  (1..numItterations).each do |iteration|
    itteration_passing = 0
    itteration_failing = 0
    log.puts("**Begining Itteration #{iteration} - #{Time.now.strftime(" %m/%d/%y @ %H:%M:%S ") }")
    i=0
    threads = []
    numUsers.times  do  #each thread simulates a user running the specified testScript
        threads << Thread.new do  
          i=i+1
          startTime = Time.now
          result = system("ruby #{testScript}")  #execute the script and wait for it to finish.
          duration = Time.now-startTime
          if result  #Record our results.
            itteration_passing = itteration_passing +1
            log.puts("Thread #{i} executed #{testScript} Successfully - Duration = #{duration}seconds")
          else 
            itteration_failing = itteration_failing + 1
            log.puts("Thread #{i} executed #{testScript} FAILED - Duration = #{duration}seconds")
          end
        end
      sleep(delayBetweenUsers)  #Wait for the specified delayBetweenUsers before we start the next user/thread.
    end
    threads.each {|t| t.join}
    log.puts("**Finished Itteration #{iteration}. #{itteration_passing} out of #{numUsers} tests Passed - #{Time.now.strftime(" %m/%d/%y @ %H:%M:%S ") }")
    total_passing = total_passing + itteration_passing
    total_failing = total_failing + itteration_failing
  end
  log.puts("**Load Test complete.  #{total_passing} out of #{numUsers*numItterations} tests Passed")
  log.close
end


#EXAMPLE USAGE:
#Here we are going to run 5 itterations of the 'googleSearch.rb' script with four concurent threads/users, using a 2 second delay between the start of each user.
#We will save/append the results to a file: 'myLoadTest.log'
loadTestRunner(Dir.pwd + '/snmp_1_walk.rb',5,4,2,Dir.pwd + '/myLoadTest.log')  
