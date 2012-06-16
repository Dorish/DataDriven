#require 'rubygems'



=begin
NetSNMP read
1) Convert the SNMP GET result into a string. This is 'my_string'
2) Split the string based on spaces. This is 'words'
3) Declare new array 'arr'
4) Convert 'words' into an array. This is 'arr'
5) Show Array length just for the sake of demonstration
6) Show the array
7) Pick the desired value from the array using the correct array 
   index number (location). This is 'my_value' 
8) Show the desired element 'my_value'
9) Convert 'my_value' to an integer (this is needed for a proper 
   compare when the value reaches the spreadsheet
10) Show my integer
=end


#1)
my_string=`snmpget -v2c -c LiebertEM 126.4.202.50 LIEBERT-GP-PDU-MIB::lgpPduPsLineEntryEpLN.1.1.1`.to_s

#2)
words = my_string.split(/ /) #split the string at each space. words = my_string.split(' ') works the same

#3)
arr = Array.new 

#4)
arr = words.to_a

#5)
puts arr.length #there are 5 elements in the array now. they are indexed 0 thru 4
puts "\n"

#6)
puts arr # show the array
puts "\n"

#7)
my_value = arr[3] # the value is a the 4th element which is index no. 4

#8)
puts my_value
puts "\n"

#9)
my_int = my_value.to_i

#10)
puts my_int
