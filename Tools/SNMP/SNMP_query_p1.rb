
# this is the command being used:
# snmpget -v2c -c LiebertEM 126.4.203.252 LIEBERT-GP-POWER-MIB::lgpPwrNoLoadWarningLimit.0

def snmp(ip,snmp_query)
	command='snmpget -v2c -c LiebertEM '<< ip << ' ' << snmp_query
	#Convert to array by splitting the string at each space, then get value of 4th element[3]
	snmp_data =`#{command}`.to_s.split(/ /)[3]
	return snmp_data
end


# Test the function
ip = '126.4.203.252'
snmp_query = 'LIEBERT-GP-POWER-MIB::lgpPwrNoLoadWarningLimit.0'

puts snmp(ip,snmp_query)


