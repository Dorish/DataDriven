 #This snippet demonstrates setting snmp oids - in this case for the Nfinity	UPS
	
	i = 0
while i < 100
	cmd = 'Snmpset -v2c -c LiebertEM 126.4.202.195 LIEBERT-GP-POWER-MIB::lgpPwrNoLoadWarningLimit.0 i 20'
	puts %x{#{cmd}}
	sleep 5
	cmd = 'Snmpset -v2c -c LiebertEM 126.4.202.195 LIEBERT-GP-POWER-MIB::lgpPwrNoLoadWarningLimit.0 i 0'
	puts %x{#{cmd}}
	puts "loop #{i}"
	sleep 5
	i+=1
end