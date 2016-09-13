#$Language="VBScript"
#$Interface="1.0"

' -----------------------------------------------------
' Southeast IAD Update
'
' Updates the ACL list on list of IAD devices
' 
'
' Eric Hansen
' 9/8/16
' ----------------------------------------------------- 

'Credentials List
g_SERUsername = "SANITIZED"
g_SERPassword = "SANITIZED"
' GLA
g_MarketUsername = "SANITIZED" 
g_MarketPassword = "SANITIZED"
g_MarketSecret = "SANITIZED"
'CF
'g_MarketUsername = "SANITIZED" 
'g_MarketPassword = "SANITIZED"
'g_MarketSecret = "SANITIZED"
'g_MarketSecret = "SANITIZED"

Sub Main()
	
	' Ask the user where the list of IADs is
	filePath = crt.Dialog.FileOpenDialog("Select Subnet List", "Open", "%USERPROFILE%\Desktop\IAD_List.txt")
	
	' Nothing was chosen, or the dialog box was canceled
	if filePath = "" then 
		crt.Dialog.MessageBox("No IAD List Selected - Canceling")
		exit sub
	end if
	
	' Create the File IO object for reading from a file
	Set objFSI = CreateObject("Scripting.FileSystemObject")
	
	' Create the File IO object for writing to the results file
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	' Open the file that the user selected
	Set objStream = objFSI.OpenTextFile(filepath)
	
	' Open the file for the results
	filePathParts = split(filepath, ".")
	Set objOutStream = objFSO.CreateTextFile(filePathParts(0) & "_Report.csv", true)
	
	' Write the CSV File Header
	objOutStream.WriteLine("IP Address,Access Method,Accessible?,TACACS?,Device Type,ACL ID,Comments")
	
	' Start a counter
	deviceCounter = 1
	
	' Start reading tht file stream
	Do While objStream.AtEndOfStream <> True
	
	' Grab the first device IP
	strLine = objStream.ReadLine
	
		' Check if a comment
		if(InStr(1, strLine, "//") <> 1) then
			' Make sure we're starting at a bash prompt

			row = crt.screen.CurrentRow

			' Get the contents of that row
			rowString = crt.Screen.Get(row, 1, row, 80)
	
			' Make sure the row contains both '@' and ':'
			if(InStr(rowString, "@") = 0 OR InStr(rowString, ":") = 0) then
			
				' Display an error message
				crt.Dialog.MessageBox("An error has occured. Expected to be back at the Linux prompt")
				Exit Sub
			end if
		
			' Echo this to the screen
			crt.screen.send chr(13) & chr(13)
			crt.sleep 10
			crt.screen.send "############################################" & chr(13)
			crt.sleep 10
			crt.screen.send "#          Beginning IAD Operation         " & chr(13)
			crt.sleep 10
			crt.screen.send "#                                          " & chr(13)
			crt.sleep 10
			crt.screen.send "# IAD Device Number: " & deviceCounter & chr(13)
			crt.sleep 10
			crt.screen.send "# IP Address: " & strLine & chr(13)
			crt.sleep 10
			crt.screen.send "############################################" & chr(13) & chr(13) & chr(13)
			crt.sleep 10

			' Add the IP to the output report
			outputLine = strLine & ","
			
			' Determine if device supports SSH (Some IAD devices are Telnet)
			checkResult = CheckForSSHAccess(strLine)
			' ********************************************** SSH Section ******************************************************
			if(checkResult = 1) then
				outputLine = outputLine & "SSH,Yes,"
				
				' Try the SER login for the device
				loginResult = SSHLogIntoDevice(strLine, g_SERUsername, g_SERPassword)
				
				' Did the ACS login work?
				if(loginResult = 1) then
					' Add the 'Tacacs?' result
					outputLine = outputLine & "Yes,"
					
					' We're logged into an IAD right?
					if(IsAnIAD = 1) then
						
						' Mark this as an 'IAD' device
						outputLine = outputLine & "IAD,"
						
						' Start on the ACL work
						outputLine = outputLine & UpdateACL()
					
					else
						
						' Mark this as an 'Other' device
						outputLine = outputLine & "Other,N/A,Non-IAD Device - Could be Edgemarc or a system router"					
					
					end if
					
					' End session
					' Cancel this session "exit"
					crt.Screen.Send "exit" & chr(13)
			
					crt.sleep 1000
				else ' No ACS
					' Add the 'Tacacs?' result
					outputLine = outputLine & "No,"
					
					' Cancel previous session ^C
					crt.Screen.Send chr(3)
					
					' Try again with market password
					loginResult = SSHLogIntoDevice(strLine, g_MarketUsername, g_MarketPassword)
					
					' Did the Market login work?
					if(loginResult = 1) then
						
						' Mark this as an 'IAD' device
						outputLine = outputLine & "IAD,"
						
						' Start on the ACL work
						outputLine = outputLine & UpdateACL()
						
						' End session
						' Cancel this session "exit"
						crt.Screen.Send "exit" & chr(13)
				
						crt.sleep 1000
					else
					
						' Cancel previous session ^C
						crt.Screen.Send chr(3)
						
						crt.sleep 1000
											
						' Mark this as an 'Other' device
						outputLine = outputLine & "Other,N/A,Non-IAD Device - Could be Edgemarc or a system router"
						
					end if
					
				end if
			' ********************************************** Telnet Section ******************************************************	
			elseif (checkResult = 2) then
				outputLine = outputLine & "Telnet,Yes,"
				
				' Try the SER login for the device
				loginResult = TelnetLogIntoDevice(strLine, g_SERUsername, g_SERPassword)
				
				' Did the ACS login work?
				if(loginResult = 1) then
					' Add the 'Tacacs?' result
					outputLine = outputLine & "Yes,"
					
					' We're logged into an IAD right?
					if(IsAnIAD = 1) then
						
						' Mark this as an 'IAD' device
						outputLine = outputLine & "IAD,"
						
						' Start on the ACL work
						outputLine = outputLine & UpdateACL()
						
					else
						
						' Mark this as an 'Other' device
						outputLine = outputLine & "Other,N/A,Non-IAD Device - Could be Edgemarc or a system router"						
					
					end if
					
					' End session
					' Cancel this session "exit"
					crt.Screen.Send "exit" & chr(13)
			
					crt.sleep 1000
					
				else ' No ACS
					' Add the 'Tacacs?' result
					outputLine = outputLine & "No,"
					
					' Cancel previous session ^C
					crt.Screen.Send chr(3)
					
					crt.sleep 1000
					
					crt.Screen.Send chr(3)
					
					crt.sleep 3000					
					
					' Try again with market password
					loginResult = TelnetLogIntoDevice(strLine, g_MarketUsername, g_MarketPassword)
					
					' Did the Market login work?
					if(loginResult = 1) then
					
						' Mark this as an 'IAD' device
						outputLine = outputLine & "IAD,"
						
						' Start on the ACL work
						outputLine = outputLine & UpdateACL()
						
						' End session
						' Cancel this session "exit"
						crt.Screen.Send "exit" & chr(13)
				
						crt.sleep 1000
					
					else
						
						' Cancel previous session ^C
						crt.Screen.Send chr(3)
						
						crt.sleep 1000
						
						crt.Screen.Send chr(3)
						
						crt.sleep 3000					
					end if
					
				end if
			' ********************************************** No Connectivity Section ******************************************************				
			else
				outputLine = outputLine & "N/A,No,N/A,N/A,N/A,Device appears to be offline"
			end if
			
			
			' Write the output to the report
			objOutStream.WriteLine(outputLine)
			
			' Increment the device counter
			deviceCounter = deviceCounter +1	
			
		end if ' Comment check

	Loop ' File read loop

	
End Sub

Function SSHLogIntoDevice(IPaddress, Username, Password)

	' Send the command to attempt a login
	crt.Screen.Send "ssh -o StrictHostKeyChecking=no -o ConnectTimeout=2 " & Username & "@" & IPAddress & chr(13)
	
	' Give enough time for a connection and a response
	valid = crt.Screen.WaitForStrings("password","Password","passcode","Passcode", 2)
	
	if (valid > 0) then ' We were prompted for password
		
		' Try the password
		crt.Screen.Send Password & chr(13)
		
		' Wait for the result
		passvalid = crt.Screen.WaitForStrings(">","#", 3)
		
		if(passvalid > 0 ) then
			'Login was successful
			SSHLogIntoDevice = 1
			
			' Courtesy timer
			crt.sleep 250
			
			exit function
		else
			' Login unsuccessful
			SSHLogIntoDevice = 0
			
			' Courtesy timer
			crt.sleep 250
			
			exit function
		end if
	else ' Password prompt never came
		' Timed out
		SSHLogIntoDevice = 0
		
		' Courtesy timer
		crt.sleep 250
	end if

End Function

Function TelnetLogIntoDevice(IPaddress, Username, Password)

	' Send the command to attempt a login
	crt.Screen.Send "telnet " & IPAddress & chr(13)
	
	' Give enough time for a connection and a response
	valid = crt.Screen.WaitForStrings("Username","username", 4)
	
	if (valid > 0) then ' We were prompted for username
		
		' Courtesy timer
		crt.sleep 500	
			
		' Try the username
		crt.Screen.Send Username & chr(13)
		
		' Give enough time for a connection and a response
		valid = crt.Screen.WaitForStrings("Password","password","passcode","Passcode", 4)
		
		if (valid > 0) then ' We were prompted for password
		
			' Send the command to attempt a login
			crt.Screen.Send  Password & chr(13)		

			' Wait for the result
			passvalid = crt.Screen.WaitForStrings(">","#", 5)
			
			if(passvalid > 0 ) then
				'Login was successful
				TelnetLogIntoDevice = 1
				
				' Courtesy timer
				crt.sleep 500
				
				exit function
			else
				' Login unsuccessful
				TelnetLogIntoDevice = 0
				
				' Courtesy timer
				crt.sleep 500
				
				exit function
			end if
		else ' Never prompted for password
			' Timed out
			TelnetLogIntoDevice = 0
		
			' Courtesy timer
			crt.sleep 250
		end if
	else ' Never prompted for username
		
		' Timed out
		TelnetLogIntoDevice = 0
		
		' Courtesy timer
		crt.sleep 250
	end if

End Function

Function CheckForSSHAccess(IPAddress)

	' Option 1 - Use NMap to check if the port is open or closed (SLOWER but more elegant)
	' Send the command to namp the host and look for ssh support
	'crt.Screen.Send "nmap -p22 " & IPAddress & " | grep ssh" & chr(13)
	
	' Option 2 - try to ssh into the device and see if we get "Connection refused" (QUICKER, but more of a hack)
	' Send the command to attempt a login
	crt.Screen.Send "ssh -o StrictHostKeyChecking=no -o ConnectTimeout=2 InvalidUser@" & IPAddress & chr(13)
	
	' Give enough time for a connection and a response
	valid = crt.Screen.WaitForStrings("password","Password","passcode","Passcode", 2)
	
	' Courtesy timer
	crt.sleep 250
	
	' Check if we were given the password prompt or not
	if(valid > 0) then
		' SSH responded, clear to login
		CheckForSSHAccess = 1
		
		' Cancel this session ^C
		crt.Screen.Send chr(3)
		
		' Courtesy timer
		crt.sleep 500
		
		exit function
	end if
	
	' Check for a "Connection refused"
	
	' Get the row number of the route result
	row = crt.screen.CurrentRow - 1

	' Get the contents of that row
	rowString = crt.Screen.Get(row, 1, row, 80)
	
	' Was the connection refused, or did the connection timeout?
	if(InStr(rowString, "Connection refused") > 0) then

		' Connection was refused
		CheckForSSHAccess = 2
		
		' Courtesy timer
		crt.sleep 500
		
		exit function
	else
		' Connection must have timed out
		CheckForSSHAccess = 3
		
		' Courtesy timer
		crt.sleep 500
		
		exit function
	end if	

End Function

Function IsAnIAD()
	' Courtesy timer
	crt.sleep 250

	' Execute a 'show version'
	crt.Screen.Send "show version | inc IAD" & chr(13)
	
	' Courtesy timer
	crt.sleep 750
	
	' Get the row number of the route result
	row = crt.screen.CurrentRow - 1

	' Get the contents of that row
	rowString = crt.Screen.Get(row, 1, row, 80)
	
	' Did the | include return the string IAD? : Make sure we're not false triggering by reading back the show command
	if(InStr(rowString, "IAD") > 0 AND InStr(rowString, "show") = 0) then
	
		' Yes we verified this is indeed an IAD
		IsAnIAD = 1

	else
	
		' Not an IAD
		IsAnIAD = 0
		
	end if
	
	' Courtesy timer
	crt.sleep 250
	
End Function

Function UpdateACL()

	' Make sure we're in privileged exec mode
	if(ExecCheck() = 1) then
	
		' Ready to start making changes
		
		' Get the access-class
		aclNumber = GetAccessClass()
		
		if( aclNumber > -1) then
		
			' Track how many updates we had to do
			numberChangesMade = 0
			
			' Good ACL number
			UpdateACL = aclNumber & ","
			
			
			' ************* 24.248.74.254 **************************************
			' Search for each of the required addresses in this access list
			if(ExistsInList("24.248.74.254", aclNumber) = 1) then
				
				' No action to take
				
			else
			
				' Add this address to the list
				success = AddToList("24.248.74.254", aclNumber)
				
				' Report that we added the entry
				UpdateACL = UpdateACL & "Added: 24.248.74.254    "
				
				' Increment the changes counter
				numberChangesMade = numberChangesMade + 1
				
			end if
			
			' ************* 98.178.246.199 **************************************
			' Search for each of the required addresses in this access list
			if(ExistsInList("98.178.246.199", aclNumber) = 1) then
				
				' No action to take
				
			else
			
				' Add this address to the list
				success = AddToList("98.178.246.199", aclNumber)
				
				' Report that we added the entry
				UpdateACL = UpdateACL & "Added: 98.178.246.199    "
				
				' Increment the changes counter
				numberChangesMade = numberChangesMade + 1
				
			end if
			
			' ************* 68.12.15.83 **************************************
			' Search for each of the required addresses in this access list
			if(ExistsInList("68.12.15.83", aclNumber) = 1) then
				
				' No action to take
				
			else
			
				' Add this address to the list
				success = AddToList("68.12.15.83", aclNumber)
				
				' Report that we added the entry
				UpdateACL = UpdateACL & "Added: 68.12.15.83    "
				
				' Increment the changes counter
				numberChangesMade = numberChangesMade + 1
				
			end if
			
			' ************* 174.79.30.7 **************************************
			' Search for each of the required addresses in this access list
			if(ExistsInList("174.79.30.7", aclNumber) = 1) then
				
				' No action to take
				
			else
			
				' Add this address to the list
				success = AddToList("174.79.30.7", aclNumber)
				
				' Report that we added the entry
				UpdateACL = UpdateACL & "Added: 174.79.30.7    "
				
				' Increment the changes counter
				numberChangesMade = numberChangesMade + 1
				
			end if
			
			' Check if we made any changes
			if(numberChangesMade > 0) then
			
				' Save the configuration
				crt.Screen.Send "wr" & chr(13)
	
				' Wait for the save to finish
				success = crt.Screen.WaitForString("[OK]", 10)
				
				' Courtesy timer
				crt.sleep 500
				
			else
				' Report that we made no changes
				UpdateACL = UpdateACL & "All required ACL entries were found - No Changes Made"	
				
			end if
			
		else
		
			' Not a good ACL number
			UpdateACL = "Failed, Valid ACL number not found"
			
		end if
		
	else
	
		' Was not able to get permission to make changes
		UpdateACL = "Failed,No Changes were made to the ACL on the device"
		
	end if


End Function

Function ExecCheck()

	' Get the row number of the current row
	row = crt.screen.CurrentRow

	' Get the contents of that row
	rowString = crt.Screen.Get(row, 1, row, 80)
	
	' Privileged exec is noted by the '#'
	if(InStr(rowString, "#") > 0) then
	
		' Already in privileged exec mode
		ExecCheck = 1
		Exit Function
		
	else
	
		' Enable privileged exec mode
		crt.Screen.Send "enable" & chr(13)
		
		' Courtesy timer
		crt.sleep 500
		
		' Try the password
		crt.Screen.Send g_MarketSecret & chr(13)
		
		' Courtesy timer
		crt.sleep 500
		
		' Get the row number of the current row
		row = crt.screen.CurrentRow

		' Get the contents of that row
		rowString = crt.Screen.Get(row, 1, row, 80)
	
		' Privileged exec is noted by the '#'
		if(InStr(rowString, "#") > 0) then
		
			' Was able to log in with the exec pass
			ExecCheck = 1
			Exit Function
			
		else
			
			' Was not able to log in with the exec pass
			ExecCheck = 0
			Exit Function
			
		end if
		
		
	end if
	
End Function

Function GetAccessClass()

	' Send the request to show the line vty configuration
	crt.Screen.Send "show run | section line vty 0 4" & chr(13)
	
	' Courtesy timer
	crt.sleep 2000
	
	'Check each of the last 15 lines to see if we can find the access-class
	lineCount = 1
	configText = ""
	
	do while lineCount < 16
	
		' Get the row number of the current row
		row = crt.screen.CurrentRow - lineCount

		' Get the contents of that row
		rowString = crt.Screen.Get(row, 1, row, 80)

		' Look for the access-class string
		if(InStr(rowString, "access-class") > 0) then

			' Stop the loop
			lineCount = 99
			
			' Grab the line
			configText = rowString
		else		
				
			' Increment the line counter
			lineCount = lineCount + 1
			
		end if

	loop
	
	' Did we find the access-class?
	if(lineCount = 99) then
	
		' Get the number from the string
		stringParts = split(configText, " ")
		
		' Is the ACL number in the standard access list range?
		aclNumber = stringParts(2)
		
		if(aclNumber > 0 AND aclNumber < 100) then

			' Return this value
			GetAccessClass = aclNumber
			Exit Function
			
		else
		
			' Send back the invalid access-class
			GetAccessClass = -1
			Exit Function
			
		end if
	
	else
	
		' Send back the invalid access-class
		GetAccessClass = -1
		Exit Function
		
	end if

End Function

Function ExistsInList(addressToFind, aclNumber)

	' Send the request to show the contencts of the ACL
	crt.Screen.Send "show access-lists " & aclNumber & " | inc " & addressToFind & chr(13)
	
	' Courtesy timer
	crt.sleep 2000	
		
	' Get the row number of the current row
	row = crt.screen.CurrentRow - 1

	' Get the contents of that row
	rowString = crt.Screen.Get(row, 1, row, 80)

	' Look for the IP address string
	if(InStr(rowString, addressToFind) > 0 AND InStr(rowString, "show") = 0) then
		
		' Address is already in list
		ExistsInList = 1
		
	else
		'Address is not found
		ExistsInList = 0
		
	end if

End Function

Function AddToList(addressToAdd, aclNumber)

	' Send the request to view the entries in the access-list
	crt.Screen.Send "show access-lists " & aclNumber & chr(13)
	
	' Courtesy timer
	crt.sleep 1500

	' ************** Get the sequence number of the last entry (DENY) **************
	' Get the row number of the current row
	row = crt.screen.CurrentRow - 1

	' Get the contents of that row
	rowString = crt.Screen.Get(row, 1, row, 80)
	stringParts = split(rowString, " ")
	denySeq = stringParts(4)
	
	' ************** Get the sequence number of the second to last entry (PERMIT) **************
	' Get the row number of the current row
	row = crt.screen.CurrentRow - 2

	' Get the contents of that row
	rowString = crt.Screen.Get(row, 1, row, 80)
	stringParts = split(rowString, " ")
	permitSeq = stringParts(4)
	
	' Generate the next sequence number
	nextSeq = permitSeq + 1
	
	' Is there enough room to add this entry?
	if(nextSeq < denySeq) then
	
		'Add the entry
		
		' Send the request to enter global configuration mode
		crt.Screen.Send "conf t" & chr(13)
		
		' Courtesy timer
		crt.sleep 750	
		
		' Send the request to add the address
		crt.Screen.Send "ip access-list standard " & aclNumber & chr(13)
		
		' Courtesy timer
		crt.sleep 750

		' Send the request to add the address
		crt.Screen.Send nextSeq & " permit " & addressToAdd & chr(13)
		
		' Courtesy timer
		crt.sleep 750		
		
		' Exit the Standard ACL configuration mode
		crt.Screen.Send "exit" & chr(13)
		
		' Courtesy timer
		crt.sleep 750	
		
		' Resequence the ACL
		crt.Screen.Send "ip access-list resequence " & aclNumber & " 10 10" & chr(13)
		
		' Courtesy timer
		crt.sleep 1500	
		
		' Send the request to end global configuration mode
		crt.Screen.Send "end" & chr(13)
		
		' Courtesy timer
		crt.sleep 750

		' Mark the return as successful
		AddToList = 1
		
	else
	
		' No room for another entry
		AddToList = 0
		
	end if


End Function