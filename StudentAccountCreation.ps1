# Student-Account-Creation
# Created by: Kevin Papenhausen
# Automatically creates, disables, and deletes student accounts in AD 
Start-Transcript -Path C:\Location\Location\Domain_Script\Transcript.log 

Set-Content "C:\Location\Location\ADaccounts.csv" -Value "userprincipalname" 
Set-Content "C:\Location\Location\DeletedAccounts.csv" -Value "StudentID"

Import-Module ActiveDirectory

#Gets all Disabled accounts studentID's in DISABLED ACCOUNT OU
Get-ADUser -Filter * -server DomainName.org -SearchBase "OU=DISABLED ACCOUNTS,DC=Domain,DC=Domain,DC=org" -Properties studentID, physicaldeliveryofficename | Select studentID, physicaldeliveryofficename | Export-CSV "C:\Location\Location\DISABLED_ACCOUNTS.csv"


#Removes leading zeros when excel is being used as .csv
#(Import-Csv 'C:\Location\Location\studentID.csv') | select -Property @{n='studentID';e={[int]$_.studentID}},* -Exclude 'studentID' | Export-Csv 'C:\Location\Location\StudentID.csv'



#Import SchoolTool enrollment information.

$Path="C:\Location\Location\StudentExport.csv"
$import = Import-Csv $Path

##################Begin Checking for already created Accounts

#Pulls all users from Disabled Accounts OU
$Path5="C:\Location\Location\DISABLED_ACCOUNTS.csv"
$import5 = Import-Csv $Path5						
										
Compare-Object $import5 $import -property studentID -passthru -includeequal| Where-Object {$_.SideIndicator -eq '==' } | Select-Object "studentID"| export-csv "C:\Location\Location\StudentID_Re-Enable.csv" 

$Path6="C:\Location\Location\StudentID_Re-Enable.csv"
$import6 = Import-Csv $Path6

foreach($SID in $import6)
{
$ID= $SID.("studentID")
$OU= "OU=Re-Enabled Accounts,OU=domain.net,OU=domain_students,DC=Domain,DC=Domain,DC=org"
# Moves and Enables the accounts
Get-ADUser -server "DomainName.org" -LDAPFilter "(studentID=$ID)" | Move-ADObject -targetpath "$OU" 
Get-ADUser -server "DomainName.org" -LDAPFilter "(studentID=$ID)" | Enable-ADAccount

Write-Host "User $ID Enabled"


}
####################End of checking for disabled accounts

#Add all User accounts to a CSV file containing their UPN.
Get-ADUser -Filter * -server DomainName.org  -Properties userPrincipalName | Select userPrincipalName | Export-CSV "C:\Location\Location\ADaccounts.csv" 
Get-ADUser -Filter * -server DomainName.org -SearchBase "OU=domain.net,OU=domain_students,DC=Domain,DC=Domain,DC=org" -Properties studentID | Select studentID | Export-CSV "C:\Location\Location\studentID.csv"

#Compares SchoolTool export with AD export and adds new users to NewAccounts csv

$Path1="C:\Location\Location\studentID.csv"
$import1 = Import-Csv $Path1

Compare-Object $import1 $import -property studentID -passthru | Where-Object {$_.SideIndicator -eq '=>' } | Select-Object "studentID", "Person_FirstName", "Person_LastName", "StudentEnrollment_Building", "StudentEnrollment_SchoolYear", "StudentEnrollment_Grade", "StudentEnrollment_ClassOf" | export-csv "C:\Location\Location\NewAccounts.csv" -NoTypeInformation
						

$import2= Import-Csv "C:\Location\Location\NewAccounts.csv"					

$arrays=@()

ForEach ($Name In $import2)
{
	#Pulls needed infromation from SchoolTool export and saves to a variable
	$Classof=$Name.("StudentEnrollment_Classof").Substring(7,2)
	$CompanyClassOf=$Name.("StudentEnrollment_Classof").Substring(5,4)
	$School= $Name.("StudentEnrollment_Building")
	$StudentID= $Name.("studentID")
	
	$Grade= $Name.("StudentEnrollment_Grade")
	$FName= $Name.("Person_Firstname")
	$First= $FName.substring(0,1)
	
	$Last1= $Name.("Person_Lastname")
	$Last=$Last1.ToLower()

	$Display=$FName + " " + $Last1
	$Office=$Name.("Student_number")

	
			#Formats the username and pulls out special characters
			If ($Last | Select-String -pattern " Jr"){$Last=$Last.replace(' jr','')}
			If ($Last | Select-String -pattern " i"){$Last=$Last.substring(0,$Last.Indexof(' i'))}
			If ($Last | Select-String -pattern " v"){$Last=$Last.substring(0,$Last.Indexof(' v'))}
			If ($Last.contains(' ')){$Last=$Last.replace(' ','')}
			If ($Last.contains('(')){$Last=$Last.substring(0,$Last.Indexof(' ('))}
			If ($Last.contains('-')){$Last=$Last.replace('-','')}
			If ($Last | Select-String -pattern "'"){$Last=$Last.replace("'",'')}
			If ($Last.contains('.')){$Last=$Last.replace('.','')}
			If ($Last.contains(',')){$Last=$Last.replace(',','')}
			


			# Measures and counts the amount of characters in the last name
			$measureObject = $Last | Measure-Object -Character;
			$count = $measureObject.Characters;

			# Test to see if the counts is greater than 17, if greater it will format it to 17
			If($count -gt 17)
			{
				$Last=$Last.Substring(0,17)
			}
			else
			{
			}


			$Domain="@sscsd.net"
			$Email= $First + $Last + $Domain
			$Reset=$First + $Last

			[double]$Int=1
	
		
		
	$Path3="C:\Location\Location\ADaccounts.csv"
	$import3 = Import-Csv $Path3

	ForEach ($UName In $import3)
	{
		$Name1= $UName.("userprincipalname")

		If ($Email -ne $Name1)
		{
		}
		Else
		{
			While($Email -eq $Name1)
			{
				write-output "$Name1"
				$Email=$Reset
				$Email= $Email + [string]$Int + $Domain
				$Int=$Int+1
			}
		}

	#write-output "$Email"
	}
	
	#Test schools
	Switch ($School)
	{
		"School1"
		{
			$Descrip="CA"+$Classof
			#Write-output $Descrip
			$LogonScript="elementary.bat"
			$OU="OU=New_Accounts,OU=location.net,OU=location,DC=Domain,DC=Domain,DC=org"
			$Department="Elem"
		}
		"School2"
		{
			$Descrip="DI"+$Classof
			#Write-output $Descrip
			$LogonScript="elementary.bat"
			$OU="OU=New_Accounts,OU=location.net,OU=location,DC=Domain,DC=Domain,DC=org"
			$Department="Elem"
		}
		"School3"
		{
			$Descrip="DN"+$Classof
			#Write-output $Descrip
			$LogonScript="elementary.bat"
			$OU="OU=New_Accounts,OU=location.net,OU=location,DC=Domain,DC=Domain,DC=org"
			$Department="Elem"
		}
		"School4"
		{
			$Descrip="GE"+$Classof
			#Write-output $Descrip
			$LogonScript="elementary.bat"
			$OU="OU=New_Accounts,OU=location.net,OU=location,DC=Domain,DC=Domain,DC=org"
			$Department="Elem"
		}
		"School5"
		{
			$Descrip="GR"+$Classof
			#Write-output $Descrip
			$LogonScript="elementary.bat"
			$OU="OU=New_Accounts,OU=location.net,OU=location,DC=Domain,DC=Domain,DC=org"
			$Department="Elem"
		}
		"School6"
		{
			$Descrip="LA"+$Classof
			#Write-output $Descrip
			$LogonScript="elementary.bat"
			$OU="OU=New_Accounts,OU=location.net,OU=location,DC=Domain,DC=Domain,DC=org"
			$Department="Elem"
		}
		"School7"
		{
			$Descrip="HS"+$Classof
			#Write-output $Descrip
			$LogonScript="student.bat"
			$OU="OU=New_Accounts,OU=location.net,OU=location,DC=Domain,DC=Domain,DC=org"
			$Department="HS"
		}
		"School8"
		{
			$Descrip="MA"+$Classof
			#Write-output $Descrip
			$LogonScript="Student.bat"
			$OU="OU=New_Accounts,OU=location.net,OU=location,DC=Domain,DC=Domain,DC=org"
			$Department="MS"
		}
		default
		{
			$Descrip = "ERR"
		}
	}
	
	
	$StudentID3="{0:000000000}" -f $StudentID
	$Test="1"
	ForEach ($Uname3 In $import1)
	{
		$StudentID2= $UName3.("studentID")
		$StudentID4="{0:000000000}" -f $StudentID2
		
		If ($StudentID3 -eq $StudentID4)
		{
			$Test="0"
			write-output "$Test"
			write-output "$StudentID3"
			write-output "$StudentID4"
			write-output "$Email"
		}
		Else
		{
			
			#write-output "Created!"
		}
		
	}



	if ($test -eq "1")
	{
	
	write-output "Created!"
	write-output "$StudentID3"
	write-output "$StudentID4"
	$Email=$Email.ToLower()
	$Reset1=$Email
	$Reset1=$Reset1.substring(0,$Reset1.Indexof('@'))
	$Reset1=$Reset1.ToLower()
	[int]$StudentID= $Name.("studentID")
	$StudentID3="{0:000000000}" -f $StudentID	
	$Proxy="SMTP:"+"$Email"	
	#$Password1= "$Password" | ConvertTo-SecureString -AsPlainText -Force
	$ErrorMessage= $Reset1+ "- " + " Account has failed to Create" + " " + "("+$TimeStamp+")"
	$ErrorMessage2= "First Name= "+$FName + " Last Name= "+ $Last1 + " DisplayName= " + $Display + " samaccountname= " + $Reset1 + " userprincipalname= " + $Email + "  Not all fields exist for account creation" + "("+$TimeStamp+")"
	

	#write-output "no match"

	# Test code to check that all variables are populated for user account creation in AD.
	
	write-output "$Reset1"
   	write-output "$FName"
   	write-output "$Last1"
    	write-output "$Email"
	write-output "$Display"
	write-output "$Descrip"
	write-output "$OU"
	write-output "$LogonScript"
	

	#Creates accounts in Active Directory UC
	
	New-ADUser -server DomainName.org -Name "$Reset1" -GivenName "$FName" -SurName "$Last1" -samaccountname "$Reset1" -userprincipalname "$Email" -Displayname "$Display" -Description "$Descrip" -Email "$Email" -Title 	"Student" -Path "$OU" -Scriptpath "$LogonScript" -Department "$Department" -Company "$CompanyClassOf" -AccountPassword (ConvertTo-SecureString -AsPlainText "password" -Force) -ChangePasswordAtLogon $True -Enabled $True
	Add-AdGroupMember -server DomainName.org -Identity "SMS_Students" -Members $Reset1
	Set-ADuser "$Reset1" -server DomainName.org -Add @{studentID=$StudentID3} 
	Set-ADuser "$Reset1" -server DomainName.org -Add @{proxyAddresses=$Proxy}
	Write-output $Reset1
		
	#exports values to a year to date account creation file
		If ($? -eq "false")
		{
		$array= New-Object PSObject
		
		$array| Add-Member -MemberType NoteProperty -Name "Displayname" -Value $Display
		$array| Add-Member -MemberType NoteProperty -Name "Student_Number" -Value $StudentID
		$array| Add-Member -MemberType NoteProperty -Name "Grad_Year" -Value $Classof
		$array| Add-Member -MemberType NoteProperty -Name "Account" -Value $Email
		$array| Add-Member -MemberType NoteProperty -Name "Samaccountname" -Value $Reset1
		$array| Add-Member -MemberType NoteProperty -Name "Proxy_Address" -Value $Proxy

		
		$TimeStamp = (Get-Date).tostring("MM-dd-yyyy")
		$arrays+=$array

		$arrays|Export-Csv -NoTypeInformation -path "C:\Location\Location\NewAccounts.csv"
		$Email | Add-Content -path "C:\Location\Location\ADaccounts.csv"
 		$AllcreatedAccounts= $Display + " " + $StudentID + " " + $Classof + " " + $Reset1 + " "+"("+$TimeStamp+")"
		$AllcreatedAccounts | out-file -append -filepath "C:\Location\Location\2017-2018_Accounts_Created"
		Write-output "$Email"
		}
		Else
		{
		$ErrorMessage | out-file -append -filepath "C:\Location\Location\errorlog.txt"
		}



		}
		else
		{
		write-output "Will Not Create!"
		$ErrorMessage= $Email+ "- " + " Cannot create an account with an existing studentID in ADCU" + " " + "("+$TimeStamp+")"
		$ErrorMessage | out-file -append -filepath "C:\Location\Location\errorlog.txt"
		}
	
		
	

}
# End of user creation code / beggining of user disabling code



$Path="C:\Location\Location\StudentExport.csv"
$import = Import-Csv $Path

$Path1="C:\Location\Location\StudentID.csv"
$import1 = Import-Csv $Path1

#pulls out the studentIDs from ADUC that are no longer in SchoolTool export
Compare-Object $import1 $import -property studentID -passthru| Where-Object {$_.SideIndicator -eq '<=' } | Select-Object "studentID"| export-csv "C:\Location\Location\StudentID_Outgoing.csv" 

$Path4="C:\Location\Location\StudentID_Outgoing.csv"
$import4 = Import-Csv $Path4

foreach($SID in $import4)
{
$ID= $SID.("studentID")
$OU= "OU=DISABLED ACCOUNTS,DC=Domain,DC=Domain,DC=org"
# Moves and disables the accounts
Get-ADUser -server "DomainName.org" -LDAPFilter "(studentID=$ID)" | Move-ADObject -targetpath "$OU" 
Get-ADUser -server "DomainName.org" -LDAPFilter "(studentID=$ID)" | Disable-ADAccount
Get-ADUser -server "DomainName.org" -LDAPFilter "(studentID=$ID)" | Set-ADuser -Office "1" 

Write-Host "User $ID disabled"


}



#Sends Email of the new accounts created

$SMTPUsername = "AD_Admin_Account@Domain.org"
$EncryptedPasswordFile = "C:\Location\Location\scriptsencrypted_password1.txt"
$SecureStringPassword = Get-Content -Path $EncryptedPasswordFile | ConvertTo-SecureString
$EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SMTPUsername,$SecureStringPassword
$Attachment = "C:\Location\Location\NewAccounts.csv"
$Attachment1 = "C:\Location\Location\StudentID_Outgoing.csv"
$Attachment2 = "C:\Location\Location\StudentID_Re-Enable.csv"
$TimeStamp = (Get-Date).tostring("MM-dd-yyyy")
$Subject= "Accounts Created on " + $TimeStamp 
$body = “Here are the new created accounts for ” + $TimeStamp




###########Account Deletion 
$Path6="C:\Location\Location\DISABLED_ACCOUNTS.csv"
$import6 = Import-Csv $Path6

foreach($Value in $import6)
{
$ID= $Value.("studentID")
[Double]$Timer= $Value.("physicaldeliveryofficename")

if ($Timer -eq 90)
	{
	
	$TimeStamp1 = (Get-Date).tostring("MM-dd-yyyy")
	$Subject1= "Deleted Account- 90 Days Since Disabled " + $TimeStamp1 
	Write-output "Delete"
	Get-ADUser -server "DomainName.org" -LDAPFilter "(studentID=$ID)" -Properties samaccountname | Select $ID | export-csv "C:\Location\Location\DeletedAccounts.csv" -NoTypeInformation
	$Info = Get-ADUser -server "DomainName.org" -LDAPFilter "(studentID=$ID)" -Properties samaccountname | Select samaccountname 
	$IDandInfo = "$ID"+"$Info"
	write-output "$IDandInfo"
	Get-ADUser -server "DomainName.org" -LDAPFilter "(studentID=$ID)" | Remove-ADUser -Confirm:$false
	
	
	}
else
	{
	[Double]$Adjusted_Timer= $Timer + 1
	#write-output "$Adjusted_Timer"
	#write-output "$ID"
	Try{Get-ADUser -server "DomainName.org" -LDAPFilter "(studentID=$ID)" | Set-ADuser -Office "$Adjusted_Timer"} 
	catch{write-output "error in username format"}
	}
}
###################Account Deletion 

$Attachment3 = "C:\Location\Location\DeletedAccounts.csv"
Send-MailMessage -To IT_USER1@domain.org, IT_USER2@domain.org, IT_USER2@domain.org -Subject $Subject -Body $body -BodyAsHtml -Attachments $Attachment, $Attachment1, $Attachment2, $Attachment3 -smtpserver smtp.office365.com -usessl -Credential $EmailCredential -Port 587 


Stop-Transcript



