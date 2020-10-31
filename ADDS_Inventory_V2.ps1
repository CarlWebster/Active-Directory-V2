#Requires -Version 3.0
#requires -Module ActiveDirectory
#requires -Module GroupPolicy
#This File is in Unicode format.  Do not edit in an ASCII editor. Notepad++ UTF-8-BOM

<#
.SYNOPSIS
	Creates a complete inventory of a Microsoft Active Directory Forest.
.DESCRIPTION
	Creates a complete inventory of a Microsoft Active Directory Forest using Microsoft 
	PowerShell, Word, plain text, or HTML.
	
	Creates a Word or PDF document, text or HTML file named after the Active Directory Forest.
	
	Word and PDF document includes a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish

	The script requires at least PowerShell version 3 but runs best in version 5.

	Word is NOT needed to run the script. This script will output in Text and HTML.
	
	You do NOT have to run this script on a domain controller. This script was developed 
	and run from a Windows 10 VM.

	While most of the script can be run with a non-admin account, there are some features 
	that will not or may not work without domain admin or enterprise admin rights.  
	The Hardware and Services parameters require domain admin privileges.  
	
	Version 2.0 of the script adds gathering information on Time Server and AD database, 
	log file, and SYSVOL locations. Those require access to the registry on each domain 
	controller, which means the script should now always be run from an elevated PowerShell 
	session with an account with a minimum of domain admin rights.
	
	Running the script in a forest with multiple domains requires Enterprise Admin rights.

	The count of all users may not be accurate if the user running the script doesn't have 
	the necessary permissions on all user objects.  In that case, there may be user accounts 
	classified as "unknown".
	
	To run the script from a workstation, RSAT is required.
	
	Remote Server Administration Tools for Windows 7 with Service Pack 1 (SP1)
		http://www.microsoft.com/en-us/download/details.aspx?id=7887
		
	Remote Server Administration Tools for Windows 8 
		http://www.microsoft.com/en-us/download/details.aspx?id=28972
		
	Remote Server Administration Tools for Windows 8.1 
		http://www.microsoft.com/en-us/download/details.aspx?id=39296
		
	Remote Server Administration Tools for Windows 10
		http://www.microsoft.com/en-us/download/details.aspx?id=45520
	
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	This parameter is disabled by default.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	This parameter requires Microsoft Word to be installed.
	This parameter uses the Word SaveAs PDF capability.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	This parameter is disabled by default.
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2020 at 6PM is 2020-06-01_1800.
	Output filename will be ReportName_2020-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
.PARAMETER ADDomain
	Specifies an Active Directory domain object by providing one of the following 
	property values. The identifier in parentheses is the LDAP display name for the 
	attribute. All values are for the domainDNS object that represents the domain.

	Distinguished Name

	Example: DC=tullahoma,DC=corp,DC=labaddomain,DC=com

	GUID (objectGUID)

	Example: b9fa5fbd-4334-4a98-85f1-3a3a44069fc6

	Security Identifier (objectSid)

	Example: S-1-5-21-3643273344-1505409314-3732760578

	DNS domain name

	Example: tullahoma.corp.labaddomain.com

	NetBIOS domain name

	Example: Tullahoma

	If both ADForest and ADDomain are specified, ADDomain takes precedence.
.PARAMETER ADForest
	Specifies an Active Directory forest object by providing one of the following 
	attribute values. 
	The identifier in parentheses is the LDAP display name for the attribute.

	Fully qualified domain name
		Example: labaddomain.com
	GUID (objectGUID)
		Example: 599c3d2e-e61e-4d20-7b77-030d99495e19
	DNS host name
		Example: labaddomain.com
	NetBIOS name
		Example: labaddomain
	
	Default value is $Env:USERDNSDOMAIN	
	
	If both ADForest and ADDomain are specified, ADDomain takes precedence.
.PARAMETER CompanyAddress
	Company Address to use for the Cover Page, if the Cover Page has the Address field.
	
	The following Cover Pages have an Address field:
		Banded (Word 2013/2016)
		Contrast (Word 2010)
		Exposure (Word 2010)
		Filigree (Word 2013/2016)
		Ion (Dark) (Word 2013/2016)
		Retrospect (Word 2013/2016)
		Semaphore (Word 2013/2016)
		Tiles (Word 2010)
		ViewMaster (Word 2013/2016)
		
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CA.
.PARAMETER CompanyEmail
	Company Email to use for the Cover Page, if the Cover Page has the Email field.  
	
	The following Cover Pages have an Email field:
		Facet (Word 2013/2016)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CE.
.PARAMETER CompanyFax
	Company Fax to use for the Cover Page, if the Cover Page has the Fax field.  
	
	The following Cover Pages have a Fax field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CF.
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in 
	HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated 
	on the computer running the script.
	This parameter has an alias of CN.
	If either registry key does not exist and this parameter is not specified, the report 
	will not contain a Company Name on the cover page.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER CompanyPhone
	Company Phone to use for the Cover Page if the Cover Page has the Phone field.  
	
	The following Cover Pages have a Phone field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CPh.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013 and 2016 are supported.
	(default cover pages in Word en-US)
	
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly 
		works in 2010 but Subtitle/Subject & Author fields need to be moved 
		after title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit; box needs to be 
		manually resized or font changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 
		36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit; box needs to be manually 
		resized or font changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 
		2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)
		
	The default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER ComputerName
	Specifies which domain controller to use to run the script against.
	If ADForest is a trusted forest, then ComputerName is required to detect the 
	existence of ADForest.
	ComputerName can be entered as the NetBIOS name, FQDN, localhost or IP Address.
	If entered as localhost, the actual computer name is determined and used.
	If entered as an IP address, an attempt is made to determine and use the actual 
	computer name.
	
	This parameter has an alias of ServerName.
	Default value is $Env:USERDNSDOMAIN	
.PARAMETER DCDNSInfo 
	Use WMI to gather, for each domain controller, the IP Address, and each DNS server 
	configured.
	This parameter requires the script be run from an elevated PowerShell session 
	using an account with permission to retrieve hardware information (i.e. Domain 
	Admin).
	Selecting this parameter will add an extra section to the report.
	This parameter is disabled by default.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER GPOInheritance
	In the Group Policies by OU section, adds Inherited GPOs in addition to the GPOs 
	directly linked.
	Adds a second column to the table GPO Type.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of GPO.
.PARAMETER Hardware
	Use WMI to gather hardware information on Computer System, Disks, Processor, and 
	Network Interface Cards
	This parameter requires the script be run from an elevated PowerShell session 
	using an account with permission to retrieve hardware information (i.e. Domain 
	Admin).
	Selecting this parameter will add to both the time it takes to run the script and 
	size of the report.
	This parameter is disabled by default.
.PARAMETER IncludeUserInfo
	For the User Miscellaneous Data section outputs a table with the SamAccountName
	and DistinguishedName of the users in the All Users counts:
	
		Disabled users
		Unknown users
		Locked out users
		All users with password expired
		All users whose password never expires
		All users with password not required
		All users who cannot change password
		All users with SID History
		All users with Homedrive set in ADUC
		All users whose Primary Group is not Domain Users
		All users with RDS HomeDrive set in ADUC
	
	The Text output option is limited to the first 25 characters of the SamAccountName
	and the first 116 characters of the DistinguishedName.
	
	This parameter is disabled by default.
	This parameter has an alias of IU.
.PARAMETER Log
	Generates a log file for troubleshooting.
.PARAMETER MaxDetails
	Adds maximum detail to the report.
	
	This is the same as using the following parameters:
		DCDNSInfo
		GPOInheritance
		Hardware
		IncludeUserInfo
		Services
	
	WARNING: Using this parameter can create an extremely large report and 
	can take a very long time to run.

	This parameter has an alias of MAX.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER Services
	Gather information on all services running on domain controllers.
	Servers that are configured to automatically start but are not running will be 
	colored in red.
	Used on Domain Controllers only.
	This parameter requires the script be run from an elevated PowerShell session
	using an account with permission to retrieve service information (i.e. Domain 
	Admin).
	Selecting this parameter will add to both the time it takes to run the script and 
	size of the report.
	This parameter is disabled by default.
.PARAMETER Section
	Processes one or more sections of the report.
	Valid options are:
		Forest
		Sites
		Domains (includes Domain Controllers and optional Hardware, Services and 
		DCDNSInfo)
		OUs (Organizational Units)
		Groups
		GPOs
		Misc (Miscellaneous data)
		All

	This parameter defaults to All sections.
	
	Multiple sections are separated by a comma. -Section forest, domains
	
.PARAMETER UserName
	Username to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	The default is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	The default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -ADForest company.tld
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	company.tld for the AD Forest.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -ADDomain child.company.tld
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	child.company.tld for the AD Domain.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -ADForest parent.company.tld 
	-ADDomain child.company.tld
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Because both ADForest and ADDomain are specified, ADDomain wins and child.company.tld 
	is used for AD Domain.
	ADForest is set to the value of ADDomain.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -ADForest company.tld -ComputerName DC01
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	company.tld for the AD Forest
	The script will be run remotely on the DC01 domain controller.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -PDF -ADForest corp.carlwebster.com
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	corp.carlwebster.com for the AD Forest.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -Text -ADForest corp.carlwebster.com
	
	Will use all default values and save the document as a formatted text file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator.

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	corp.carlwebster.com for the AD Forest.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -HTML -ADForest corp.carlwebster.com
	
	Will use all default values and save the document as an HTML file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator.

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	corp.carlwebster.com for the AD Forest.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -hardware
	
	Will use all default values and add additional information for each domain controller 
	about its hardware.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator.

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -services
	
	Will use all default values and add additional information for the services running 
	on each domain controller.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator.

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -DCDNSInfo
	
	Will use all default values and add additional information for each domain controller 
	about its DNS IP configuration.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	An extra section will be added to the end of the report.

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
.EXAMPLE
	PS C:\PSScript .\ADDS_Inventory_V2.ps1 -CompanyName "Carl Webster Consulting" 
	-CoverPage "Mod" -UserName "Carl Webster" -ComputerName ADDC01

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
		ADForest defaults to the value of $Env:USERDNSDOMAIN.
		Domain Controller named ADDC01 for the ComputerName.
.EXAMPLE
	PS C:\PSScript .\ADDS_Inventory_V2.ps1 -CN "Carl Webster Consulting" -CP "Mod" 
	-UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
.EXAMPLE
	PS C:\PSScript .\ADDS_Inventory_V2.ps1 -CompanyName "Sherlock Holmes 
	Consulting"
	-CoverPage Exposure -UserName "Dr. Watson"
	-CompanyAddress "221B Baker Street, London, England"
	-CompanyFax "+44 1753 276600"
	-CompanyPhone "+44 1753 276200"
	
	Will use:
		Sherlock Holmes Consulting for the Company Name.
		Exposure for the Cover Page format.
		Dr. Watson for the User Name.
		221B Baker Street, London, England for the Company Address.
		+44 1753 276600 for the Company Fax.
		+44 1753 276200 for the Company Phone.

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
.EXAMPLE
	PS C:\PSScript .\ADDS_Inventory_V2.ps1 -CompanyName "Sherlock Holmes 
	Consulting"
	-CoverPage Facet -UserName "Dr. Watson"
	-CompanyEmail SuperSleuth@SherlockHolmes.com

	Will use:
		Sherlock Holmes Consulting for the Company Name.
		Facet for the Cover Page format.
		Dr. Watson for the User Name.
		SuperSleuth@SherlockHolmes.com for the Company Email.

	ADForest defaults to the value of $Env:USERDNSDOMAIN.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -ADForest company.tld -AddDateTime
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator.

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	company.tld for the AD Forest.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.

	Adds a date time stamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2020 at 6PM is 2020-06-01_1800.
	Output filename will be company.tld_2020-06-01_1800.docx.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -PDF -ADForest corp.carlwebster.com 
	-AddDateTime
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	corp.carlwebster.com for the AD Forest.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.

	Adds a date time stamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2020 at 6PM is 2020-06-01_1800.
	Output filename will be corp.carlwebster.com_2020-06-01_1800.PDF
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -ADForest corp.carlwebster.com 
	-Folder \\FileServer\ShareName
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	corp.carlwebster.com for the AD Forest.

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
	
	The output file will be saved in the path \\FileServer\ShareName.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -Section Forest

	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	ADForest defaults to the value of $Env:USERDNSDOMAIN

	ComputerName defaults to the value of $Env:USERDNSDOMAIN, then the script queries for 
	a domain controller that is also a global catalog server and will use that as the 
	value for ComputerName.
	
	The report will include only the Forest section.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -Section groups, misc 
	-ADForest WebstersLab.com -ServerName PrimaryDC.websterslab.com

	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	WebstersLab.com for ADForest.
	PrimaryDC.websterslab.com for ComputerName.
	
	The report will include only the Groups and Miscellaneous sections.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -MaxDetails
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Set the following parameter values:
		DCDNSInfo       = True
		GPOInheritance  = True
		Hardware        = True
		IncludeUserInfo = True
		Services        = True
		
		Section         = "All"
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 -ADForest corp.carlwebster.com 
	-SmtpServer mail.domain.tld
	-From XDAdmin@domain.tld 
	-To ITGroup@domain.tld	
	-ComputerName Server01

	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	corp.carlwebster.com for the AD Forest.
	
	The script will be run remotely against server Server01.
	The script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, 
	sending to ITGroup@domain.tld.

	The script will use the default SMTP port 25 and will not use SSL.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 
	-SmtpServer mailrelay.domain.tld
	-From Anonymous@domain.tld 
	-To ITGroup@domain.tld	

	***SENDING UNAUTHENTICATED EMAIL***

	The script will use the email server mailrelay.domain.tld, sending from 
	anonymous@domain.tld, sending to ITGroup@domain.tld.

	To send unauthenticated email using an email relay server requires the From email account 
	to use the name Anonymous.

	The script will use the default SMTP port 25 and will not use SSL.
	
	***GMAIL/G SUITE SMTP RELAY***
	https://support.google.com/a/answer/2956491?hl=en
	https://support.google.com/a/answer/176600?hl=en

	To send email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	***GMAIL/G SUITE SMTP RELAY***

	The script will generate an anonymous secure password for the anonymous@domain.tld 
	account.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 
	-SmtpServer labaddomain-com.mail.protection.outlook.com
	-UseSSL
	-From SomeEmailAddress@labaddomain.com 
	-To ITGroupDL@labaddomain.com	

	***OFFICE 365 Example***

	https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-office-3
	
	This uses Option 2 from the above link.
	
	***OFFICE 365 Example***

	The script will use the email server labaddomain-com.mail.protection.outlook.com, 
	sending from SomeEmailAddress@labaddomain.com, sending to ITGroupDL@labaddomain.com.

	The script will use the default SMTP port 25 and will use SSL.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 
	-SmtpServer smtp.office365.com 
	-SmtpPort 587
	-UseSSL 
	-From Webster@CarlWebster.com 
	-To ITGroup@CarlWebster.com	

	The script will use the email server smtp.office365.com on port 587 using SSL, 
	sending from webster@carlwebster.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\ADDS_Inventory_V2.ps1 
	-SmtpServer smtp.gmail.com 
	-SmtpPort 587
	-UseSSL 
	-From Webster@CarlWebster.com 
	-To ITGroup@CarlWebster.com	

	*** NOTE ***
	To send email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	*** NOTE ***
	
	The script will use the email server smtp.gmail.com on port 587 using SSL, 
	sending from webster@gmail.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word or PDF document.
.NOTES
	NAME: ADDS_Inventory_V2.ps1
	VERSION: 2.26
	AUTHOR: Carl Webster and Michael B. Smith
	LASTEDIT: May 8, 2020
#>


#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Word") ]

Param(
	[parameter(ParameterSetName="HTML",Mandatory=$False)] 
	[Switch]$HTML=$False,

	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName="Text",Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$AddDateTime=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$ADDomain="", 

	[parameter(Mandatory=$False)] 
	[string]$ADForest=$Env:USERDNSDOMAIN, 

	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CA")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyAddress="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CE")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyEmail="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CF")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyFax="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CPh")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyPhone="",
    
	[parameter(Mandatory=$False)] 
	[Alias("ServerName")]
	[string]$ComputerName=$Env:USERDNSDOMAIN,
	
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(Mandatory=$False)] 
	[Switch]$DCDNSInfo=$False, 

	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(Mandatory=$False)] 
	[Switch]$GPOInheritance=$False, 

	[parameter(Mandatory=$False)] 
	[Switch]$Hardware=$False, 

	[parameter(Mandatory=$False)] 
	[Alias("IU")]
	[Switch]$IncludeUserInfo=$False,
	
	[parameter(Mandatory=$False)] 
	[Switch]$Log=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("MAX")]
	[Switch]$MaxDetails=$False,

	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False,
	
	[parameter(Mandatory=$False)] 
	[array]$Section="All",
	
	[parameter(Mandatory=$False )] 
	[Switch]$Services=$False,
	
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(Mandatory=$False)] 
	[string]$SmtpServer="",

	[parameter(Mandatory=$False)] 
	[int]$SmtpPort=25,

	[parameter(Mandatory=$False)] 
	[switch]$UseSSL=$False,

	[parameter(Mandatory=$False)] 
	[string]$From="",

	[parameter(Mandatory=$False)] 
	[string]$To=""

	)

	
#webster@carlwebster.com
#Created by Carl Webster and Michael B. Smith
#webster@carlwebster.com
#@carlwebster on Twitter
#https://www.CarlWebster.com
#
#michael@smithcons.com
#@essentialexch on Twitter
#https://www.essential.exchange/blog/
#
#Created on April 10, 2014

#Version 1.0 released to the community on May 31, 2014
#
#Version 2.0 is based on version 1.20
#
#Version 2.26 8-May-2020
#	Add checking for a Word version of 0, which indicates the Office installation needs repairing
#	Add Receive Side Scaling setting to Function OutputNICItem
#	Change color variables $wdColorGray15 and $wdColorGray05 from [long] to [int]
#	Change location of the -Dev, -Log, and -ScriptInfo output files from the script folder to the -Folder location (Thanks to Guy Leech for the "suggestion")
#	Change Text output to use [System.Text.StringBuilder]
#		Updated Functions Line and SaveAndCloseTextDocument
#	Reformatted the terminating Write-Error messages to make them more visible and readable in the console
#	Update Function SetWordCellFormat to change parameter $BackgroundColor to [int]
#	Update Functions GetComputerWMIInfo and OutputNicInfo to fix two bugs in NIC Power Management settings
#
#Version 2.25 21-Apr-2020
#	Remove the SMTP parameterset and manually verify the parameters
#	Reorder parameters
#	Update Function SendEmail to handle anonymous unauthenticated email
#		Update Help Text with examples
#
#Version 2.24 13-Feb-2020
#	Fixed several variable name typos
#	General code cleanup
#	Updated the following Exchange Schema Versions:
#		"15312" = "Exchange 2013 CU7 through CU23"
#		"15317" = "Exchange 2016 Preview and RTM"
#		"15332" = "Exchange 2016 CU7 through CU15"
#		"17000" = "Exchange 2019 RTM/CU1"
#		"17001" = "Exchange 2019 CU2-CU4"
#
#Version 2.23 17-Dec-2019
#	Fix Swedish Table of Contents (Thanks to Johan Kallio)
#		From 
#			'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
#		To
#			'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
#	Updated help text
#
#Version 2.22 14-Feb-2019
#	Added a line under the OU table stating how many OUs are not protected
#	Added color $wdColorYellow
#	Added Exchange schema version 17000 for Exchange 2019
#	Added to the "Gathering user misc data" section, the following console message if there are more than 100,000 user accounts in AD:
#		There are $($UsersCount) user accounts to process. The following 17 actions will take a long time. Be patient.
#	Changed section heading "Domain trusts" to "Domain Trusts" to match the capitalization of other sections
#	Changed several $Var -eq $Null to $Null -eq $Var and on Get-Process line for WinWord (thanks to MBS)
#	Changed test for "No Certification Authority Root(s) were retrieved" by Michael B. Smith who contributed the original code
#	For HTML and Text output, for Heading1 and Heading2 output, added "///  " and "  \\\" surrounding the heading text
#		This will help for those of us who read reports that contain > 100,000 OUs and users
#		and > 1,000 GPOs
#	Removed "Preview" from Windows Server 2019 AD Schema version 88
#	Remove unused variables
#	Updated help text
#
#Version 2.21 11-Nov-2018
#	For HTML output, reverted the output Hardware and Service functions back to using $rowdata = @()
#		Using $rowdata = New-Object System.Collections.ArrayList did not always work, which is weird
#
#Version 2.20 28-Sep-2018
#	Added Domain Functional Level of 7 for Windows Server 2016
#	Added Forest Functional Level of 7 for Windows Server 2016
#	Added Domain Schema version 88 for Server 2019 Preview
#	Added to Domain Information:
#		Last logon replication interval
#		Public key required password rolling (2012+)
#	Added to Forest Information, Domain Controllers:
#		Operation System
#		Server Core (Y/N)
#	Changed "renamed" to "changed" as it was freaking people out thinking I was renaming their domain or computer
#	Changed all but the Word and HTML arrays from "@() +=" to "New-Object System.Collections.ArrayList .Add()"
#	Changed the code where I checked for Singletons and -is [array] to use @() around the cmdlets so the result
#		is always an array. Thanks to fellow CTP Sam Jacobs for the tip. This reduced the code by almost 500 lines.
#	Changed the functions getting the computer WMI hardware and service info to use 
#		"New-Object System.Collections.ArrayList .Add()" for Word and HTML tables
#	Changed the width of the Domain Controllers table to accommodate the new columns
#	Change the width of the AD Schema Items table to match the other tables
#	Remove all the duplicate $VarName = $Null from Function ProcessDomains
#	Reorder Change Log so the most recent is on top and the oldest is at the bottom
#	Reorder most of the domain properties to be in alphabetical order
#	Reorder most of the forest properties to be in alphabetical order
#	Updated Exchange schema version information
#
#Version 2.19 5-Apr-2018
#	Added Event Log information to each domain controller and an appendix
#		If the script is run from an elevated PowerShell session by a user with Domain Admin rights
#	Added Operating System information to Functions GetComputerWMIInfo and OutputComputerItem
#	Code clean-up for most recommendations made by Visual Studio Code
#
#Version 2.18 10-Mar-2018
#	Added Log switch to create a transcript log
#
#Version 2.17 8-Dec-2017
#	Updated Function WriteHTMLLine with fixes from the script template
#
#Version 2.16 4-Dec-2017
#	Add checking for users with home drive set in Active Directory Users and Computers (ADUC)
#		Added function OutputHDUserInfo
#	Add checking for users with RDS home drive set in ADUC
#		Added function from Jeff Hicks Get-RDUserSetting
#		Added function OutputRDSHDUserInfo
#	Add checking for users whose Primary Group is not Domain Users
#		Added function OutputPGUserInfo
#	Add "DC: " in fron tof the domain controller name, in text output, for domain controller information
#	Add new parameter ADDomain to restrict report to a single domain in a multi-domain Forest
#	Add schema extension checking for the following items and add to Forest section:
#		'User-Account-Control', #Flags that control the behavior of a user account
#		'msNPAllowDialin', #RAS Server
#		'ms-Mcs-AdmPwd', #LAPS
#		'ms-Mcs-AdmPwdExpirationTime', #LAPS
#		'ms-SMS-Assignment-Site-Code', #SCCM
#		'ms-SMS-Capabilities', #SCCM
#		'msRTCSIP-UserRoutingGroupId', #Lync/SfB
#		'msRTCSIP-MirrorBackEndServer' #Lync/SfB
#		'ms-exch-schema-version-pt' #Exchange
#	Add "Site: " in front of Site name when listing Subnets, Servers, and Connection Objects
#	Remove several large blocks of code that had been commented out
#	Revise how $LinkedGPOs and $InheritedGPOs variables are set to work around invalid property 
#		name DisplayName when collection is empty
#	Sort Enabled Scopes in AD Optional Features
#	Text output changes to tabular data:
#		Domain Controllers (in Forest section)
#		AD Schema Items (in Forest section)
#		Services
#		Organizational Units
#		Domain Admins
#		Enterprise Admins
#		Schema Admins
#		Users with AdminCount=1
#	Updated Exchange schema versions
#	Updated help text
#	When reporting on the domain controllers in the Forest, if unable to get data from a domain controller,
#		instead of reporting "Unknown", use:
#		Unable to retrieve Global Catalog status on <DCName>
#		Unable to retrieve Read-only status on <DCName>
#	When run for a single domain in a multi-domain forest
#		Revise gathering list of domains
#		Revise testing for $ComputerName 
#		Revise variable $ADContext in Function ProcessAllDCsInTheForest
#
#Version 2.15 26-Jun-2017
#	Added new parameter MaxDetails:
#		This is the same as using the following parameters:
#			DCDNSInfo
#			GPOInheritance
#			HardWare
#			IncludeUserInfo
#			Services
#	Fixed wrong loop variable for CA
#	Removed code that made sure all Parameters were set to default values if for some reason they did exist or values were $Null
#	Reordered the parameters in the help text and parameter list so they match and are grouped better
#	Replaced _SetDocumentProperty function with Jim Moyle's Set-DocumentProperty function
#	Updated Function ProcessScriptEnd for the new Cover Page properties and Parameters
#	Updated Function ShowScriptOptions for the new Cover Page properties and Parameters
#	Updated Function UpdateDocumentProperties for the new Cover Page properties and Parameters
#	Updated help text
#
#Version 2.14 12-May-2017
#	Add Certificate Authority Information section to Forest Information
#		Check for the following CA related errors:
#			Possible error: There are more than one Certification Authority Root(s)
#			Error: Certification Authority Root(s) exist, but no Certification Authority Issuers(s) (also known as Enrollment Agents) exist
#			Error: More Certification Authority Root(s) exist than there are Certification Authority Issuers(s) (also known as Enrollment Agents)
#			Error: Certification Authority Issuers(s) (also known as Enrollment Agents) exist, but no Certification Authority Root(s) exist
#	Change "Users with AdminCount=1 ($($AdminsCountStr) members):" to "Users with AdminCount=1 ($($AdminsCountStr) users):"
#	Reorder the Forest Information section
#
#Version 2.13 13-Feb-2017
#	Fix French wording for Table of Contents 2
#
#Version 2.12 10-Nov-2016
#	Add Chinese language support
#	Add table for Time Server information if script is run from an elevated PowerShell session
#	Remove "Appendix A" from DC DNS Info table
#
#Version 2.11 6-Nov-2016
#	Fixed Domain Trust Attributes (thanks GT)
#	Fixed several Write-Warning statements that had no message
#	Fixed using -AddDateTime with -HTML
#	Remove duplicate setting for $Script:Title
#	Reworked the use of [gc]::Collect()
#
#Version 2.10 released on 19-Oct-2016 (Happy Birthday Linz)
#	Add a new parameter, IncludeUserInfo
#		Add to the User Miscellaneous Data section, outputs a table with the SamAccountName
#		and DistinguishedName of the users in the All Users counts
#	Add to the Domain section, listing Fine Grained Password Policies
#	Add to the Forest section, Tombstone Lifetime
#	Changed the HTML header for AD Optional Features from a table header to a section header
#	Changed "Site and Services" heading to "Sites and Services"
#	Fixed formatting issues with HTML headings output
#	The $AdminsCountStr variable was used, when it should not have been used, 
#		when privileged groups had no members or members could not be retrieved
#	Update Forest and Domain schema tables for the released Server 2016 product
#
#Version 2.0 released 26-Sep-2016
#
#	Added a parameter, GPOInheritance, to set whether to use the new GPOs by OU with linked and inherited GPOs
#		By default, the script will use the original GPOs by OU with only directly linked GPOs
#	Added a function, ElevatedSession, to test if the PowerShell session running the script is elevated
#	Added a Section parameter to allow specific sections only to be in the report
#		Valid options are:
#			Forest
#			Sites
#			Domains (includes Domain Controllers and optional Hardware, Services and DCDNSInfo)
#			OUs (Organizational Units)
#			Groups
#			GPOs
#			Misc (Miscellaneous data)
#			All (Default value)
#	Added AD Database, logfile and SYSVOL locations along with AD Database size
#	Added AD Optional Features
#	Added an alias to the ComputerName parameter to ServerName
#	Added checking the NIC's "Allow the computer to turn off this device to save power" setting
#	Added requires line for the GroupPolicy module
#	Added Text and HTML output
#	Added Time Server information
#	Change checking for both DA rights and an elevated session for the Time Server and AD file locations
#		If the check fails, added a warning message as write-host with white foreground
#	Change object created for DCDNSINFO to storing blank data for DNS properties
#		HTML output would not display a row if any of the DNS values were blank or Null
#	Fix test for domain admin rights for the user account
#	Fix text and HTML output for the -Hardware parameter
#	Fix the DC DNS Info table to handle two IP Addresses
#	Fix the ProcessScriptSetup function
#		Add checking for an elevated PowerShell session
#		Add checking for DA rights and elevated session if using DCDNSINFO parameter
#		Add checking for elevated session if using the Hardware and Services parameters
#		Change the elevated session warning to write-host with a white foreground to make it stand out
#		Fix where variables were not being set properly
#	Fix the user name not being displayed in the warning message about not having domain admin rights
#	If no ComputerName value is entered and $ComputerName –eq $Env:USERDNSDOMAIN then the script queries for 
#		a domain controller that is also a global catalog server and will use that as the value for ComputerName
#	Modified GPOs by OU to show if the GPO is Linked or Inherited
#		This necessitated a change in the Word/PDF/HTML table format
#	Modified GPOs by OU to use the Get-GPInheritance cmdlet to list all directly linked and inherited GPOs
#	Organize script into functions and regions
#	Replace Jeremy Saunder's Get-ComputerCountByOS function with his latest version
#	The ADForest parameter is no longer mandatory. It will now default to the value in $Env:USERDNSDOMAIN
#	The ComputerName parameter will also now default to the value in $Env:USERDNSDOMAIN
#	Update forest and domain schema information for the latest updates for Exchange 2013/2016 and Server 2016 TP4/5
#	Update help text
#	Update Verbose messages for testing to see if -ComputerName is a domain controller
#	Worked around Get-ADDomainController issue when run from a child domain
#


Set-StrictMode -Version 2

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

If($Null -eq $MSWord)
{
	If($Text -or $HTML -or $PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$MSWord = $True
}

Write-Verbose "$(Get-Date): Testing output parameters"

If($MSWord)
{
	Write-Verbose "$(Get-Date): MSWord is set"
}
ElseIf($PDF)
{
	Write-Verbose "$(Get-Date): PDF is set"
}
ElseIf($Text)
{
	Write-Verbose "$(Get-Date): Text is set"
}
ElseIf($HTML)
{
	Write-Verbose "$(Get-Date): HTML is set"
}
Else
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Verbose "$(Get-Date): Unable to determine output parameter"
	If($Null -eq $MSWord)
	{
		Write-Verbose "$(Get-Date): MSWord is Null"
	}
	ElseIf($Null -eq $PDF)
	{
		Write-Verbose "$(Get-Date): PDF is Null"
	}
	ElseIf($Null -eq $Text)
	{
		Write-Verbose "$(Get-Date): Text is Null"
	}
	ElseIf($Null -eq $HTML)
	{
		Write-Verbose "$(Get-Date): HTML is Null"
	}
	Else
	{
		Write-Verbose "$(Get-Date): MSWord is $($MSWord)"
		Write-Verbose "$(Get-Date): PDF is $($PDF)"
		Write-Verbose "$(Get-Date): Text is $($Text)"
		Write-Verbose "$(Get-Date): HTML is $($HTML)"
	}
	Write-Error "
	`n`n
	`t`t
	Unable to determine output parameter.
	`n`n
	`t`t
	Script cannot continue.
	`n`n
	"
	Exit
}

If($ADForest -ne "" -and $ADDomain -ne "")
{
	#2.16
	#Make ADForest equal to ADDomain so no code has to change in the script
	$ADForest = $ADDomain
}

#If the MaxDetails parameter is used, set a bunch of stuff true
If($MaxDetails)
{
	$DCDNSInfo       	= $True
	$GPOInheritance  	= $True
	$HardWare        	= $True
	$IncludeUserInfo	= $True
	$Services        	= $True
	$Section			= "All"
}

$ValidSection = $False
Switch ($Section)
{
	"Forest"	{$ValidSection = $True}
	"Sites"		{$ValidSection = $True}
	"Domains"	{$ValidSection = $True}
	"OUs"		{$ValidSection = $True}
	"Groups"	{$ValidSection = $True}
	"GPOs"		{$ValidSection = $True}
	"Misc"		{$ValidSection = $True}
	"All"		{$ValidSection = $True}
}

If($ValidSection -eq $False)
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Error -Message "
	`n`n
	`t`t
	The Section parameter specified, $($Section), is an invalid Section option.
	`n`n
	`t`t
	Valid options are:
	
	`t`tForest
	`t`tSites
	`t`tDomains
	`t`tOUs
	`t`tGroups
	`t`tGPOs
	`t`tMisc
	`t`tAll
	
	`t`t
	Script cannot continue.
	`n`n
	"
	Exit
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date): Folder path $Folder exists and is a folder"
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "
			`n`n
			`t`t
			Folder $Folder is a file, not a folder.
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "
		`n`n
		`t`t
		Folder $Folder does not exist.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		Exit
	}
}

If($Folder -eq "")
{
	$Script:pwdpath = $pwd.Path
}
Else
{
	$Script:pwdpath = $Folder
}

If($Script:pwdpath.EndsWith("\"))
{
	#remove the trailing \
	$Script:pwdpath = $Script:pwdpath.SubString(0, ($Script:pwdpath.Length - 1))
}


#V2.18 added
If($Log) 
{
	#start transcript logging
	$Script:LogPath = "$($Script:pwdpath)\ADDSDocScriptTranscript_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	
	try 
	{
		Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
		Write-Verbose "$(Get-Date): Transcript/log started at $Script:LogPath"
		$Script:StartLog = $true
	} 
	catch 
	{
		Write-Verbose "$(Get-Date): Transcript/log failed at $Script:LogPath"
		$Script:StartLog = $false
	}
}

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$($Script:pwdpath)\ADInventoryScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}


If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer but did not include a From or To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a To email address but did not include a From email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($SmtpServer) -and [String]::IsNullOrEmpty($To) -and ![String]::IsNullOrEmpty($From))
{
	Write-Error "
	`n`n
	`t`t
	You specified an SmtpServer and a From email address but did not include a To email address.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($From) -and ![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified From and To email addresses but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($From) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a From email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}
If(![String]::IsNullOrEmpty($To) -and [String]::IsNullOrEmpty($SmtpServer))
{
	Write-Error "
	`n`n
	`t`t
	You specified a To email address but did not include the SmtpServer.
	`n`n
	`t`t
	Script cannot continue.
	`n`n"
	Exit
}

#region initialize variables for word html and text
[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$Script:CoName = $CompanyName
	Write-Verbose "$(Get-Date): CoName is $($Script:CoName)"
	
	#the following values were attained from 
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[int]$wdColorGray15 = 14277081
	[int]$wdColorGray05 = 15987699 
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[long]$wdColorRed = 255
	[int]$wdColorBlack = 0
	[long]$wdColorYellow = 65535 #added in ADDS script V2.22
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdWord2016 = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	[int]$wdAlignParagraphLeft = 0
	[int]$wdAlignParagraphCenter = 1
	[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	[int]$wdCellAlignVerticalTop = 0
	[int]$wdCellAlignVerticalCenter = 1
	[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	[int]$wdAdjustFirstColumn = 2
	[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	[int]$Indent1TabStops = 1 * $PointsPerTabStop
	[int]$Indent2TabStops = 2 * $PointsPerTabStop
	[int]$Indent3TabStops = 3 * $PointsPerTabStop
	[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	[int]$wdHeadingFormatFalse = 0 
}

If($HTML)
{
    Set-Variable htmlredmask         -Option AllScope -Value "#FF0000" 4>$Null
    Set-Variable htmlcyanmask        -Option AllScope -Value "#00FFFF" 4>$Null
    Set-Variable htmlbluemask        -Option AllScope -Value "#0000FF" 4>$Null
    Set-Variable htmldarkbluemask    -Option AllScope -Value "#0000A0" 4>$Null
    Set-Variable htmllightbluemask   -Option AllScope -Value "#ADD8E6" 4>$Null
    Set-Variable htmlpurplemask      -Option AllScope -Value "#800080" 4>$Null
    Set-Variable htmlyellowmask      -Option AllScope -Value "#FFFF00" 4>$Null
    Set-Variable htmllimemask        -Option AllScope -Value "#00FF00" 4>$Null
    Set-Variable htmlmagentamask     -Option AllScope -Value "#FF00FF" 4>$Null
    Set-Variable htmlwhitemask       -Option AllScope -Value "#FFFFFF" 4>$Null
    Set-Variable htmlsilvermask      -Option AllScope -Value "#C0C0C0" 4>$Null
    Set-Variable htmlgraymask        -Option AllScope -Value "#808080" 4>$Null
    Set-Variable htmlblackmask       -Option AllScope -Value "#000000" 4>$Null
    Set-Variable htmlorangemask      -Option AllScope -Value "#FFA500" 4>$Null
    Set-Variable htmlmaroonmask      -Option AllScope -Value "#800000" 4>$Null
    Set-Variable htmlgreenmask       -Option AllScope -Value "#008000" 4>$Null
    Set-Variable htmlolivemask       -Option AllScope -Value "#808000" 4>$Null

    Set-Variable htmlbold        -Option AllScope -Value 1 4>$Null
    Set-Variable htmlitalics     -Option AllScope -Value 2 4>$Null
    Set-Variable htmlred         -Option AllScope -Value 4 4>$Null
    Set-Variable htmlcyan        -Option AllScope -Value 8 4>$Null
    Set-Variable htmlblue        -Option AllScope -Value 16 4>$Null
    Set-Variable htmldarkblue    -Option AllScope -Value 32 4>$Null
    Set-Variable htmllightblue   -Option AllScope -Value 64 4>$Null
    Set-Variable htmlpurple      -Option AllScope -Value 128 4>$Null
    Set-Variable htmlyellow      -Option AllScope -Value 256 4>$Null
    Set-Variable htmllime        -Option AllScope -Value 512 4>$Null
    Set-Variable htmlmagenta     -Option AllScope -Value 1024 4>$Null
    Set-Variable htmlwhite       -Option AllScope -Value 2048 4>$Null
    Set-Variable htmlsilver      -Option AllScope -Value 4096 4>$Null
    Set-Variable htmlgray        -Option AllScope -Value 8192 4>$Null
    Set-Variable htmlolive       -Option AllScope -Value 16384 4>$Null
    Set-Variable htmlorange      -Option AllScope -Value 32768 4>$Null
    Set-Variable htmlmaroon      -Option AllScope -Value 65536 4>$Null
    Set-Variable htmlgreen       -Option AllScope -Value 131072 4>$Null
    Set-Variable htmlblack       -Option AllScope -Value 262144 4>$Null
}

If($TEXT)
{
	[System.Text.StringBuilder] $global:Output = New-Object System.Text.StringBuilder( 16384 )
}
#endregion

#region email function
Function SendEmail
{
	Param([string]$Attachments)
	Write-Verbose "$(Get-Date): Prepare to email"

	$emailAttachment = $Attachments
	$emailSubject = $Script:Title
	$emailBody = @"
Hello, <br />
<br />
$Script:Title is attached.

"@ 

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
	}

	$error.Clear()
	
	If($From -Like "anonymous@*")
	{
		#https://serverfault.com/questions/543052/sending-unauthenticated-mail-through-ms-exchange-with-powershell-windows-server
		$anonUsername = "anonymous"
		$anonPassword = ConvertTo-SecureString -String "anonymous" -AsPlainText -Force
		$anonCredentials = New-Object System.Management.Automation.PSCredential($anonUsername,$anonPassword)

		If($UseSSL)
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL -credential $anonCredentials *>$Null 
		}
		Else
		{
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-credential $anonCredentials *>$Null 
		}
		
		If($?)
		{
			Write-Verbose "$(Get-Date): Email successfully sent using anonymous credentials"
		}
		ElseIf(!$?)
		{
			$e = $error[0]

			Write-Verbose "$(Get-Date): Email was not sent:"
			Write-Warning "$(Get-Date): Exception: $e.Exception" 
		}
	}
	Else
	{
		If($UseSSL)
		{
			Write-Verbose "$(Get-Date): Trying to send email using current user's credentials with SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL *>$Null
		}
		Else
		{
			Write-Verbose  "$(Get-Date): Trying to send email using current user's credentials without SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To *>$Null
		}

		If(!$?)
		{
			$e = $error[0]
			
			#error 5.7.57 is O365 and error 5.7.0 is gmail
			If($null -ne $e.Exception -and $e.Exception.ToString().Contains("5.7"))
			{
				#The server response was: 5.7.xx SMTP; Client was not authenticated to send anonymous mail during MAIL FROM
				Write-Verbose "$(Get-Date): Current user's credentials failed. Ask for usable credentials."

				If($Dev)
				{
					Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
				}

				$error.Clear()

				$emailCredentials = Get-Credential -UserName $From -Message "Enter the password to send email"

				If($UseSSL)
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-UseSSL -credential $emailCredentials *>$Null 
				}
				Else
				{
					Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
					-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
					-credential $emailCredentials *>$Null 
				}

				If($?)
				{
					Write-Verbose "$(Get-Date): Email successfully sent using new credentials"
				}
				ElseIf(!$?)
				{
					$e = $error[0]

					Write-Verbose "$(Get-Date): Email was not sent:"
					Write-Warning "$(Get-Date): Exception: $e.Exception" 
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date): Email was not sent:"
				Write-Warning "$(Get-Date): Exception: $e.Exception" 
			}
		}
	}
}
#endregion

#region code for -hardware switch
Function GetComputerWMIInfo
{
	Param([string]$RemoteComputerName)
	
	# original work by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com
	# modified 1-May-2014 to work in trusted AD Forests and using different domain admin credentials	
	# modified 17-Aug-2016 to fix a few issues with Text and HTML output
	# modified 29-Apr-2018 to change from Arrays to New-Object System.Collections.ArrayList

	#Get Computer info
	Write-Verbose "$(Get-Date): `t`tProcessing WMI Computer information"
	Write-Verbose "$(Get-Date): `t`t`tHardware information"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteWordLine 4 0 "General Computer"
	}
	ElseIf($Text)
	{
		Line 0 "Computer Information: $($RemoteComputerName)"
		Line 1 "General Computer"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteHTMLLine 4 0 "General Computer"
	}
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_computersystem
	}
	
	Catch
	{
		$Results = $Null
	}
	
	If($? -and $Null -ne $Results)
	{
		$ComputerItems = $Results | Select-Object Manufacturer, Model, Domain, `
		@{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}, `
		NumberOfProcessors, NumberOfLogicalProcessors
		$Results = $Null
		[string]$ComputerOS = (Get-WmiObject -class Win32_OperatingSystem -computername $RemoteComputerName -EA 0).Caption

		ForEach($Item in $ComputerItems)
		{
			OutputComputerItem $Item $ComputerOS
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
			Line 2 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Computer information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Computer information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
	}
	
	#Get Disk info
	Write-Verbose "$(Get-Date): `t`t`tDrive information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Drive(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Drive(s)"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Drive(s)"
	}

	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName Win32_LogicalDisk
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$drives = $Results | Select-Object caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
		filesystem, @{N="drivefreespace"; E={[math]::round(($_.freespace / 1GB),0)}}, 
		volumename, drivetype, volumedirty, volumeserialnumber
		$Results = $Null
		ForEach($drive in $drives)
		{
			If($drive.caption -ne "A:" -and $drive.caption -ne "B:")
			{
				OutputDriveItem $drive
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Drive information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Drive information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
	}
	
	#Get CPU's and stepping
	Write-Verbose "$(Get-Date): `t`t`tProcessor information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Processor(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Processor(s)"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Processor(s)"
	}

	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_Processor
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$Processors = $Results | Select-Object availability, name, description, maxclockspeed, 
		l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		$Results = $Null
		ForEach($processor in $processors)
		{
			OutputProcessorItem $processor
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for Processor information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Processor information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
	}

	#Get Nics
	Write-Verbose "$(Get-Date): `t`t`tNIC information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Network Interface(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Network Interface(s)"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Network Interface(s)"
	}

	[bool]$GotNics = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_networkadapterconfiguration
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$Nics = $Results | Where-Object {$Null -ne $_.ipaddress}
		$Results = $Null

		If($Null -eq $Nics) 
		{ 
			$GotNics = $False 
		} 
		Else 
		{ 
			$GotNics = !($Nics.__PROPERTY_COUNT -eq 0) 
		} 
	
		If($GotNics)
		{
			ForEach($nic in $nics)
			{
				Try
				{
					$ThisNic = Get-WmiObject -computername $RemoteComputerName win32_networkadapter | Where-Object {$_.index -eq $nic.index}
				}
				
				Catch 
				{
					$ThisNic = $Null
				}
				
				If($? -and $Null -ne $ThisNic)
				{
					OutputNicItem $Nic $ThisNic $RemoteComputerName
				}
				ElseIf(!$?)
				{
					Write-Warning "$(Get-Date): Error retrieving NIC information"
					Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 2 "Error retrieving NIC information"
						Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
						Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
						Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
						Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
				}
				Else
				{
					Write-Verbose "$(Get-Date): No results Returned for NIC information"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 2 "No results Returned for NIC information"
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
				}
			}
		}	
	}
	ElseIf(!$?)
	{
		Write-Warning "$(Get-Date): Error retrieving NIC configuration information"
		Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Error retrieving NIC configuration information"
			Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for NIC configuration information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for NIC configuration information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
	}
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 0 0 ""
	}
}

Function OutputComputerItem
{
	Param([object]$Item, [string]$OS)
	
	If($MSWord -or $PDF)
	{
		$ItemInformation = New-Object System.Collections.ArrayList
		$ItemInformation.Add(@{ Data = "Manufacturer"; Value = $Item.manufacturer; }) > $Null
		$ItemInformation.Add(@{ Data = "Model"; Value = $Item.model; }) > $Null
		$ItemInformation.Add(@{ Data = "Domain"; Value = $Item.domain; }) > $Null
		$ItemInformation.Add(@{ Data = "Operating System"; Value = $OS; }) > $Null
		$ItemInformation.Add(@{ Data = "Total Ram"; Value = "$($Item.totalphysicalram) GB"; }) > $Null
		$ItemInformation.Add(@{ Data = "Physical Processors (sockets)"; Value = $Item.NumberOfProcessors; }) > $Null
		$ItemInformation.Add(@{ Data = "Logical Processors (cores w/HT)"; Value = $Item.NumberOfLogicalProcessors; }) > $Null
		$Table = AddWordTable -Hashtable $ItemInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 2 "Manufacturer`t`t`t: " $Item.manufacturer
		Line 2 "Model`t`t`t`t: " $Item.model
		Line 2 "Domain`t`t`t`t: " $Item.domain
		Line 2 "Operating System`t`t: " $OS
		Line 2 "Total Ram`t`t`t: $($Item.totalphysicalram) GB"
		Line 2 "Physical Processors (sockets)`t: " $Item.NumberOfProcessors
		Line 2 "Logical Processors (cores w/HT)`t: " $Item.NumberOfLogicalProcessors
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Manufacturer",($htmlsilver -bor $htmlbold),$Item.manufacturer,$htmlwhite)
		$rowdata += @(,('Model',($htmlsilver -bor $htmlbold),$Item.model,$htmlwhite))
		$rowdata += @(,('Domain',($htmlsilver -bor $htmlbold),$Item.domain,$htmlwhite))
		$rowdata += @(,('Total Ram',($htmlsilver -bor $htmlbold),"$($Item.totalphysicalram) GB",$htmlwhite))
		$rowdata += @(,('Physical Processors (sockets)',($htmlsilver -bor $htmlbold),$Item.NumberOfProcessors,$htmlwhite))
		$rowdata += @(,('Logical Processors (cores w/HT)',($htmlsilver -bor $htmlbold),$Item.NumberOfLogicalProcessors,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputDriveItem
{
	Param([object]$Drive)
	
	$xDriveType = ""
	Switch ($drive.drivetype)
	{
		0	{$xDriveType = "Unknown"; Break}
		1	{$xDriveType = "No Root Directory"; Break}
		2	{$xDriveType = "Removable Disk"; Break}
		3	{$xDriveType = "Local Disk"; Break}
		4	{$xDriveType = "Network Drive"; Break}
		5	{$xDriveType = "Compact Disc"; Break}
		6	{$xDriveType = "RAM Disk"; Break}
		Default {$xDriveType = "Unknown"; Break}
	}
	
	$xVolumeDirty = ""
	If(![String]::IsNullOrEmpty($drive.volumedirty))
	{
		If($drive.volumedirty)
		{
			$xVolumeDirty = "Yes"
		}
		Else
		{
			$xVolumeDirty = "No"
		}
	}

	If($MSWORD -or $PDF)
	{
		$DriveInformation = New-Object System.Collections.ArrayList
		$DriveInformation.Add(@{ Data = "Caption"; Value = $Drive.caption; }) > $Null
		$DriveInformation.Add(@{ Data = "Size"; Value = "$($drive.drivesize) GB"; }) > $Null
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$DriveInformation.Add(@{ Data = "File System"; Value = $Drive.filesystem; }) > $Null
		}
		$DriveInformation.Add(@{ Data = "Free Space"; Value = "$($drive.drivefreespace) GB"; }) > $Null
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$DriveInformation.Add(@{ Data = "Volume Name"; Value = $Drive.volumename; }) > $Null
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$DriveInformation.Add(@{ Data = "Volume is Dirty"; Value = $xVolumeDirty; }) > $Null
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$DriveInformation.Add(@{ Data = "Volume Serial Number"; Value = $Drive.volumeserialnumber; }) > $Null
		}
		$DriveInformation.Add(@{ Data = "Drive Type"; Value = $xDriveType; }) > $Null
		$Table = AddWordTable -Hashtable $DriveInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells `
		-Bold `
		-BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
	}
	ElseIf($Text)
	{
		Line 2 "Caption`t`t: " $drive.caption
		Line 2 "Size`t`t: $($drive.drivesize) GB"
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			Line 2 "File System`t: " $drive.filesystem
		}
		Line 2 "Free Space`t: $($drive.drivefreespace) GB"
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			Line 2 "Volume Name`t: " $drive.volumename
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			Line 2 "Volume is Dirty`t: " $xVolumeDirty
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			Line 2 "Volume Serial #`t: " $drive.volumeserialnumber
		}
		Line 2 "Drive Type`t: " $xDriveType
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Caption",($htmlsilver -bor $htmlbold),$Drive.caption,$htmlwhite)
		$rowdata += @(,('Size',($htmlsilver -bor $htmlbold),"$($drive.drivesize) GB",$htmlwhite))

		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$rowdata += @(,('File System',($htmlsilver -bor $htmlbold),$Drive.filesystem,$htmlwhite))
		}
		$rowdata += @(,('Free Space',($htmlsilver -bor $htmlbold),"$($drive.drivefreespace) GB",$htmlwhite))
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$rowdata += @(,('Volume Name',($htmlsilver -bor $htmlbold),$Drive.volumename,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$rowdata += @(,('Volume is Dirty',($htmlsilver -bor $htmlbold),$xVolumeDirty,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$rowdata += @(,('Volume Serial Number',($htmlsilver -bor $htmlbold),$Drive.volumeserialnumber,$htmlwhite))
		}
		$rowdata += @(,('Drive Type',($htmlsilver -bor $htmlbold),$xDriveType,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputProcessorItem
{
	Param([object]$Processor)
	
	$xAvailability = ""
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"; Break}
		2	{$xAvailability = "Unknown"; Break}
		3	{$xAvailability = "Running or Full Power"; Break}
		4	{$xAvailability = "Warning"; Break}
		5	{$xAvailability = "In Test"; Break}
		6	{$xAvailability = "Not Applicable"; Break}
		7	{$xAvailability = "Power Off"; Break}
		8	{$xAvailability = "Off Line"; Break}
		9	{$xAvailability = "Off Duty"; Break}
		10	{$xAvailability = "Degraded"; Break}
		11	{$xAvailability = "Not Installed"; Break}
		12	{$xAvailability = "Install Error"; Break}
		13	{$xAvailability = "Power Save - Unknown"; Break}
		14	{$xAvailability = "Power Save - Low Power Mode"; Break}
		15	{$xAvailability = "Power Save - Standby"; Break}
		16	{$xAvailability = "Power Cycle"; Break}
		17	{$xAvailability = "Power Save - Warning"; Break}
		Default	{$xAvailability = "Unknown"; Break}
	}

	If($MSWORD -or $PDF)
	{
		$ProcessorInformation = New-Object System.Collections.ArrayList
		$ProcessorInformation.Add(@{ Data = "Name"; Value = $Processor.name; }) > $Null
		$ProcessorInformation.Add(@{ Data = "Description"; Value = $Processor.description; }) > $Null
		$ProcessorInformation.Add(@{ Data = "Max Clock Speed"; Value = "$($processor.maxclockspeed) MHz"; }) > $Null
		If($processor.l2cachesize -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "L2 Cache Size"; Value = "$($processor.l2cachesize) KB"; }) > $Null
		}
		If($processor.l3cachesize -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "L3 Cache Size"; Value = "$($processor.l3cachesize) KB"; }) > $Null
		}
		If($processor.numberofcores -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "Number of Cores"; Value = $Processor.numberofcores; }) > $Null
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$ProcessorInformation.Add(@{ Data = "Number of Logical Processors (cores w/HT)"; Value = $Processor.numberoflogicalprocessors; }) > $Null
		}
		$ProcessorInformation.Add(@{ Data = "Availability"; Value = $xAvailability; }) > $Null
		$Table = AddWordTable -Hashtable $ProcessorInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 2 "Name`t`t`t`t: " $processor.name
		Line 2 "Description`t`t`t: " $processor.description
		Line 2 "Max Clock Speed`t`t`t: $($processor.maxclockspeed) MHz"
		If($processor.l2cachesize -gt 0)
		{
			Line 2 "L2 Cache Size`t`t`t: $($processor.l2cachesize) KB"
		}
		If($processor.l3cachesize -gt 0)
		{
			Line 2 "L3 Cache Size`t`t`t: $($processor.l3cachesize) KB"
		}
		If($processor.numberofcores -gt 0)
		{
			Line 2 "# of Cores`t`t`t: " $processor.numberofcores
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			Line 2 "# of Logical Procs (cores w/HT)`t: " $processor.numberoflogicalprocessors
		}
		Line 2 "Availability`t`t`t: " $xAvailability
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$Processor.name,$htmlwhite)
		$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Processor.description,$htmlwhite))

		$rowdata += @(,('Max Clock Speed',($htmlsilver -bor $htmlbold),"$($processor.maxclockspeed) MHz",$htmlwhite))
		If($processor.l2cachesize -gt 0)
		{
			$rowdata += @(,('L2 Cache Size',($htmlsilver -bor $htmlbold),"$($processor.l2cachesize) KB",$htmlwhite))
		}
		If($processor.l3cachesize -gt 0)
		{
			$rowdata += @(,('L3 Cache Size',($htmlsilver -bor $htmlbold),"$($processor.l3cachesize) KB",$htmlwhite))
		}
		If($processor.numberofcores -gt 0)
		{
			$rowdata += @(,('Number of Cores',($htmlsilver -bor $htmlbold),$Processor.numberofcores,$htmlwhite))
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$rowdata += @(,('Number of Logical Processors (cores w/HT)',($htmlsilver -bor $htmlbold),$Processor.numberoflogicalprocessors,$htmlwhite))
		}
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlbold),$xAvailability,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic, [string]$RemoteComputerName)
	
	$powerMgmt = Get-WmiObject -computername $RemoteComputerName MSPower_DeviceEnable -Namespace root\wmi | Where-Object{$_.InstanceName -match [regex]::Escape($ThisNic.PNPDeviceID)}

	If($? -and $Null -ne $powerMgmt)
	{
		If($powerMgmt.Enable -eq $True)
		{
			$PowerSaving = "Enabled"
		}
		Else
		{
			$PowerSaving = "Disabled"
		}
	}
	Else
	{
        $PowerSaving = "N/A"
	}
	
	$xAvailability = ""
	Switch ($ThisNic.availability)
	{
		1		{$xAvailability = "Other"; Break}
		2		{$xAvailability = "Unknown"; Break}
		3		{$xAvailability = "Running or Full Power"; Break}
		4		{$xAvailability = "Warning"; Break}
		5		{$xAvailability = "In Test"; Break}
		6		{$xAvailability = "Not Applicable"; Break}
		7		{$xAvailability = "Power Off"; Break}
		8		{$xAvailability = "Off Line"; Break}
		9		{$xAvailability = "Off Duty"; Break}
		10		{$xAvailability = "Degraded"; Break}
		11		{$xAvailability = "Not Installed"; Break}
		12		{$xAvailability = "Install Error"; Break}
		13		{$xAvailability = "Power Save - Unknown"; Break}
		14		{$xAvailability = "Power Save - Low Power Mode"; Break}
		15		{$xAvailability = "Power Save - Standby"; Break}
		16		{$xAvailability = "Power Cycle"; Break}
		17		{$xAvailability = "Power Save - Warning"; Break}
		Default	{$xAvailability = "Unknown"; Break}
	}

	#attempt to get Receive Side Scaling setting
	$RSSEnabled = "N/A"
	Try
	{
		#https://ios.developreference.com/article/10085450/How+do+I+enable+VRSS+(Virtual+Receive+Side+Scaling)+for+a+Windows+VM+without+relying+on+Enable-NetAdapterRSS%3F
		$RSSEnabled = (Get-WmiObject -ComputerName $RemoteComputerName MSFT_NetAdapterRssSettingData -Namespace "root\StandardCimV2" -ea 0).Enabled

		If($RSSEnabled)
		{
			$RSSEnabled = "Enabled"
		}
		ELse
		{
			$RSSEnabled = "Disabled"
		}
	}
	
	Catch
	{
		$RSSEnabled = "Not available on $Script:RunningOS"
	}
	
	$xIPAddress = New-Object System.Collections.ArrayList
	ForEach($IPAddress in $Nic.ipaddress)
	{
		$xIPAddress.Add("$($IPAddress)") > $Null
	}

	$xIPSubnet = New-Object System.Collections.ArrayList
	ForEach($IPSubnet in $Nic.ipsubnet)
	{
		$xIPSubnet.Add("$($IPSubnet)") > $Null
	}

	If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
		$xnicdnsdomainsuffixsearchorder = New-Object System.Collections.ArrayList
		ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
		{
			$xnicdnsdomainsuffixsearchorder.Add("$($DNSDomain)") > $Null
		}
	}
	
	If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
	{
		$nicdnsserversearchorder = $nic.dnsserversearchorder
		$xnicdnsserversearchorder = New-Object System.Collections.ArrayList
		ForEach($DNSServer in $nicdnsserversearchorder)
		{
			$xnicdnsserversearchorder.Add("$($DNSServer)") > $Null
		}
	}

	$xdnsenabledforwinsresolution = ""
	If($nic.dnsenabledforwinsresolution)
	{
		$xdnsenabledforwinsresolution = "Yes"
	}
	Else
	{
		$xdnsenabledforwinsresolution = "No"
	}
	
	$xTcpipNetbiosOptions = ""
	Switch ($nic.TcpipNetbiosOptions)
	{
		0	{$xTcpipNetbiosOptions = "Use NetBIOS setting from DHCP Server"; Break}
		1	{$xTcpipNetbiosOptions = "Enable NetBIOS"; Break}
		2	{$xTcpipNetbiosOptions = "Disable NetBIOS"; Break}
		Default	{$xTcpipNetbiosOptions = "Unknown"; Break}
	}
	
	$xwinsenablelmhostslookup = ""
	If($nic.winsenablelmhostslookup)
	{
		$xwinsenablelmhostslookup = "Yes"
	}
	Else
	{
		$xwinsenablelmhostslookup = "No"
	}

	If($MSWORD -or $PDF)
	{
		$NicInformation = New-Object System.Collections.ArrayList
		$NicInformation.Add(@{ Data = "Name"; Value = $ThisNic.Name; }) > $Null
		If($ThisNic.Name -ne $nic.description)
		{
			$NicInformation.Add(@{ Data = "Description"; Value = $Nic.description; }) > $Null
		}
		$NicInformation.Add(@{ Data = "Connection ID"; Value = $ThisNic.NetConnectionID; }) > $Null
		If(validObject $Nic Manufacturer)
		{
			$NicInformation.Add(@{ Data = "Manufacturer"; Value = $Nic.manufacturer; }) > $Null
		}
		$NicInformation.Add(@{ Data = "Availability"; Value = $xAvailability; }) > $Null
		$NicInformation.Add(@{ Data = "Allow the computer to turn off this device to save power"; Value = $PowerSaving; }) > $Null
		$NicInformation.Add(@{ Data = "Receive Side Scaling"; Value = $RSSEnabled; }) > $Null
		$NicInformation.Add(@{ Data = "Physical Address"; Value = $Nic.macaddress; }) > $Null
		If($xIPAddress.Count -gt 1)
		{
			$NicInformation.Add(@{ Data = "IP Address"; Value = $xIPAddress[0]; }) > $Null
			$NicInformation.Add(@{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }) > $Null
			$NicInformation.Add(@{ Data = "Subnet Mask"; Value = $xIPSubnet[0]; }) > $Null
			$cnt = -1
			ForEach($tmp in $xIPAddress)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation.Add(@{ Data = "IP Address"; Value = $tmp; }) > $Null
					$NicInformation.Add(@{ Data = "Subnet Mask"; Value = $xIPSubnet[$cnt]; }) > $Null
				}
			}
		}
		Else
		{
			$NicInformation.Add(@{ Data = "IP Address"; Value = $xIPAddress; }) > $Null
			$NicInformation.Add(@{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }) > $Null
			$NicInformation.Add(@{ Data = "Subnet Mask"; Value = $xIPSubnet; }) > $Null
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$NicInformation.Add(@{ Data = "DHCP Enabled"; Value = $Nic.dhcpenabled; }) > $Null
			$NicInformation.Add(@{ Data = "DHCP Lease Obtained"; Value = $dhcpleaseobtaineddate; }) > $Null
			$NicInformation.Add(@{ Data = "DHCP Lease Expires"; Value = $dhcpleaseexpiresdate; }) > $Null
			$NicInformation.Add(@{ Data = "DHCP Server"; Value = $Nic.dhcpserver; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$NicInformation.Add(@{ Data = "DNS Domain"; Value = $Nic.dnsdomain; }) > $Null
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$NicInformation.Add(@{ Data = "DNS Search Suffixes"; Value = $xnicdnsdomainsuffixsearchorder[0]; }) > $Null
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation.Add(@{ Data = ""; Value = $tmp; }) > $Null
				}
			}
		}
		$NicInformation.Add(@{ Data = "DNS WINS Enabled"; Value = $xdnsenabledforwinsresolution; }) > $Null
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$NicInformation.Add(@{ Data = "DNS Servers"; Value = $xnicdnsserversearchorder[0]; }) > $Null
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation.Add(@{ Data = ""; Value = $tmp; }) > $Null
				}
			}
		}
		$NicInformation.Add(@{ Data = "NetBIOS Setting"; Value = $xTcpipNetbiosOptions; }) > $Null
		$NicInformation.Add(@{ Data = "WINS: Enabled LMHosts"; Value = $xwinsenablelmhostslookup; }) > $Null
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$NicInformation.Add(@{ Data = "Host Lookup File"; Value = $Nic.winshostlookupfile; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$NicInformation.Add(@{ Data = "Primary Server"; Value = $Nic.winsprimaryserver; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$NicInformation.Add(@{ Data = "Secondary Server"; Value = $Nic.winssecondaryserver; }) > $Null
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$NicInformation.Add(@{ Data = "Scope ID"; Value = $Nic.winsscopeid; }) > $Null
		}
		$Table = AddWordTable -Hashtable $NicInformation -Columns Data,Value -List -AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 2 "Name`t`t`t: " $ThisNic.Name
		If($ThisNic.Name -ne $nic.description)
		{
			Line 2 "Description`t`t: " $nic.description
		}
		Line 2 "Connection ID`t`t: " $ThisNic.NetConnectionID
		If(validObject $Nic Manufacturer)
		{
			Line 2 "Manufacturer`t`t: " $Nic.manufacturer
		}
		Line 2 "Availability`t`t: " $xAvailability
		Line 2 "Allow computer to turn "
		Line 2 "off device to save power: " $PowerSaving
		Line 2 "Receive Side Scaling`t: " $RSSEnabled
		Line 2 "Physical Address`t: " $nic.macaddress
		Line 2 "IP Address`t`t: " $xIPAddress[0]
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "  " $tmp
			}
		}
		Line 2 "Default Gateway`t`t: " $Nic.Defaultipgateway
		Line 2 "Subnet Mask`t`t: " $xIPSubnet[0]
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "  " $tmp
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			Line 2 "DHCP Enabled`t`t: " $nic.dhcpenabled
			Line 2 "DHCP Lease Obtained`t: " $dhcpleaseobtaineddate
			Line 2 "DHCP Lease Expires`t: " $dhcpleaseexpiresdate
			Line 2 "DHCP Server`t`t:" $nic.dhcpserver
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			Line 2 "DNS Domain`t`t: " $nic.dnsdomain
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Search Suffixes`t: " $xnicdnsdomainsuffixsearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 5 "  " $tmp
				}
			}
		}
		Line 2 "DNS WINS Enabled`t: " $xdnsenabledforwinsresolution
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Servers`t`t: " $xnicdnsserversearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 5 "  " $tmp
				}
			}
		}
		Line 2 "NetBIOS Setting`t`t: " $xTcpipNetbiosOptions
		Line 2 "WINS:"
		Line 3 "Enabled LMHosts`t: " $xwinsenablelmhostslookup
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			Line 3 "Host Lookup File`t: " $nic.winshostlookupfile
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			Line 3 "Primary Server`t: " $nic.winsprimaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			Line 3 "Secondary Server`t: " $nic.winssecondaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			Line 3 "Scope ID`t`t: " $nic.winsscopeid
		}
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$ThisNic.Name,$htmlwhite)
		If($ThisNic.Name -ne $nic.description)
		{
			$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Nic.description,$htmlwhite))
		}
		$rowdata += @(,('Connection ID',($htmlsilver -bor $htmlbold),$ThisNic.NetConnectionID,$htmlwhite))
		If(validObject $Nic Manufacturer)
		{
			$rowdata += @(,('Manufacturer',($htmlsilver -bor $htmlbold),$Nic.manufacturer,$htmlwhite))
		}
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlbold),$xAvailability,$htmlwhite))
		$rowdata += @(,('Allow the computer to turn off this device to save power',($htmlsilver -bor $htmlbold),$PowerSaving,$htmlwhite))
		$rowdata += @(,('Physical Address',($htmlsilver -bor $htmlbold),$Nic.macaddress,$htmlwhite))
		$rowdata += @(,('Receive Side Scaling',($htmlsilver -bor $htmlbold),$RSSEnabled,$htmlwhite))
		$rowdata += @(,('IP Address',($htmlsilver -bor $htmlbold),$xIPAddress[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('IP Address',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		$rowdata += @(,('Default Gateway',($htmlsilver -bor $htmlbold),$Nic.Defaultipgateway[0],$htmlwhite))
		$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlbold),$xIPSubnet[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$rowdata += @(,('DHCP Enabled',($htmlsilver -bor $htmlbold),$Nic.dhcpenabled,$htmlwhite))
			$rowdata += @(,('DHCP Lease Obtained',($htmlsilver -bor $htmlbold),$dhcpleaseobtaineddate,$htmlwhite))
			$rowdata += @(,('DHCP Lease Expires',($htmlsilver -bor $htmlbold),$dhcpleaseexpiresdate,$htmlwhite))
			$rowdata += @(,('DHCP Server',($htmlsilver -bor $htmlbold),$Nic.dhcpserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$rowdata += @(,('DNS Domain',($htmlsilver -bor $htmlbold),$Nic.dnsdomain,$htmlwhite))
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Search Suffixes',($htmlsilver -bor $htmlbold),$xnicdnsdomainsuffixsearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('DNS WINS Enabled',($htmlsilver -bor $htmlbold),$xdnsenabledforwinsresolution,$htmlwhite))
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Servers',($htmlsilver -bor $htmlbold),$xnicdnsserversearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('NetBIOS Setting',($htmlsilver -bor $htmlbold),$xTcpipNetbiosOptions,$htmlwhite))
		$rowdata += @(,('WINS: Enabled LMHosts',($htmlsilver -bor $htmlbold),$xwinsenablelmhostslookup,$htmlwhite))
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$rowdata += @(,('Host Lookup File',($htmlsilver -bor $htmlbold),$Nic.winshostlookupfile,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$rowdata += @(,('Primary Server',($htmlsilver -bor $htmlbold),$Nic.winsprimaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$rowdata += @(,('Secondary Server',($htmlsilver -bor $htmlbold),$Nic.winssecondaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$rowdata += @(,('Scope ID',($htmlsilver -bor $htmlbold),$Nic.winsscopeid,$htmlwhite))
		}

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region GetComputerServices
Function GetComputerServices 
{
	Param([string]$RemoteComputerName)
	# modified 29-Apr-2018 to change from Arrays to New-Object System.Collections.ArrayList
	
	#Get Computer services info
	Write-Verbose "$(Get-Date): `t`tProcessing Computer services information"
	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 "Services"
	}
	ElseIf($Text)
	{
		Line 0 "Services"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "Services"
	}

	Try
	{
		#Iain Brighton optimization 5-Jun-2014
		#Replaced with a single call to retrieve services via WMI. The repeated
		## "Get-WMIObject Win32_Service -Filter" calls were the major delays in the script.
		## If we need to retrieve the StartUp type might as well just use WMI.
		
		#V2.20 changed to @()
		$Services = @(Get-WMIObject Win32_Service -ComputerName $RemoteComputerName | Sort-Object DisplayName)
	}
	
	Catch
	{
		$Services = $Null
	}
	
	If($? -and $Null -ne $Services)
	{
		[int]$NumServices = $Services.count
		Write-Verbose "$(Get-Date): `t`t$($NumServices) Services found"

		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "Services ($NumServices Services found)"

			$ServicesWordTable = New-Object System.Collections.ArrayList
			## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
			$HighlightedCells = New-Object System.Collections.ArrayList
			## Seed the $Services row index from the second row
			[int] $CurrentServiceIndex = 2;
		}
		ElseIf($Text)
		{
			Line 0 "Services ($NumServices Services found)"
			Line 0 ""
			
			#V2.16 addition
			[int]$MaxDisplayNameLength = ($Services.DisplayName | Measure-Object -Maximum -Property Length).Maximum
			If($MaxDisplayNameLength -gt 12) #12 is length of "Display Name"
			{
				#10 is length of "Display Name" minus 2 to allow for spacing between columns
				Line 1 ("Display Name" + (' ' * ($MaxDisplayNameLength - 10))) -NoNewLine
			}
			Else
			{
				Line 1 "Display Name " -NoNewLine
			}
			Line 1 "Status  " -NoNewLine
			Line 1 "Startup Type"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "Services ($NumServices Services found)"
			$rowdata = @()
		}

		ForEach($Service in $Services) 
		{
			#Write-Verbose "$(Get-Date): `t`t`t Processing service $($Service.DisplayName)";

			If($MSWord -or $PDF)
			{

				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ 
				DisplayName = $Service.DisplayName; 
				Status = $Service.State; 
				StartMode = $Service.StartMode
				}

				## Add the hash to the array
				$ServicesWordTable.Add($WordTableRowHash) > $Null

				## Store "to highlight" cell references
				If($Service.State -like "Stopped" -and $Service.StartMode -like "Auto") 
				{
					$HighlightedCells.Add(@{ Row = $CurrentServiceIndex; Column = 2; }) > $Null
				}
				$CurrentServiceIndex++;
			}
			ElseIf($Text)
			{
				#V2.16 change
				If(($Service.DisplayName).Length -lt ($MaxDisplayNameLength))
				{
					[int]$NumOfSpaces = (($MaxDisplayNameLength) - ($Service.DisplayName.Length)) + 2 #+2 to allow for column spacing
					$tmp1 = ($($Service.DisplayName) + (' ' * $NumOfSPaces))
					Line 1 $tmp1 -NoNewLine
				}
				Else
				{
					Line 1 "$($Service.DisplayName)  " -NoNewLine
				}
				Line 1 "$($Service.State) " -NoNewLine
				Line 1 $Service.StartMode
			}
			ElseIf($HTML)
			{
				If($Service.State -like "Stopped" -and $Service.StartMode -like "Auto") 
				{
					$HighlightedCells = $htmlred
				}
				Else
				{
					$HighlightedCells = $htmlwhite
				} 
				$rowdata += @(,($Service.DisplayName,$htmlwhite,
								$Service.State,$HighlightedCells,
								$Service.StartMode,$htmlwhite))
			}
		}

		If($MSWord -or $PDF)
		{
			## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
			$Table = AddWordTable -Hashtable $ServicesWordTable `
			-Columns DisplayName, Status, StartMode `
			-Headers "Display Name", "Status", "Startup Type" `
			-AutoFit $wdAutoFitContent;

			## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
			## IB - Set the required highlighted cells
			SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

			#indent the entire table 1 tab stop
			$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			#V2.16 change
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$columnHeaders = @('Display Name',($htmlsilver -bor $htmlbold),'Status',($htmlsilver -bor $htmlbold),'Startup Type',($htmlsilver -bor $htmlbold))
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 " "
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "No services were retrieved."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Warning: No Services were retrieved" "" $Null 0 $False $True
			WriteWordLine 0 1 "If this is a trusted Forest, you may need to rerun the" "" $Null 0 $False $True
			WriteWordLine 0 1 "script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Warning: No Services were retrieved"
			Line 1 "If this is a trusted Forest, you may need to rerun the"
			Line 1 "script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Warning: No Services were retrieved" "" $Null 0 $htmlbold
			WriteHTMLLine 0 1 "If this is a trusted Forest, you may need to rerun the" "" $Null 0 $htmlbold
			WriteHTMLLine 0 1 "script with Domain Admin credentials from the trusted Forest." "" $Null 0 $htmlbold
		}
	}
	Else
	{
		Write-Warning "Services retrieval was successful but no services were returned."
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "Services retrieval was successful but no services were returned." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 0 "Services retrieval was successful but no services were returned."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Services retrieval was successful but no services were returned." "" $Null 0 $htmlbold
		}
	}
}
#endregion

#region BuildDCDNSIPConfigTable
Function BuildDCDNSIPConfigTable
{
	Param([string]$RemoteComputerName, [string]$Site)
	
	[bool]$GotNics = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_networkadapterconfiguration
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$Nics = $Results | Where-Object {$Null -ne $_.ipaddress}
		$Results = $Null

		If($Null -eq $Nics) 
		{ 
			$GotNics = $False 
		} 
		Else 
		{ 
			$GotNics = !($Nics.__PROPERTY_COUNT -eq 0) 
		} 
	
		If($GotNics)
		{
			ForEach($nic in $nics)
			{
				Try
				{
					$ThisNic = Get-WmiObject -computername $RemoteComputerName win32_networkadapter | Where-Object {$_.index -eq $nic.index}
				}
				
				Catch 
				{
					$ThisNic = $Null
				}
				
				If($? -and $Null -ne $ThisNic)
				{
					Write-Verbose "$(Get-Date): `t`t`tGather DC DNS IP Config info"
					$xIPAddress = @()
					ForEach($IPAddress in $Nic.ipaddress)
					{
						$xIPAddress += "$($IPAddress)"
					}
					
					If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
					{
						$nicdnsserversearchorder = $nic.dnsserversearchorder
						$xnicdnsserversearchorder = @()
						ForEach($DNSServer in $nicdnsserversearchorder)
						{
							$xnicdnsserversearchorder += "$($DNSServer)"
						}
					}

					$obj = New-Object -TypeName PSObject
					$obj | Add-Member -MemberType NoteProperty -Name DCName -Value $RemoteComputerName
					$obj | Add-Member -MemberType NoteProperty -Name DCSite -Value $Site
					If($xIPAddress.Count -gt 1)
					{
						$obj | Add-Member -MemberType NoteProperty -Name DCIpAddress1 -Value $xIPAddress[0]
						$obj | Add-Member -MemberType NoteProperty -Name DCIpAddress2 -Value $xIPAddress[1]
					}
					Else
					{
						$obj | Add-Member -MemberType NoteProperty -Name DCIpAddress1 -Value $xIPAddress[0]
						$obj | Add-Member -MemberType NoteProperty -Name DCIpAddress2 -Value ""
					}
					
					If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
					{
						$obj | Add-Member -MemberType NoteProperty -Name DCDNS1 -Value $xnicdnsserversearchorder[0]
						If($Null -ne $xnicdnsserversearchorder[1])
						{
							$obj | Add-Member -MemberType NoteProperty -Name DCDNS2 -Value $xnicdnsserversearchorder[1]
						}
						Else
						{
							$obj | Add-Member -MemberType NoteProperty -Name DCDNS2 -Value " "
						}
						
						If($Null -ne $xnicdnsserversearchorder[2])
						{
							$obj | Add-Member -MemberType NoteProperty -Name DCDNS3 -Value $xnicdnsserversearchorder[2]
						}
						Else
						{
							$obj | Add-Member -MemberType NoteProperty -Name DCDNS3 -Value " "
						}
						
						If($Null -ne $xnicdnsserversearchorder[3])
						{
							$obj | Add-Member -MemberType NoteProperty -Name DCDNS4 -Value $xnicdnsserversearchorder[3]
						}
						Else
						{
							$obj | Add-Member -MemberType NoteProperty -Name DCDNS4 -Value " "
						}
					}

					[void]$Script:DCDNSIPInfo.Add($obj)
				}
			}
		}	
	}
	Else
	{
		Write-Verbose "$(Get-Date): No results Returned for NIC configuration information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for NIC configuration information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
	}
}
#endregion

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. SMith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break } #changed 13-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			# fix in 2.23 thanks to Johan Kallio 'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
			'zh-'	{ '自动目录 2'; Break }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$ChineseArray = 2052,3076,5124,4100
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_} {$CultureCode = "ca-"}
		{$ChineseArray -contains $_} {$CultureCode = "zh-"}
		{$DanishArray -contains $_} {$CultureCode = "da-"}
		{$DutchArray -contains $_} {$CultureCode = "nl-"}
		{$EnglishArray -contains $_} {$CultureCode = "en-"}
		{$FinnishArray -contains $_} {$CultureCode = "fi-"}
		{$FrenchArray -contains $_} {$CultureCode = "fr-"}
		{$GermanArray -contains $_} {$CultureCode = "de-"}
		{$NorwegianArray -contains $_} {$CultureCode = "nb-"}
		{$PortugueseArray -contains $_} {$CultureCode = "pt-"}
		{$SpanishArray -contains $_} {$CultureCode = "es-"}
		{$SwedishArray -contains $_} {$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		'zh-'	{
				If($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ('奥斯汀', '边线型', '花丝', '怀旧', '积分',
					'离子(浅色)', '离子(深色)', '母版型', '平面', '切片(浅色)',
					'切片(深色)', '丝状', '网格', '镶边', '信号灯',
					'运动型')
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	#[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|Where-Object {$_.SessionId -eq $SessionID}) -ne $Null
	[bool]$wordrunning = $null –ne ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID})	
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n"
		Exit
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function Set-DocumentProperty {
    <#
	.SYNOPSIS
	Function to set the Title Page document properties in MS Word
	.DESCRIPTION
	Long description
	.PARAMETER Document
	Current Document Object
	.PARAMETER DocProperty
	Parameter description
	.PARAMETER Value
	Parameter description
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value 'MyTitle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value 'MyCompany'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value 'Jim Moyle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value 'MySubjectTitle'
	.NOTES
	Function Created by Jim Moyle June 2017
	Twitter : @JimMoyle
	#>
    param (
        [object]$Document,
        [String]$DocProperty,
        [string]$Value
    )
    try {
        $binding = "System.Reflection.BindingFlags" -as [type]
        $builtInProperties = $Document.BuiltInDocumentProperties
        $property = [System.__ComObject].invokemember("item", $binding::GetProperty, $null, $BuiltinProperties, $DocProperty)
        [System.__ComObject].invokemember("value", $binding::SetProperty, $null, $property, $Value)
    }
    catch {
        Write-Warning "Failed to set $DocProperty to $Value"
    }
}

Function FindWordDocumentEnd
{
	#return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function SetupWord
{
	Write-Verbose "$(Get-Date): Setting up Word"
    
	# Setup word for output
	Write-Verbose "$(Get-Date): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
		`t`t
		The Word object could not be created.
		`n`n
		`t`t
		You may need to repair your Word installation.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		Exit
	}

	Write-Verbose "$(Get-Date): Determine Word language value"
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
		`t`t
		Unable to determine the Word language value.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}
	Write-Verbose "$(Get-Date): Word language value is $($Script:WordLanguageValue)"
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
		`t`t
		Microsoft Word 2007 is no longer supported.
		`n`n
		`t`t
		Script will end.
		`n`n
		"
		AbortScript
	}
	ElseIf($Script:WordVersion -eq 0)
	{
		Write-Error "
		`n`n
		`t`t
		The Word Version is 0. You should run a full online repair of your Office installation.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		Exit
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
		`t`t
		You are running an untested or unsupported version of Microsoft Word.
		`n`n
		`t`t
		Script will end.
		`n`n
		`t`t
		Please send info on your version of Word to webster@carlwebster.com
		`n`n
		"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($Script:CoName))
	{
		Write-Verbose "$(Get-Date): Company name is blank.  Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
			Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
			Write-Warning "`n`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.`n`n"
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date): Updated company name to $($Script:CoName)"
		}
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date): Check Default Cover Page for $($WordCultureCode)"
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}

			'zh-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "边线型"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Culture code $($Script:WordCultureCode)"
		Write-Error "
		`n`n
		`t`t
		For $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}

	ShowScriptOptions

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where-Object {$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach-Object{
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($Null -ne $BuildingBlocks)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($Null -ne $part)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "This report will not have a Cover Page."
	}

	Write-Verbose "$(Get-Date): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Null -eq $Script:Doc)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
		`t`t
		An empty Word document could not be created.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
		`t`t
		An unknown error happened selecting the entire Word document for default formatting options.
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($Null -eq $toc)
		{
			Write-Verbose "$(Get-Date): "
			Write-Verbose "$(Get-Date): Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved."
			Write-Warning "This report will not have a Table of Contents."
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Table of Contents are not installed."
		Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
	}

	#set the footer
	Write-Verbose "$(Get-Date): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date): Get the footer and format font"
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach ($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Verbose "$(Get-Date): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	Write-Verbose "$(Get-Date):"
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#updated 8-Jun-2017 with additional cover page fields
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Verbose "$(Get-Date): Set Cover Page Properties"
			#8-Jun-2017 put these 4 items in alpha order
            Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value $UserName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value $Script:CoName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value $SubjectTitle
            Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value $Script:title

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where-Object {$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "Abstract"}
			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
			}
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "CompanyAddress"}
			#set the text
			[string]$abstract = $CompanyAddress
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "CompanyEmail"}
			#set the text
			[string]$abstract = $CompanyEmail
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "CompanyFax"}
			#set the text
			[string]$abstract = $CompanyFax
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "CompanyPhone"}
			#set the text
			[string]$abstract = $CompanyPhone
			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where-Object {$_.basename -eq "PublishDate"}
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Verbose "$(Get-Date): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}
#endregion

#region registry functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified local registry value or $Null if it is missing
Function Get-LocalRegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

Function Get-RegistryValue
{
	# Gets the specified registry value or $Null if it is missing
	[CmdletBinding()]
	Param([string]$path, [string]$name, [string]$ComputerName)
	If($ComputerName -eq $env:computername -or $ComputerName -eq "LocalHost")
	{
		$key = Get-Item -LiteralPath $path -EA 0
		If($key)
		{
			Return $key.GetValue($name, $Null)
		}
		Else
		{
			Return $Null
		}
	}
	Else
	{
		#path needed here is different for remote registry access
		$path1 = $path.SubString(6)
		$path2 = $path1.Replace('\','\\')
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
		$RegKey= $Reg.OpenSubKey($path2)
		$Results = $RegKey.GetValue($name)
		If($Null -ne $Results)
		{
			Return $Results
		}
		Else
		{
			Return $Null
		}
	}
}
#endregion

#region word, text and html line output functions
Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexch on Twitter
#https://essential.exchange/blog
#for creating the formatted text report
#created March 2011
#updated March 2014
# updated March 2019 to use StringBuilder (about 100 times more efficient than simple strings)
{
	Param
	(
		[Int]    $tabs = 0, 
		[String] $name = '', 
		[String] $value = '', 
		[String] $newline = [System.Environment]::NewLine, 
		[Switch] $nonewline
	)

	while( $tabs -gt 0 )
	{
		$null = $global:Output.Append( "`t" )
		$tabs--
	}

	If( $nonewline )
	{
		$null = $global:Output.Append( $name + $value )
	}
	Else
	{
		$null = $global:Output.AppendLine( $name + $value )
	}
}
	
Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1; Break}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2; Break}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3; Break}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4; Break}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}

#***********************************************************************************************************
# WriteHTMLLine
#***********************************************************************************************************

<#
.Synopsis
	Writes a line of output for HTML output
.DESCRIPTION
	This function formats an HTML line
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

	0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	WriteHTMLLine 0 0 " "

	Writes a blank line with no style or tab stops, obviously none needed.

.EXAMPLE
	WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"

	Writes a line with 1 tab stop.

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $Null 0 $htmlitalics

	Writes a line omitting font and font size and setting the italics attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $Null 0 $htmlbold

	Writes a line omitting font and font size and setting the bold attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $Null 0 ($htmlbold -bor $htmlitalics)

	Writes a line omitting font and font size and setting both italics and bold options

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $Null 2  # 10 point font

	Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 

	Writes a line using Courier New Font and 0 font point size (default = 2 if set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlbold -bor $htmlred -bor $htmlitalics)

	Writes a line using Courier New Font with first and second string values to be used, also uses 10 point font with bold, italics and red color options set.

.NOTES

	Font Size - Unlike word, there is a limited set of font sizes that can be used in HTML.  They are:
		0 - default which actually gives it a 2 or 10 point.
		1 - 7.5 point font size
		2 - 10 point
		3 - 13.5 point
		4 - 15 point
		5 - 18 point
		6 - 24 point
		7 - 36 point
	Any number larger than 7 defaults to 7

	Style - Refers to the headers that are used with output and resemble the headers in word, 
	HTML supports headers h1-h6 and h1-h4 are more commonly used.  Unlike word, H1 will not 
	give you a blue colored font, you will have to set that yourself.

	Colors and Bold/Italics Flags are:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack       
#>

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
#headings fixed 12-Oct-2016 by Webster
#errors with $HTMLStyle fixed 7-Dec-2017 by Webster
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName="Calibri",
	[int]$fontSize=1,
	[int]$options=$htmlblack)


	#Build output style
	[string]$output = ""

	If([String]::IsNullOrEmpty($Name))	
	{
		$HTMLBody = "<p></p>"
	}
	Else
	{
		$color = CheckHTMLColor $options

		#build # of tabs

		While($tabs -gt 0)
		{ 
			$output += "&nbsp;&nbsp;&nbsp;&nbsp;"; $tabs--; 
		}

		$HTMLFontName = $fontName		

		$HTMLBody = ""

		If($options -band $htmlitalics) 
		{
			$HTMLBody += "<i>"
		} 

		If($options -band $htmlbold) 
		{
			$HTMLBody += "<b>"
		} 

		#output the rest of the parameters.
		$output += $name + $value

		Switch ($style)
		{
			1 {$HTMLStyle = "<h1>"; Break}
			2 {$HTMLStyle = "<h2>"; Break}
			3 {$HTMLStyle = "<h3>"; Break}
			4 {$HTMLStyle = "<h4>"; Break}
			Default {$HTMLStyle = ""; Break}
		}

		$HTMLBody += $HTMLStyle + $output

		Switch ($style)
		{
			1 {$HTMLStyle = "</h1>"; Break}
			2 {$HTMLStyle = "</h2>"; Break}
			3 {$HTMLStyle = "</h3>"; Break}
			4 {$HTMLStyle = "</h4>"; Break}
			Default {$HTMLStyle = ""; Break}
		}

		#added by webster 12-oct-2016
		#if a heading, don't add the <br>
		#moved to after the two switch statements on 7-Dec-2017 to fix $HTMLStyle has not been set error
		If($HTMLStyle -eq "")
		{
			$HTMLBody += "<br><font face='" + $HTMLFontName + "' " + "color='" + $color + "' size='"  + $fontsize + "'>"
		}
		Else
		{
			$HTMLBody += "<font face='" + $HTMLFontName + "' " + "color='" + $color + "' size='"  + $fontsize + "'>"
		}
		
		$HTMLBody += $HTMLStyle +  "</font>"

		If($options -band $htmlitalics) 
		{
			$HTMLBody += "</i>"
		} 

		If($options -band $htmlbold) 
		{
			$HTMLBody += "</b>"
		} 

		#added by webster 12-oct-2016
		#if a heading, don't add the <br />
		#moved to inside the Else statement on 7-Dec-2017 to fix $HTMLStyle has not been set error
		If($HTMLStyle -eq "")
		{
			$HTMLBody += "<br />"
		}
	}

	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null
}
#endregion

#region HTML table functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable function
# Created by Ken Avram
# modified by Jake Rutski
#***********************************************************************************************************
Function AddHTMLTable
{
	Param([string]$fontName="Calibri",
	[int]$fontSize=2,
	[int]$colCount=0,
	[int]$rowCount=0,
	[object[]]$rowInfo=@(),
	[object[]]$fixedInfo=@())

	For($rowidx = $RowIndex;$rowidx -le $rowCount;$rowidx++)
	{
		$rd = @($rowInfo[$rowidx - 2])
		$htmlbody = $htmlbody + "<tr>"
		For($columnIndex = 0; $columnIndex -lt $colCount; $columnindex+=2)
		{
			$tmp = CheckHTMLColor $rd[$columnIndex+1]

			If($fixedInfo.Length -eq 0)
			{
				$htmlbody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$htmlbody += "<td style=""width:$($fixedInfo[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If($rd[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "<b>"
			}
			If($rd[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "<i>"
			}
			If($Null -ne $rd[$columnIndex])
			{
				$cell = $rd[$columnIndex].tostring()
				If($cell -eq " " -or $cell.length -eq 0)
				{
					$htmlbody += "&nbsp;&nbsp;&nbsp;"
				}
				Else
				{
					For($i=0;$i -lt $cell.length;$i++)
					{
						If($cell[$i] -eq " ")
						{
							$htmlbody += "&nbsp;"
						}
						If($cell[$i] -ne " ")
						{
							Break
						}
					}
					$htmlbody += $cell
				}
			}
			Else
			{
				$htmlbody += "&nbsp;&nbsp;&nbsp;"
			}
			If($rd[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "</b>"
			}
			If($rd[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "</i>"
			}
			$htmlbody += "</font></td>"
		}
		$htmlbody += "</tr>"
	}
	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
}

#***********************************************************************************************************
# FormatHTMLTable 
# Created by Ken Avram
# modified by Jake Rutski
#***********************************************************************************************************

<#
.Synopsis
	Format table for HTML output document
.DESCRIPTION
	This function formats a table for HTML from an array of strings
.PARAMETER noBorder
	If set to $true, a table will be generated without a border (border='0')
.PARAMETER noHeadCols
	This parameter should be used when generating tables without column headers
	Set this parameter equal to the number of columns in the table
.PARAMETER rowArray
	This parameter contains the row data array for the table
.PARAMETER columnArray
	This parameter contains column header data for the table
.PARAMETER fixedWidth
	This parameter contains widths for columns in pixel format ("100px") to override auto column widths
	The variable should contain a width for each column you wish to override the auto-size setting
	For example: $columnWidths = @("100px","110px","120px","130px","140px")

.USAGE
	FormatHTMLTable <Table Header> <Table Format> <Font Name> <Font Size>

.EXAMPLE
	FormatHTMLTable "Table Heading" "auto" "Calibri" 3

	This example formats a table and writes it out into an html file.  All of the parameters are optional
	defaults are used if not supplied.

	for <Table format>, the default is auto which will autofit the text into the columns and adjust to the longest text in that column.  You can also use percentage i.e. 25%
	which will take only 25% of the line and will auto word wrap the text to the next line in the column.  Also, instead of using a percentage, you can use pixels i.e. 400px.

	FormatHTMLTable "Table Heading" "auto" -rowArray $rowData -columnArray $columnData

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, column header data from $columnData and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -noHeadCols 3

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, no header, and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -fixedWidth $fixedColumns

	This example creates an HTML table with a heading of 'Table Heading, no header, row data from $rowData, and fixed columns defined by $fixedColumns

.NOTES
	In order to use the formatted table it first has to be loaded with data.  Examples below will show how to load the table:

	First, initialize the table array

	$rowdata = @()

	Then Load the array.  If you are using column headers then load those into the column headers array, otherwise the first line of the table goes into the column headers array
	and the second and subsequent lines go into the $rowdata table as shown below:

	$columnHeaders = @('Display Name',($htmlsilver -bor $htmlbold),'Status',($htmlsilver -bor $htmlbold),'Startup Type',($htmlsilver -bor $htmlbold))

	The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics.  For the anding, parens are required or it will
	not format correctly.

	This is following by adding rowdata as shown below.  As more columns are added the columns will auto adjust to fit the size of the page.

	$rowdata = @()
	$columnHeaders = @("User Name",($htmlsilver -bor $htmlbold),$UserName,$htmlwhite)
	$rowdata += @(,('Save as PDF',($htmlsilver -bor $htmlbold),$PDF.ToString(),$htmlwhite))
	$rowdata += @(,('Save as TEXT',($htmlsilver -bor $htmlbold),$TEXT.ToString(),$htmlwhite))
	$rowdata += @(,('Save as WORD',($htmlsilver -bor $htmlbold),$MSWORD.ToString(),$htmlwhite))
	$rowdata += @(,('Save as HTML',($htmlsilver -bor $htmlbold),$HTML.ToString(),$htmlwhite))
	$rowdata += @(,('Add DateTime',($htmlsilver -bor $htmlbold),$AddDateTime.ToString(),$htmlwhite))
	$rowdata += @(,('Hardware Inventory',($htmlsilver -bor $htmlbold),$Hardware.ToString(),$htmlwhite))
	$rowdata += @(,('Computer Name',($htmlsilver -bor $htmlbold),$ComputerName,$htmlwhite))
	$rowdata += @(,('Filename1',($htmlsilver -bor $htmlbold),$Script:FileName1,$htmlwhite))
	$rowdata += @(,('OS Detected',($htmlsilver -bor $htmlbold),$Script:RunningOS,$htmlwhite))
	$rowdata += @(,('PSUICulture',($htmlsilver -bor $htmlbold),$PSCulture,$htmlwhite))
	$rowdata += @(,('PoSH version',($htmlsilver -bor $htmlbold),$Host.Version.ToString(),$htmlwhite))
	FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table" -rowArray $rowdata

	The 'rowArray' paramater is mandatory to build the table, but it is not set as such in the function - if nothing is passed, the table will be empty.

	Colors and Bold/Italics Flags are shown below:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack     

#>

Function FormatHTMLTable
{
	Param([string]$tableheader,
	[string]$tablewidth="auto",
	[string]$fontName="Calibri",
	[int]$fontSize=2,
	[switch]$noBorder=$false,
	[int]$noHeadCols=1,
	[object[]]$rowArray=@(),
	[object[]]$fixedWidth=@(),
	[object[]]$columnArray=@())

	$HTMLBody = "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>"

	If($columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If($Null -ne $rowArray)
	{
		$NumRows = $rowArray.length + 1
	}
	Else
	{
		$NumRows = 1
	}

	If($noBorder)
	{
		$htmlbody += "<table border='0' width='" + $tablewidth + "'>"
	}
	Else
	{
		$htmlbody += "<table border='1' width='" + $tablewidth + "'>"
	}

	If(!($columnArray.Length -eq 0))
	{
		$htmlbody += "<tr>"

		For($columnIndex = 0; $columnIndex -lt $NumCols; $columnindex+=2)
		{
			$tmp = CheckHTMLColor $columnArray[$columnIndex+1]
			If($fixedWidth.Length -eq 0)
			{
				$htmlbody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$htmlbody += "<td style=""width:$($fixedWidth[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If($columnArray[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "<b>"
			}
			If($columnArray[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "<i>"
			}
			If($Null -ne $columnArray[$columnIndex])
			{
				If($columnArray[$columnIndex] -eq " " -or $columnArray[$columnIndex].length -eq 0)
				{
					$htmlbody += "&nbsp;&nbsp;&nbsp;"
				}
				Else
				{
					For($i=0;$i -lt $columnArray[$columnIndex].length;$i+=2)
					{
						If($columnArray[$columnIndex][$i] -eq " ")
						{
							$htmlbody += "&nbsp;"
						}
						If($columnArray[$columnIndex][$i] -ne " ")
						{
							Break
						}
					}
					$htmlbody += $columnArray[$columnIndex]
				}
			}
			Else
			{
				$htmlbody += "&nbsp;&nbsp;&nbsp;"
			}
			If($columnArray[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "</b>"
			}
			If($columnArray[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "</i>"
			}
			$htmlbody += "</font></td>"
		}
		$htmlbody += "</tr>"
	}
	$rowindex = 2
	If($Null -ne $rowArray)
	{
		AddHTMLTable $fontName $fontSize -colCount $numCols -rowCount $NumRows -rowInfo $rowArray -fixedInfo $fixedWidth
		$rowArray = @()
		$htmlbody = "</table>"
	}
	Else
	{
		$HTMLBody += "</table>"
	}	
	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
}
#endregion

#region other HTML functions
#***********************************************************************************************************
# CheckHTMLColor - Called from AddHTMLTable WriteHTMLLine and FormatHTMLTable
#***********************************************************************************************************
Function CheckHTMLColor
{
	Param($hash)

	If($hash -band $htmlwhite)
	{
		Return $htmlwhitemask
	}
	If($hash -band $htmlred)
	{
		Return $htmlredmask
	}
	If($hash -band $htmlcyan)
	{
		Return $htmlcyanmask
	}
	If($hash -band $htmlblue)
	{
		Return $htmlbluemask
	}
	If($hash -band $htmldarkblue)
	{
		Return $htmldarkbluemask
	}
	If($hash -band $htmllightblue)
	{
		Return $htmllightbluemask
	}
	If($hash -band $htmlpurple)
	{
		Return $htmlpurplemask
	}
	If($hash -band $htmlyellow)
	{
		Return $htmlyellowmask
	}
	If($hash -band $htmllime)
	{
		Return $htmllimemask
	}
	If($hash -band $htmlmagenta)
	{
		Return $htmlmagentamask
	}
	If($hash -band $htmlsilver)
	{
		Return $htmlsilvermask
	}
	If($hash -band $htmlgray)
	{
		Return $htmlgraymask
	}
	If($hash -band $htmlblack)
	{
		Return $htmlblackmask
	}
	If($hash -band $htmlorange)
	{
		Return $htmlorangemask
	}
	If($hash -band $htmlmaroon)
	{
		Return $htmlmaroonmask
	}
	If($hash -band $htmlgreen)
	{
		Return $htmlgreenmask
	}
	If($hash -band $htmlolive)
	{
		Return $htmlolivemask
	}
}

Function SetupHTML
{
	Write-Verbose "$(Get-Date): Setting up HTML"
	If(!$AddDateTime)
	{
		[string]$Script:FileName1 = "$($Script:pwdpath)\$($OutputFileName).html"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:FileName1 = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}

	$htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
	out-file -FilePath $Script:Filename1 -Force -InputObject $HTMLHead 4>$Null
}
#endregion

#region Iain's Word table functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Null -eq $Columns) -and ($Null -ne $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Null -ne $Columns) -and ($Null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end ElseIf
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$true, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$true, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $Null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $Null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] [int]$BackgroundColor = $Null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end ForEach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $true; }
				If($Italic) { $Cell.Range.Font.Italic = $true; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
				If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $true; }
					If($Italic) { $Cell.Range.Font.Italic = $true; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this function is called by the AddWordTable function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$true, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$true, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region general script functions
Function Get-ADTrustInfo
{
	Param ($TrustInfo)

	$propertyhash = @{
		TrustType = $NULL
		TrustAttribute = $NULL
		TrustDirection = $NULL
	}

	$TrustDirectionNumber = $TrustInfo.TrustDirection
	$TrustTypeNumber = $TrustInfo.TrustType
	$TrustAttributesNumber = $TrustInfo.TrustAttributes
	
	#http://msdn.microsoft.com/en-us/library/cc234293.aspx
	Switch ($TrustTypeNumber) 
	{ 
		1 { $propertyhash['TrustType'] = "Trust with a Windows domain not running Active Directory"; Break} 
		2 { $propertyhash['TrustType'] = "Trust with a Windows domain running Active Directory"; Break} 
		3 { $propertyhash['TrustType'] = "Trust with a non-Windows-compliant Kerberos distribution"; Break} 
		4 { $propertyhash['TrustType'] = "Trust with a DCE realm (not used)"; Break} 
		Default { $propertyhash['TrustType'] = "Invalid Trust Type of $($TrustTypeNumber)" ; Break}
	}
	
	$propertyhash['TrustAttribute'] = @()
	#$hextrustAttributesValue = '{0:X}' -f $trustAttributesNumber
	Switch ($trustAttributesNumber)
	{
		1 {$propertyhash['TrustAttribute'] += "Non-Transitive"}
		2 {$propertyhash['TrustAttribute'] += "Uplevel clients only"}
		4 {$propertyhash['TrustAttribute'] += "Quarantined Domain (External, SID Filtering)"}
		8 {$propertyhash['TrustAttribute'] += "Cross-Organizational Trust (Selective Authentication)"}
		16 {$propertyhash['TrustAttribute'] += "Interforest Trust"}
		32 {$propertyhash['TrustAttribute'] += "Intraforest Trust"}
		64 {$propertyhash['TrustAttribute'] += "MIT Trust using RC4 Encryption"}
		512 {$propertyhash['TrustAttribute'] += "Cross organization Trust no TGT delegation"}
	}
	
	Switch ($TrustDirectionNumber) 
	{ 
		0 { $propertyhash['TrustDirection'] = "Disabled"; Break} 
		1 { $propertyhash['TrustDirection'] = "Inbound"; Break} 
		2 { $propertyhash['TrustDirection'] = "Outbound"; Break} 
		3 { $propertyhash['TrustDirection'] = "Bidirectional"; Break} 
		Default { $propertyhash['TrustDirection'] = $TrustDirectionNumber ; Break}
	}
	
	New-Object -Type PSObject -property $propertyhash
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			If((Get-Member -Name $secondLevel -InputObject $object.$topLevel))
			{
				Return $True
			}
		}
	}
	Return $False
}

Function validObject( [object] $object, [string] $topLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			Return $True
		}
	}
	Return $False
}

Function AbortScript
{
	If($MSWord -or $PDF)
	{
		$Script:Word.quit()
		Write-Verbose "$(Get-Date): System Cleanup"
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
		If(Test-Path variable:global:word)
		{
			Remove-Variable -Name word -Scope Global
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Function BuildMultiColumnTable
{
	Param([Array]$xArray, [String]$xType)
	
	#divide by 0 bug reported 9-Apr-2014 by Lee Dehmer 
	#if security group name or OU name was longer than 60 characters it caused a divide by 0 error
	
	#added a second parameter to the function so the verbose message would say whether 
	#the function is processing servers, security groups or OUs.
	
	If(-not ($xArray -is [Array]))
	{
		$xArray = (,$xArray)
	}
	[int]$MaxLength = 0
	[int]$TmpLength = 0
	#remove 60 as a hard-coded value
	#60 is the max width the table can be when indented 36 points
	[int]$MaxTableWidth = 60
	ForEach($xName in $xArray)
	{
		$TmpLength = $xName.Length
		If($TmpLength -gt $MaxLength)
		{
			$MaxLength = $TmpLength
		}
	}
	$TableRange = $doc.Application.Selection.Range
	#removed hard-coded value of 60 and replace with MaxTableWidth variable
	[int]$Columns = [Math]::Floor($MaxTableWidth / $MaxLength)
	If($xArray.count -lt $Columns)
	{
		[int]$Rows = 1
		#not enough array items to fill columns so use array count
		$MaxCells  = $xArray.Count
		#reset column count so there are no empty columns
		$Columns   = $xArray.Count 
	}
	ElseIf($Columns -eq 0)
	{
		#divide by 0 bug if this condition is not handled
		#number was larger than $MaxTableWidth so there can only be one column
		#with one cell per row
		[int]$Rows = $xArray.count
		$Columns   = 1
		$MaxCells  = 1
	}
	Else
	{
		[int]$Rows = [Math]::Floor( ( $xArray.count + $Columns - 1 ) / $Columns)
		#more array items than columns so don't go past last column
		$MaxCells  = $Columns
	}
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$Table.Style = $Script:MyHash.Word_TableGrid
	
	$Table.Borders.InsideLineStyle = $wdLineStyleSingle
	$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
	[int]$xRow = 1
	[int]$ArrayItem = 0
	While($xRow -le $Rows)
	{
		For($xCell=1; $xCell -le $MaxCells; $xCell++)
		{
			$Table.Cell($xRow,$xCell).Range.Text = $xArray[$ArrayItem]
			$ArrayItem++
		}
		$xRow++
	}
	$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
	$Table.AutoFitBehavior($wdAutoFitContent)

	FindWordDocumentEnd
	$TableRange = $Null
	$Table = $Null
	$xArray = $Null
}

Function UserIsaDomainAdmin
{
	#function adapted from sample code provided by Thomas Vuylsteke
	$IsDA = $False
	$name = $env:username
	Write-Verbose "$(Get-Date): TokenGroups - Checking groups for $name"

	$root = [ADSI]""
	$filter = "(sAMAccountName=$name)"
	$props = @("distinguishedName")
	$Searcher = new-Object System.DirectoryServices.DirectorySearcher($root,$filter,$props)
	$account = $Searcher.FindOne().properties.distinguishedname

	$user = [ADSI]"LDAP://$Account"
	$user.GetInfoEx(@("tokengroups"),0)
	$groups = $user.Get("tokengroups")

	$domainAdminsSID = New-Object System.Security.Principal.SecurityIdentifier (((Get-ADDomain -Server $ADForest -EA 0).DomainSid).Value+"-512") 

	ForEach($group in $groups)
	{     
		$ID = New-Object System.Security.Principal.SecurityIdentifier($group,0)       
		If($ID.CompareTo($domainAdminsSID) -eq 0)
		{
			$IsDA = $True
			Break
		}     
	}

	Return $IsDA
}

Function ElevatedSession
{
	$currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

	If($currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator ))
	{
		Write-Verbose "$(Get-Date): This is an elevated PowerShell session"
		Return $True
	}
	Else
	{
		Write-Host "" -Foreground White
		Write-Host "$(Get-Date): This is NOT an elevated PowerShell session" -Foreground White
		Write-Host "" -Foreground White
		Return $False
	}
}

Function Get-ComputerCountByOS
{
	<#
	This function is provided by Jeremy Saunders and used with his permission

	http://www.jhouseconsulting.com/2014/06/22/script-to-create-an-overview-of-all-computer-objects-in-a-domain-1385

	Jeremy sent me version 1.8 of his script to use as the basis for this function
	
	This function will provide an overview and count of all computer objects in a
	domain based on Operating System and Service Pack. It helps an organisation
	to understand the number of stale and active computers against the different
	types of operating systems deployed in their environment.

	Computer objects are filtered into 4 categories:
	1) Windows Servers
	2) Windows Workstations
	3) Other non-Windows (Linux, Mac, etc)
	4) Windows Cluster Name Objects (CNOs) and Virtual Computer Objects (VCOs)

	A Stale object is derived from 2 values ANDed together:
	1) PasswordLastChanged > $MaxPasswordLastChanged days ago
	2) LastLogonDate > $MaxLastLogonDate days ago

	By default the script variable for $MaxPasswordLastChanged is set to 90 and
	the variable for $MaxLastLogonDate is set to 30. These can easily be adjusted
	to suite your definition of a stale object.

	The Active objects column is calculated by subtracting the Enabled_Stale
	value from the Enabled value. This gives us an accurate number of active
	objects against each Operating System.

	To help provide a high level overview of the computer object landscape, we
	calculate the number of stale objects of enabled and disabled objects
	separately. Disabled objects are often ignored, but it's pointless leaving
	old disabled computer objects in the domain.

	For viewing purposes we sort the output by Operating System and not count.

	You may notice a question mark (?) in some of the OperatingSystem strings.
	This is a representation of each Double-Byte character that was unable to
	be translated. Refer to Microsoft KB829856 for an explanation.

	Computers change their passwword if and when they feel like it. The domain
	doesn't initiate the change. It is controlled by three values under the
	following registry key:
	- HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters
	- DisablePasswordChange
	- MaximumPasswordAge
	- RefusePasswordChange
	If these values are not present, the default value of 30 days will be used.

	Non Windows operating systems and appliances vary in the way they manage
	their password change, with some not doing it at all. It's best to do some
	research before drawing any conclusions about the validity of these objects.
	i.e. Don't just go and delete them because my script says they are stale.

	Be aware that a cluster updates the lastLogonTimeStamp of the CNO/VNO when
	it brings a clustered network name resource online. So it could be running
	for months without an update to the lastLogonTimeStamp attribute.

	Syntax examples:

	- To execute the script in the current Domain:
	Get-ComputerCountByOS.ps1

	- To execute the script against a trusted Domain:
	Get-ComputerCountByOS.ps1 -TrustedDomain mydemosthatrock.com

	Script Name: Get-ComputerCountByOS.ps1
	Release: 1.6
	Written by Jeremy@jhouseconsulting.com 20th May 2012
	Modified by Jeremy@jhouseconsulting.com 4th December 2015

	#>

	#-------------------------------------------------------------
	param([String]$TrustedDomain)

	#-------------------------------------------------------------

	# Set this to true to include service pack level. This makes the
	# output more ganular, as the counts are then based on Operating
	# System + Service Pack.
	$OperatingSystemIncludesServicePack = $True

	# Set this to the maximum value in number of days when the computer
	# password last changed. Do not go beyond 90 days.
	$MaxPasswordLastChanged = 90

	# Set this to the maximum value in number of days when the computer
	# last logged onto the domain.
	$MaxLastLogonDate = 30

	#-------------------------------------------------------------

	$TotalComputersProcessed = 0
	$ComputerCount = 0
	$TotalStaleObjects = 0
	$TotalEnabledStaleObjects = 0
	$TotalEnabledObjects = 0
	$TotalDisabledObjects = 0
	$TotalDisabledStaleObjects = 0
	$AllComputerObjects = New-Object System.Collections.ArrayList
	$WindowsServerObjects = New-Object System.Collections.ArrayList
	$WindowsWorkstationObjects = New-Object System.Collections.ArrayList
	$NonWindowsComputerObjects = New-Object System.Collections.ArrayList
	$CNOandVCOObjects = New-Object System.Collections.ArrayList
	$ComputersHashTable = @{}

	$context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$TrustedDomain)
	Try 
	{
		$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($context)
	}
	Catch [exception] 
	{
		Write-Error $_.Exception.Message
		Exit
	}

	# Get AD Distinguished Name
	$DomainDistinguishedName = $Domain.GetDirectoryEntry() | Select-Object -ExpandProperty DistinguishedName  

	$ADSearchBase = $DomainDistinguishedName

	Write-Verbose "$(Get-Date): `t`tGathering computer misc data"

	# Create an LDAP search for all computer objects
	$ADFilter = "(objectCategory=computer)"

	# There is a known bug in PowerShell requiring the DirectorySearcher
	# properties to be in lower case for reliability.
	$ADPropertyList = @("distinguishedname","name","operatingsystem","operatingsystemversion", `
	"operatingsystemservicepack", "description", "info", "useraccountcontrol", `
	"pwdlastset","lastlogontimestamp","whencreated","serviceprincipalname")

	$ADScope = "SUBTREE"
	$ADPageSize = 1000
	$ADSearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($ADSearchBase)") 
	$ADSearcher = New-Object System.DirectoryServices.DirectorySearcher
	$ADSearcher.SearchRoot = $ADSearchRoot
	$ADSearcher.PageSize = $ADPageSize 
	$ADSearcher.Filter = $ADFilter 
	$ADSearcher.SearchScope = $ADScope
	If($ADPropertyList) 
	{
		ForEach($ADProperty in $ADPropertyList) 
		{
			[Void]$ADSearcher.PropertiesToLoad.Add($ADProperty)
		}
	}
	Try 
	{
		Write-Verbose "Please be patient whilst the script retrieves all computer objects and specified attributes..."
		$colResults = $ADSearcher.Findall()
		# Dispose of the search and results properly to avoid a memory leak
		$ADSearcher.Dispose()
		$ComputerCount = $colResults.Count
	}
	Catch 
	{
		$ComputerCount = 0
		Write-Warning "The $ADSearchBase structure cannot be found!"
	}

	If($ComputerCount -ne 0) 
	{
		Write-Verbose "Processing $ComputerCount computer objects in the $domain Domain..."
		ForEach($objResult in $colResults) 
		{
			$Name = $objResult.Properties.name[0]
			$DistinguishedName = $objResult.Properties.distinguishedname[0]

			$ParentDN = $DistinguishedName -split '(?<![\\]),'
			$ParentDN = $ParentDN[1..$($ParentDN.Count-1)] -join ','

			Try 
			{
				If(($objResult.Properties.operatingsystem | Measure-Object).Count -gt 0) 
				{
					$OperatingSystem = $objResult.Properties.operatingsystem[0]
				} 
				Else 
				{
					$OperatingSystem = "Undefined"
				}
			}
			Catch 
			{
				$OperatingSystem = "Undefined"
			}
			Try 
			{
				If(($objResult.Properties.operatingsystemversion | Measure-Object).Count -gt 0) 
				{
					$OperatingSystemVersion = $objResult.Properties.operatingsystemversion[0]
				} 
				Else 
				{
					$OperatingSystemVersion = ""
				}
			}
			Catch 
			{
				$OperatingSystemVersion = ""
			}
			Try 
			{
				If(($objResult.Properties.operatingsystemservicepack | Measure-Object).Count -gt 0)
				{					
					$OperatingSystemServicePack = $objResult.Properties.operatingsystemservicepack[0]
				}
				Else 
				{
					$OperatingSystemServicePack = ""
				}
			}
			Catch 
			{
				$OperatingSystemServicePack = ""
			}
			Try 
			{
				If(($objResult.Properties.description | Measure-Object).Count -gt 0) 
				{
					$Description = $objResult.Properties.description[0]
				} 
				Else 
				{
					$Description = ""
				}
			}
			Catch 
			{
				$Description = ""
			}
			$PasswordTooOld = $False
			$PasswordLastSet = [System.DateTime]::FromFileTime($objResult.Properties.pwdlastset[0])
			If($PasswordLastSet -lt (Get-Date).AddDays(-$MaxPasswordLastChanged)) 
			{
				$PasswordTooOld = $True
			}
			$HasNotRecentlyLoggedOn = $False
			Try 
			{
				If(($objResult.Properties.lastlogontimestamp | Measure-Object).Count -gt 0) 
				{
					$LastLogonTimeStamp = $objResult.Properties.lastlogontimestamp[0]
					$LastLogon = [System.DateTime]::FromFileTime($LastLogonTimeStamp)
					If($LastLogon -le (Get-Date).AddDays(-$MaxLastLogonDate)) 
					{
						$HasNotRecentlyLoggedOn = $True
					}
					If($LastLogon -match "1/01/1601") 
					{
						$LastLogon = "Never logged on before"
					}
				} 
				Else 
				{
					$LastLogon = "Never logged on before"
				}
			}
			Catch 
			{
				$LastLogon = "Never logged on before"
			}
			$WhenCreated = $objResult.Properties.whencreated[0]

			# If it's never logged on before and was created more than $MaxLastLogonDate days
			# ago, set the $HasNotRecentlyLoggedOn variable to True.
			# An example of this would be if you prestaged the account but never ended up using
			# it.
			If($lastLogon -eq "Never logged on before") 
			{
			  If($whencreated -le (Get-Date).AddDays(-$MaxLastLogonDate)) 
			  {
				$HasNotRecentlyLoggedOn = $True
			  }
			}

			# Check if it's a stale object.
			$IsStale = $False
			If($PasswordTooOld -eq $True -AND $HasNotRecentlyLoggedOn -eq $True) 
			{
				$IsStale = $True
			}
			Try 
			{
				$ServicePrincipalName = $objResult.Properties.serviceprincipalname
			}
			Catch 
			{
				$ServicePrincipalName = ""
			}
			$UserAccountControl = $objResult.Properties.useraccountcontrol[0]
			$Enabled = $True
			Switch($UserAccountControl)
			{
				{($UserAccountControl -bor 0x0002) -eq $UserAccountControl} 
				{
					$Enabled = $False
				}
			}
			Try 
			{
				If(($objResult.Properties.info | Measure-Object).Count -gt 0) 
				{
					$notes = $objResult.Properties.info[0]
					$notes = $notes -replace "`r`n", "|"
				} 
				Else 
				{
					$notes = ""
				}
			}
			Catch 
			{
				$notes = ""
			}
			If($IsStale) 
			{
				$TotalStaleObjects = $TotalStaleObjects + 1
			}
			If($Enabled) 
			{
				$TotalEnabledObjects = $TotalEnabledObjects + 1
			}
			If($Enabled -eq $False) 
			{
				$TotalDisabledObjects = $TotalDisabledObjects + 1
			}
			If($IsStale -AND $Enabled) 
			{
				$TotalEnabledStaleObjects = $TotalEnabledStaleObjects + 1
			}
			If($IsStale -AND $Enabled -eq $False) 
			{
				$TotalDisabledStaleObjects = $TotalDisabledStaleObjects + 1
			}

			$obj = New-Object -TypeName PSObject
			$obj | Add-Member -MemberType NoteProperty -Name "Name" -value $Name
			$obj | Add-Member -MemberType NoteProperty -Name "ParentOU" -value $ParentDN
			$obj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -value $OperatingSystem
			$obj | Add-Member -MemberType NoteProperty -Name "Version" -value $OperatingSystemVersion
			$obj | Add-Member -MemberType NoteProperty -Name "ServicePack" -value $OperatingSystemServicePack
			$obj | Add-Member -MemberType NoteProperty -Name "Description" -value $Description

			If(!($ServicePrincipalName -match 'MSClusterVirtualServer')) 
			{
				If($OperatingSystem -match 'windows' -AND $OperatingSystem -match 'server') 
				{
					$Category = "Server"
				}
				If($OperatingSystem -match 'windows' -AND !($OperatingSystem -match 'server')) 
				{
					$Category = "Workstation"
				}
				If(!($OperatingSystem -match 'windows')) 
				{
					$Category = "Other"
				}
			} 
			Else 
			{
				$Category = "CNO or VCO"
				$OperatingSystem = $OperatingSystem + " - " + $Category
			}
			If($Category -eq "") 
			{
				$Category = "Undefined"
			}

			$obj | Add-Member -MemberType NoteProperty -Name "Category" -value $Category
			$obj | Add-Member -MemberType NoteProperty -Name "PasswordLastSet" -value $PasswordLastSet
			$obj | Add-Member -MemberType NoteProperty -Name "LastLogon" -value $LastLogon
			$obj | Add-Member -MemberType NoteProperty -Name "Enabled" -value $Enabled
			$obj | Add-Member -MemberType NoteProperty -Name "IsStale" -value $IsStale
			$obj | Add-Member -MemberType NoteProperty -Name "WhenCreated" -value $WhenCreated
			$obj | Add-Member -MemberType NoteProperty -Name "Notes" -value $notes

			$AllComputerObjects.Add($obj) > $Null

			Switch($Category)
			{
				"Server"		{$WindowsServerObjects.Add($obj) > $Null; break}
				"Workstation"	{$WindowsWorkstationObjects.Add($obj) > $Null; break}
				"Other"			{$NonWindowsComputerObjects.Add($obj) > $Null; break}
				"CNO or VCO"	{$CNOandVCOObjects.Add($obj) > $Null; break}
				"Undefined"		{$NonWindowsComputerObjects.Add($obj) > $Null; break}
			}
			$obj = New-Object -TypeName PSObject
			$obj | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -value $OperatingSystem
			If($OperatingSystemIncludesServicePack -eq $False) 
			{
				$FullOperatingSystem = $OperatingSystem
			} 
			Else 
			{
				$FullOperatingSystem = $OperatingSystem + " " + $OperatingSystemServicePack
				$obj | Add-Member -MemberType NoteProperty -Name "ServicePack" -value $OperatingSystemServicePack
			}
			$obj | Add-Member -MemberType NoteProperty -Name "Category" -value $Category

			# Create a hashtable to capture a count of each Operating System
			If(!($ComputersHashTable.ContainsKey($FullOperatingSystem))) 
			{
				$TotalCount = 1
				$StaleCount = 0
				$EnabledStaleCount = 0
				$DisabledStaleCount = 0
				If($IsStale -eq $True) 
				{ 
					$StaleCount = 1 
				}
				If($Enabled -eq $True) 
				{
					$EnabledCount = 1
					$DisabledCount = 0
					If($IsStale -eq $True) 
					{ 
						$EnabledStaleCount = 1 
					}
				}
				If($Enabled -eq $False) 
				{
					$DisabledCount = 1
					$EnabledCount = 0
					If($IsStale -eq $True) 
					{ 
						$DisabledStaleCount = 1 
					}
				}
				$obj | Add-Member -MemberType NoteProperty -Name "Total" -value $TotalCount
				$obj | Add-Member -MemberType NoteProperty -Name "Stale" -value $StaleCount
				$obj | Add-Member -MemberType NoteProperty -Name "Enabled" -value $EnabledCount
				$obj | Add-Member -MemberType NoteProperty -Name "Enabled_Stale" -value $EnabledStaleCount
				$obj | Add-Member -MemberType NoteProperty -Name "Active" -value ($EnabledCount - $EnabledStaleCount)
				$obj | Add-Member -MemberType NoteProperty -Name "Disabled" -value $DisabledCount
				$obj | Add-Member -MemberType NoteProperty -Name "Disabled_Stale" -value $DisabledStaleCount
				$ComputersHashTable = $ComputersHashTable + @{$FullOperatingSystem = $obj}
			} 
			Else 
			{
				$value = $ComputersHashTable.Get_Item($FullOperatingSystem)
				$TotalCount = $value.Total + 1
				$StaleCount = $value.Stale
				$EnabledStaleCount = $value.Enabled_Stale
				$DisabledStaleCount = $value.Disabled_Stale
				If($IsStale -eq $True) 
				{ 
					$StaleCount = $value.Stale + 1 
				}
				If($Enabled -eq $True) 
				{
					$EnabledCount = $value.Enabled + 1
					$DisabledCount = $value.Disabled
					If($IsStale -eq $True) 
					{ 
						$EnabledStaleCount = $value.Enabled_Stale + 1 
					}
				}
				If($Enabled -eq $False) 
				{ 
					$DisabledCount = $value.Disabled + 1
					$EnabledCount = $value.Enabled
					If($IsStale -eq $True) 
					{ 
						$DisabledStaleCount = $value.Disabled_Stale + 1 
					}
				}
				$obj | Add-Member -MemberType NoteProperty -Name "Total" -value $TotalCount
				$obj | Add-Member -MemberType NoteProperty -Name "Stale" -value $StaleCount
				$obj | Add-Member -MemberType NoteProperty -Name "Enabled" -value $EnabledCount
				$obj | Add-Member -MemberType NoteProperty -Name "Enabled_Stale" -value $EnabledStaleCount
				$obj | Add-Member -MemberType NoteProperty -Name "Active" -value ($EnabledCount - $EnabledStaleCount)
				$obj | Add-Member -MemberType NoteProperty -Name "Disabled" -value $DisabledCount
				$obj | Add-Member -MemberType NoteProperty -Name "Disabled_Stale" -value $DisabledStaleCount
				$ComputersHashTable.Set_Item($FullOperatingSystem,$obj)
			} # end if
			$TotalComputersProcessed ++
		}

		# Dispose of the search and results properly to avoid a memory leak
		$colResults.Dispose()

		$Output = $ComputersHashTable.values | ForEach-Object {$_ } | ForEach-Object {$_ } | Sort-Object OperatingSystem -descending
		
		If($MSWORD -or $PDF)
		{
			WriteWordLine 3 0 "Computer Operating Systems"
		}
		ElseIf($Text)
		{
			Line 0 "Computer Operating Systems"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 3 0 "Computer Operating Systems"
		}

		ForEach($Item in $Output)
		{
			If($MSWORD -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Operating System"; Value = $Item.OperatingSystem; }
				$ScriptInformation += @{ Data = "Service Pack"; Value = $Item.ServicePack; }
				$ScriptInformation += @{ Data = "Category"; Value = $Item.Category; }
				$ScriptInformation += @{ Data = "Total"; Value = $Item.Total; }
				$ScriptInformation += @{ Data = "Stale"; Value = $Item.Stale; }
				$ScriptInformation += @{ Data = "Enabled"; Value = $Item.Enabled; }
				$ScriptInformation += @{ Data = "Enabled/Stale"; Value = $Item.Enabled_Stale; }
				$ScriptInformation += @{ Data = "Active"; Value = $Item.Active; }
				$ScriptInformation += @{ Data = "Disabled"; Value = $Item.Disabled; }
				$ScriptInformation += @{ Data = "Disabled/Stale"; Value = $Item.Disabled_Stale; }
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				## IB - set column widths without recursion
				$Table.Columns.Item(1).Width = 100;
				$Table.Columns.Item(2).Width = 200;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 1 "Operating System`t: " $Item.OperatingSystem
				Line 1 "Service Pack`t`t: " $Item.ServicePack
				Line 1 "Category`t`t: " $Item.Category
				Line 1 "Total`t`t`t: " $Item.Total
				Line 1 "Stale`t`t`t: " $Item.Stale
				Line 1 "Enabled`t`t`t: " $Item.Enabled
				Line 1 "Enabled/Stale`t`t: " $Item.Enabled_Stale
				Line 1 "Active`t`t`t: " $Item.Active
				Line 1 "Disabled`t`t: " $Item.Disabled
				Line 1 "Disabled/Stale`t`t: " $Item.Disabled_Stale
				Line 0 ""
			}
			ElseIf($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Operating System",($htmlsilver -bor $htmlbold),$Item.OperatingSystem,$htmlwhite)
				$rowdata += @(,('Service Pack',($htmlsilver -bor $htmlbold),$Item.ServicePack,$htmlwhite))
				$rowdata += @(,('Category',($htmlsilver -bor $htmlbold),$Item.Category,$htmlwhite))
				$rowdata += @(,('Total',($htmlsilver -bor $htmlbold),$Item.Total,$htmlwhite))
				$rowdata += @(,('Stale',($htmlsilver -bor $htmlbold),$Item.Stale,$htmlwhite))
				$rowdata += @(,('Enabled',($htmlsilver -bor $htmlbold),$Item.Enabled,$htmlwhite))
				$rowdata += @(,('Enabled/Stale',($htmlsilver -bor $htmlbold),$Item.Enabled_Stale,$htmlwhite))
				$rowdata += @(,('Active',($htmlsilver -bor $htmlbold),$Item.Active,$htmlwhite))
				$rowdata += @(,('Disabled',($htmlsilver -bor $htmlbold),$Item.Disabled,$htmlwhite))
				$rowdata += @(,('Disabled/Stale',($htmlsilver -bor $htmlbold),$Item.Disabled_Stale,$htmlwhite))
				
				$msg = ""
				$columnWidths = @("100","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "300"
				WriteHTMLLine 0 0 " "
			}
		}

		$percent = "{0:P}" -f ($TotalStaleObjects/$ComputerCount)
		
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "A breakdown of the $ComputerCount Computer Objects in the $domain Domain"
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Total Computer Objects"; Value = $ComputerCount; }
			$ScriptInformation += @{ Data = "Total Stale Computer Objects (count)"; Value = $TotalStaleObjects; }
			$ScriptInformation += @{ Data = "Total Stale Computer Objects (percent)"; Value = $percent; }
			$ScriptInformation += @{ Data = "Total Enabled Computer Objects"; Value = $TotalEnabledObjects; }
			$ScriptInformation += @{ Data = "Total Enabled Stale Computer Objects"; Value = $TotalEnabledStaleObjects; }
			$ScriptInformation += @{ Data = "Total Active Computer Objects"; Value = $($TotalEnabledObjects - $TotalEnabledStaleObjects); }
			$ScriptInformation += @{ Data = "Total Disabled Computer Objects"; Value = $TotalDisabledObjects; }
			$ScriptInformation += @{ Data = "Total Disabled Stale Computer Objects"; Value = $TotalDisabledStaleObjects; }
			
			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 0 "A breakdown of the $ComputerCount Computer Objects in the $domain Domain"
			
			Line 1 "Total Computer Objects`t`t`t: " $ComputerCount
			Line 1 "Total Stale Computer Objects (count)`t: " $TotalStaleObjects
			Line 1 "Total Stale Computer Objects (percent)`t: " $percent
			Line 1 "Total Enabled Computer Objects`t`t: " $TotalEnabledObjects
			Line 1 "Total Enabled Stale Computer Objects`t: " $TotalEnabledStaleObjects
			Line 1 "Total Active Computer Objects`t`t: " $($TotalEnabledObjects - $TotalEnabledStaleObjects)
			Line 1 "Total Disabled Computer Objects`t`t: " $TotalDisabledObjects
			Line 1 "Total Disabled Stale Computer Objects`t: " $TotalDisabledStaleObjects
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata = @()
			$columnHeaders = @("Total Computer Objects",($htmlsilver -bor $htmlbold),$ComputerCount.ToString(),$htmlwhite)
			$rowdata += @(,('Total Stale Computer Objects (count)',($htmlsilver -bor $htmlbold),$TotalStaleObjects,$htmlwhite))
			$rowdata += @(,('Total Stale Computer Objects (percent)',($htmlsilver -bor $htmlbold),$percent,$htmlwhite))
			$rowdata += @(,('Total Enabled Computer Objects',($htmlsilver -bor $htmlbold),$TotalEnabledObjects,$htmlwhite))
			$rowdata += @(,('Total Enabled Stale Computer Objects',($htmlsilver -bor $htmlbold),$TotalEnabledStaleObjects,$htmlwhite))
			$rowdata += @(,('Total Active Computer Objects',($htmlsilver -bor $htmlbold),$($TotalEnabledObjects - $TotalEnabledStaleObjects),$htmlwhite))
			$rowdata += @(,('Total Disabled Computer Objects',($htmlsilver -bor $htmlbold),$TotalDisabledObjects,$htmlwhite))
			$rowdata += @(,('Total Disabled Stale Computer Objects',($htmlsilver -bor $htmlbold),$TotalDisabledStaleObjects,$htmlwhite))
			
			$msg = "A breakdown of the $ComputerCount Computer Objects in the $domain Domain"
			$columnWidths = @("250","50")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "300"
			WriteHTMLLine 0 0 " "
		}

		#Notes:
		# - Computer objects are filtered into 4 categories:
		#	 - Windows Servers
		#	 - Windows Workstations
		#	 - Other non-Windows (Linux, Mac, etc)
		#	 - CNO or VCO (Windows Cluster Name Objects and Virtual Computer Objects)
		# - A Stale object is derived from 2 values ANDed together:
		#     PasswordLastChanged  > $MaxPasswordLastChanged days ago
		#     AND
		#     LastLogonDate > $MaxLastLogonDate days ago
		# - If it's never logged on before and was created more than $MaxLastLogonDate days ago, set the
		#   HasNotRecentlyLoggedOn variable to True. This will also be used to help determine if
		#   it's a stale account. An example of this would be if you prestaged the account but
		#   never ended up using it.
		# - The Active objects column is calculated by subtracting the Enabled_Stale value from
		#   the Enabled value. This gives us an accurate number of active objects against each
		#   Operating System.
		# - To help provide a high level overview of the computer object landscape we calculate
		#   the number of stale objects of enabled and disabled objects separately.
		#   Disabled objects are often ignored, but it's pointless leaving old disabled computer
		#   objects in the domain.
		# - For viewing purposes we sort the output by Operating System and not count.
		# - You may notice a question mark (?) in some of the OperatingSystem strings. This is a
		#   representation of each Double-Byte character that was unable to be translated. Refer
		#   to Microsoft KB829856 for an explanation.
		# - Be aware that a cluster updates the lastLogonTimeStamp of the CNO/VNO when it brings
		#   a clustered network name resource online. So it could be running for months without
		#   an update to the lastLogonTimeStamp attribute.
		# - When a VMware ESX host has been added to Active Directory its associated computer
		#   object will appear as an Operating System of unknown, version: unknown, and service
		#   pack: Likewise Identity 5.3.0. The lsassd.conf manages the machine password expiration
		#   lifespan, which by default is set to 30 days.
		# - When a Riverbed SteelHead has been added to Active Directory the password set on the
		#   Computer Object will never change by default. This must be enabled on the command
		#   line using the 'domain settings pwd-refresh-int <number of days>' command. The
		#   lastLogon timestamp is only updated when the SteelHead appliance is restarted.
	}
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): AddDateTime     : $AddDateTime"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Company Name    : $Script:CoName"
	}
	Write-Verbose "$(Get-Date): ComputerName    : $ComputerName"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Company Address : $CompanyAddress"
		Write-Verbose "$(Get-Date): Company Email   : $CompanyEmail"
		Write-Verbose "$(Get-Date): Company Fax     : $CompanyFax"
		Write-Verbose "$(Get-Date): Company Phone   : $CompanyPhone"
		Write-Verbose "$(Get-Date): Cover Page      : $CoverPage"
	}
	Write-Verbose "$(Get-Date): DCDNSInfo       : $DCDNSInfo"
	Write-Verbose "$(Get-Date): Dev             : $Dev"
	If($Dev)
	{
		Write-Verbose "$(Get-Date): DevErrorFile    : $Script:DevErrorFile"
	}
	Write-Verbose "$(Get-Date): Domain Name     : $ADDomain"
	Write-Verbose "$(Get-Date): Elevated        : $Script:Elevated"
	Write-Verbose "$(Get-Date): Filename1       : $Script:filename1"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2       : $Script:filename2"
	}
	Write-Verbose "$(Get-Date): Folder          : $Folder"
	Write-Verbose "$(Get-Date): Forest Name     : $ADForest"
	Write-Verbose "$(Get-Date): From            : $From"
	Write-Verbose "$(Get-Date): GPOInheritance  : $GPOInheritance"
	Write-Verbose "$(Get-Date): HW Inventory    : $Hardware"
	Write-Verbose "$(Get-Date): IncludeUserInfo : $IncludeUserInfo"
	Write-Verbose "$(Get-Date): Log             : $($Log)"
	Write-Verbose "$(Get-Date): MaxDetail       : $MaxDetails"
	Write-Verbose "$(Get-Date): Save As HTML    : $HTML"
	Write-Verbose "$(Get-Date): Save As PDF     : $PDF"
	Write-Verbose "$(Get-Date): Save As TEXT    : $TEXT"
	Write-Verbose "$(Get-Date): Save As WORD    : $MSWORD"
	Write-Verbose "$(Get-Date): ScriptInfo      : $ScriptInfo"
	Write-Verbose "$(Get-Date): Section         : $Section"
	Write-Verbose "$(Get-Date): Services        : $Services"
	Write-Verbose "$(Get-Date): Smtp Port       : $SmtpPort"
	Write-Verbose "$(Get-Date): Smtp Server     : $SmtpServer"
	Write-Verbose "$(Get-Date): Title           : $Script:Title"
	Write-Verbose "$(Get-Date): To              : $To"
	Write-Verbose "$(Get-Date): Use SSL         : $UseSSL"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): User Name       : $UserName"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): OS Detected     : $Script:RunningOS"
	Write-Verbose "$(Get-Date): PoSH version    : $($Host.Version)"
	Write-Verbose "$(Get-Date): PSCulture       : $PSCulture"
	Write-Verbose "$(Get-Date): PSUICulture     : $PSUICulture"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Word language   : $Script:WordLanguageValue"
		Write-Verbose "$(Get-Date): Word version    : $Script:WordProduct"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start    : $Script:StartTime"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date): Saving DOCX file"
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Verbose "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:FileName1, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	If($PDF)
	{
		[int]$cnt = 0
		While(Test-Path $Script:FileName1)
		{
			$cnt++
			If($cnt -gt 1)
			{
				Write-Verbose "$(Get-Date): Waiting another 10 seconds to allow Word to fully close (try # $($cnt))"
				Start-Sleep -Seconds 10
				$Script:Word.Quit()
				If($cnt -gt 2)
				{
					#kill the winword process

					#find out our session (usually "1" except on TS/RDC or Citrix)
					$SessionID = (Get-Process -PID $PID).SessionId
					
					#Find out if winword is running in our session
					$wordprocess = ((Get-Process 'WinWord' -ea 0)|Where-Object {$_.SessionId -eq $SessionID}).Id
					If($wordprocess -gt 0)
					{
						Write-Verbose "$(Get-Date): Attempting to stop WinWord process # $($wordprocess)"
						Stop-Process $wordprocess -EA 0
					}
				}
			}
			Write-Verbose "$(Get-Date): Attempting to delete $($Script:FileName1) since only $($Script:FileName2) is needed (try # $($cnt))"
			Remove-Item $Script:FileName1 -EA 0 4>$Null
		}
	}
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global 4>$Null
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword is running in our session
	$wordprocess = $Null
	$wordprocess = ((Get-Process 'WinWord' -ea 0)|Where-Object {$_.SessionId -eq $SessionID}).Id
	If($null -ne $wordprocess -and $wordprocess -gt 0)
	{
		Write-Verbose "$(Get-Date): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess)"
		Stop-Process $wordprocess -EA 0
	}
}

Function SaveandCloseTextDocument
{
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}

	Write-Output $global:Output.ToString() | Out-File $Script:Filename1 4>$Null
}

Function SaveandCloseHTMLDocument
{
	Out-File -FilePath $Script:FileName1 -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFileName1andFileName2
{
	Param([string]$OutputFileName)
	
	#set $filename1 and $filename2 with no file extension
	If($AddDateTime)
	{
		[string]$Script:FileName1 = "$($Script:pwdpath)\$($OutputFileName)"
		If($PDF)
		{
			[string]$Script:FileName2 = "$($Script:pwdpath)\$($OutputFileName)"
		}
	}

	If($MSWord -or $PDF)
	{
		CheckWordPreReq

		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($Script:pwdpath)\$($OutputFileName).docx"
			If($PDF)
			{
				[string]$Script:FileName2 = "$($Script:pwdpath)\$($OutputFileName).pdf"
			}
		}

		SetupWord
	}
	ElseIf($Text)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($Script:pwdpath)\$($OutputFileName).txt"
		}
		ShowScriptOptions
	}
	ElseIf($HTML)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($Script:pwdpath)\$($OutputFileName).html"
		}
		SetupHTML
		ShowScriptOptions
	}
}

#endregion

#Script begins

#region script setup function
Function ProcessScriptSetup
{
	#If hardware inventory or services are requested, make sure user is running the script with Domain Admin rights
	Write-Verbose "$(Get-Date): `tTesting to see if $env:username has Domain Admin rights"
	$Script:DARights = $False
	$Script:Elevated = $False
	
	$AmIReallyDA = UserIsaDomainAdmin
	If($AmIReallyDA -eq $True)
	{
		#user has Domain Admin rights
		If($ADDomain -ne "")
		{
			Write-Verbose "$(Get-Date): $env:username has Domain Admin rights in the $ADDomain Domain"
		}
		Else
		{
			Write-Verbose "$(Get-Date): $env:username has Domain Admin rights in the $ADForest Forest"
		}
		$Script:DARights = $True
	}
	Else
	{
		#user does nto have Domain Admin rights
		If($ADDomain -ne "")
		{
			Write-Verbose "$(Get-Date): $env:username does not have Domain Admin rights in the $ADDomain Domain"
		}
		Else
		{
			Write-Verbose "$(Get-Date): $env:username does not have Domain Admin rights in the $ADForest Forest"
		}
	}
	
	$Script:Elevated = ElevatedSession
	
	If($Hardware -or $Services -or $DCDNSINFO)
	{
		If($Hardware -and -not $Services)
		{
			Write-Verbose "$(Get-Date): Hardware inventory requested"
		}
		ElseIf($Services -and -not $Hardware)
		{
			Write-Verbose "$(Get-Date): Services requested"
		}
		ElseIf($Hardware -and $Services)
		{
			Write-Verbose "$(Get-Date): Hardware inventory and Services requested"
		}
		
		If($DCDNSINFO)
		{
			Write-Verbose "$(Get-Date): Domain Controller DNS configuration information requested"
		}

		If($Script:DARights -eq $False)
		{
			#user does not have Domain Admin rights
			If($Hardware -and -not $Services)
			{
				#don't abort script, set $hardware to false
				Write-Warning "`n`n`t`tHardware inventory was requested but $($env:username) does not have Domain Admin rights."
				Write-Warning "`n`n`t`tHardware inventory option will be turned off."
				$Script:Hardware = $False
			}
			ElseIf($Services -and -not $Hardware)
			{
				#don't abort script, set $services to false
				Write-Warning "`n`n`t`tServices were requested but $($env:username) does not have Domain Admin rights."
				Write-Warning "`n`n`t`tServices option will be turned off."
				$Script:Services = $False
			}
			ElseIf($Hardware -and $Services)
			{
				#don't abort script, set $hardware and $services to false
				Write-Warning "`n`n`t`tHardware inventory and Services were requested but $($env:username) does not have Domain Admin rights."
				Write-Warning "`n`n`t`tHardware inventory and Services options will be turned off."
				$Script:Hardware = $False
				$Script:Services = $False
			}

			If($DCDNSINFO)
			{
				#don't abort script, set $DCDNSINFO to false
				Write-Warning "`n`n`t`tDCDNSINFO information was requested but $($env:username) does not have Domain Admin rights."
				Write-Warning "`n`n`t`tDCDNSINFO option will be turned off."
				$Script:DCDNSINFO = $False
			}
		}
		
		If( ($Hardware -or $Services) -and -not $Script:Elevated )
		{
			Write-Host "Warning: " -Foreground White
			Write-Host "Warning: Hardware inventory or Services were requested but this is not an elevated PowerShell session." -Foreground White
			Write-Host "Warning: Hardware inventory and Services options will be turned off." -Foreground White
			Write-Host "Warning: To obtain Hardware inventory and Services data, please run the script from an elevated PowerShell session." -Foreground White
			Write-Host "Warning: " -Foreground White
			$Script:Hardware = $False
			$Script:Services = $False
		}

		If( $DCDNSINFO -and -not $Script:Elevated )
		{
			Write-Host "Warning: " -Foreground White
			Write-Host "Warning: Domain Controller DNS information was requested but this is not an elevated PowerShell session." -Foreground White
			Write-Host "Warning: DCDNSINFO option will be turned off." -Foreground White
			Write-Host "Warning: To obtain DCDNSINFO data, please run the script from an elevated PowerShell session." -Foreground White
			Write-Host "Warning: " -Foreground White
			$Script:DCDNSINFO = $False
		}

		If(!$Script:DARights -and !$Script:Elevated)
		{
			Write-Host "Warning: " -Foreground White
			Write-Host "Warning: To obtain Time Server and AD file location data, please run the script from an elevated PowerShell session using an account with Domain Admin rights." -Foreground White
			Write-Host "Warning: " -Foreground White
		}
	}

	#if computer name is localhost, get actual server name
	If($ComputerName -eq "localhost")
	{
		$Script:ComputerName = $env:ComputerName
		#V2.20 change from "renamed" to "changed"
		Write-Verbose "$(Get-Date): Server name has been changed from localhost to $($ComputerName)"
	}
	
	#see if default value of $Env:USERDNSDOMAIN was used
	If($ComputerName -eq $Env:USERDNSDOMAIN)
	{
		#change $ComputerName to a found global catalog server
		$Results = (Get-ADDomainController -DomainName $ADForest -Discover -Service GlobalCatalog -EA 0).Name
		
		If($? -and $Null -ne $Results)
		{
			$Script:ComputerName = $Results
			#V2.20 change from "renamed" to "changed"
			Write-Verbose "$(Get-Date): Server name has been changed from $Env:USERDNSDOMAIN to $ComputerName"
		}
		ElseIf(!$?) #changed for 2.16
		{
			#may be in a child domain where -Service GlobalCatalog doesn't work. Try PrimaryDC
			$Results = (Get-ADDomainController -DomainName $ADForest -Discover -Service PrimaryDC -EA 0).Name

			If($? -and $Null -ne $Results)
			{
				$Script:ComputerName = $Results
				#V2.20 change from "renamed" to "changed"
				Write-Verbose "$(Get-Date): Server name has been changed from $Env:USERDNSDOMAIN to $ComputerName"
			}
		}
	}

	#if computer name is an IP address, get host name from DNS
	#http://blogs.technet.com/b/gary/archive/2009/08/29/resolve-ip-addresses-to-hostname-using-powershell.aspx
	#help from Michael B. Smith
	$ip = $ComputerName -as [System.Net.IpAddress]
	If($ip)
	{
		$Result = [System.Net.Dns]::gethostentry($ip)
		
		If($? -and $Null -ne $Result)
		{
			$Script:ComputerName = $Result.HostName
			#V2.20 change from "renamed" to "changed"
			Write-Verbose "$(Get-Date): Server name has been changed from $ip to $ComputerName"
		}
		Else
		{
			Write-Warning "Unable to resolve $ComputerName to a hostname"
		}
	}
	Else
	{
		#server is online but for some reason $ComputerName cannot be converted to a System.Net.IpAddress
	}

	If(![String]::IsNullOrEmpty($ComputerName)) 
	{
		#get server name
		#first test to make sure the server is reachable
		Write-Verbose "$(Get-Date): Testing to see if $ComputerName is online and reachable"
		If(Test-Connection -ComputerName $ComputerName -quiet -EA 0)
		{
			Write-Verbose "$(Get-Date): Server $ComputerName is online."
			Write-Verbose "$(Get-Date): `tTest #1 to see if $ComputerName is a Domain Controller."
			#the server may be online but is it really a domain controller?

			#is the ComputerName in the current domain
			$Results = Get-ADDomainController $ComputerName -EA 0
			
			If(!$? -or $Null -eq $Results)
			{
				#try using the Forest name
				Write-Verbose "$(Get-Date): `tTest #2 to see if $ComputerName is a Domain Controller."
				$Results = Get-ADDomainController $ComputerName -Server $ADForest -EA 0
				If(!$?)
				{
					$ErrorActionPreference = $SaveEAPreference
					Write-Error "
					`n`n
					`t`t
					$ComputerName is not a domain controller for $ADForest.
					`n`n
					`t`t
					Script cannot continue.
					`n`n
					"
					Exit
				}
				Else
				{
					Write-Verbose "$(Get-Date): `tTest #2 succeeded. $ComputerName is a Domain Controller."
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date): `tTest #1 succeeded. $ComputerName is a Domain Controller."
			}
			
			$Results = $Null
		}
		Else
		{
			Write-Verbose "$(Get-Date): Computer $ComputerName is offline"
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "
			`n`n
			`t`t
			Computer $ComputerName is offline.
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
			Exit
		}
	}

	If($ADForest -ne $ADDomain)
	{
		#get forest information so output filename can be generated
		Write-Verbose "$(Get-Date): Testing to see if $($ADForest) is a valid forest name"
		If([String]::IsNullOrEmpty($ComputerName))
		{
			$Script:Forest = Get-ADForest -Identity $ADForest -EA 0
			
			If(!$?)
			{
				$ErrorActionPreference = $SaveEAPreference
				Write-Error "
				`n`n
				`t`t
				Could not find a forest identified by: $ADForest.
				`n`n
				`t`t
				Script cannot continue.
				`n`n
				"
				Exit
			}
		}
		Else
		{
			$Script:Forest = Get-ADForest -Identity $ADForest -Server $ComputerName -EA 0

			If(!$?)
			{
				$ErrorActionPreference = $SaveEAPreference
				Write-Error "
				`n`n
				`t`t
				Could not find a forest with the name of $ADForest.
				`n`n
				`t`t
				Script cannot continue.
				`n`n
				`t`t
				Is $ComputerName running Active Directory Web Services?
				`n`n
				"
				Exit
			}
		}
		Write-Verbose "$(Get-Date): $ADForest is a valid forest name"
		[string]$Script:Title = "AD Inventory Report for the $ADForest Forest"
		$Script:Domains       = $Script:Forest.Domains | Sort-Object 
		$Script:ConfigNC      = (Get-ADRootDSE -Server $ADForest -EA 0).ConfigurationNamingContext
	}
	
	If($ADDomain -ne "")
	{
		If([String]::IsNullOrEmpty($ComputerName))
		{
			$results = Get-ADDomain -Identity $ADDomain -EA 0
			
			If(!$?)
			{
				$ErrorActionPreference = $SaveEAPreference
				Write-Error "
				`n`n
				`t`t
				Could not find a domain identified by: $ADDomain.
				`n`n
				`t`t
				Script cannot continue.
				`n`n
				"
				Exit
			}
		}
		Else
		{
			$results = Get-ADDomain -Identity $ADDomain -Server $ComputerName -EA 0

			If(!$?)
			{
				$ErrorActionPreference = $SaveEAPreference
				Write-Error "
				`n`n
				`t`t
				Could not find a domain with the name of $ADDomain.
				`n`n
				`t`t
				Script cannot continue.
				`n`n
				`t`t
				Is $ComputerName running Active Directory Web Services?
				`n`n
				"
				Exit
			}
		}
		Write-Verbose "$(Get-Date): $ADDomain is a valid domain name"
		$Script:Domains       = $results.DNSRoot
		$Script:DomainDNSRoot = $results.DNSRoot
		[string]$Script:Title = "AD Inventory Report for the $Script:Domains Domain"
		
		$tmp = $results.Forest
		#get forest info 
		Write-Verbose "$(Get-Date): Retrieving forest information"
		If([String]::IsNullOrEmpty($ComputerName))
		{
			$Script:Forest = Get-ADForest -Identity $tmp -EA 0
			
			If(!$?)
			{
				$ErrorActionPreference = $SaveEAPreference
				Write-Error "
				`n`n
				`t`t
				Could not find a forest identified by: $tmp.
				`n`n
				`t`t
				Script cannot continue.
				`n`n
				"
				Exit
			}
		}
		Else
		{
			$Script:Forest = Get-ADForest -Identity $tmp -Server $ComputerName -EA 0

			If(!$?)
			{
				$ErrorActionPreference = $SaveEAPreference
				Write-Error "
				`n`n
				`t`t
				Could not find a forest with the name of $tmp.
				`n`n
				`t`t
				Script cannot continue.
				`n`n
				`t`t
				Is $ComputerName running Active Directory Web Services?
				`n`n
				"
				Exit
			}
		}
		Write-Verbose "$(Get-Date): Found forest information for $tmp"
		$Script:ConfigNC = (Get-ADRootDSE -Server $tmp -EA 0).ConfigurationNamingContext
	}
	
	#store root domain so it only has to be accessed once
	[string]$Script:ForestRootDomain = $Script:Forest.RootDomain
	[string]$Script:ForestName       = $Script:Forest.Name
	#set naming context
}
#endregion

######################START OF BUILDING REPORT

#region Forest information
Function ProcessForestInformation
{
	Write-Verbose "$(Get-Date): Writing forest data"

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Forest Information"
	}
	ElseIf($Text)
	{
		Line 0 "///  Forest Information  \\\"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\"
	}

	Switch ($Script:Forest.ForestMode)
	{
		"0"	{$ForestMode = "Windows 2000"; Break}
		"1" {$ForestMode = "Windows Server 2003 interim"; Break}
		"2" {$ForestMode = "Windows Server 2003"; Break}
		"3" {$ForestMode = "Windows Server 2008"; Break}
		"4" {$ForestMode = "Windows Server 2008 R2"; Break}
		"5" {$ForestMode = "Windows Server 2012"; Break}
		"6" {$ForestMode = "Windows Server 2012 R2"; Break}
		"7" {$ForestMode = "Windows Server 2016"; Break}	#added V2.20
		"Windows2000Forest"        {$ForestMode = "Windows 2000"; Break}
		"Windows2003InterimForest" {$ForestMode = "Windows Server 2003 interim"; Break}
		"Windows2003Forest"        {$ForestMode = "Windows Server 2003"; Break}
		"Windows2008Forest"        {$ForestMode = "Windows Server 2008"; Break}
		"Windows2008R2Forest"      {$ForestMode = "Windows Server 2008 R2"; Break}
		"Windows2012Forest"        {$ForestMode = "Windows Server 2012"; Break}
		"Windows2012R2Forest"      {$ForestMode = "Windows Server 2012 R2"; Break}
		"WindowsThresholdForest"   {$ForestMode = "Windows Server 2016 TP4"; Break}
		"Windows2016Forest"		   {$ForestMode = "Windows Server 2016"; Break}
		"UnknownForest"            {$ForestMode = "Unknown Forest Mode"; Break}
		Default                    {$ForestMode = "Unable to determine Forest Mode: $($Script:Forest.ForestMode)"; Break}
	}

	$AppPartitions         = $Script:Forest.ApplicationPartitions | Sort-Object 
	$CrossForestReferences = $Script:Forest.CrossForestReferences | Sort-Object 
	$SPNSuffixes           = $Script:Forest.SPNSuffixes | Sort-Object 
	$UPNSuffixes           = $Script:Forest.UPNSuffixes | Sort-Object 
	$Sites                 = $Script:Forest.Sites | Sort-Object 
	
	#added 9-oct-2016
	#https://adsecurity.org/?p=81
	$DirectoryServicesConfigPartition = Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,$Script:ConfigNC" -Partition $Script:ConfigNC -Properties *
	
	$TombstoneLifetime = $DirectoryServicesConfigPartition.tombstoneLifetime
	
	If($Null -eq $TombstoneLifetime -or $TombstoneLifetime -eq 0)
	{
		$TombstoneLifetime = 60
	}

	#2.16
	#move this duplicated block of code outside the output format test
	If($ADDomain -ne "")
	{
		#2.16 don't mess with the $Script:Domains variable
		#redo list of domains so forest root domain is listed first
		[array]$tmpDomains = "$Script:ForestRootDomain"
		[array]$tmpDomains2 = "$($Script:ForestRootDomain)"
		ForEach($Domain in $Forest.Domains)
		{
			If($Domain -ne $Script:ForestRootDomain)
			{
				$tmpDomains += "$($Domain.ToString())"
				$tmpDomains2 += "$($Domain.ToString())"
			}
		}
	}
	Else
	{
		#redo list of domains so forest root domain is listed first
		[array]$tmpDomains = "$Script:ForestRootDomain"
		[array]$tmpDomains2 = "$($Script:ForestRootDomain)"
		ForEach($Domain in $Script:Domains)
		{
			If($Domain -ne $Script:ForestRootDomain)
			{
				$tmpDomains += "$($Domain.ToString())"
				$tmpDomains2 += "$($Domain.ToString())"
			}
		}
		
		$Script:Domains = $tmpDomains
	}

	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Forest mode"; Value = $ForestMode; }
		$ScriptInformation += @{ Data = "Forest name"; Value = $Script:Forest.Name; }
		#V2.20 reorder to alpha order
		$tmp = ""
		If($Null -eq $AppPartitions)
		{
			$tmp = "<None>"
			$ScriptInformation += @{ Data = "Application partitions"; Value = $tmp; }
		}
		Else
		{
			$cnt = 0
			ForEach($AppPartition in $AppPartitions)
			{
				$cnt++
				$tmp = "$($AppPartition.ToString())"
				
				If($cnt -eq 1)
				{
					$ScriptInformation += @{ Data = "Application partitions"; Value = $tmp; }
				}
				Else
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$tmp = ""
		If($Null -eq $CrossForestReferences)
		{
			$tmp = "<None>"
			$ScriptInformation += @{ Data = "Cross forest references"; Value = $tmp; }
		}
		Else
		{
			$cnt = 0
			ForEach($CrossForestReference in $CrossForestReferences)
			{
				$cnt++
				$tmp = "$($CrossForestReference.ToString())"
				
				If($cnt -eq 1)
				{
					$ScriptInformation += @{ Data = "Cross forest references"; Value = $tmp; }
				}
				Else
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$ScriptInformation += @{ Data = "Domain naming master"; Value = $Script:Forest.DomainNamingMaster; }
		$tmp = ""
		If($Null -eq $Script:Domains)
		{
			$tmp = "<None>"
			$ScriptInformation += @{ Data = "Domains in forest"; Value = $tmp; }
		}
		Else
		{
			$cnt = 0
			ForEach($Domain in $tmpDomains2)
			{
				$cnt++
				$tmp = "$($Domain.ToString())"
				
				If($cnt -eq 1)
				{
					$ScriptInformation += @{ Data = "Domains in forest"; Value = $tmp; }
				}
				Else
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$ScriptInformation += @{ Data = "Partitions container"; Value = $Script:Forest.PartitionsContainer; }
		$ScriptInformation += @{ Data = "Root domain"; Value = $Script:ForestRootDomain; }
		$ScriptInformation += @{ Data = "Schema master"; Value = $Script:Forest.SchemaMaster; }
		$tmp = ""
		If($Null -eq $Sites)
		{
			$tmp = "<None>"
			$ScriptInformation += @{ Data = "Sites"; Value = $tmp; }
		}
		Else
		{
			$cnt = 0
			ForEach($Site in $Sites)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					$ScriptInformation += @{ Data = "Sites"; Value = $Site.ToString(); }
				}
				Else
				{
					$ScriptInformation += @{ Data = ""; Value = $Site.ToString(); }
				}
			}
		}
		$tmp = ""
		If($Null -eq $SPNSuffixes)
		{
			$tmp = "<None>"
			$ScriptInformation += @{ Data = "SPN suffixes"; Value = $tmp; }
		}
		Else
		{
			$cnt = 0
			ForEach($SPNSuffix in $SPNSuffixes)
			{
				$cnt++
				$tmp = "$($SPNSuffix.ToString())"
				
				If($cnt -eq 1)
				{
					$ScriptInformation += @{ Data = "SPN suffixes"; Value = $tmp; }
				}
				Else
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$ScriptInformation += @{ Data = "Tombstone lifetime"; Value = "$($TombstoneLifetime) days"; }
		$tmp = ""
		If($Null -eq $UPNSuffixes)
		{
			$tmp = "<None>"
			$ScriptInformation += @{ Data = "UPN suffixes"; Value = $tmp; }
		}
		Else
		{
			$cnt = 0
			ForEach($UPNSuffix in $UPNSuffixes)
			{
				$cnt++
				$tmp = "$($UPNSuffix.ToString())"
				
				If($cnt -eq 1)
				{
					$ScriptInformation += @{ Data = "UPN suffixes"; Value = $tmp; }
				}
				Else
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
		}
		$tmp = $Null

		Write-Verbose "$(Get-Date): `t`tCreate Forest Word table"
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 125;
		$Table.Columns.Item(2).Width = 300;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0  "Forest mode`t`t: " $ForestMode
		Line 0  "Forest name`t`t: " $Script:Forest.Name
		#V2.20 reorder to alpha order
		If($Null -eq $AppPartitions)
		{
			Line 0 "Application partitions`t: <None>"
		}
		Else
		{
			Line 0 "Application partitions`t: " -NoNewLine
			$cnt = 0
			ForEach($AppPartition in $AppPartitions)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					Line 0 "$($AppPartition.ToString())"
				}
				Else
				{
					Line 3 "  $($AppPartition.ToString())"
				}
			}
		}
		If($Null -eq $CrossForestReferences)
		{
			Line 0 "Cross forest references`t: <None>"
		}
		Else
		{
			Line 0 "Cross forest references`t: " -NoNewLine
			$cnt = 0
			ForEach($CrossForestReference in $CrossForestReferences)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					Line 0 "$($CrossForestReference.ToString())"
				}
				Else
				{
					Line 3 "  $($CrossForestReference.ToString())"
				}
			}
		}
		Line 0  "Domain naming master`t: " $Script:Forest.DomainNamingMaster
		If($Null -eq $Script:Domains)
		{
			Line 0 "Domains in forest`t: <None>"
		}
		Else
		{
			Line 0 "Domains in forest`t: " -NoNewLine
			$cnt = 0
			
			ForEach($Domain in $tmpDomains2)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					Line 0 $Domain
				}
				Else
				{
					Line 3 "  " $Domain
				}
			}
		}
		Line 0  "Partitions container`t: " $Script:Forest.PartitionsContainer
		Line 0  "Root domain`t`t: " $Script:ForestRootDomain
		Line 0  "Schema master`t`t: " $Script:Forest.SchemaMaster
		If($Null -eq $Sites)
		{
			Line 0 "Sites`t`t`t: <None>"
		}
		Else
		{
			Line 0 "Sites`t`t`t: " -NoNewLine
			$cnt = 0
			ForEach($Site in $Sites)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					Line 0 $Site.ToString()
				}
				Else
				{
					Line 3 "  $($Site.ToString())"
				}
			}
		}
		If($Null -eq $SPNSuffixes)
		{
			Line 0 "SPN suffixes`t`t: <None>"
		}
		Else
		{
			Line 0 "SPN suffixes`t`t: " -NoNewLine
			$cnt = 0
			ForEach($SPNSuffix in $SPNSuffixes)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					Line 0 "$($SPNSuffix.ToString())"
				}
				Else
				{
					Line 3 "  $($SPNSuffix.ToString())"
				}
			}
		}
		Line 0 "Tombstone lifetime`t: " "$($TombstoneLifetime) days"
		If($Null -eq $UPNSuffixes)
		{
			Line 0 "UPN Suffixes`t`t: <None>"
		}
		Else
		{
			Line 0 "UPN Suffixes`t`t: " -NoNewLine
			$cnt = 0
			ForEach($UPNSuffix in $UPNSuffixes)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					Line 0 "$($UPNSuffix.ToString())"
				}
				Else
				{
					Line 3 "  $($UPNSuffix.ToString())"
				}
			}
		}
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Forest mode",($htmlsilver -bor $htmlbold),$ForestMode,$htmlwhite)
		$rowdata += @(,('Forest name',($htmlsilver -bor $htmlbold),$Script:Forest.Name,$htmlwhite))
		$rowdata += @(,('Root domain',($htmlsilver -bor $htmlbold),$Script:ForestRootDomain,$htmlwhite))
		#V2.20 reorder to alpha order
		$tmp = ""
		If($Null -eq $AppPartitions)
		{
			$tmp = "None"
			
			$rowdata += @(,('Application partitions',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
		}
		Else
		{
			$cnt = 0
			ForEach($AppPartition in $AppPartitions)
			{
				$cnt++
				$tmp = "$($AppPartition.ToString())"
				
				If($cnt -eq 1)
				{
					$rowdata += @(,('Application partitions',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$tmp = ""
		If($Null -eq $CrossForestReferences)
		{
			$tmp = "None"
			$rowdata += @(,('Cross forest references',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
		}
		Else
		{
			$cnt = 0
			ForEach($CrossForestReference in $CrossForestReferences)
			{
				$cnt++
				$tmp = "$($CrossForestReference.ToString())"
				
				If($cnt -eq 1)
				{
					$rowdata += @(,('Cross forest references',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('Domain naming master',($htmlsilver -bor $htmlbold),$Script:Forest.DomainNamingMaster,$htmlwhite))
		$tmp = ""
		If($Null -eq $Script:Domains)
		{
			$tmp = "None"
			$rowdata += @(,('Domains in forest',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
		}
		Else
		{
			$cnt = 0
			ForEach($Domain in $tmpDomains2)
			{
				$cnt++
				$tmp = "$($Domain.ToString())"
				
				If($cnt -eq 1)
				{
					$rowdata += @(,('Domains in forest',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('Partitions container',($htmlsilver -bor $htmlbold),$Script:Forest.PartitionsContainer,$htmlwhite))
		$rowdata += @(,('Schema master',($htmlsilver -bor $htmlbold),$Script:Forest.SchemaMaster,$htmlwhite))
		$tmp = ""
		If($Null -eq $Sites)
		{
			$tmp = "None"
			$rowdata += @(,('Sites',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
		}
		Else
		{
			$cnt = 0
			ForEach($Site in $Sites)
			{
				$cnt++
				
				If($cnt -eq 1)
				{
					$rowdata += @(,('Sites',($htmlsilver -bor $htmlbold),$Site.ToString(),$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$Site.ToString(),$htmlwhite))
				}
			}
		}
		$tmp = ""
		If($Null -eq $SPNSuffixes)
		{
			$tmp = "None"
			$rowdata += @(,('SPN suffixes',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
		}
		Else
		{
			$cnt = 0
			ForEach($SPNSuffix in $SPNSuffixes)
			{
				$cnt++
				$tmp = "$($SPNSuffix.ToString())"
				
				If($cnt -eq 1)
				{
					$rowdata += @(,('SPN suffixes',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('Tombstone lifetime',($htmlsilver -bor $htmlbold),"$($TombstoneLifetime) days",$htmlwhite))
		$tmp = ""
		If($Null -eq $UPNSuffixes)
		{
			$tmp = "None"
			$rowdata += @(,('UPN suffixes',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
		}
		Else
		{
			$cnt = 0
			ForEach($UPNSuffix in $UPNSuffixes)
			{
				$cnt++
				$tmp = "$($UPNSuffix.ToString())"
				
				If($cnt -eq 1)
				{
					$rowdata += @(,('UPN suffixes',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$tmp = $Null

		$msg = ""
		$columnWidths = @("125","300")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "425"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region get all DCs in the forest
Function ProcessAllDCsInTheForest
{
	Write-Verbose "$(Get-Date): `tDomain controllers"

	$txt = "Domain Controllers"
	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 $txt
	}

	#http://www.superedge.net/2012/09/how-to-get-ad-forest-in-powershell.html
	#http://msdn.microsoft.com/en-us/library/vstudio/system.directoryservices.activedirectory.forest.getforest%28v=vs.90%29
	#$ADContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("forest", $ADForest) 
	#2.16 change
	$ADContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("forest", $Script:Forest.Name)
	$Forest2 = [system.directoryservices.activedirectory.Forest]::GetForest($ADContext)
	Write-Verbose "$(Get-Date): `t`tBuild list of Domain controllers in the Forest"
	$AllDCs = $Forest2.domains | ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name} 
	Write-Verbose "$(Get-Date): `t`tSort list of all Domain controllers"
	$AllDCs = $AllDCs | Sort-Object 
	$ADContext = $Null
	$Forest2 = $Null

	If($Null -eq $AllDCs)
	{
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 "<None>"
		}
		ElseIf($Text)
		{
			Line 0 "<None>"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "None"
		}
	}
	Else
	{
		If($MSWORD -or $PDF)
		{
			[System.Collections.Hashtable[]] $WordTableRowHash = @();
			ForEach($DC in $AllDCs)
			{
				Write-Verbose "$(Get-Date): `t`t`t$DC"
				$DCName = $DC.SubString(0,$DC.IndexOf("."))
				$SrvName = $DC.SubString($DC.IndexOf(".")+1)
				
				$Results = Get-ADDomainController -Identity $DCName -Server $SrvName -EA 0
				
				If($? -and $Null -ne $Results)
				{
					$GC = $Results.IsGlobalCatalog.ToString()
					$ReadOnly = $Results.IsReadOnly.ToString()
					#ServerOS and ServerCore added in V2.20
					$ServerOS = $Results.OperatingSystem
					#https://blogs.msmvps.com/russel/2017/03/16/how-to-tell-if-youre-running-on-windows-server-core/
					$tmp = Get-RegistryValue "HKLM:\software\microsoft\windows nt\currentversion" "installationtype" $DCName
					If($tmp -eq "Server Core")
					{
						$ServerCore = "Yes"
					}
					Else
					{
						$ServerCore = "No"
					}
				}
				Else
				{
					$GC = "Unable to retrieve status"
					$ReadOnly = "Unable to retrieve status"
					$ServerOS = "Unable to retrieve status"
					$ServerCore = "Unable to retrieve status"
				}
				
				$WordTableRowHash += @{ 
				DCName = $DC; 
				GC = $GC; 
				ReadOnly = $ReadOnly;
				ServerOS = $ServerOS;
				ServerCore = $ServerCore
				}
				
				$Results = $Null
			}
		}
		ElseIf($Text)
		{
			#V2.16 addition
			[int]$MaxDCNameLength = ($AllDCs | Measure-Object -Maximum -Property Length).Maximum
			
			If($MaxDCNameLength -gt 4) #4 is length of "Name"
			{
				#2 is to allow for spacing between columns
				Line 1 ("Name" + (' ' * ($MaxDCNameLength - 2))) -NoNewLine
				Line 0 "Global Catalog  Read-only  Server OS                       Server Core"
				Line 1 ('=' * $MaxDCNameLength) -NoNewLine
				Line 0 "==============================================================================================="
			}
			Else
			{
				Line 1 "Name  Global Catalog  Read-only  Server OS                       Server Core"
				Line 1 "============================================================================"
			}
			
			ForEach($DC in $AllDCs)
			{
				Write-Verbose "$(Get-Date): `t`t`t$DC"
				$DCName = $DC.SubString(0,$DC.IndexOf("."))
				$SrvName = $DC.SubString($DC.IndexOf(".")+1)
				
				$Results = Get-ADDomainController -Identity $DCName -Server $SrvName -EA 0
				
				#V2.16 change
				If($? -and $Null -ne $Results)
				{
					$xGC = $Results.IsGlobalCatalog.ToString()
					$xRO = $Results.IsReadOnly.ToString()
					#ServerOS and ServerCore added in V2.20
					$ServerOS = $Results.OperatingSystem
					#https://blogs.msmvps.com/russel/2017/03/16/how-to-tell-if-youre-running-on-windows-server-core/
					$tmp = Get-RegistryValue "HKLM:\software\microsoft\windows nt\currentversion" "installationtype" $DCName
					If($tmp -eq "Server Core")
					{
						$ServerCore = "Yes"
					}
					Else
					{
						$ServerCore = "No"
					}
				}
				Else
				{
					$xGC = "Unable to retrieve status"
					$xRO = "Unable to retrieve status"
					$ServerOS = "Unable to retrieve"
					$ServerCore = "N/A"
				}

				If(($DC).Length -lt ($MaxDCNameLength))
				{
					[int]$NumOfSpaces = ($MaxDCNameLength * -1) 
				}
				Else
				{
					[int]$NumOfSpaces = -4
				}
				Line 1 ( "{0,$NumOfSpaces}  {1,-15} {2,-10} {3,-31} {4,-3}" -f $DC,$xGC,$xRO,$ServerOS,$ServerCore)

				$Results = $Null
			}
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata = @()
			
			ForEach($DC in $AllDCs)
			{
				Write-Verbose "$(Get-Date): `t`t`t$DC"
				$DCName = $DC.SubString(0,$DC.IndexOf("."))
				$SrvName = $DC.SubString($DC.IndexOf(".")+1)
				
				$Results = Get-ADDomainController -Identity $DCName -Server $SrvName -EA 0
				
				If($? -and $Null -ne $Results)
				{
					$GC = $Results.IsGlobalCatalog.ToString()
					$ReadOnly = $Results.IsReadOnly.ToString()
					#ServerOS and ServerCore added in V2.20
					$ServerOS = $Results.OperatingSystem
					#https://blogs.msmvps.com/russel/2017/03/16/how-to-tell-if-youre-running-on-windows-server-core/
					$tmp = Get-RegistryValue "HKLM:\software\microsoft\windows nt\currentversion" "installationtype" $DCName
					If($tmp -eq "Server Core")
					{
						$ServerCore = "Yes"
					}
					Else
					{
						$ServerCore = "No"
					}
				}
				Else
				{
					$GC = "Unable to retrieve status"
					$ReadOnly = "Unable to retrieve status"
					$ServerOS = "Unable to retrieve status"
					$ServerCore = "Unable to retrieve status"
				}
				
				$rowdata += @(,($DC,$htmlwhite,
								$GC,$htmlwhite,
								$ReadOnly,$htmlwhite,
								$ServerOS,$htmlwhite,
								$ServerCore,$htmlwhite
								))
			}
		}
	}
	If($MSWord -or $PDF)
	{
		Write-Verbose "$(Get-Date): `t`tCreate Domain Controller in Forest Word table"
		$Table = AddWordTable -Hashtable $WordTableRowHash `
		-Columns DCName, GC, ReadOnly, ServerOS, ServerCore `
		-Headers "Name", "Global Catalog", "Read-only", "Server OS", "Server Core" `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 50;
		$Table.Columns.Item(3).Width = 50;
		$Table.Columns.Item(4).Width = 130;
		$Table.Columns.Item(5).Width = 45;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		#nothing to do
	}
	ElseIf($HTML)
	{
		$columnHeaders = @('Name',($htmlsilver -bor $htmlbold),
							'Global Catalog',($htmlsilver -bor $htmlbold),
							'Read-only',($htmlsilver -bor $htmlbold),
							'Server OS',($htmlsilver -bor $htmlbold),
							'Server Core',($htmlsilver -bor $htmlbold))
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 " "
	}
	$AllDCs = $Null
}
#endregion

#region process CA information
Function ProcessCAInformation
{
	Write-Verbose "$(Get-Date): `tCA Information"
	
	$txt = "Certificate Authority Information"
	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 $txt
	}

	$txt = "Certification Authority Root(s)"
	If($MSWORD -or $PDF)
	{
		WriteWordLine 4 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 $txt
	}

	$rootDSE = [ADSI]"LDAP://RootDSE"

	$configNC = $rootDSE.Properties[ 'configurationNamingContext' ].Value -as [String]

	$rootCA = 'CN=Certification Authorities,CN=Public Key Services,CN=Services,' + $configNC
	$rootObj = [ADSI] ( 'LDAP://' + $rootCA )
	$RootCnt = 0
	$AllCnt = 0
	
	If($Null -ne $rootObj)
	{
		ForEach($obj in $rootObj.psbase.children)
		{
			$RootCnt++
			If($MSWORD -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Common name"; Value = $obj.cn; }
				$ScriptInformation += @{ Data = "Distinguished name"; Value = $obj.distinguishedName; }
				Write-Verbose "$(Get-Date): `t`tCreate Certification Authority Root(s) Word table"
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 125;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 1 "Common name`t`t: " $obj.cn
				Line 1 "Distinguished name`t: " $obj.distinguishedName
				Line 1 ""
			}
			ElseIf($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Common name",($htmlsilver -bor $htmlbold),$obj.cn,$htmlwhite)
				$rowdata += @(,('Distinguished name',($htmlsilver -bor $htmlbold),$obj.distinguishedName,$htmlwhite))
				$msg = ""
				$columnWidths = @("125","400")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "525"
				WriteHTMLLine 0 0 " "
			}
		}
		
		If($RootCnt -gt 1)
		{
			$txt = "Possible error: There are more than one Certification Authority Root(s)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 0 $txt
				Line 0 ""
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
				WriteHTMLLine 0 0 ""
			}
		}
	}
	#ElseIf($Null -eq $rootObj) changed in V2.22 by Michael B. Smith
	If($RootCnt -eq 0 -or $Null -eq $rootObj)
	{
		$txt = "No Certification Authority Root(s) were retrieved"
		Write-Warning $txt
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 $txt "" $Null 0 $False $True
			WriteHTMLLine 0 0 " "
		}
	}

	$txt = "Certification Authority Issuer(s)"
	If($MSWORD -or $PDF)
	{
		WriteWordLine 4 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 $txt
	}

	$allCA = 'CN=Enrollment Services,CN=Public Key Services,CN=Services,' + $configNC
	$allObj = [ADSI] ( 'LDAP://' + $allCA )
	
	If([string]::isnullorempty($allObj.psbase.children) -and !([string]::isnullorempty($rootObj.psbase.children)))
	{
		#uh oh error
		$txt = "Error: Certification Authority Root(s) exist, but no Certification Authority Issuers(s) (also known as Enrollment Agents) exist"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 $txt
			WriteHTMLLine 0 0 ""
		}
	}
	ElseIf(!([string]::isnullorempty($allObj.psbase.children)) -and !([string]::isnullorempty($rootObj.psbase.children)))
	{
		$AllCnt = 0
		ForEach($obj in $allObj.psbase.children)
		{
			$AllCnt++
			If($MSWORD -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Common name"; Value = $obj.cn; }
				$ScriptInformation += @{ Data = "Distinguished name"; Value = $obj.distinguishedName; }
				Write-Verbose "$(Get-Date): `t`tCreate CA Authorities Word table"
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 125;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 1 "Common name`t`t: " $obj.cn
				Line 1 "Distinguished name`t: " $obj.distinguishedName
				Line 1 ""
			}
			ElseIf($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Common name",($htmlsilver -bor $htmlbold),$obj.cn,$htmlwhite)
				$rowdata += @(,('Distinguished name',($htmlsilver -bor $htmlbold),$obj.distinguishedName,$htmlwhite))
				$msg = ""
				$columnWidths = @("125","400")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "525"
				WriteHTMLLine 0 0 " "
			}
		}
		If($AllCnt -lt $RootCnt)
		{
			$txt = "Error: More Certification Authority Root(s) exist than there are Certification Authority Issuers(s) (also known as Enrollment Agents)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 0 $txt
				Line 0 ""
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
				WriteHTMLLine 0 0 ""
			}
		}
	}
	ElseIf(([string]::isnullorempty($allObj.psbase.children)) -and ([string]::isnullorempty($rootObj.psbase.children)))
	{
		$txt = "No Certification Authority Issuer(s) were retrieved"
		Write-Warning $txt
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 $txt "" $Null 0 $False $True
			WriteHTMLLine 0 0 " "
		}
	}
	
	#if you have enrollment authorities and no roots – that’s a BIG error
	If($AllCnt -gt 0 -and $RootCnt -eq 0)
	{
		$txt = "Error: Certification Authority Issuers(s) (also known as Enrollment Agents) exist, but no Certification Authority Root(s) exist"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 $txt
			WriteHTMLLine 0 0 ""
		}
	}
	
}
#endregion

#region process ad optional features
Function ProcessADOptionalFeatures
{
	Write-Verbose "$(Get-Date): `tAD Optional Features"
	
	$txt = "AD Optional Features"
	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 $txt
	}
	ElseIf($Text)
	{
		Line 0 $txt
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 $txt
	}

	$ADOptionalFeatures = Get-ADOptionalFeature -Filter * -EA 0
	
	If($? -and $Null -ne $ADOptionalFeatures)
	{
		ForEach($Item in $ADOptionalFeatures)
		{
			$Enabled = "No"
			If($Item.EnabledScopes.Count -gt 0)
			{
				$Enabled = "Yes"
				$EnabledScopes = $Item.EnabledScopes | Sort-Object 
			}
			
			If($MSWORD -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Feature name"; Value = $Item.Name; }
				$ScriptInformation += @{ Data = "Enabled"; Value = $Enabled; }
				
				If($Enabled -eq "Yes")
				{
					
					$cnt = 0
					ForEach($Scope in $EnabledScopes)
					{
						$cnt++
					
						If($cnt -eq 1)
						{
							$ScriptInformation += @{ Data = "Enabled Scopes"; Value = $Scope; }
						}
						Else
						{
							$ScriptInformation += @{ Data = ""; Value = $($Scope); }
						}
					}
				}
				Write-Verbose "$(Get-Date): `t`tCreate AD Optional Features Word table"
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 125;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 1 "Feature Name`t: " $Item.Name
				Line 1 "Enabled`t`t: " $Enabled
				
				If($Enabled -eq "Yes")
				{
					Line 1 "Enabled Scopes`t: " -NoNewLine
					
					$cnt = 0
					ForEach($Scope in $EnabledScopes)
					{
						$cnt++
					
						If($cnt -eq 1)
						{
							Line 0 $Scope
						}
						Else
						{
							Line 3 "  $($Scope)"
						}
					}
					Line 0 ""
				}
				Else
				{
					Line 0 ""
				}
			}
			ElseIf($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Feature Name",($htmlsilver -bor $htmlbold),$Item.Name,$htmlwhite)
				$rowdata += @(,('Enabled',($htmlsilver -bor $htmlbold),$Enabled,$htmlwhite))
				
				If($Enabled -eq "Yes")
				{
					
					$cnt = 0
					ForEach($Scope in $EnabledScopes)
					{
						$cnt++
					
						If($cnt -eq 1)
						{
							$rowdata += @(,('Enabled Scopes',($htmlsilver -bor $htmlbold),$Scope,$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),$Scope,$htmlwhite))
						}
					}
				}
				$msg = ""
				$columnWidths = @("125","400")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "525"
				WriteHTMLLine 0 0 " "
			}
		}
	}
	ElseIf($? -and $Null -eq $ADOptionalFeatures)
	{
		$txt = "No AD Optional Features were retrieved"
		Write-Warning $txt
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 $txt "" $Null 0 $False $True
			WriteHTMLLine 0 0 " "
		}
	}
	Else
	{
		$txt = "Error retieving AD Optional Features"
		Write-Warning $txt
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt "" $Null 0 $False $True
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 0 $txt
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 $txt "" $Null 0 $False $True
			WriteHTMLLine 0 0 " "
		}
	}
}
#endregion

#region process ad schema items

#new for 2.16
Function ProcessADSchemaItems
{
	Param(
		[String []] $Name = @( 
		'User-Account-Control', #Flags that control the behavior of a user account
		'msNPAllowDialin', #RAS Server
		'ms-Mcs-AdmPwd', #LAPS
		'ms-Mcs-AdmPwdExpirationTime', #LAPS
		'ms-SMS-Assignment-Site-Code', #SCCM
		'ms-SMS-Capabilities', #SCCM
		'msRTCSIP-UserRoutingGroupId', #Lync/SfB
		'msRTCSIP-MirrorBackEndServer', #Lync/SfB
		'ms-exch-schema-version-pt' #Exchange
		)
	)

	Write-Verbose "$(Get-Date): `tAD Schema Items"
	
	$txt = "AD Schema Items"
	$txt1 = "Just because a schema extension is Present does not mean it is in use."
	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 $txt
		WriteWordLine 0 0 $txt1 "" $Null 8 $False $True	
	}
	ElseIf($Text)
	{
		Line 0 $txt
		Line 0 $txt1
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 $txt
		WriteHTMLLine 0 0 $txt1 "" "Calibri" 1
	}

	$rootDS   = [ADSI] 'LDAP://RootDSE'
	$schemaNC = $rootDS.schemaNamingContext.Item( 0 )

	$SchemaItems = New-Object System.Collections.ArrayList
	ForEach( $item in $Name )
	{
		Write-Verbose "$(Get-Date): `t`tChecking for $item declared in schema"

		$objDN = 'LDAP://' + 'CN=' + $item + ',' + $schemaNC

		$obj = [ADSI] $objDN
		$mem = Get-Member -Name name -InputObject $obj

		$Itemobj = New-Object -TypeName PSObject

		Switch ($item)
		{
			'User-Account-Control'			{$tmp = "Flags that control the behavior of a user account"}
			'msNPAllowDialin'				{$tmp = "RAS Server"}
			'ms-Mcs-AdmPwd'					{$tmp = "LAPS"}
			'ms-Mcs-AdmPwdExpirationTime'	{$tmp = "LAPS"}
			'ms-SMS-Assignment-Site-Code'	{$tmp = "SCCM"}
			'ms-SMS-Capabilities'			{$tmp = "SCCM"}
			'msRTCSIP-UserRoutingGroupId'	{$tmp = "Lync/Skype for Business"}
			'msRTCSIP-MirrorBackEndServer'	{$tmp = "Lync/Skype for Business"}
			'ms-exch-schema-version-pt' 	{$tmp = "Exchange"}
			Default							{$tmp = "Unknown"}
		}
		
		If( $mem )
		{
			$Itemobj | Add-Member -MemberType NoteProperty -Name ItemName	-Value $item
			$Itemobj | Add-Member -MemberType NoteProperty -Name ItemState	-Value "Present"
			$Itemobj | Add-Member -MemberType NoteProperty -Name ItemDesc	-Value $tmp
			$SchemaItems.Add($Itemobj) > $Null
		}
		Else
		{
			$Itemobj | Add-Member -MemberType NoteProperty -Name ItemName	-Value $item
			$Itemobj | Add-Member -MemberType NoteProperty -Name ItemState	-Value "Not Present"
			$Itemobj | Add-Member -MemberType NoteProperty -Name ItemDesc	-Value $tmp
			$SchemaItems.Add($Itemobj) > $Null
		}
		$mem = $null
		$obj = $null
	}

	If($MSWORD -or $PDF)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 3
		[int]$Rows = $SchemaItems.Count + 1
		[int]$xRow = 1

		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.AutoFitBehavior($wdAutoFitFixed)
		$Table.Style = $Script:MyHash.Word_TableGrid
	
		$Table.rows.first.headingformat = $wdHeadingFormatTrue
		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

		$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Schema item name"
		
		$Table.Cell($xRow,2).Range.Font.Bold = $True
		$Table.Cell($xRow,2).Range.Text = "Present"
		
		$Table.Cell($xRow,3).Range.Font.Bold = $True
		$Table.Cell($xRow,3).Range.Text = "Used for"
		
	}
	ElseIf($Text)
	{
		#V2.16 change
		Line 1 "Schema item name                Present      Used for                                         "
		Line 1 "=============================================================================================="
	}
	ElseIf($HTML)
	{
		$rowdata = @()
	}
	
	ForEach($item in $SchemaItems)
	{
		If($MSWORD -or $PDF)
		{
			$xRow++
			If($xRow % 2 -eq 0)
			{
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
				$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray05
				$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray05
			}
			$Table.Cell($xRow,1).Range.Text = $Item.ItemName
			$Table.Cell($xRow,2).Range.Text = $Item.ItemState
			$Table.Cell($xRow,3).Range.Text = $Item.ItemDesc
		}
		ElseIf($Text)
		{
			#V2.16 change
			Line 1 ( "{0,-30}  {1,-11}  {2,-50}" -f $Item.ItemName,$Item.ItemState,$Item.ItemDesc)
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
			$Item.ItemName,$htmlwhite,
			$Item.ItemState,$htmlwhite,
			$Item.ItemDesc,$htmlwhite))
		}
	}

	If($MSWORD -or $PDF)
	{
		#set column widths
		$xcols = $table.columns

		ForEach($xcol in $xcols)
		{
			switch ($xcol.Index)
			{
			  1 {$xcol.width = 175; Break}
			  2 {$xcol.width = 75; Break}
			  3 {$xcol.width = 175; Break}
			}
		}
		
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
		$Table.AutoFitBehavior($wdAutoFitFixed)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		WriteWordLine 0 0 ""
		$TableRange = $Null
		$Table = $Null
	}
	ElseIf($Text)
	{
		#V2.16 change
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$columnHeaders = @('Schema item name',($htmlsilver -bor $htmlbold),
							'Present',($htmlsilver -bor $htmlbold),
							'Used for',($htmlsilver -bor $htmlbold)
							)
		$msg = ""
		$columnWidths = @("175","75","175")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "475"
		WriteHTMLLine 0 0 " "
	}
	
	$rootDS      = $null
	$schemaNC    = $null
	$objDN       = $null
	$SchemaItems = $null
}
#endregion

#region Site information
Function ProcessSiteInformation
{
	Write-Verbose "$(Get-Date): Writing sites and services data"

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Sites and Services"
	}
	ElseIf($Text)
	{
		Line 0 "///  Sites and Services  \\\"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Sites and Services&nbsp;&nbsp;\\\"
	}
	
	#get site information
	#some of the following was taken from
	#http://blogs.msdn.com/b/adpowershell/archive/2009/08/18/active-directory-powershell-to-manage-sites-and-subnets-part-3-getting-site-and-subnets.aspx

	$tmp = $Script:Forest.PartitionsContainer
	$ConfigurationBase = $tmp.SubString($tmp.IndexOf(",") + 1)
	$Sites = $Null
	$Sites = Get-ADObject -Filter 'ObjectClass -eq "site"' -SearchBase $ConfigurationBase -Properties Name, SiteObjectBl -Server $ADForest -EA 0 | Sort-Object Name

	$siteContainerDN = ("CN=Sites," + $Script:ConfigNC)

	If($? -and $Null -ne $Sites)
	{
		Write-Verbose "$(Get-Date): `tProcessing Inter-Site Transports"
		$AllSiteLinks = Get-ADObject -Searchbase $Script:ConfigNC -Server $ADForest `
		-Filter 'objectClass -eq "siteLink"' -Property Description, Options, Cost, ReplInterval, SiteList, Schedule -EA 0 `
		| Select-Object Name, Description, @{Name="SiteCount";Expression={$_.SiteList.Count}}, Cost, ReplInterval, `
		@{Name="Schedule";Expression={If($_.Schedule){If(($_.Schedule -Join " ").Contains("240")){"NonDefault"}Else{"24x7"}}Else{"24x7"}}}, `
		Options, SiteList, DistinguishedName

		If($MSWORD -or $PDF)
		{
			WriteWordLine 2 0 "Inter-Site Transports"
			#adapted from code provided by Goatee PFE
			#http://blogs.technet.com/b/ashleymcglone/archive/2011/06/29/report-and-edit-ad-site-links-from-powershell-turbo-your-ad-replication.aspx
			# Report of all site links and related settings
			
			If($? -and $Null -ne $AllSiteLinks)
			{
				ForEach($SiteLink in $AllSiteLinks)
				{
					Write-Verbose "$(Get-Date): `t`tProcessing site link $($SiteLink.Name)"
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$SiteLinkTypeDN = @()
					$SiteLinkTypeDN = $SiteLink.DistinguishedName.Split(",")
					$SiteLinkType = $SiteLinkTypeDN[1].SubString(3)
					$SitesInLink = New-Object System.Collections.ArrayList
					$SiteLinkSiteList = $SiteLink.SiteList
					ForEach($xSite in $SiteLinkSiteList)
					{
						$tmp = $xSite.Split(",")
						$SitesInLink.Add("$($tmp[0].SubString(3))") > $Null
					}
					
					$ScriptInformation += @{ Data = "Name"; Value = $SiteLink.Name; }
					If(![String]::IsNullOrEmpty($SiteLink.Description))
					{
						$ScriptInformation += @{ Data = "Description"; Value = $SiteLink.Description; }
					}
					If($SitesInLink -ne "")
					{
						$cnt = 0
						ForEach($xSite in $SitesInLink)
						{
							$cnt++
							
							If($cnt -eq 1)
							{
								$ScriptInformation += @{ Data = "Sites in Link"; Value = $xSite; }
							}
							Else
							{
								$ScriptInformation += @{ Data = ""; Value = $xSite; }
							}
						}
					}
					Else
					{
						$ScriptInformation += @{ Data = "Sites in Link"; Value = <None>; }
					}
					$ScriptInformation += @{ Data = "Cost"; Value = $SiteLink.Cost.ToString(); }
					$ScriptInformation += @{ Data = "Replication Interval"; Value = $SiteLink.ReplInterval.ToString(); }
					$ScriptInformation += @{ Data = "Schedule"; Value = $SiteLink.Schedule; }

					#https://msdn.microsoft.com/en-us/library/cc223552.aspx
					$tmp = ""
					If([String]::IsNullOrEmpty($SiteLink.Options) -or $SiteLink.Options -eq "0")
					{
						$tmp = "Change Notification is Disabled"
					}
					ElseIf($SiteLink.Options -eq "1")
					{
						$tmp = "Change Notification is Enabled with Compression"
					}
					ElseIf($SiteLink.Options -eq "2")
					{
						$tmp = "Force sync in opposite direction at end of sync"
					}
					ElseIf($SiteLink.Options -eq "3")
					{
						$tmp = "Change Notification is Enabled with Compression and Force sync in opposite direction at end of sync"
					}
					ElseIf($SiteLink.Options -eq "4")
					{
						$tmp = "Disable compression of Change Notification messages"
					}
					ElseIf($SiteLink.Options -eq "5")
					{
						$tmp = "Change Notification is Enabled without Compression"
					}
					ElseIf($SiteLink.Options -eq "6")
					{
						$tmp = "Force sync in opposite direction at end of sync and Disable compression of Change Notification messages"
					}
					ElseIf($SiteLink.Options -eq "7")
					{
						$tmp = "Change Notification is Enabled without Compression and Force sync in opposite direction at end of sync"
					}
					Else
					{
						$tmp = "Unknown"
					}
					$ScriptInformation += @{ Data = "Options"; Value = $tmp; }
					$ScriptInformation += @{ Data = "Type"; Value = $SiteLinkType; }

					Write-Verbose "$(Get-Date): `t`t`tCreate Site Links Word table"
					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitContent;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
			}
			$AllSiteLinks = $Null
			
			ForEach($Site in $Sites)
			{
				Write-Verbose "$(Get-Date): `tProcessing site $($Site.Name)"
				WriteWordLine 2 0 "Site: " $Site.Name

				WriteWordLine 3 0 "Subnets"
				Write-Verbose "$(Get-Date): `t`tProcessing subnets"
				$subnetArray = New-Object -Type string[] -ArgumentList $Site.siteObjectBL.Count
				$i = 0
				$SitesiteObjectBL = $Site.siteObjectBL
				ForEach($subnetDN in $SitesiteObjectBL) 
				{
					$subnetName = $subnetDN.SubString(3, $subnetDN.IndexOf(",CN=Subnets,CN=Sites,") - 3)
					$subnetArray[$i] = $subnetName
					$i++
				}
				$subnetArray = $subnetArray | Sort-Object 
				If($Null -eq $subnetArray)
				{
					WriteWordLine 0 0 "<None>"
				}
				Else
				{
					BuildMultiColumnTable $subnetArray "Subnets"
				}
				
				Write-Verbose "$(Get-Date): `t`tProcessing servers"
				WriteWordLine 3 0 "Servers"
				$siteName = $Site.Name
				
				#build array of connect objects
				Write-Verbose "$(Get-Date): `t`t`tProcessing automatic connection objects"
				$Connections = New-Object System.Collections.ArrayList
				$ConnectionObjects = $Null
				$ConnectionObjects = Get-ADObject -Filter 'objectClass -eq "nTDSConnection" -and options -bor 1' -Searchbase $Script:ConfigNC -Property DistinguishedName, fromServer -Server $ADForest -EA 0
				
				If($? -and $Null -ne $ConnectionObjects)
				{
					ForEach($ConnectionObject in $ConnectionObjects)
					{
						$xArray = $ConnectionObject.DistinguishedName.Split(",")
						#server name is 3rd item in array (element 2)
						$ToServer = $xArray[2].SubString($xArray[2].IndexOf("=")+1) #get past the = sign
						$xArray = $ConnectionObject.FromServer.Split(",")
						#server name is 2nd item in array (element 1)
						$FromServer = $xArray[1].SubString($xArray[1].IndexOf("=")+1) #get past the = sign
						#site name is 4th item in array (element 3)
						$FromServerSite = $xArray[3].SubString($xArray[3].IndexOf("=")+1) #get past the = sign
						$xArray = $Null
						$obj = New-Object -TypeName PSObject
						$obj | Add-Member -MemberType NoteProperty -Name Name           -Value "<automatically generated>"
						$obj | Add-Member -MemberType NoteProperty -Name ToServer       -Value $ToServer
						$obj | Add-Member -MemberType NoteProperty -Name FromServer     -Value $FromServer
						$obj | Add-Member -MemberType NoteProperty -Name FromServerSite -Value $FromServerSite
						$Connections.Add($obj) > $Null
					}
				}
				
				Write-Verbose "$(Get-Date): `t`t`tProcessing manual connection objects"
				$ConnectionObjects = $Null
				$ConnectionObjects = Get-ADObject -Filter 'objectClass -eq "nTDSConnection" -and -not options -bor 1' -Searchbase $Script:ConfigNC -Property Name, DistinguishedName, fromServer -Server $ADForest -EA 0
				
				If($? -and $Null -ne $ConnectionObjects)
				{
					ForEach($ConnectionObject in $ConnectionObjects)
					{
						$xArray = $ConnectionObject.DistinguishedName.Split(",")
						#server name is 3rd item in array (element 2)
						$ToServer = $xArray[2].SubString($xArray[2].IndexOf("=")+1) #get past the = sign
						$xArray = $ConnectionObject.FromServer.Split(",")
						#server name is 2nd item in array (element 1)
						$FromServer = $xArray[1].SubString($xArray[1].IndexOf("=")+1) #get past the = sign
						#site name is 4th item in array (element 3)
						$FromServerSite = $xArray[3].SubString($xArray[3].IndexOf("=")+1) #get past the = sign
						$xArray = $Null
						$obj = New-Object -TypeName PSObject
						$obj | Add-Member -MemberType NoteProperty -Name Name           -Value $ConnectionObject.Name
						$obj | Add-Member -MemberType NoteProperty -Name ToServer       -Value $ToServer
						$obj | Add-Member -MemberType NoteProperty -Name FromServer     -Value $FromServer
						$obj | Add-Member -MemberType NoteProperty -Name FromServerSite -Value $FromServerSite
						$Connections += $obj
					}
				}

				If($Null -ne $Connections)
				{
					$Connections = $Connections | Sort-Object Name, ToServer, FromServer
				}
				
				#list each server
				$serverContainerDN = "CN=Servers,CN=" + $siteName + "," + $siteContainerDN
				$SiteServers = $Null
				$SiteServers = Get-ADObject -SearchBase $serverContainerDN -SearchScope OneLevel `
				-Filter { objectClass -eq "Server" } -Properties "DNSHostName" -Server $ADForest -EA 0 | `
				Select-Object DNSHostName, Name | Sort-Object DNSHostName
				
				If($? -and $Null -ne $SiteServers)
				{
					$First = $True
					ForEach($SiteServer in $SiteServers)
					{
						If(!$First)
						{
							WriteWordLine 0 0 ""
						}
						WriteWordLine 0 0 $SiteServer.DNSHostName
						#for each server list each connection object
						If($Null -ne $Connections)
						{
							$Results = $Connections | Where-Object {$_.ToServer -eq $SiteServer.Name}

							If($? -and $Null -ne $Results)
							{
								WriteWordLine 0 1 "Connection Objects to source server $($SiteServer.Name)"
								[System.Collections.Hashtable[]] $WordTableRowHash = @();
								ForEach($Result in $Results)
								{
									$WordTableRowHash += @{ 
									ConnectionName = $Result.Name; 
									FromServer = $Result.FromServer; 
									FromSite = $Result.FromServerSite
									}
								}
								$Table = AddWordTable -Hashtable $WordTableRowHash `
								-Columns ConnectionName, FromServer, FromSite `
								-Headers "Name", "From Server", "From Site" `
								-AutoFit $wdAutoFitFixed;

								SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

								$Table.Columns.Item(1).Width = 200;
								$Table.Columns.Item(2).Width = 100;
								$Table.Columns.Item(3).Width = 100;

								#indent the entire table 1 tab stop
								$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)

								FindWordDocumentEnd
								$Table = $Null
								WriteWordLine 0 0 ""
							}
						}
						Else
						{
							WriteWordLine 0 3 "Connection Objects: "
							WriteWordLine 0 4 "<None>"
							WriteWordLine 0 0 ""
						}
						$First = $False
					}
				}
				ElseIf(!$?)
				{
					Write-Warning "No Site Servers were retrieved."
					WriteWordLine 0 0 "Warning: No Site Servers were retrieved" "" $Null 0 $False $True
					WriteWordLine 0 0 ""
				}
				Else
				{
					WriteWordLine 0 0 "No servers in this site"
					WriteWordLine 0 0 ""
				}
			}
		}
		ElseIf($Text)
		{
			Line 0 "///  Inter-Site Transports  \\\"
			If($? -and $Null -ne $AllSiteLinks)
			{
				ForEach($SiteLink in $AllSiteLinks)
				{
					Write-Verbose "$(Get-Date): `t`tProcessing site link $($SiteLink.Name)"
					$SiteLinkTypeDN = @()
					$SiteLinkTypeDN = $SiteLink.DistinguishedName.Split(",")
					$SiteLinkType = $SiteLinkTypeDN[1].SubString(3)
					$SitesInLink = New-Object System.Collections.ArrayList
					$SiteLinkSiteList = $SiteLink.SiteList
					ForEach($xSite in $SiteLinkSiteList)
					{
						$tmp = $xSite.Split(",")
						$SitesInLink.Add("$($tmp[0].SubString(3))") > $Null
					}
					
					Line 0 "Name`t`t`t: " $SiteLink.Name
					If(![String]::IsNullOrEmpty($SiteLink.Description))
					{
						Line 0 "Description`t`t: " $SiteLink.Description
					}
					Line 0 "Sites in Link`t`t: " -NoNewLine
					If($SitesInLink -ne "")
					{
						$cnt = 0
						ForEach($xSite in $SitesInLink)
						{
							$cnt++
							
							If($cnt -eq 1)
							{
								Line 0 $xSite
							}
							Else
							{
								Line 3 "  $($xSite)"
							}
						}
					}
					Else
					{
						Line 0 "<None>"
					}
					Line 0 "Cost`t`t`t: " $SiteLink.Cost.ToString()
					Line 0 "Replication Interval`t: " $SiteLink.ReplInterval.ToString()
					Line 0 "Schedule`t`t: " $SiteLink.Schedule
					Line 0 "Options`t`t`t: " -NoNewLine
					#https://msdn.microsoft.com/en-us/library/cc223552.aspx
					If([String]::IsNullOrEmpty($SiteLink.Options) -or $SiteLink.Options -eq "0")
					{
						Line 0 "Change Notification is Disabled"
					}
					ElseIf($SiteLink.Options -eq "1")
					{
						Line 0 "Change Notification is Enabled with Compression"
					}
					ElseIf($SiteLink.Options -eq "2")
					{
						Line 0 "Force sync in opposite direction at end of sync"
					}
					ElseIf($SiteLink.Options -eq "3")
					{
						Line 0 "Change Notification is Enabled with Compression and Force sync in opposite direction at end of sync"
					}
					ElseIf($SiteLink.Options -eq "4")
					{
						Line 0 "Disable compression of Change Notification messages"
					}
					ElseIf($SiteLink.Options -eq "5")
					{
						Line 0 "Change Notification is Enabled without Compression"
					}
					ElseIf($SiteLink.Options -eq "6")
					{
						Line 0 "Force sync in opposite direction at end of sync and Disable compression of Change Notification messages"
					}
					ElseIf($SiteLink.Options -eq "7")
					{
						Line 0 "Change Notification is Enabled without Compression and Force sync in opposite direction at end of sync"
					}
					Else
					{
						Line 0 "Unknown"
					}
					Line 0 "Type`t`t`t: " $SiteLinkType
					Line 0 ""
				}
			}
			$AllSiteLinks = $Null
			
			ForEach($Site in $Sites)
			{
				Write-Verbose "$(Get-Date): `tProcessing site $($Site.Name)"
				Line 0 "///  Site: $($Site.Name)  \\\"

				Line 1 "Subnets"
				Write-Verbose "$(Get-Date): `t`tProcessing subnets"
				$subnetArray = New-Object -Type string[] -ArgumentList $Site.siteObjectBL.Count
				$i = 0
				$SitesiteObjectBL = $Site.siteObjectBL
				ForEach($subnetDN in $SitesiteObjectBL) 
				{
					$subnetName = $subnetDN.SubString(3, $subnetDN.IndexOf(",CN=Subnets,CN=Sites,") - 3)
					$subnetArray[$i] = $subnetName
					$i++
				}
				$subnetArray = $subnetArray | Sort-Object 
				If($Null -eq $subnetArray)
				{
					Line 2 "<None>"
				}
				Else
				{
					ForEach($xSubnet in $subnetArray)
					{
						Line 2 $xSubnet
					}
				}
				Line 0 ""
				
				Write-Verbose "$(Get-Date): `t`tProcessing servers"
				Line 1 "Servers"
				$siteName = $Site.Name
				
				#build array of connect objects
				Write-Verbose "$(Get-Date): `t`t`tProcessing automatic connection objects"
				$Connections = New-Object System.Collections.ArrayList
				$ConnectionObjects = $Null
				$ConnectionObjects = Get-ADObject -Filter 'objectClass -eq "nTDSConnection" -and options -bor 1' `
				-Searchbase $Script:ConfigNC -Property DistinguishedName, fromServer -Server $ADForest -EA 0
				
				If($? -and $Null -ne $ConnectionObjects)
				{
					ForEach($ConnectionObject in $ConnectionObjects)
					{
						$xArray = $ConnectionObject.DistinguishedName.Split(",")
						#server name is 3rd item in array (element 2)
						$ToServer = $xArray[2].SubString($xArray[2].IndexOf("=")+1) #get past the = sign
						$xArray = $ConnectionObject.FromServer.Split(",")
						#server name is 2nd item in array (element 1)
						$FromServer = $xArray[1].SubString($xArray[1].IndexOf("=")+1) #get past the = sign
						#site name is 4th item in array (element 3)
						$FromServerSite = $xArray[3].SubString($xArray[3].IndexOf("=")+1) #get past the = sign
						$xArray = $Null
						$obj = New-Object -TypeName PSObject
						$obj | Add-Member -MemberType NoteProperty -Name Name           -Value "<automatically generated>"
						$obj | Add-Member -MemberType NoteProperty -Name ToServer       -Value $ToServer
						$obj | Add-Member -MemberType NoteProperty -Name FromServer     -Value $FromServer
						$obj | Add-Member -MemberType NoteProperty -Name FromServerSite -Value $FromServerSite
						$Connections.Add($obj) > $Null
					}
				}
				
				Write-Verbose "$(Get-Date): `t`t`tProcessing manual connection objects"
				$ConnectionObjects = $Null
				$ConnectionObjects = Get-ADObject -Filter 'objectClass -eq "nTDSConnection" -and -not options -bor 1' `
				-Searchbase $Script:ConfigNC -Property Name, DistinguishedName, fromServer -Server $ADForest -EA 0
				
				If($? -and $Null -ne $ConnectionObjects)
				{
					ForEach($ConnectionObject in $ConnectionObjects)
					{
						$xArray = $ConnectionObject.DistinguishedName.Split(",")
						#server name is 3rd item in array (element 2)
						$ToServer = $xArray[2].SubString($xArray[2].IndexOf("=")+1) #get past the = sign
						$xArray = $ConnectionObject.FromServer.Split(",")
						#server name is 2nd item in array (element 1)
						$FromServer = $xArray[1].SubString($xArray[1].IndexOf("=")+1) #get past the = sign
						#site name is 4th item in array (element 3)
						$FromServerSite = $xArray[3].SubString($xArray[3].IndexOf("=")+1) #get past the = sign
						$xArray = $Null
						$obj = New-Object -TypeName PSObject
						$obj | Add-Member -MemberType NoteProperty -Name Name           -Value $ConnectionObject.Name
						$obj | Add-Member -MemberType NoteProperty -Name ToServer       -Value $ToServer
						$obj | Add-Member -MemberType NoteProperty -Name FromServer     -Value $FromServer
						$obj | Add-Member -MemberType NoteProperty -Name FromServerSite -Value $FromServerSite
						$Connections += $obj
					}
				}

				If($Null -ne $Connections)
				{
					$Connections = $Connections | Sort-Object Name, ToServer, FromServer
				}
				
				#list each server
				$serverContainerDN = "CN=Servers,CN=" + $siteName + "," + $siteContainerDN
				$SiteServers = $Null
				$SiteServers = Get-ADObject -SearchBase $serverContainerDN -SearchScope OneLevel `
				-Filter { objectClass -eq "Server" } -Properties "DNSHostName" -Server $ADForest -EA 0 | `
				Select-Object DNSHostName, Name | Sort-Object DNSHostName
				
				If($? -and $Null -ne $SiteServers)
				{
					ForEach($SiteServer in $SiteServers)
					{
						Line 2 $SiteServer.DNSHostName
						Line 0 ""
						#for each server list each connection object
						If($Null -ne $Connections)
						{
							$Results = $Connections | Where-Object {$_.ToServer -eq $SiteServer.Name}

							If($? -and $Null -ne $Results)
							{
								Line 2 "Connection Objects to source server $($SiteServer.Name)"
								ForEach($Result in $Results)
								{
									Line 3 "Name`t`t: " $Result.Name
									Line 3 "From Server`t: " $Result.FromServer
									Line 3 "From Site`t: " $Result.FromServerSite
									Line 0 ""
								}
							}
						}
						Else
						{
							Line 3 "Connection Objects: <None>"
						}
					}
				}
				ElseIf(!$?)
				{
					Write-Warning "No Site Servers were retrieved."
					Line 2 "Warning: No Site Servers were retrieved"
					Line 0 ""
				}
				Else
				{
					Line 2 "No servers in this site"
					Line 0 ""
				}
			}
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 2 0 "///&nbsp;&nbsp;Inter-Site Transports&nbsp;&nbsp;\\\"
			#adapted from code provided by Goatee PFE
			#http://blogs.technet.com/b/ashleymcglone/archive/2011/06/29/report-and-edit-ad-site-links-from-powershell-turbo-your-ad-replication.aspx
			# Report of all site links and related settings
			
			If($? -and $Null -ne $AllSiteLinks)
			{
				ForEach($SiteLink in $AllSiteLinks)
				{
					Write-Verbose "$(Get-Date): `t`tProcessing site link $($SiteLink.Name)"
					$rowdata = @()
					$SiteLinkTypeDN = @()
					$SiteLinkTypeDN = $SiteLink.DistinguishedName.Split(",")
					$SiteLinkType = $SiteLinkTypeDN[1].SubString(3)
					$SitesInLink = New-Object System.Collections.ArrayList
					$SiteLinkSiteList = $SiteLink.SiteList
					ForEach($xSite in $SiteLinkSiteList)
					{
						$tmp = $xSite.Split(",")
						$SitesInLink.Add("$($tmp[0].SubString(3))") > $Null
					}
					
					$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$SiteLink.Name,$htmlwhite)
					If(![String]::IsNullOrEmpty($SiteLink.Description))
					{
						$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$SiteLink.Description,$htmlwhite))
					}
					If($SitesInLink -ne "")
					{
						$cnt = 0
						ForEach($xSite in $SitesInLink)
						{
							$cnt++
							
							If($cnt -eq 1)
							{
								$rowdata += @(,('Sites in Link',($htmlsilver -bor $htmlbold),$xSite,$htmlwhite))
							}
							Else
							{
								$rowdata += @(,('',($htmlsilver -bor $htmlbold),$xSite,$htmlwhite))
							}
						}
					}
					Else
					{
						$rowdata += @(,('Sites in Link',($htmlsilver -bor $htmlbold),"None",$htmlwhite))
					}
					$rowdata += @(,('Cost',($htmlsilver -bor $htmlbold),$SiteLink.Cost.ToString(),$htmlwhite))
					$rowdata += @(,('Replication Interval',($htmlsilver -bor $htmlbold),$SiteLink.ReplInterval.ToString(),$htmlwhite))
					$rowdata += @(,('Schedule',($htmlsilver -bor $htmlbold),$SiteLink.Schedule,$htmlwhite))

					#https://msdn.microsoft.com/en-us/library/cc223552.aspx
					$tmp = ""
					If([String]::IsNullOrEmpty($SiteLink.Options) -or $SiteLink.Options -eq "0")
					{
						$tmp = "Change Notification is Disabled"
					}
					ElseIf($SiteLink.Options -eq "1")
					{
						$tmp = "Change Notification is Enabled with Compression"
					}
					ElseIf($SiteLink.Options -eq "2")
					{
						$tmp = "Force sync in opposite direction at end of sync"
					}
					ElseIf($SiteLink.Options -eq "3")
					{
						$tmp = "Change Notification is Enabled with Compression and Force sync in opposite direction at end of sync"
					}
					ElseIf($SiteLink.Options -eq "4")
					{
						$tmp = "Disable compression of Change Notification messages"
					}
					ElseIf($SiteLink.Options -eq "5")
					{
						$tmp = "Change Notification is Enabled without Compression"
					}
					ElseIf($SiteLink.Options -eq "6")
					{
						$tmp = "Force sync in opposite direction at end of sync and Disable compression of Change Notification messages"
					}
					ElseIf($SiteLink.Options -eq "7")
					{
						$tmp = "Change Notification is Enabled without Compression and Force sync in opposite direction at end of sync"
					}
					Else
					{
						$tmp = "Unknown"
					}
					$rowdata += @(,('Options',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					$rowdata += @(,('Type',($htmlsilver -bor $htmlbold),$SiteLinkType,$htmlwhite))
					$msg = ""
					$columnWidths = @("200","250")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
					WriteHTMLLine 0 0 " "
				}
			}
			$AllSiteLinks = $Null
			
			ForEach($Site in $Sites)
			{
				Write-Verbose "$(Get-Date): `tProcessing site $($Site.Name)"
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;Site: $($Site.Name)&nbsp;&nbsp;\\\"

				WriteHTMLLine 3 0 "Subnets"
				Write-Verbose "$(Get-Date): `t`tProcessing subnets"
				$subnetArray = New-Object -Type string[] -ArgumentList $Site.siteObjectBL.Count
				$i = 0
				$SitesiteObjectBL = $Site.siteObjectBL
				ForEach($subnetDN in $SitesiteObjectBL) 
				{
					$subnetName = $subnetDN.SubString(3, $subnetDN.IndexOf(",CN=Subnets,CN=Sites,") - 3)
					$subnetArray[$i] = $subnetName
					$i++
				}
				$subnetArray = $subnetArray | Sort-Object 
				If($Null -eq $subnetArray)
				{
					WriteHTMLLine 0 0 "None"
				}
				Else
				{
					$rowdata = @()
					ForEach($xSubnet in $subnetArray)
					{
						$rowdata += @(,($xSubnet,$htmlwhite))
					}
					$columnHeaders = @('Subnets',($htmlsilver -bor $htmlbold))
					$msg = ""
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
					WriteHTMLLine 0 0 " "
				}
				
				Write-Verbose "$(Get-Date): `t`tProcessing servers"
				WriteHTMLLine 3 0 "Servers"
				$siteName = $Site.Name
				
				#build array of connect objects
				Write-Verbose "$(Get-Date): `t`t`tProcessing automatic connection objects"
				$Connections = New-Object System.Collections.ArrayList
				$ConnectionObjects = $Null
				$ConnectionObjects = Get-ADObject -Filter 'objectClass -eq "nTDSConnection" -and options -bor 1' `
				-Searchbase $Script:ConfigNC -Property DistinguishedName, fromServer -Server $ADForest -EA 0
				
				If($? -and $Null -ne $ConnectionObjects)
				{
					ForEach($ConnectionObject in $ConnectionObjects)
					{
						$xArray = $ConnectionObject.DistinguishedName.Split(",")
						#server name is 3rd item in array (element 2)
						$ToServer = $xArray[2].SubString($xArray[2].IndexOf("=")+1) #get past the = sign
						$xArray = $ConnectionObject.FromServer.Split(",")
						#server name is 2nd item in array (element 1)
						$FromServer = $xArray[1].SubString($xArray[1].IndexOf("=")+1) #get past the = sign
						#site name is 4th item in array (element 3)
						$FromServerSite = $xArray[3].SubString($xArray[3].IndexOf("=")+1) #get past the = sign
						$xArray = $Null
						$obj = New-Object -TypeName PSObject
						$obj | Add-Member -MemberType NoteProperty -Name Name           -Value "<automatically generated>"
						$obj | Add-Member -MemberType NoteProperty -Name ToServer       -Value $ToServer
						$obj | Add-Member -MemberType NoteProperty -Name FromServer     -Value $FromServer
						$obj | Add-Member -MemberType NoteProperty -Name FromServerSite -Value $FromServerSite
						$Connections.Add($obj) > $Null
					}
				}
				
				Write-Verbose "$(Get-Date): `t`t`tProcessing manual connection objects"
				$ConnectionObjects = $Null
				$ConnectionObjects = Get-ADObject -Filter 'objectClass -eq "nTDSConnection" -and -not options -bor 1' `
				-Searchbase $Script:ConfigNC -Property Name, DistinguishedName, fromServer -Server $ADForest -EA 0
				
				If($? -and $Null -ne $ConnectionObjects)
				{
					ForEach($ConnectionObject in $ConnectionObjects)
					{
						$xArray = $ConnectionObject.DistinguishedName.Split(",")
						#server name is 3rd item in array (element 2)
						$ToServer = $xArray[2].SubString($xArray[2].IndexOf("=")+1) #get past the = sign
						$xArray = $ConnectionObject.FromServer.Split(",")
						#server name is 2nd item in array (element 1)
						$FromServer = $xArray[1].SubString($xArray[1].IndexOf("=")+1) #get past the = sign
						#site name is 4th item in array (element 3)
						$FromServerSite = $xArray[3].SubString($xArray[3].IndexOf("=")+1) #get past the = sign
						$xArray = $Null
						$obj = New-Object -TypeName PSObject
						$obj | Add-Member -MemberType NoteProperty -Name Name           -Value $ConnectionObject.Name
						$obj | Add-Member -MemberType NoteProperty -Name ToServer       -Value $ToServer
						$obj | Add-Member -MemberType NoteProperty -Name FromServer     -Value $FromServer
						$obj | Add-Member -MemberType NoteProperty -Name FromServerSite -Value $FromServerSite
						$Connections += $obj
					}
				}

				If($Null -ne $Connections)
				{
					$Connections = $Connections | Sort-Object Name, ToServer, FromServer
				}
				
				#list each server
				$serverContainerDN = "CN=Servers,CN=" + $siteName + "," + $siteContainerDN
				$SiteServers = $Null
				$SiteServers = Get-ADObject -SearchBase $serverContainerDN -SearchScope OneLevel `
				-Filter { objectClass -eq "Server" } -Properties "DNSHostName" -Server $ADForest -EA 0 | `
				Select-Object DNSHostName, Name | Sort-Object DNSHostName
				
				If($? -and $Null -ne $SiteServers)
				{
					ForEach($SiteServer in $SiteServers)
					{
						WriteHTMLLine 0 0 $SiteServer.DNSHostName
						WriteHTMLLine 0 0 " "
						#for each server list each connection object
						If($Null -ne $Connections)
						{
							$Results = $Connections | Where-Object {$_.ToServer -eq $SiteServer.Name}

							If($? -and $Null -ne $Results)
							{
								$rowdata = @()
								ForEach($Result in $Results)
								{
									#replace the <> characters since HTML doesn't like those in data
									$tmp = $Result.Name
									$tmp = $tmp.Replace(">","")
									$tmp = $tmp.Replace("<","")
									
									$rowdata += @(,($tmp,$htmlwhite,
													$Result.FromServer,$htmlwhite,
													$Result.FromServerSite,$htmlwhite))
								}
								$columnWidths = @("175px","125px","150px")
								$columnHeaders = @('Name',($htmlsilver -bor $htmlbold),
													'From Server',($htmlsilver -bor $htmlbold),
													'From Site',($htmlsilver -bor $htmlbold))
								$msg = "Connection Objects to source server $($SiteServer.Name)"
								FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "450"
								WriteHTMLLine 0 0 " "
							}
						}
						Else
						{
							WriteHTMLLine 0 3 "Connection Objects: None"
						}
					}
				}
				ElseIf(!$?)
				{
					Write-Warning "No Site Servers were retrieved."
					WriteHTMLLine 0 0 "Warning: No Site Servers were retrieved" "" $Null 0 $False $True
				}
				Else
				{
					WriteHTMLLine 0 0 "No servers in this site"
				}
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "No Sites were retrieved."
		$txt = "Warning: No Sites were retrieved"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			WriteWordLine 0 0 $txt
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 $txt
		}
	}
	Else
	{
		$txt = "There were no sites found to retrieve"
		Write-Warning $txt
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 0 $txt "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			WriteWordLine 0 0 $txt
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 $txt
		}
	}
}
#endregion

#region domains
Function ProcessDomains
{
	Write-Verbose "$(Get-Date): Writing domain data"

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Domain Information"
	}
	ElseIf($Text)
	{
		Line 0 "///  Domain Information  \\\"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Domain Information&nbsp;&nbsp;\\\"
	}

	$Script:AllDomainControllers = New-Object System.Collections.ArrayList
	$First = $True

	#http://technet.microsoft.com/en-us/library/bb125224(v=exchg.150).aspx
	#http://support.microsoft.com/kb/556086/he
	#https://eightwone.com/references/ad-schema-versions/
	#https://eightwone.com/references/schema-versions/
	
	$SchemaVersionTable = @{ 
	"13" = "Windows 2000"; 
	"30" = "Windows 2003 RTM, SP1, SP2"; 
	"31" = "Windows 2003 R2";
	"44" = "Windows 2008"; 
	"47" = "Windows 2008 R2";
	"56" = "Windows Server 2012";
	"69" = "Windows Server 2012 R2";
	"72" = "Windows Server 2016 TP4";
	"87" = "Windows Server 2016";
	"88" = "Windows Server 2019";	#added V2.20, updated in 2.22
	"4397" = "Exchange 2000 RTM"; 
	"4406" = "Exchange 2000 SP3";
	"6870" = "Exchange 2003 RTM, SP1, SP2"; 
	"6936" = "Exchange 2003 SP3"; 
	"10637" = "Exchange 2007 RTM";
	"11116" = "Exchange 2007 SP1"; 
	"14622" = "Exchange 2007 SP2, Exchange 2010 RTM";
	"14625" = "Exchange 2007 SP3";
	"14726" = "Exchange 2010 SP1";
	"14732" = "Exchange 2010 SP2";
	"14734" = "Exchange 2010 SP3";
	"15137" = "Exchange 2013 RTM";
	"15254" = "Exchange 2013 CU1";
	"15281" = "Exchange 2013 CU2";
	"15283" = "Exchange 2013 CU3";
	"15292" = "Exchange 2013 SP1/CU4";
	"15300" = "Exchange 2013 CU5";
	"15303" = "Exchange 2013 CU6";
	"15312" = "Exchange 2013 CU7 through CU23"; #updated in 2.20, updated in 2.24
	"15317" = "Exchange 2016 Preview and RTM"; #updated in 2.24
	"15323" = "Exchange 2016 CU1";
	"15325" = "Exchange 2016 CU2";
	"15326" = "Exchange 2016 CU3/CU4/CU5"; #added in 2.16
	"15330" = "Exchange 2016 CU6"; #added in 2.16
	"15332" = "Exchange 2016 CU7 through CU15"; #added in 2.16 and updated in 2.20, updated in 2.22, updated in 2.24
	"17000" = "Exchange 2019 RTM/CU1"; #added in 2.22, updated in 2.24
	"17001" = "Exchange 2019 CU2-CU4"; #added in 2.24
	}

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date): `tProcessing domain $($Domain)"

		$DomainInfo = Get-ADDomain -Identity $Domain -EA 0
		
		If($? -and $Null -ne $DomainInfo)
		{
			If(($MSWORD -or $PDF) -and !$First)
			{
				#put each domain, starting with the second, on a new page
				$Script:selection.InsertNewPage()
			}
			
			If($Domain -eq $Script:ForestRootDomain)
			{
				If($MSWORD -or $PDF)
				{
					WriteWordLine 2 0 "$($Domain) (Forest Root)"
				}
				ElseIf($Text)
				{
					Line 0 "///  $($Domain) (Forest Root)  \\\"
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($Domain) (Forest Root)&nbsp;&nbsp;\\\"
				}
			}
			Else
			{
				If($MSWORD -or $PDF)
				{
					WriteWordLine 2 0 $Domain
				}
				ElseIf($Text)
				{
					Line 0 "///  $($Domain)  \\\"
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($Domain)&nbsp;&nbsp;\\\"
				}
			}

			Switch ($DomainInfo.DomainMode)
			{
				"0"	{$DomainMode = "Windows 2000"; Break}
				"1" {$DomainMode = "Windows Server 2003 mixed"; Break}
				"2" {$DomainMode = "Windows Server 2003"; Break}
				"3" {$DomainMode = "Windows Server 2008"; Break}
				"4" {$DomainMode = "Windows Server 2008 R2"; Break}
				"5" {$DomainMode = "Windows Server 2012"; Break}
				"6" {$DomainMode = "Windows Server 2012 R2"; Break}
				"7" {$DomainMode = "Windows Server 2016"; Break}	#added V2.20
				"Windows2000Domain"   		{$DomainMode = "Windows 2000"; Break}
				"Windows2003Mixed"    		{$DomainMode = "Windows Server 2003 mixed"; Break}
				"Windows2003Domain"   		{$DomainMode = "Windows Server 2003"; Break}
				"Windows2008Domain"   		{$DomainMode = "Windows Server 2008"; Break}
				"Windows2008R2Domain" 		{$DomainMode = "Windows Server 2008 R2"; Break}
				"Windows2012Domain"   		{$DomainMode = "Windows Server 2012"; Break}
				"Windows2012R2Domain" 		{$DomainMode = "Windows Server 2012 R2"; Break}
				"WindowsThresholdDomain"	{$DomainMode = "Windows Server 2016 TP"; Break}
				"Windows2016Domain"			{$DomainMode = "Windows Server 2016"; Break}
				"UnknownDomain"       		{$DomainMode = "Unknown Domain Mode"; Break}
				Default               		{$DomainMode = "Unable to determine Domain Mode: $($DomainInfo.DomainMode)"; Break}
			}
			
			#http://blogs.technet.com/b/poshchap/archive/2014/03/07/ad-schema-version.aspx
			$ADSchemaInfo = $Null
			$ExchangeSchemaInfo = $Null
			
			$ADSchemaInfo = Get-ADObject (Get-ADRootDSE -Server $Domain -EA 0).schemaNamingContext `
			-Property objectVersion -Server $Domain -EA 0
			
			If($? -and $Null -ne $ADSchemaInfo)
			{
				$ADSchemaVersion = $ADSchemaInfo.objectversion
				$ADSchemaVersionName = $SchemaVersionTable.Get_Item("$ADSchemaVersion")
				If($Null -eq $ADSchemaVersionName)
				{
					$ADSchemaVersionName = "Unknown"
				}
			}
			Else
			{
				$ADSchemaVersion = "Unknown"
				$ADSchemaVersionName = "Unknown"
			}
			
			If($Domain -eq $Script:ForestRootDomain)
			{
				$ExchangeSchemaInfo = Get-ADObject "cn=ms-exch-schema-version-pt,cn=Schema,cn=Configuration,$($DomainInfo.DistinguishedName)" `
				-properties rangeupper -Server $Domain -EA 0

				If($? -and $Null -ne $ExchangeSchemaInfo)
				{
					$ExchangeSchemaVersion = $ExchangeSchemaInfo.rangeupper
					$ExchangeSchemaVersionName = $SchemaVersionTable.Get_Item("$ExchangeSchemaVersion")
					If($Null -eq $ExchangeSchemaVersionName)
					{
						$ExchangeSchemaVersionName = "Unknown"
					}
				}
				Else
				{
					$ExchangeSchemaVersion = "Unknown"
					$ExchangeSchemaVersionName = "Unknown"
				}
			}

			If($Null -eq $DomainInfo.LastLogonReplicationInterval)
			{
				$LastLogonReplicationInterval = "Default 1 day"
			}
			Else
			{
				$LastLogonReplicationInterval = $DomainInfo.LastLogonReplicationInterval.ToString()
			}
			
			If($MSWORD -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Domain mode"; Value = $DomainMode; }
				$ScriptInformation += @{ Data = "Domain name"; Value = $DomainInfo.Name; }
				$ScriptInformation += @{ Data = "NetBIOS name"; Value = $DomainInfo.NetBIOSName; }
				#V2.20 reorder the following properties in alpha order
				$ScriptInformation += @{ Data = "AD Schema"; Value = "($($ADSchemaVersion)) - $($ADSchemaVersionName)"; }
				$DNSSuffixes = $DomainInfo.AllowedDNSSuffixes | Sort-Object 
				If($Null -eq $DNSSuffixes)
				{
					$ScriptInformation += @{ Data = "Allowed DNS Suffixes"; Value = "<None>"; }
				}
				Else
				{
					$cnt = 0
					ForEach($DNSSuffix in $DNSSuffixes)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$ScriptInformation += @{ Data = "Allowed DNS Suffixes"; Value = "$($DNSSuffix.ToString())"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = ""; Value = "$($DNSSuffix.ToString())"; }
						}
					}
				}
				$ChildDomains = $DomainInfo.ChildDomains | Sort-Object 
				If($Null -eq $ChildDomains)
				{
					$ScriptInformation += @{ Data = "Child domains"; Value = "<None>"; }
				}
				Else
				{
					$cnt = 0 
					ForEach($ChildDomain in $ChildDomains)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$ScriptInformation += @{ Data = "Child domains"; Value = "$($ChildDomain.ToString())"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = ""; Value = "$($ChildDomain.ToString())"; }
						}
					}
				}
				$ScriptInformation += @{ Data = "Default computers container"; Value = $DomainInfo.ComputersContainer; }
				$ScriptInformation += @{ Data = "Default users container"; Value = $DomainInfo.UsersContainer; }
				$ScriptInformation += @{ Data = "Deleted objects container"; Value = $DomainInfo.DeletedObjectsContainer; }
				$ScriptInformation += @{ Data = "Distinguished name"; Value = $DomainInfo.DistinguishedName; }
				$ScriptInformation += @{ Data = "DNS root"; Value = $DomainInfo.DNSRoot; }
				$ScriptInformation += @{ Data = "Domain controllers container"; Value = $DomainInfo.DomainControllersContainer; }
				If(![String]::IsNullOrEmpty($ExchangeSchemaInfo))
				{
					$ScriptInformation += @{ Data = "Exchange Schema"; Value = "($($ExchangeSchemaVersion)) - $($ExchangeSchemaVersionName)"; }
				}
				$ScriptInformation += @{ Data = "Foreign security principals container"; Value = $DomainInfo.ForeignSecurityPrincipalsContainer; }
				$ScriptInformation += @{ Data = "Infrastructure master"; Value = $DomainInfo.InfrastructureMaster; }
				#V2.20 added
				$ScriptInformation += @{ Data = "Last logon replication interval"; Value = $LastLogonReplicationInterval; }
				$ScriptInformation += @{ Data = "Lost and Found container"; Value = $DomainInfo.LostAndFoundContainer; }
				If(![String]::IsNullOrEmpty($DomainInfo.ManagedBy))
				{
					$ScriptInformation += @{ Data = "Managed by"; Value = $DomainInfo.ManagedBy; }
				}
				$ScriptInformation += @{ Data = "PDC Emulator"; Value = $DomainInfo.PDCEmulator; }
				#V2.20 added
				If(validObject $DomainInfo PublicKeyRequiredPasswordRolling)
				{
					$ScriptInformation += @{ Data = "Public key required password rolling"; Value = $DomainInfo.PublicKeyRequiredPasswordRolling.ToString(); }
				}
				$ScriptInformation += @{ Data = "Quotas container"; Value = $DomainInfo.QuotasContainer; }
				$ReadOnlyReplicas = $DomainInfo.ReadOnlyReplicaDirectoryServers | Sort-Object 
				If($Null -eq $ReadOnlyReplicas)
				{
					$ScriptInformation += @{ Data = "Read-only replica directory servers"; Value = "<None>"; }
				}
				Else
				{
					$cnt = 0 
					ForEach($ReadOnlyReplica in $ReadOnlyReplicas)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$ScriptInformation += @{ Data = "Read-only replica directory servers"; Value = "$($ReadOnlyReplica.ToString())"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = ""; Value = "$($ReadOnlyReplica.ToString())"; }
						}
					}
				}
				$Replicas = $DomainInfo.ReplicaDirectoryServers | Sort-Object 
				If($Null -eq $Replicas)
				{
					$ScriptInformation += @{ Data = "Replica directory servers"; Value = "<None>"; }
				}
				Else
				{
					$cnt = 0 
					ForEach($Replica in $Replicas)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$ScriptInformation += @{ Data = "Replica directory servers"; Value = "$($Replica.ToString())"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = ""; Value = "$($Replica.ToString())"; }
						}
					}
				}
				$ScriptInformation += @{ Data = "RID Master"; Value = $DomainInfo.RIDMaster; }
				$SubordinateReferences = $DomainInfo.SubordinateReferences | Sort-Object 
				If($Null -eq $SubordinateReferences)
				{
					$ScriptInformation += @{ Data = "Subordinate references"; Value = "<None>"; }
				}
				Else
				{
					$cnt = 0
					ForEach($SubordinateReference in $SubordinateReferences)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$ScriptInformation += @{ Data = "Subordinate references"; Value = "$($SubordinateReference.ToString())"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = ""; Value = "$($SubordinateReference.ToString())"; }
						}
					}
				}
				$ScriptInformation += @{ Data = "Systems container"; Value = $DomainInfo.SystemsContainer; }
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 175;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				Write-Verbose "$(Get-Date): `t`tGetting domain trusts"
				WriteWordLine 3 0 "Domain Trusts"
				
				$ADDomainTrusts = $Null
				$ADDomainTrusts = Get-ADObject -Filter {ObjectClass -eq "trustedDomain"} `
				-Server $Domain -Properties * -EA 0

				If($? -and $Null -ne $ADDomainTrusts)
				{
					
					ForEach($Trust in $ADDomainTrusts) 
					{ 
						[System.Collections.Hashtable[]] $ScriptInformation = @()
						$ScriptInformation += @{ Data = "Name"; Value = $Trust.Name; }
						
						If(![String]::IsNullOrEmpty($Trust.Description))
						{
							$ScriptInformation += @{ Data = "Description"; Value = $Trust.Description; }
						}
						
						$ScriptInformation += @{ Data = "Created"; Value = $Trust.Created; }
						$ScriptInformation += @{ Data = "Modified"; Value = $Trust.Modified; }

						$TrustExtendedAttributes = Get-ADTrustInfo $Trust
						
						$ScriptInformation += @{ Data = "Type"; Value = $TrustExtendedAttributes.TrustType; }

						$cnt = 0
						ForEach($attribute in $TrustExtendedAttributes.TrustAttribute)
						{
							$cnt++
							
							If($cnt -eq 1)
							{
								$ScriptInformation += @{ Data = "Attributes"; Value = $attribute.ToString(); }
							}
							Else
							{
								$ScriptInformation += @{ Data = ""; Value = "$($attribute.ToString())"; }
							}
						}

						$ScriptInformation += @{ Data = "Direction"; Value = $TrustDirection; }

						$Table = AddWordTable -Hashtable $ScriptInformation `
						-Columns Data,Value `
						-List `
						-Format $wdTableGrid `
						-AutoFit $wdAutoFitFixed;

						SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						$Table.Columns.Item(1).Width = 175;
						$Table.Columns.Item(2).Width = 300;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

						FindWordDocumentEnd
						$Table = $Null
						WriteWordLine 0 0 ""
					}
				}
				ElseIf(!$?)
				{
					#error retrieving domain trusts
					Write-Warning "Error retrieving domain trusts for $($Domain)"
					WriteWordLine 0 0 "Error retrieving domain trusts for $($Domain)" "" $Null 0 $False $True
				}
				Else
				{
					#no domain trust data
					WriteWordLine 0 0 "<None>"
				}
				
				Write-Verbose "$(Get-Date): `t`tProcessing domain controllers"
				$DomainControllers = $Null
				$DomainControllers = Get-ADDomainController -Filter * -Server $DomainInfo.DNSRoot -EA 0 | Sort-Object Name
				
				If($? -and $Null -ne $DomainControllers)
				{
					$Script:AllDomainControllers.Add($DomainControllers) > $Null
					[System.Collections.Hashtable[]] $WordTable = @();
					WriteWordLine 3 0 "Domain Controllers"
					ForEach($DomainController in $DomainControllers)
					{
						$WordTableRowHash = @{
						DCName = $DomainController.Name; 
						}
						$WordTable += $WordTableRowHash;
					}
					#set column widths
					$Table = AddWordTable -Hashtable $WordTable `
					-Columns  DCName `
					-Headers "Name" `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 105;
					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				ElseIf(!$?)
				{
					Write-Warning "Error retrieving domain controller data for domain $($Domain)"
					WriteWordLine 0 0 "Error retrieving domain controller data for domain $($Domain)" "" $Null 0 $False $True
				}
				Else
				{
					WriteWordLine 0 0 "No Domain controller data was retrieved for domain $($Domain)" "" $Null 0 $False $True
				}

				Write-Verbose "$(Get-Date): `t`tProcessing Fine Grained Password Policies"
				
				#are FGPP cmdlets available
				If(Get-Command -Name "Get-ADFineGrainedPasswordPolicy" -ea 0)
				{
					$FGPPs = $Null
					$FGPPs = Get-ADFineGrainedPasswordPolicy -Searchbase $DomainInfo.DistinguishedName -Filter * -Properties * -Server $DomainInfo.DNSRoot -EA 0 | Sort-Object Precedence, ObjectGUID
					
					If($? -and $Null -ne $FGPPs)
					{
						WriteWordLine 3 0 "Fine Grained Password Policies"
						
						ForEach($FGPP in $FGPPs)
						{
							[System.Collections.Hashtable[]] $ScriptInformation = @()
							$ScriptInformation += @{ Data = "Name"; Value = $FGPP.Name; }
							$ScriptInformation += @{ Data = "Precedence"; Value = $FGPP.Precedence.ToString(); }
							
							If($FGPP.MinPasswordLength -eq 0)
							{
								$ScriptInformation += @{ Data = "Enforce minimum password length"; Value = "Not enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "Enforce minimum password length"; Value = "Enabled"; }
								$ScriptInformation += @{ Data = "     Minimum password length (characters)"; Value = $FGPP.MinPasswordLength.ToString(); }
							}
							
							If($FGPP.PasswordHistoryCount -eq 0)
							{
								$ScriptInformation += @{ Data = "Enforce password history"; Value = "Not enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "Enforce password history"; Value = "Enabled"; }
								$ScriptInformation += @{ Data = "     Number of passwords remembered"; Value = $FGPP.PasswordHistoryCount.ToString(); }
							}
							
							If($FGPP.ComplexityEnabled -eq $True)
							{
								$ScriptInformation += @{ Data = "Password must meet complexity requirements"; Value = "Enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "Password must meet complexity requirements"; Value = "Not enabled"; }
							}
							
							If($FGPP.ReversibleEncryptionEnabled -eq $True)
							{
								$ScriptInformation += @{ Data = "Store password using reversible encryption"; Value = "Enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "Store password using reversible encryption"; Value = "Not enabled"; }
							}
							
							If($FGPP.ProtectedFromAccidentalDeletion -eq $True)
							{
								$ScriptInformation += @{ Data = "Protect from accidental deletion"; Value = "Enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "Protect from accidental deletion"; Value = "Not enabled"; }
							}
							
							$ScriptInformation += @{ Data = "Password age options"; Value = ""; }
							If($FGPP.MinPasswordAge.Days -eq 0)
							{
								$ScriptInformation += @{ Data = "     Enforce minimum password age"; Value = "Not enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "     Enforce minimum password age"; Value = "Enabled"; }
								$ScriptInformation += @{ Data = "          User cannot change the password within (days)"; Value = $FGPP.MinPasswordAge.TotalDays.ToString(); }
							}
							
							If($FGPP.MaxPasswordAge -eq 0)
							{
								$ScriptInformation += @{ Data = "     Enforce maximum password age"; Value = "Not enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "     Enforce maximum password age"; Value = "Enabled"; }
								$ScriptInformation += @{ Data = "          User must change the password after (days)"; Value = $FGPP.MaxPasswordAge.TotalDays.ToString(); }
							}
							
							If($FGPP.LockoutThreshold -eq 0)
							{
								$ScriptInformation += @{ Data = "Enforce account lockout policy"; Value = "Not enabled"; }
							}
							Else
							{
								$ScriptInformation += @{ Data = "Enforce account lockout policy"; Value = "Enabled"; }
								$ScriptInformation += @{ Data = "     Number of failed logon attempts allowed"; Value = $FGPP.LockoutThreshold.ToString(); }
								$ScriptInformation += @{ Data = "     Reset failed logon attempts count after (mins)"; Value = $FGPP.LockoutObservationWindow.TotalMinutes.ToString(); }
								If($FGPP.LockoutDuration -eq 0)
								{
									$ScriptInformation += @{ Data = "     Account will be locked out"; Value = ""; }
									$ScriptInformation += @{ Data = "          Until an administrator manually unlocks the account"; Value = ""; }
								}
								Else
								{
									$ScriptInformation += @{ Data = "     Account will be locked out for a duration of (mins)"; Value = $FGPP.LockoutDuration.TotalMinutes.ToString(); }
								}
								
							}
							
							$ScriptInformation += @{ Data = "Description"; Value = $FGPP.Description; }
							
							$results = Get-ADFineGrainedPasswordPolicySubject -Identity $FGPP.Name -EA 0 | Sort-Object Name
							
							If($? -and $Null -ne $results)
							{
								$cnt = 0
								ForEach($Item in $results)
								{
									$cnt++
									
									If($cnt -eq 1)
									{
										$ScriptInformation += @{ Data = "Directly Applies To"; Value = $Item.Name; }
									}
									Else
									{
										$ScriptInformation += @{ Data = ""; Value = $($Item.Name); }
									}
								}
							}
							Else
							{
							}
							
							
							$Table = AddWordTable -Hashtable $ScriptInformation `
							-Columns Data,Value `
							-List `
							-Format $wdTableGrid `
							-AutoFit $wdAutoFitFixed;

							SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

							$Table.Columns.Item(1).Width = 275;
							$Table.Columns.Item(2).Width = 200;

							$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

							FindWordDocumentEnd
							$Table = $Null
							WriteWordLine 0 0 ""
						}
					}
					ElseIf(!$?)
					{
						Write-Warning "Error retrieving Fine Grained Password Policy data for domain $($Domain)"
						WriteWordLine 0 0 "Error retrieving Fine Grained Password Policy data for domain $($Domain)"
					}
					Else
					{
						WriteWordLine 0 0 "No Fine Grained Password Policy data was retrieved for domain $($Domain)"
					}
					WriteWordLine 0 0 ""
				}
				Else
				{
					#FGPP cmdlets are not available
				}
			}
			ElseIf($Text)
			{
				Line 1 "Domain mode`t`t`t`t: " $DomainMode
				Line 1 "Domain name`t`t`t`t: " $DomainInfo.Name
				Line 1 "NetBIOS name`t`t`t`t: " $DomainInfo.NetBIOSName
				#V2.20 reorder the following properties in alpha order
				Line 1 "AD Schema`t`t`t`t: ($($ADSchemaVersion)) - $($ADSchemaVersionName)"
				Line 1 "Allowed DNS Suffixes`t`t`t: " -NoNewLine
				$DNSSuffixes = $DomainInfo.AllowedDNSSuffixes | Sort-Object 
				If($Null -eq $DNSSuffixes)
				{
					 Line 0 "<None>"
				}
				Else
				{
					$cnt = 0
					ForEach($DNSSuffix in $DNSSuffixes)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							Line 0 $DNSSuffix.ToString()
						}
						Else
						{
							Line 6 "  $($DNSSuffix.ToString())"
						}
					}
				}
				Line 1 "Child domains`t`t`t`t: " -NoNewLine
				$ChildDomains = $DomainInfo.ChildDomains | Sort-Object 
				If($Null -eq $ChildDomains)
				{
					Line 0 "<None>"
				}
				Else
				{
					$cnt = 0
					ForEach($ChildDomain in $ChildDomains)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							Line 0 $ChildDomain.ToString()
						}
						Else
						{
							Line 6 "  $($ChildDomain.ToString())"
						}
					}
				}
				Line 1 "Default computers container`t`t: " $DomainInfo.ComputersContainer
				Line 1 "Default users container`t`t`t: " $DomainInfo.UsersContainer
				Line 1 "Deleted objects container`t`t: " $DomainInfo.DeletedObjectsContainer
				Line 1 "Distinguished name`t`t`t: " $DomainInfo.DistinguishedName
				Line 1 "DNS root`t`t`t`t: " $DomainInfo.DNSRoot
				Line 1 "Domain controllers container`t`t: " $DomainInfo.DomainControllersContainer
				If(![String]::IsNullOrEmpty($ExchangeSchemaInfo))
				{
					Line 1 "Exchange Schema`t`t`t`t: ($($ExchangeSchemaVersion)) - $($ExchangeSchemaVersionName)"
				}
				Line 1 "Foreign security principals container`t: " $DomainInfo.ForeignSecurityPrincipalsContainer
				Line 1 "Infrastructure master`t`t`t: " $DomainInfo.InfrastructureMaster
				#V2.20 added
				Line 1 "Last logon replication interval`t`t: " $LastLogonReplicationInterval
				Line 1 "Lost and Found container`t`t: " $DomainInfo.LostAndFoundContainer
				If(![String]::IsNullOrEmpty($DomainInfo.ManagedBy))
				{
					Line 1 "Managed by`t`t`t`t: " $DomainInfo.ManagedBy
				}
				Line 1 "PDC Emulator`t`t`t`t: " $DomainInfo.PDCEmulator
				#V2.20 added
				If(validObject $DomainInfo PublicKeyRequiredPasswordRolling)
				{
					Line 1 "Public key required password rolling`t: " $DomainInfo.PublicKeyRequiredPasswordRolling.ToString()
				}
				Line 1 "Quotas container`t`t`t: " $DomainInfo.QuotasContainer
				Line 1 "Read-only replica directory servers`t: " -NoNewLine
				$ReadOnlyReplicas = $DomainInfo.ReadOnlyReplicaDirectoryServers | Sort-Object 
				If($Null -eq $ReadOnlyReplicas)
				{
					Line 0 "<None>"
				}
				Else
				{
					$cnt = 0
					ForEach($ReadOnlyReplica in $ReadOnlyReplicas)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							Line 0 $ReadOnlyReplica.ToString()
						}
						Else
						{
							Line 6 "  $($ReadOnlyReplica.ToString())"
						}
					}
				}
				Line 1 "Replica directory servers`t`t: " -NoNewLine
				$Replicas = $DomainInfo.ReplicaDirectoryServers | Sort-Object 
				If($Null -eq $Replicas)
				{
					Line 0 "<None>"
				}
				Else
				{
					$cnt = 0
					ForEach($Replica in $Replicas)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							Line 0 $Replica.ToString()
						}
						Else
						{
							Line 6 "  $($Replica.ToString())"
						}
					}
				}
				Line 1 "RID Master`t`t`t`t: " $DomainInfo.RIDMaster
				Line 1 "Subordinate references`t`t`t: " -NoNewLine
				$SubordinateReferences = $DomainInfo.SubordinateReferences | Sort-Object 
				If($Null -eq $SubordinateReferences)
				{
					Line 0 "<None>"
				}
				Else
				{
					$cnt = 0
					ForEach($SubordinateReference in $SubordinateReferences)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							Line 0 $SubordinateReference.ToString()
						}
						Else
						{
							Line 6 "  $($SubordinateReference.ToString())"
						}
					}
				}
				Line 1 "Systems container`t`t`t: " $DomainInfo.SystemsContainer
				
				Write-Verbose "$(Get-Date): `t`tGetting domain trusts"
				Line 0 "Domain Trusts: "
				
				$ADDomainTrusts = $Null
				$ADDomainTrusts = Get-ADObject -Filter {ObjectClass -eq "trustedDomain"} -Server $Domain -Properties * -EA 0

				If($? -and $Null -ne $ADDomainTrusts)
				{
					
					ForEach($Trust in $ADDomainTrusts) 
					{ 
						Line 1 "Name`t`t: " $Trust.Name 
						
						If(![String]::IsNullOrEmpty($Trust.Description))
						{
							Line 1 "Description`t: " $Trust.Description
						}
						
						Line 1 "Created`t`t: " $Trust.Created
						Line 1 "Modified`t: " $Trust.Modified

						$TrustExtendedAttributes = Get-ADTrustInfo $Trust
						
						Line 1 "Type`t`t: " $TrustExtendedAttributes.TrustType
						Line 1 "Attributes`t: " -NoNewLine
						$cnt = 0
						ForEach($attribute in $TrustExtendedAttributes.Trustattribute)
						{
							$cnt++
							
							If($cnt -eq 1)
							{
								Line 0 $attribute.ToString()
							}
							Else
							{
								Line 3 "  $($attribute.ToString())"
							}
						}

						Line 1 "Direction`t: " $TrustDirection
						Line 0 ""
					}
				}
				ElseIf(!$?)
				{
					#error retrieving domain trusts
					Write-Warning "Error retrieving domain trusts for $($Domain)"
					Line 0 "Error retrieving domain trusts for $($Domain)"
				}
				Else
				{
					#no domain trust data
					Line 1 "<None>"
				}
				Line 0 ""
				
				Write-Verbose "$(Get-Date): `t`tProcessing domain controllers"
				$DomainControllers = $Null
				$DomainControllers = Get-ADDomainController -Filter * -Server $DomainInfo.DNSRoot -EA 0 | Sort-Object Name
				
				If($? -and $Null -ne $DomainControllers)
				{
					$Script:AllDomainControllers.Add($DomainControllers) > $Null
					Line 0 "Domain Controllers: "
					ForEach($DomainController in $DomainControllers)
					{
						Line 1 $DomainController.Name
					}
				}
				ElseIf(!$?)
				{
					Write-Warning "Error retrieving domain controller data for domain $($Domain)"
					Line 0 "Error retrieving domain controller data for domain $($Domain)"
				}
				Else
				{
					Line 0 "No Domain controller data was retrieved for domain $($Domain)"
				}
				Line 0 ""
				
				Write-Verbose "$(Get-Date): `t`tProcessing Fine Grained Password Policies"
				
				#are FGPP cmdlets available
				If(Get-Command -Name "Get-ADFineGrainedPasswordPolicy" -ea 0)
				{
					$FGPPs = $Null
					$FGPPs = Get-ADFineGrainedPasswordPolicy -Searchbase $DomainInfo.DistinguishedName -Filter * -Properties * -Server $DomainInfo.DNSRoot -EA 0 | Sort-Object Precedence, ObjectGUID
					
					If($? -and $Null -ne $FGPPs)
					{
						Line 0 "Fine Grained Password Policies"
						
						ForEach($FGPP in $FGPPs)
						{
							Line 1 "Name`t`t`t`t`t`t`t`t: " $FGPP.Name
							Line 1 "Precedence`t`t`t`t`t`t`t: " $FGPP.Precedence.ToString()
							
							Line 1 "Enforce minimum password length`t`t`t`t`t: " -NoNewLine
							If($FGPP.MinPasswordLength -eq 0)
							{
								Line 0 "Not enabled"
							}
							Else
							{
								Line 0 "Enabled"
								Line 3 "Minimum password length (characters)`t`t: " $FGPP.MinPasswordLength.ToString()
							}
							
							Line 1 "Enforce password history`t`t`t`t`t: " -NoNewLine
							If($FGPP.PasswordHistoryCount -eq 0)
							{
								Line 0 "Not enabled"
							}
							Else
							{
								Line 0 "Enabled"
								Line 3 "Number of passwords remembered`t`t`t: " $FGPP.PasswordHistoryCount.ToString()
							}
							
							Line 1 "Password must meet complexity requirements`t`t`t: " -NoNewLine
							If($FGPP.ComplexityEnabled -eq $True)
							{
								Line 0 "Enabled"
							}
							Else
							{
								Line 0 "Not enabled"
							}
							
							Line 1 "Store password using reversible encryption`t`t`t: " -NoNewLine
							If($FGPP.ReversibleEncryptionEnabled -eq $True)
							{
								Line 0 "Enabled"
							}
							Else
							{
								Line 0 "Not enabled"
							}
							
							Line 1 "Protect from accidental deletion`t`t`t`t: " -NoNewLine
							If($FGPP.ProtectedFromAccidentalDeletion -eq $True)
							{
								Line 0 "Enabled"
							}
							Else
							{
								Line 0 "Not enabled"
							}
							
							Line 1 "Password age options:"
							If($FGPP.MinPasswordAge.Days -eq 0)
							{
								Line 2 "Enforce minimum password age`t`t`t`t: Not enabled"
							}
							Else
							{
								Line 2 "Enforce minimum password age`t`t`t`t: Enabled" 
								Line 3 "User cannot change the password within (days)`t: " $FGPP.MinPasswordAge.TotalDays.ToString()
							}
							
							If($FGPP.MaxPasswordAge -eq 0)
							{
								Line 2 "Enforce maximum password age`t`t`t`t: Not enabled"
							}
							Else
							{
								Line 2 "Enforce maximum password age`t`t`t`t: Enabled"
								Line 3 "User must change the password after (days)`t: " $FGPP.MaxPasswordAge.TotalDays.ToString()
							}
							
							Line 1 "Enforce account lockout policy`t`t`t`t`t: " -NoNewLine
							If($FGPP.LockoutThreshold -eq 0)
							{
								Line 0 "Not enabled"
							}
							Else
							{
								Line 0 "Enabled"
								Line 2 "Number of failed logon attempts allowed`t`t`t: " $FGPP.LockoutThreshold.ToString()
								Line 2 "Reset failed logon attempts count after (mins)`t`t: " $FGPP.LockoutObservationWindow.TotalMinutes.ToString()
								If($FGPP.LockoutDuration -eq 0)
								{
									Line 2 "Account will be locked out"
									Line 3 "Until an administrator manually unlocks the account"
								}
								Else
								{
									Line 2 "Account will be locked out for a duration of (mins)`t: " $FGPP.LockoutDuration.TotalMinutes.ToString()
								}
								
							}
							
							Line 1 "Description`t`t`t`t`t`t`t: " $FGPP.Description
							
							$results = Get-ADFineGrainedPasswordPolicySubject -Identity $FGPP.Name -EA 0 | Sort-Object Name
							
							If($? -and $Null -ne $results)
							{
								Line 1 "Directly Applies To`t`t`t`t`t`t: " -NoNewLine
								$cnt = 0
								ForEach($Item in $results)
								{
									$cnt++
									
									If($cnt -eq 1)
									{
										Line 0 $Item.Name
									}
									Else
									{
										Line 9 "  $($Item.Name)"
									}
								}
							}
							Else
							{
							}
							
							Line 0 ""
						}
					}
					ElseIf(!$?)
					{
						Write-Warning "Error retrieving Fine Grained Password Policy data for domain $($Domain)"
						Line 0 "Error retrieving Fine Grained Password Policy data for domain $($Domain)"
					}
					Else
					{
						Line 0 "No Fine Grained Password Policy data was retrieved for domain $($Domain)"
					}
					Line 0 ""
				}
				Else
				{
					#FGPP cmdlets are not available
				}
			}
			ElseIf($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Domain mode",($htmlsilver -bor $htmlbold),$DomainMode,$htmlwhite)
				$rowdata += @(,('Domain name',($htmlsilver -bor $htmlbold),$DomainInfo.Name,$htmlwhite))
				$rowdata += @(,('NetBIOS name',($htmlsilver -bor $htmlbold),$DomainInfo.NetBIOSName,$htmlwhite))
				#V2.20 reorder the following properties in alpha order
				$rowdata += @(,('AD Schema',($htmlsilver -bor $htmlbold),"($($ADSchemaVersion)) - $($ADSchemaVersionName)",$htmlwhite))
				$DNSSuffixes = $DomainInfo.AllowedDNSSuffixes | Sort-Object 
				If($Null -eq $DNSSuffixes)
				{
					$rowdata += @(,('Allowed DNS Suffixes',($htmlsilver -bor $htmlbold),"None",$htmlwhite))
				}
				Else
				{
					$cnt = 0
					ForEach($DNSSuffix in $DNSSuffixes)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$rowdata += @(,('Allowed DNS Suffixes',($htmlsilver -bor $htmlbold),"$($DNSSuffix.ToString())",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),"$($DNSSuffix.ToString())",$htmlwhite))
						}
					}
				}
				$ChildDomains = $DomainInfo.ChildDomains | Sort-Object 
				If($Null -eq $ChildDomains)
				{
					$rowdata += @(,('Child domains',($htmlsilver -bor $htmlbold),"None",$htmlwhite))
				}
				Else
				{
					$cnt = 0 
					ForEach($ChildDomain in $ChildDomains)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$rowdata += @(,('Child domains',($htmlsilver -bor $htmlbold),"$($ChildDomain.ToString())",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),"$($ChildDomain.ToString())",$htmlwhite))
						}
					}
				}
				$rowdata += @(,('Default computers container',($htmlsilver -bor $htmlbold),$DomainInfo.ComputersContainer,$htmlwhite))
				$rowdata += @(,('Default users container',($htmlsilver -bor $htmlbold),$DomainInfo.UsersContainer,$htmlwhite))
				$rowdata += @(,('Deleted objects container',($htmlsilver -bor $htmlbold),$DomainInfo.DeletedObjectsContainer,$htmlwhite))
				$rowdata += @(,('Distinguished name',($htmlsilver -bor $htmlbold),$DomainInfo.DistinguishedName,$htmlwhite))
				$rowdata += @(,('DNS root',($htmlsilver -bor $htmlbold),$DomainInfo.DNSRoot,$htmlwhite))
				$rowdata += @(,('Domain controllers container',($htmlsilver -bor $htmlbold),$DomainInfo.DomainControllersContainer,$htmlwhite))
				If(![String]::IsNullOrEmpty($ExchangeSchemaInfo))
				{
					$rowdata += @(,('Exchange Schema',($htmlsilver -bor $htmlbold),"($($ExchangeSchemaVersion)) - $($ExchangeSchemaVersionName)",$htmlwhite))
				}
				$rowdata += @(,('Foreign security principals container',($htmlsilver -bor $htmlbold),$DomainInfo.ForeignSecurityPrincipalsContainer,$htmlwhite))
				$rowdata += @(,('Infrastructure master',($htmlsilver -bor $htmlbold),$DomainInfo.InfrastructureMaster,$htmlwhite))
				#V2.20 added
				$rowdata += @(,("Last logon replication interval",($htmlsilver -bor $htmlbold),$LastLogonReplicationInterval,$htmlwhite))
				$rowdata += @(,('Lost and Found container',($htmlsilver -bor $htmlbold),$DomainInfo.LostAndFoundContainer,$htmlwhite))
				If(![String]::IsNullOrEmpty($DomainInfo.ManagedBy))
				{
					$rowdata += @(,('Managed by',($htmlsilver -bor $htmlbold),$DomainInfo.ManagedBy,$htmlwhite))
				}
				$rowdata += @(,('PDC Emulator',($htmlsilver -bor $htmlbold),$DomainInfo.PDCEmulator,$htmlwhite))
				#V2.20 added
				If(validObject $DomainInfo PublicKeyRequiredPasswordRolling)
				{
					$rowdata += @(,("Public key required password rolling",($htmlsilver -bor $htmlbold),$DomainInfo.PublicKeyRequiredPasswordRolling.ToString(),$htmlwhite))
				}
				$rowdata += @(,('Quotas container',($htmlsilver -bor $htmlbold),$DomainInfo.QuotasContainer,$htmlwhite))
				$ReadOnlyReplicas = $DomainInfo.ReadOnlyReplicaDirectoryServers | Sort-Object 
				If($Null -eq $ReadOnlyReplicas)
				{
					$rowdata += @(,('Read-only replica directory servers',($htmlsilver -bor $htmlbold),"None",$htmlwhite))
				}
				Else
				{
					$cnt = 0 
					ForEach($ReadOnlyReplica in $ReadOnlyReplicas)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$rowdata += @(,('Read-only replica directory servers',($htmlsilver -bor $htmlbold),"$($ReadOnlyReplica.ToString())",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),"$($ReadOnlyReplica.ToString())",$htmlwhite))
						}
					}
				}
				$Replicas = $DomainInfo.ReplicaDirectoryServers | Sort-Object 
				If($Null -eq $Replicas)
				{
					$rowdata += @(,('Replica directory servers',($htmlsilver -bor $htmlbold),"None",$htmlwhite))
				}
				Else
				{
					$cnt = 0 
					ForEach($Replica in $Replicas)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$rowdata += @(,('Replica directory servers',($htmlsilver -bor $htmlbold),"$($Replica.ToString())",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),"$($Replica.ToString())",$htmlwhite))
						}
					}
				}
				$rowdata += @(,('RID Master',($htmlsilver -bor $htmlbold),$DomainInfo.RIDMaster,$htmlwhite))
				$SubordinateReferences = $DomainInfo.SubordinateReferences | Sort-Object 
				If($Null -eq $SubordinateReferences)
				{
					$rowdata += @(,('Subordinate references',($htmlsilver -bor $htmlbold),"None",$htmlwhite))
				}
				Else
				{
					$cnt = 0
					ForEach($SubordinateReference in $SubordinateReferences)
					{
						$cnt++
						
						If($cnt -eq 1)
						{
							$rowdata += @(,('Subordinate references',($htmlsilver -bor $htmlbold),"$($SubordinateReference.ToString())",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('',($htmlsilver -bor $htmlbold),"$($SubordinateReference.ToString())",$htmlwhite))
						}
					}
				}
				$rowdata += @(,('Systems container',($htmlsilver -bor $htmlbold),$DomainInfo.SystemsContainer,$htmlwhite))
				
				$msg = ""
				$columnWidths = @("175","300")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "475"
				WriteHTMLLine 0 0 " "

				Write-Verbose "$(Get-Date): `t`tGetting domain trusts"
				WriteHTMLLine 3 0 "Domain Trusts"
				
				$ADDomainTrusts = $Null
				$ADDomainTrusts = Get-ADObject -Filter {ObjectClass -eq "trustedDomain"} -Server $Domain -Properties * -EA 0

				If($? -and $Null -ne $ADDomainTrusts)
				{
					
					ForEach($Trust in $ADDomainTrusts) 
					{ 
						$rowdata = @()
						$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$Trust.Name,$htmlwhite)
						
						If(![String]::IsNullOrEmpty($Trust.Description))
						{
							$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Trust.Description,$htmlwhite))
						}
						
						$rowdata += @(,('Created',($htmlsilver -bor $htmlbold),$Trust.Created,$htmlwhite))
						$rowdata += @(,('Modified',($htmlsilver -bor $htmlbold),$Trust.Modified,$htmlwhite))
	
						$TrustExtendedAttributes = Get-ADTrustInfo $Trust
						 
						$rowdata += @(,('Type',($htmlsilver -bor $htmlbold),$TrustExtendedAttributes.TrustType,$htmlwhite))

						
						$cnt = 0
						ForEach($attribute in $TrustExtendedAttributes.Trustattribute)
						{
							$cnt++
							
							If($cnt -eq 1)
							{
								$rowdata += @(,('Attributes',($htmlsilver -bor $htmlbold),$attribute.ToString(),$htmlwhite))
							}
							Else
							{
								$rowdata += @(,('',($htmlsilver -bor $htmlbold),$attribute.ToString(),$htmlwhite))
							}
						}

						$rowdata += @(,('Direction',($htmlsilver -bor $htmlbold),$TrustDirection,$htmlwhite))

						$msg = ""
						$columnWidths = @("175","300")
						FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "475"
						WriteHTMLLine 0 0 " "
					}
				}
				ElseIf(!$?)
				{
					#error retrieving domain trusts
					Write-Warning "Error retrieving domain trusts for $($Domain)"
					WriteHTMLLine 0 0 "Error retrieving domain trusts for $($Domain)" "" $Null 0 $False $True
				}
				Else
				{
					#no domain trust data
					WriteHTMLLine 0 0 "None"
				}
				
				Write-Verbose "$(Get-Date): `t`tProcessing domain controllers"
				$DomainControllers = $Null
				$DomainControllers = Get-ADDomainController -Filter * -Server $DomainInfo.DNSRoot -EA 0 | Sort-Object Name
				
				If($? -and $Null -ne $DomainControllers)
				{
					$Script:AllDomainControllers.Add($DomainControllers) > $Null
					$rowdata = @()
					WriteHTMLLine 3 0 "Domain Controllers"
					ForEach($DomainController in $DomainControllers)
					{
						$rowdata += @(,($DomainController.Name,$htmlwhite))
					}
					$msg = ""
					$columnHeaders = @("Name",($htmlsilver -bor $htmlbold))
					$columnWidths = @("105")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "105"
					WriteHTMLLine 0 0 " "
				}
				ElseIf(!$?)
				{
					Write-Warning "Error retrieving domain controller data for domain $($Domain)"
					WriteHTMLLine 0 0 "Error retrieving domain controller data for domain $($Domain)" "" $Null 0 $False $True
				}
				Else
				{
					WriteHTMLLine 0 0 "No Domain controller data was retrieved for domain $($Domain)" "" $Null 0 $False $True
				}
				WriteHTMLLine 0 0 " "
				
				Write-Verbose "$(Get-Date): `t`tProcessing Fine Grained Password Policies"
				
				#are FGPP cmdlets available
				If(Get-Command -Name "Get-ADFineGrainedPasswordPolicy" -ea 0)
				{
					$FGPPs = $Null
					$FGPPs = Get-ADFineGrainedPasswordPolicy -Searchbase $DomainInfo.DistinguishedName -Filter * -Properties * -Server $DomainInfo.DNSRoot -EA 0 | Sort-Object Precedence, ObjectGUID
					
					If($? -and $Null -ne $FGPPs)
					{
						WriteHTMLLine 3 0 "Fine Grained Password Policies"
						
						ForEach($FGPP in $FGPPs)
						{
							$rowdata = @()
							$columnHeaders = @("Precedence",($htmlsilver -bor $htmlbold),$FGPP.Precedence.ToString(),$htmlwhite)
							
							If($FGPP.MinPasswordLength -eq 0)
							{
								$rowdata += @(,("Enforce minimum password length",($htmlsilver -bor $htmlbold),"Not enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("Enforce minimum password length",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
								$rowdata += @(,("     Minimum password length (characters)",($htmlsilver -bor $htmlbold),$FGPP.MinPasswordLength.ToString(),$htmlwhite))
							}
							
							If($FGPP.PasswordHistoryCount -eq 0)
							{
								$rowdata += @(,("Enforce password history",($htmlsilver -bor $htmlbold),"Not enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("Enforce password history",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
								$rowdata += @(,("     Number of passwords remembered",($htmlsilver -bor $htmlbold),$FGPP.PasswordHistoryCount.ToString(),$htmlwhite))
							}
							
							If($FGPP.ComplexityEnabled -eq $True)
							{
								$rowdata += @(,("Password must meet complexity requirements",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("Password must meet complexity requirements",($htmlsilver -bor $htmlbold),"Not enabled",$htmlwhite))
							}
							
							If($FGPP.ReversibleEncryptionEnabled -eq $True)
							{
								$rowdata += @(,("Store password using reversible encryption",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("Store password using reversible encryption",($htmlsilver -bor $htmlbold),"Not enabled",$htmlwhite))
							}
							
							If($FGPP.ProtectedFromAccidentalDeletion -eq $True)
							{
								$rowdata += @(,("Protect from accidental deletion",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("Protect from accidental deletion",($htmlsilver -bor $htmlbold),"Not enabled",$htmlwhite))
							}
							
							$rowdata += @(,("Password age options",($htmlsilver -bor $htmlbold),"",$htmlwhite))
							If($FGPP.MinPasswordAge.Days -eq 0)
							{
								$rowdata += @(,("     Enforce minimum password age",($htmlsilver -bor $htmlbold),"Not enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("     Enforce minimum password age",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
								$rowdata += @(,("          User cannot change the password within (days)",($htmlsilver -bor $htmlbold),$FGPP.MinPasswordAge.TotalDays.ToString(),$htmlwhite))
							}
							
							If($FGPP.MaxPasswordAge -eq 0)
							{
								$rowdata += @(,("     Enforce maximum password age",($htmlsilver -bor $htmlbold),"Not enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("     Enforce maximum password age",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
								$rowdata += @(,("          User must change the password after (days)",($htmlsilver -bor $htmlbold),$FGPP.MaxPasswordAge.TotalDays.ToString(),$htmlwhite))
							}
							
							If($FGPP.LockoutThreshold -eq 0)
							{
								$rowdata += @(,("Enforce account lockout policy",($htmlsilver -bor $htmlbold),"Not enabled",$htmlwhite))
							}
							Else
							{
								$rowdata += @(,("Enforce account lockout policy",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
								$rowdata += @(,("     Number of failed logon attempts allowed",($htmlsilver -bor $htmlbold),$FGPP.LockoutThreshold.ToString(),$htmlwhite))
								$rowdata += @(,("     Reset failed logon attempts count after (mins)",($htmlsilver -bor $htmlbold),$FGPP.LockoutObservationWindow.TotalMinutes.ToString(),$htmlwhite))
								If($FGPP.LockoutDuration -eq 0)
								{
									$rowdata += @(,("     Account will be locked out",($htmlsilver -bor $htmlbold),"",$htmlwhite))
									$rowdata += @(,("          Until an administrator manually unlocks the account",($htmlsilver -bor $htmlbold),"",$htmlwhite))
								}
								Else
								{
									$rowdata += @(,("     Account will be locked out for a duration of (mins)",($htmlsilver -bor $htmlbold),$FGPP.LockoutDuration.TotalMinutes.ToString(),$htmlwhite))
								}
								
							}
							
							$rowdata += @(,("Description",($htmlsilver -bor $htmlbold),$FGPP.Description,$htmlwhite))
							
							$results = Get-ADFineGrainedPasswordPolicySubject -Identity $FGPP.Name -EA 0 | Sort-Object Name
							
							If($? -and $Null -ne $results)
							{
								$cnt = 0
								ForEach($Item in $results)
								{
									$cnt++
									
									If($cnt -eq 1)
									{
										$rowdata += @(,("Directly Applies To",($htmlsilver -bor $htmlbold),$Item.Name,$htmlwhite))
									}
									Else
									{
										$rowdata += @(,("",($htmlsilver -bor $htmlbold),$($Item.Name),$htmlwhite))
									}
								}
							}
							Else
							{
							}
							
							$msg = "Name: $($FGPP.Name)"
							$columnWidths = @("300","225")
							FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "525"
							WriteHTMLLine 0 0 " "
						}
					}
					ElseIf(!$?)
					{
						Write-Warning "Error retrieving Fine Grained Password Policy data for domain $($Domain)"
						WriteHTMLLine 0 0 "Error retrieving Fine Grained Password Policy data for domain $($Domain)"
					}
					Else
					{
						WriteHTMLLine 0 0 "No Fine Grained Password Policy data was retrieved for domain $($Domain)"
					}
					WriteHTMLLine 0 0 " "
				}
				Else
				{
					#FGPP cmdlets are not available
				}
			}

			$First = $False
		}
		ElseIf(!$?)
		{
			$txt = "Error retrieving domain data for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 $txt
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}
		}
		Else
		{
			$txt = "No Domain data was retrieved for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 $txt
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}
		}
	}
	$ADDomainTrusts        = $Null
	$ADSchemaInfo          = $Null
	$ChildDomains          = $Null
	$DNSSuffixes           = $Null
	$DomainControllers     = $Null
	$ExchangeSchemaInfo    = $Null
	$FGPPs                 = $Null
	$First                 = $Null
	$ReadOnlyReplicas      = $Null
	$Replicas              = $Null
	$SubordinateReferences = $Null
	$Table                 = $Null
}
#endregion

#region domain controllers
Function ProcessDomainControllers
{
	Write-Verbose "$(Get-Date): Writing domain controller data"

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Domain Controllers in $($Script:ForestName)"
	}
	ElseIf($Text)
	{
		Line 0 "///  Domain Controllers in $($Script:ForestName)  \\\"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Domain Controllers in $($Script:ForestName)&nbsp;&nbsp;\\\"
	}

	$Script:DCDNSIPInfo = New-Object System.Collections.ArrayList
	#V2.19 added
	$Script:DCEventLogInfo = New-Object System.Collections.ArrayList
	$Script:TimeServerInfo = New-Object System.Collections.ArrayList
	$Script:AllDomainControllers = $Script:AllDomainControllers | Sort-Object Name
	$First = $True

	ForEach($DC in $Script:AllDomainControllers)
	{
		Write-Verbose "$(Get-Date): `tProcessing domain controller $($DC.name)"
		$FSMORoles = $DC.OperationMasterRoles | Sort-Object 
		$Partitions = $DC.Partitions | Sort-Object 
		
		If(($MSWORD -or $PDF) -and !$First)
		{
			#put each DC, starting with the second, on a new page
			$Script:selection.InsertNewPage()
		}
		
		If($MSWORD -or $PDF)
		{
			WriteWordLine 2 0 $DC.Name
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Default partition"; Value = $DC.DefaultPartition; }
			$ScriptInformation += @{ Data = "Domain"; Value = $DC.domain; }
			If($DC.Enabled -eq $True)
			{
				$tmp = "True"
			}
			Else
			{
				$tmp = "False"
			}
			$ScriptInformation += @{ Data = "Enabled"; Value = $tmp; }
			$ScriptInformation += @{ Data = "Hostname"; Value = $DC.HostName; }
			If($DC.IsGlobalCatalog -eq $True)
			{
				$tmp = "Yes" 
			}
			Else
			{
				$tmp = "No"
			}
			$ScriptInformation += @{ Data = "Global Catalog"; Value = $tmp; }
			If($DC.IsReadOnly -eq $True)
			{
				$tmp = "Yes"
			}
			Else
			{
				$tmp = "No"
			}
			$ScriptInformation += @{ Data = "Read-only"; Value = $tmp; }
			$ScriptInformation += @{ Data = "LDAP port"; Value = $DC.LdapPort.ToString(); }
			$ScriptInformation += @{ Data = "SSL port"; Value = $DC.SslPort.ToString(); }
			If($Null -eq $FSMORoles)
			{
				$ScriptInformation += @{ Data = "Operation Master roles"; Value = "<None>"; }
			}
			Else
			{
				$cnt = 0
				ForEach($FSMORole in $FSMORoles)
				{
					$cnt++
					
					If($cnt -eq 1)
					{
						$ScriptInformation += @{ Data = "Operation Master roles"; Value = $FSMORole.ToString(); }
					}
					Else
					{
						$ScriptInformation += @{ Data = ""; Value = $FSMORole.ToString(); }
					}
				}
			}
			If($Null -eq $Partitions)
			{
				$ScriptInformation += @{ Data = "Partitions"; Value = "<None>"; }
			}
			Else
			{
				$cnt = 0
				ForEach($Partition in $Partitions)
				{
					$cnt++
					
					If($cnt -eq 1)
					{
						$ScriptInformation += @{ Data = "Partitions"; Value = $Partition.ToString(); }
					}
					Else
					{
						$ScriptInformation += @{ Data = ""; Value = $Partition.ToString(); }
					}
					
				}
			}
			$ScriptInformation += @{ Data = "Site"; Value = $DC.Site; }
			$ScriptInformation += @{ Data = "Operating System"; Value = $DC.OperatingSystem; }
			
			If(![String]::IsNullOrEmpty($DC.OperatingSystemServicePack))
			{
				$ScriptInformation += @{ Data = "Service Pack"; Value = $DC.OperatingSystemServicePack; }
			}
			$ScriptInformation += @{ Data = "Operating System version"; Value = $DC.OperatingSystemVersion; }
			
			If(!$Hardware)
			{
				If([String]::IsNullOrEmpty($DC.IPv4Address))
				{
					$tmp = "<None>"
				}
				Else
				{
					$tmp = $DC.IPv4Address
				}
				$ScriptInformation += @{ Data = "IPv4 Address"; Value = $tmp; }

				If([String]::IsNullOrEmpty($DC.IPv6Address))
				{
					$tmp = "<None>"
				}
				Else
				{
					$tmp = $DC.IPv6Address
				}
				$ScriptInformation += @{ Data = "IPv6 Address"; Value = $tmp; }
			}
			
			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 140;
			$Table.Columns.Item(2).Width = 300;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 0 "///  DC: $($DC.Name)  \\\"
			Line 1 "Default partition`t`t: " $DC.DefaultPartition
			Line 1 "Domain`t`t`t`t: " $DC.domain
			If($DC.Enabled -eq $True)
			{
				$tmp = "True"
			}
			Else
			{
				$tmp = "False"
			}
			Line 1 "Enabled`t`t`t`t: " $tmp
			Line 1 "Hostname`t`t`t: " $DC.HostName
			If($DC.IsGlobalCatalog -eq $True)
			{
				$tmp = "Yes" 
			}
			Else
			{
				$tmp = "No"
			}
			Line 1 "Global Catalog`t`t`t: " $tmp
			If($DC.IsReadOnly -eq $True)
			{
				$tmp = "Yes"
			}
			Else
			{
				$tmp = "No"
			}
			Line 1 "Read-only`t`t`t: " $tmp
			Line 1 "LDAP port`t`t`t: " $DC.LdapPort.ToString()
			Line 1 "SSL port`t`t`t: " $DC.SslPort.ToString()
			Line 1 "Operation Master roles`t`t: " -NoNewLine
			If($Null -eq $FSMORoles)
			{
				Line 0 "<None>"
			}
			Else
			{
				$cnt = 0
				ForEach($FSMORole in $FSMORoles)
				{
					$cnt++
					
					If($cnt -eq 1)
					{
						Line 0 $FSMORole.ToString()
					}
					Else
					{
						Line 5 "  $($FSMORole.ToString())"
					}
					
				}
			}
			Line 1 "Partitions`t`t`t: " -NoNewLine
			If($Null -eq $Partitions)
			{
				Line 0 "<None>"
			}
			Else
			{
				$cnt = 0
				ForEach($Partition in $Partitions)
				{
					$cnt++
					
					If($cnt -eq 1)
					{
						Line 0 $Partition.ToString()
					}
					Else
					{
						Line 5 "  $($Partition.ToString())"
					}
				}
			}
			Line 1 "Site`t`t`t`t: " $DC.Site
			Line 1 "Operating System`t`t: " $DC.OperatingSystem
			If(![String]::IsNullOrEmpty($DC.OperatingSystemServicePack))
			{
				Line 1 "Service Pack`t`t`t: " $DC.OperatingSystemServicePack
			}
			Line 1 "Operating System version`t: " $DC.OperatingSystemVersion
			
			If(!$Hardware)
			{
				Line 1 "IPv4 Address`t`t`t: " -NoNewLine
				If([String]::IsNullOrEmpty($DC.IPv4Address))
				{
					Line 0 "<None>"
				}
				Else
				{
					Line 0 $DC.IPv4Address
				}
				Line 1 "IPv6 Address`t`t`t: " -NoNewLine
				If([String]::IsNullOrEmpty($DC.IPv6Address))
				{
					Line 0 "<None>"
				}
				Else
				{
					Line 0 $DC.IPv6Address
				}
			}
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($DC.Name)&nbsp;&nbsp;\\\"
			$rowdata = @()
			$columnHeaders = @("Default partition",($htmlsilver -bor $htmlbold),$DC.DefaultPartition,$htmlwhite)
			$rowdata += @(,('Domain',($htmlsilver -bor $htmlbold),$DC.domain,$htmlwhite))
			If($DC.Enabled -eq $True)
			{
				$tmp = "True"
			}
			Else
			{
				$tmp = "False"
			}
			$rowdata += @(,('Enabled',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			$rowdata += @(,('Hostname',($htmlsilver -bor $htmlbold),$DC.HostName,$htmlwhite))
			If($DC.IsGlobalCatalog -eq $True)
			{
				$tmp = "Yes" 
			}
			Else
			{
				$tmp = "No"
			}
			$rowdata += @(,('Global Catalog',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			If($DC.IsReadOnly -eq $True)
			{
				$tmp = "Yes"
			}
			Else
			{
				$tmp = "No"
			}
			$rowdata += @(,('Read-only',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			$rowdata += @(,('LDAP port',($htmlsilver -bor $htmlbold),$DC.LdapPort.ToString(),$htmlwhite))
			$rowdata += @(,('SSL port',($htmlsilver -bor $htmlbold),$DC.SslPort.ToString(),$htmlwhite))
			If($Null -eq $FSMORoles)
			{
				$rowdata += @(,('Operation Master roles',($htmlsilver -bor $htmlbold),"None",$htmlwhite))
			}
			Else
			{
				$cnt = 0
				ForEach($FSMORole in $FSMORoles)
				{
					$cnt++
					
					If($cnt -eq 1)
					{
						$rowdata += @(,('Operation Master roles',($htmlsilver -bor $htmlbold),$FSMORole.ToString(),$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('',($htmlsilver -bor $htmlbold),$FSMORole.ToString(),$htmlwhite))
					}
				}
			}
			If($Null -eq $Partitions)
			{
				$rowdata += @(,('Partitions',($htmlsilver -bor $htmlbold),"None",$htmlwhite))
			}
			Else
			{
				$cnt = 0
				ForEach($Partition in $Partitions)
				{
					$cnt++
					
					If($cnt -eq 1)
					{
						$rowdata += @(,('Partitions',($htmlsilver -bor $htmlbold),$Partition.ToString(),$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('',($htmlsilver -bor $htmlbold),$Partition.ToString(),$htmlwhite))
					}
				}
			}
			$rowdata += @(,('Site',($htmlsilver -bor $htmlbold),$DC.Site,$htmlwhite))
			$rowdata += @(,('Operating System',($htmlsilver -bor $htmlbold),$DC.OperatingSystem,$htmlwhite))
			
			If(![String]::IsNullOrEmpty($DC.OperatingSystemServicePack))
			{
				$rowdata += @(,('Service Pack',($htmlsilver -bor $htmlbold),$DC.OperatingSystemServicePack,$htmlwhite))
			}
			$rowdata += @(,('Operating System version',($htmlsilver -bor $htmlbold),$DC.OperatingSystemVersion,$htmlwhite))
			
			If(!$Hardware)
			{
				If([String]::IsNullOrEmpty($DC.IPv4Address))
				{
					$tmp = "None"
				}
				Else
				{
					$tmp = $DC.IPv4Address
				}
				$rowdata += @(,('IPv4 Address',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))

				If([String]::IsNullOrEmpty($DC.IPv6Address))
				{
					$tmp = "None"
				}
				Else
				{
					$tmp = $DC.IPv6Address
				}
				$rowdata += @(,('IPv6 Address',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
			
			$msg = ""
			$columnWidths = @("140","300")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "440"
			WriteHTMLLine 0 0 " "
		}
		
		If($Script:DARights -and $Script:Elevated)
		{
			OutputTimeServerRegistryKeys $DC.HostName
		
			OutputADFileLocations $DC.HostName
			
			OutputEventLogInfo $DC.HostName
		}
		
		If($Hardware -or $Services -or $DCDNSInfo)
		{
			If(Test-Connection -ComputerName $DC.HostName -quiet -EA 0)
			{
				If($Hardware)
				{
					GetComputerWMIInfo $DC.HostName
				}
				
				If($DCDNSInfo)
				{
					BuildDCDNSIPConfigTable $DC.HostName $DC.Site
				}

				If($Services)
				{
					GetComputerServices $DC.HostName
				}
				
			}
			Else
			{
				$txt = "$(Get-Date): `t`t$($DC.Name) is offline or unreachable.  Hardware inventory is skipped."
				Write-Verbose $txt
				If($MSWORD -or $PDF)
				{
					WriteWordLine 0 0 $txt
				}
				ElseIf($Text)
				{
					Line 0 $txt
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 0 0 $txt
				}

				If($Hardware -and -not $Services)
				{
					$txt = "Hardware inventory was skipped."
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 0 $txt
					}
					ElseIf($Text)
					{
						Line 0 $txt
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 0 $txt
					}
				}
				ElseIf($Services -and -not $Hardware)
				{
					$txt = "Services was skipped."
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 0 $txt
					}
					ElseIf($Text)
					{
						Line 0 $txt
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 0 $txt
					}
				}
				ElseIf($Hardware -and $Services)
				{
					$txt = "Hardware inventory and Services were skipped."
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 0 $txt
					}
					ElseIf($Text)
					{
						Line 0 $txt
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 0 $txt
					}
				}
			}
		}
		$First = $False
	}
	$Script:AllDomainControllers = $Null
}

Function OutputTimeServerRegistryKeys 
{
	Param( [string] $DCName )
	
	Write-Verbose "$(Get-Date): `tTimeServer Registry Keys"
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config	AnnounceFlags
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config	MaxNegPhaseCorrection
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Config	MaxPosPhaseCorrection
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Parameters	NtpServer
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\Parameters	Type 	
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient	SpecialPollInterval
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\VMICTimeProvider Enabled
	
	$AnnounceFlags           = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config" "AnnounceFlags" $DCName
	$MaxNegPhaseCorrection   = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config" "MaxNegPhaseCorrection" $DCName
	$MaxPosPhaseCorrection   = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config" "MaxPosPhaseCorrection" $DCName
	$NtpServer               = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" "NtpServer" $DCName
	$NtpType                 = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" "Type" $DCName
	$SpecialPollInterval     = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient" "SpecialPollInterval" $DCName
	$VMICTimeProviderEnabled = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\VMICTimeProvider" "Enabled" $DCName
	$NTPSource               = w32tm /query /computer:$DCName /source
	
	If($VMICTimeProviderEnabled -eq 0)
	{
		$VMICEnabled = "Disabled"
	}
	Else
	{
		$VMICEnabled = "Enabled"
	}
	
	#create time server info array
	$obj = New-Object -TypeName PSObject
	$obj | Add-Member -MemberType NoteProperty -Name DCName                -Value $DCName
	$obj | Add-Member -MemberType NoteProperty -Name TimeSource            -Value $NTPSource
	$obj | Add-Member -MemberType NoteProperty -Name AnnounceFlags         -Value $AnnounceFlags
	$obj | Add-Member -MemberType NoteProperty -Name MaxNegPhaseCorrection -Value $MaxNegPhaseCorrection
	$obj | Add-Member -MemberType NoteProperty -Name MaxPosPhaseCorrection -Value $MaxPosPhaseCorrection
	$obj | Add-Member -MemberType NoteProperty -Name NtpServer             -Value $NtpServer
	$obj | Add-Member -MemberType NoteProperty -Name NtpType               -Value $NtpType
	$obj | Add-Member -MemberType NoteProperty -Name SpecialPollInterval   -Value $SpecialPollInterval
	$obj | Add-Member -MemberType NoteProperty -Name VMICTimeProvider      -Value $VMICEnabled
	
	[void]$Script:TimeServerInfo.Add($obj)
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 "Time Server Information"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		
		$Scriptinformation += @{ Data = "Time source"; Value = $NTPSource; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CCS\Services\W32Time\Config\AnnounceFlags"; Value = $AnnounceFlags; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CCS\Services\W32Time\Config\MaxNegPhaseCorrection"; Value = $MaxNegPhaseCorrection; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CCS\Services\W32Time\Config\MaxPosPhaseCorrection"; Value = $MaxPosPhaseCorrection; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CCS\Services\W32Time\Parameters\NtpServer"; Value = $NtpServer; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CCS\Services\W32Time\Parameters\Type"; Value = $NtpType; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CCS\Services\W32Time\TimeProviders\NtpClient\SpecialPollInterval"; Value = $SpecialPollInterval; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CCS\Services\W32Time\TimeProviders\VMICTimeProvider\Enabled"; Value = $VMICEnabled; }
		
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
		SetWordCellFormat -Collection $Table -Size 9;

		$Table.Columns.Item(1).Width = 335;
		$Table.Columns.Item(2).Width = 130;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "Time Server Information"
		Line 0 ""
		Line 1 "Time source: " $NTPSource
		Line 1 "HKLM:\SYSTEM\CCS\Services\W32Time\"
		Line 2 "Config\AnnounceFlags`t`t`t`t: " $AnnounceFlags
		Line 2 "Config\MaxNegPhaseCorrection`t`t`t: " $MaxNegPhaseCorrection
		Line 2 "Config\MaxPosPhaseCorrection`t`t`t: " $MaxPosPhaseCorrection
		Line 2 "Parameters\NtpServer`t`t`t`t: " $NtpServer
		Line 2 "Parameters\Type`t`t`t`t`t: " $NtpType
		Line 2 "TimeProviders\NtpClient\SpecialPollInterval`t: " $SpecialPollInterval
		Line 2 "TimeProviders\VMICTimeProvider\Enabled`t`t: " $VMICEnabled
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "Time Server Information"
		$rowdata = @()
		$columnHeaders = @("Time source",($htmlsilver -bor $htmlbold),$NTPSource,$htmlwhite)
		$rowdata += @(,('HKLM:\SYSTEM\CCS\Services\W32Time\Config\AnnounceFlags',($htmlsilver -bor $htmlbold),$AnnounceFlags,$htmlwhite))
		$rowdata += @(,('HKLM:\SYSTEM\CCS\Services\W32Time\Config\MaxNegPhaseCorrection',($htmlsilver -bor $htmlbold),$MaxNegPhaseCorrection,$htmlwhite))
		$rowdata += @(,('HKLM:\SYSTEM\CCS\Services\W32Time\Config\MaxPosPhaseCorrection',($htmlsilver -bor $htmlbold),$MaxPosPhaseCorrection,$htmlwhite))
		$rowdata += @(,('HKLM:\SYSTEM\CCS\Services\W32Time\Parameters\NtpServer',($htmlsilver -bor $htmlbold),$NtpServer,$htmlwhite))
		$rowdata += @(,('HKLM:\SYSTEM\CCS\Services\W32Time\Parameters\Type',($htmlsilver -bor $htmlbold),$NtpType,$htmlwhite))
		$rowdata += @(,('HKLM:\SYSTEM\CCS\Services\W32Time\TimeProviders\NtpClient\SpecialPollInterval',($htmlsilver -bor $htmlbold),$SpecialPollInterval,$htmlwhite))
		$rowdata += @(,('HKLM:\SYSTEM\CCS\Services\W32Time\TimeProviders\VMICTimeProvider\Enabled',($htmlsilver -bor $htmlbold),$VMICEnabled,$htmlwhite))

		$msg = ""
		$columnWidths = @("335","130")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "465"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputADFileLocations
{
	Param( [string] $DCName )
	
	Write-Verbose "$(Get-Date): `tAD Database, Logfile and SYSVOL locations"
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NTDS\Parameters	'DSA Database file'
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NTDS\Parameters	'Database log files path'
	#HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters	SysVol
	
	$DSADatabaseFile = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters" "DSA Database file" $DCName
	$DatabaseLogFilesPath = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters" "Database log files path" $DCName
	$SysVol = Get-RegistryValue "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters" "SysVol" $DCName
	
	#calculation is taken from http://blogs.metcorpconsulting.com/tech/?p=177
	$DITRemotePath = $DSADatabaseFile.Replace(":", "$")
	$DITFile = "\\$DCName\$DITRemotePath"
	$DITsize = ([System.IO.FileInfo]$DITFile).Length
	$DITsize = ($DITsize/1GB)
	$DSADatabaseFileSize = "{0:N3}" -f $DITsize
		
	If($MSWORD -or $PDF)
	{
		WriteWordLine 3 0 "AD Database, Logfile and SYSVOL Locations"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters\DSA Database file"; Value = $DSADatabaseFile; }
		$Scriptinformation += @{ Data = "DSA Database file size "; Value = "$($DSADatabaseFileSize) GB"; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters\Database log files path"; Value = $DatabaseLogFilesPath; }
		$Scriptinformation += @{ Data = "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters\SysVol"; Value = $SysVol; }
		
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
		SetWordCellFormat -Collection $Table -Size 9;

		$Table.Columns.Item(1).Width = 335;
		$Table.Columns.Item(2).Width = 130;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "AD Database, Logfile and SYSVOL Locations"
		Line 0 ""
		Line 1 "HKLM:\SYSTEM\CCS\Services\"
		Line 2 "NTDS\Parameters\DSA Database file`t: " $DSADatabaseFile
		Line 2 "DSA Database file size`t`t`t: $($DSADatabaseFileSize) GB"
		Line 2 "NTDS\Parameters\Database log files path`t: " $DatabaseLogFilesPath
		Line 2 "Netlogon\Parameters\SysVol`t`t: " $SysVol
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "AD Database, Logfile and SYSVOL Locations"
		$rowdata = @()
		$columnHeaders = @("HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters\DSA Database file",($htmlsilver -bor $htmlbold),$DSADatabaseFile,$htmlwhite)
		$rowdata += @(,('DSA Database file size',($htmlsilver -bor $htmlbold),"$($DSADatabaseFileSize) GB",$htmlwhite))
		$rowdata += @(,('HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters\Database log files path',($htmlsilver -bor $htmlbold),$DatabaseLogFilesPath,$htmlwhite))
		$rowdata += @(,('HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters\SysVol',($htmlsilver -bor $htmlbold),$SysVol,$htmlwhite))

		$msg = ""
		$columnWidths = @("335","130")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "465"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputEventLogInfo
{
	Param( [string] $DCName )
	#V2.19 added
	
	Write-Verbose "$(Get-Date): `tEvent Log Information"
	$ELInfo = New-Object System.Collections.ArrayList
	
	$EventLogs = Get-EventLog -List -ComputerName $DCName | Select-Object MaximumKilobytes, Log | Sort-Object Log 
	
	If($? -and $Null -ne $EventLogs)
	{
		ForEach($EventLog in $EventLogs)
		{
			[string]$ELSize = "{0,10:N0}" -f $EventLog.MaximumKilobytes
			
			$obj = New-Object -TypeName PSObject
			$obj | Add-Member -MemberType NoteProperty -Name DCName                -Value $DCName
			$obj | Add-Member -MemberType NoteProperty -Name EventLogName          -Value $EventLog.Log
			$obj | Add-Member -MemberType NoteProperty -Name EventLogSize          -Value $ELSize
			
			[void]$Script:DCEventLogInfo.Add($obj)
			[void]$ELInfo.Add($obj)
		}
	}
	Else
	{
		[string]$ELSize = "{0,10:N0}" -f 0
	
		$obj = New-Object -TypeName PSObject
		$obj | Add-Member -MemberType NoteProperty -Name DCName                -Value $DCName
		$obj | Add-Member -MemberType NoteProperty -Name EventLogName          -Value "Unable to retrieve Event Log data"
		$obj | Add-Member -MemberType NoteProperty -Name EventLogSize          -Value $ELSize
		
		[void]$Script:DCEventLogInfo.Add($obj)
		[void]$ELInfo.Add($obj)
	}
	
	#V2.20 changed to @()
	$xEventLogInfo = @($ELInfo | Sort-Object EventLogName)
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Event Log Information"
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		[int]$Rows = $xEventLogInfo.Count + 1
		[int]$xRow = 1
		
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.AutoFitBehavior($wdAutoFitFixed)
		$Table.Style = $Script:MyHash.Word_TableGrid
	
		$Table.rows.first.headingformat = $wdHeadingFormatTrue
		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

		$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Event Log Name"
		
		$Table.Cell($xRow,2).Range.Font.Bold = $True
		$Table.Cell($xRow,2).Range.Text = "Event Log Size (KB)"
	}
	ElseIf($Text)
	{
		Line 0 "Event Log Information"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "Event Log Information"
		$rowdata = @()
	}

	ForEach($Item in $xEventLogInfo)
	{
		If($MSWord -or $PDF)
		{
			$xRow++
			$Table.Cell($xRow,1).Range.Text = $Item.EventLogName
			$Table.Cell($xRow,2).Range.ParagraphFormat.Alignment = $wdCellAlignVerticalTop
			$Table.Cell($xRow,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
			$Table.Cell($xRow,2).Range.Text = $Item.EventLogSize
		}
		ElseIf($Text)
		{
			Line 1 "Event Log Name`t`t: " $Item.EventLogName
			Line 1 "Event Log Size (KB)`t: " $Item.EventLogSize
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
				$Item.EventLogName,$htmlwhite,
				$Item.EventLogSize,$htmlwhite
			))
		}
	}

	If($MSWord -or $PDF)
	{
		#set column widths
		$xcols = $table.columns

		ForEach($xcol in $xcols)
		{
			switch ($xcol.Index)
			{
			  1 {$xcol.width = 150; Break}
			  2 {$xcol.width = 100; Break}
			}
		}
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
		$Table.AutoFitBehavior($wdAutoFitFixed)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		#nothing to do
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Event Log Name',($htmlsilver -bor $htmlbold),
		'Event Log Size (KB)',($htmlsilver -bor $htmlbold)
		)

		$msg = ""
		$columnWidths = @("175px","125px")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "300"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region organizational units
Function ProcessOrganizationalUnits
{
	Write-Verbose "$(Get-Date): Writing OU data by Domain"
	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Organizational Units"
	}
	ElseIf($Text)
	{
		Line 0 "///  Organizational Units  \\\"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Organizational Units&nbsp;&nbsp;\\\"
	}
	$First = $True

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date): `tProcessing domain $($Domain)"
		
		If(($MSWORD -or $PDF) -and !$First)
		{
			#put each domain, starting with the second, on a new page
			$Script:selection.InsertNewPage()
		}
		
		If($Domain -eq $Script:ForestRootDomain)
		{
			$txt = "OUs in Domain $($Domain) (Forest Root)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
			}
			ElseIf($Text)
			{
				Line 0 "///  $($txt)  \\\"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
			}
		}
		Else
		{
			$txt = "OUs in Domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
			}
			ElseIf($Text)
			{
				Line 0 "///  $($txt)  \\\"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
			}
		}
		
		#get all OUs for the domain
		$OUs = $Null
		#V2.20 changed to @()
		$OUs = @(Get-ADOrganizationalUnit -Filter * -Server $Domain `
		-Properties CanonicalName, DistinguishedName, Name, Created, ProtectedFromAccidentalDeletion -EA 0 | `
		Select-Object CanonicalName, DistinguishedName, Name, Created, ProtectedFromAccidentalDeletion | `
		Sort-Object CanonicalName)
		
		If($? -and $Null -ne $OUs)
		{
			[int]$OUCount = 0
			If($MSWORD -or $PDF)
			{
				$TableRange = $doc.Application.Selection.Range
				[int]$Columns = 6
				[int]$Rows = $OUs.Count + 1
				[int]$NumOUs = $OUs.Count
				[int]$xRow = 1
				[int]$UnprotectedOUs = 0 #added in V2.22

				$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				$Table.AutoFitBehavior($wdAutoFitFixed)
				$Table.Style = $Script:MyHash.Word_TableGrid
			
				$Table.rows.first.headingformat = $wdHeadingFormatTrue
				$Table.Borders.InsideLineStyle = $wdLineStyleSingle
				$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

				$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Text = "Name"
				
				$Table.Cell($xRow,2).Range.Font.Bold = $True
				$Table.Cell($xRow,2).Range.Text = "Created"
				
				$Table.Cell($xRow,3).Range.Font.Bold = $True
				$Table.Cell($xRow,3).Range.Text = "Protected"
				
				$Table.Cell($xRow,4).Range.Font.Bold = $True
				$Table.Cell($xRow,4).Range.Text = "# Users"
				
				$Table.Cell($xRow,5).Range.Font.Bold = $True
				$Table.Cell($xRow,5).Range.Text = "# Computers"
				
				$Table.Cell($xRow,6).Range.Font.Bold = $True
				$Table.Cell($xRow,6).Range.Text = "# Groups"

				ForEach($OU in $OUs)
				{
					$xRow++
					$OUCount++
					If($xRow % 2 -eq 0)
					{
						$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray05
						$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray05
						$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray05
						$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorGray05
						$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorGray05
						$Table.Cell($xRow,6).Shading.BackgroundPatternColor = $wdColorGray05
					}
					$OUDisplayName = $OU.CanonicalName.SubString($OU.CanonicalName.IndexOf("/")+1)
					Write-Verbose "$(Get-Date): `t`tProcessing OU $($OU.CanonicalName) - OU # $OUCount of $NumOUs"
					
					#get counts of users, computers and groups in the OU
					
					[int]$UserCount = 0
					[int]$ComputerCount = 0
					[int]$GroupCount = 0
					
					#V2.20 changed to @()
					$Results = @(Get-ADUser -Filter * -SearchBase $OU.DistinguishedName -Server $Domain -EA 0)
					$UserCount = $Results.Count

					#V2.20 changed to @()
					$Results = @(Get-ADComputer -Filter * -SearchBase $OU.DistinguishedName -Server $Domain -EA 0)
					$ComputerCount = $Results.Count

					#V2.20 changed to @()
					$Results = @(Get-ADGroup -Filter * -SearchBase $OU.DistinguishedName -Server $Domain -EA 0)
					$GroupCount = $Results.Count
					
					$Table.Cell($xRow,1).Range.Text = $OUDisplayName
					$Table.Cell($xRow,2).Range.Text = $OU.Created.ToString()
					If($OU.ProtectedFromAccidentalDeletion -eq $True)
					{
						$Table.Cell($xRow,3).Range.Text = "Yes"
					}
					Else
					{
						$Table.Cell($xRow,3).Range.Text = "No"
						#not added in V2.22 now
						##$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorYellow
						$Table.Cell($xRow,3).Range.Font.Bold = $True
						$UnprotectedOUs++
					}
					
					[string]$UserCountStr = "{0,7:N0}" -f $UserCount
					[string]$ComputerCountStr = "{0,7:N0}" -f $ComputerCount
					[string]$GroupCountStr = "{0,7:N0}" -f $GroupCount

					$Table.Cell($xRow,4).Range.ParagraphFormat.Alignment = $wdCellAlignVerticalTop
					$Table.Cell($xRow,4).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
					$Table.Cell($xRow,4).Range.Text = $UserCountStr
					$Table.Cell($xRow,5).Range.ParagraphFormat.Alignment = $wdCellAlignVerticalTop
					$Table.Cell($xRow,5).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
					$Table.Cell($xRow,5).Range.Text = $ComputerCountStr
					$Table.Cell($xRow,6).Range.ParagraphFormat.Alignment = $wdCellAlignVerticalTop
					$Table.Cell($xRow,6).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
					$Table.Cell($xRow,6).Range.Text = $GroupCountStr
					$Results = $Null
					$UserCountStr = $Null
					$ComputerCountStr = $Null
					$GroupCountStr = $Null
				}
				
				#set column widths
				$xcols = $table.columns

				ForEach($xcol in $xcols)
				{
					switch ($xcol.Index)
					{
					  1 {$xcol.width = 214; Break}
					  2 {$xcol.width = 68; Break}
					  3 {$xcol.width = 56; Break}
					  4 {$xcol.width = 56; Break}
					  5 {$xcol.width = 70; Break}
					  6 {$xcol.width = 56; Break}
					}
				}
				
				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
				$Table.AutoFitBehavior($wdAutoFitFixed)

				#return focus back to document
				$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

				#move to the end of the current document
				$selection.EndKey($wdStory,$wdMove) | Out-Null
				$TableRange = $Null
				$Table = $Null
				$Results = $Null
				$UserCountStr = $Null
				$ComputerCountStr = $Null
				$GroupCountStr = $Null

				#added in V2.22
				If($UnprotectedOUs -gt 0)
				{
					WriteWordLine 0 0 "There are $($UnprotectedOUs) unprotected OUs out of $($NumOUs) OUs"
				}
			}
			ElseIf($Text)
			{
				[int]$NumOUs = $OUs.Count
				[int]$UnprotectedOUs = 0 #added in V2.22
				#V2.16 addition
				[int]$MaxOUNameLength = ($OUs.CanonicalName.SubString($OUs[0].CanonicalName.IndexOf("/")+1) | measure-object -maximum -property length).maximum
				
				If($MaxOUNameLength -gt 4) #4 is length of "Name"
				{
					#2 is to allow for spacing between columns
					Line 1 ("Name" + (' ' * ($MaxOUNameLength - 2))) -NoNewLine
					Line 0 "Created                Protected # Users # Computers # Groups"
					Line 1 ('=' * $MaxOUNameLength) -NoNewLine
					Line 0 "==============================================================="
				}
				Else
				{
					Line 1 "Name  Created                Protected # Users # Computers # Groups"
					Line 1 "==================================================================="
				}

				ForEach($OU in $OUs)
				{
					$OUCount++
					$OUDisplayName = $OU.CanonicalName.SubString($OU.CanonicalName.IndexOf("/")+1)
					Write-Verbose "$(Get-Date): `t`tProcessing OU $($OU.CanonicalName) - OU # $OUCount of $NumOUs"
					
					#get counts of users, computers and groups in the OU
					
					[int]$UserCount = 0
					[int]$ComputerCount = 0
					[int]$GroupCount = 0
					
					#V2.20 changed to @()
					$Results = @(Get-ADUser -Filter * -SearchBase $OU.DistinguishedName -Server $Domain -EA 0)
					$UserCount = $Results.Count

					#V2.20 changed to @()
					$Results = @(Get-ADComputer -Filter * -SearchBase $OU.DistinguishedName -Server $Domain -EA 0)
					$ComputerCount = $Results.Count

					#V2.20 changed to @()
					$Results = @(Get-ADGroup -Filter * -SearchBase $OU.DistinguishedName -Server $Domain -EA 0)
					$GroupCount = $Results.Count
					
					If($OU.ProtectedFromAccidentalDeletion -eq $True)
					{
						$tmp = "Yes"
					}
					Else
					{
						$tmp = "NO"
						$UnprotectedOUs++
					}
					[string]$UserCountStr = "{0,7:N0}" -f $UserCount
					[string]$ComputerCountStr = "{0,11:N0}" -f $ComputerCount
					[string]$GroupCountStr = "{0,7:N0}" -f $GroupCount

					#V2.16 change
					If(($OUDisplayName).Length -lt ($MaxOUNameLength))
					{
						[int]$NumOfSpaces = ($MaxOUNameLength * -1) 
					}
                    Else
                    {
                        [int]$NumOfSpaces = -4
                    }
					Line 1 ( "{0,$NumOfSpaces}  {1,-22} {2,-9} {3,-7} {4,-12} {5,-7}" -f $OUDisplayName,$OU.Created.ToString(),$tmp,$UserCountStr,$ComputerCountStr,$GroupCountStr)

					$Results = $Null
					$UserCountStr = $Null
					$ComputerCountStr = $Null
					$GroupCountStr = $Null
				}
				Line 0 ""
				#added in V2.22
				If($UnprotectedOUs -gt 0)
				{
					Line 0 "There are $($UnprotectedOUs) unprotected OUs out of $($NumOUs) OUs"
				}
				$Results = $Null
				$UserCountStr = $Null
				$ComputerCountStr = $Null
				$GroupCountStr = $Null
			}
			ElseIf($HTML)
			{
				[int]$NumOUs = $OUs.Count
				[int]$UnprotectedOUs = 0 #added in V2.22
				$rowdata = @()
				ForEach($OU in $OUs)
				{
					$OUCount++
					$OUDisplayName = $OU.CanonicalName.SubString($OU.CanonicalName.IndexOf("/")+1)
					Write-Verbose "$(Get-Date): `t`tProcessing OU $($OU.CanonicalName) - OU # $OUCount of $NumOUs"
					
					#get counts of users, computers and groups in the OU
					
					[int]$UserCount = 0
					[int]$ComputerCount = 0
					[int]$GroupCount = 0
					
					#V2.20 changed to @()
					$Results = @(Get-ADUser -Filter * -SearchBase $OU.DistinguishedName -Server $Domain -EA 0)
					$UserCount = $Results.Count

					#V2.20 changed to @()
					$Results = @(Get-ADComputer -Filter * -SearchBase $OU.DistinguishedName -Server $Domain -EA 0)
					$ComputerCount = $Results.Count

					#V2.20 changed to @()
					$Results = @(Get-ADGroup -Filter * -SearchBase $OU.DistinguishedName -Server $Domain -EA 0)
					$GroupCount = $Results.Count
					
					[string]$UserCountStr = "{0,7:N0}" -f $UserCount
					[string]$ComputerCountStr = "{0,7:N0}" -f $ComputerCount
					[string]$GroupCountStr = "{0,7:N0}" -f $GroupCount

					If($OU.ProtectedFromAccidentalDeletion -eq $True)
					{
						$Protected = "Yes"
						$rowdata += @(,(
						$OUDisplayName,$htmlwhite,
						$OU.Created.ToString(),$htmlwhite,
						$Protected,$htmlwhite,
						$UserCountStr,$htmlwhite,
						$ComputerCountStr,$htmlwhite,
						$GroupCountStr,$htmlwhite))
					}
					Else
					{
						$Protected = "No"
						$UnprotectedOUs++
						$rowdata += @(,(
						$OUDisplayName,$htmlwhite,
						$OU.Created.ToString(),$htmlwhite,
						$Protected,$htmlwhite,
						$UserCountStr,$htmlwhite,
						$ComputerCountStr,$htmlwhite,
						$GroupCountStr,$htmlwhite))
					}

					$Results = $Null
					$UserCountStr = $Null
					$ComputerCountStr = $Null
					$GroupCountStr = $Null
				}
				$columnHeaders = @('Name',($htmlsilver -bor $htmlbold),
									'Created',($htmlsilver -bor $htmlbold),
									'Protected',($htmlsilver -bor $htmlbold),
									'# Users',($htmlsilver -bor $htmlbold),
									'# Computers',($htmlsilver -bor $htmlbold),
									'# Groups',($htmlsilver -bor $htmlbold)
									)
				$msg = ""
				$columnWidths = @("214","68","56","56","75","56")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "525"
				#added in V2.22
				If($UnprotectedOUs -gt 0)
				{
					WriteHTMLLine 0 0 "There are $($UnprotectedOUs) unprotected OUs out of $($NumOUs) OUs"
				}
			}
		}
		ElseIf(!$?)
		{
			$txt = "Error retrieving OU data for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 $txt
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}
		}
		Else
		{
			$txt = "No OU data was retrieved for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 $txt
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}
		}
		$First = $False
	}
}
#endregion

#region Group information
Function ProcessGroupInformation
{
	Write-Verbose "$(Get-Date): Writing group data"

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Groups"
	}
	ElseIf($Text)
	{
		Line 0 "///  Groups  \\\"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Groups&nbsp;&nbsp;\\\"
	}

	$First = $True

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date): `tProcessing groups in domain $($Domain)"
		If(($MSWORD -or $PDF) -and !$First)
		{
			#put each domain, starting with the second, on a new page
			$Script:selection.InsertNewPage()
		}
		
		If($Domain -eq $Script:ForestRootDomain)
		{
			$txt = "Domain $($Domain) (Forest Root)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
			}
			ElseIf($Text)
			{
				Line 0 "///  $($txt)  \\\"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
			}
		}
		Else
		{
			$txt = "Domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
			}
			ElseIf($Text)
			{
				Line 0 "///  $($txt)  \\\"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
			}
		}

		#get all Groups for the domain
		$Groups = $Null
		$Groups = Get-ADGroup -Filter * -Server $Domain -Properties Name, GroupCategory, GroupType -EA 0 | Sort-Object Name

		If($? -and $Null -ne $Groups)
		{
			#get counts
			
			Write-Verbose "$(Get-Date): `t`tGetting counts"
			
			[int]$SecurityCount        = 0
			[int]$DistributionCount    = 0
			[int]$GlobalCount          = 0
			[int]$UniversalCount       = 0
			[int]$DomainLocalCount     = 0
			[int]$ContactsCount        = 0
			[int]$GroupsWithSIDHistory = 0
			
			Write-Verbose "$(Get-Date): `t`t`tSecurity Groups"
			#V2.20 changed to @()
			$Results = @($groups | Where-Object {$_.groupcategory -eq "Security"})
			
			[int]$SecurityCount = $Results.Count
			
			Write-Verbose "$(Get-Date): `t`t`tDistribution Groups"
			#V2.20 changed to @()
			$Results = @($groups | Where-Object {$_.groupcategory -eq "Distribution"})
			
			[int]$DistributionCount = $Results.Count

			Write-Verbose "$(Get-Date): `t`t`tGlobal Groups"
			#V2.20 changed to @()
			$Results = @($groups | Where-Object {$_.groupscope -eq "Global"})

			[int]$GlobalCount = $Results.Count

			Write-Verbose "$(Get-Date): `t`t`tUniversal Groups"
			#V2.20 changed to @()
			$Results = @($groups | Where-Object {$_.groupscope -eq "Universal"})

			[int]$UniversalCount = $Results.Count
			
			Write-Verbose "$(Get-Date): `t`t`tDomain Local Groups"
			#V2.20 changed to @()
			$Results = @($groups | Where-Object {$_.groupscope -eq "DomainLocal"})

			[int]$DomainLocalCount = $Results.Count

			Write-Verbose "$(Get-Date): `t`t`tGroups with SID History"
			$Results = $Null
			#V2.20 changed to @()
			$Results = @(Get-ADObject -LDAPFilter "(sIDHistory=*)" -Server $Domain -Property objectClass, sIDHistory -EA 0)

			[int]$GroupsWithSIDHistory = ($Results | Where-Object {$_.objectClass -eq 'group'}).Count

			Write-Verbose "$(Get-Date): `t`t`tContacts"
			$Results = $Null
			#V2.20 changed to @()
			$Results = @(Get-ADObject -LDAPFilter "objectClass=Contact" -Server $Domain -EA 0)

			[int]$ContactsCount = $Results.Count

			[string]$TotalCountStr           = "{0,7:N0}" -f ($SecurityCount + $DistributionCount)
			[string]$SecurityCountStr        = "{0,7:N0}" -f $SecurityCount
			[string]$DomainLocalCountStr     = "{0,7:N0}" -f $DomainLocalCount
			[string]$GlobalCountStr          = "{0,7:N0}" -f $GlobalCount
			[string]$UniversalCountStr       = "{0,7:N0}" -f $UniversalCount
			[string]$DistributionCountStr    = "{0,7:N0}" -f $DistributionCount
			[string]$GroupsWithSIDHistoryStr = "{0,7:N0}" -f $GroupsWithSIDHistory
			[string]$ContactsCountStr        = "{0,7:N0}" -f $ContactsCount
			
			Write-Verbose "$(Get-Date): `t`tBuild groups table"
			If($MSWORD -or $PDF)
			{
				$TableRange = $Script:doc.Application.Selection.Range
				[int]$Columns = 2
				[int]$Rows = 8
				$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
				$Table.Style = $Script:MyHash.Word_TableGrid
			
				$Table.Borders.InsideLineStyle = $wdLineStyleSingle
				$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
				$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(1,1).Range.Font.Bold = $True
				$Table.Cell(1,1).Range.Text = "Total Groups"
				$Table.Cell(1,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(1,2).Range.Text = $TotalCountStr
				$Table.Cell(2,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(2,1).Range.Font.Bold = $True
				$Table.Cell(2,1).Range.Text = "`tSecurity Groups"
				$Table.Cell(2,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(2,2).Range.Text = $SecurityCountStr
				$Table.Cell(3,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(3,1).Range.Font.Bold = $True
				$Table.Cell(3,1).Range.Text = "`t`tDomain Local"
				$Table.Cell(3,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(3,2).Range.Text = $DomainLocalCountStr
				$Table.Cell(4,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(4,1).Range.Font.Bold = $True
				$Table.Cell(4,1).Range.Text = "`t`tGlobal"
				$Table.Cell(4,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(4,2).Range.Text = $GlobalCountStr
				$Table.Cell(5,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(5,1).Range.Font.Bold = $True
				$Table.Cell(5,1).Range.Text = "`t`tUniversal"
				$Table.Cell(5,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(5,2).Range.Text = $UniversalCountStr
				$Table.Cell(6,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(6,1).Range.Font.Bold = $True
				$Table.Cell(6,1).Range.Text = "`tDistribution Groups"
				$Table.Cell(6,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(6,2).Range.Text = $DistributionCountStr
				$Table.Cell(7,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(7,1).Range.Font.Bold = $True
				$Table.Cell(7,1).Range.Text = "Groups with SID History"
				$Table.Cell(7,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(7,2).Range.Text = $GroupsWithSIDHistoryStr
				$Table.Cell(8,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(8,1).Range.Font.Bold = $True
				$Table.Cell(8,1).Range.Text = "Contacts"
				$Table.Cell(8,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(8,2).Range.Text = $ContactsCountStr

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
				$Table.AutoFitBehavior($wdAutoFitContent)

				#return focus back to document
				$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

				#move to the end of the current document
				$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
				$TableRange = $Null
				$Table = $Null
			}
			ElseIf($Text)
			{
				Line 1 "Total Groups`t`t`t: " $TotalCountStr
				Line 1 "`tSecurity Groups`t`t: " $SecurityCountStr
				Line 1 "`t`tDomain Local`t: " $DomainLocalCountStr
				Line 1 "`t`tGlobal`t`t: " $GlobalCountStr
				Line 1 "`t`tUniversal`t: " $UniversalCountStr
				Line 1 "`tDistribution Groups`t: " $DistributionCountStr
				Line 1 "Groups with SID History`t`t: " $GroupsWithSIDHistoryStr
				Line 1 "Contacts`t`t`t: " $ContactsCountStr
				Line 0 ""
			}
			ElseIf($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Total Groups",($htmlsilver -bor $htmlbold),$TotalCountStr,$htmlwhite)
				$rowdata += @(,("     Security Groups",($htmlsilver -bor $htmlbold),$SecurityCountStr,$htmlwhite))
				$rowdata += @(,("          Domain Local",($htmlsilver -bor $htmlbold),$DomainLocalCountStr,$htmlwhite))
				$rowdata += @(,("          Global",($htmlsilver -bor $htmlbold),$GlobalCountStr,$htmlwhite))
				$rowdata += @(,("          Universal",($htmlsilver -bor $htmlbold),$UniversalCountStr,$htmlwhite))
				$rowdata += @(,("     Distribution Groups",($htmlsilver -bor $htmlbold),$DistributionCountStr,$htmlwhite))
				$rowdata += @(,("Groups with SID History",($htmlsilver -bor $htmlbold),$GroupsWithSIDHistoryStr,$htmlwhite))
				$rowdata += @(,("Contacts",($htmlsilver -bor $htmlbold),$ContactsCountStr,$htmlwhite))

				$msg = ""
				$columnWidths = @("150","75")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "225"
				WriteHTMLLine 0 0 " "
			}
			
			#get members of privileged groups
			$DomainInfo = $Null
			$DomainInfo = Get-ADDomain -Identity $Domain -EA 0
			
			If($? -and $Null -ne $DomainInfo)
			{
				$DomainAdminsSID     = "$($DomainInfo.DomainSID)-512"
				$EnterpriseAdminsSID = "$($DomainInfo.DomainSID)-519"
				$SchemaAdminsSID     = "$($DomainInfo.DomainSID)-518"
			}
			Else
			{
				$DomainAdminsSID     = $Null
				$EnterpriseAdminsSID = $Null
				$SchemaAdminsSID     = $Null
			}
			
			Write-Verbose "$(Get-Date): `t`tListing domain admins"
			$Admins = $Null
			#V2.20 changed to @()
			$Admins = @(Get-ADGroupMember -Identity $DomainAdminsSID -Server $Domain -EA 0)
			
			If($? -and $Null -ne $Admins)
			{
				[int]$AdminsCount = $Admins.Count
				$Admins = $Admins | Sort-Object Name
				[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
				
				If($MSWORD -or $PDF)
				{
					WriteWordLine 3 0 "Privileged Groups"
					WriteWordLine 4 0 "Domain Admins ($($AdminsCountStr) members):"
					$TableRange = $Script:doc.Application.Selection.Range
					[int]$Columns = 4
					[int]$Rows = $AdminsCount + 1
					[int]$xRow = 1
					$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.AutoFitBehavior($wdAutoFitFixed)
					$Table.Style = $Script:MyHash.Word_TableGrid
			
					$Table.rows.first.headingformat = $wdHeadingFormatTrue
					$Table.Borders.InsideLineStyle = $wdLineStyleSingle
					$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

					$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,1).Range.Font.Bold = $True
					$Table.Cell($xRow,1).Range.Text = "Name"
					$Table.Cell($xRow,2).Range.Font.Bold = $True
					$Table.Cell($xRow,2).Range.Text = "Password Last Changed"
					$Table.Cell($xRow,3).Range.Font.Bold = $True
					$Table.Cell($xRow,3).Range.Text = "Password Never Expires"
					$Table.Cell($xRow,4).Range.Font.Bold = $True
					$Table.Cell($xRow,4).Range.Text = "Account Enabled"
					ForEach($Admin in $Admins)
					{
						$xRow++
						
						$User = Get-ADUser -Identity $Admin.SID -Server $Domain -Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0 

						If($? -and $Null -ne $User)
						{
							$Table.Cell($xRow,1).Range.Text = $User.Name
							If($Null -eq $User.PasswordLastSet)
							{
								$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorRed
								$Table.Cell($xRow,2).Range.Font.Bold  = $True
								$Table.Cell($xRow,2).Range.Font.Color = $WDColorBlack
								$Table.Cell($xRow,2).Range.Text = "No Date Set"
							}
							Else
							{
								$Table.Cell($xRow,2).Range.Text = (get-date $User.PasswordLastSet -f d)
							}
							If($User.PasswordNeverExpires -eq $True)
							{
								$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorRed
								$Table.Cell($xRow,3).Range.Font.Bold  = $True
								$Table.Cell($xRow,3).Range.Font.Color = $WDColorBlack
								$Table.Cell($xRow,3).Range.Text = "True"
							}
							Else
							{
								$Table.Cell($xRow,3).Range.Text = "False"
							}
							If($User.Enabled -eq $True)
							{
								$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorRed
								$Table.Cell($xRow,4).Range.Font.Bold  = $True
								$Table.Cell($xRow,4).Range.Font.Color = $WDColorBlack
								$Table.Cell($xRow,4).Range.Text = "True"
							}
							Else
							{
								$Table.Cell($xRow,4).Range.Text = "False"
							}
						}
						Else
						{
							$Table.Cell($xRow,1).Range.Text = $Admin.SID
							$Table.Cell($xRow,2).Range.Text = "Unknown"
							$Table.Cell($xRow,3).Range.Text = "Unknown"
							$Table.Cell($xRow,4).Range.Text = "Unknown"
						}
					}
					
					#set column widths
					$xcols = $table.columns

					ForEach($xcol in $xcols)
					{
						switch ($xcol.Index)
						{
						  1 {$xcol.width = 200; Break}
						  2 {$xcol.width = 66; Break}
						  3 {$xcol.width = 56; Break}
						  4 {$xcol.width = 56; Break}
						}
					}
					
					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
					$Table.AutoFitBehavior($wdAutoFitFixed)

					#return focus back to document
					$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

					#move to the end of the current document
					$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
					$TableRange = $Null
					$Table = $Null
				}
				ElseIf($Text)
				{
					Line 0 "Privileged Groups"
					Line 1 "Domain Admins ($AdminsCountStr members):"
					#V2.16 addition
					Line 2 "                                                   Password    Password          "
					Line 2 "                                                   Last        Never      Account"
					Line 2 "Name                                               Changed     Expires    Enabled"
					Line 2 "================================================================================="
					ForEach($Admin in $Admins)
					{
						$User = Get-ADUser -Identity $Admin.SID -Server $Domain -Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0

						If($? -and $Null -ne $User)
						{
							If($Null -eq $User.PasswordLastSet)
							{
								$PasswordLastSet = "No Date Set"
							}
							Else
							{
								$PasswordLastSet = (get-date $User.PasswordLastSet -f d)
							}
							If($User.PasswordNeverExpires -eq $True)
							{
								$PasswordNeverExpires = "True"
							}
							Else
							{
								$PasswordNeverExpires = "False"
							}
							If($User.Enabled -eq $True)
							{
								$UserEnabled = "True"
							}
							Else
							{
								$UserEnabled = "False"
							}
							#V2.16 change
							Line 2 ( "{0,-50} {1,-11} {2,-10} {3,-5}" -f $User.Name,$PasswordLastSet,$PasswordNeverExpires,$UserEnabled)
						}
						Else
						{
							#V2.16 change
							Line 2 ( "{0,-50} {1,-11} {2,-10} {3,-5}" -f $Admin.SID,"Unknown","Unknown","Unknown")
						}
					}
					Line 0 ""
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 3 0 "Privileged Groups"
					WriteHTMLLine 4 0 "Domain Admins ($($AdminsCountStr) members):"
					$rowdata = @()
					ForEach($Admin in $Admins)
					{
						$User = Get-ADUser -Identity $Admin.SID -Server $Domain -Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0 

						If($? -and $Null -ne $User)
						{
							$UserName = $User.Name
							If($Null -eq $User.PasswordLastSet)
							{
								$PasswordLastSet = "No Date Set"
							}
							Else
							{
								$PasswordLastSet = (get-date $User.PasswordLastSet -f d)
							}
							If($User.PasswordNeverExpires -eq $True)
							{
								$PasswordNeverExpires = "True"
							}
							Else
							{
								$PasswordNeverExpires = "False"
							}
							If($User.Enabled -eq $True)
							{
								$Enabled = "True"
							}
							Else
							{
								$Enabled = "False"
							}
						}
						Else
						{
							$UserName = $Admin.SID
							$PasswordLastSet = "Unknown"
							$PasswordNeverExpires = "Unknown"
							$Enabled = "Unknown"
						}
						$rowdata += @(,(
						$UserName,$htmlwhite,
						$PasswordLastSet,$htmlwhite,
						$PasswordNeverExpires,$htmlwhite,
						$Enabled,$htmlwhite))
					}
					
					$columnHeaders = @(
					'Name',($htmlsilver -bor $htmlbold),
					'Password Last Changed',($htmlsilver -bor $htmlbold),
					'Password Never Expires',($htmlsilver -bor $htmlbold),
					'Account Enabled',($htmlsilver -bor $htmlbold)
					)
					
					$columnWidths = @("200","66","56","56")
					$msg = ""
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "378"
					WriteHTMLLine 0 0 " "
				}
			}
			ElseIf(!$?)
			{
				$txt = "Unable to retrieve Domain Admins group membership"
				If($MSWORD -or $PDF)
				{
					WriteWordLine 0 0 $txt "" $Null 0 $False $True
				}
				ElseIf($Text)
				{
					Line 0 $txt
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 0 0 $txt
				}
			}
			Else
			{
				$txt = "Domain Admins: <None>"
				If($MSWORD -or $PDF)
				{
					WriteWordLine 4 0 $txt
				}
				ElseIf($Text)
				{
					Line 0 $txt
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 4 0 "Domain Admins: None"
				}
			}

			If($Domain -eq $Script:ForestRootDomain)
			{
				Write-Verbose "$(Get-Date): `t`tListing enterprise admins"
			
				#V2.20 changed to @()
				$Admins = @(Get-ADGroupMember -Identity $EnterpriseAdminsSID -Server $Domain -EA 0)
				
				If($? -and $Null -ne $Admins)
				{
					[int]$AdminsCount = $Admins.Count
					$Admins = $Admins | Sort-Object Name
					[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
					
					If($MSWORD -or $PDF)
					{
						WriteWordLine 4 0 "Enterprise Admins ($($AdminsCountStr) members):"
						$TableRange = $Script:doc.Application.Selection.Range
						[int]$Columns = 5
						[int]$Rows = $AdminsCount + 1
						[int]$xRow = 1
						$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
						$Table.AutoFitBehavior($wdAutoFitFixed)
						$Table.Style = $Script:MyHash.Word_TableGrid
			
						$Table.rows.first.headingformat = $wdHeadingFormatTrue
						$Table.Borders.InsideLineStyle = $wdLineStyleSingle
						$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

						$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell($xRow,1).Range.Font.Bold = $True
						$Table.Cell($xRow,1).Range.Text = "Name"
						$Table.Cell($xRow,2).Range.Font.Bold = $True
						$Table.Cell($xRow,2).Range.Text = "Domain"
						$Table.Cell($xRow,3).Range.Font.Bold = $True
						$Table.Cell($xRow,3).Range.Text = "Password Last Changed"
						$Table.Cell($xRow,4).Range.Font.Bold = $True
						$Table.Cell($xRow,4).Range.Text = "Password Never Expires"
						$Table.Cell($xRow,5).Range.Font.Bold = $True
						$Table.Cell($xRow,5).Range.Text = "Account Enabled"
						ForEach($Admin in $Admins)
						{
							$xRow++
							$xArray = $Admin.DistinguishedName.Split(",")
							$xServer = ""
							$xCnt = 0
							ForEach($xItem in $xArray)
							{
								$xCnt++
								If($xItem.StartsWith("DC="))
								{
									$xtmp = $xItem.Substring($xItem.IndexOf("=")+1)
									If($xCnt -eq $xArray.Count)
									{
										$xServer += $xTmp
									}
									Else
									{
										$xServer += "$($xTmp)."
									}
								}
							}

							If($Admin.ObjectClass -eq 'user')
							{
								$User = Get-ADUser -Identity $Admin.SID.value -Server $xServer `
								-Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0
							}
							ElseIf($Admin.ObjectClass -eq 'group')
							{
								$User = Get-ADGroup -Identity $Admin.SID.value -Server $xServer -EA 0
							}
							Else
							{
								$User = $Null
							}
							
							If($? -and $Null -ne $User)
							{
								If($Admin.ObjectClass -eq 'user')
								{
									$Table.Cell($xRow,1).Range.Text = $User.Name
									$Table.Cell($xRow,2).Range.Text = $xServer
									If($Null -eq $User.PasswordLastSet)
									{
										$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorRed
										$Table.Cell($xRow,3).Range.Font.Bold  = $True
										$Table.Cell($xRow,3).Range.Font.Color = $WDColorBlack
										$Table.Cell($xRow,3).Range.Text = "No Date Set"
									}
									Else
									{
										$Table.Cell($xRow,3).Range.Text = (get-date $User.PasswordLastSet -f d)
									}
									If($User.PasswordNeverExpires -eq $True)
									{
										$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorRed
										$Table.Cell($xRow,4).Range.Font.Bold  = $True
										$Table.Cell($xRow,4).Range.Font.Color = $WDColorBlack
									}
									$Table.Cell($xRow,4).Range.Text = $User.PasswordNeverExpires.ToString()
									If($User.Enabled -eq $False)
									{
										$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorRed
										$Table.Cell($xRow,5).Range.Font.Bold  = $True
										$Table.Cell($xRow,5).Range.Font.Color = $WDColorBlack
									}
									$Table.Cell($xRow,5).Range.Text = $User.Enabled.ToString()
								}
								ElseIf($Admin.ObjectClass -eq 'group')
								{
									$Table.Cell($xRow,1).Range.Text = "$($User.Name) (group)"
									$Table.Cell($xRow,2).Range.Text = $xServer
									$Table.Cell($xRow,3).Range.Text = "N/A"
									$Table.Cell($xRow,4).Range.Text = "N/A"
									$Table.Cell($xRow,5).Range.Text = "N/A"
								}
								
							}
							Else
							{
								$Table.Cell($xRow,1).Range.Text = $Admin.SID.Value
								$Table.Cell($xRow,2).Range.Text = $xServer
								$Table.Cell($xRow,3).Range.Text = "Unknown"
								$Table.Cell($xRow,4).Range.Text = "Unknown"
								$Table.Cell($xRow,5).Range.Text = "Unknown"
							}
						}
					
						#set column widths
						$xcols = $table.columns

						ForEach($xcol in $xcols)
						{
							switch ($xcol.Index)
							{
							  1 {$xcol.width = 100; Break}
							  2 {$xcol.width = 108; Break}
							  3 {$xcol.width = 66; Break}
							  4 {$xcol.width = 56; Break}
							  5 {$xcol.width = 56; Break}
							}
						}

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
						$Table.AutoFitBehavior($wdAutoFitFixed)

						#return focus back to document
						$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

						#move to the end of the current document
						$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
						$TableRange = $Null
						$Table = $Null
					}
					ElseIf($Text)
					{
						Line 1 "Enterprise Admins ($AdminsCountStr members):"
						#V2.16 addition
						Line 2 "                                                                              Password   Password          "
						Line 2 "                                                                              Last       Never      Account"
						Line 2 "Name                                                Domain                    Changed    Expires    Enabled"
						Line 2 "==========================================================================================================="
						ForEach($Admin in $Admins)
						{
							$xArray = $Admin.DistinguishedName.Split(",")
							$xServer = ""
							$xCnt = 0
							ForEach($xItem in $xArray)
							{
								$xCnt++
								If($xItem.StartsWith("DC="))
								{
									$xtmp = $xItem.Substring($xItem.IndexOf("=")+1)
									If($xCnt -eq $xArray.Count)
									{
										$xServer += $xTmp
									}
									Else
									{
										$xServer += "$($xTmp)."
									}
								}
							}

							If($Admin.ObjectClass -eq 'user')
							{
								$User = Get-ADUser -Identity $Admin.SID.value -Server $xServer `
								-Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0
							}
							ElseIf($Admin.ObjectClass -eq 'group')
							{
								$User = Get-ADGroup -Identity $Admin.SID.value -Server $xServer -EA 0
							}
							Else
							{
								$User = $Null
							}
							
							If($? -and $Null -ne $User)
							{
								If($Admin.ObjectClass -eq 'user')
								{
									If($Null -eq $User.PasswordLastSet)
									{
										$PasswordLastSet = "No Date Set"
									}
									Else
									{
										$PasswordLastSet = (get-date $User.PasswordLastSet -f d)
									}
									#V2.16 change
									Line 2 ( "{0,-50}  {1,-25} {2,-10} {3,-10} {4,-5}" -f $User.Name,$xServer,$PasswordLastSet,$User.PasswordNeverExpires.ToString(),$User.Enabled.ToString())
								}
								ElseIf($Admin.ObjectClass -eq 'group')
								{
									#V2.16 change
									Line 2 ( "{0,-43} (group) {1,-25} {2,-10} {3,-10} {4,-5}" -f $User.Name,$xServer,"N/A","N/A","N/A")
								}
							}
							Else
							{
								#v2.16 change
								Line 2 ( "{0,-50} {1,-25} {2,-10} {3,-10} {4,-5}" -f $Admin.SID.Value,$xServer,"Unknown","Unknown","Unknown")
							}
						}
						Line 0 ""
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 4 0 "Enterprise Admins ($($AdminsCountStr) members):"
						$rowdata = @()
						ForEach($Admin in $Admins)
						{
							$xArray = $Admin.DistinguishedName.Split(",")
							$xServer = ""
							$xCnt = 0
							ForEach($xItem in $xArray)
							{
								$xCnt++
								If($xItem.StartsWith("DC="))
								{
									$xtmp = $xItem.Substring($xItem.IndexOf("=")+1)
									If($xCnt -eq $xArray.Count)
									{
										$xServer += $xTmp
									}
									Else
									{
										$xServer += "$($xTmp)."
									}
								}
							}

							If($Admin.ObjectClass -eq 'user')
							{
								$User = Get-ADUser -Identity $Admin.SID.value -Server $xServer `
								-Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0
							}
							ElseIf($Admin.ObjectClass -eq 'group')
							{
								$User = Get-ADGroup -Identity $Admin.SID.value -Server $xServer -EA 0
							}
							Else
							{
								$User = $Null
							}
							
							If($? -and $Null -ne $User)
							{
								If($Admin.ObjectClass -eq 'user')
								{
									$UserName = $User.Name
									$Domain = $xServer
									If($Null -eq $User.PasswordLastSet)
									{
										$PasswordLastSet = "No Date Set"
									}
									Else
									{
										$PasswordLastSet = (get-date $User.PasswordLastSet -f d)
									}
									$PasswordNeverExpires = $User.PasswordNeverExpires.ToString()
									$Enabled = $User.Enabled.ToString()
								}
								ElseIf($Admin.ObjectClass -eq 'group')
								{
									$UserName = "$($User.Name) (group)"
									$Domain = $xServer
									$PasswordLastSet = "N/A"
									$PasswordNeverExpires = "N/A"
									$Enabled = "N/A"
								}
								
							}
							Else
							{
								$UserName = $Admin.SID.Value
								$Domain = $xServer
								$PasswordLastSet = "Unknown"
								$PasswordNeverExpires = "Unknown"
								$Enabled = "Unknown"
							}
							$rowdata += @(,(
							$UserName,$htmlwhite,
							$Domain,$htmlwhite,
							$PasswordLastSet,$htmlwhite,
							$PasswordNeverExpires,$htmlwhite,
							$Enabled,$htmlwhite))
						}

						$columnHeaders = @(
						'Name',($htmlsilver -bor $htmlbold),
						'Domain',($htmlsilver -bor $htmlbold),
						'Password Last Changed',($htmlsilver -bor $htmlbold),
						'Password Never Expires',($htmlsilver -bor $htmlbold),
						'Account Enabled',($htmlsilver -bor $htmlbold)
						)
						
						$columnWidths = @("100","108","66","56","56")
						$msg = ""
						FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "386"
						WriteHTMLLine 0 0 " "
					}
				}
				ElseIf(!$?)
				{
					$txt1 = "Enterprise Admins:"
					$txt2 = "Unable to retrieve Enterprise Admins group membership"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 4 0 $txt1
						WriteWordLine 0 0 $txt2 "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 0 $txt1
						Line 0 $txt2
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 4 0 $txt1
						WriteHTMLLine 0 0 $txt2
					}
				}
				Else
				{
					$txt1 = "Enterprise Admins:"
					$txt2 = "<None>"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 4 0 $txt1
						WriteWordLine 0 0 $txt2
					}
					ElseIf($Text)
					{
						Line 0 $txt1
						Line 0 $txt2
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 4 0 $txt1
						WriteHTMLLine 0 0 "None"
					}
				}
			}
			
			If($Domain -eq $Script:ForestRootDomain)
			{
				Write-Verbose "$(Get-Date): `t`tListing schema admins"
			
				#V2.20 changed to @()
				$Admins = @(Get-ADGroupMember -Identity $SchemaAdminsSID -Server $Domain -EA 0)
				
				If($? -and $Null -ne $Admins)
				{
					[int]$AdminsCount = $Admins.Count
					$Admins = $Admins | Sort-Object Name
					[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
					
					If($MSWORD -or $PDF)
					{
						WriteWordLine 4 0 "Schema Admins ($($AdminsCountStr) members): "
						$TableRange = $Script:doc.Application.Selection.Range
						[int]$Columns = 5
						[int]$Rows = $AdminsCount + 1
						[int]$xRow = 1
						$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
						$Table.AutoFitBehavior($wdAutoFitFixed)
						$Table.Style = $Script:MyHash.Word_TableGrid
			
						$Table.rows.first.headingformat = $wdHeadingFormatTrue
						$Table.Borders.InsideLineStyle = $wdLineStyleSingle
						$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

						$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell($xRow,1).Range.Font.Bold = $True
						$Table.Cell($xRow,1).Range.Text = "Name"
						$Table.Cell($xRow,2).Range.Font.Bold = $True
						$Table.Cell($xRow,2).Range.Text = "Domain"
						$Table.Cell($xRow,3).Range.Font.Bold = $True
						$Table.Cell($xRow,3).Range.Text = "Password Last Changed"
						$Table.Cell($xRow,4).Range.Font.Bold = $True
						$Table.Cell($xRow,4).Range.Text = "Password Never Expires"
						$Table.Cell($xRow,5).Range.Font.Bold = $True
						$Table.Cell($xRow,5).Range.Text = "Account Enabled"
						ForEach($Admin in $Admins)
						{
							$xRow++
							$xArray = $Admin.DistinguishedName.Split(",")
							$xServer = ""
							$xCnt = 0
							ForEach($xItem in $xArray)
							{
								$xCnt++
								If($xItem.StartsWith("DC="))
								{
									$xtmp = $xItem.Substring($xItem.IndexOf("=")+1)
									If($xCnt -eq $xArray.Count)
									{
										$xServer += $xTmp
									}
									Else
									{
										$xServer += "$($xTmp)."
									}
								}
							}

							If($Admin.ObjectClass -eq 'user')
							{
								$User = Get-ADUser -Identity $Admin.SID.value -Server $xServer `
								-Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0 
							}
							ElseIf($Admin.ObjectClass -eq 'group')
							{
								$User = Get-ADGroup -Identity $Admin.SID.value -Server $xServer -EA 0 
							}
							Else
							{
								$User = $Null
							}
							
							If($? -and $Null -ne $User)
							{
								If($Admin.ObjectClass -eq 'user')
								{
									$Table.Cell($xRow,1).Range.Text = $User.Name
									$Table.Cell($xRow,2).Range.Text = $xServer
									If($Null -eq $User.PasswordLastSet)
									{
										$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorRed
										$Table.Cell($xRow,3).Range.Font.Bold  = $True
										$Table.Cell($xRow,3).Range.Font.Color = $WDColorBlack
										$Table.Cell($xRow,3).Range.Text = "No Date Set"
									}
									Else
									{
										$Table.Cell($xRow,3).Range.Text = (get-date $User.PasswordLastSet -f d)
									}
									If($User.PasswordNeverExpires -eq $True)
									{
										$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorRed
										$Table.Cell($xRow,4).Range.Font.Bold  = $True
										$Table.Cell($xRow,4).Range.Font.Color = $WDColorBlack
									}
									$Table.Cell($xRow,4).Range.Text = $User.PasswordNeverExpires.ToString()
									If($User.Enabled -eq $False)
									{
										$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorRed
										$Table.Cell($xRow,5).Range.Font.Bold  = $True
										$Table.Cell($xRow,5).Range.Font.Color = $WDColorBlack
									}
									$Table.Cell($xRow,5).Range.Text = $User.Enabled.ToString()
									#$Table.Cell($xRow,6).Range.Text = ""
								}
								ElseIf($Admin.ObjectClass -eq 'group')
								{
									$Table.Cell($xRow,1).Range.Text = "$($User.Name) (group)"
									$Table.Cell($xRow,2).Range.Text = $xServer
									$Table.Cell($xRow,3).Range.Text = "N/A"
									$Table.Cell($xRow,4).Range.Text = "N/A"
									$Table.Cell($xRow,5).Range.Text = "N/A"
								}
								
							}
							Else
							{
								$Table.Cell($xRow,1).Range.Text = $Admin.SID.Value
								$Table.Cell($xRow,2).Range.Text = $xServer
								$Table.Cell($xRow,3).Range.Text = "Unknown"
								$Table.Cell($xRow,4).Range.Text = "Unknown"
								$Table.Cell($xRow,5).Range.Text = "Unknown"
							}
						}
					
						#set column widths
						$xcols = $table.columns

						ForEach($xcol in $xcols)
						{
							switch ($xcol.Index)
							{
							  1 {$xcol.width = 100; Break}
							  2 {$xcol.width = 108; Break}
							  3 {$xcol.width = 66; Break}
							  4 {$xcol.width = 56; Break}
							  5 {$xcol.width = 56; Break}
							}
						}
						
						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
						$Table.AutoFitBehavior($wdAutoFitFixed)

						#return focus back to document
						$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

						#move to the end of the current document
						$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
						$TableRange = $Null
						$Table = $Null
					}
					ElseIf($Text)
					{
						Line 1 "Schema Admins ($($AdminsCountStr) members): "
						#V2.16 addition
						Line 2 "                                                                              Password   Password          "
						Line 2 "                                                                              Last       Never      Account"
						Line 2 "Name                                                Domain                    Changed    Expires    Enabled"
						Line 2 "==========================================================================================================="
						ForEach($Admin in $Admins)
						{
							$xArray = $Admin.DistinguishedName.Split(",")
							$xServer = ""
							$xCnt = 0
							ForEach($xItem in $xArray)
							{
								$xCnt++
								If($xItem.StartsWith("DC="))
								{
									$xtmp = $xItem.Substring($xItem.IndexOf("=")+1)
									If($xCnt -eq $xArray.Count)
									{
										$xServer += $xTmp
									}
									Else
									{
										$xServer += "$($xTmp)."
									}
								}
							}

							If($Admin.ObjectClass -eq 'user')
							{
								$User = Get-ADUser -Identity $Admin.SID.value -Server $xServer `
								-Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0 
							}
							ElseIf($Admin.ObjectClass -eq 'group')
							{
								$User = Get-ADGroup -Identity $Admin.SID.value -Server $xServer -EA 0 
							}
							Else
							{
								$User = $Null
							}
							
							If($? -and $Null -ne $User)
							{
								If($Admin.ObjectClass -eq 'user')
								{
									If($Null -eq $User.PasswordLastSet)
									{
										$PasswordLastSet = "No Date Set"
									}
									Else
									{
										$PasswordLastSet = (get-date $User.PasswordLastSet -f d)
									}
									#V2.16 change
									Line 2 ( "{0,-50}  {1,-25} {2,-10} {3,-10} {4,-5}" -f $User.Name,$xServer,$PasswordLastSet,$User.PasswordNeverExpires.ToString(),$User.Enabled.ToString())
								}
								ElseIf($Admin.ObjectClass -eq 'group')
								{
									#V2.16 change
									Line 2 ( "{0,-43} (group) {1,-25} {2,-10} {3,-10} {4,-5}" -f $User.Name,$xServer,"N/A","N/A","N/A")
								}
							}
							Else
							{
								#v2.16 change
								Line 2 ( "{0,-50} {1,-25} {2,-10} {3,-10} {4,-5}" -f $Admin.SID.Value,$xServer,"Unknown","Unknown","Unknown")
							}
						}
						Line 0 ""
					}
					ElseIf($HTML)
					{
						$rowdata = @()
						WriteHTMLLine 4 0 "Schema Admins ($($AdminsCountStr) members): "
						ForEach($Admin in $Admins)
						{
							$xArray = $Admin.DistinguishedName.Split(",")
							$xServer = ""
							$xCnt = 0
							ForEach($xItem in $xArray)
							{
								$xCnt++
								If($xItem.StartsWith("DC="))
								{
									$xtmp = $xItem.Substring($xItem.IndexOf("=")+1)
									If($xCnt -eq $xArray.Count)
									{
										$xServer += $xTmp
									}
									Else
									{
										$xServer += "$($xTmp)."
									}
								}
							}

							If($Admin.ObjectClass -eq 'user')
							{
								$User = Get-ADUser -Identity $Admin.SID.value -Server $xServer `
								-Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0
							}
							ElseIf($Admin.ObjectClass -eq 'group')
							{
								$User = Get-ADGroup -Identity $Admin.SID.value -Server $xServer -EA 0
							}
							Else
							{
								$User = $Null
							}
							
							If($? -and $Null -ne $User)
							{
								If($Admin.ObjectClass -eq 'user')
								{
									$UserName = $User.Name
									$Domain = $xServer
									If($Null -eq $User.PasswordLastSet)
									{
										$PasswordLastSet = "No Date Set"
									}
									Else
									{
										$PasswordLastSet = (get-date $User.PasswordLastSet -f d)
									}
									$PasswordNeverExpires = $User.PasswordNeverExpires.ToString()
									$Enabled = $User.Enabled.ToString()
								}
								ElseIf($Admin.ObjectClass -eq 'group')
								{
									$UserName = "$($User.Name) (group)"
									$Domain = $xServer
									$PasswordLastSet = "N/A"
									$PasswordNeverExpires = "N/A"
									$Enabled = "N/A"
								}
								
							}
							Else
							{
								$UserName = $Admin.SID.Value
								$Domain = $xServer
								$PasswordLastSet = "Unknown"
								$PasswordNeverExpires = "Unknown"
								$Enabled = "Unknown"
							}
							$rowdata += @(,(
							$UserName,$htmlwhite,
							$Domain,$htmlwhite,
							$PasswordLastSet,$htmlwhite,
							$PasswordNeverExpires,$htmlwhite,
							$Enabled,$htmlwhite))
						}

						$columnHeaders = @(
						'Name',($htmlsilver -bor $htmlbold),
						'Domain',($htmlsilver -bor $htmlbold),
						'Password Last Changed',($htmlsilver -bor $htmlbold),
						'Password Never Expires',($htmlsilver -bor $htmlbold),
						'Account Enabled',($htmlsilver -bor $htmlbold)
						)
						
						$columnWidths = @("100","108","66","56","56")
						$msg = ""
						FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "386"
						WriteHTMLLine 0 0 " "
					}
				}
				ElseIf(!$?)
				{
					$txt1 = "Schema Admins: "
					$txt2 = "Unable to retrieve Schema Admins group membership"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 4 0 $txt1
						WriteWordLine 0 0 $txt2 "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 0 $txt1
						Line 0 $txt2
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 4 0 $txt1
						WriteHTMLLine 0 0 $txt2
					}
				}
				Else
				{
					$txt1 = "Schema Admins: "
					$txt2 = "<None>"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 4 0 $txt1
						WriteWordLine 0 0 $txt2
					}
					ElseIf($Text)
					{
						Line 0 $txt1
						Line 0 $txt2
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 4 0 $txt1
						WriteHTMLLine 0 0 "None"
					}
				}
			}

			#http://www.shariqsheikh.com/blog/index.php/200908/use-powershell-to-look-up-admincount-from-adminsdholder-and-sdprop/		
			Write-Verbose "$(Get-Date): `t`tListing users with AdminCount=1"
			#V2.20 changed to @()
			$AdminCounts = @(Get-ADUser -LDAPFilter "(admincount=1)" -Server $Domain -EA 0)
			
			If($? -and $Null -ne $AdminCounts)
			{
				$AdminCounts = $AdminCounts | Sort-Object Name
				[int]$AdminsCount = $AdminCounts.Count
				[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
				
				If($MSWORD -or $PDF)
				{
					WriteWordLine 4 0 "Users with AdminCount=1 ($AdminsCountStr users):"
					$TableRange = $Script:doc.Application.Selection.Range
					[int]$Columns = 4
					[int]$Rows = $AdminCounts.Count + 1
					[int]$xRow = 1
					$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.AutoFitBehavior($wdAutoFitFixed)
					$Table.Style = $Script:MyHash.Word_TableGrid
			
					$Table.rows.first.headingformat = $wdHeadingFormatTrue
					$Table.Borders.InsideLineStyle = $wdLineStyleSingle
					$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
					
					$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,1).Range.Font.Bold = $True
					$Table.Cell($xRow,1).Range.Text = "Name"
					$Table.Cell($xRow,2).Range.Font.Bold = $True
					$Table.Cell($xRow,2).Range.Text = "Password Last Changed"
					$Table.Cell($xRow,3).Range.Font.Bold = $True
					$Table.Cell($xRow,3).Range.Text = "Password Never Expires"
					$Table.Cell($xRow,4).Range.Font.Bold = $True
					$Table.Cell($xRow,4).Range.Text = "Account Enabled"
					ForEach($Admin in $AdminCounts)
					{
						$User = Get-ADUser -Identity $Admin.SID -Server $Domain `
						-Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0 

						$xRow++
						
						If($? -and $Null -ne $User)
						{
							$Table.Cell($xRow,1).Range.Text = $User.Name
							If($Null -eq $User.PasswordLastSet)
							{
								$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorRed
								$Table.Cell($xRow,2).Range.Font.Bold  = $True
								$Table.Cell($xRow,2).Range.Font.Color = $WDColorBlack
								$Table.Cell($xRow,2).Range.Text = "No Date Set"
							}
							Else
							{
								$Table.Cell($xRow,2).Range.Text = (get-date $User.PasswordLastSet -f d)
							}
							If($User.PasswordNeverExpires -eq $True)
							{
								$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorRed
								$Table.Cell($xRow,3).Range.Font.Bold  = $True
								$Table.Cell($xRow,3).Range.Font.Color = $WDColorBlack
							}
							$Table.Cell($xRow,3).Range.Text = $User.PasswordNeverExpires.ToString()
							If($User.Enabled -eq $False)
							{
								$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorRed
								$Table.Cell($xRow,4).Range.Font.Bold  = $True
								$Table.Cell($xRow,4).Range.Font.Color = $WDColorBlack
							}
							$Table.Cell($xRow,4).Range.Text = $User.Enabled.ToString()
						}
						Else
						{
							$Table.Cell($xRow,2).Range.Text = $Admin.SID
							$Table.Cell($xRow,3).Range.Text = "Unknown"
							$Table.Cell($xRow,4).Range.Text = "Unknown"
						}
					}
					
					#set column widths
					$xcols = $table.columns

					ForEach($xcol in $xcols)
					{
						switch ($xcol.Index)
						{
						  1 {$xcol.width = 200; Break}
						  2 {$xcol.width = 66; Break}
						  3 {$xcol.width = 56; Break}
						  4 {$xcol.width = 56; Break}
						}
					}

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
					$Table.AutoFitBehavior($wdAutoFitFixed)

					#return focus back to document
					$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

					#move to the end of the current document
					$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
					$TableRange = $Null
					$Table = $Null
				}
				ElseIf($Text)
				{
					Line 1 "Users with AdminCount=1 ($AdminsCountStr users):"
					#V2.16 addition
					Line 2 "                                                   Password   Password          "
					Line 2 "                                                   Last       Never      Account"
					Line 2 "Name                                               Changed    Expires    Enabled"
					Line 2 "================================================================================"
					ForEach($Admin in $AdminCounts)
					{
						$User = Get-ADUser -Identity $Admin.SID -Server $Domain `
						-Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0 

						If($? -and $Null -ne $User)
						{
							If($Null -eq $User.PasswordLastSet)
							{
								$PasswordLastSet = "No Date Set"
							}
							Else
							{
								$PasswordLastSet = (get-date $User.PasswordLastSet -f d)
							}
							If($User.PasswordNeverExpires -eq $True)
							{
								$PasswordNeverExpires = "True"
							}
							Else
							{
								$PasswordNeverExpires = "False"
							}
							If($User.Enabled -eq $False)
							{
								$UserEnabled = "True"
							}
							Else
							{
								$UserEnabled = "False"
							}
							#V2.16 change
							Line 2 ( "{0,-50} {1,-10} {2,-10} {3,-5}" -f $User.Name,$PasswordLastSet,$PasswordNeverExpires,$UserEnabled)
						}
						Else
						{
							#V2.16 change
							Line 2 ( "{0,-50} {1,-10} {2,-10} {3,-5}" -f $Admin.SID,"Unknown","Unknown","Unknown")
						}
					}
					Line 0 ""
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 4 0 "Users with AdminCount=1 ($($AdminsCountStr) users):"
					$rowdata = @()
					ForEach($Admin in $AdminCounts)
					{
						$User = Get-ADUser -Identity $Admin.SID -Server $Domain `
						-Properties PasswordLastSet, Enabled, PasswordNeverExpires -EA 0 

						If($? -and $Null -ne $User)
						{
							$UserName = $User.Name
							If($Null -eq $User.PasswordLastSet)
							{
								$PasswordLastSet = "No Date Set"
							}
							Else
							{
								$PasswordLastSet = (get-date $User.PasswordLastSet -f d)
							}
							If($User.PasswordNeverExpires -eq $True)
							{
								$PasswordNeverExpires = "True"
							}
							Else
							{
								$PasswordNeverExpires = "False"
							}
							If($User.Enabled -eq $True)
							{
								$Enabled = "True"
							}
							Else
							{
								$Enabled = "False"
							}
						}
						Else
						{
							$UserName = $Admin.SID
							$PasswordLastSet = "Unknown"
							$PasswordNeverExpires = "Unknown"
							$Enabled = "Unknown"
						}
						$rowdata += @(,(
						$UserName,$htmlwhite,
						$PasswordLastSet,$htmlwhite,
						$PasswordNeverExpires,$htmlwhite,
						$Enabled,$htmlwhite))
					}
					$columnHeaders = @(
					'Name',($htmlsilver -bor $htmlbold),
					'Password Last Changed',($htmlsilver -bor $htmlbold),
					'Password Never Expires',($htmlsilver -bor $htmlbold),
					'Account Enabled',($htmlsilver -bor $htmlbold)
					)
					
					$columnWidths = @("200","66","56","56")
					$msg = ""
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "378"
					WriteHTMLLine 0 0 " "
				}
			}
			ElseIf(!$?)
			{
				$txt = "Unable to retrieve users with AdminCount=1"
				If($MSWORD -or $PDF)
				{
					WriteWordLine 0 0 $txt "" $Null 0 $False $True
				}
				ElseIf($Text)
				{
					Line 0 $txt
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 0 0 $txt
				}
			}
			Else
			{
				$txt1 = "Users with AdminCount=1: "
				$txt2 = "<None>"
				If($MSWORD -or $PDF)
				{
					WriteWordLine 4 0 $txt1
					WriteWordLIne 0 0 $txt2
				}
				ElseIf($Text)
				{
					Line 0 $txt1
					Line 0 $txt2
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 4 0 $txt1
					WriteHTMLLIne 0 0 "None"
				}
			}
			
			Write-Verbose "$(Get-Date): `t`tListing groups with AdminCount = 1"
			#V2.20 changed to @()
			$AdminCounts = @(Get-ADGroup -LDAPFilter "(admincount=1)" -Server $Domain -EA 0 | Select-Object Name)
			
			If($? -and $Null -ne $AdminCounts)
			{
				$AdminCounts = $AdminCounts | Sort-Object Name
				[int]$AdminsCount = $AdminCounts.Count
				[string]$AdminsCountStr = "{0:N0}" -f $AdminsCount
				
				If($MSWORD -or $PDF)
				{
					WriteWordLine 4 0 "Groups with AdminCount=1 ($($AdminsCountStr) members):"
					$TableRange = $Script:doc.Application.Selection.Range
					[int]$Columns = 2
					[int]$Rows = $AdminCounts.Count + 1
					[int]$xRow = 1
					$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.AutoFitBehavior($wdAutoFitFixed)
					$Table.Style = $Script:MyHash.Word_TableGrid
			
					$Table.rows.first.headingformat = $wdHeadingFormatTrue
					$Table.Borders.InsideLineStyle = $wdLineStyleSingle
					$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
					$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,1).Range.Font.Bold = $True
					$Table.Cell($xRow,1).Range.Text = "Group Name"
					$Table.Cell($xRow,2).Range.Font.Bold = $True
					$Table.Cell($xRow,2).Range.Text = "Members"
					ForEach($Admin in $AdminCounts)
					{
						Write-Verbose "$(Get-Date): `t`t`t$($Admin.Name)"
						$xRow++
						#V2.20 changed to @()
						$Members = @(Get-ADGroupMember -Identity $Admin.Name -Server $Domain -EA 0 | Sort-Object Name)
						
						If($? -and $Null -ne $Members)
						{
							[int]$MembersCount = $Members.Count
						}
						Else
						{
							[int]$MembersCount = 0
						}

						[string]$MembersCountStr = "{0:N0}" -f $MembersCount
						$Table.Cell($xRow,1).Range.Text = "$($Admin.Name) ($($MembersCountStr) members)"
						$MbrStr = ""
						If($MembersCount -gt 0)
						{
							ForEach($Member in $Members)
							{
								$MbrStr += "$($Member.Name)`r"
							}
							$Table.Cell($xRow,2).Range.Text = $MbrStr
						}
					}
					
					#set column widths
					$xcols = $table.columns

					ForEach($xcol in $xcols)
					{
						switch ($xcol.Index)
						{
						  1 {$xcol.width = 200; Break}
						  2 {$xcol.width = 172; Break}
						}
					}
					
					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
					$Table.AutoFitBehavior($wdAutoFitFixed)

					#return focus back to document
					$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

					#move to the end of the current document
					$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
					$TableRange = $Null
					$Table = $Null
				}
				ElseIf($Text)
				{
					Line 1 "Groups with AdminCount=1 ($($AdminsCountStr) members):"
					ForEach($Admin in $AdminCounts)
					{
						Write-Verbose "$(Get-Date): `t`t`t$($Admin.Name)"
						#V2.20 changed to @()
						$Members = @(Get-ADGroupMember -Identity $Admin.Name -Server $Domain -EA 0 | Sort-Object Name)
						
						If($? -and $Null -ne $Members)
						{
							[int]$MembersCount = $Members.Count
						}
						Else
						{
							[int]$MembersCount = 0
						}

						[string]$MembersCountStr = "{0:N0}" -f $MembersCount
						Line 2 "Group Name`t: $($Admin.Name) ($($MembersCountStr) members)"
						$MbrStr = ""
						If($MembersCount -gt 0)
						{
							Line 2 "Members`t`t: " -NoNewLine
							$cnt = 0
							ForEach($Member in $Members)
							{
								$cnt++
								
								If($cnt -eq 1)
								{
									Line 0 $Member.Name
								}
								Else
								{
									Line 4 "  $($Member.Name)"
								}
							}
						}
						Else
						{
							Line 2 "Members`t`t: None"
						}
						Line 0 ""
					}
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 4 0 "Groups with AdminCount=1 ($($AdminsCountStr) members):"
					$rowdata = @()
					ForEach($Admin in $AdminCounts)
					{
						Write-Verbose "$(Get-Date): `t`t`t$($Admin.Name)"
						#V2.20 changed to @()
						$Members = @(Get-ADGroupMember -Identity $Admin.Name -Server $Domain -EA 0 | Sort-Object Name)
						
						If($? -and $Null -ne $Members)
						{
							$MembersCount = $Members.Count
						}
						Else
						{
							[int]$MembersCount = 0
						}

						[string]$MembersCountStr = "{0:N0}" -f $MembersCount
						$GroupName = "$($Admin.Name) ($($MembersCountStr) members)"
						$MbrStr = ""
						If($MembersCount -gt 0)
						{
							$cnt = 0
							ForEach($Member in $Members)
							{
								$cnt++
								
								If($cnt -eq 1)
								{
									$rowdata += @(,(
									$GroupName,$htmlwhite,
									$Member.Name,$htmlwhite))
								}
								Else
								{
									$rowdata += @(,(
									"",$htmlwhite,
									$Member.Name,$htmlwhite))
								}
							}
						}
						Else
						{
							$rowdata += @(,(
							$GroupName,$htmlwhite,
							"",$htmlwhite))
						}
					}
					$columnHeaders = @(
					'Group Name',($htmlsilver -bor $htmlbold),
					'Members',($htmlsilver -bor $htmlbold)
					)
					
					$columnWidths = @("200","172")
					$msg = ""
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "372"
					WriteHTMLLine 0 0 " "
				}
			}
			ElseIf(!$?)
			{
				$txt = "Unable to retrieve Groups with AdminCount=1"
				If($MSWORD -or $PDF)
				{
					WriteWordLine 0 0 $txt "" $Null 0 $False $True
				}
				ElseIf($Text)
				{
					Line 0 $txt
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 0 0 $txt
				}
			}
			Else
			{
				$txt1 = "Groups with AdminCount=1: "
				$txt2 = "<None>"
				If($MSWORD -or $PDF)
				{
					WriteWordLine 4 0 $txt1
					WriteWordLine 0 0 $txt2
				}
				ElseIf($Text)
				{
					Line 0 $txt1
					Line 0 $txt2
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 4 0 $txt1
					WriteHTMLLine 0 0 "None"
				}
			}
		}
		ElseIf(!$?)
		{
			$txt = "Error retrieving Group data for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 $txt
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}
		}
		Else
		{
			$txt = "No Group data was retrieved for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 $txt
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}
		}
		$First = $False
	}
}
#endregion

#region GPOs by domain
Function ProcessGPOsByDomain
{
	Write-Verbose "$(Get-Date): Writing domain group policy data"

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Group Policies by Domain"
	}
	ElseIf($Text)
	{
		Line 0 "///  Group Policies by Domain  \\\"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Group Policies by Domain&nbsp;&nbsp;\\\"
	}
	$First = $True

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date): `tProcessing group policies for domain $($Domain)"

		$DomainInfo = Get-ADDomain -Identity $Domain -EA 0 
		
		If($? -and $Null -ne $DomainInfo)
		{
			If(($MSWORD -or $PDF) -and !$First)
			{
				#put each domain, starting with the second, on a new page
				$Script:selection.InsertNewPage()
			}
			
			If($Domain -eq $Script:ForestRootDomain)
			{
				$txt = "$($Domain) (Forest Root)"
				If($MSWORD -or $PDF)
				{
					WriteWordLine 2 0 $txt
				}
				ElseIf($Text)
				{
					Line 1 "///  $($txt)  \\\"
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
				}
			}
			Else
			{
				If($MSWORD -or $PDF)
				{
					WriteWordLine 2 0 $Domain
				}
				ElseIf($Text)
				{
					Line 1 "///  $($Domain)  \\\"
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($Domain)&nbsp;&nbsp;\\\"
				}
			}

			Write-Verbose "$(Get-Date): `t`tGetting linked GPOs"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 3 0 "Linked Group Policy Objects" 
			}
			ElseIf($Text)
			{
				Line 0 "Linked Group Policy Objects" 
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 3 0 "Linked Group Policy Objects" 
			}

			#V2.20 changed to @()
			$LinkedGPOs = @($DomainInfo.LinkedGroupPolicyObjects | Sort-Object)
			If($Null -eq $LinkedGpos)
			{
				If($MSWORD -or $PDF)
				{
					WriteWordLine 0 0 "<None>"
				}
				ElseIf($Text)
				{
					Line 2 "<None>"
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 0 0 "None"
				}
			}
			Else
			{
				$GPOArray = New-Object System.Collections.ArrayList
				ForEach($LinkedGpo in $LinkedGpos)
				{
					#taken from Michael B. Smith's work on the XenApp 6.x scripts
					#this way we don't need the GroupPolicy module
					$gpObject = [ADSI]( "LDAP://" + $LinkedGPO )
					If($Null -eq $gpObject.DisplayName)
					{
						$p1 = $LinkedGPO.IndexOf("{")
						#38 is length of guid (32) plus the four "-"s plus the beginning "{" plus the ending "}"
						$GUID = $LinkedGPO.SubString($p1,38)
						$tmp = "GPO with GUID $($GUID) was not found in this domain"
					}
					Else
					{
						$tmp = $gpObject.DisplayName	### name of the group policy object
					}
					$GPOArray.Add($tmp) > $Null
				}

				$GPOArray = $GPOArray | Sort-Object 

				If($MSWORD -or $PDF)
				{
					$TableRange = $Script:doc.Application.Selection.Range
					[int]$Columns = 1
					[int]$Rows = $LinkedGpos.Count
					$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.Style = $Script:MyHash.Word_TableGrid
			
					$Table.Borders.InsideLineStyle = $wdLineStyleNone
					$Table.Borders.OutsideLineStyle = $wdLineStyleNone
					
					[int]$xRow = 0
					ForEach($Item in $GPOArray)
					{
						$xRow++
						$Table.Cell($xRow,1).Range.Text = $Item
					}

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
					$Table.AutoFitBehavior($wdAutoFitContent)

					#return focus back to document
					$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

					#move to the end of the current document
					$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
					$TableRange = $Null
					$Table = $Null
				}
				ElseIf($Text)
				{
					ForEach($Item in $GPOArray)
					{
						Line 2 $Item
					}
					Line 0 ""
				}
				ElseIf($HTML)
				{
					$rowdata = @()
					ForEach($Item in $GPOArray)
					{
						$rowdata += @(,($Item,$htmlwhite))
					}
					$columnHeaders = @('Name',($htmlsilver -bor $htmlbold))
					$msg = ""
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
					WriteHTMLLine 0 0 " "
				}
				$GPOArray = $Null
			}
			$LinkedGPOs = $Null
			$First = $False
		}
		ElseIf(!$?)
		{
			$txt = "Error retrieving domain data for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 $txt
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}
		}
		Else
		{
			$txt = "No Domain data was retrieved for domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 $txt
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}
		}
	}
}
#endregion

#region group policies by organizational units
Function ProcessgGPOsByOUOld
{
	Write-Verbose "$(Get-Date): Writing Group Policy data by Domain by OU"
	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Group Policies by Organizational Unit"
	}
	ElseIf($Text)
	{
		Line 0 "///  Group Policies by Organizational Unit  \\\"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Group Policies by Organizational Unit&nbsp;&nbsp;\\\"
	}
	$First = $True

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date): `tProcessing domain $($Domain)"
		If(($MSWORD -or $PDF) -and !$First)
		{
			#put each domain, starting with the second, on a new page
			$Script:selection.InsertNewPage()
		}

		$Disclaimer = "(Contains only OUs with linked Group Policies)"
		If($Domain -eq $Script:ForestRootDomain)
		{
			$txt = "Group Policies by OUs in Domain $($Domain) (Forest Root)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
				#print disclaimer line in 8 point bold italics
				WriteWordLine 0 0 $Disclaimer "" $Null 8 $True $True
			}
			ElseIf($Text)
			{
				Line 0 "///  $($txt)  \\\"
				Line 0 $Disclaimer
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
				WriteHTMLLine 0 0 $Disclaimer "" $Null 1 $True $True
			}
		}
		Else
		{
			$txt = "Group Policies by OUs in Domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
				WriteWordLine 0 0 $Disclaimer "" $Null 8 $True $True
			}
			ElseIf($Text)
			{
				Line 1 "///  $($txt)  \\\"
				Line 1 $Disclaimer
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
				WriteHTMLLine 0 0 $Disclaimer "" $Null 1 $True $True
			}
		}
		
		#get all OUs for the domain
		#V2.20 changed to @()
		$OUs = @(Get-ADOrganizationalUnit -Filter * -Server $Domain `
		-Properties CanonicalName, DistinguishedName, Name -EA 0 | `
		Select-Object CanonicalName, DistinguishedName, Name | Sort-Object CanonicalName)
		
		If($? -and $Null -ne $OUs)
		{
			[int]$NumOUs = $OUs.Count
			[int]$OUCount = 0

			ForEach($OU in $OUs)
			{
				$OUCount++
				$OUDisplayName = $OU.CanonicalName.SubString($OU.CanonicalName.IndexOf("/")+1)
				Write-Verbose "$(Get-Date): `t`tProcessing OU $($OU.CanonicalName) - OU # $OUCount of $NumOUs"
				
				#get data for the individual OU
				$OUInfo = Get-ADOrganizationalUnit -Identity $OU.DistinguishedName -Server $Domain -Properties * -EA 0 
				
				If($? -and $Null -ne $OUInfo)
				{
					Write-Verbose "$(Get-Date): `t`t`tGetting linked GPOs"
					$LinkedGPOs = $OUInfo.LinkedGroupPolicyObjects | Sort-Object 
					If($Null -eq $LinkedGpos)
					{
						# do nothing
					}
					Else
					{
						$GPOArray = New-Object System.Collections.ArrayList
						ForEach($LinkedGpo in $LinkedGpos)
						{
							#taken from Michael B. Smith's work on the XenApp 6.x scripts
							#this way we don't need the GroupPolicy module
							$gpObject = [ADSI]( "LDAP://" + $LinkedGPO )
							If($Null -eq $gpObject.DisplayName)
							{
								$p1 = $LinkedGPO.IndexOf("{")
								#38 is length of guid (32) plus the four "-"s plus the beginning "{" plus the ending "}"
								$GUID = $LinkedGPO.SubString($p1,38)
								$tmp = "GPO with GUID $($GUID) was not found in this domain"
							}
							Else
							{
								$tmp = $gpObject.DisplayName	### name of the group policy object
							}
							$GPOArray.Add($tmp) > $Null
						}

						$GPOArray = $GPOArray | Sort-Object 

						[int]$Rows = $LinkedGpos.Count

						If($MSWORD -or $PDF)
						{
							[int]$Columns = 1
							WriteWordLine 3 0 "$($OUDisplayName) ($($Rows))"
							$TableRange = $Script:doc.Application.Selection.Range
							$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
							$Table.Style = $Script:MyHash.Word_TableGrid
			
							$Table.Borders.InsideLineStyle = $wdLineStyleNone
							$Table.Borders.OutsideLineStyle = $wdLineStyleNone
							
							[int]$xRow = 0
							ForEach($Item in $GPOArray)
							{
								$xRow++
								$Table.Cell($xRow,1).Range.Text = $Item
							}

							$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
							$Table.AutoFitBehavior($wdAutoFitContent)

							#return focus back to document
							$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

							#move to the end of the current document
							$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
							$TableRange = $Null
							$Table = $Null
						}
						ElseIf($Text)
						{
							Line 2 "$($OUDisplayName) ($($Rows))"
							ForEach($Item in $GPOArray)
							{
								Line 3 $Item
							}
							Line 0 ""
						}
						ElseIf($HTML)
						{
							WriteHTMLLine 3 0 "$($OUDisplayName) ($($Rows))"
							$rowdata = @()
							ForEach($Item in $GPOArray)
							{
								$rowdata += @(,($Item,$htmlwhite))
							}
							$columnHeaders = @('Name',($htmlsilver -bor $htmlbold))
							$msg = ""
							FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
							WriteHTMLLine 0 0 " "
						}
						$GPOArray = $Null
					}
				}
				ElseIf(!$?)
				{
					Write-Warning "Error retrieving OU data for OU $($OU.CanonicalName)"
				}
				Else
				{
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 0 "<None>"
					}
					ElseIf($Text)
					{
						Line 0 "<None>"
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 0 "None"
					}
				}
			}
		}
		ElseIf(!$?)
		{
			$txt = "Error retrieving OU data for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 $txt
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}
		}
		Else
		{
			$txt = "No OU data was retrieved for domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 $txt
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}
		}
		$First = $False
	}
}

Function ProcessgGPOsByOUNew
{
	Write-Verbose "$(Get-Date): Writing Group Policy data by Domain by OU"
	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Group Policies by Organizational Unit"
	}
	ElseIf($Text)
	{
		Line 0 "///  Group Policies by Organizational Unit  \\\"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Group Policies by Organizational Unit&nbsp;&nbsp;\\\"
	}
	$First = $True

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date): `tProcessing domain $($Domain)"
		If(($MSWORD -or $PDF) -and !$First)
		{
			#put each domain, starting with the second, on a new page
			$Script:selection.InsertNewPage()
		}

		$Disclaimer = "(Contains only OUs with linked or inherited Group Policies)"
		If($Domain -eq $Script:ForestRootDomain)
		{
			$txt = "Group Policies by OUs in Domain $($Domain) (Forest Root)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
				#print disclaimer line in 8 point bold italics
				WriteWordLine 0 0 $Disclaimer "" $Null 8 $True $True
			}
			ElseIf($Text)
			{
				Line 0 "///  $($txt)  \\\"
				Line 0 $Disclaimer
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
				WriteHTMLLine 0 0 $Disclaimer "" $Null 1 $True $True
			}
		}
		Else
		{
			$txt = "Group Policies by OUs in Domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 2 0 $txt
				WriteWordLine 0 0 $Disclaimer "" $Null 8 $True $True
			}
			ElseIf($Text)
			{
				Line 1 "///  $($txt)  \\\"
				Line 1 $Disclaimer
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
				WriteHTMLLine 0 0 $Disclaimer "" $Null 1 $True $True
			}
		}
		
		#get all OUs for the domain
		#V2.20 changed to @()
		$OUs = @(Get-ADOrganizationalUnit -Filter * -Server $Domain `
		-Properties CanonicalName, DistinguishedName, Name -EA 0 | `
		Select-Object CanonicalName, DistinguishedName, Name | Sort-Object CanonicalName)
		
		If($? -and $Null -ne $OUs)
		{
			[int]$NumOUs = $OUs.Count
			[int]$OUCount = 0

			ForEach($OU in $OUs)
			{
				$OUCount++
				$OUDisplayName = $OU.CanonicalName.SubString($OU.CanonicalName.IndexOf("/")+1)
				Write-Verbose "$(Get-Date): `t`tProcessing OU $($OU.CanonicalName) - OU # $OUCount of $NumOUs"
				
				#get data for the individual OU
				$OUInfo = Get-ADOrganizationalUnit -Identity $OU.DistinguishedName -Server $Domain -Properties * -EA 0 
				
				If($? -and $Null -ne $OUInfo)
				{
					Write-Verbose "$(Get-Date): `t`t`tGetting linked and inherited GPOs"
					
					#change for 2.16
					#work around invalid property DisplayName when the gpolinks and inheritedgpolinks collections are empty
					
					$Results = Get-GPInheritance -target $OU.DistinguishedName -EA 0
					
					If(($Results.gpoLinks).Count -gt 0)
					{
						$LinkedGPOs = $Results.gpolinks.DisplayName
					}
					Else
					{
						$LinkedGPOs = $Null
					}
					
					If(($Results.inheritedgpoLinks).Count -gt 0)
					{
						$InheritedGPOs = $Results.inheritedgpolinks.DisplayName
					}
					Else
					{
						$InheritedGPOs = $Null
					}
					
					If($Null -eq $LinkedGPOs -and $Null -eq $InheritedGPOs)
					{
						# do nothing
					}
					Else
					{
						$AllGPOs = New-Object System.Collections.ArrayList

						ForEach($item in $InheritedGPOs)
						{
							$obj = New-Object -TypeName PSObject
							$GPOType = ""
							if(!($LinkedGPOs -contains $item))
							{
								$GPOType = "Inherited"
							}
							else
							{
								$GPOType = "Linked"
							}
							$obj | Add-Member -MemberType NoteProperty -Name "GPOName" -value $Item
							$obj | Add-Member -MemberType NoteProperty -Name "GPOType" -value $GPOType
							
							$AllGPOs.Add($obj) > $Null
						}

						$AllGPOS = $AllGPOs | Sort-Object GPOName						

						[int]$Rows = 0
						$Rows = $AllGPOS.Count
						
						If($MSWORD -or $PDF)
						{
							WriteWordLine 3 0 "$($OUDisplayName) ($($Rows))"
							[int]$Columns = 2
							$TableRange = $Script:doc.Application.Selection.Range
							#increment $Rows to account for the last row
							$Table = $Script:doc.Tables.Add($TableRange, ($Rows+1), $Columns)
							$Table.Style = $Script:MyHash.Word_TableGrid
			
							$Table.rows.first.headingformat = $wdHeadingFormatTrue
							$Table.Borders.InsideLineStyle = $wdLineStyleSingle
							$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
							
							$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
							[int]$xRow = 1
							$Table.Cell($xRow,1).Range.Font.Bold = $True
							$Table.Cell($xRow,1).Range.Text = "GPO Name"
							$Table.Cell($xRow,2).Range.Font.Bold = $True
							$Table.Cell($xRow,2).Range.Text = "GPO Type"
							ForEach($Item in $AllGPOS)
							{
								$xRow++
								$Table.Cell($xRow,1).Range.Text = $Item.GPOName
								$Table.Cell($xRow,2).Range.Text = $Item.GPOType
							}

							#set column widths
							$xcols = $table.columns

							ForEach($xcol in $xcols)
							{
								switch ($xcol.Index)
								{
								  1 {$xcol.width = 300; Break}
								  2 {$xcol.width = 65; Break}
								}
							}
							
							$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
							$Table.AutoFitBehavior($wdAutoFitFixed)

							#return focus back to document
							$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

							#move to the end of the current document
							$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
							$TableRange = $Null
							$Table = $Null
							WriteWordLine 0 0 ""
						}
						ElseIf($Text)
						{
							Line 2 "$($OUDisplayName) ($($Rows))"
							ForEach($Item in $AllGPOS)
							{
								Line 3 "Name: " $Item.GPOName
								Line 3 "Type: " $Item.GPOType
								Line 0 ""
							}
							Line 0 ""
						}
						ElseIf($HTML)
						{
							WriteHTMLLine 3 0 "$($OUDisplayName) ($($Rows))"
							$rowdata = @()
							ForEach($Item in $AllGPOS)
							{
								$rowdata += @(,(
								$Item.GPOName,$htmlwhite,
								$Item.GPOType,$htmlwhite))
							}
							$columnHeaders = @(
							'GPO Name',($htmlsilver -bor $htmlbold),
							'GPO Type',($htmlsilver -bor $htmlbold)
							)
							
							$columnWidths = @("350","65")
							$msg = ""
							FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "415"
							WriteHTMLLine 0 0 " "
						}
					}
				}
				ElseIf(!$?)
				{
					Write-Warning "Error retrieving OU data for OU $($OU.CanonicalName)"
				}
				Else
				{
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 0 "<None>"
					}
					ElseIf($Text)
					{
						Line 0 "<None>"
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 0 "None"
					}
				}
			}
		}
		ElseIf(!$?)
		{
			$txt = "Error retrieving OU data for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 $txt
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}
		}
		Else
		{
			$txt = "No OU data was retrieved for domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 $txt
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}
		}
		$First = $False
	}
}
#endregion

#region misc info by domain
#From Jeff Hicks
#modified from his original
#https://www.petri.com/powershell-problem-solver-active-directory-remote-desktop-settings
#added for 2.16
Function Get-RDUserSetting 
{
	[cmdletbinding(DefaultParameterSetName="SAM")]
	 
	Param(
	[Parameter(Position=0,Mandatory,HelpMessage="Enter a user's sAMAccountName",
	ValueFromPipeline,ParameterSetName="SAM")]
	[ValidateNotNullorEmpty()]
	[Alias("Name")]
	[string]$SAMAccountname,
	[Parameter(ParameterSetName="SAM")]
	[string]$SearchRoot,
	 
	[Parameter(Mandatory,HelpMessage="Enter a user's distingished name",
	ValueFromPipelineByPropertyName,ParameterSetName="DN")]
	[ValidateNotNullorEmpty()]
	[Alias("DN")]
	[string]$DistinguishedName,
	 
	[string]$Server
	 
	)
	 
	Begin 
	{
		#remote desktop properties
		$TSSettings = @("TerminalServicesProfilePath","TerminalServicesHomeDirectory","TerminalServicesHomeDrive")
	}
	 
	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			"SAM" 
			{
				$searcher = New-Object DirectoryServices.DirectorySearcher
				$searcher.Filter = "(&(objectcategory=person)(objectclass=user)(samAccountname=$sAMAccountname))"
				If($SearchRoot) 
				{
					If($Server) 
					{
						$searchPath = "LDAP://$server/$SearchRoot"
					}
					Else 
					{
						$searchPath = "LDAP://$SearchRoot"
					}
					$r = New-Object System.DirectoryServices.DirectoryEntry $SearchPath

					$searcher.SearchRoot = $r
				}
				$user = $searcher.FindOne().GetDirectoryEntry()
			} 
			"DN" 
			{
				If($server) 
				{
					[ADSI]$User = "LDAP://$Server/$DistinguishedName"
				}
				Else 
				{
					[ADSI]$User = "LDAP://$DistinguishedName"
				}
			}
		} #close Switch
	 
		If($user.path) 
		{
			#initialize a hashtable
			Try 
			{
				$hash=[ordered]@{
					DistinguishedName = $User.DistinguishedName.Value
					Name = $user.name.Value
					samAccountName = $user.samAccountName.value
					AllowLogon = $user.psbase.InvokeGet("AllowLogon") -as [Boolean]
				}

				ForEach($property in $TSSettings) 
				{
					$hash.Add($property,$user.psbase.invokeGet($property))
				} #ForEach

				#create an object
				New-Object -TypeName PSObject -Property $hash
			}
			Catch 
			{
				#nothing
			}
		} #if user found
		Else 
		{
			#nothing
		}
	 
	} #Process
	 
	End 
	{
		#nothing
	} #End
 
} #end function

Function ProcessMiscDataByDomain
{
	Write-Verbose "$(Get-Date): Writing miscellaneous data by domain"

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Miscellaneous Data by Domain"
	}
	ElseIf($Text)
	{
		Line 0 "///  Miscellaneous Data by Domain  \\\"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Miscellaneous Data by Domain&nbsp;&nbsp;\\\"
	}
	
	$First = $True

	ForEach($Domain in $Script:Domains)
	{
		Write-Verbose "$(Get-Date): `tProcessing misc data for domain $($Domain)"

		$DomainInfo = Get-ADDomain -Identity $Domain -EA 0 
		
		If($? -and $Null -ne $DomainInfo)
		{
			If(($MSWORD -or $PDF) -and !$First)
			{
				#put each domain, starting with the second, on a new page
				$Script:selection.InsertNewPage()
			}
			
			If($Domain -eq $Script:ForestRootDomain)
			{
				$txt = "$($Domain) (Forest Root)"
				If($MSWORD -or $PDF)
				{
					WriteWordLine 2 0 $txt
				}
				ElseIf($Text)
				{
					Line 0 "///  $($txt)  \\\"
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($txt)&nbsp;&nbsp;\\\"
				}
			}
			Else
			{
				If($MSWORD -or $PDF)
				{
					WriteWordLine 2 0 $Domain
				}
				ElseIf($Text)
				{
					Line 0 "///  $($Domain)  \\\"
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 2 0 "///&nbsp;&nbsp;$($Domain)&nbsp;&nbsp;\\\"
				}
			}

			Write-Verbose "$(Get-Date): `t`tGathering user misc data"
			
			#added for 2.16 HomeDrive, HomeDirectory, ProfilePath, ScriptPath, PrimaryGroup
			#V2.20 changed to @()
			$Users = @(Get-ADUser -Filter * -Server $Domain `
			-Properties CannotChangePassword, Enabled, LockedOut, PasswordExpired, PasswordNeverExpires, `
			PasswordNotRequired, lastLogonTimestamp, DistinguishedName, SamAccountName, UserPrincipalName, `
			HomeDrive, HomeDirectory, ProfilePath, ScriptPath, PrimaryGroup -EA 0)
			
			If($? -and $Null -ne $Users)
			{
				[int]$UsersCount = $Users.Count
				
				#added in V2.22
				If($UsersCount -gt 100000)
				{
					Write-Verbose "$(Get-Date): `t`t`t******************************************************************************************************"
					Write-Verbose "$(Get-Date): `t`t`tThere are $($UsersCount) user accounts to process. The following 17 actions will take a long time. Be patient."
					Write-Verbose "$(Get-Date): `t`t`t******************************************************************************************************"
				}
				
				Write-Verbose "$(Get-Date): `t`t`tDisabled users"
				#V2.20 changed to @()
				$DisabledUsers = @($Users | Where-Object {$_.Enabled -eq $False})
			
				[int]$UsersDisabledcnt = $DisabledUsers.Count
				
				Write-Verbose "$(Get-Date): `t`t`tUnknown users"
				#V2.20 changed to @()
				$UnknownUsers = @($Users | Where-Object {$Null -eq $_.Enabled})
			
				[int]$UsersUnknowncnt = $UnknownUsers.Count

				Write-Verbose "$(Get-Date): `t`t`tLocked out users"
				#V2.20 changed to @()
				$LockedOutUsers = @($Users | Where-Object {$_.LockedOut -eq $True})
			
				[int]$UsersLockedOutcnt = $LockedOutUsers.Count

				Write-Verbose "$(Get-Date): `t`t`tAll users with password expired"
				#V2.20 changed to @()
				$AllUsersWithPasswordExpired = @($Users | Where-Object {$_.PasswordExpired -eq $True})
			
				[int]$UsersPasswordExpiredcnt = $AllUsersWithPasswordExpired.Count

				Write-Verbose "$(Get-Date): `t`t`tAll users whose password never expires"
				#V2.20 changed to @()
				$AllUsersWhosePasswordNeverExpires = @($Users | Where-Object {$_.PasswordNeverExpires -eq $True})
			
				[int]$UsersPasswordNeverExpirescnt = $AllUsersWhosePasswordNeverExpires.Count

				Write-Verbose "$(Get-Date): `t`t`tAll users with password not required"
				#V2.20 changed to @()
				$AllUsersWithPasswordNotRequired = @($Users | Where-Object {$_.PasswordNotRequired -eq $True})
			
				[int]$UsersPasswordNotRequiredcnt = $AllUsersWithPasswordNotRequired.Count

				Write-Verbose "$(Get-Date): `t`t`tAll users who cannot change password"
				#V2.20 changed to @()
				$AllUsersWhoCannotChangePassword = @($Users | Where-Object {$_.CannotChangePassword -eq $True})
			
				[int]$UsersCannotChangePasswordcnt = $AllUsersWhoCannotChangePassword.Count

				Write-Verbose "$(Get-Date): `t`t`tAll users with SID History"
				#V2.20 changed to @()
				$AllUsersWithSIDHistory = @(Get-ADObject -LDAPFilter "(sIDHistory=*)" -Server $Domain `
				-Property objectClass, sIDHistory -EA 0)

				[int]$UsersWithSIDHistorycnt = ($AllUsersWithSIDHistory | Where-Object {$_.objectClass -eq 'user'}).Count

				#2.16
				Write-Verbose "$(Get-Date): `t`t`tAll users with Homedrive set in ADUC"
				#V2.20 changed to @()
				$HomeDriveUsers = @($Users | Where-Object {$Null -ne $_.HomeDrive}) #fixed in 2.24
			
				[int]$UsersHomeDrivecnt = $HomeDriveUsers.Count
				
				#2.16
				Write-Verbose "$(Get-Date): `t`t`tAll users whose Primary Group is not Domain Users"
				#V2.20 changed to @()
				$PrimaryGroupUsers = @($Users | Where-Object {$_.SamAccountName -ne 'Guest' -and $_.PrimaryGroup -notmatch 'Domain Users'})
			
				[int]$UsersPrimaryGroupcnt = $PrimaryGroupUsers.Count

				#2.16
				Write-Verbose "$(Get-Date): `t`t`tAll users with RDS HomeDrive set in ADUC"
				#V2.20 changed to @()
				$RDSHomeDriveUsers = @($users | Get-RDUserSetting | Where-Object {$_.TerminalServicesHomeDrive -gt 0})
			
				[int]$UsersRDSHomeDrivecnt = $RDSHomeDriveUsers.Count

				#active users now
				Write-Verbose "$(Get-Date): `t`t`tActive users"
				#V2.20 changed to @()
				$EnabledUsers = @($Users | Where-Object {$_.Enabled -eq $True})
			
				[int]$ActiveUsersCount = $EnabledUsers.Count

				Write-Verbose "$(Get-Date): `t`t`tActive users password expired"
				#V2.20 changed to @()
				$Results = @($EnabledUsers | Where-Object {$_.PasswordExpired -eq $True})
			
				[int]$ActiveUsersPasswordExpired = $Results.Count

				Write-Verbose "$(Get-Date): `t`t`tActive users password never expires"
				#V2.20 changed to @()
				$Results = @($EnabledUsers | Where-Object {$_.PasswordNeverExpires -eq $True})
			
				[int]$ActiveUsersPasswordNeverExpires = $Results.Count

				Write-Verbose "$(Get-Date): `t`t`tActive users password not required"
				#V2.20 changed to @()
				$Results = @($EnabledUsers | Where-Object {$_.PasswordNotRequired -eq $True})
			
				[int]$ActiveUsersPasswordNotRequired = $Results.Count

				Write-Verbose "$(Get-Date): `t`t`tActive Users cannot change password"
				#V2.20 changed to @()
				$Results = @($EnabledUsers | Where-Object {$_.CannotChangePassword -eq $True})
			
				[int]$ActiveUsersCannotChangePassword = $Results.Count

				Write-Verbose "$(Get-Date): `t`t`tActive Users no lastLogonTimestamp"
				#V2.20 changed to @()
				$Results = @($EnabledUsers | Where-Object {$Null -eq $_.lastLogonTimestamp})
			
				[int]$ActiveUserslastLogonTimestamp = $Results.Count
			}
			Else
			{
				[int]$UsersCount                      = 0
				[int]$UsersDisabledcnt                = 0
				[int]$UsersLockedOutcnt               = 0
				[int]$UsersPasswordExpiredcnt         = 0
				[int]$UsersPasswordNeverExpirescnt    = 0
				[int]$UsersPasswordNotRequiredcnt     = 0
				[int]$UsersCannotChangePasswordcnt    = 0
				[int]$UsersWithSIDHistorycnt          = 0
				[int]$UsersHomeDrivecnt               = 0
				[int]$UsersPrimaryGroupcnt            = 0
				[int]$UsersRDSHomeDrivecnt            = 0
				[int]$ActiveUsersCount                = 0 #fixed 2.24
				[int]$ActiveUsersPasswordExpired      = 0 #fixed 2.24
				[int]$ActiveUsersPasswordNeverExpires = 0 #fixed 2.24
				[int]$ActiveUsersPasswordNotRequired  = 0 #fixed 2.24
				[int]$ActiveUsersCannotChangePassword = 0 #fixed 2.24
				[int]$ActiveUserslastLogonTimestamp   = 0 #fixed 2.24
			}

			Write-Verbose "$(Get-Date): `t`tFormat numbers into strings"
			[string]$UsersCountStr                      = "{0,7:N0}" -f $UsersCount
			[string]$UsersDisabledStr                   = "{0,7:N0}" -f $UsersDisabledcnt
			[string]$UsersUnknownStr                    = "{0,7:N0}" -f $UsersUnknowncnt
			[string]$UsersLockedOutStr                  = "{0,7:N0}" -f $UsersLockedOutcnt
			[string]$UsersPasswordExpiredStr            = "{0,7:N0}" -f $UsersPasswordExpiredcnt
			[string]$UsersPasswordNeverExpiresStr       = "{0,7:N0}" -f $UsersPasswordNeverExpirescnt
			[string]$UsersPasswordNotRequiredStr        = "{0,7:N0}" -f $UsersPasswordNotRequiredcnt
			[string]$UsersCannotChangePasswordStr       = "{0,7:N0}" -f $UsersCannotChangePasswordcnt
			[string]$UsersWithSIDHistoryStr             = "{0,7:N0}" -f $UsersWithSIDHistorycnt
			[string]$UsersHomeDriveStr                  = "{0,7:N0}" -f $UsersHomeDrivecnt
			[string]$UsersPrimaryGroupStr               = "{0,7:N0}" -f $UsersPrimaryGroupcnt
			[string]$UsersRDSHomeDriveStr               = "{0,7:N0}" -f $UsersRDSHomeDrivecnt
			[string]$ActiveUsersCountStr                = "{0,7:N0}" -f $ActiveUsersCount
			[string]$ActiveUsersPasswordExpiredStr      = "{0,7:N0}" -f $ActiveUsersPasswordExpired
			[string]$ActiveUsersPasswordNeverExpiresStr = "{0,7:N0}" -f $ActiveUsersPasswordNeverExpires
			[string]$ActiveUsersPasswordNotRequiredStr  = "{0,7:N0}" -f $ActiveUsersPasswordNotRequired
			[string]$ActiveUsersCannotChangePasswordStr = "{0,7:N0}" -f $ActiveUsersCannotChangePassword
			[string]$ActiveUserslastLogonTimestampStr   = "{0,7:N0}" -f $ActiveUserslastLogonTimestamp

			If($MSWORD -or $PDF)
			{
				Write-Verbose "$(Get-Date): `t`tBuild table for All Users"
				WriteWordLine 3 0 "All Users"
				$TableRange   = $Script:doc.Application.Selection.Range
				[int]$Columns = 3
				[int]$Rows = 12
				$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
				$Table.Style = $Script:MyHash.Word_TableGrid
			
				$Table.Borders.InsideLineStyle = $wdLineStyleSingle
				$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
				$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(1,1).Range.Font.Bold = $True
				$Table.Cell(1,1).Range.Text = "Total Users"
				$Table.Cell(1,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(1,2).Range.Text = $UsersCountStr
				$Table.Cell(2,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(2,1).Range.Font.Bold = $True
				$Table.Cell(2,1).Range.Text = "Disabled users"
				$Table.Cell(2,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(2,2).Range.Text = $UsersDisabledStr
				[single]$pct = (($UsersDisabledcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(2,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(2,3).Range.Text = "$($pctstr)% of Total Users"
				$Table.Cell(3,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(3,1).Range.Font.Bold = $True
				$Table.Cell(3,1).Range.Text = "Unknown users*"
				$Table.Cell(3,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(3,2).Range.Text = $UsersUnknownStr
				[single]$pct = (($UsersUnknowncnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(3,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(3,3).Range.Text = "$($pctstr)% of Total Users"
				$Table.Cell(4,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(4,1).Range.Font.Bold = $True
				$Table.Cell(4,1).Range.Text = "Locked out users"
				$Table.Cell(4,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(4,2).Range.Text = $UsersLockedOutStr
				[single]$pct = (($UsersLockedOutcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(4,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(4,3).Range.Text = "$($pctstr)% of Total Users"
				$Table.Cell(5,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(5,1).Range.Font.Bold = $True
				$Table.Cell(5,1).Range.Text = "Password expired"
				$Table.Cell(5,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(5,2).Range.Text = $UsersPasswordExpiredStr
				[single]$pct = (($UsersPasswordExpiredcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(5,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(5,3).Range.Text = "$($pctstr)% of Total Users"
				$Table.Cell(6,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(6,1).Range.Font.Bold = $True
				$Table.Cell(6,1).Range.Text = "Password never expires"
				$Table.Cell(6,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(6,2).Range.Text = $UsersPasswordNeverExpiresStr
				[single]$pct = (($UsersPasswordNeverExpirescnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(6,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(6,3).Range.Text = "$($pctstr)% of Total Users"
				$Table.Cell(7,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(7,1).Range.Font.Bold = $True
				$Table.Cell(7,1).Range.Text = "Password not required"
				$Table.Cell(7,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(7,2).Range.Text = $UsersPasswordNotRequiredStr
				[single]$pct = (($UsersPasswordNotRequiredcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(7,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(7,3).Range.Text = "$($pctstr)% of Total Users"
				$Table.Cell(8,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(8,1).Range.Font.Bold = $True
				$Table.Cell(8,1).Range.Text = "Can't change password"
				$Table.Cell(8,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(8,2).Range.Text = $UsersCannotChangePasswordStr
				[single]$pct = (($UsersCannotChangePasswordcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(8,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(8,3).Range.Text = "$($pctstr)% of Total Users"
				$Table.Cell(9,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(9,1).Range.Font.Bold = $True
				$Table.Cell(9,1).Range.Text = "With SID History"
				$Table.Cell(9,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(9,2).Range.Text = $UsersWithSIDHistoryStr
				[single]$pct = (($UsersWithSIDHistorycnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(9,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(9,3).Range.Text = "$($pctstr)% of Total Users"
				$Table.Cell(10,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(10,1).Range.Font.Bold = $True
				$Table.Cell(10,1).Range.Text = "HomeDrive users"
				$Table.Cell(10,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(10,2).Range.Text = $UsersHomeDriveStr
				[single]$pct = (($UsersHomeDrivecnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(10,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(10,3).Range.Text = "$($pctstr)% of Total Users"
				$Table.Cell(11,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(11,1).Range.Font.Bold = $True
				$Table.Cell(11,1).Range.Text = "PrimaryGroup users"
				$Table.Cell(11,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(11,2).Range.Text = $UsersPrimaryGroupStr
				[single]$pct = (($UsersPrimaryGroupcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(11,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(11,3).Range.Text = "$($pctstr)% of Total Users"
				$Table.Cell(12,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(12,1).Range.Font.Bold = $True
				$Table.Cell(12,1).Range.Text = "RDS HomeDrive users"
				$Table.Cell(12,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(12,2).Range.Text = $UsersRDSHomeDriveStr
				[single]$pct = (($UsersRDSHomeDrivecnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(12,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(12,3).Range.Text = "$($pctstr)% of Total Users"

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
				$Table.AutoFitBehavior($wdAutoFitContent)

				#return focus back to document
				$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

				#move to the end of the current document
				$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
				$TableRange = $Null
				$Table = $Null

				WriteWordLine 0 0 "*Unknown users are user accounts with no Enabled property." "" $Null 8 $False $True
				If($Script:DARights -eq $False)
				{
					WriteWordLine 0 0 "*Rerun the script with Domain Admin rights in $($ADForest)." "" $Null 8 $False $True
				}
				Else
				{
					WriteWordLine 0 0 "*This may be a permissions issue if this is a Trusted Forest." "" $Null 8 $False $True
				}
				
				Write-Verbose "$(Get-Date): `t`tBuild table for Active Users"
				WriteWordLine 3 0 "Active Users"
				$TableRange   = $Script:doc.Application.Selection.Range
				[int]$Columns = 3
				[int]$Rows = 6
				$Table = $Script:doc.Tables.Add($TableRange, $Rows, $Columns)
				$Table.Style = $Script:MyHash.Word_TableGrid
			
				$Table.Borders.InsideLineStyle = $wdLineStyleSingle
				$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
				$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(1,1).Range.Font.Bold = $True
				$Table.Cell(1,1).Range.Text = "Total Active Users"
				$Table.Cell(1,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(1,2).Range.Text = $ActiveUsersCountStr
				$Table.Cell(2,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(2,1).Range.Font.Bold = $True
				$Table.Cell(2,1).Range.Text = "Password expired"
				$Table.Cell(2,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(2,2).Range.Text = $ActiveUsersPasswordExpiredStr
				[single]$pct = (($ActiveUsersPasswordExpired / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(2,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(2,3).Range.Text = "$($pctstr)% of Active Users"
				$Table.Cell(3,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(3,1).Range.Font.Bold = $True
				$Table.Cell(3,1).Range.Text = "Password never expires"
				$Table.Cell(3,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(3,2).Range.Text = $ActiveUsersPasswordNeverExpiresStr
				[single]$pct = (($ActiveUsersPasswordNeverExpires / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(3,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(3,3).Range.Text = "$($pctstr)% of Active Users"
				$Table.Cell(4,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(4,1).Range.Font.Bold = $True
				$Table.Cell(4,1).Range.Text = "Password not required"
				$Table.Cell(4,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(4,2).Range.Text = $ActiveUsersPasswordNotRequiredStr
				[single]$pct = (($ActiveUsersPasswordNotRequired / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(4,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(4,3).Range.Text = "$($pctstr)% of Active Users"
				$Table.Cell(5,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(5,1).Range.Font.Bold = $True
				$Table.Cell(5,1).Range.Text = "Can't change password"
				$Table.Cell(5,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(5,2).Range.Text = $ActiveUsersCannotChangePasswordStr
				[single]$pct = (($ActiveUsersCannotChangePassword / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(5,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(5,3).Range.Text = "$($pctstr)% of Active Users"
				$Table.Cell(6,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell(6,1).Range.Font.Bold = $True
				$Table.Cell(6,1).Range.Text = "No lastLogonTimestamp"
				$Table.Cell(6,2).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(6,2).Range.Text = $ActiveUserslastLogonTimestampStr
				[single]$pct = (($ActiveUserslastLogonTimestamp / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$Table.Cell(6,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
				$Table.Cell(6,3).Range.Text = "$($pctstr)% of Active Users"

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
				$Table.AutoFitBehavior($wdAutoFitContent)

				#return focus back to document
				$Script:doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

				#move to the end of the current document
				$Script:selection.EndKey($wdStory,$wdMove) | Out-Null
				$TableRange = $Null
				$Table = $Null

				#put computer info on a separate page
				$Script:selection.InsertNewPage()
			}
			ElseIf($Text)
			{
				Write-Verbose "$(Get-Date): `t`tBuild table for All Users"
				Line 0 "All Users"
				Line 1 "Total Users`t`t: " $UsersCountStr

				[single]$pct = (($UsersDisabledcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "Disabled users`t`t: $($UsersDisabledStr)`t$($pctstr)% of Total Users"

				[single]$pct = (($UsersUnknowncnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "Unknown users*`t`t: $($UsersUnknownStr)`t$($pctstr)% of Total Users"

				[single]$pct = (($UsersLockedOutcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "Locked out users`t: $($UsersLockedOutStr)`t$($pctstr)% of Total Users"

				[single]$pct = (($UsersPasswordExpiredcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "Password expired`t: $($UsersPasswordExpiredStr)`t$($pctstr)% of Total Users"

				[single]$pct = (($UsersPasswordNeverExpirescnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "Password never expires`t: $($UsersPasswordNeverExpiresStr)`t$($pctstr)% of Total Users"

				[single]$pct = (($UsersPasswordNotRequiredcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "Password not required`t: $($UsersPasswordNotRequiredStr)`t$($pctstr)% of Total Users"

				[single]$pct = (($UsersCannotChangePasswordcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "Can't change password`t: $($UsersCannotChangePasswordStr)`t$($pctstr)% of Total Users"

				[single]$pct = (($UsersWithSIDHistorycnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "With SID History`t: $($UsersWithSIDHistoryStr)`t$($pctstr)% of Total Users"

				[single]$pct = (($UsersHomeDrivecnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "With HomeDrive`t`t: $($UsersHomeDriveStr)`t$($pctstr)% of Total Users"

				[single]$pct = (($UsersPrimaryGroupcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "With Primary Group`t: $($UsersPrimaryGroupStr)`t$($pctstr)% of Total Users"

				[single]$pct = (($UsersRDSHomeDrivecnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "With RDS HomeDrive`t: $($UsersRDSHomeDriveStr)`t$($pctstr)% of Total Users"

				Line 1 "*Unknown users are user accounts with no Enabled property"
				If($Script:DARights -eq $False)
				{
					Line 1 "*Rerun the script with Domain Admin rights in $($ADForest)"
				}
				Else
				{
					Line 1 "*This may be a permissions issue if this is a Trusted Forest"
				}
				Line 0 ""
				
				Write-Verbose "$(Get-Date): `t`tBuild table for Active Users"
				Line 0 "Active Users"

				Line 1 "Total Active Users`t: " $ActiveUsersCountStr

				[single]$pct = (($ActiveUsersPasswordExpired / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "Password expired`t: $($ActiveUsersPasswordExpiredStr)`t$($pctstr)% of Active Users"

				[single]$pct = (($ActiveUsersPasswordNeverExpires / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "Password never expires`t: $($ActiveUsersPasswordNeverExpiresStr)`t$($pctstr)% of Active Users"

				[single]$pct = (($ActiveUsersPasswordNotRequired / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "Password not required`t: $($ActiveUsersPasswordNotRequiredStr)`t$($pctstr)% of Active Users"

				[single]$pct = (($ActiveUsersCannotChangePassword / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "Can't change password`t: $($ActiveUsersCannotChangePasswordStr)`t$($pctstr)% of Active Users"

				[single]$pct = (($ActiveUserslastLogonTimestamp / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				Line 1 "No lastLogonTimestamp`t: $($ActiveUserslastLogonTimestampStr)`t$($pctstr)% of Active Users"
				Line 0 ""
			}
			ElseIf($HTML)
			{
				Write-Verbose "$(Get-Date): `t`tBuild table for All Users"
				WriteHTMLLine 3 0 "All Users"
				$rowdata = @()
				$columnHeaders = @("Total Users",($htmlsilver -bor $htmlbold),
									$UsersCountStr,($htmlwhite),
									"",($htmlwhite))

				[single]$pct = (($UsersDisabledcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,('Disabled users',($htmlsilver -bor $htmlbold),
				$UsersDisabledStr,($htmlwhite),
				"$($pctstr)% of Total Users",($htmlwhite)))

				[single]$pct = (($UsersUnknowncnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,('Unknown users*',($htmlsilver -bor $htmlbold),
				$UsersUnknownStr,($htmlwhite),
				"$($pctstr)% of Total Users",($htmlwhite)))

				[single]$pct = (($UsersLockedOutcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,('Locked out users',($htmlsilver -bor $htmlbold),
				$UsersLockedOutStr,($htmlwhite),
				"$($pctstr)% of Total Users",($htmlwhite)))

				[single]$pct = (($UsersPasswordExpiredcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,('Password expired',($htmlsilver -bor $htmlbold),
				$UsersPasswordExpiredStr,($htmlwhite),
				"$($pctstr)% of Total User",($htmlwhite)))

				[single]$pct = (($UsersPasswordNeverExpirescnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,('Password never expires',($htmlsilver -bor $htmlbold),
				$UsersPasswordNeverExpiresStr,($htmlwhite),
				"$($pctstr)% of Total Users",($htmlwhite)))

				[single]$pct = (($UsersPasswordNotRequiredcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,('Password not required',($htmlsilver -bor $htmlbold),
				$UsersPasswordNotRequiredStr,($htmlwhite),
				"$($pctstr)% of Total Users",($htmlwhite)))

				[single]$pct = (($UsersCannotChangePasswordcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,("Can't change password",($htmlsilver -bor $htmlbold),
				$UsersCannotChangePasswordStr,($htmlwhite),
				"$($pctstr)% of Total Users",($htmlwhite)))

				[single]$pct = (($UsersWithSIDHistorycnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,('With SID History',($htmlsilver -bor $htmlbold),
				$UsersWithSIDHistoryStr,($htmlwhite),
				"$($pctstr)% of Total Users",($htmlwhite)))
				
				[single]$pct = (($UsersHomeDrivecnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,('With HomeDrive',($htmlsilver -bor $htmlbold),
				$UsersHomeDriveStr,($htmlwhite),
				"$($pctstr)% of Total Users",($htmlwhite)))
				
				[single]$pct = (($UsersPrimaryGroupcnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,('With PrimaryGroup',($htmlsilver -bor $htmlbold),
				$UsersPrimaryGroupStr,($htmlwhite),
				"$($pctstr)% of Total Users",($htmlwhite)))
				
				[single]$pct = (($UsersRDSHomeDrivecnt / $UsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,('With RDS HomeDrive',($htmlsilver -bor $htmlbold),
				$UsersRDSHomeDriveStr,($htmlwhite),
				"$($pctstr)% of Total Users",($htmlwhite)))
				
				$msg = ""
				$columnWidths = @("150","50","150")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"

				WriteHTMLLine 0 0 "*Unknown users are user accounts with no Enabled property." "" $Null 1 $False $True
				If($Script:DARights -eq $False)
				{
					WriteHTMLLine 0 0 "*Rerun the script with Domain Admin rights in $($ADForest)." "" $Null 1 $False $True
				}
				Else
				{
					WriteHTMLLine 0 0 "*This may be a permissions issue if this is a Trusted Forest." "" $Null 1 $False $True
				}
				
				Write-Verbose "$(Get-Date): `t`tBuild table for Active Users"
				WriteHTMLLine 3 0 "Active Users"

				$rowdata = @()
				$rowdata += @(,('Total Active Users',($htmlsilver -bor $htmlbold),
				$ActiveUsersCountStr,($htmlwhite),
				"",($htmlwhite)))

				[single]$pct = (($ActiveUsersPasswordExpired / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,('Password expired',($htmlsilver -bor $htmlbold),
				$ActiveUsersPasswordExpiredStr,($htmlwhite),
				"$($pctstr)% of Active Users",($htmlwhite)))

				[single]$pct = (($ActiveUsersPasswordNeverExpires / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,('Password never expires',($htmlsilver -bor $htmlbold),
				$ActiveUsersPasswordNeverExpiresStr,($htmlwhite),
				"$($pctstr)% of Active Users",($htmlwhite)))

				[single]$pct = (($ActiveUsersPasswordNotRequired / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,('Password not required',($htmlsilver -bor $htmlbold),
				$ActiveUsersPasswordNotRequiredStr,($htmlwhite),
				"$($pctstr)% of Active Users",($htmlwhite)))

				[single]$pct = (($ActiveUsersCannotChangePassword / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,("Can't change password",($htmlsilver -bor $htmlbold),
				$ActiveUsersCannotChangePasswordStr,($htmlwhite),
				"$($pctstr)% of Active Users",($htmlwhite)))

				[single]$pct = (($ActiveUserslastLogonTimestamp / $ActiveUsersCount)*100)
				$pctstr = "{0,5:N2}" -f $pct
				$rowdata += @(,('No lastLogonTimestamp',($htmlsilver -bor $htmlbold),
				$ActiveUserslastLogonTimestampStr,($htmlwhite),
				"$($pctstr)% of Active Users",($htmlwhite)))

				$msg = ""
				$columnWidths = @("150","50","150")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
				WriteHTMLLine 0 0 " "
			}
			
			If($UsersDisabledcnt -gt 0 -and $IncludeUserInfo -eq $True)
			{
				OutputUserInfo $DisabledUsers "Disabled users"
			}

			If($UsersUnknowncnt -gt 0 -and $IncludeUserInfo -eq $True)
			{
				OutputUserInfo $UnknownUsers "Unknown users"
			}

			If($UsersLockedOutcnt -gt 0 -and $IncludeUserInfo -eq $True)
			{
				OutputUserInfo $LockedOutUsers "Locked out users"
			}

			If($UsersPasswordExpiredcnt -gt 0 -and $IncludeUserInfo -eq $True)
			{
				OutputUserInfo $AllUsersWithPasswordExpired "All users with password expired"
			}

			If($UsersPasswordNeverExpirescnt -gt 0 -and $IncludeUserInfo -eq $True)
			{
				OutputUserInfo $AllUsersWhosePasswordNeverExpires "All users whose password never expires"
			}

			If($UsersPasswordNotRequiredcnt -gt 0 -and $IncludeUserInfo -eq $True)
			{
				OutputUserInfo $AllUsersWithPasswordNotRequired "All users with password not required"
			}

			If($UsersCannotChangePasswordcnt -gt 0 -and $IncludeUserInfo -eq $True)
			{
				OutputUserInfo $AllUsersWhoCannotChangePassword "All users who cannot change password"
			}

			If($UsersWithSIDHistorycnt -gt 0 -and $IncludeUserInfo -eq $True)
			{
				OutputUserInfo $AllUsersWithSIDHistory "All users with SID History"
			}

			If($UsersHomeDrivecnt -gt 0 -and $IncludeUserInfo -eq $True)
			{
				OutputHDUserInfo $HomeDriveUsers "All users with HomeDrive set in ADUC"
			}

			If($UsersPrimaryGroupcnt -gt 0 -and $IncludeUserInfo -eq $True)
			{
				OutputPGUserInfo $PrimaryGroupUsers "All users whose Primary Group is not Domain Users"
			}

			If($UsersRDSHomeDrivecnt -gt 0 -and $IncludeUserInfo -eq $True)
			{
				OutputRDSHDUserInfo $RDSHomeDriveUsers "All users with RDS HomeDrive set in ADUC"
			}

			Get-ComputerCountByOS $Domain
		}
		ElseIf(!$?)
		{
			$txt = "Error retrieving domain data for domain $($Domain)"
			Write-Warning $txt
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 $txt
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}
		}
		Else
		{
			$txt = "No Domain data was retrieved for domain $($Domain)"
			If($MSWORD -or $PDF)
			{
				WriteWordLine 0 0 $txt "" $Null 0 $False $True
			}
			ElseIf($Text)
			{
				Line 0 $txt
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $txt
			}
		}
		$First = $False
	}
	$Script:Domains = $Null
}

Function OutputUserInfo
{
	Param([object] $Users, [string] $title)
	
	Write-Verbose "$(Get-Date): `t`t`t`tOutput $($title)"
	$Users = $Users | Sort-Object samAccountName
	
	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $UsersWordTable = @();
		[int] $CurrentServiceIndex = 2;

		WriteWordLine 4 0 $title

		ForEach($User in $Users)
		{
			$WordTableRowHash = @{ 
			SamAccountName = $User.SamAccountName; 
			DN = $User.DistinguishedName
			}

			## Add the hash to the array
			$UsersWordTable += $WordTableRowHash;

			$CurrentServiceIndex++;
		}
		
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $UsersWordTable `
		-Columns SamAccountName, DN `
		-Headers "SamAccountName", "DistinguishedName" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 350;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 $title
		Line 0 ""
		Line 1 "samAccountName            DistinguishedName"
		Line 1 "=============================================================================================================================================="
		###### "1234512345123451234512345

		ForEach($User in $Users)
		{
			Line 1 ( "{0,-25} {1,-116}" -f $User.samAccountName, $User.DistinguishedName )
		}
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 $title
		$rowdata = @()
		
		ForEach($User in $Users)
		{
			$rowdata += @(,($User.SamAccountName,$htmlwhite,
							$User.DistinguishedName,$htmlwhite))
		}
		
		$columnHeaders = @('SamAccountName',($htmlsilver -bor $htmlbold),'DistinguishedName',($htmlsilver -bor $htmlbold))
		$columnWidths = @("150px","350px")
		$msg = ""
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputHDUserInfo
{
	#new for 2.16
	Param([object] $Users, [string] $title)
	
	Write-Verbose "$(Get-Date): `t`t`t`tOutput $($title)"
	$Users = $Users | Sort-Object samAccountName
	
	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $UsersWordTable = @();
		[int] $CurrentServiceIndex = 2;

		WriteWordLine 4 0 $title

		ForEach($User in $Users)
		{
			$WordTableRowHash = @{ 
			SamAccountName = $User.SamAccountName; 
			DN = $User.DistinguishedName;
			HomeDrive = $User.HomeDrive;
			HomeDir = $User.HomeDirectory;
			ProfilePath = $User.ProfilePath;
			ScriptPath = $User.ScriptPath
			}

			## Add the hash to the array
			$UsersWordTable += $WordTableRowHash;

			$CurrentServiceIndex++;
		}
		
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $UsersWordTable `
		-Columns SamAccountName, DN, HomeDrive, HomeDir, ProfilePath, ScriptPath `
		-Headers "SamAccountName", "DistinguishedName", "Home drive", "Home folder", "Profile path", "Login script" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table -Size 9
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 100;
		$Table.Columns.Item(2).Width = 110;
		$Table.Columns.Item(3).Width = 35;
		$Table.Columns.Item(4).Width = 85;
		$Table.Columns.Item(5).Width = 85;
		$Table.Columns.Item(6).Width = 85;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 $title
		Line 0 ""

		ForEach($User in $Users)
		{
			Line 1 "SamAccountName`t`t: " $User.samAccountName
			Line 1 "DistinguishedName`t: " $User.DistinguishedName
			Line 1 "Home drive`t`t: " $User.HomeDrive
			Line 1 "Home folder`t`t: " $User.HomeDirectory
			Line 1 "Profile path`t`t: " $User.ProfilePath
			Line 1 "Login script`t`t: " $User.ScriptPath
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 $title
		$rowdata = @()
		
		ForEach($User in $Users)
		{
			$rowdata += @(,($User.SamAccountName,$htmlwhite,
							$User.DistinguishedName,$htmlwhite,
							$User.HomeDrive,$htmlwhite,
							$User.HomeDirectory,$htmlwhite,
							$User.ProfilePath,$htmlwhite,
							$User.ScriptPath,$htmlwhite))
		}
		
		$columnHeaders = @(
		'SamAccountName',($htmlsilver -bor $htmlbold),
		'DistinguishedName',($htmlsilver -bor $htmlbold),
		'Home Drive',($htmlsilver -bor $htmlbold),
		'Home folder',($htmlsilver -bor $htmlbold),
		'Profile path',($htmlsilver -bor $htmlbold),
		'Login script',($htmlsilver -bor $htmlbold))
		$columnWidths = @("100px","100px","75px","75px","75px","75px")
		$msg = ""
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputPGUserInfo
{
	#new for 2.16
	Param([object] $Users, [string] $title)
	
	Write-Verbose "$(Get-Date): `t`t`t`tOutput $($title)"
	$Users = $Users | Sort-Object samAccountName
	
	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $UsersWordTable = @();
		[int] $CurrentServiceIndex = 2;

		WriteWordLine 4 0 $title

		ForEach($User in $Users)
		{
			$WordTableRowHash = @{ 
			SamAccountName = $User.SamAccountName; 
			DN = $User.DistinguishedName;
			PG = $User.PrimaryGroup
			}

			## Add the hash to the array
			$UsersWordTable += $WordTableRowHash;

			$CurrentServiceIndex++;
		}
		
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $UsersWordTable `
		-Columns SamAccountName, DN, PG `
		-Headers "SamAccountName", "DistinguishedName", "Primary Group" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table -Size 9
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 100;
		$Table.Columns.Item(2).Width = 200;
		$Table.Columns.Item(3).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 $title
		Line 0 ""

		ForEach($User in $Users)
		{
			Line 1 "samAccountName`t`t: " $User.samAccountName
			Line 1 "DistinguishedName`t: " $User.DistinguishedName
			Line 1 "Primary Group`t`t: " $User.PrimaryGroup
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 $title
		$rowdata = @()
		
		ForEach($User in $Users)
		{
			$rowdata += @(,($User.SamAccountName,$htmlwhite,
							$User.DistinguishedName,$htmlwhite,
							$User.PrimaryGroup,$htmlwhite))
		}
		
		$columnHeaders = @(
		'SamAccountName',($htmlsilver -bor $htmlbold),
		'DistinguishedName',($htmlsilver -bor $htmlbold),
		'Primary Group',($htmlsilver -bor $htmlbold))
		$columnWidths = @("100px","200px","200px")
		$msg = ""
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
		WriteHTMLLine 0 0 " "
	}
}

Function OutputRDSHDUserInfo
{
	#new for 2.16
	Param([object] $Users, [string] $title)
	
	Write-Verbose "$(Get-Date): `t`t`t`tOutput $($title)"
	$Users = $Users | Sort-Object samAccountName
	
	If($MSWORD -or $PDF)
	{
		[System.Collections.Hashtable[]] $UsersWordTable = @();
		[int] $CurrentServiceIndex = 2;

		WriteWordLine 4 0 $title

		ForEach($User in $Users)
		{
			$WordTableRowHash = @{ 
			SamAccountName = $User.SamAccountName; 
			DN = $User.DistinguishedName;
			HomeDrive = $User.TerminalServicesHomeDrive;
			HomeDir = $User.TerminalServicesHomeDirectory;
			ProfilePath = $User.TerminalServicesProfilePath;
			AllowLogon = $User.AllowLogon
			}

			## Add the hash to the array
			$UsersWordTable += $WordTableRowHash;

			$CurrentServiceIndex++;
		}
		
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $UsersWordTable `
		-Columns SamAccountName, DN, HOmeDrive, HomeDir, ProfilePath, ALlowLogon `
		-Headers "SamAccountName", "DistinguishedName", "RDS Home drive", "RDS Home folder", "RDS Profile path", "Allow Logon" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table -Size 9
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 100;
		$Table.Columns.Item(2).Width = 140;
		$Table.Columns.Item(3).Width = 35;
		$Table.Columns.Item(4).Width = 75;
		$Table.Columns.Item(5).Width = 90;
		$Table.Columns.Item(6).Width = 60;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 $title
		Line 0 ""

		ForEach($User in $Users)
		{
			Line 1 "SamAccountName`t`t: " $User.samAccountName
			Line 1 "DistinguishedName`t: " $User.DistinguishedName
			Line 1 "RDS Home drive`t`t: " $User.TerminalServicesHomeDrive
			Line 1 "RDS Home folder`t`t: " $User.TerminalServicesHomeDirectory
			Line 1 "RDS Profile path`t: " $User.TerminalServicesProfilePath
			Line 1 "Allow Logon`t`t: " $User.AllowLogon
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 $title
		$rowdata = @()
		
		ForEach($User in $Users)
		{
			$rowdata += @(,($User.SamAccountName,$htmlwhite,
							$User.DistinguishedName,$htmlwhite,
							$User.TerminalServicesHomeDrive,$htmlwhite,
							$User.TerminalServicesHomeDirectory,$htmlwhite,
							$User.TerminalServicesProfilePath,$htmlwhite,
							$User.AllowLogon,$htmlwhite))
		}
		
		$columnHeaders = @(
		'SamAccountName',($htmlsilver -bor $htmlbold),
		'DistinguishedName',($htmlsilver -bor $htmlbold),
		'RDS Home Drive',($htmlsilver -bor $htmlbold),
		'RDS Home folder',($htmlsilver -bor $htmlbold),
		'RDS Profile path',($htmlsilver -bor $htmlbold),
		'Allow Logon',($htmlsilver -bor $htmlbold))
		$columnWidths = @("100px","100px","75px","75px","90px","60px")
		$msg = ""
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region DCDNSInfo
Function ProcessDCDNSInfo
{
	If($DCDNSInfo)
	{
		#Domain Controller DNS IP Configuration
		Write-Verbose "$(Get-Date): Create Domain Controller DNS IP Configuration"
		Write-Verbose "$(Get-Date): `tAdd Domain Controller DNS IP Configuration table to doc"
		
		#sort by site then by DC
		$xDCDNSIPInfo = $Script:DCDNSIPInfo | Sort-Object DCSite, DCName
		
		If($MSWord -or $PDF)
		{
			$Script:selection.InsertNewPage()
			WriteWordLine 1 0 "Domain Controller DNS IP Configuration"
			[System.Collections.Hashtable[]] $ItemsWordTable = @();
			[int] $CurrentServiceIndex = 2;
		}
		ElseIf($Text)
		{
			Line 0 "///  Domain Controller DNS IP Configuration  \\\"
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 1 0 "///&nbsp;&nbsp;Domain Controller DNS IP Configuration&nbsp;&nbsp;\\\"
			$rowdata = @()
		}

		ForEach($Item in $xDCDNSIPInfo)
		{
			If($MSWord -or $PDF)
			{
				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ 
				DCName = $Item.DCName;
				DCSite = $Item.DCSite;
				DCIpAddress1 = $Item.DCIpAddress1;
				DCIpAddress2 = $Item.DCIpAddress2;
				DCDNS1 = $Item.DCDNS1; 
				DCDNS2 = $Item.DCDNS2; 
				DCDNS3 = $Item.DCDNS3; 
				DCDNS4 = $Item.DCDNS4
				}

				## Add the hash to the array
				$ItemsWordTable += $WordTableRowHash;

				$CurrentServiceIndex++;
			}
			ElseIf($Text)
			{
				Line 1 "DC Name`t`t: " $Item.DCName
				Line 1 "Site Name`t: " $Item.DCSite
				Line 1 "IP Address1`t: " $Item.DCIpAddress1
				Line 1 "IP Address2`t: " $Item.DCIpAddress2
				Line 1 "Preferred DNS`t: " $Item.DCDNS1
				Line 1 "Alternate DNS`t: " $Item.DCDNS2
				Line 1 "DNS 3`t`t: " $Item.DCDNS3
				Line 1 "DNS 4`t`t: " $Item.DCDNS4
				Line 0 ""
			}
			ElseIf($HTML)
			{
				$rowdata += @(,(
				$Item.DCName,$htmlwhite,
				$Item.DCSite,$htmlwhite,
				$Item.DCIpAddress1,$htmlwhite,
				$Item.DCIpAddress2,$htmlwhite,
				$Item.DCDNS1,$htmlwhite,
				$Item.DCDNS2,$htmlwhite,
				$Item.DCDNS3,$htmlwhite,
				$Item.DCDNS4,$htmlwhite))
			}
		}

		If($MSWord -or $PDF)
		{
			## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns DCName, DCSite, DCIpAddress1, DCIpAddress2, DCDNS1, DCDNS2, DCDNS3, DCDNS4 `
			-Headers "DC Name", "Site", "IP Address 1", "IP Address 2", "DNS 1", "DNS 2", "DNS 3", "DNS 4" `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
			SetWordCellFormat -Collection $Table -Size 8

			$Table.Columns.Item(1).Width = 100;
			$Table.Columns.Item(2).Width = 60;
			$Table.Columns.Item(3).Width = 67;
			$Table.Columns.Item(4).Width = 67;
			$Table.Columns.Item(5).Width = 40;
			$Table.Columns.Item(6).Width = 40;
			$Table.Columns.Item(7).Width = 40;
			$Table.Columns.Item(8).Width = 40;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		ElseIf($Text)
		{
			#nothing to do
		}
		ElseIf($HTML)
		{
			$columnHeaders = @(
			'DC Name',($htmlsilver -bor $htmlbold),
			'Site',($htmlsilver -bor $htmlbold),
			'IP Address 1',($htmlsilver -bor $htmlbold),
			'IP Address 2',($htmlsilver -bor $htmlbold),
			'Preferred DNS',($htmlsilver -bor $htmlbold),
			'Alternate DNS',($htmlsilver -bor $htmlbold),
			'DNS 3',($htmlsilver -bor $htmlbold),
			'DNS 4',($htmlsilver -bor $htmlbold)
			)

			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 " "
		}

		Write-Verbose "$(Get-Date): Finished Create Domain Controller DNS IP Configuration"
		Write-Verbose "$(Get-Date): "
	}
}
#endregion

#region TimeServerInfo
Function ProcessTimeServerInfo
{
	#Domain Controller Time Server Configuration
	Write-Verbose "$(Get-Date): Create Domain Controller Time Server Configuration"
	Write-Verbose "$(Get-Date): `tAdd Domain Controller Time Server Configuration table to doc"
	
	#sort by DC
	$xTimeServerInfo = $Script:TimeServerInfo | Sort-Object DCName
	
	If($MSWord -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Domain Controller Time Server Configuration"
		[System.Collections.Hashtable[]] $ItemsWordTable = @();
		[int] $CurrentServiceIndex = 2;
	}
	ElseIf($Text)
	{
		Line 0 "///  Domain Controller Time Server Configuration  \\\"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Domain Controller Time Server Configuration&nbsp;&nbsp;\\\"
		$rowdata = @()
	}

	ForEach($Item in $xTimeServerInfo)
	{
		If($MSWord -or $PDF)
		{
			## Add the required key/values to the hashtable
			$WordTableRowHash = @{ 
				DCName = $Item.DCName;
				DCTimeSource = $Item.TimeSource;
				DCAnnounceFlags = $Item.AnnounceFlags;
				DCMaxNegPhaseCorrection = $Item.MaxNegPhaseCorrection;
				DCMaxPosPhaseCorrection = $Item.MaxPosPhaseCorrection;
				DCNtpServer = $Item.NtpServer;
				DCNtpType = $Item.NtpType;
				DCSpecialPollInterval = $Item.SpecialPollInterval;
				DCVMICTimeProvider = $Item.VMICTimeProvider
			}

			## Add the hash to the array
			$ItemsWordTable += $WordTableRowHash;

			$CurrentServiceIndex++;
		}
		ElseIf($Text)
		{
			Line 1 "DC Name`t`t`t: " $Item.DCName
			Line 1 "Time source`t`t: " $Item.TimeSource
			Line 1 "Announce flags`t`t: " $Item.AnnounceFlags
			Line 1 "Max Neg Phase Correction: " $Item.MaxNegPhaseCorrection
			Line 1 "Max Pos Phase Correction: " $Item.MaxPosPhaseCorrection
			Line 1 "NTP Server`t`t: " $Item.NtpServer
			Line 1 "Type`t`t`t: " $Item.NtpType
			Line 1 "Special Poll Interval`t: " $Item.SpecialPollInterval
			Line 1 "VMIC Time Provider`t: " $Item.VMICTimeProvider
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
				$Item.DCName,$htmlwhite,
				$Item.TimeSource,$htmlwhite,
				$Item.AnnounceFlags,$htmlwhite,
				$Item.MaxNegPhaseCorrection,$htmlwhite,
				$Item.MaxPosPhaseCorrection,$htmlwhite,
				$Item.NtpServer,$htmlwhite,
				$Item.NtpType,$htmlwhite,
				$Item.SpecialPollInterval,$htmlwhite,
				$Item.VMICTimeProvider,$htmlwhite
			))
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $ItemsWordTable `
		-Columns DCName, DCTimeSource, DCAnnounceFlags, DCMaxNegPhaseCorrection, DCMaxPosPhaseCorrection, DCNtpServer, DCNtpType, DCSpecialPollInterval, DCVMICTimeProvider `
		-Headers "DC Name", "Time Source", "Announce Flags", "Max Neg Phase Correction", "Max Pos Phase Correction", "NTP Server", "Type", "Special Poll Interval", "VMIC Time Provider" `
		-AutoFit $wdAutoFitFixed;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
		SetWordCellFormat -Collection $Table -Size 8

		$Table.Columns.Item(1).Width = 105;
		$Table.Columns.Item(2).Width = 73;
		$Table.Columns.Item(3).Width = 45;
		$Table.Columns.Item(4).Width = 47;
		$Table.Columns.Item(5).Width = 47;
		$Table.Columns.Item(6).Width = 73;
		$Table.Columns.Item(7).Width = 33;
		$Table.Columns.Item(8).Width = 37;
		$Table.Columns.Item(9).Width = 40;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	ElseIf($Text)
	{
		#nothing to do
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'DC Name',($htmlsilver -bor $htmlbold),
		'Time Source',($htmlsilver -bor $htmlbold),
		'Announce Flags',($htmlsilver -bor $htmlbold),
		'Max Neg Phase Correction',($htmlsilver -bor $htmlbold),
		'Max Pos Phase Correction',($htmlsilver -bor $htmlbold),
		'NTP Server',($htmlsilver -bor $htmlbold),
		'Type',($htmlsilver -bor $htmlbold),
		'Special Poll Interval',($htmlsilver -bor $htmlbold),
		'VMIC Time Provider',($htmlsilver -bor $htmlbold)
		)

		$msg = ""
		$columnWidths = @("100px","70px","45px","45px","45px","75px","40px","40px","40px")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
		WriteHTMLLine 0 0 " "
	}

	Write-Verbose "$(Get-Date): Finished Create Domain Controller Time Server Configuration"
	Write-Verbose "$(Get-Date): "
}
#endregion

#region EventLogInfo
Function ProcessEventLogInfo
{
	#Domain Controller Event Log Data
	Write-Verbose "$(Get-Date): Create Domain Controller Event Log Data"
	Write-Verbose "$(Get-Date): `tAdd Domain Controller Event Log Data table to doc"
	
	#sort by DC and then event log name
	#V2.20 changed to @()
	$xEventLogInfo = @($Script:DCEventLogInfo | Sort-Object EventLogName, DCName)
	
	If($MSWord -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Domain Controller Event Log Data"
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 3
		[int]$Rows = $xEventLogInfo.Count + 1
		[int]$xRow = 1
		
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$Table.AutoFitBehavior($wdAutoFitFixed)
		$Table.Style = $Script:MyHash.Word_TableGrid
	
		$Table.rows.first.headingformat = $wdHeadingFormatTrue
		$Table.Borders.InsideLineStyle = $wdLineStyleSingle
		$Table.Borders.OutsideLineStyle = $wdLineStyleSingle

		$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell($xRow,1).Range.Font.Bold = $True
		$Table.Cell($xRow,1).Range.Text = "Event Log Name"
		
		$Table.Cell($xRow,2).Range.Font.Bold = $True
		$Table.Cell($xRow,2).Range.Text = "DC Name"
		
		$Table.Cell($xRow,3).Range.Font.Bold = $True
		$Table.Cell($xRow,3).Range.Text = "Event Log Size (KB)"
	}
	ElseIf($Text)
	{
		Line 0 "///  Domain Controller Event Log Data  \\\"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Domain Controller Event Log Data&nbsp;&nbsp;\\\"
		$rowdata = @()
	}

	ForEach($Item in $xEventLogInfo)
	{
		If($MSWord -or $PDF)
		{
			$xRow++
			$Table.Cell($xRow,1).Range.Text = $Item.EventLogName
			$Table.Cell($xRow,2).Range.Text = $Item.DCName
			$Table.Cell($xRow,3).Range.ParagraphFormat.Alignment = $wdCellAlignVerticalTop
			$Table.Cell($xRow,3).Range.ParagraphFormat.Alignment = $wdAlignParagraphRight
			$Table.Cell($xRow,3).Range.Text = $Item.EventLogSize
		}
		ElseIf($Text)
		{
			Line 1 "Event Log Name`t`t: " $Item.EventLogName
			Line 1 "DC Name`t`t`t: " $Item.DCName
			Line 1 "Event Log Size (KB)`t: " $Item.EventLogSize
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
				$Item.EventLogName,$htmlwhite,
				$Item.DCName,$htmlwhite,
				$Item.EventLogSize,$htmlwhite
			))
		}
	}

	If($MSWord -or $PDF)
	{
		#set column widths
		$xcols = $table.columns

		ForEach($xcol in $xcols)
		{
			switch ($xcol.Index)
			{
			  1 {$xcol.width = 150; Break}
			  2 {$xcol.width = 150; Break}
			  3 {$xcol.width = 100; Break}
			}
		}
		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)
		$Table.AutoFitBehavior($wdAutoFitFixed)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		#nothing to do
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Event Log Name',($htmlsilver -bor $htmlbold),
		'DC Name',($htmlsilver -bor $htmlbold),
		'Event Log Size (KB)',($htmlsilver -bor $htmlbold)
		)

		$msg = ""
		$columnWidths = @("150px","150px","100px")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "400"
		WriteHTMLLine 0 0 " "
	}

	Write-Verbose "$(Get-Date): Finished Create Domain Controller Event Log Data"
	Write-Verbose "$(Get-Date): "
}
#endregion

#region general script functions
Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	ElseIf($Text)
	{
		SaveandCloseTextDocument
	}
	ElseIf($HTML)
	{
		SaveandCloseHTMLDocument
	}

	$GotFile = $False

	If($PDF)
	{
		If(Test-Path "$($Script:FileName2)")
		{
			Write-Verbose "$(Get-Date): $($Script:FileName2) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName2)"
			Write-Error "Unable to save the output file, $($Script:FileName2)"
		}
	}
	Else
	{
		If(Test-Path "$($Script:FileName1)")
		{
			Write-Verbose "$(Get-Date): $($Script:FileName1) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName1)"
			Write-Error "Unable to save the output file, $($Script:FileName1)"
		}
	}
	
	#email output file if requested
	If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		If($PDF)
		{
			$emailAttachment = $Script:FileName2
		}
		Else
		{
			$emailAttachment = $Script:FileName1
		}
		SendEmail $emailAttachment
	}
}
#endregion

#region script start function
Function ProcessScriptStart
{
	$script:startTime = Get-Date
}
#endregion

#region script end
Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date): Script has completed"
	Write-Verbose "$(Get-Date): "

	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
	Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date): Elapsed time: $($Str)"

	If($Dev)
	{
		If($SmtpServer -eq "")
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error 4>$Null
		}
		Else
		{
			Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
		}
	}

	If($ScriptInfo)
	{
		$SIFile = "$($Script:pwdpath)\ADInventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime   : $AddDateTime" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name   : $Script:CoName" 4>$Null		
		}
		Out-File -FilePath $SIFile -Append -InputObject "ComputerName   : $ComputerName" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Address: $CompanyAddress" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Email  : $CompanyEmail" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Fax    : $CompanyFax" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Phone  : $CompanyPhone" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page     : $CoverPage" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "DCDNSInfo      : $DCDNSInfo" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Dev            : $Dev" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile   : $Script:DevErrorFile" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Domain name    : $ADDomain" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elevated       : $Script:Elevated" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Filename1      : $Script:FileName1" 4>$Null
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Filename2      : $Script:FileName2" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder         : $Folder" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Forest name    : $ADForest" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From           : $From" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "GPOInheritance : $GPOInheritance" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "HW Inventory   : $Hardware" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "IncludeUserInfo: $IncludeUserInfo" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log            : $($Log)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "MaxDetails     : $MaxDetails" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML   : $HTML" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF    : $PDF" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT   : $TEXT" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD   : $MSWORD" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info    : $ScriptInfo" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Services       : $Services" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port      : $SmtpPort" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server    : $SmtpServer" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title          : $Script:Title" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To             : $To" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL        : $UseSSL" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "User Name      : $UserName" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected    : $Script:RunningOS" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version   : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture      : $PSCulture" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture    : $PSUICulture" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word language  : $Script:WordLanguageValue" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word version   : $Script:WordProduct" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start   : $Script:StartTime" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time   : $Str" 4>$Null
	}

	#V2.18 added
	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $true) 
		{
			try 
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date): $Script:LogPath is ready for use"
			} 
			catch 
			{
				Write-Verbose "$(Get-Date): Transcript/log stop failed"
			}
		}
	}
	$ErrorActionPreference = $SaveEAPreference
}
#endregion

#region script core
#Script begins

ProcessScriptStart

ProcessScriptSetup

If($ADDomain -ne "")
{
	SetFilename1andFilename2 "$Script:DomainDNSRoot"
}
Else
{
	SetFilename1andFilename2 "$Script:ForestRootDomain"
}

If($Section -eq "All" -or $Section -eq "Forest")
{
	ProcessForestInformation

	ProcessAllDCsInTheForest
	
	ProcessCAInformation
	
	ProcessADOptionalFeatures
	
	ProcessADSchemaItems
	[gc]::collect()
}

If($Section -eq "All" -or $Section -eq "Sites")
{
	ProcessSiteInformation
	[gc]::collect()
}

If($Section -eq "All" -or $Section -eq "Domains")
{
	ProcessDomains
	ProcessDomainControllers
	[gc]::collect()
}

If($Section -eq "All" -or $Section -eq "OUs")
{
	ProcessOrganizationalUnits
	[gc]::collect()
}

If($Section -eq "All" -or $Section -eq "Groups")
{
	ProcessGroupInformation
	[gc]::collect()
}

If($Section -eq "All" -or $Section -eq "GPOs")
{
	ProcessGPOsByDomain

	If($GPOInheritance -eq $True)
	{
		ProcessgGPOsByOUNew
	}
	Else
	{
		ProcessgGPOsByOUOld
	}
	[gc]::collect()
}

If($Section -eq "All" -or $Section -eq "Misc")
{
	ProcessMiscDataByDomain
	[gc]::collect()
}

If($Section -eq "All" -or $Section -eq "Domains")
{
	ProcessDCDNSInfo
	[gc]::collect()
}

If($Script:Elevated -and ($Section -eq "All" -or $Section -eq "Domains"))
{
	ProcessTimeServerInfo
	[gc]::collect()
}

If($Script:Elevated -and ($Section -eq "All" -or $Section -eq "Domains"))
{
	ProcessEventLogInfo
	[gc]::collect()
}

#endregion

#region finish script
Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

$AbstractTitle = "Microsoft Active Directory Inventory Report V2.16"
$SubjectTitle = "Active Directory Inventory Report V2.16"
UpdateDocumentProperties $AbstractTitle $SubjectTitle

ProcessDocumentOutput

ProcessScriptEnd
[gc]::collect()

#endregion
