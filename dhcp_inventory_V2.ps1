#requires -Version 4.0
#requires -Module DHCPServer
#This File is in Unicode format.  Do not edit in an ASCII editor. Notepad++ UTF-8-BOM

#region help text

<#
.SYNOPSIS
	Creates a complete inventory of a Microsoft 2012+ DHCP server.
.DESCRIPTION
	Creates a complete inventory of a Microsoft 2012+ DHCP server using Microsoft 
	PowerShell, Word, plain text, or HTML.
	
	Creates a Word or PDF document, text or HTML file named either:
		DHCP Inventory Report for Server <DHCPServerName> for the Domain <domain>.HTML 
		(or .DOCX or .PDF or .TXT).
		DHCP Inventory Report for All DHCP Servers for the Domain <domain>.HTML (or .DOCX 
		or .PDF or .TXT).

	Version 2.0 changes the default output report from Word to HTML.
	
	The script requires at least PowerShell version 4 but runs best in version 5.
	
	Word is NOT needed to run the script. This script outputs in Text and HTML.

	You do NOT have to run this script on a DHCP server. This script was developed 
	and run from a Windows 10 VM.

	Requires the DHCPServer module.
	
	The script can run on a DHCP server or a Windows 8.x or Windows 10 computer with RSAT 
	installed.
		
	Remote Server Administration Tools for Windows 8 
		https://carlwebster.sharefile.com/d-s791075d451fc415ca83ec8958b385a95
		
	Remote Server Administration Tools for Windows 8.1 
		https://carlwebster.sharefile.com/d-s37209afb73dc48f497745924ed854226
		
	Remote Server Administration Tools for Windows 10
		http://www.microsoft.com/en-us/download/details.aspx?id=45520
	
	For Windows Server 2003, 2008, and 2008 R2, use the following to export and import the 
	DHCP data:
		Export from the 2003, 2008, or 2008 R2 server:
			netsh dhcp server export C:\DHCPExport.txt all
			
			Copy the C:\DHCPExport.txt file to the 2012+ server.
			
		Import on the 2012+ server:
			netsh dhcp server import c:\DHCPExport.txt all
			
		The script can now be run on the 2012+ DHCP server to document the older DHCP 
		information.

	For Windows Server 2008 and Server 2008 R2, the 2012+ DHCP Server PowerShell cmdlets 
	can be used for export and import.
		From the 2012+ DHCP server:
			Export-DhcpServer -ComputerName 2008R2Server.domain.tld -Leases -File 
			C:\DHCPExport.xml 
			
			Import-DhcpServer -ComputerName 2012Server.domain.tld -Leases -File 
			C:\DHCPExport.xml -BackupPath C:\dhcp\backup\ 
			
			Note: The c:\dhcp\backup path must exist before running the 
			Import-DhcpServer cmdlet.
	
	Using netsh is much faster than using the PowerShell export and import cmdlets.
	
	Processing of IPv4 Multicast Scopes is only available with Server 2012 R2 DHCP.
	
	Word and PDF Documents include a Cover Page, Table of Contents, and Footer.
	
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
		
.PARAMETER ComputerName
	DHCP server to run the script against.
	The computername is used for the report title.
	ComputerName can be entered as the NetBIOS name, FQDN, localhost or IP Address.
	If entered as localhost, the actual computer name is determined and used.
	If entered as an IP address, an attempt is made to determine and use the actual 
	computer name.
	
	If both ComputerName and AllDHCPServers are used, AllDHCPServers is used.
.PARAMETER AllDHCPServers
	The script processes all Authorized DHCP servers that are online.
	"All DHCP Servers" is used for the report title.
	This parameter is disabled by default.
	
	If both ComputerName and AllDHCPServers are used, AllDHCPServers is used.
	This parameter has an alias of ALL.
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	This parameter is set True if no other output format is selected.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	This parameter is disabled by default.
.PARAMETER Hardware
	Use WMI to gather hardware information on Computer System, Disks, Processor, and 
	Network Interface Cards

	This parameter may require the script runs from an elevated PowerShell session 
	using an account with permission to retrieve hardware information (i.e., Domain 
	Admin or Local Administrator).

	Selecting this parameter adds to both the time it takes to run the script and 
	size of the report.

	This parameter is disabled by default.
	This parameter has an alias of HW.
.PARAMETER IncludeLeases
	Include DHCP lease information.
	The default is to not included lease information.
.PARAMETER IncludeOptions
	Include DHCP Options information.
	The default is to not included Options information.
.PARAMETER AddDateTime
	Adds a date Timestamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2021 at 6PM is 2021-06-01_1800.
	Output filename will either be:
		DHCP Inventory Report for Server <server> for the Domain <domain>_2021-06-01_1800.html (or .txt or .docx or .pdf)
		DHCP Inventory for All DHCP Servers for the Domain <domain>_2021-06-01_1800.html (or .txt or .docx or .pdf)
	This parameter is disabled by default.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	The text file is placed in the same folder from where the script runs.
	
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER Log
	Generates a log file for troubleshooting.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	The text file is placed in the same folder from where the script runs.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER ReportFooter
	Outputs a footer section at the end of the report.

	This parameter has an alias of RF.
	
	Report Footer
		Report information:
			Created with: <Script Name> - Release Date: <Script Release Date>
			Script version: <Script Version>
			Started on <Date Time in Local Format>
			Elapsed time: nn days, nn hours, nn minutes, nn.nn seconds
			Ran from domain <Domain Name> by user <Username>
			Ran from the folder <Folder Name>

	Script Name and Script Release date are script-specific variables.
	Start Date Time in Local Format is a script variable.
	Elapsed time is a calculated value.
	Domain Name is $env:USERDNSDOMAIN.
	Username is $env:USERNAME.
	Folder Name is a script variable.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is disabled by default.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
	This parameter requires Microsoft Word to be installed.
	This parameter uses the Word SaveAs PDF capability.
.PARAMETER CompanyAddress
	Company Address used for the Cover Page, if the Cover Page has the Address field.
	
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
	Company Email used for the Cover Page, if the Cover Page has the Email field.  
	
	The following Cover Pages have an Email field:
		Facet (Word 2013/2016)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CE.
.PARAMETER CompanyFax
	Company Fax used for the Cover Page, if the Cover Page has the Fax field.  
	
	The following Cover Pages have a Fax field:
		Contrast (Word 2010)
		Exposure (Word 2010)
	
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CF.
.PARAMETER CompanyName
	Company Name used for the Cover Page.  
	The default value is contained in 
	HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated 
	on the computer running the script.
	This parameter has an alias of CN.
.PARAMETER CompanyPhone
	Company Phone used for the Cover Page if the Cover Page has the Phone field.  
	
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
.PARAMETER UserName
	Username used for the Cover Page and Footer.
	The default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER SmtpServer
	Specifies the optional email server to send the output report. 
.PARAMETER SmtpPort
	Specifies the SMTP port. 
	Default is 25.
.PARAMETER UseSSL
	Specifies whether to use SSL for the SmtpServer.
	Default is False.
.PARAMETER From
	Specifies the username for the From email address.
	If SmtpServer is used, this is a required parameter.
.PARAMETER To
	Specifies the username for the To email address.
	If SmtpServer is used, this is a required parameter.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -ComputerName DHCPServer01 -MSWord
	
	Uses all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	The script runs remotely against the DHCP server DHCPServer01.

	Creates a Microsoft Word document.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -ComputerName localhost
	
	Script will resolve localhost to $env:computername, for example, DHCPServer01.
	The script runs remotely against the DHCP server DHCPServer01 and not localhost.
	The output filename uses the server name DHCPServer01 and not localhost.
	
	Creates an HTML file.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -ComputerName 192.168.1.222
	
	Script will resolve 192.168.1.222 to the DNS hostname, for example, DHCPServer01.
	The script runs remotely against the DHCP server DHCPServer01 and not 192.18.1.222.
	The output filename uses the server name DHCPServer01 and not 192.168.1.222.

	Creates an HTML file.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -PDF -ComputerName DHCPServer02
	
	Uses all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	The script runs remotely against the DHCP server DHCPServer02.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -Text -ComputerName DHCPServer02
	
	The script runs remotely against the DHCP server DHCPServer02.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -MSWord -ComputerName DHCPServer02
	
	Uses all Default values and save the document as a Word DOCX file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	The script runs remotely against the DHCP server DHCPServer02.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -Hardware -ComputerName DHCPServer02
	
	The script runs remotely against the DHCP server DHCPServer02.
	Creates an HTML file.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -ComputerName DHCPServer03 -IncludeLeases 
	-MSWord
	
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	The script runs remotely against the DHCP server DHCPServer03.
	The output contains DHCP lease information.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -AllDHCPServers -HTML -IncludeOptions
		
	The script finds all Authorized DHCP servers and processes all servers that are 
	online.
	The output contains DHCP Options information.
.EXAMPLE
	PS C:\PSScript .\DHCP_Inventory_V2.ps1 -CompanyName "Carl Webster Consulting" 
	-CoverPage "Mod" -UserName "Carl Webster" -ComputerName DHCPServer01 -MSWord

	Uses:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
	
	The script runs remotely against the DHCP server DHCPServer01.
.EXAMPLE
	PS C:\PSScript .\DHCP_Inventory_V2.ps1 -CN "Carl Webster Consulting" -CP "Mod" 
	-UN "Carl Webster" -ComputerName DHCPServer02 -IncludeLeases -MSWord

	Uses:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
	
	The script runs remotely against the DHCP server DHCPServer02.
	The output contains DHCP lease information.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -AddDateTime
	
	Adds a date Timestamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	July 25, 2021 at 6PM is 2021-07-25_1800.
	The output filename will be DHCP Inventory Report for Server <server> for the Domain 
	<domain>_2021-07-25_1800.html
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -PDF -AddDateTime
	
	Will use all default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date Timestamp to the end of the file name.
	The timestamp is in the format of yyyy-MM-dd_HHmm.
	July 25, 2021 at 6PM is 2021-07-25_1800.
	The output filename will be DHCP Inventory Report for Server <server> for the Domain 
	<domain>_2021-07-25_1800.pdf
.EXAMPLE
	PS C:\PSScript .\DHCP_Inventory_V2.ps1 -MSWord -CompanyName "Sherlock Holmes 
	Consulting" -CoverPage Exposure -UserName "Dr. Watson" -CompanyAddress "221B Baker 
	Street, London, England" -CompanyFax "+44 1753 276600" -CompanyPhone "+44 1753 276200"
	
	Uses:
		Sherlock Holmes Consulting for the Company Name.
		Exposure for the Cover Page format.
		Dr. Watson for the User Name.
		221B Baker Street, London, England for the Company Address.
		+44 1753 276600 for the Company Fax.
		+44 1753 276200 for the Company Phone.
.EXAMPLE
	PS C:\PSScript .\DHCP_Inventory_V2.ps1 -MSWord -CompanyName "Sherlock Holmes 
	Consulting" -CoverPage Facet -UserName "Dr. Watson" -CompanyEmail 
	SuperSleuth@SherlockHolmes.com
	
	Uses:
		Sherlock Holmes Consulting for the Company Name.
		Facet for the Cover Page format.
		Dr. Watson for the User Name.
		SuperSleuth@SherlockHolmes.com for the Company Email.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -Folder \\FileServer\ShareName
	
	Output HTML file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -HTML -MSWord -PDF -Text -Dev -ScriptInfo -Log 
	-ComputerName DHCPServer
	
	Creates four reports: HTML, Microsoft Word, PDF, and plain text.
	
	Creates a text file named DHCPInventoryScriptErrors_yyyy-MM-dd_HHmm for the Domain 
	<domain>.txt that contains up to the last 250 errors reported by the script.
	
	Creates a text file named DHCPInventoryScriptInfo_yyyy-MM-dd_HHmm for the Domain 
	<domain>.txt that contains all the script parameters and other basic information.
	
	Creates a text file for transcript logging named 
	DHCPDocScriptTranscript_yyyy-MM-dd_HHmm for the Domain <domain>.txt.

	For Microsoft Word and PDF, uses all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	The script runs remotely against the DHCP server DHCPServer.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -SmtpServer mail.domain.tld -From 
	XDAdmin@domain.tld -To ITGroup@domain.tld	

	The script uses the email server mail.domain.tld, sending from XDAdmin@domain.tld, 
	sending to ITGroup@domain.tld.

	The script uses the default SMTP port 25 and does not use SSL.

	If the current user's credentials are not valid to send an email, 
	the script prompts the user to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -SmtpServer mailrelay.domain.tld -From 
	Anonymous@domain.tld -To ITGroup@domain.tld	

	***SENDING UNAUTHENTICATED EMAIL***

	The script uses the email server mailrelay.domain.tld, sending from 
	anonymous@domain.tld, sending to ITGroup@domain.tld.

	To send unauthenticated email using an email relay server requires the From email account 
	to use the name Anonymous.

	The script uses the default SMTP port 25 and does not use SSL.
	
	***GMAIL/G SUITE SMTP RELAY***
	https://support.google.com/a/answer/2956491?hl=en
	https://support.google.com/a/answer/176600?hl=en

	To send an email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	***GMAIL/G SUITE SMTP RELAY***

	The script generates an anonymous, secure password for the anonymous@domain.tld 
	account.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -SmtpServer 
	labaddomain-com.mail.protection.outlook.com -UseSSL -From 
	SomeEmailAddress@labaddomain.com -To ITGroupDL@labaddomain.com	

	***OFFICE 365 Example***

	https://docs.microsoft.com/en-us/exchange/mail-flow-best-practices/how-to-set-up-a-multifunction-device-or-application-to-send-email-using-office-3
	
	This uses Option 2 from the above link.
	
	***OFFICE 365 Example***

	The script uses the email server labaddomain-com.mail.protection.outlook.com, 
	sending from SomeEmailAddress@labaddomain.com, sending to ITGroupDL@labaddomain.com.

	The script uses the default SMTP port 25 and uses SSL.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -SmtpServer smtp.office365.com -SmtpPort 587 
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	The script uses the email server smtp.office365.com on port 587 using SSL, 
	sending from webster@carlwebster.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send an email, 
	the script prompts the user to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -SmtpServer smtp.gmail.com -SmtpPort 587 
	-UseSSL -From Webster@CarlWebster.com -To ITGroup@CarlWebster.com	

	*** NOTE ***
	To send an email using a Gmail or g-suite account, you may have to turn ON
	the "Less secure app access" option on your account.
	*** NOTE ***
	
	The script uses the email server smtp.gmail.com on port 587 using SSL, 
	sending from webster@gmail.com, sending to ITGroup@carlwebster.com.

	If the current user's credentials are not valid to send an email, 
	the script prompts the user to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -Dev -ScriptInfo -Log
	
	Creates a text file named DHCPInventoryScriptErrors_yyyy-MM-dd_HHmm for the Domain 
	<domain>.txt that contains up to the last 250 errors reported by the script.
	
	Creates a text file named DHCPInventoryScriptInfo_yyyy-MM-dd_HHmm for the Domain 
	<domain>.txt that contains all the script parameters and other basic information.
	
	Creates a text file for transcript logging named 
	DHCPDocScriptTranscript_yyyy-MM-dd_HHmm for the Domain <domain>.txt.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -ComputerName DHCPServer01 -Hardware
	
	Adds additional information for the server about its hardware.
	
	The script runs remotely against the DHCP server DHCPServer01.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -AllDHCPServers
	
	The script finds all Authorized DHCP servers and processes all servers that are 
	online.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V2.ps1 -ComputerName DHCPServer01 -AllDHCPServers
	
	Even though DHCPServer01 is specified, the script finds all Authorized DHCP servers 
	and processes all online servers.
.INPUTS
	None. You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word, PDF, HTML or 
	formatted text document.
.NOTES
	NAME: DHCP_Inventory_V2.ps1
	VERSION: 2.04
	AUTHOR: Carl Webster and Michael B. Smith
	LASTEDIT: September 11, 2021
#>

#endregion


#region script parameters
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "") ]

Param(
	[parameter(Mandatory=$False)] 
	[string]$ComputerName="LocalHost",
	
	[parameter(Mandatory=$False)] 
	[Alias("ALL")]
	[Switch]$AllDHCPServers=$False,
	
	[parameter(Mandatory=$False)] 
	[Switch]$HTML=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(Mandatory=$False)] 
	[Alias("HW")]
	[Switch]$Hardware=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$IncludeLeases=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$IncludeOptions=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$AddDateTime=$False,
	
	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(Mandatory=$False)] 
	[Switch]$Log=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("RF")]
	[Switch]$ReportFooter=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CA")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyAddress="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CE")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyEmail="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CF")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyFax="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CPh")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyPhone="",
    
	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(ParameterSetName="WordPDF",Mandatory=$False)] 
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
#endregion

#region script change log	
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

#Version 2.04 11-Sep-2021
#	Added array error checking for non-empty arrays before attempting to create the Word table for most Word tables
#	Added Function OutputReportFooter
#	Added Parameter ReportFooter
#		Outputs a footer section at the end of the report.
#		Report Footer
#			Report information:
#				Created with: <Script Name> - Release Date: <Script Release Date>
#				Script version: <Script Version>
#				Started on <Date Time in Local Format>
#				Elapsed time: nn days, nn hours, nn minutes, nn.nn seconds
#				Ran from domain <Domain Name> by user <Username>
#				Ran from the folder <Folder Name>
#	Updated Functions SaveandCloseTextDocument and SaveandCloseHTMLDocument to add a "Report Complete" line
#	Updated Functions ShowScriptOptions and ProcessScriptEnd to add $ReportFooter
#	Updated the help text
#	Updated the ReadMe file
#
#Version 2.03 9-Jan-2021
#	Added to the Computer Hardware section, the server's Power Plan
#	Fixed Date calculation errors with IPv4 and IPv6 statistics and script runtime
#	Reordered parameters in an order recommended by Guy Leech
#	Updated help text
#	Updated ReadMe file
#
#Version 2.02 5-Nov-2020
#	Added to the server properties, "Is a domain controller" with a value of Yes or No
#	Changed all Write-Verbose $(Get-Date) to add -Format G to put the dates in the user's locale
#
#Version 2.01 30-Oct-2020
#	Added variable $Script:RptDomain
#	For the -Dev, -Log, and -ScriptInfo output files, add the text "for the Domain <domain>"
#	Updated the help text
#	Updated the ReadMe file (https://carlwebster.sharefile.com/d-s6b941ce3a4643df8)
#	Updated the report title and output filenames to:
#		For using -ComputerName:
#			DHCP Inventory Report for Server <DHCPServerName> for the Domain <domain>
#		For using AllDHCPServers:
#			DHCP Inventory Report for All DHCP Servers for the Domain <domain>
#
#Version 2.00 26-Oct-2020
#	Changed all Word/PDF tables to Ian Brighton's table functions
#	Changed color variables $wdColorGray15 and $wdColorGray05 from [long] to [int]
#	Changed Network Access Protection Status values to a Switch table to get the full text label
#	Changed several sections from outputting in WriteWordLine to Word/PDF tables
#	Changed sorting of DHCP Reservations from IPAddress to Name
#	Changed formatting of Statistics percentages from (x.x)% to x.x%
#	Cleanup formatting of Text output
#	Combine functions ProcessIPv4Bindings and ProcessIPv6Bindings into one function ProcessIPBindings
#	Fixed several typos in text labels
#	General code cleanup
#	HTML is now the default output format.
#	In IPv4 protocol and IPv4 Scopes DNS Settings, add missing options:
#		Dynamically update DNS records for DHCP clients to do not request updates
#		Disable dynamic updates for DNS PTR record
#	In IPv4 protocol and IPv4 Scopes Policies, add missing Policy data
#		General
#			Lease duration for DHCP client
#		Conditions
#			Condition
#			Conditions
#			Operator
#			Value
#		IP Address Range
#			Added support for multiple IP address ranges
#		Options
#			Option Name
#			Vendor
#			Value
#			Policy Name
#		DNS
#			Enable DNS dynamic updates
#				Dynamically update DNS A (or AAAA) and PTR records only if requested by the DHCP clients
#				Always dynamically update DNS A (or AAAA) and PTR records
#			Discard A (or AAAA) and PTR records when lease deleted
#			Dynamically update DNS records for DHCP clients to do not request updates
#			Disable dynamic updates for DNS PTR record
#			Register DHCP clients using the following DNS suffix
#			Name Protection
#	In the various Statistics text output sections, remove the ":" from the text labels
#	Optimize the output of Failover data
#	Removed the calls to [gc]::collect
#	Removed function ValidateWordTableValues as it is no longer needed
#	Updated Function CheckWordPrereq to the latest 
#	Updated Function SetupWord to the latest
#	Updated Function SetWordCellFormat to the latest
#	Updated Hardware inventory functions to the latest
#	Updated help text
#	Updated ReadMe file (https://carlwebster.sharefile.com/d-s6b941ce3a4643df8)
#	You can now select multiple output formats. This required extensive code changes.
#
#Version 1.44 8-May-2020
#	Add checking for a Word version of 0, which indicates the Office installation needs repairing
#	Add Receive Side Scaling setting to Function OutputNICItem
#	Change color variables $wdColorGray15 and $wdColorGray05 from [long] to [int]
#	Change location of the -Dev, -Log, and -ScriptInfo output files from the script folder to the -Folder location (Thanks to Guy Leech for the "suggestion")
#	Reformatted the terminating Write-Error messages to make them more visible and readable in the console
#	Remove manually checking for multiple output formats
#	Remove the SMTP parameterset and manually verify the parameters
#	Update Function SendEmail to handle anonymous unauthenticated email
#	Update Function SetWordCellFormat to change parameter $BackgroundColor to [int]
#	Update Functions GetComputerWMIInfo and OutputNicInfo to fix two bugs in NIC Power Management settings
#	Update Help Text

#Version 1.43 17-Apr-2020
#	Add parameter IncludeOptions to add DHCP Options to report.
#		New Function ProcessDHCPOptions
#		Update Functions ShowScriptOptions and ProcessScriptEnd 
#		Update Help Text
#	Cleanup spacing in some of the Write-Verbose statements
#	In the GetIPv4ScopeData functions, ignore Option ID 81
#		This Option ID is set when Name Protection is disabled in the DNS
#		tab in a Scope's Properties. Option ID 81 is not in the Predefined Options.
#		https://carlwebster.com/the-mysterious-microsoft-dhcp-option-id-81/
#	Reorder parameters
#	Update Function SendEmail to handle anonymous unauthenticated email
#		Update Help Text with examples
#	Update script to match updates to other documentation scripts
#		Checking for multiple output formats selected
#		Change Text output to use [System.Text.StringBuilder]
#		Function Line
#		Function SaveandCloseTextDocument
#		Function WriteWordLine
#		Function WriteHTMLLine
#		Function AddHTMLTable
#		Function FormatHTMLTable
#		Function FormatHTMLTable
#		Function CheckHTMLColor
#
#Version 1.42 17-Dec-2019
#	Fix Swedish Table of Contents (Thanks to Johan Kallio)
#		From 
#			'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
#		To
#			'sv-'	{ 'Automatisk innehållsförteckn2'; Break }
#	Updated help text
#
#Version 1.41 8-Jan-2019
#	Cleaned up help text
#	Reorganized parameters
#
#Version 1.40 5-Apr-2018
#	Added -AllDHCPServers (ALL) parameter to process all Authorized DHCP servers that are online
#		Added text file (BadDHCPServers_yyyy-MM-dd_HHmm.txt) of the authorized DHCP servers that 
#		are either offline or no longer have DHCP installed
#	Added -Hardware parameter
#		Added functions to output hardware information
#	Code clean-up for most recommendations made by Visual Studio Code
#	Fixed several minor issues found during testing from the code cleanup
#	Grouped code into functions and functions into regions
#	In the Scope Options, if all Scope Options inherit from Server Options and the only 
#		scope option is the implied Option ID 51; then blank lines were inserted. This is now 
#		fixed so "None" is reported, just like all the other items. For some reason, Option ID 
#		51 is implied and even though it does not show in the console, the PowerShell cmdlet 
#		exposes it. If I try and retrieve the properties of that option, it can crash the computer 
#		running the script. Not a good thing if you are running the script on a DHCP server. I now 
#		check for this specific condition, and it is now handled properly for all output types.
#		Many thanks to my exhaustive tester, David McSpadden, for helping find and fix this logic flaw.
#	Updated help text
#
#Version 1.35 10-Feb-2017
#	Added four new Cover Page properties
#		Company Address
#		Company Email
#		Company Fax
#		Company Phone
#	Added Log switch to create a transcript log
#	Replaced _SetDocumentProperty function with Jim Moyle's Set-DocumentProperty function
#	Removed code that made sure all Parameters were set to default values if for some reason they did exist or values were $Null
#	Updated Function ProcessScriptEnd for the new Cover Page properties and Parameters
#	Updated Function ShowScriptOptions for the new Cover Page properties and Parameters
#	Updated Function UpdateDocumentProperties for the new Cover Page properties and Parameters
#	Updated help text
#
#Version 1.34 8-Dec-2017
#	Updated Function WriteHTMLLine with fixes from the script template
#
#Version 1.33 13-Feb-2017
#	Fixed French wording for Table of Contents 2 (Thanks to David Rouquier)
#
#Version 1.32 7-Nov-2016
#	Added Chinese language support
#
#Version 1.31 24-Oct-2016
#	Add HTML output
#	Fix typo on failover status "iitializing" -> "initializing"
#	Fix numerous issues where I used .day/.hour/.minute instead of .days/.hours/.minutes when formatting times
#
#Version 1.30 4-May-2016
#	Fixed numerous issues discovered with the latest update to PowerShell V5
#	Color variables needed to be [long] and not [int] except for $wdColorBlack which is 0
#	Changed from using arrays to populating data in tables to strings
#	Fixed several incorrect variable names that kept PDFs from saving in Windows 10 and Office 2013
#	Fixed the rest of the $Var -eq $Null to $Null -eq $Var
#	Removed blocks of old commented out code
#	Removed the 10 second pauses waiting for Word to save and close.
#	Added -Dev parameter to create a text file of script errors
#	Added -ScriptInfo (SI) parameter to create a text file of script information
#	Added more script information to the console output when script starts
#	Cleaned up some issues in the help text
#	Commented out HTML parameters as HTML output is not ready
#	Added HTML functions to prep for adding HTML output
#
#Version 1.24 8-Feb-2016
#	Added specifying an optional output folder
#	Added the option to email the output file
#	Fixed several spacing and typo errors
#
#Version 1.23 1-Feb-2016
#	Added DNS Dynamic update credentials from protocol properties, advanced tab
#
#Version 1.22 25-Nov-2015
#	Updated help text and ReadMe for RSAT for Windows 10
#	Updated ReadMe with an example of running the script remotely
#	Tested script on Windows 10 x64 and Word 2016 x64
#
#Version 1.21 5-Oct-2015
#	Added support for Word 2016
#
#Version 1.2
#	Cleanup some of the console output
#	Added error checking:
#	If script is run without -ComputerName, resolve LocalHost to computer and very it is a DHCP server
#	If script is run with -ComputerName, very it is a DHCP server
#
#Version 1.1
#	Cleanup the script's parameters section
#	Code cleanup and standardization with the master template script
#	Requires PowerShell V3 or later
#	Removed support for Word 2007
#	Word 2007 references in help text removed
#	Cover page parameter now states only Word 2010 and 2013 are supported
#	Most Word 2007 references in script removed:
#		Function ValidateCoverPage
#		Function SetupWord
#		Function SaveandCloseDocumentandShutdownWord
#	Function CheckWord2007SaveAsPDFInstalled removed
#	If Word 2007 is detected, an error message is now given and the script is aborted
#	Fix when -ComputerName was entered as LocalHost, output filename said LocalHost and not actual server name
#	Cleanup Word table code for the first row and background color
#	Add Iain Brighton's Word table functions
#
#Version 1.01
#	Added an AddDateTime parameter
#
#endregion


#region initial variable testing and setup
Set-StrictMode -Version Latest

#force  on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
$SaveEAPreference         = $ErrorActionPreference
$ErrorActionPreference    = 'SilentlyContinue'
$global:emailCredentials  = $Null
$Script:RptDomain         = (Get-WmiObject -computername $ComputerName win32_computersystem).Domain
$script:MyVersion         = '2.04'
$Script:ScriptName        = "DHCP_Inventory_V2.ps1"
$tmpdate                  = [datetime] "09/11/2021"
$Script:ReleaseDate       = $tmpdate.ToUniversalTime().ToShortDateString()


If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$HTML = $True
}

If($MSWord)
{
	Write-Verbose "$(Get-Date -Format G): MSWord is set"
}
If($PDF)
{
	Write-Verbose "$(Get-Date -Format G): PDF is set"
}
If($Text)
{
	Write-Verbose "$(Get-Date -Format G): Text is set"
}
If($HTML)
{
	Write-Verbose "$(Get-Date -Format G): HTML is set"
}

If($Folder -ne "")
{
	Write-Verbose "$(Get-Date -Format G): Testing folder path"
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Verbose "$(Get-Date -Format G): Folder path $Folder exists and is a folder"
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


#V1.35 added
If($Log) 
{
	#start transcript logging
	$Script:LogPath = "$($Script:pwdpath)\DHCPDocScriptTranscript_$(Get-Date -f yyyy-MM-dd_HHmm) for the Domain $Script:RptDomain.txt"
	
	try 
	{
		Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
		Write-Verbose "$(Get-Date -Format G): Transcript/log started at $Script:LogPath"
		$Script:StartLog = $true
	} 
	catch 
	{
		Write-Verbose "$(Get-Date -Format G): Transcript/log failed at $Script:LogPath"
		$Script:StartLog = $false
	}
}

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$($Script:pwdpath)\DHCPInventoryScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm) for the Domain $Script:RptDomain.txt"
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

#endregion

#region initialize variables for Word, HTML, and text
[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$Script:CoName = $CompanyName
	Write-Verbose "$(Get-Date -Format G): CoName is $($Script:CoName)"
	
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[int]$wdColorGray15 = 14277081
	[int]$wdColorGray05 = 15987699 
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[int]$wdColorRed = 255
	[int]$wdColorWhite = 16777215
	[int]$wdColorBlack = 0
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
	[int]$wdTableLightListAccent3 = -206

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	[int]$wdHeadingFormatFalse = 0 
	
	[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption
}

If($HTML)
{
	#V2.23 Prior versions used Set-Variable. That hid the variables
	#from @code. So MBS switched to using $global:

    $global:htmlredmask       = "#FF0000" 4>$Null
    $global:htmlcyanmask      = "#00FFFF" 4>$Null
    $global:htmlbluemask      = "#0000FF" 4>$Null
    $global:htmldarkbluemask  = "#0000A0" 4>$Null
    $global:htmllightbluemask = "#ADD8E6" 4>$Null
    $global:htmlpurplemask    = "#800080" 4>$Null
    $global:htmlyellowmask    = "#FFFF00" 4>$Null
    $global:htmllimemask      = "#00FF00" 4>$Null
    $global:htmlmagentamask   = "#FF00FF" 4>$Null
    $global:htmlwhitemask     = "#FFFFFF" 4>$Null
    $global:htmlsilvermask    = "#C0C0C0" 4>$Null
    $global:htmlgraymask      = "#808080" 4>$Null
    $global:htmlblackmask     = "#000000" 4>$Null
    $global:htmlorangemask    = "#FFA500" 4>$Null
    $global:htmlmaroonmask    = "#800000" 4>$Null
    $global:htmlgreenmask     = "#008000" 4>$Null
    $global:htmlolivemask     = "#808000" 4>$Null

    $global:htmlbold        = 1 4>$Null
    $global:htmlitalics     = 2 4>$Null
    $global:htmlred         = 4 4>$Null
    $global:htmlcyan        = 8 4>$Null
    $global:htmlblue        = 16 4>$Null
    $global:htmldarkblue    = 32 4>$Null
    $global:htmllightblue   = 64 4>$Null
    $global:htmlpurple      = 128 4>$Null
    $global:htmlyellow      = 256 4>$Null
    $global:htmllime        = 512 4>$Null
    $global:htmlmagenta     = 1024 4>$Null
    $global:htmlwhite       = 2048 4>$Null
    $global:htmlsilver      = 4096 4>$Null
    $global:htmlgray        = 8192 4>$Null
    $global:htmlolive       = 16384 4>$Null
    $global:htmlorange      = 32768 4>$Null
    $global:htmlmaroon      = 65536 4>$Null
    $global:htmlgreen       = 131072 4>$Null
	$global:htmlblack       = 262144 4>$Null

	$global:htmlsb          = ( $htmlsilver -bor $htmlBold ) ## point optimization

	$global:htmlColor = 
	@{
		$htmlred       = $htmlredmask
		$htmlcyan      = $htmlcyanmask
		$htmlblue      = $htmlbluemask
		$htmldarkblue  = $htmldarkbluemask
		$htmllightblue = $htmllightbluemask
		$htmlpurple    = $htmlpurplemask
		$htmlyellow    = $htmlyellowmask
		$htmllime      = $htmllimemask
		$htmlmagenta   = $htmlmagentamask
		$htmlwhite     = $htmlwhitemask
		$htmlsilver    = $htmlsilvermask
		$htmlgray      = $htmlgraymask
		$htmlolive     = $htmlolivemask
		$htmlorange    = $htmlorangemask
		$htmlmaroon    = $htmlmaroonmask
		$htmlgreen     = $htmlgreenmask
		$htmlblack     = $htmlblackmask
	}
}
#endregion

#region code for hardware data
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
	Write-Verbose "$(Get-Date -Format G): `t`tProcessing WMI Computer information"
	Write-Verbose "$(Get-Date -Format G): `t`t`tHardware information"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteWordLine 4 0 "General Computer"
	}
	If($Text)
	{
		Line 0 "Computer Information: $($RemoteComputerName)"
		Line 1 "General Computer"
	}
	If($HTML)
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
			OutputComputerItem $Item $ComputerOS $RemoteComputerName
		}
	}
	ElseIf(!$?)
	{
		Write-Verbose "$(Get-Date -Format G): Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
			Line 2 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" -option $htmlBold
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" -option $htmlBold
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" -Option $htmlBold
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." -Option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): No results Returned for Computer information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "No results Returned for Computer information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Computer information" -Option $htmlBold
		}
	}
	
	#Get Disk info
	Write-Verbose "$(Get-Date -Format G): `t`t`tDrive information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Drive(s)"
	}
	If($Text)
	{
		Line 1 "Drive(s)"
	}
	If($HTML)
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
		Write-Verbose "$(Get-Date -Format G): Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" -Option $htmlBold
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" -Option $htmlBold
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" -Option $htmlBold
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." -Option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): No results Returned for Drive information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "No results Returned for Drive information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Drive information" -Option $htmlBold
		}
	}
	
	#Get CPU's and stepping
	Write-Verbose "$(Get-Date -Format G): `t`t`tProcessor information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Processor(s)"
	}
	If($Text)
	{
		Line 1 "Processor(s)"
	}
	If($HTML)
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
		Write-Verbose "$(Get-Date -Format G): Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" -Option $htmlBold
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" -Option $htmlBold
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" -Option $htmlBold
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." -Option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): No results Returned for Processor information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "No results Returned for Processor information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Processor information" -Option $htmlBold
		}
	}

	#Get Nics
	Write-Verbose "$(Get-Date -Format G): `t`t`tNIC information"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Network Interface(s)"
	}
	If($Text)
	{
		Line 1 "Network Interface(s)"
	}
	If($HTML)
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
					Write-Warning "$(Get-Date -Format G): Error retrieving NIC information"
					Write-Verbose "$(Get-Date -Format G): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteWordLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
					If($Text)
					{
						Line 2 "Error retrieving NIC information"
						Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
						Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
						Line 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may"
						Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
					}
					If($HTML)
					{
						WriteHTMLLine 0 2 "Error retrieving NIC information" -Option $htmlBold
						WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" -Option $htmlBold
						WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" -Option $htmlBold
						WriteHTMLLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" -Option $htmlBold
						WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." -Option $htmlBold
					}
				}
				Else
				{
					Write-Verbose "$(Get-Date -Format G): No results Returned for NIC information"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
					If($Text)
					{
						Line 2 "No results Returned for NIC information"
					}
					If($HTML)
					{
						WriteHTMLLine 0 2 "No results Returned for NIC information" -Option $htmlBold
					}
				}
			}
		}	
	}
	ElseIf(!$?)
	{
		Write-Warning "$(Get-Date -Format G): Error retrieving NIC configuration information"
		Write-Verbose "$(Get-Date -Format G): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "Error retrieving NIC configuration information"
			Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "Error retrieving NIC configuration information" -Option $htmlBold
			WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" -Option $htmlBold
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" -Option $htmlBold
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository. If this is a trusted Forest, you may" -Option $htmlBold
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." -Option $htmlBold
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): No results Returned for NIC configuration information"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
		If($Text)
		{
			Line 2 "No results Returned for NIC configuration information"
		}
		If($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for NIC configuration information" -Option $htmlBold
		}
	}
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 0 0 ""
	}
}

Function OutputComputerItem
{
	Param([object]$Item, [string]$OS, [string]$RemoteComputerName)
	
	#get computer's power plan
	#https://techcommunity.microsoft.com/t5/core-infrastructure-and-security/get-the-active-power-plan-of-multiple-servers-with-powershell/ba-p/370429
	
	try 
	{

		$PowerPlan = (Get-WmiObject -ComputerName $RemoteComputerName -Class Win32_PowerPlan -Namespace "root\cimv2\power" |
			Where-Object {$_.IsActive -eq $true} |
			Select-Object @{Name = "PowerPlan"; Expression = {$_.ElementName}}).PowerPlan
	}

	catch 
	{

		$PowerPlan = $_.Exception

	}	
	
	If($MSWord -or $PDF)
	{
		$ItemInformation = New-Object System.Collections.ArrayList
		$ItemInformation.Add(@{ Data = "Manufacturer"; Value = $Item.manufacturer; }) > $Null
		$ItemInformation.Add(@{ Data = "Model"; Value = $Item.model; }) > $Null
		$ItemInformation.Add(@{ Data = "Domain"; Value = $Item.domain; }) > $Null
		$ItemInformation.Add(@{ Data = "Operating System"; Value = $OS; }) > $Null
		$ItemInformation.Add(@{ Data = "Power Plan"; Value = $PowerPlan; }) > $Null
		$ItemInformation.Add(@{ Data = "Total Ram"; Value = "$($Item.totalphysicalram) GB"; }) > $Null
		$ItemInformation.Add(@{ Data = "Physical Processors (sockets)"; Value = $Item.NumberOfProcessors; }) > $Null
		$ItemInformation.Add(@{ Data = "Logical Processors (cores w/HT)"; Value = $Item.NumberOfLogicalProcessors; }) > $Null
		$Table = AddWordTable -Hashtable $ItemInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
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
	If($Text)
	{
		Line 2 "Manufacturer`t`t`t: " $Item.manufacturer
		Line 2 "Model`t`t`t`t: " $Item.model
		Line 2 "Domain`t`t`t`t: " $Item.domain
		Line 2 "Operating System`t`t: " $OS
		Line 2 "Power Plan`t`t`t: " $PowerPlan
		Line 2 "Total Ram`t`t`t: $($Item.totalphysicalram) GB"
		Line 2 "Physical Processors (sockets)`t: " $Item.NumberOfProcessors
		Line 2 "Logical Processors (cores w/HT)`t: " $Item.NumberOfLogicalProcessors
		Line 2 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Manufacturer",($htmlsilver -bor $htmlBold),$Item.manufacturer,$htmlwhite)
		$rowdata += @(,('Model',($htmlsilver -bor $htmlBold),$Item.model,$htmlwhite))
		$rowdata += @(,('Domain',($htmlsilver -bor $htmlBold),$Item.domain,$htmlwhite))
		$rowdata += @(,('Operating System',($htmlsilver -bor $htmlBold),$OS,$htmlwhite))
		$rowdata += @(,('Power Plan',($htmlsilver -bor $htmlBold),$PowerPlan,$htmlwhite))
		$rowdata += @(,('Total Ram',($htmlsilver -bor $htmlBold),"$($Item.totalphysicalram) GB",$htmlwhite))
		$rowdata += @(,('Physical Processors (sockets)',($htmlsilver -bor $htmlBold),$Item.NumberOfProcessors,$htmlwhite))
		$rowdata += @(,('Logical Processors (cores w/HT)',($htmlsilver -bor $htmlBold),$Item.NumberOfLogicalProcessors,$htmlwhite))

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
		-Format $wdTableGrid `
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
	If($Text)
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
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Caption",($htmlsilver -bor $htmlBold),$Drive.caption,$htmlwhite)
		$rowdata += @(,('Size',($htmlsilver -bor $htmlBold),"$($drive.drivesize) GB",$htmlwhite))

		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$rowdata += @(,('File System',($htmlsilver -bor $htmlBold),$Drive.filesystem,$htmlwhite))
		}
		$rowdata += @(,('Free Space',($htmlsilver -bor $htmlBold),"$($drive.drivefreespace) GB",$htmlwhite))
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$rowdata += @(,('Volume Name',($htmlsilver -bor $htmlBold),$Drive.volumename,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$rowdata += @(,('Volume is Dirty',($htmlsilver -bor $htmlBold),$xVolumeDirty,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$rowdata += @(,('Volume Serial Number',($htmlsilver -bor $htmlBold),$Drive.volumeserialnumber,$htmlwhite))
		}
		$rowdata += @(,('Drive Type',($htmlsilver -bor $htmlBold),$xDriveType,$htmlwhite))

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
		-Format $wdTableGrid `
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
	If($Text)
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
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlBold),$Processor.name,$htmlwhite)
		$rowdata += @(,('Description',($htmlsilver -bor $htmlBold),$Processor.description,$htmlwhite))

		$rowdata += @(,('Max Clock Speed',($htmlsilver -bor $htmlBold),"$($processor.maxclockspeed) MHz",$htmlwhite))
		If($processor.l2cachesize -gt 0)
		{
			$rowdata += @(,('L2 Cache Size',($htmlsilver -bor $htmlBold),"$($processor.l2cachesize) KB",$htmlwhite))
		}
		If($processor.l3cachesize -gt 0)
		{
			$rowdata += @(,('L3 Cache Size',($htmlsilver -bor $htmlBold),"$($processor.l3cachesize) KB",$htmlwhite))
		}
		If($processor.numberofcores -gt 0)
		{
			$rowdata += @(,('Number of Cores',($htmlsilver -bor $htmlBold),$Processor.numberofcores,$htmlwhite))
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$rowdata += @(,('Number of Logical Processors (cores w/HT)',($htmlsilver -bor $htmlBold),$Processor.numberoflogicalprocessors,$htmlwhite))
		}
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlBold),$xAvailability,$htmlwhite))

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
		Else
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
			If($Nic.Defaultipgateway)
			{
				$NicInformation.Add(@{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }) > $Null
			}
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
			If($Nic.Defaultipgateway)
			{
				$NicInformation.Add(@{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }) > $Null
			}
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
		$Table = AddWordTable -Hashtable $NicInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
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
	If($Text)
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
		If($xIPAddress.Count -gt 1)
		{
			Line 2 "IP Address: " $xIPAddress[0]
			If($Nic.Defaultipgateway)
			{
				Line 2 "Default Gateway: " $Nic.Defaultipgateway
			}
			Line 2 "Subnet Mask: " $xIPSubnet[0]
			$cnt = -1
			ForEach($tmp in $xIPAddress)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 2 "IP Address: " $tmp
					Line 2 "Subnet Mask: " $xIPSubnet[$cnt]
				}
			}
		}
		Else
		{
			Line 2 "IP Address: " $xIPAddress
			If($Nic.Defaultipgateway)
			{
				Line 2 "Default Gateway: " $Nic.Defaultipgateway
			}
			Line 2 "Subnet Mask: " $xIPSubnet
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
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlBold),$ThisNic.Name,$htmlwhite)
		If($ThisNic.Name -ne $nic.description)
		{
			$rowdata += @(,('Description',($htmlsilver -bor $htmlBold),$Nic.description,$htmlwhite))
		}
		$rowdata += @(,('Connection ID',($htmlsilver -bor $htmlBold),$ThisNic.NetConnectionID,$htmlwhite))
		If(validObject $Nic Manufacturer)
		{
			$rowdata += @(,('Manufacturer',($htmlsilver -bor $htmlBold),$Nic.manufacturer,$htmlwhite))
		}
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlBold),$xAvailability,$htmlwhite))
		$rowdata += @(,('Allow the computer to turn off this device to save power',($htmlsilver -bor $htmlBold),$PowerSaving,$htmlwhite))
		$rowdata += @(,('Receive Side Scaling',($htmlsilver -bor $htmlbold),$RSSEnabled,$htmlwhite))
		$rowdata += @(,('Physical Address',($htmlsilver -bor $htmlBold),$Nic.macaddress,$htmlwhite))
		If($xIPAddress.Count -gt 1)
		{
			$rowdata += @(,("IP Address",($htmlsilver -bor $htmlBold),$xIPAddress[0],$htmlwhite))
			If($Nic.Defaultipgateway)
			{
				$rowdata += @(,("Default Gateway",($htmlsilver -bor $htmlBold),$Nic.Defaultipgateway,$htmlwhite))
			}
			$rowdata += @(,("Subnet Mask",($htmlsilver -bor $htmlBold),$xIPSubnet[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xIPAddress)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,("IP Address",($htmlsilver -bor $htmlBold),$tmp,$htmlwhite))
					$rowdata += @(,("Subnet Mask",($htmlsilver -bor $htmlBold),$xIPSubnet[$cnt],$htmlwhite))
				}
			}
		}
		Else
		{
			$rowdata += @(,("IP Address",($htmlsilver -bor $htmlBold),$xIPAddress,$htmlwhite))
			If($Nic.Defaultipgateway)
			{
				$rowdata += @(,("Default Gateway",($htmlsilver -bor $htmlBold),$Nic.Defaultipgateway,$htmlwhite))
			}
			$rowdata += @(,("Subnet Mask",($htmlsilver -bor $htmlBold),$xIPSubnet,$htmlwhite))
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$rowdata += @(,('DHCP Enabled',($htmlsilver -bor $htmlBold),$Nic.dhcpenabled,$htmlwhite))
			$rowdata += @(,('DHCP Lease Obtained',($htmlsilver -bor $htmlBold),$dhcpleaseobtaineddate,$htmlwhite))
			$rowdata += @(,('DHCP Lease Expires',($htmlsilver -bor $htmlBold),$dhcpleaseexpiresdate,$htmlwhite))
			$rowdata += @(,('DHCP Server',($htmlsilver -bor $htmlBold),$Nic.dhcpserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$rowdata += @(,('DNS Domain',($htmlsilver -bor $htmlBold),$Nic.dnsdomain,$htmlwhite))
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Search Suffixes',($htmlsilver -bor $htmlBold),$xnicdnsdomainsuffixsearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlBold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('DNS WINS Enabled',($htmlsilver -bor $htmlBold),$xdnsenabledforwinsresolution,$htmlwhite))
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Servers',($htmlsilver -bor $htmlBold),$xnicdnsserversearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlBold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('NetBIOS Setting',($htmlsilver -bor $htmlBold),$xTcpipNetbiosOptions,$htmlwhite))
		$rowdata += @(,('WINS: Enabled LMHosts',($htmlsilver -bor $htmlBold),$xwinsenablelmhostslookup,$htmlwhite))
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$rowdata += @(,('Host Lookup File',($htmlsilver -bor $htmlBold),$Nic.winshostlookupfile,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$rowdata += @(,('Primary Server',($htmlsilver -bor $htmlBold),$Nic.winsprimaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$rowdata += @(,('Secondary Server',($htmlsilver -bor $htmlBold),$Nic.winssecondaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$rowdata += @(,('Scope ID',($htmlsilver -bor $htmlBold),$Nic.winsscopeid,$htmlwhite))
		}

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths -tablewidth "350"
		WriteHTMLLine 0 0 " "
	}
}
#endregion

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. Smith
	
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
			'fr-'	{ 'Table automatique 2'; Break } #changed 10-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
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
		{$CatalanArray -contains $_}	{$CultureCode = "ca-"}
		{$ChineseArray -contains $_}	{$CultureCode = "zh-"}
		{$DanishArray -contains $_}		{$CultureCode = "da-"}
		{$DutchArray -contains $_}		{$CultureCode = "nl-"}
		{$EnglishArray -contains $_}	{$CultureCode = "en-"}
		{$FinnishArray -contains $_}	{$CultureCode = "fi-"}
		{$FrenchArray -contains $_}		{$CultureCode = "fr-"}
		{$GermanArray -contains $_}		{$CultureCode = "de-"}
		{$NorwegianArray -contains $_}	{$CultureCode = "nb-"}
		{$PortugueseArray -contains $_}	{$CultureCode = "pt-"}
		{$SpanishArray -contains $_}	{$CultureCode = "es-"}
		{$SwedishArray -contains $_}	{$CultureCode = "sv-"}
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
		
		If(($MSWord -eq $False) -and ($PDF -eq $True))
		{
			Write-Host "`n`n`t`tThis script uses Microsoft Word's SaveAs PDF function, please install Microsoft Word`n`n"
			Exit
		}
		Else
		{
			Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
			Exit
		}
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
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
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
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
	#Return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function SetupWord
{
	Write-Verbose "$(Get-Date -Format G): Setting up Word"
    
	If(!$AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName).docx"
		If($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName).pdf"
		}
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:WordFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
		If($PDF)
		{
			[string]$Script:PDFFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
		}
	}

	# Setup word for output
	Write-Verbose "$(Get-Date -Format G): Create Word comObject."
	$Script:Word = New-Object -comobject "Word.Application" -EA 0 4>$Null
	
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created. You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
		`t`t
		The Word object could not be created. You may need to repair your Word installation.
		`n`n
		`t`t
		Script cannot Continue.
		`n`n"
		Exit
	}

	Write-Verbose "$(Get-Date -Format G): Determine Word language value"
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
		Unable to determine the Word language value. You may need to repair your Word installation.
		`n`n
		`t`t
		Script cannot Continue.
		`n`n
		"
		AbortScript
	}
	Write-Verbose "$(Get-Date -Format G): Word language value is $($Script:WordLanguageValue)"
	
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
		Microsoft Word 2007 is no longer supported.`n`n`t`tScript will end.
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
		Script cannot Continue.
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
		`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com
		`n`n
		"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($CompanyName))
	{
		Write-Verbose "$(Get-Date -Format G): Company name is blank. Retrieve company name from registry."
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Host "
		Company Name is blank so Cover Page will not show a Company Name.
		Check HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value.
		You may want to use the -CompanyName parameter if you need a Company Name on the cover page.
			" -Foreground White
			$Script:CoName = $TmpName
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Verbose "$(Get-Date -Format G): Updated company name to $($Script:CoName)"
		}
	}
	Else
	{
		$Script:CoName = $CompanyName
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Verbose "$(Get-Date -Format G): Check Default Cover Page for $($WordCultureCode)"
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
			Write-Verbose "$(Get-Date -Format G): Changed Default Cover Page from Sideline to $($CoverPage)"
		}
	}

	Write-Verbose "$(Get-Date -Format G): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)"
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Verbose "$(Get-Date -Format G): Word language value $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date -Format G): Culture code $($Script:WordCultureCode)"
		Write-Error "
		`n`n
		`t`t
		For $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.
		`n`n
		`t`t
		Script cannot Continue.
		`n`n
		"
		AbortScript
	}

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Verbose "$(Get-Date -Format G): Load Word Templates"

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where-Object{$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Verbose "$(Get-Date -Format G): Attempt to load cover page $($CoverPage)"
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach-Object {
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
		Write-Verbose "$(Get-Date -Format G): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Host "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist." -Foreground White
		Write-Host "This report will not have a Cover Page." -Foreground White
	}

	Write-Verbose "$(Get-Date -Format G): Create empty word doc"
	$Script:Doc = $Script:Word.Documents.Add()
	If($Null -eq $Script:Doc)
	{
		Write-Verbose "$(Get-Date -Format G): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	An empty Word document could not be created. You may need to repair your Word installation.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Verbose "$(Get-Date -Format G): "
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "
		`n`n
	An unknown error happened selecting the entire Word document for default formatting options.
		`n`n
	Script cannot Continue.
		`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Verbose "$(Get-Date -Format G): Disable grammar and spell checking"
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Verbose "$(Get-Date -Format G): Insert new page, getting ready for table of contents"
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Verbose "$(Get-Date -Format G): Table of Contents - $($Script:MyHash.Word_TableOfContents)"
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($Null -eq $toc)
		{
			Write-Verbose "$(Get-Date -Format G): "
			Write-Host "Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved." -Foreground White
			Write-Host "This report will not have a Table of Contents." -Foreground White
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Host "Table of Contents are not installed." -Foreground White
		Write-Host "Table of Contents are not installed so this report will not have a Table of Contents." -Foreground White
	}

	#set the footer
	Write-Verbose "$(Get-Date -Format G): Set the footer"
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Verbose "$(Get-Date -Format G): Get the footer and format font"
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
	Write-Verbose "$(Get-Date -Format G): Footer text"
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Verbose "$(Get-Date -Format G): Add page numbering"
	$Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	Write-Verbose "$(Get-Date -Format G):"
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
			Write-Verbose "$(Get-Date -Format G): Set Cover Page Properties"
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

			Write-Verbose "$(Get-Date -Format G): Update the Table of Contents"
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}
#endregion

#region general script functions
Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Verbose "$(Get-Date -Format G): Save and Close document and Shutdown Word"
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following link
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date -Format G): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:WordFileName, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Saving as DOCX file first before saving to PDF"
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Saving DOCX file"
		}
		Write-Verbose "$(Get-Date -Format G): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)"
		$Script:Doc.SaveAs2([REF]$Script:WordFileName, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Verbose "$(Get-Date -Format G): Now saving as PDF"
			$Script:Doc.SaveAs([REF]$Script:PDFFileName, [ref]$wdFormatPDF)
		}
	}

	Write-Verbose "$(Get-Date -Format G): Closing Word"
	$Script:Doc.Close()
	$Script:Word.Quit()
	Write-Verbose "$(Get-Date -Format G): System Cleanup"
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
	$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}).Id
	If($null -ne $wordprocess -and $wordprocess -gt 0)
	{
		Write-Verbose "$(Get-Date -Format G): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess)"
		Stop-Process $wordprocess -EA 0
	}
}

Function SetupText
{
	Write-Verbose "$(Get-Date -Format G): Setting up Text"

	[System.Text.StringBuilder] $global:Output = New-Object System.Text.StringBuilder( 16384 )

	If(!$AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName).txt"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:TextFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}
}

Function SaveandCloseTextDocument
{
	Write-Verbose "$(Get-Date -Format G): Saving Text file"
	Line 0 ""
	Line 0 "Report Complete"
	Write-Output $global:Output.ToString() | Out-File $Script:TextFileName 4>$Null
}

Function SaveandCloseHTMLDocument
{
	Write-Verbose "$(Get-Date -Format G): Saving HTML file"
	WriteHTMLLine 0 0 ""
	WriteHTMLLine 0 0 "Report Complete"
	Out-File -FilePath $Script:HTMLFileName -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFilenames
{
	Param([string]$OutputFileName)
	
	If($MSWord -or $PDF)
	{
		CheckWordPreReq
		
		SetupWord
	}
	If($Text)
	{
		SetupText
	}
	If($HTML)
	{
		SetupHTML
	}
	ShowScriptOptions
}

Function OutputReportFooter
{
	#Added in 2.04
	<#
	Report Footer
		Report information:
			Created with: <Script Name> - Release Date: <Script Release Date>
			Script version: <Script Version>
			Started on <Date Time in Local Format>
			Elapsed time: nn days, nn hours, nn minutes, nn.nn seconds
			Ran from domain <Domain Name> by user <Username>
			Ran from the folder <Folder Name>

	Script Name and Script Release date are script-specific variables.
	Script version is a script variable.
	Start Date Time in Local Format is a script variable.
	Domain Name is $env:USERDNSDOMAIN.
	Username is $env:USERNAME.
	Folder Name is a script variable.
	#>

	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)

	If($MSWORD -or $PDF)
	{
		$Script:selection.InsertNewPage()
		WriteWordLine 1 0 "Report Footer"
		WriteWordLine 2 0 "Report Information:"
		WriteWordLine 0 1 "Created with: $Script:ScriptName - Release Date: $Script:ReleaseDate"
		WriteWordLine 0 1 "Script version: $Script:MyVersion"
		WriteWordLine 0 1 "Started on $Script:StartTime"
		WriteWordLine 0 1 "Elapsed time: $Str"
		WriteWordLine 0 1 "Ran from domain $env:USERDNSDOMAIN by user $env:USERNAME"
		WriteWordLine 0 1 "Ran from the folder $Script:pwdpath"
	}
	If($Text)
	{
		Line 0 "///  Report Footer  \\\"
		Line 1 "Report Information:"
		Line 2 "Created with: $Script:ScriptName - Release Date: $Script:ReleaseDate"
		Line 2 "Script version: $Script:MyVersion"
		Line 2 "Started on $Script:StartTime"
		Line 2 "Elapsed time: $Str"
		Line 2 "Ran from domain $env:USERDNSDOMAIN by user $env:USERNAME"
		Line 2 "Ran from the folder $Script:pwdpath"
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "///&nbsp;&nbsp;Report Footer&nbsp;&nbsp;\\\"
		WriteHTMLLine 2 0 "Report Information:"
		WriteHTMLLine 0 1 "Created with: $Script:ScriptName - Release Date: $Script:ReleaseDate"
		WriteHTMLLine 0 1 "Script version: $Script:MyVersion"
		WriteHTMLLine 0 1 "Started on $Script:StartTime"
		WriteHTMLLine 0 1 "Elapsed time: $Str"
		WriteHTMLLine 0 1 "Ran from domain $env:USERDNSDOMAIN by user $env:USERNAME"
		WriteHTMLLine 0 1 "Ran from the folder $Script:pwdpath"
	}
}

Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	If($Text)
	{
		SaveandCloseTextDocument
	}
	If($HTML)
	{
		SaveandCloseHTMLDocument
	}

	$GotFile = $False

	If($MSWord)
	{
		If(Test-Path "$($Script:WordFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:WordFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date -Format G): Unable to save the output file, $($Script:WordFileName)"
			Write-Error "Unable to save the output file, $($Script:WordFileName)"
		}
	}
	If($PDF)
	{
		If(Test-Path "$($Script:PDFFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:PDFFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date -Format G): Unable to save the output file, $($Script:PDFFileName)"
			Write-Error "Unable to save the output file, $($Script:PDFFileName)"
		}
	}
	If($Text)
	{
		If(Test-Path "$($Script:TextFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:TextFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date -Format G): Unable to save the output file, $($Script:TextFileName)"
			Write-Error "Unable to save the output file, $($Script:TextFileName)"
		}
	}
	If($HTML)
	{
		If(Test-Path "$($Script:HTMLFileName)")
		{
			Write-Verbose "$(Get-Date -Format G): $($Script:HTMLFileName) is ready for use"
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date -Format G): Unable to save the output file, $($Script:HTMLFileName)"
			Write-Error "Unable to save the output file, $($Script:HTMLFileName)"
		}
	}
	
	#email output file if requested
	If($GotFile -and ![System.String]::IsNullOrEmpty( $SmtpServer ))
	{
		$emailattachments = @()
		If($MSWord)
		{
			$emailAttachments += $Script:WordFileName
		}
		If($PDF)
		{
			$emailAttachments += $Script:PDFFileName
		}
		If($Text)
		{
			$emailAttachments += $Script:TextFileName
		}
		If($HTML)
		{
			$emailAttachments += $Script:HTMLFileName
		}
		SendEmail $emailAttachments
	}
}

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): AddDateTime     : $AddDateTime"
	If($MSWord -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): Company Name    : $Script:CoName"
		Write-Verbose "$(Get-Date -Format G): Company Address : $CompanyAddress"
		Write-Verbose "$(Get-Date -Format G): Company Email   : $CompanyEmail"
		Write-Verbose "$(Get-Date -Format G): Company Fax     : $CompanyFax"
		Write-Verbose "$(Get-Date -Format G): Company Phone   : $CompanyPhone"
		Write-Verbose "$(Get-Date -Format G): Cover Page      : $CoverPage"
	}
	Write-Verbose "$(Get-Date -Format G): ComputerName    : $ComputerName"
	Write-Verbose "$(Get-Date -Format G): Dev             : $Dev"
	If($Dev)
	{
		Write-Verbose "$(Get-Date -Format G): DevErrorFile    : $Script:DevErrorFile"
	}
	If($MSWord)
	{
		Write-Verbose "$(Get-Date -Format G): Word FileName   : $Script:WordFileName"
	}
	If($HTML)
	{
		Write-Verbose "$(Get-Date -Format G): HTML FileName   : $Script:HTMLFileName"
	} 
	If($PDF)
	{
		Write-Verbose "$(Get-Date -Format G): PDF FileName    : $Script:PDFFileName"
	}
	If($Text)
	{
		Write-Verbose "$(Get-Date -Format G): Text FileName   : $Script:TextFileName"
	}
	Write-Verbose "$(Get-Date -Format G): Folder          : $Folder"
	Write-Verbose "$(Get-Date -Format G): From            : $From"
	Write-Verbose "$(Get-Date -Format G): HW Inventory    : $Hardware"
	Write-Verbose "$(Get-Date -Format G): Include Leases  : $IncludeLeases"
	Write-Verbose "$(Get-Date -Format G): Include Options : $IncludeOptions"
	Write-Verbose "$(Get-Date -Format G): Log             : $Log"
	Write-Verbose "$(Get-Date -Format G): Report Footer   : $ReportFooter"
	Write-Verbose "$(Get-Date -Format G): Save As HTML    : $HTML"
	Write-Verbose "$(Get-Date -Format G): Save As PDF     : $PDF"
	Write-Verbose "$(Get-Date -Format G): Save As TEXT    : $TEXT"
	Write-Verbose "$(Get-Date -Format G): Save As WORD    : $MSWORD"
	Write-Verbose "$(Get-Date -Format G): ScriptInfo      : $ScriptInfo"
	Write-Verbose "$(Get-Date -Format G): Smtp Port       : $SmtpPort"
	Write-Verbose "$(Get-Date -Format G): Smtp Server     : $SmtpServer"
	Write-Verbose "$(Get-Date -Format G): Title           : $Script:Title"
	Write-Verbose "$(Get-Date -Format G): To              : $To"
	Write-Verbose "$(Get-Date -Format G): Use SSL         : $UseSSL"
	Write-Verbose "$(Get-Date -Format G): User Name       : $UserName"
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): OS Detected     : $Script:RunningOS"
	Write-Verbose "$(Get-Date -Format G): PoSH version    : $Host.Version"
	Write-Verbose "$(Get-Date -Format G): PSCulture       : $PSCulture"
	Write-Verbose "$(Get-Date -Format G): PSUICulture     : $PSUICulture"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date -Format G): Word language   : $Script:WordLanguageValue"
		Write-Verbose "$(Get-Date -Format G): Word version    : $Script:WordProduct"
	}
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): Script start    : $Script:StartTime"
	Write-Verbose "$(Get-Date -Format G): "
	Write-Verbose "$(Get-Date -Format G): "

}

Function AbortScript
{
	If($MSWord -or $PDF)
	{
		$Script:Word.quit()
		Write-Verbose "$(Get-Date -Format G): System Cleanup"
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
		If(Test-Path variable:global:word)
		{
			Remove-Variable -Name word -Scope Global 4>$Null
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date -Format G): Script has been aborted"
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Function OutputWarning
{
	Param([string] $txt)
	Write-Warning $txt
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 $txt
		WriteWordLIne 0 0 ""
	}
	If($Text)
	{
		Line 1 $txt
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 0 1 $txt
		WriteHTMLLine 0 0 " "
	}
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If( ( Get-Member -Name $topLevel -InputObject $object ) )
		{
			If( ( Get-Member -Name $secondLevel -InputObject $object.$topLevel ) )
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

Function InsertBlankLine
{
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 0 ""
	}
	If($Text)
	{
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 0 0 " "
	}
}

Function TestComputerName
{
	Param([string]$Cname)
	If(![String]::IsNullOrEmpty($CName)) 
	{
		#get computer name
		#first test to make sure the computer is reachable
		Write-Verbose "$(Get-Date -Format G): Testing to see if $($CName) is online and reachable"
		If(Test-Connection -ComputerName $CName -quiet)
		{
			Write-Verbose "$(Get-Date -Format G): Server $($CName) is online."
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Computer $($CName) is offline"
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "
			`n`n
			`t`t
			Computer $($CName) is offline.
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
			Exit
		}
	}

	#if computer name is localhost, get actual computer name
	If($CName -eq "localhost")
	{
		$CName = $env:ComputerName
		Write-Verbose "$(Get-Date -Format G): Computer name has been renamed from localhost to $($CName)"
		Write-Verbose "$(Get-Date -Format G): Testing to see if $($CName) is a DHCP Server"
		$results = Get-DHCPServerVersion -ComputerName $CName -EA 0
		If($? -and $Null -ne $results)
		{
			#the computer is a dhcp server
			Write-Verbose "$(Get-Date -Format G): Computer $($CName) is a DHCP Server"
			Return $CName
		}
		ElseIf(!$? -or $Null -eq $results)
		{
			#the computer is not a dhcp server
			Write-Verbose "$(Get-Date -Format G): Computer $($CName) is not a DHCP Server"
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "
			`n`n
			`t`t
			Computer $($CName) is not a DHCP Server.
			`n`n
			`t`t
			Rerun the script using -ComputerName with a valid DHCP server name.
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
			Exit
		}
	}

	#if computer name is an IP address, get host name from DNS
	#http://blogs.technet.com/b/gary/archive/2009/08/29/resolve-ip-addresses-to-hostname-using-powershell.aspx
	#help from Michael B. Smith
	$ip = $CName -as [System.Net.IpAddress]
	If($ip)
	{
		$Result = [System.Net.Dns]::gethostentry($ip)
		
		If($? -and $Null -ne $Result)
		{
			$CName = $Result.HostName
			Write-Verbose "$(Get-Date -Format G): Computer name has been renamed from $($ip) to $($CName)"
			Write-Verbose "$(Get-Date -Format G): Testing to see if $($CName) is a DHCP Server"
			$results = Get-DHCPServerVersion -ComputerName $CName -EA 0
			If($? -and $Null -ne $results)
			{
				#the computer is a dhcp server
				Write-Verbose "$(Get-Date -Format G): Computer $($CName) is a DHCP Server"
				Return $CName
			}
			ElseIf(!$? -or $Null -eq $results)
			{
				#the computer is not a dhcp server
				Write-Verbose "$(Get-Date -Format G): Computer $($CName) is not a DHCP Server"
				$ErrorActionPreference = $SaveEAPreference
				Write-Error "
				`n`n
				`t`t
				Computer $($CName) is not a DHCP Server.
				`n`n
				`t`t
				Rerun the script using -ComputerName with a valid DHCP server name.
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
			Write-Warning "Unable to resolve $($CName) to a hostname"
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): Testing to see if $($CName) is a DHCP Server"
		$results = Get-DHCPServerVersion -ComputerName $CName -EA 0
		If($? -and $Null -ne $results)
		{
			#the computer is a dhcp server
			Write-Verbose "$(Get-Date -Format G): Computer $($CName) is a DHCP Server"
			Return $CName
		}
		ElseIf(!$? -or $Null -eq $results)
		{
			#the computer is not a dhcp server
			Write-Verbose "$(Get-Date -Format G): Computer $($CName) is not a DHCP Server"
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "
			`n`n
			`t`t
			Computer $($CName) is not a DHCP Server.
			`n`n
			`t`t
			Rerun the script using -ComputerName with a valid DHCP server name.
			`n`n
			`t`t
			Script cannot continue.
			`n`n
			"
			Exit
		}
	}
	Return $CName
}

Function TestComputerName2
{
	Param([string]$Cname)
	
	If(![String]::IsNullOrEmpty($CName)) 
	{
		#get computer name
		#first test to make sure the computer is reachable
		Write-Verbose "$(Get-Date -Format G): Testing to see if $($CName) is online and reachable"
		If(Test-Connection -ComputerName $CName -quiet)
		{
			Write-Verbose "$(Get-Date -Format G): Server $($CName) is online."
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Computer $($CName) is offline"
			Write-Output "$(Get-Date -Format G): Computer $($CName) is offline" | Out-File $Script:BadDHCPErrorFile -Append 4>$Null
			Return "BAD"
		}
	}

	#if computer name is localhost, get actual computer name
	If($CName -eq "localhost")
	{
		$CName = $env:ComputerName
		Write-Verbose "$(Get-Date -Format G): Computer name has been renamed from localhost to $($CName)"
		Write-Verbose "$(Get-Date -Format G): Testing to see if $($CName) is a DHCP Server"
		$results = Get-DHCPServerVersion -ComputerName $CName -EA 0
		If($? -and $Null -ne $results)
		{
			#the computer is a dhcp server
			Write-Verbose "$(Get-Date -Format G): Computer $($CName) is a DHCP Server"
			Return $CName
		}
		ElseIf(!$? -or $Null -eq $results)
		{
			#the computer is not a dhcp server
			Write-Verbose "$(Get-Date -Format G): Computer $($CName) is not a DHCP Server"
			Write-Output "$(Get-Date -Format G): Computer $($CName) is not a DHCP Server" | Out-File $Script:BadDHCPErrorFile -Append 4>$Null
			Return "BAD"
		}
	}

	#if computer name is an IP address, get host name from DNS
	#http://blogs.technet.com/b/gary/archive/2009/08/29/resolve-ip-addresses-to-hostname-using-powershell.aspx
	#help from Michael B. Smith
	$ip = $CName -as [System.Net.IpAddress]
	If($ip)
	{
		$Result = [System.Net.Dns]::gethostentry($ip)
		
		If($? -and $Null -ne $Result)
		{
			$CName = $Result.HostName
			Write-Verbose "$(Get-Date -Format G): Computer name has been renamed from $($ip) to $($CName)"
			Write-Verbose "$(Get-Date -Format G): Testing to see if $($CName) is a DHCP Server"
			$results = Get-DHCPServerVersion -ComputerName $CName -EA 0
			If($? -and $Null -ne $results)
			{
				#the computer is a dhcp server
				Write-Verbose "$(Get-Date -Format G): Computer $($CName) is a DHCP Server"
				Return $CName
			}
			ElseIf(!$? -or $Null -eq $results)
			{
				#the computer is not a dhcp server
				Write-Verbose "$(Get-Date -Format G): Computer $($CName) is not a DHCP Server"
				Write-Output "$(Get-Date -Format G): Computer $($CName) is not a DHCP Server" | Out-File $Script:BadDHCPErrorFile -Append 4>$Null
				Return "BAD"
			}
		}
		Else
		{
			Write-Verbose "$(Get-Date -Format G): Unable to resolve $($CName) to a hostname"
			Write-Output "$(Get-Date -Format G): Unable to resolve $($CName) to a hostname" | Out-File $Script:BadDHCPErrorFile -Append 4>$Null
			Return "BAD"
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): Testing to see if $($CName) is a DHCP Server"
		$results = Get-DHCPServerVersion -ComputerName $CName -EA 0
		If($? -and $Null -ne $results)
		{
			#the computer is a dhcp server
			Write-Verbose "$(Get-Date -Format G): Computer $($CName) is a DHCP Server"
			Return $CName
		}
		ElseIf(!$? -or $Null -eq $results)
		{
			#the computer is not a dhcp server
			Write-Verbose "$(Get-Date -Format G): Computer $($CName) is not a DHCP Server"
			Write-Output "$(Get-Date -Format G): Computer $($CName) is not a DHCP Server" | Out-File $Script:BadDHCPErrorFile -Append 4>$Null
			Return "BAD"
		}
	}

	Write-Verbose "$(Get-Date -Format G): "
	Return $CName
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

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
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

	While( $tabs -gt 0 )
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
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing}
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
	This Function formats an HTML line
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
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $Null 0 $htmlBold

	Writes a line omitting font and font size and setting the bold attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $Null 0 ($htmlBold -bor $htmlitalics)

	Writes a line omitting font and font size and setting both italics and bold options

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $Null 2  # 10 point font

	Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 

	Writes a line using Courier New Font and 0 font point size (default = 2 if set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlBold -bor $htmlred -bor $htmlitalics)

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

#V3.00
# to suppress $crlf in HTML documents, replace this with '' (empty string)
# but this was added to make the HTML readable
$crlf = [System.Environment]::NewLine

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
#headings fixed 12-Oct-2016 by Webster
#errors with $HTMLStyle fixed 7-Dec-2017 by Webster
# re-implemented/re-based for v3.00 by Michael B. Smith
{
	Param
	(
		[Int]    $style    = 0, 
		[Int]    $tabs     = 0, 
		[String] $name     = '', 
		[String] $value    = '', 
		[String] $fontName = $null,
		[Int]    $fontSize = 1,
		[Int]    $options  = $htmlblack
	)

	#V3.00 - FIXME - long story short, this Function was wrong and had been wrong for a long time. 
	## The Function generated invalid HTML, and ignored fontname and fontsize parameters. I fixed
	## those items, but that made the report unreadable, because all of the formatting had been based
	## on this Function not working properly.

	## here is a typical H1 previously generated:
	## <h1>///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\<font face='Calibri' color='#000000' size='1'></h1></font>

	## fixing the Function generated this (unreadably small):
	## <h1><font face='Calibri' color='#000000' size='1'>///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\</font></h1>

	## So I took all the fixes out. This routine now generates valid HTML, but the fontName, fontSize,
	## and options parameters are ignored; so the routine generates equivalent output as before. I took
	## the fixes out instead of fixing all the call sites, because there are 225 call sites! If you are
	## willing to update all the call sites, you can easily re-instate the fixes. They have only been
	## commented out with '##' below.

	## If( [String]::IsNullOrEmpty( $fontName ) )
	## {
	##	$fontName = 'Calibri'
	## }
	## If( $fontSize -le 0 )
	## {
	##	$fontSize = 1
	## }

	## ## output data is stored here
	## [String] $output = ''
	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 1024 )

	If( [String]::IsNullOrEmpty( $name ) )	
	{
		## $HTMLBody = '<p></p>'
		$null = $sb.Append( '<p></p>' )
	}
	Else
	{
		## #V3.00
		[Bool] $ital = $options -band $htmlitalics
		[Bool] $bold = $options -band $htmlBold
		## $color = $global:htmlColor[ $options -band 0xffffc ]

		## ## build the HTML output string
##		$HTMLBody = ''
##		If( $ital ) { $HTMLBody += '<i>' }
##		If( $bold ) { $HTMLBody += '<b>' } 
		If( $ital ) { $null = $sb.Append( '<i>' ) }
		If( $bold ) { $null = $sb.Append( '<b>' ) } 

		Switch( $style )
		{
			1 { $HTMLOpen = '<h1>'; $HTMLClose = '</h1>'; Break }
			2 { $HTMLOpen = '<h2>'; $HTMLClose = '</h2>'; Break }
			3 { $HTMLOpen = '<h3>'; $HTMLClose = '</h3>'; Break }
			4 { $HTMLOpen = '<h4>'; $HTMLClose = '</h4>'; Break }
			Default { $HTMLOpen = ''; $HTMLClose = ''; Break }
		}

		## $HTMLBody += $HTMLOpen
		$null = $sb.Append( $HTMLOpen )

		## If($HTMLClose -eq '')
		## {
		##	$HTMLBody += "<br><font face='" + $fontName + "' " + "color='" + $color + "' size='"  + $fontSize + "'>"
		## }
		## Else
		## {
		##	$HTMLBody += "<font face='" + $fontName + "' " + "color='" + $color + "' size='"  + $fontSize + "'>"
		## }
		
##		While( $tabs -gt 0 )
##		{ 
##			$output += '&nbsp;&nbsp;&nbsp;&nbsp;'
##			$tabs--
##		}
		## output the rest of the parameters.
##		$output += $name + $value
		## $HTMLBody += $output
		$null = $sb.Append( ( '&nbsp;&nbsp;&nbsp;&nbsp;' * $tabs ) + $name + $value )

		## $HTMLBody += '</font>'
##		If( $HTMLClose -eq '' ) { $HTMLBody += '<br>'     }
##		Else                    { $HTMLBody += $HTMLClose }

##		If( $ital ) { $HTMLBody += '</i>' }
##		If( $bold ) { $HTMLBody += '</b>' } 

##		If( $HTMLClose -eq '' ) { $HTMLBody += '<br />' }

		If( $HTMLClose -eq '' ) { $null = $sb.Append( '<br>' )     }
		Else                    { $null = $sb.Append( $HTMLClose ) }

		If( $ital ) { $null = $sb.Append( '</i>' ) }
		If( $bold ) { $null = $sb.Append( '</b>' ) } 

		If( $HTMLClose -eq '' ) { $null = $sb.Append( '<br />' ) }
	}
	##$HTMLBody += $crlf
	$null = $sb.AppendLine( '' )

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $sb.ToString() 4>$Null
}
#endregion

#region HTML table functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable Function
# Created by Ken Avram
# modified by Jake Rutski
# re-implemented by Michael B. Smith for v2.00. Also made the documentation match reality.
#***********************************************************************************************************
Function AddHTMLTable
{
	Param
	(
		[String]   $fontName  = 'Calibri',
		[Int]      $fontSize  = 2,
		[Int]      $colCount  = 0,
		[Int]      $rowCount  = 0,
		[Object[]] $rowInfo   = $null,
		[Object[]] $fixedInfo = $null
	)
	#V3.00 - Use StringBuilder - MBS
	## In the normal case, tables are only a few dozen cells. But in the case
	## of Sites, OUs, and Users - there may be many hundreds of thousands of 
	## cells. Using normal strings is too slow.

	#V3.00
	## If( $ExtraSpecialVerbose )
	## {
	##	$global:rowInfo1 = $rowInfo
	## }
<#
	If( $SuperVerbose )
	{
		wv "AddHTMLTable: fontName '$fontName', fontsize $fontSize, colCount $colCount, rowCount $rowCount"
		If( $null -ne $rowInfo -and $rowInfo.Count -gt 0 )
		{
			wv "AddHTMLTable: rowInfo has $( $rowInfo.Count ) elements"
			If( $ExtraSpecialVerbose )
			{
				wv "AddHTMLTable: rowInfo length $( $rowInfo.Length )"
				For( $ii = 0; $ii -lt $rowInfo.Length; $ii++ )
				{
					$row = $rowInfo[ $ii ]
					wv "AddHTMLTable: index $ii, type $( $row.GetType().FullName ), length $( $row.Length )"
					For( $yyy = 0; $yyy -lt $row.Length; $yyy++ )
					{
						wv "AddHTMLTable: index $ii, yyy = $yyy, val = '$( $row[ $yyy ] )'"
					}
					wv "AddHTMLTable: done"
				}
			}
		}
		Else
		{
			wv "AddHTMLTable: rowInfo is empty"
		}
		If( $null -ne $fixedInfo -and $fixedInfo.Count -gt 0 )
		{
			wv "AddHTMLTable: fixedInfo has $( $fixedInfo.Count ) elements"
		}
		Else
		{
			wv "AddHTMLTable: fixedInfo is empty"
		}
	}
#>

	$fwLength = If( $null -ne $fixedInfo ) { $fixedInfo.Count } else { 0 }

	##$htmlbody = ''
	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 8192 )

	If( $rowInfo -and $rowInfo.Length -lt $rowCount )
	{
##		$oldCount = $rowCount
		$rowCount = $rowInfo.Length
##		If( $SuperVerbose )
##		{
##			wv "AddHTMLTable: updated rowCount to $rowCount from $oldCount, based on rowInfo.Length"
##		}
	}

	For( $rowCountIndex = 0; $rowCountIndex -lt $rowCount; $rowCountIndex++ )
	{
		$null = $sb.AppendLine( '<tr>' )
		## $htmlbody += '<tr>'
		## $htmlbody += $crlf #V3.00 - make the HTML readable

		## each row of rowInfo is an array
		## each row consists of tuples: an item of text followed by an item of formatting data
<#		
		$row = $rowInfo[ $rowCountIndex ]
		If( $ExtraSpecialVerbose )
		{
			wv "!!!!! AddHTMLTable: rowCountIndex = $rowCountIndex, row.Length = $( $row.Length ), row gettype = $( $row.GetType().FullName )"
			wv "!!!!! AddHTMLTable: colCount $colCount"
			wv "!!!!! AddHTMLTable: row[0].Length $( $row[0].Length )"
			wv "!!!!! AddHTMLTable: row[0].GetType $( $row[0].GetType().FullName )"
			$subRow = $row
			If( $subRow -is [Array] -and $subRow[ 0 ] -is [Array] )
			{
				$subRow = $subRow[ 0 ]
				wv "!!!!! AddHTMLTable: deref subRow.Length $( $subRow.Length ), subRow.GetType $( $subRow.GetType().FullName )"
			}

			For( $columnIndex = 0; $columnIndex -lt $subRow.Length; $columnIndex += 2 )
			{
				$item = $subRow[ $columnIndex ]
				wv "!!!!! AddHTMLTable: item.GetType $( $item.GetType().FullName )"
				## If( !( $item -is [String] ) -and $item -is [Array] )
##				If( $item -is [Array] -and $item[ 0 ] -is [Array] )				
##				{
##					$item = $item[ 0 ]
##					wv "!!!!! AddHTMLTable: dereferenced item.GetType $( $item.GetType().FullName )"
##				}
				wv "!!!!! AddHTMLTable: rowCountIndex = $rowCountIndex, columnIndex = $columnIndex, val '$item'"
			}
			wv "!!!!! AddHTMLTable: done"
		}
#>

		## reset
		$row = $rowInfo[ $rowCountIndex ]

		$subRow = $row
		If( $subRow -is [Array] -and $subRow[ 0 ] -is [Array] )
		{
			$subRow = $subRow[ 0 ]
			## wv "***** AddHTMLTable: deref rowCountIndex $rowCountIndex, subRow.Length $( $subRow.Length ), subRow.GetType $( $subRow.GetType().FullName )"
		}

		$subRowLength = $subRow.Count
		For( $columnIndex = 0; $columnIndex -lt $colCount; $columnIndex += 2 )
		{
			$item = If( $columnIndex -lt $subRowLength ) { $subRow[ $columnIndex ] } Else { 0 }
			## If( !( $item -is [String] ) -and $item -is [Array] )
##			If( $item -is [Array] -and $item[ 0 ] -is [Array] )
##			{
##				$item = $item[ 0 ]
##			}

			$text   = If( $item ) { $item.ToString() } Else { '' }
			$format = If( ( $columnIndex + 1 ) -lt $subRowLength ) { $subRow[ $columnIndex + 1 ] } Else { 0 }
			## item, text, and format ALWAYS have values, even if empty values
			$color  = $global:htmlColor[ $format -band 0xffffc ]
			[Bool] $bold = $format -band $htmlBold
			[Bool] $ital = $format -band $htmlitalics
<#			
			If( $ExtraSpecialVerbose )
			{
				wv "***** columnIndex $columnIndex, subRow.Length $( $subRow.Length ), item GetType $( $item.GetType().FullName ), item '$item'"
				wv "***** format $format, color $color, text '$text'"
				wv "***** format gettype $( $format.GetType().Fullname ), text gettype $( $text.GetType().Fullname )"
			}
#>

			If( $fwLength -eq 0 )
			{
				$null = $sb.Append( "<td style=""background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>" )
				##$htmlbody += "<td style=""background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>"
			}
			Else
			{
				$null = $sb.Append( "<td style=""width:$( $fixedInfo[ $columnIndex / 2 ] ); background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>" )
				##$htmlbody += "<td style=""width:$( $fixedInfo[ $columnIndex / 2 ] ); background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>"
			}

			##If( $bold ) { $htmlbody += '<b>' }
			##If( $ital ) { $htmlbody += '<i>' }
			If( $bold ) { $null = $sb.Append( '<b>' ) }
			If( $ital ) { $null = $sb.Append( '<i>' ) }

			If( $text -eq ' ' -or $text.length -eq 0)
			{
				##$htmlbody += '&nbsp;&nbsp;&nbsp;'
				$null = $sb.Append( '&nbsp;&nbsp;&nbsp;' )
			}
			Else
			{
				For($inx = 0; $inx -lt $text.length; $inx++ )
				{
					If( $text[ $inx ] -eq ' ' )
					{
						##$htmlbody += '&nbsp;'
						$null = $sb.Append( '&nbsp;' )
					}
					Else
					{
						Break
					}
				}
				##$htmlbody += $text
				$null = $sb.Append( $text )
			}

##			If( $bold ) { $htmlbody += '</b>' }
##			If( $ital ) { $htmlbody += '</i>' }
			If( $bold ) { $null = $sb.Append( '</b>' ) }
			If( $ital ) { $null = $sb.Append( '</i>' ) }

			$null = $sb.AppendLine( '</font></td>' )
##			$htmlbody += '</font></td>'
##			$htmlbody += $crlf
		}

		$null = $sb.AppendLine( '</tr>' )
##		$htmlbody += '</tr>'
##		$htmlbody += $crlf
	}

##	If( $ExtraSpecialVerbose )
##	{
##		$global:rowInfo = $rowInfo
##		wv "!!!!! AddHTMLTable: HTML = '$htmlbody'"
##	}

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $sb.ToString() 4>$Null 
}

#***********************************************************************************************************
# FormatHTMLTable 
# Created by Ken Avram
# modified by Jake Rutski
# reworked by Michael B. Smith for v2.23
#***********************************************************************************************************

<#
.Synopsis
	Format table for a HTML output document.
.DESCRIPTION
	This function formats a table for HTML from multiple arrays of strings.
.PARAMETER noBorder
	If set to $true, a table will be generated without a border (border = '0'). Otherwise the table will be generated
	with a border (border = '1').
.PARAMETER noHeadCols
	This parameter should be used when generating tables which do not have a separate array containing column headers
	(columnArray is not specified). Set this parameter equal to the number of columns in the table.
.PARAMETER rowArray
	This parameter contains the row data array for the table.
.PARAMETER columnArray
	This parameter contains column header data for the table.
.PARAMETER fixedWidth
	This parameter contains widths for columns in pixel format ("100px") to override auto column widths
	The variable should contain a width for each column you wish to override the auto-size setting
	For example: $fixedWidth = @("100px","110px","120px","130px","140px")
.PARAMETER tableHeader
	A string containing the header for the table (printed at the top of the table, left justified). The
	default is a blank string.
.PARAMETER tableWidth
	The width of the table in pixels, or 'auto'. The default is 'auto'.
.PARAMETER fontName
	The name of the font to use in the table. The default is 'Calibri'.
.PARAMETER fontSize
	The size of the font to use in the table. The default is 2. Note that this is the HTML size, not the pixel size.

.USAGE
	FormatHTMLTable <Table Header> <Table Width> <Font Name> <Font Size>

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

	$columnHeaders = @('Display Name',$htmlsb,'Status',$htmlsb,'Startup Type',$htmlsb)

	The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics.  For the anding, parens are required or it will
	not format correctly.

	This is following by adding rowdata as shown below.  As more columns are added the columns will auto adjust to fit the size of the page.

	$rowdata = @()
	$columnHeaders = @("User Name",$htmlsb,$UserName,$htmlwhite)
	$rowdata += @(,('Save as PDF',$htmlsb,$PDF.ToString(),$htmlwhite))
	$rowdata += @(,('Save as TEXT',$htmlsb,$TEXT.ToString(),$htmlwhite))
	$rowdata += @(,('Save as WORD',$htmlsb,$MSWORD.ToString(),$htmlwhite))
	$rowdata += @(,('Save as HTML',$htmlsb,$HTML.ToString(),$htmlwhite))
	$rowdata += @(,('Add DateTime',$htmlsb,$AddDateTime.ToString(),$htmlwhite))
	$rowdata += @(,('Hardware Inventory',$htmlsb,$Hardware.ToString(),$htmlwhite))
	$rowdata += @(,('Computer Name',$htmlsb,$ComputerName,$htmlwhite))
	$rowdata += @(,('Filename1',$htmlsb,$Script:FileName1,$htmlwhite))
	$rowdata += @(,('OS Detected',$htmlsb,$Script:RunningOS,$htmlwhite))
	$rowdata += @(,('PSUICulture',$htmlsb,$PSCulture,$htmlwhite))
	$rowdata += @(,('PoSH version',$htmlsb,$Host.Version.ToString(),$htmlwhite))
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
	Param
	(
		[String]   $tableheader = '',
		[String]   $tablewidth  = 'auto',
		[String]   $fontName    = 'Calibri',
		[Int]      $fontSize    = 2,
		[Switch]   $noBorder    = $false,
		[Int]      $noHeadCols  = 1,
		[Object[]] $rowArray    = $null,
		[Object[]] $fixedWidth  = $null,
		[Object[]] $columnArray = $null
	)

	## FIXME - the help text for this Function is wacky wrong - MBS
	## FIXME - Use StringBuilder - MBS - this only builds the table header - benefit relatively small
<#
	If( $SuperVerbose )
	{
		wv "FormatHTMLTable: fontname '$fontname', size $fontSize, tableheader '$tableheader'"
		wv "FormatHTMLTable: noborder $noborder, noheadcols $noheadcols"
		If( $rowarray -and $rowarray.count -gt 0 )
		{
			wv "FormatHTMLTable: rowarray has $( $rowarray.count ) elements"
		}
		Else
		{
			wv "FormatHTMLTable: rowarray is empty"
		}
		If( $columnarray -and $columnarray.count -gt 0 )
		{
			wv "FormatHTMLTable: columnarray has $( $columnarray.count ) elements"
		}
		Else
		{
			wv "FormatHTMLTable: columnarray is empty"
		}
		If( $fixedwidth -and $fixedwidth.count -gt 0 )
		{
			wv "FormatHTMLTable: fixedwidth has $( $fixedwidth.count ) elements"
		}
		Else
		{
			wv "FormatHTMLTable: fixedwidth is empty"
		}
	}
#>

	$HTMLBody = ''
	If( $tableheader.Length -gt 0 )
	{
		$HTMLBody += "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>" + $crlf
	}

	$fwSize = If( $null -eq $fixedWidth ) { 0 } else { $fixedWidth.Count }

	If( $null -eq $columnArray -or $columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If( $null -eq $rowArray )
	{
		$NumRows = 1
	}
	Else
	{
		$NumRows = $rowArray.length + 1
	}

	If( $noBorder )
	{
		$HTMLBody += "<table border='0' width='" + $tablewidth + "'>"
	}
	Else
	{
		$HTMLBody += "<table border='1' width='" + $tablewidth + "'>"
	}
	$HTMLBody += $crlf

	If( $columnArray -and $columnArray.Length -gt 0 )
	{
		$HTMLBody += '<tr>' + $crlf

		For( $columnIndex = 0; $columnIndex -lt $NumCols; $columnindex += 2 )
		{
			#V3.00
			$val = $columnArray[ $columnIndex + 1 ]
			$tmp = $global:htmlColor[ $val -band 0xffffc ]
			[Bool] $bold = $val -band $htmlBold
			[Bool] $ital = $val -band $htmlitalics

			If( $fwSize -eq 0 )
			{
				$HTMLBody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$HTMLBody += "<td style=""width:$($fixedWidth[$columnIndex / 2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If( $bold ) { $HTMLBody += '<b>' }
			If( $ital ) { $HTMLBody += '<i>' }

			$array = $columnArray[ $columnIndex ]
			If( $array )
			{
				If( $array -eq ' ' -or $array.Length -eq 0 )
				{
					$HTMLBody += '&nbsp;&nbsp;&nbsp;'
				}
				Else
				{
					For( $i = 0; $i -lt $array.Length; $i += 2 )
					{
						If( $array[ $i ] -eq ' ' )
						{
							$HTMLBody += '&nbsp;'
						}
						Else
						{
							Break
						}
					}
					$HTMLBody += $array
				}
			}
			Else
			{
				$HTMLBody += '&nbsp;&nbsp;&nbsp;'
			}
			
			If( $bold ) { $HTMLBody += '</b>' }
			If( $ital ) { $HTMLBody += '</i>' }

			$HTMLBody += '</font></td>'
			$HTMLBody += $crlf
		}

		$HTMLBody += '</tr>' + $crlf
	}

	#V3.00
	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $HTMLBody 4>$Null 
	$HTMLBody = ''

	##$rowindex = 2
	If( $rowArray )
	{
<#
		If( $ExtraSpecialVerbose )
		{
			wv "***** FormatHTMLTable: rowarray length $( $rowArray.Length )"
			For( $ii = 0; $ii -lt $rowArray.Length; $ii++ )
			{
				$row = $rowArray[ $ii ]
				wv "***** FormatHTMLTable: index $ii, type $( $row.GetType().FullName ), length $( $row.Length )"
				For( $yyy = 0; $yyy -lt $row.Length; $yyy++ )
				{
					wv "***** FormatHTMLTable: index $ii, yyy = $yyy, val = '$( $row[ $yyy ] )'"
				}
				wv "***** done"
			}
			wv "***** FormatHTMLTable: rowCount $NumRows"
		}
#>

		AddHTMLTable -fontName $fontName -fontSize $fontSize `
			-colCount $numCols -rowCount $NumRows `
			-rowInfo $rowArray -fixedInfo $fixedWidth
		##$rowArray = @()
		$rowArray = $null
		$HTMLBody = '</table>'
	}
	Else
	{
		$HTMLBody += '</table>'
	}

	Out-File -FilePath $Script:HTMLFileName -Append -InputObject $HTMLBody 4>$Null 
}
#endregion

#region other HTML functions
<#
#***********************************************************************************************************
# CheckHTMLColor - Called from AddHTMLTable WriteHTMLLine and FormatHTMLTable
#***********************************************************************************************************
Function CheckHTMLColor
{
	Param($hash)

	#V2.23 -- this is really slow. several ways to fixit. so fixit. MBS
	#V2.23 - obsolete. replaced by using $global:htmlColor lookup table
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
#>

Function SetupHTML
{
	Write-Verbose "$(Get-Date -Format G): Setting up HTML"
	If(!$AddDateTime)
	{
		[string]$Script:HTMLFileName = "$($Script:pwdpath)\$($OutputFileName).html"
	}
	ElseIf($AddDateTime)
	{
		[string]$Script:HTMLFileName = "$($Script:pwdpath)\$($OutputFileName)_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}

	$htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
	Out-File -FilePath $Script:HTMLFileName -Force -InputObject $HTMLHead 4>$Null
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
		} ## end Elseif
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
					## Build the available columns from all availble PSCustomObject note properties
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
					Write-Debug ("$(Get-Date -Format G): `t`tBuilding table headers");
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
				Write-Debug ("$(Get-Date -Format G): `t`tBuilding table rows");
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
				} ## end foreach
				Write-Debug ("$(Get-Date -Format G): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
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
					Write-Debug ("$(Get-Date -Format G): `t`tBuilding table headers");
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
				Write-Debug ("$(Get-Date -Format G): `t`tBuilding table rows");
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
				} ## end foreach

				Write-Debug ("$(Get-Date -Format G): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date -Format G): `t`tBuilding table range");
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
		Write-Debug ("$(Get-Date -Format G): `t`tConverting range to table");
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
			Write-Debug ("$(Get-Date -Format G): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page Breaks
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
		} # end Switch
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

#region email function
Function SendEmail
{
	Param([array]$Attachments)
	Write-Verbose "$(Get-Date -Format G): Prepare to email"

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
			Write-Verbose "$(Get-Date -Format G): Email successfully sent using anonymous credentials"
		}
		ElseIf(!$?)
		{
			$e = $error[0]

			Write-Verbose "$(Get-Date -Format G): Email was not sent:"
			Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
		}
	}
	Else
	{
		If($UseSSL)
		{
			Write-Verbose "$(Get-Date -Format G): Trying to send an email using current user's credentials with SSL"
			Send-MailMessage -Attachments $emailAttachment -Body $emailBody -BodyAsHtml -From $From `
			-Port $SmtpPort -SmtpServer $SmtpServer -Subject $emailSubject -To $To `
			-UseSSL *>$Null
		}
		Else
		{
			Write-Verbose  "$(Get-Date -Format G): Trying to send an email using current user's credentials without SSL"
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
				Write-Verbose "$(Get-Date -Format G): Current user's credentials failed. Ask for usable credentials."

				If($Dev)
				{
					Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append 4>$Null
				}

				$error.Clear()

				$emailCredentials = Get-Credential -UserName $From -Message "Enter the password to send an email"

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
					Write-Verbose "$(Get-Date -Format G): Email successfully sent using new credentials"
				}
				ElseIf(!$?)
				{
					$e = $error[0]

					Write-Verbose "$(Get-Date -Format G): Email was not sent:"
					Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
				}
			}
			Else
			{
				Write-Verbose "$(Get-Date -Format G): Email was not sent:"
				Write-Warning "$(Get-Date -Format G): Exception: $e.Exception" 
			}
		}
	}
}
#endregion

#region DHCP script functions
Function GetShortStatistics
{
	Param([object]$Statistics)
	
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $StatWordTable = @()
		[decimal]$TotalAddresses = "{0:N0}" -f ($Statistics.AddressesFree + $Statistics.AddressesInUse)
		[int]$InUsePercent = "{0:N0}" -f $Statistics.PercentageInUse.ToString()
		If($TotalAddresses -ne 0)
		{
			[int]$AvailablePercent = "{0:N0}" -f (($Statistics.AddressesFree / $TotalAddresses) * 100)
		}
		Else
		{
			[int]$AvailablePercent = 0
		}

		$WordTableRowHash = @{ 
		Description = "Total Addresses"; `
		Detail = $TotalAddresses.ToString()
		}

		## Add the hash to the array
		$StatWordTable += $WordTableRowHash;

		$WordTableRowHash = @{ 
		Description = "In Use"; `
		Detail = "$($Statistics.AddressesInUse) - $($InUsePercent)%"
		}

		## Add the hash to the array
		$StatWordTable += $WordTableRowHash;

		$WordTableRowHash = @{ 
		Description = "Available"; `
		Detail = "$($Statistics.AddressesFree) - $($AvailablePercent)%"
		}

		## Add the hash to the array
		$StatWordTable += $WordTableRowHash;

		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		If($StatWordTable.Count -gt 0)
		{
			$Table = AddWordTable -Hashtable $StatWordTable `
			-Columns Description,Detail `
			-Headers "Description","Details" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
	}
	If($Text)
	{
		Line 2 "Description" -NoNewLine
		Line 1 "Details"
		Line 2 "Total Addresses`t" -NoNewLine
		[decimal]$TotalAddresses = ($Statistics.AddressesFree + $Statistics.AddressesInUse)
		$tmp = "{0:N0}" -f $TotalAddresses
		Line 0 $tmp
		Line 2 "In Use`t`t" -NoNewLine
		[int]$InUsePercent = "{0:N0}" -f $Statistics.PercentageInUse
		$tmp = "{0:N0}" -f $Statistics.AddressesInUse
		Line 0 "$($tmp) - $($InUsePercent)%"
		Line 2 "Available`t" -NoNewLine
		If($TotalAddresses -ne 0)
		{
			[int]$AvailablePercent = "{0:N0}" -f (($Statistics.AddressesFree / $TotalAddresses) * 100)
		}
		Else
		{
			[int]$AvailablePercent = 0
		}
		$tmp = "{0:N0}" -f $Statistics.AddressesFree
		Line 0 "$($tmp) - $($AvailablePercent)%"
	}
	If($HTML)
	{
		$rowdata = @()

		[decimal]$TotalAddresses = "{0:N0}" -f ($Statistics.AddressesFree + $Statistics.AddressesInUse)
		$rowdata += @(,("Total Addresses",$htmlwhite,
						$TotalAddresses.ToString(),$htmlwhite))

		[int]$InUsePercent = "{0:N0}" -f $Statistics.PercentageInUse.ToString()
		$rowdata += @(,("In Use",$htmlwhite,
						"$($Statistics.AddressesInUse) - $($InUsePercent)%",$htmlwhite))

		If($TotalAddresses -ne 0)
		{
			[int]$AvailablePercent = "{0:N0}" -f (($Statistics.AddressesFree / $TotalAddresses) * 100)
		}
		Else
		{
			[int]$AvailablePercent = 0
		}
		$rowdata += @(,("Available",$htmlwhite,
						"$($Statistics.AddressesFree) - $($AvailablePercent)%",$htmlwhite))

		$columnHeaders = @('Description',($htmlsilver -bor $htmlbold),'Details',($htmlsilver -bor $htmlbold))
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}
	InsertBlankLine
}

Function ProcessServerProperties
{
	Write-Verbose "$(Get-Date -Format G): Server Properties and Configuration"
	Write-Verbose "$(Get-Date -Format G): "

	Write-Verbose "$(Get-Date -Format G): Getting DHCP server information"
	
	#added in V2.10, see if server is a domain controller
	$osInfo = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Script:DHCPServerName
	If($osInfo.ProductType -eq 2)
	{
		$IsADomainController = "Yes"
	}
	Else
	{
		$IsADomainController = "No"
	}
	
	$tmp = $Script:DHCPServerName.Split(".")
	$NetBIOSName = $tmp[0]
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "DHCP Server Information: " $NetBIOSName
		WriteWordLine 2 0 "Server Properties"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Server name"; Value = $Script:DHCPServerName; }
		$ScriptInformation += @{ Data = "Is a domain controller"; Value = $IsADomainController; }
	}
	If($Text)
	{
		Line 0 "DHCP Server Information: " $NetBIOSName
		Line 0 "Server Properties"
		Line 1 "Server name`t`t: " $Script:DHCPServerName
		Line 1 "Is a domain controller`t: " $IsADomainController
	}
	If($HTML)
	{
		WriteHTMLLine 1 0 "DHCP Server Information: " $NetBIOSName
		WriteHTMLLine 2 0 "Server Properties"
		$rowdata = @()
		$columnHeaders = @("Server name",($htmlsilver -bor $htmlbold),$Script:DHCPServerName,$htmlwhite)
		$rowdata += @(,('Is a domain controller',($htmlsilver -bor $htmlbold),$IsADomainController,$htmlwhite))
	}

	$DHCPDB = Get-DHCPServerDatabase -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $DHCPDB)
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Database path"; Value = $DHCPDB.FileName.SubString(0,($DHCPDB.FileName.LastIndexOf('\'))); }
			$ScriptInformation += @{ Data = "Backup path"; Value = $DHCPDB.BackupPath; }
		}
		If($Text)
		{
			Line 1 "Database path`t`t: " $DHCPDB.FileName.SubString(0,($DHCPDB.FileName.LastIndexOf('\')))
			Line 1 "Backup path`t`t: " $DHCPDB.BackupPath
		}
		If($HTML)
		{
			$rowdata += @(,('Database path',($htmlsilver -bor $htmlbold),$DHCPDB.FileName.SubString(0,($DHCPDB.FileName.LastIndexOf('\'))),$htmlwhite))
			$rowdata += @(,('Backup path',($htmlsilver -bor $htmlbold),$DHCPDB.BackupPath,$htmlwhite))
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Database path"; Value = "Error retrieving DHCP Server Database information"; }
			$ScriptInformation += @{ Data = "Backup path"; Value = "Error retrieving DHCP Server Database information"; }
		}
		If($Text)
		{
			Line 0 "Error retrieving DHCP Server Database information"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,('Database path',($htmlsilver -bor $htmlbold),"Error retrieving DHCP Server Database information",$htmlwhite))
			$rowdata += @(,('Backup path',($htmlsilver -bor $htmlbold),"Error retrieving DHCP Server Database information",$htmlwhite))
		}
	}

	$DHCPDB = $Null

	[bool]$Script:GotServerSettings = $False
	$Script:ServerSettings = Get-DHCPServerSetting -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $Script:ServerSettings)
	{
		$Script:GotServerSettings = $True
		#some properties of $Script:ServerSettings will be needed later
		If($Script:ServerSettings.IsAuthorized)
		{
			If($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "DHCP server"; Value = "Authorized"; }
			}
			If($Text)
			{
				Line 1 "DHCP server is authorized"
			}
			If($HTML)
			{
				$rowdata += @(,('DHCP server',($htmlsilver -bor $htmlbold),"Authorized",$htmlwhite))
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "DHCP server"; Value = "Not Authorized"; }
			}
			If($Text)
			{
				Line 1 "DHCP server is not authorized"
			}
			If($HTML)
			{
				$rowdata += @(,('DHCP server',($htmlsilver -bor $htmlbold),"Not Authorized",$htmlwhite))
			}
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "DHCP server"; Value = "Error retrieving DHCP Server setting information"; }
		}
		If($Text)
		{
			Line 0 "Error retrieving DHCP Server setting information"
		}
		If($HTML)
		{
			$rowdata += @(,('DHCP server',($htmlsilver -bor $htmlbold),"Error retrieving DHCP Server setting information",$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	If($HTML)
	{
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}
	
	InsertBlankLine
}

Function ProcessIPBindings
{
	Write-Verbose "$(Get-Date -Format G): `tGetting IPv4 bindings"
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $BindingsWordTable = @()
		WriteWordLine 2 0 "Connections and server bindings"
	}
	If($Text)
	{
		Line 0 "Connections and server bindings"
	}
	If($HTML)
	{
		$rowdata = @()
		WriteHTMLLine 2 0 "Connections and server bindings"
	}
	
	$IPv4Bindings = Get-DHCPServerV4Binding -ComputerName $Script:DHCPServerName -EA 0 | Sort-Object IPAddress

	If($? -and $Null -ne $IPv4Bindings)
	{
		ForEach($IPv4Binding in $IPv4Bindings)
		{
			If($IPv4Binding.BindingState)
			{
				If($MSWord -or $PDF)
				{
					$WordTableRowHash = @{ 
					Status = "Enabled"; `
					Binding = "$($IPv4Binding.IPAddress) $($IPv4Binding.InterfaceAlias)"
					}

					## Add the hash to the array
					$BindingsWordTable += $WordTableRowHash;
				}
				If($Text)
				{
					Line 1 "Enabled: " "$($IPv4Binding.IPAddress) $($IPv4Binding.InterfaceAlias)"
				}
				If($HTML)
				{
					$rowdata += @(,('Enabled',$htmlwhite,"$($IPv4Binding.IPAddress) $($IPv4Binding.InterfaceAlias)",$htmlwhite))
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					$WordTableRowHash = @{ 
					Status = "Disabled"; `
					Binding = "$($IPv4Binding.IPAddress) $($IPv4Binding.InterfaceAlias)"
					}

					## Add the hash to the array
					$BindingsWordTable += $WordTableRowHash;
				}
				If($Text)
				{
					Line 1 "Disabled: " "$($IPv4Binding.IPAddress) $($IPv4Binding.InterfaceAlias)"
				}
				If($HTML)
				{
					$rowdata += @(,('Disabled',$htmlwhite,"$($IPv4Binding.IPAddress) $($IPv4Binding.InterfaceAlias)",$htmlwhite))
				}
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			Status = "Error retrieving IPv4 server bindings"; `
			Binding = ""
			}

			## Add the hash to the array
			$BindingsWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv4 server bindings"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,('Error retrieving IPv4 server bindings',$htmlwhite,"",$htmlwhite))
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			Status = "There were no IPv4 server bindings"; `
			Binding = ""
			}

			## Add the hash to the array
			$BindingsWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 1 "There were no IPv4 server bindings"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,('There were no IPv4 server bindings',$htmlwhite,"",$htmlwhite))
		}
	}
	$IPv4Bindings = $Null

	Write-Verbose "$(Get-Date -Format G): `tGetting IPv6 bindings"
	$IPv6Bindings = Get-DHCPServerV6Binding -ComputerName $Script:DHCPServerName -EA 0 | Sort-Object IPAddress

	If($? -and $Null -ne $IPv6Bindings)
	{
		ForEach($IPv6Binding in $IPv6Bindings)
		{
			If($IPv6Binding.BindingState)
			{
				If($MSWord -or $PDF)
				{
					$WordTableRowHash = @{ 
					Status = "Enabled"; `
					Binding = "$($IPv6Binding.IPAddress) $($IPv6Binding.InterfaceAlias)"
					}

					## Add the hash to the array
					$BindingsWordTable += $WordTableRowHash;
				}
				If($Text)
				{
					Line 1 "Enabled: " "$($IPv6Binding.IPAddress) $($IPv6Binding.InterfaceAlias)"
				}
				If($HTML)
				{
					$rowdata += @(,('Enabled',$htmlwhite,"$($IPv6Binding.IPAddress) $($IPv6Binding.InterfaceAlias)",$htmlwhite))
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					$WordTableRowHash = @{ 
					Status = "Disabled"; `
					Binding = "$($IPv6Binding.IPAddress) $($IPv6Binding.InterfaceAlias)"
					}

					## Add the hash to the array
					$BindingsWordTable += $WordTableRowHash;
				}
				If($Text)
				{
					Line 1 "Disabled: " "$($IPv6Binding.IPAddress) $($IPv6Binding.InterfaceAlias)"
				}
				If($HTML)
				{
					$rowdata += @(,('Disabled',$htmlwhite,"$($IPv6Binding.IPAddress) $($IPv6Binding.InterfaceAlias)",$htmlwhite))
				}
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			Status = "Error retrieving IPv6 server bindings"; `
			Binding = ""
			}

			## Add the hash to the array
			$BindingsWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv6 server bindings"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,('Error retrieving IPv6 server bindings',$htmlwhite,"",$htmlwhite))
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			Status = "There were no IPv6 server bindings"; `
			Binding = ""
			}

			## Add the hash to the array
			$BindingsWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 1 "There were no IPv6 server bindings"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,('There were no IPv6 server bindings',$htmlwhite,"",$htmlwhite))
		}
	}
	$IPv6Bindings = $Null
	
	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		If($BindingsWordTable.Count -gt 0)
		{
			$Table = AddWordTable -Hashtable $BindingsWordTable `
			-Columns Status,Binding `
			-Headers "Status","Binding" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
	}
	If($HTML)
	{
		$columnHeaders = @('Status',($htmlsilver -bor $htmlbold),'Binding',($htmlsilver -bor $htmlbold))
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}
}

Function ProcessIPv4Properties
{
	Write-Verbose "$(Get-Date -Format G): Getting IPv4 properties"
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 2 0 "IPv4"
		WriteWordLine 3 0 "Properties"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	}
	If($Text)
	{
		Line 0 ""
		Line 0 "IPv4"
		Line 0 "Properties"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "IPv4"
		WriteHTMLLine 3 0 "Properties"
		$rowdata = @()
	}

	Write-Verbose "$(Get-Date -Format G): `tGetting IPv4 general settings"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "General"
	}
	If($Text)
	{
		Line 1 "General"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "General"
	}

	[bool]$Script:GotAuditSettings = $False
	$Script:AuditSettings = Get-DHCPServerAuditLog -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $Script:AuditSettings)
	{
		$Script:GotAuditSettings = $True
		If($Script:AuditSettings.Enable)
		{
			If($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "DHCP audit logging is"; Value = "Enabled"; }
			}
			If($Text)
			{
				Line 2 "DHCP audit logging is enabled"
			}
			If($HTML)
			{
				$columnHeaders = @("DHCP audit logging is",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "DHCP audit logging is"; Value = "Disabled"; }
			}
			If($Text)
			{
				Line 2 "DHCP audit logging is disabled"
			}
			If($HTML)
			{
				$columnHeaders = @("DHCP audit logging is",($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite)
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "DHCP audit logging is"; Value = "Error retrieving audit log settings"; }
		}
		If($Text)
		{
			Line 0 "Error retrieving audit log settings"
		}
		If($HTML)
		{
			$columnHeaders = @("DHCP audit logging is",($htmlsilver -bor $htmlbold),"Error retrieving audit log settings",$htmlwhite)
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "DHCP audit logging is"; Value = "There were no audit log settings"; }
		}
		If($Text)
		{
			Line 0 "There were no audit log settings"
		}
		If($HTML)
		{
			$columnHeaders = @("DHCP audit logging i",($htmlsilver -bor $htmlbold),"There were no audit log settings",$htmlwhite)
		}
	}

	#"HKLM:\SYSTEM\CurrentControlSet\Services\DHCPServer\Parameters" "BootFileTable"

	#Define the variable to hold the BOOTP Table
	$BOOTPKey="SYSTEM\CurrentControlSet\Services\DHCPServer\Parameters" 

	#Create an instance of the Registry Object and open the HKLM base key
	$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Script:DHCPServerName) 

	#Drill down into the BOOTP key using the OpenSubKey Method
	$regkey1=$reg.OpenSubKey($BOOTPKey) 

	#Retrieve an array of string that contain all the subkey names
	If($Null -ne $regkey1)
	{
		$Script:BOOTPTable = $regkey1.GetValue("BootFileTable") 
	}
	Else
	{
		$Script:BOOTPTable = $Null
	}

	If($Null -ne $Script:BOOTPTable)
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Show the BOOTP table folder is"; Value = "Enabled"; }
		}
		If($Text)
		{
			Line 2 "Show the BOOTP table folder is enabled"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,('Show the BOOTP table folder is',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Show the BOOTP table folder is"; Value = "Disabled"; }
		}
		If($Text)
		{
			Line 2 "Show the BOOTP table folder is Disabled"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,('Show the BOOTP table folder is',($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
		}
	}
	
	If($MSWord -or $PDF)
	{
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
	If($HTML)
	{
		$msg = ""
		$columnWidths = @("200","100")
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}

	#DNS settings
	Write-Verbose "$(Get-Date -Format G): `tGetting IPv4 DNS settings"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "DNS"
	}
	If($Text)
	{
		Line 1 "DNS"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "DNS"
	}

	$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $Script:DHCPServerName -EA 0
	If($? -and $Null -ne $DNSSettings)
	{
		If($DNSSettings.DeleteDnsRROnLeaseExpiry)
		{
			$DeleteDnsRROnLeaseExpiry = "Enabled"
		}
		Else
		{
			$DeleteDnsRROnLeaseExpiry = "Disabled"
		}

		If($DNSSettings.UpdateDnsRRForOlderClients)
		{
			$UpdateDnsRRForOlderClients = "Enabled"
		}
		Else
		{
			$UpdateDnsRRForOlderClients = "Disabled"
		}

		If($DNSSettings.DisableDnsPtrRRUpdate)
		{
			$DisableDnsPtrRRUpdate = "Enabled"
		}
		Else
		{
			$DisableDnsPtrRRUpdate = "Disabled"
		}

		If($DNSSettings.NameProtection)
		{
			$NameProtection = "Enabled"
		}
		Else
		{
			$NameProtection = "Disabled"
		}

		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			If($DNSSettings.DynamicUpdates -eq "Never")
			{
				$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Disabled"; }
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
			{
				$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
				$ScriptInformation += @{ Data = "Dynamically update DNS A and PTR records only if requested by the DHCP clients"; Value = "Enabled"; }
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "Always")
			{
				$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
				$ScriptInformation += @{ Data = "Always dynamically update DNS A and PTR records"; Value = "Enabled"; }
			}
			$ScriptInformation += @{ Data = "Discard A and PTR records when lease is deleted"; Value = $DeleteDnsRROnLeaseExpiry; }
			$ScriptInformation += @{ Data = "Dynamically update DNS records for DHCP clients that do not request updates"; Value = $UpdateDnsRRForOlderClients; }
			$ScriptInformation += @{ Data = "Disable dynamic updates for DNS PTR record"; Value = $DisableDnsPtrRRUpdate; }
			$ScriptInformation += @{ Data = "Name Protection"; Value = $NameProtection; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 400;
			$Table.Columns.Item(2).Width = 50;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			$ScriptInformation = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 2 "Enable DNS dynamic updates`t`t`t: " -NoNewLine
			If($DNSSettings.DynamicUpdates -eq "Never")
			{
				Line 0 "Disabled"
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
			{
				Line 0 "Enabled"
				Line 2 "Dynamically update DNS A and PTR records only "
				Line 2 "if requested by the DHCP clients`t`t: Enabled"
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "Always")
			{
				Line 0 "Enabled"
				Line 2 "Always dynamically update DNS A and PTR records: Enabled"
			}
			Line 2 "Discard A and PTR records when lease deleted`t: " $DeleteDnsRROnLeaseExpiry
			Line 2 "Dynamically update DNS records for DHCP "
			Line 2 "clients that do not request updates`t`t: " $UpdateDnsRRForOlderClients
			Line 2 "Disable dynamic updates for DNS PTR records`t: " $DisableDnsPtrRRUpdate$DisableDnsPtrRRUpdate
			Line 2 "Name Protection`t`t`t`t`t: " $NameProtection
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()
			If($DNSSettings.DynamicUpdates -eq "Never")
			{
				$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite)
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
			{
				$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
				$rowdata += @(,("Dynamically update DNS A and PTR records only if requested by the DHCP clients",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "Always")
			{
				$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
				$rowdata += @(,("Always dynamically update DNS A and PTR records",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
			}
			$rowdata += @(,("Discard A and PTR records when lease deleted",($htmlsilver -bor $htmlbold),$DeleteDnsRROnLeaseExpiry,$htmlwhite))
			$rowdata += @(,('Dynamically update DNS records for DHCP clients that do not request updates',($htmlsilver -bor $htmlbold),$UpdateDnsRRForOlderClients,$htmlwhite))
			$rowdata += @(,('Disable dynamic updates for DNS PTR records',($htmlsilver -bor $htmlbold),$DisableDnsPtrRRUpdate,$htmlwhite))
			$rowdata += @(,('Name Protection',($htmlsilver -bor $htmlbold),$NameProtection,$htmlwhite))
			$msg = ""
			$columnWidths = @("450","50")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
			WriteHTMLLine 0 0 
			$rowdata = $Null
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv4 DNS Settings for DHCP server $Script:DHCPServerName"
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv4 DNS Settings for DHCP server $Script:DHCPServerName"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv4 DNS Settings for DHCP server $Script:DHCPServerName"
		}
	}
	$DNSSettings = $Null

	#now back to some server settings
	Write-Verbose "$(Get-Date -Format G): `tGetting IPv4 NAP settings"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Network Access Protection"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	}
	If($Text)
	{
		Line 1 "Network Access Protection"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Network Access Protection"
		$rowdata = @()
	}

	If($Script:GotServerSettings)
	{
		Switch($Script:ServerSettings.NpsUnreachableAction)
		{
			"Full"			{$NpsUnreachableAction = "Full Access"; Break}
			"Restricted"	{$NpsUnreachableAction = "Restricted Access"; Break}
			"NoAccess"		{$NpsUnreachableAction = "Drop Client Packet"; Break}
			Default			{$NpsUnreachableAction = "Unable to determine NPS unreachable action: $($Script:ServerSettings.NpsUnreachableAction)"; Break}
		}

		If($Script:ServerSettings.NapEnabled)
		{
			$NapEnabled = "Enabled on all scopes"
		}
		Else
		{
			$NapEnabled = "Disabled on all scopes"
		}

		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Network Access Protection is"; Value = $NapEnabled; }
			$ScriptInformation += @{ Data = "DHCP server behavior when NPS is unreachable"; Value = $NpsUnreachableAction; }

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
		If($Text)
		{
			Line 2 "Network Access Protection is`t`t`t: " $NapEnabled
			Line 2 "DHCP server behavior when NPS is unreachable`t: " $NpsUnreachableAction
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @("Network Access Protection is",($htmlsilver -bor $htmlbold),$NapEnabled,$htmlwhite)
			$rowdata += @(,('DHCP server behavior when NPS is unreachable',($htmlsilver -bor $htmlbold),$NpsUnreachableAction,$htmlwhite))
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 ""
		}
	}

	#filters
	Write-Verbose "$(Get-Date -Format G): `tGetting IPv4 filters"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Filters"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	}
	If($Text)
	{
		Line 1 "Filters"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Filters"
		$rowdata = @()
	}

	$MACFilters = Get-DHCPServerV4FilterList -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $MACFilters)
	{
		If($MACFilters.Allow)
		{
			$MACFiltersAllow = "Enabled"
		}
		Else
		{
			$MACFiltersAllow = "Disabled"
		}
		If($MACFilters.Deny)
		{
			$MACFiltersDeny = "Enabled"
		}
		Else
		{
			$MACFiltersDeny = "Disabled"
		}

		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Enable Allow list"; Value = $MACFiltersAllow; }
			$ScriptInformation += @{ Data = "Enable Deny list"; Value = $MACFiltersDeny; }
		}
		If($Text)
		{
			Line 2 "Enable Allow list`t: " $MACFiltersAllow
			Line 2 "Enable Deny list`t: " $MACFiltersDeny
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @("Enable Allow list",($htmlsilver -bor $htmlbold),$MACFiltersAllow,$htmlwhite)
			$rowdata += @(,('Enable Deny list',($htmlsilver -bor $htmlbold),$MACFiltersDeny,$htmlwhite))
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Error retrieving MAC filters for DHCP server $Script:DHCPServerName"; Value = ""; }
		}
		If($Text)
		{
			Line 0 "Error retrieving MAC filters for DHCP server $Script:DHCPServerName"
		}
		If($HTML)
		{
			$columnHeaders = @("Error retrieving MAC filters for DHCP server $Script:DHCPServerName",($htmlsilver -bor $htmlbold),"",$htmlwhite)
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "There were no MAC filters for DHCP server $Script:DHCPServerName"; Value = ""; }
		}
		ElseIf($Text)
		{
			Line 2 "There were no MAC filters for DHCP server $Script:DHCPServerName"
		}
		ElseIf($HTML)
		{
			$columnHeaders = @("There were no MAC filters for DHCP server $Script:DHCPServerName",($htmlsilver -bor $htmlbold),"",$htmlwhite)
		}
	}
	$MACFilters = $Null
	
	If($MSWord -or $PDF)
	{
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
	If($HTML)
	{
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}

	#failover
	Write-Verbose "$(Get-Date -Format G): `tGetting IPv4 Failover"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Failover"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	}
	If($Text)
	{
		Line 1 "Failover"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Failover"
		$rowdata = @()
	}

	$Failovers = Get-DHCPServerV4Failover -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $Failovers)
	{
		ForEach($Failover in $Failovers)
		{
			Write-Verbose "$(Get-Date -Format G): `t`tProcessing failover $($Failover.Name)"
					
			Switch($Failover.State)
			{
				"NoState"					{$FailoverState = "No State"; Break}
				"Normal"					{$FailoverState = "Normal"; Break}
				"Init"						{$FailoverState = "Initializing"; Break}
				"CommunicationInterrupted"	{$FailoverState = "Communication Interrupted"; Break}
				"PartnerDown"				{$FailoverState = "Normal"; Break}
				"PotentialConflict"			{$FailoverState = "Potential Conflict"; Break}
				"Startup"					{$FailoverState = "Starting Up"; Break}
				"ResolutionInterrupted"		{$FailoverState = "Resolution Interrupted"; Break}
				"ConflictDone"				{$FailoverState = "Conflict Done"; Break}
				"Recover"					{$FailoverState = "Recover"; Break}
				"RecoverWait"				{$FailoverState = "Recover Wait"; Break}
				"RecoverDone"				{$FailoverState = "Recover Done"; Break}
				Default						{$FailoverState = "Unable to determine server failover state: $($Failover.State)"; Break}
			}
			If($Failover.EnableAuth)
			{
				$EnableAuth = "Enabled"
			}
			Else
			{
				$EnableAuth = "Disabled"
			}
			If($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "Relationship name"; Value = $Failover.Name; }
				$ScriptInformation += @{ Data = "State of the server"; Value = $FailoverState; }
				$ScriptInformation += @{ Data = "Partner Server"; Value = $Failover.PartnerServer; }
				$ScriptInformation += @{ Data = "Mode"; Value = $Failover.Mode; }
				$ScriptInformation += @{ Data = "Message Authentication"; Value = $EnableAuth; }
						
				If($Failover.Mode -eq "LoadBalance")
				{
					$ScriptInformation += @{ Data = "Local server"; Value = "$($Failover.LoadBalancePercent)%"; }
						
					$tmp = (100 - $($Failover.LoadBalancePercent))
					$ScriptInformation += @{ Data = "Partner Server"; Value = "$($tmp)%"; }
					$tmp = $Null
				}
				Else
				{
					$ScriptInformation += @{ Data = "Role of this server"; Value = $Failover.ServerRole; }
					$ScriptInformation += @{ Data = "Addresses reserved for standby server"; Value = "$($Failover.ReservePercent)%"; }
				}
			}
			If($Text)
			{
				Line 2 "Relationship name: " $Failover.Name
				Line 2 "State of the server`t: " $FailoverState
				Line 2 "Partner Server`t`t: " $Failover.PartnerServer
				Line 2 "Mode`t`t`t: " $Failover.Mode
				Line 2 "Message Authentication`t: " $EnableAuth
				If($Failover.Mode -eq "LoadBalance")
				{
					$tmp = (100 - $($Failover.LoadBalancePercent))
					Line 2 "Local server`t`t: $($Failover.LoadBalancePercent)%"
					Line 2 "Partner Server`t`t: $($tmp)%"
					$tmp = $Null
				}
				Else
				{
					Line 2 "Role of this server`t: " $Failover.ServerRole
					Line 2 "Addresses reserved for standby server: $($Failover.ReservePercent)%"
					Line 0 ""
				}
			}
			If($HTML)
			{
				$columnHeaders = @("Relationship name",($htmlsilver -bor $htmlbold),$Failover.Name,$htmlwhite)
				$rowdata += @(,('State of the server',($htmlsilver -bor $htmlbold),$FailoverState,$htmlwhite))
				$rowdata += @(,('Partner Server',($htmlsilver -bor $htmlbold),$Failover.PartnerServer,$htmlwhite))
				$rowdata += @(,('Mode',($htmlsilver -bor $htmlbold),$Failover.Mode,$htmlwhite))
				$rowdata += @(,('Message Authentication',($htmlsilver -bor $htmlbold),$EnableAuth,$htmlwhite))
				If($Failover.Mode -eq "LoadBalance")
				{
					$tmp = (100 - $($Failover.LoadBalancePercent))
					$rowdata += @(,('Local server',($htmlsilver -bor $htmlbold),"$($Failover.LoadBalancePercent)%",$htmlwhite))
					$rowdata += @(,('Partner Server',($htmlsilver -bor $htmlbold),"$($tmp)%",$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('Role of this server',($htmlsilver -bor $htmlbold),$Failover.ServerRole,$htmlwhite))
					$rowdata += @(,('Addresses reserved for standby server',($htmlsilver -bor $htmlbold),"$($Failover.ReservePercent)%",$htmlwhite))
				}
			}
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "There was no Failover configured for DHCP server"; Value = $Script:DHCPServerName; }
		}
		If($Text)
		{
			Line 2 "There was no Failover configured for DHCP server $Script:DHCPServerName"
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @("There was no Failover configured for DHCP server",($htmlsilver -bor $htmlbold),$Script:DHCPServerName,$htmlwhite)
		}
	}
	$Failovers = $Null
	
	If($MSWord -or $PDF)
	{
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
	If($HTML)
	{
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}

	#Advanced
	Write-Verbose "$(Get-Date -Format G): `tGetting IPv4 advanced settings"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Advanced"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	}
	If($Text)
	{
		Line 1 "Advanced"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Advanced"
		$rowdata = @()
	}

	If($Script:GotServerSettings)
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Conflict detection attempts"; Value = $Script:ServerSettings.ConflictDetectionAttempts.ToString(); }
		}
		If($Text)
		{
			Line 2 "Conflict detection attempts`t: " $Script:ServerSettings.ConflictDetectionAttempts.ToString()
		}
		If($HTML)
		{
			$columnHeaders = @("Conflict detection attempts",($htmlsilver -bor $htmlbold),$Script:ServerSettings.ConflictDetectionAttempts.ToString(),$htmlwhite)
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Conflict detection attempts"; Value = "Unable to determine"; }
		}
		If($Text)
		{
			Line 2 "Conflict detection attempts`t: " "Unable to determine"
		}
		If($HTML)
		{
			$columnHeaders = @("Conflict detection attempts",($htmlsilver -bor $htmlbold),"Unable to determine",$htmlwhite)
		}
	}

	If($Script:GotAuditSettings)
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Audit log file path"; Value = $Script:AuditSettings.Path; }
		}
		If($Text)
		{
			Line 2 "Audit log file path`t`t: " $Script:AuditSettings.Path
		}
		If($HTML)
		{
			$rowdata += @(,('Audit log file path',($htmlsilver -bor $htmlbold),$Script:AuditSettings.Path,$htmlwhite))
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Audit log file path"; Value = "Unable to determine"; }
		}
		If($Text)
		{
			Line 2 "Audit log file path`t`t: " "Unable to determine"
		}
		If($HTML)
		{
			$rowdata += @(,('Audit log file path',($htmlsilver -bor $htmlbold),"Unable to determine",$htmlwhite))
		}
	}

	#added 18-Jan-2016
	#get dns update credentials
	Write-Verbose "$(Get-Date -Format G): `tGetting DNS dynamic update registration credentials"
	$DNSUpdateSettings = Get-DhcpServerDnsCredential -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $DNSUpdateSettings)
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "DNS dynamic update registration credentials"; Value = ""; }
			$ScriptInformation += @{ Data = "     User name"; Value = $DNSUpdateSettings.UserName; }
			$ScriptInformation += @{ Data = "     Domain"; Value = $DNSUpdateSettings.DomainName; }
		}
		If($Text)
		{
			Line 2 "DNS dynamic update registration credentials: "
			Line 3 "User name`t: " $DNSUpdateSettings.UserName
			Line 3 "Domain`t`t: " $DNSUpdateSettings.DomainName
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,('DNS dynamic update registration credentials',($htmlsilver -bor $htmlbold),"",$htmlwhite))
			$rowdata += @(,('     User name',($htmlsilver -bor $htmlbold),$DNSUpdateSettings.UserName,$htmlwhite))
			$rowdata += @(,('     Domain',($htmlsilver -bor $htmlbold),$DNSUpdateSettings.DomainName,$htmlwhite))
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Error retrieving DNS Update Credentials for DHCP server"; Value = $Script:DHCPServerName; }
		}
		If($Text)
		{
			Line 0 "Error retrieving DNS Update Credentials for DHCP server $Script:DHCPServerName"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,('Error retrieving DNS Update Credentials for DHCP server',($htmlsilver -bor $htmlbold),$Script:DHCPServerName,$htmlwhite))
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "There were no DNS Update Credentials for DHCP server"; Value = $Script:DHCPServerName; }
		}
		If($Text)
		{
			Line 2 "There were no DNS Update Credentials for DHCP server $Script:DHCPServerName"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,('There were no DNS Update Credentials for DHCP server',($htmlsilver -bor $htmlbold),$Script:DHCPServerName,$htmlwhite))
		}
	}
	$DNSUpdateSettings = $Null
	
	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	If($HTML)
	{
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}
}

Function ProcessIPv4Statistics
{
	Write-Verbose "$(Get-Date -Format G): Getting IPv4 Statistics"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Statistics"
		[System.Collections.Hashtable[]] $StatWordTable = @()
	}
	If($Text)
	{
		Line 1 "Statistics"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Statistics"
		$rowdata = @()
	}

	$Statistics = Get-DHCPServerV4Statistics -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $Statistics)
	{
		$UpTime = $(Get-Date) - $Statistics.ServerStartTime
		$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3} seconds", `
			$UpTime.Days, `
			$UpTime.Hours, `
			$UpTime.Minutes, `
			$UpTime.Seconds)
		[int]$InUsePercent = "{0:N0}" -f $Statistics.PercentageInUse
		[int]$AvailablePercent = "{0:N0}" -f $Statistics.PercentageAvailable
		
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			Description = "Start Time"; `
			Detail = $Statistics.ServerStartTime.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Up Time"; `
			Detail =  $Str
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Discovers"; `
			Detail = $Statistics.Discovers.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Offers"; `
			Detail = $Statistics.Offers.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Delayed Offers"; `
			Detail = $Statistics.DelayedOffers.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Requests"; `
			Detail = $Statistics.Requests.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Acks"; `
			Detail = $Statistics.Acks.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Nacks"; `
			Detail = $Statistics.Naks.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Declines"; `
			Detail = $Statistics.Declines.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Releases"; `
			Detail = $Statistics.Releases.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Total Scopes"; `
			Detail = $Statistics.TotalScopes.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Scopes with delay configured"; `
			Detail = $Statistics.ScopesWithDelayConfigured.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Total Addresses"; `
			Detail = "{0:N0}" -f $Statistics.TotalAddresses.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "In Use"; `
			Detail = "$($Statistics.AddressesInUse) - $($InUsePercent)%"
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Available"; `
			Detail = "{0:N0}" -f "$($Statistics.AddressesAvailable) - $($AvailablePercent)%"
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 2 "Description" -NoNewLine
			Line 3 "Details"

			Line 2 "Start Time" -NoNewLine
			Line 3 $Statistics.ServerStartTime
			Line 2 "Up Time" -NoNewLine
			Line 4 $Str
			Line 2 "Discovers" -NoNewLine
			Line 3 $Statistics.Discovers
			Line 2 "Offers" -NoNewLine
			Line 4 $Statistics.Offers
			Line 2 "Delayed Offers" -NoNewLine
			Line 3 $Statistics.DelayedOffers
			Line 2 "Requests" -NoNewLine
			Line 3 $Statistics.Requests
			Line 2 "Acks" -NoNewLine
			Line 4 $Statistics.Acks
			Line 2 "Nacks" -NoNewLine
			Line 4 $Statistics.Naks
			Line 2 "Declines" -NoNewLine
			Line 3 $Statistics.Declines
			Line 2 "Releases" -NoNewLine
			Line 3 $Statistics.Releases
			Line 2 "Total Scopes" -NoNewLine
			Line 3 $Statistics.TotalScopes
			Line 2 "Scopes w/delay configured" -NoNewLine
			Line 1 $Statistics.ScopesWithDelayConfigured
			Line 2 "Total Addresses" -NoNewLine
			$tmp = "{0:N0}" -f $Statistics.TotalAddresses
			Line 3 $tmp
			Line 2 "In Use" -NoNewLine
			Line 4 "$($Statistics.AddressesInUse) - $($InUsePercent)%"
			Line 2 "Available" -NoNewLine
			$tmp = "{0:N0}" -f $Statistics.AddressesAvailable
			Line 3 "$($tmp) - $($AvailablePercent)%"
		}
		If($HTML)
		{
			$rowdata += @(,("Start Time",$htmlwhite,$Statistics.ServerStartTime.ToString(),$htmlwhite))
			$rowdata += @(,("Up Time",$htmlwhite,$Str,$htmlwhite))
			$rowdata += @(,("Discovers",$htmlwhite,$Statistics.Discovers.ToString(),$htmlwhite))
			$rowdata += @(,("Offers",$htmlwhite,$Statistics.Offers.ToString(),$htmlwhite))
			$rowdata += @(,("Delayed Offers",$htmlwhite,$Statistics.DelayedOffers.ToString(),$htmlwhite))
			$rowdata += @(,("Requests",$htmlwhite,$Statistics.Requests.ToString(),$htmlwhite))
			$rowdata += @(,("Acks",$htmlwhite,$Statistics.Acks.ToString(),$htmlwhite))
			$rowdata += @(,("Nacks",$htmlwhite,$Statistics.Naks.ToString(),$htmlwhite))
			$rowdata += @(,("Declines",$htmlwhite,$Statistics.Declines.ToString(),$htmlwhite))
			$rowdata += @(,("Releases",$htmlwhite,$Statistics.Releases.ToString(),$htmlwhite))
			$rowdata += @(,("Total Scopes",$htmlwhite,$Statistics.TotalScopes.ToString(),$htmlwhite))
			$rowdata += @(,("Scopes with delay configured",$htmlwhite,$Statistics.ScopesWithDelayConfigured.ToString(),$htmlwhite))
			$tmp = "{0:N0}" -f $Statistics.TotalAddresses.ToString()
			$rowdata += @(,("Total Addresses",$htmlwhite,$tmp,$htmlwhite))
			$rowdata += @(,("In Use",$htmlwhite,"$($Statistics.AddressesInUse) - $($InUsePercent)%",$htmlwhite))
			$tmp = "{0:N0}" -f "$($Statistics.AddressesAvailable) - $($AvailablePercent)%"
			$rowdata += @(,("Available",$htmlwhite,$tmp,$htmlwhite))
		}

		Write-Verbose "$(Get-Date -Format G): `tFinished IPv4 statistics table"
		Write-Verbose "$(Get-Date -Format G): "
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			Description = "Error retrieving IPv4 statistics"; `
			Detail = ""
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv4 statistics"
		}
		If($HTML)
		{
			$rowdata += @(,("Error retrieving IPv4 statistics",$htmlwhite,"",$htmlwhite))
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			Description = "There were no IPv4 statistics"; `
			Detail = ""
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 0 "There were no IPv4 statistics"
		}
		If($HTML)
		{
			$rowdata += @(,("There were no IPv4 statistics",$htmlwhite,"",$htmlwhite))
		}
	}
	$Statistics = $Null
	
	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		If($StatWordTable.Count -gt 0)
		{
			$Table = AddWordTable -Hashtable $StatWordTable `
			-Columns Description,Detail `
			-Headers "Description","Details" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
	}
	If($HTML)
	{
		$columnHeaders = @('Description',($htmlsilver -bor $htmlbold),'Details',($htmlsilver -bor $htmlbold))
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}
}

Function ProcessIPv4Superscopes
{
	Write-Verbose "$(Get-Date -Format G): Getting IPv4 Superscopes"
	$IPv4Superscopes = Get-DHCPServerV4Superscope -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $IPv4Superscopes)
	{
		ForEach($IPv4Superscope in $IPv4Superscopes)
		{
			If(![string]::IsNullOrEmpty($IPv4Superscope.SuperscopeName))
			{
				Write-Verbose "$(Get-Date -Format G): `tGetting IPv4 superscope data for scope $($IPv4Superscope.SuperscopeName)"
				If($MSWord -or $PDF)
				{
					#put each superscope on a new page
					$selection.InsertNewPage()
					WriteWordLine 3 0 "Superscope [$($IPv4Superscope.SuperscopeName)]"

					#get superscope statistics first
					$Statistics = Get-DHCPServerV4SuperscopeStatistics -ComputerName $Script:DHCPServerName -Name $IPv4Superscope.SuperscopeName -EA 0

					If($? -and $Null -ne $Statistics)
					{
						GetShortStatistics $Statistics
					}
					ElseIf(!$?)
					{
						WriteWordLine 0 0 "Error retrieving superscope statistics"
					}
					Else
					{
						WriteWordLine 0 1 "There were no statistics for the superscope"
					}
					$Statistics = $Null
				
					$xScopeIds = $IPv4Superscope.ScopeId
					[int]$StartLevel = 4
					ForEach($xScopeId in $xScopeIds)
					{
						$IPv4Scope = Get-DHCPServerV4Scope -ComputerName $Script:DHCPServerName -ScopeId $xScopeId -EA 0
						
						If($? -and $Null -ne $IPv4Scope)
						{
							GetIPv4ScopeData $IPv4Scope $StartLevel
						}
						Else
						{
							WriteWordLine 0 0 "Error retrieving Superscope data for scope $($xScopeId)"
						}
					}
				}
				If($Text)
				{
					Line 0 ""
					Line 0 "Superscope [$($IPv4Superscope.SuperscopeName)]"

					#get superscope statistics first
					$Statistics = Get-DHCPServerV4SuperscopeStatistics -ComputerName $Script:DHCPServerName -Name $IPv4Superscope.SuperscopeName -EA 0

					If($? -and $Null -ne $Statistics)
					{
						Line 1 "Statistics:"
						GetShortStatistics $Statistics
					}
					ElseIf(!$?)
					{
						Line 0 "Error retrieving superscope statistics"
					}
					Else
					{
						Line 2 "There were no statistics for the superscope"
					}
					$Statistics = $Null
				
					$xScopeIds = $IPv4Superscope.ScopeId
					[int]$StartLevel = 4
					ForEach($xScopeId in $xScopeIds)
					{
						$IPv4Scope = Get-DHCPServerV4Scope -ComputerName $Script:DHCPServerName -ScopeId $xScopeId -EA 0
						
						If($? -and $Null -ne $IPv4Scope)
						{
							GetIPv4ScopeData $IPv4Scope $StartLevel
						}
						Else
						{
							Line 0 "Error retrieving Superscope data for scope $($xScopeId)"
						}
					}
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "Superscope [$($IPv4Superscope.SuperscopeName)]"

					#get superscope statistics first
					$Statistics = Get-DHCPServerV4SuperscopeStatistics -ComputerName $Script:DHCPServerName -Name $IPv4Superscope.SuperscopeName -EA 0

					If($? -and $Null -ne $Statistics)
					{
						GetShortStatistics $Statistics
					}
					ElseIf(!$?)
					{
						WriteHTMLLine 0 0 "Error retrieving superscope statistics"
					}
					Else
					{
						WriteHTMLLine 0 1 "There were no statistics for the superscope"
					}
					$Statistics = $Null
				
					$xScopeIds = $IPv4Superscope.ScopeId
					[int]$StartLevel = 4
					ForEach($xScopeId in $xScopeIds)
					{
						$IPv4Scope = Get-DHCPServerV4Scope -ComputerName $Script:DHCPServerName -ScopeId $xScopeId -EA 0
						
						If($? -and $Null -ne $IPv4Scope)
						{
							GetIPv4ScopeData $IPv4Scope $StartLevel
						}
						Else
						{
							WriteHTMLLine 0 0 "Error retrieving Superscope data for scope $($xScopeId)"
						}
					}
				}
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv4 Superscopes"
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv4 Superscopes"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv4 Superscopes"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "There were no IPv4 Superscopes"
		}
		If($Text)
		{
			Line 0 "There were no IPv4 Superscopes"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "There were no IPv4 Superscopes"
		}
	}
	$IPv4Superscopes = $Null
	InsertBlankLine
}

Function ProcessIPv4Scopes
{
	Write-Verbose "$(Get-Date -Format G): Getting IPv4 scopes"
	$IPv4Scopes = Get-DHCPServerV4Scope -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $IPv4Scopes)
	{
		[int]$StartLevel = 3
		ForEach($IPv4Scope in $IPv4Scopes)
		{
			GetIPv4ScopeData $IPv4Scope $StartLevel
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv4 scopes"
		}
		ElseIf($Text)
		{
			Line 0 "Error retrieving IPv4 scopes"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv4 scopes"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no IPv4 scopes"
		}
		ElseIf($Text)
		{
			Line 2 "There were no IPv4 scopes"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "There were no IPv4 scopes"
		}
	}
	$IPv4Scopes = $Null
}

Function GetIPv4ScopeData
{
	Param([object]$IPv4Scope, [int]$xStartLevel)
	
	Write-Verbose "$(Get-Date -Format G): `tBuild array of Allow/Deny filters"
	$Filters = Get-DHCPServerV4Filter -ComputerName $Script:DHCPServerName -EA 0

	ProcessIPv4ScopeData $IPv4Scope $xStartLevel $Filters
}

Function ProcessIPv4ScopeData
{
	Param([object]$IPv4Scope, [int] $xStartLevel, [object]$filters)
	Write-Verbose "$(Get-Date -Format G): `tGetting IPv4 scope data for scope $($IPv4Scope.Name)"

	If($MSWord -or $PDF)
	{
		#put each scope on a new page
		$selection.InsertNewPage()
		WriteWordLine $xStartLevel 0 "Scope [$($IPv4Scope.ScopeId)] $($IPv4Scope.Name)"
		WriteWordLine ($xStartLevel + 1) 0 "Address Pool"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Start IP Address"; Value = $IPv4Scope.StartRange.ToString(); }
		$ScriptInformation += @{ Data = "End IP Address"; Value = $IPv4Scope.EndRange.ToString(); }
		$ScriptInformation += @{ Data = "Subnet Mask"; Value = $IPv4Scope.SubnetMask.ToString(); }
		If($IPv4Scope.LeaseDuration -eq "00:00:00")
		{
			$Str = "Unlimited"
		}
		Else
		{
			$Str = [string]::format("{0} days, {1} hours, {2} minutes", `
				$IPv4Scope.LeaseDuration.Days, `
				$IPv4Scope.LeaseDuration.Hours, `
				$IPv4Scope.LeaseDuration.Minutes)
		}
		$ScriptInformation += @{ Data = "Lease duration"; Value = $Str; }
		$ScriptInformation += @{ Data = "Description"; Value = $IPv4Scope.Description; }
		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
	}
	If($Text)
	{
		Line 0 "Scope [$($IPv4Scope.ScopeId)] $($IPv4Scope.Name)"
		Line 1 "Address Pool:"
		Line 2 "Start IP Address`t: " $IPv4Scope.StartRange
		Line 2 "End IP Address`t`t: " $IPv4Scope.EndRange
		Line 2 "Subnet Mask`t`t: " $IPv4Scope.SubnetMask
		Line 2 "Lease duration`t`t: " -NoNewLine
		If($IPv4Scope.LeaseDuration -eq "00:00:00")
		{
			Line 0 "Unlimited"
		}
		Else
		{
			$Str = [string]::format("{0} days, {1} hours, {2} minutes", `
				$IPv4Scope.LeaseDuration.Days, `
				$IPv4Scope.LeaseDuration.Hours, `
				$IPv4Scope.LeaseDuration.Minutes)

			Line 0 $Str
		}
	}
	If($HTML)
	{
		WriteHTMLLine $xStartLevel 0 "Scope [$($IPv4Scope.ScopeId)] $($IPv4Scope.Name)"
		WriteHTMLLine ($xStartLevel + 1) 0 "Address Pool"
		$rowdata = @()
		$columnHeaders = @("Start IP Address",($htmlsilver -bor $htmlbold),$IPv4Scope.StartRange.ToString(),$htmlwhite)
		$rowdata += @(,('End IP Address',($htmlsilver -bor $htmlbold),$IPv4Scope.EndRange.ToString(),$htmlwhite))
		$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlbold),$IPv4Scope.SubnetMask.ToString(),$htmlwhite))
		$tmp = ""
		If($IPv4Scope.LeaseDuration -eq "00:00:00")
		{
			$tmp = "Unlimited"
		}
		Else
		{
			$Str = [string]::format("{0} days, {1} hours, {2} minutes", `
				$IPv4Scope.LeaseDuration.Days, `
				$IPv4Scope.LeaseDuration.Hours, `
				$IPv4Scope.LeaseDuration.Minutes)

			$tmp = $Str
		}
		$rowdata += @(,('Lease duration',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
		$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$IPv4Scope.Description,$htmlwhite))
								
		$msg = ""
		$columnWidths = @("150","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
	}
	InsertBlankLine

	If($IncludeLeases)
	{
		Write-Verbose "$(Get-Date -Format G):	`t`tGetting leases"
		$Leases = Get-DHCPServerV4Lease -ComputerName $Script:DHCPServerName -ScopeId  $IPv4Scope.ScopeId -EA 0 | Sort-Object IPAddress

		If($MSWord -or $PDF)
		{
			WriteWordLine ($xStartLevel + 1) 0 "Address Leases"
		}
		If($Text)
		{
			Line 1 "Address Leases:"
		}
		If($HTML)
		{
			WriteHTMLLine ($xStartLevel + 1) 0 "Address Leases"
		}
		
		If($? -and $Null -ne $Leases)
		{
			ForEach($Lease in $Leases)
			{
				Write-Verbose "$(Get-Date -Format G):	`t`t`tProcessing lease $($Lease.IPAddress)"
				If($Null -ne $Lease.LeaseExpiryTime)
				{
					$LeaseStr = [string]::format("{0} days, {1} hours, {2} minutes", `
						$Lease.LeaseExpiryTime.Day, `
						$Lease.LeaseExpiryTime.Hour, `
						$Lease.LeaseExpiryTime.Minute)
				}
				Else
				{
					$LeaseStr = ""
				}

				If($Null -ne $Lease.ProbationEnds)
				{
					$ProbationStr = [string]::format("{0} days, {1} hours, {2} minutes", `
						$Lease.ProbationEnds.Day, `
						$Lease.ProbationEnds.Hour, `
						$Lease.ProbationEnds.Minute)
				}
				Else
				{
					$ProbationStr = ""
				}

				Switch($Lease.NapStatus)
				{
					"FullAccess"				{$NapStatus = "Full Access";Break}
					"RestrictedAccess"			{$NapStatus = "Restricted Access";Break}
					"DropPacket"				{$NapStatus = "Drop Packet";Break}
					"InProbation"				{$NapStatus = "In Probation";Break}
					"Exempt"					{$NapStatus = "Exempt";Break}
					"DefaultQuarantineSetting"	{$NapStatus = "Default Quarantine Setting";Break}
					"NoQuarantineInfo"			{$NapStatus = "No Quarantine Info";Break}
					Default						{$NapStatus = "Unable to determine NAP Status: $($Lease.NapStatus)";Break}
				}
				
				If($MSWord -or $PDF)
				{
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$ScriptInformation += @{ Data = "Client IP address"; Value = $Lease.IPAddress.ToString(); }
					$ScriptInformation += @{ Data = "Name"; Value = $Lease.HostName; }
					
					If([string]::IsNullOrEmpty($Lease.LeaseExpiryTime))
					{
						If($Lease.AddressState -eq "ActiveReservation")
						{
							$ScriptInformation += @{ Data = "Lease Expiration"; Value = "Reservation (active)"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = "Lease Expiration"; Value = "Reservation (inactive)"; }
						}
					}
					Else
					{
						$ScriptInformation += @{ Data = "Lease Expiration"; Value = $LeaseStr; }
					}
					
					$ScriptInformation += @{ Data = "Type"; Value = $Lease.ClientType; }
					$ScriptInformation += @{ Data = "Unique ID"; Value = $Lease.ClientID; }
					$ScriptInformation += @{ Data = "Description"; Value = $Lease.Description; }
					$ScriptInformation += @{ Data = "Network Access Protection"; Value = $NapStatus; }

					If([string]::IsNullOrEmpty($Lease.ProbationEnds))
					{
						$ScriptInformation += @{ Data = "Probation Expiration"; Value = "N/A"; }
					}
					Else
					{
						$ScriptInformation += @{ Data = "Probation Expiration"; Value = $ProbationStr; }
					}
					
					$Index = $Null
					ForEach($Filter in $Filters)
					{
						If( (ValidObject $Filter MacAddress) -and ($Filter.MacAddress -eq $Lease.ClientID) )
						{
							$Index = $Filter
						}
					}
					
					If($Null -ne $Index)
					{
						$ScriptInformation += @{ Data = "Filter"; Value = $Index.List; }
					}
					Else
					{
						$ScriptInformation += @{ Data = "Filter"; Value = "None"; }
					}

					If([string]::IsNullOrEmpty($Lease.PolicyName))
					{
						$ScriptInformation += @{ Data = "Policy"; Value = "None"; }
					}
					Else
					{
						$ScriptInformation += @{ Data = "Policy"; Value = $Lease.PolicyName; }
					}
					
					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 150;
					$Table.Columns.Item(2).Width = 275;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 2 "Name: " $Lease.HostName
					Line 2 "Client IP address`t`t: " $Lease.IPAddress
					Line 2 "Lease Expiration`t`t: " -NoNewLine
					If([string]::IsNullOrEmpty($Lease.LeaseExpiryTime))
					{
						If($Lease.AddressState -eq "ActiveReservation")
						{
							Line 0 "Reservation (active)"
						}
						Else
						{
							Line 0 "Reservation (inactive)"
						}
					}
					Else
					{
						Line 0 $LeaseStr
					}
					Line 2 "Type`t`t`t`t: " $Lease.ClientType
					Line 2 "Unique ID`t`t`t: " $Lease.ClientID
					Line 2 "Description`t`t`t: " $Lease.Description
					Line 2 "Network Access Protection`t: " $NapStatus
					Line 2 "Probation Expiration`t`t: " -NoNewLine
					If([string]::IsNullOrEmpty($Lease.ProbationEnds))
					{
						Line 0 "N/A"
					}
					Else
					{
						Line 0 $ProbationStr
					}
					Line 2 "Filter`t`t`t`t: " -NoNewLine
					
					$Index = $Null
					ForEach($Filter in $Filters)
					{
						If( (ValidObject $Filter MacAddress) -and ($Filter.MacAddress -eq $Lease.ClientID) )
						{
							$Index = $Filter
						}
					}
					
					If($Null -ne $Index)
					{
						Line 0 $Index.List
					}
					Else
					{
						Line 0 "None"
					}
					Line 2 "Policy`t`t`t`t: " -NoNewLine
					If([string]::IsNullOrEmpty($Lease.PolicyName))
					{
						Line 0 "None"
					}
					Else
					{
						Line 0 $Lease.PolicyName
					}
					
					#skip a row for spacing
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					If($Null -ne $Lease.LeaseExpiryTime)
					{
						$LeaseStr = [string]::format("{0} days, {1} hours, {2} minutes", `
							$Lease.LeaseExpiryTime.Day, `
							$Lease.LeaseExpiryTime.Hour, `
							$Lease.LeaseExpiryTime.Minute)
					}
					Else
					{
						$LeaseStr = ""
					}

					If($Null -ne $Lease.ProbationEnds)
					{
						$ProbationStr = [string]::format("{0} days, {1} hours, {2} minutes", `
							$Lease.ProbationEnds.Day, `
							$Lease.ProbationEnds.Hour, `
							$Lease.ProbationEnds.Minute)
					}
					Else
					{
						$ProbationStr = ""
					}

					Switch($Lease.NapStatus)
					{
						"FullAccess"				{$NapStatus = "Full Access";Break}
						"RestrictedAccess"			{$NapStatus = "Restricted Access";Break}
						"DropPacket"				{$NapStatus = "Drop Packet";Break}
						"InProbation"				{$NapStatus = "In Probation";Break}
						"Exempt"					{$NapStatus = "Exempt";Break}
						"DefaultQuarantineSetting"	{$NapStatus = "Default Quarantine Setting";Break}
						"NoQuarantineInfo"			{$NapStatus = "No Quarantine Info";Break}
						Default						{$NapStatus = "Unable to determine NAP Status: $($Lease.NapStatus)";Break}
					}
					
					$columnHeaders = @("Client IP address",($htmlsilver -bor $htmlbold),$Lease.IPAddress.ToString(),$htmlwhite)
					
					$rowdata += @(,('Name',($htmlsilver -bor $htmlbold),$Lease.HostName,$htmlwhite))
					
					$tmp = ""
					If([string]::IsNullOrEmpty($Lease.LeaseExpiryTime))
					{
						If($Lease.AddressState -eq "ActiveReservation")
						{
							$tmp = "Reservation (active)"
						}
						Else
						{
							$tmp = "Reservation (inactive)"
						}
					}
					Else
					{
						$tmp = $LeaseStr
					}
					$rowdata += @(,('Lease Expiration',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					$rowdata += @(,('Type',($htmlsilver -bor $htmlbold),$Lease.ClientType,$htmlwhite))
					$rowdata += @(,('Unique ID',($htmlsilver -bor $htmlbold),$Lease.ClientID,$htmlwhite))
					$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Lease.Description,$htmlwhite))
					$rowdata += @(,('Network Access Protection',($htmlsilver -bor $htmlbold),$NapStatus,$htmlwhite))

					$tmp = ""
					If([string]::IsNullOrEmpty($Lease.ProbationEnds))
					{
						$tmp = "N/A"
					}
					Else
					{
						$tmp = $ProbationStr
					}
					$rowdata += @(,('Probation Expiration',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					
					#$Filters | ForEach-Object { $index = $Null }{ If( $_.MacAddress -eq $Lease.ClientID ) { $index = $_ } }
					
					$Index = $Null
					ForEach($Filter in $Filters)
					{
						If( (ValidObject $Filter MacAddress) -and ($Filter.MacAddress -eq $Lease.ClientID) )
						{
							$Index = $Filter
						}
					}
					
					$tmp = ""
					If($Null -ne $Index)
					{
						$tmp = $Index.List
					}
					Else
					{
						$tmp = "None"
					}
					$rowdata += @(,('Filter',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))

					$tmp = ""
					If([string]::IsNullOrEmpty($Lease.PolicyName))
					{
						$tmp = "None"
					}
					Else
					{
						$tmp = $Lease.PolicyName
					}
					$rowdata += @(,('Policy',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					
					$msg = ""
					$columnWidths = @("150","200")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
					WriteHTMLLine 0 0 ""
				}
			}
		}
		ElseIf(!$?)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "Error retrieving leases for scope $IPv4Scope.ScopeId"
			}
			If($Text)
			{
				Line 0 "Error retrieving leases for scope $IPv4Scope.ScopeId"
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "Error retrieving leases for scope $IPv4Scope.ScopeId"
			}
			InsertBlankLine
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "None"
			}
			If($Text)
			{
				Line 0 "None"
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "None"
			}
			InsertBlankLine
		}
		$Leases = $Null
	}

	Write-Verbose "$(Get-Date -Format G):	`t`tGetting exclusions"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine ($xStartLevel + 1) 0 "Exclusions"
		[System.Collections.Hashtable[]] $ExclusionsWordTable = @()
	}
	If($Text)
	{
		Line 1 "Exclusions:"
	}
	If($HTML)
	{
		WriteHTMLLine ($xStartLevel + 1) 0 "Exclusions"
		$rowdata = @()
	}
	
	$Exclusions = Get-DHCPServerV4ExclusionRange -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object StartRange
	If($? -and $Null -ne $Exclusions)
	{
		ForEach($Exclusion in $Exclusions)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				Start = $Exclusion.StartRange.ToString(); `
				Ending = $Exclusion.EndRange.ToString()
				}

				## Add the hash to the array
				$ExclusionsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 2 "Start IP Address`t: " $Exclusion.StartRange.ToString()
				Line 2 "End IP Address`t`t: " $Exclusion.EndRange.ToString() 
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata += @(,($Exclusion.StartRange.ToString(),$htmlwhite,
								$Exclusion.EndRange.ToString(),$htmlwhite))
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			Start = "Error retrieving exclusions for scope $IPv4Scope.ScopeId"; `
			Ending = ""
			}

			## Add the hash to the array
			$ExclusionsWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 0 "Error retrieving exclusions for scope $IPv4Scope.ScopeId"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,("Error retrieving exclusions for scope $IPv4Scope.ScopeId",$htmlwhite,
							"",$htmlwhite))
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			Start = "None"; `
			Ending = "None"
			}

			## Add the hash to the array
			$ExclusionsWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 2 "None"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,("None",$htmlwhite,
							"None",$htmlwhite))
		}
	}
	$Exclusions = $Null
	
	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		If($ExclusionsWordTable.Count -gt 0)
		{
			$Table = AddWordTable -Hashtable $ExclusionsWordTable `
			-Columns Start,Ending `
			-Headers "Start IP Address","End IP Address" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	If($HTML)
	{
		$columnHeaders = @('Start IP Address',($htmlsilver -bor $htmlbold),'End IP Address',($htmlsilver -bor $htmlbold))
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}

	Write-Verbose "$(Get-Date -Format G):	`t`tGetting reservations"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine ($xStartLevel + 1) 0 "Reservations"
	}
	If($Text)
	{
		Line 1 "Reservations:"
	}
	If($HTML)
	{
		WriteHTMLLine ($xStartLevel + 1) 0 "Reservations"
	}
	
	$Reservations = Get-DHCPServerV4Reservation -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object Name
	If($? -and $Null -ne $Reservations)
	{
		ForEach($Reservation in $Reservations)
		{
			Write-Verbose "$(Get-Date -Format G):	`t`t`tProcessing reservation $($Reservation.Name)"
			
			If($MSWord -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Reservation name"; Value = $Reservation.Name; }
				$ScriptInformation += @{ Data = "IP address"; Value = $Reservation.IPAddress.ToString(); }
				$ScriptInformation += @{ Data = "MAC address"; Value = $Reservation.ClientId; }
				$ScriptInformation += @{ Data = "Supported types"; Value = $Reservation.Type; }
				$ScriptInformation += @{ Data = "Description"; Value = $Reservation.Description; }

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
				$Table.Columns.Item(1).Width = 100;
				$Table.Columns.Item(2).Width = 175;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				Write-Verbose "$(Get-Date -Format G):	`t`t`t`tGetting DNS settings"
				$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $Script:DHCPServerName -IPAddress $Reservation.IPAddress -EA 0
				If($? -and $Null -ne $DNSSettings)
				{
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					If($DNSSettings.DynamicUpdates -eq "Never")
					{
						$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Disabled"; }
					}
					ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
					{
						$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
						$ScriptInformation += @{ Data = "Dynamically update DNS A and PTR records only if requested by the DHCP clients"; Value = "Enabled"; }
					}
					ElseIf($DNSSettings.DynamicUpdates -eq "Always")
					{
						$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
						$ScriptInformation += @{ Data = "Always dynamically update DNS A and PTR records"; Value = "Enabled"; }
					}
					If($DNSSettings.DeleteDnsRROnLeaseExpiry)
					{
						$ScriptInformation += @{ Data = "Discard A and PTR records when lease is deleted"; Value = "Enabled"; }
					}
					Else
					{
						$ScriptInformation += @{ Data = "Discard A and PTR records when lease is deleted"; Value = "Disabled"; }
					}
					If($DNSSettings.UpdateDnsRRForOlderClients)
					{
						$Tmp = "Enabled"
					}
					Else
					{
						$Tmp = "Disabled"
					}
					$ScriptInformation += @{ Data = "Dynamically update DNS records for DHCP clients that do not request updates"; Value = $Tmp; }
					If($DNSSettings.DisableDnsPtrRRUpdate)
					{
						$ScriptInformation += @{ Data = "Disable dynamic updates for DNS PTR record"; Value = "Enabled"; }
					}
					Else
					{
						$ScriptInformation += @{ Data = "Disable dynamic updates for DNS PTR record"; Value = "Disabled"; }
					}
					If($DNSSettings.NameProtection)
					{
						$Tmp = "Enabled"
					}
					Else
					{
						$Tmp = "Disabled"
					}
					$ScriptInformation += @{ Data = "Name Protection"; Value = $Tmp; }

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 400;
					$Table.Columns.Item(2).Width = 50;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				Else
				{
					WriteWordLine 0 0 "Error retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
					WriteWordLine 0 0 ""
				}
				$DNSSettings = $Null
			}
			If($Text)
			{
				Line 2 "Reservation name`t: " $Reservation.Name
				Line 2 "IP address`t`t: " $Reservation.IPAddress
				Line 2 "MAC address`t`t: " $Reservation.ClientId
				Line 2 "Supported types`t`t: " $Reservation.Type
				Line 2 "Description`t`t: " $Reservation.Description
				Line 0 ""

				Write-Verbose "$(Get-Date -Format G): `t`t`t`tGetting DNS settings"
				$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $Script:DHCPServerName -IPAddress $Reservation.IPAddress -EA 0
				If($? -and $Null -ne $DNSSettings)
				{
					Line 2 "Enable DNS dynamic updates`t`t`t: " -NoNewLine
					If($DNSSettings.DynamicUpdates -eq "Never")
					{
						Line 0 "Disabled"
					}
					ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
					{
						Line 0 "Enabled"
						Line 2 "Dynamically update DNS A and PTR records only "
						Line 2 "if requested by the DHCP clients`t`t: Enabled"
					}
					ElseIf($DNSSettings.DynamicUpdates -eq "Always")
					{
						Line 0 "Enabled"
						Line 2 "Always dynamically update DNS A and PTR records: Enabled"
					}
					Line 2 "Discard A and PTR records when lease deleted`t: " -NoNewLine
					If($DNSSettings.DeleteDnsRROnLeaseExpiry)
					{
						Line 0 "Enabled"
					}
					Else
					{
						Line 0 "Disabled"
					}
					Line 2 "Dynamically update DNS records for DHCP "
					Line 2 "clients that do not request updates`t`t: " -NoNewLine
					If($DNSSettings.UpdateDnsRRForOlderClients)
					{
						Line 0 "Enabled"
					}
					Else
					{
						Line 0 "Disabled"
					}
					Line 2 "Disable dynamic updates for DNS PTR records`t: " -NoNewLine
					If($DNSSettings.DisableDnsPtrRRUpdate)
					{
						Line 0 "Enabled"
					}
					Else
					{
						Line 0 "Disabled"
					}
					Line 2 "Name Protection`t`t`t`t`t: " -NoNewLine
					If($DNSSettings.NameProtection)
					{
						Line 0 "Enabled"
					}
					Else
					{
						Line 0 "Disabled"
					}
					Line 0 ""
				}
				Else
				{
					Line 0 "Error retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
					Line 0 ""
				}
				$DNSSettings = $Null
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Reservation name",($htmlsilver -bor $htmlbold),$Reservation.Name,$htmlwhite)
				$rowdata += @(,('IP address',($htmlsilver -bor $htmlbold),$Reservation.IPAddress.ToString(),$htmlwhite))
				$rowdata += @(,('MAC address',($htmlsilver -bor $htmlbold),$Reservation.ClientId,$htmlwhite))
				$rowdata += @(,('Supported types',($htmlsilver -bor $htmlbold),$Reservation.Type,$htmlwhite))
				$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Reservation.Description,$htmlwhite))
				$msg = ""
				$columnWidths = @("150","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
				InsertBlankLine

				Write-Verbose "$(Get-Date -Format G):	`t`t`t`tGetting DNS settings"
				$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $Script:DHCPServerName -IPAddress $Reservation.IPAddress -EA 0
				If($? -and $Null -ne $DNSSettings)
				{
					$rowdata = @()
					If($DNSSettings.DynamicUpdates -eq "Never")
					{
						$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite)
					}
					ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
					{
						$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
						$rowdata += @(,("Dynamically update DNS A and PTR records only if requested by the DHCP clients",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
					}
					ElseIf($DNSSettings.DynamicUpdates -eq "Always")
					{
						$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
						$rowdata += @(,("Always dynamically update DNS A and PTR records",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
					}
					If($DNSSettings.DeleteDnsRROnLeaseExpiry)
					{
						$rowdata += @(,("Discard A and PTR records when lease deleted",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
					}
					Else
					{
						$rowdata += @(,("Discard A and PTR records when lease deleted",($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
					}
					If($DNSSettings.UpdateDnsRRForOlderClients)
					{
						$rowdata += @(,('Dynamically update DNS records for DHCP clients that do not request updates',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('Dynamically update DNS records for DHCP clients that do not request updates',($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
					}
					If($DNSSettings.DisableDnsPtrRRUpdate)
					{
						$rowdata += @(,('Disable dynamic updates for DNS PTR records',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('Disable dynamic updates for DNS PTR records',($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
					}
					If($DNSSettings.NameProtection)
					{
						$rowdata += @(,('Name Protection',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('Name Protection',($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
					}
					$msg = ""
					$columnWidths = @("450","50")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
					WriteHTMLLine 0 0 
				}
				Else
				{
					WriteHTMLLine 0 0 "Error retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
					WriteHTMLLine 0 0 ""
				}
				$DNSSettings = $Null
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving reservations for scope $IPv4Scope.ScopeId"
		}
		If($Text)
		{
			Line 0 "Error retrieving reservations for scope $IPv4Scope.ScopeId"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving reservations for scope $IPv4Scope.ScopeId"
		}
		InsertBlankLine
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "None"
		}
		If($Text)
		{
			Line 2 "None"
		}
		If($HTML)
		{
			WriteHTMLLine 0 1 "None"
		}
		InsertBlankLine
	}
	$Reservations = $Null

	Write-Verbose "$(Get-Date -Format G):	`t`tGetting scope options"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine ($xStartLevel + 1) 0 "Scope Options"
	}
	If($Text)
	{
		Line 1 "Scope Options:"
	}
	If($HTML)
	{
		WriteHTMLLine ($xStartLevel + 1) 0 "Scope Options"
	}
	
	$ScopeOptions = Get-DHCPServerV4OptionValue -All -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object OptionId

	If($? -and $Null -ne $ScopeOptions)
	{
		If($ScopeOptions -is [array] -and $ScopeOptions.Count -eq 2 -and $ScopeOptions[0].OptionId -eq 51 -and $ScopeOptions[1].OptionId -eq 81)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "None"
			}
			If($Text)
			{
				Line 2 "None"
			}
			If($HTML)
			{
				WriteHTMLLine 0 1 "None"
			}
		}
		Else
		{
			ForEach($ScopeOption in $ScopeOptions)
			{
				If($MSWord -or $PDF)
				{
					[System.Collections.Hashtable[]] $ScriptInformation = @()
				}
				If($HTML)
				{
					$rowdata = @()
				}
				
				If($ScopeOption.OptionId -eq 51 -or $ScopeOption.OptionId -eq 81)
				{
					#ignore these two option IDs
					#https://carlwebster.com/the-mysterious-microsoft-dhcp-option-id-81/
					#https://jimswirelessworld.wordpress.com/2019/01/03/you-should-care-about-dhcp-option-51/
				}
				Else
				{
					Write-Verbose "$(Get-Date -Format G):	`t`t`tProcessing option name $($ScopeOption.Name)"
					If([string]::IsNullOrEmpty($ScopeOption.VendorClass))
					{
						$VendorClass = "Standard" 
					}
					Else
					{
						$VendorClass = $ScopeOption.VendorClass 
					}

					If([string]::IsNullOrEmpty($ScopeOption.PolicyName))
					{
						$PolicyName = "None"
					}
					Else
					{
						$PolicyName = $ScopeOption.PolicyName
					}

					If($MSWord -or $PDF)
					{
						$ScriptInformation += @{ Data = "Option Name"; Value = "$($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)"; }
						$ScriptInformation += @{ Data = "Vendor"; Value = $VendorClass; }
						$ScriptInformation += @{ Data = "Value"; Value = "$($ScopeOption.Value)"; }
						$ScriptInformation += @{ Data = "Policy Name"; Value = $PolicyName; }

						$Table = AddWordTable -Hashtable $ScriptInformation `
						-Columns Data,Value `
						-List `
						-Format $wdTableGrid `
						-AutoFit $wdAutoFitFixed;

						SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
						
						$Table.Columns.Item(1).Width = 75;
						$Table.Columns.Item(2).Width = 300;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

						FindWordDocumentEnd
						$Table = $Null
						WriteWordLine 0 0 ""
					}
					If($Text)
					{
						Line 2 "Option Name`t: $($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)" 
						Line 2 "Vendor`t`t: " $VendorClass
						Line 2 "Value`t`t: $($ScopeOption.Value)" 
						Line 2 "Policy Name`t: " $PolicyName
						
						#for spacing
						Line 0 ""
					}
					If($HTML)
					{
						$columnHeaders = @("Option Name",($htmlsilver -bor $htmlbold),"$($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)",$htmlwhite)
						$rowdata += @(,('Vendor',($htmlsilver -bor $htmlbold),$VendorClass,$htmlwhite))
						$rowdata += @(,('Value',($htmlsilver -bor $htmlbold),"$($ScopeOption.Value)",$htmlwhite))
						$rowdata += @(,('Policy Name',($htmlsilver -bor $htmlbold),$PolicyName,$htmlwhite))
					
						$msg = ""
						$columnWidths = @("100","400")
						FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
						WriteHTMLLine 0 0 ""
					}
				}
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving scope options for $IPv4Scope.ScopeId"
		}
		If($Text)
		{
			Line 0 "Error retrieving scope options for $IPv4Scope.ScopeId"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving scope options for $IPv4Scope.ScopeId"
		}
		InsertBlankLine
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "None"
		}
		If($Text)
		{
			Line 2 "None"
		}
		If($HTML)
		{
			WriteHTMLLine 0 1 "None"
		}
		InsertBlankLine
	}
	$ScopeOptions = $Null
	
	Write-Verbose "$(Get-Date -Format G):	`t`tGetting policies"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine ($xStartLevel + 1) 0 "Policies"
	}
	If($Text)
	{
		Line 1 "Policies:"
	}
	If($HTML)
	{
		WriteHTMLLine ($xStartLevel + 1) 0 "Policies"
	}

	#V1.45 add all the missing policy data
	$ScopePolicies = Get-DHCPServerV4Policy -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object ProcessingOrder

	If($? -and $Null -ne $ScopePolicies)
	{
		ForEach($Policy in $ScopePolicies)
		{
			Write-Verbose "$(Get-Date -Format G):	`t`t`tProcessing policy name $($Policy.Name)"
			If($Policy.Enabled)
			{
				$PolicyEnabled = "Enabled"
			}
			Else
			{
				$PolicyEnabled = "Disabled"
			}

			$LeaseDuration = [string]::format("{0} days, {1} hours, {2} minutes", `
				$Policy.LeaseDuration.Days, `
				$Policy.LeaseDuration.Hours, `
				$Policy.LeaseDuration.Minutes)

			If($MSWord -or $PDF)
			{
				WriteWordLine ($xStartLevel + 2) 0 "General"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Policy Name"; Value = $Policy.Name; }
				$ScriptInformation += @{ Data = "Description"; Value = $Policy.Description; }
				$ScriptInformation += @{ Data = "Processing Order"; Value = $Policy.ProcessingOrder.ToString(); }
				$ScriptInformation += @{ Data = "Level"; Value = "Scope"; }
				$ScriptInformation += @{ Data = "State"; Value = $PolicyEnabled; }

				If($Policy.LeaseDuration.ToString() -eq "00:00:00")	#lease duration is not set
				{
					$ScriptInformation += @{ Data = "Set lease duration for the policy"; Value = "Not selected"; }
				}
				Else
				{
					#lease duration is set
					$ScriptInformation += @{ Data = "Set lease duration for the policy"; Value = "Selected"; }
					If($Policy.LeaseDuration.ToString() -eq "10675199.02:48:05.4775807") #unlimited
					{
						$ScriptInformation += @{ Data = "Lease duration for DHCP clients"; Value = "Unlimited"; }
					}
					Else
					{
						$ScriptInformation += @{ Data = "Lease duration for DHCP clients"; Value = $LeaseDuration; }
					}
				}

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 250;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 "General"
				Line 3 "Policy Name`t`t`t`t: " $Policy.Name
				Line 3 "Description`t`t`t`t: " $Policy.Description
				Line 3 "Processing Order`t`t`t: " $Policy.ProcessingOrder
				Line 3 "Level`t`t`t`t`t: Scope"
				Line 3 "State`t`t`t`t`t: " $PolicyEnabled
				If($Policy.LeaseDuration.ToString() -eq "00:00:00")	#lease duration is not set
				{
					Line 3 "Set lease duration for the policy`t: Not selected" 
				}
				Else
				{
					#lease duration is set
					Line 3 "Set lease duration for the policy`t: Selected" 
					If($Policy.LeaseDuration.ToString() -eq "10675199.02:48:05.4775807") #unlimited
					{
						Line 3 "Lease duration for DHCP clients`t`t: Unlimited"
					}
					Else
					{
						Line 3 "Lease duration for DHCP clients`t`t: " $LeaseDuration
					}
				}
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Policy Name",($htmlsilver -bor $htmlbold),$Policy.Name,$htmlwhite)
				$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Policy.Description,$htmlwhite))
				$rowdata += @(,('Processing Order',($htmlsilver -bor $htmlbold),$Policy.ProcessingOrder.ToString(),$htmlwhite))
				$rowdata += @(,('Level',($htmlsilver -bor $htmlbold),"Scope",$htmlwhite))
				$rowdata += @(,('State',($htmlsilver -bor $htmlbold),$PolicyEnabled,$htmlwhite))
				If($Policy.LeaseDuration.ToString() -eq "00:00:00")	#lease duration is not set
				{
					$rowdata += @(,('Set lease duration for the policy',($htmlsilver -bor $htmlbold),"Not selected" ,$htmlwhite))
				}
				Else
				{
					#lease duration is set
					$rowdata += @(,('Set lease duration for the policy',($htmlsilver -bor $htmlbold),"Selected" ,$htmlwhite))
					If($Policy.LeaseDuration.ToString() -eq "10675199.02:48:05.4775807") #unlimited
					{
						$rowdata += @(,('Lease duration for DHCP clients',($htmlsilver -bor $htmlbold),"Unlimited",$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('Lease duration for DHCP clients',($htmlsilver -bor $htmlbold),$LeaseDuration,$htmlwhite))
					}
				}
				$msg = ""
				$columnWidths = @("250","250")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
				WriteHTMLLine 0 0 ""
			}

			If($MSWord -or $PDF)
			{
				WriteWordLine ($xStartLevel + 2) 0 "Conditions"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Condition"; Value = $Policy.Condition; }
			}
			If($Text)
			{
				Line 2 "Conditions"
				Line 3 "Condition: " $Policy.Condition
				Line 3 "Conditions                                  Operator    Value                                   "
				Line 3 "================================================================================================"
					   #123456789012345678901234567890123456789012SS1234567890SS1234567890123456789012345678901234567890
					   #Relay Agent Information - Agent Circuit Id
			}
			If($HTML)
			{
				$rowdata = @()
				WriteHTMLLine ($xStartLevel + 2) 0 "Conditions"
				$columnHeaders = @("Condition",($htmlsilver -bor $htmlbold),$Policy.Condition,$htmlwhite)
			}
			
			If(![string]::IsNullOrEmpty($Policy.VendorClass))
			{
				For( $i = 0; $i -lt $Policy.VendorClass.Length; $i += 2 )
				{
					$xCond  = "Vendor Class"
					If($Policy.VendorClass[ $i ] -eq "EQ")
					{
						$xOper = "Equals" 
					}
					Else
					{
						$xOper = "Not Equals" 
					}
					$xValue = $Policy.VendorClass[ $i + 1 ]

					If($MSWord -or $PDF)
					{
						$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
					}
					If($Text)
					{
						Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
					}
					If($HTML)
					{
						$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
					}
				}
			}
			If(![string]::IsNullOrEmpty($Policy.UserClass))
			{
				For( $i = 0; $i -lt $Policy.UserClass.Length; $i += 2 )
				{
					$xCond  = "User Class"
					If($Policy.UserClass[ $i ] -eq "EQ")
					{
						$xOper = "Equals" 
					}
					Else
					{
						$xOper = "Not Equals" 
					}
					$xValue = $Policy.UserClass[ $i + 1 ]

					If($MSWord -or $PDF)
					{
						$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
					}
					If($Text)
					{
						Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
					}
					If($HTML)
					{
						$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
					}
				}
			}
			If(![string]::IsNullOrEmpty($Policy.MacAddress))
			{
				For( $i = 0; $i -lt $Policy.MacAddress.Length; $i += 2 )
				{
					$xCond  = "MAC Address"
					If($Policy.MacAddress[ $i ] -eq "EQ")
					{
						$xOper = "Equals" 
					}
					Else
					{
						$xOper = "Not Equals" 
					}
					$xValue = $Policy.MacAddress[ $i + 1 ]

					If($MSWord -or $PDF)
					{
						$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
					}
					If($Text)
					{
						Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
					}
					If($HTML)
					{
						$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
					}
				}
			}
			If(![string]::IsNullOrEmpty($Policy.ClientId))
			{
				For( $i = 0; $i -lt $Policy.ClientId.Length; $i += 2 )
				{
					$xCond  = "Client Identifier"
					If($Policy.ClientId[ $i ] -eq "EQ")
					{
						$xOper = "Equals" 
					}
					Else
					{
						$xOper = "Not Equals" 
					}
					$xValue = $Policy.ClientId[ $i + 1 ]

					If($MSWord -or $PDF)
					{
						$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
					}
					If($Text)
					{
						Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
					}
					If($HTML)
					{
						$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
					}
				}
			}
			If(![string]::IsNullOrEmpty($Policy.Fqdn))
			{
				For( $i = 0; $i -lt $Policy.Fqdn.Length; $i += 2 )
				{
					$xCond  = "Fully Qualified Domain Name"
					If($Policy.Fqdn[ $i ] -eq "EQ")
					{
						$xOper = "Equals" 
					}
					Else
					{
						$xOper = "Not Equals" 
					}
					$xValue = $Policy.Fqdn[ $i + 1 ]

					If($MSWord -or $PDF)
					{
						$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
					}
					If($Text)
					{
						Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
					}
					If($HTML)
					{
						$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
					}
				}
			}
			If(![string]::IsNullOrEmpty($Policy.RelayAgent))
			{
				For( $i = 0; $i -lt $Policy.RelayAgent.Length; $i += 2 )
				{
					$xCond  = "Relay Agent Information"
					If($Policy.RelayAgent[ $i ] -eq "EQ")
					{
						$xOper = "Equals" 
					}
					Else
					{
						$xOper = "Not Equals" 
					}
					$xValue = $Policy.RelayAgent[ $i + 1 ]

					If($MSWord -or $PDF)
					{
						$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
					}
					If($Text)
					{
						Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
					}
					If($HTML)
					{
						$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
					}
				}
			}
			If(![string]::IsNullOrEmpty($Policy.CircuitId))
			{
				For( $i = 0; $i -lt $Policy.CircuitId.Length; $i += 2 )
				{
					$xCond  = "Relay Agent Information - Agent Circuit Id"
					If($Policy.CircuitId[ $i ] -eq "EQ")
					{
						$xOper = "Equals" 
					}
					Else
					{
						$xOper = "Not Equals" 
					}
					$xValue = $Policy.CircuitId[ $i + 1 ]

					If($MSWord -or $PDF)
					{
						$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
					}
					If($Text)
					{
						Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
					}
					If($HTML)
					{
						$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
					}
				}
			}
			If(![string]::IsNullOrEmpty($Policy.RemoteId))
			{
				For( $i = 0; $i -lt $Policy.RemoteId.Length; $i += 2 )
				{
					$xCond  = "Relay Agent Information - Agent Remote Id"
					If($Policy.RemoteId[ $i ] -eq "EQ")
					{
						$xOper = "Equals" 
					}
					Else
					{
						$xOper = "Not Equals" 
					}
					$xValue = $Policy.RemoteId[ $i + 1 ]

					If($MSWord -or $PDF)
					{
						$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
					}
					If($Text)
					{
						Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
					}
					If($HTML)
					{
						$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
					}
				}
			}
			If(![string]::IsNullOrEmpty($Policy.SubscriberId))
			{
				For( $i = 0; $i -lt $Policy.SubscriberId.Length; $i += 2 )
				{
					$xCond  = "Relay Agent Information - Subscriber Id"
					If($Policy.SubscriberId[ $i ] -eq "EQ")
					{
						$xOper = "Equals" 
					}
					Else
					{
						$xOper = "Not Equals" 
					}
					$xValue = $Policy.SubscriberId[ $i + 1 ]

					If($MSWord -or $PDF)
					{
						$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
					}
					If($Text)
					{
						Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
					}
					If($HTML)
					{
						$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
					}
				}
			}

			If($MSWord -or $PDF)
			{
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 250;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 0 ""
			}
			If($HTML)
			{
				$msg = ""
				$columnWidths = @("250","250")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
				WriteHTMLLine 0 0 ""
			}

			If($MSWord -or $PDF)
			{
				WriteWordLine ($xStartLevel + 2) 0 "IP Address Range"
				[System.Collections.Hashtable[]] $ScriptInformation = @()
			}
			If($Text)
			{
				Line 2 "IP Address Range"
			}
			If($HTML)
			{
				WriteHTMLLine ($xStartLevel + 2) 0 "IP Address Range"
				$rowdata = @()
				$cnt = -1
			}
			
			$IPRanges = Get-DHCPServerV4PolicyIPRange -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -Name $Policy.Name -EA 0

			If($? -and $Null -ne $IPRanges)
			{
				ForEach($IPRange in $IPRanges)
				{
					If($MSWord -or $PDF)
					{
						$ScriptInformation += @{ Data = "Address Range"; Value = "$($IPRange.StartRange) - $($IPRange.EndRange)"; }
					}
					If($Text)
					{
						Line 3 "$($IPRange.StartRange) - $($IPRange.EndRange)"
					}
					If($HTML)
					{
						$cnt++
						If($cnt -eq 0)
						{
							$columnHeaders = @("Address Range",($htmlsilver -bor $htmlbold),"$($IPRange.StartRange) - $($IPRange.EndRange)",$htmlwhite)
						}
						Else
						{
							$rowdata += @(,('Address Range',($htmlsilver -bor $htmlbold),"$($IPRange.StartRange) - $($IPRange.EndRange)",$htmlwhite))
						}
					}
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					$ScriptInformation += @{ Data = "Address Range"; Value = "None"; }
				}
				If($Text)
				{
					Line 3 "None"
				}
				If($HTML)
				{
					$rowdata += @(,('Address Range',($htmlsilver -bor $htmlbold),"None",$htmlwhite))
				}
			}

			$IPRanges = $Null
			If($MSWord -or $PDF)
			{
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 250;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 0 ""
			}
			If($HTML)
			{
				$msg = ""
				$columnWidths = @("250","250")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
				WriteHTMLLine 0 0 ""
			}

			If($MSWord -or $PDF)
			{
				WriteWordLine ($xStartLevel + 2) 0 "Options"
			}
			If($Text)
			{
				Line 2 "Options"
			}
			If($HTML)
			{
				WriteHTMLLine ($xStartLevel + 2) 0 "Options"
			}

			$ScopePolicyOptions = Get-DHCPServerV4OptionValue -ComputerName $Script:DHCPServerName -PolicyName $Policy.Name -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object OptionId

			If($? -and $Null -ne $ScopePolicyOptions)
			{
				If($ScopePolicyOptions -is [array] -and $ScopePolicyOptions.Count -eq 2 -and $ScopePolicyOptions[0].OptionId -eq 51 -and $ScopePolicyOptions[1].OptionId -eq 81)
				{
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 0 "None"
					}
					If($Text)
					{
						Line 3 "None"
					}
					If($HTML)
					{
						WriteHTMLLine 0 0 "None"
					}
				}
				Else
				{
					ForEach($ScopePolicyOption in $ScopePolicyOptions)
					{
						If($ScopePolicyOption.OptionId -eq 51 -or $ScopePolicyOption.OptionId -eq 81)
						{
							#ignore these two option IDs
							#https://carlwebster.com/the-mysterious-microsoft-dhcp-option-id-81/
							#https://jimswirelessworld.wordpress.com/2019/01/03/you-should-care-about-dhcp-option-51/
						}
						Else
						{
							Write-Verbose "$(Get-Date -Format G):	`t`t`tProcessing option name $($ScopePolicyOption.Name)"
							If([string]::IsNullOrEmpty($ScopePolicyOption.VendorClass))
							{
								$VendorClass = "Standard" 
							}
							Else
							{
								$VendorClass = $ScopePolicyOption.VendorClass 
							}

							If([string]::IsNullOrEmpty($ScopePolicyOption.PolicyName))
							{
								$PolicyName = "None"
							}
							Else
							{
								$PolicyName = $ScopePolicyOption.PolicyName
							}

							If($MSWord -or $PDF)
							{
								[System.Collections.Hashtable[]] $ScriptInformation = @()
								$ScriptInformation += @{ Data = "Option Name"; Value = "$($ScopePolicyOption.OptionId.ToString("00000")) $($ScopePolicyOption.Name)"; }
								$ScriptInformation += @{ Data = "Vendor"; Value = $VendorClass; }
								$ScriptInformation += @{ Data = "Value"; Value = "$($ScopePolicyOption.Value)"; }
								$ScriptInformation += @{ Data = "Policy Name"; Value = $PolicyName; }

								$Table = AddWordTable -Hashtable $ScriptInformation `
								-Columns Data,Value `
								-List `
								-Format $wdTableGrid `
								-AutoFit $wdAutoFitFixed;

								SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
								
								$Table.Columns.Item(1).Width = 75;
								$Table.Columns.Item(2).Width = 300;

								$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

								FindWordDocumentEnd
								$Table = $Null
								WriteWordLine 0 0 ""
							}
							If($Text)
							{
								Line 3 "Option Name`t: $($ScopePolicyOption.OptionId.ToString("00000")) $($ScopePolicyOption.Name)" 
								Line 3 "Vendor`t`t: " $VendorClass
								Line 3 "Value`t`t: $($ScopePolicyOption.Value)" 
								Line 3 "Policy Name`t: " $PolicyName
								
								#for spacing
								Line 0 ""
							}
							If($HTML)
							{
								$rowdata = @()
								$columnHeaders = @("Option Name",($htmlsilver -bor $htmlbold),"$($ScopePolicyOption.OptionId.ToString("00000")) $($ScopePolicyOption.Name)",$htmlwhite)
								$rowdata += @(,('Vendor',($htmlsilver -bor $htmlbold),$VendorClass,$htmlwhite))
								$rowdata += @(,('Value',($htmlsilver -bor $htmlbold),"$($ScopePolicyOption.Value)",$htmlwhite))
								$rowdata += @(,('Policy Name',($htmlsilver -bor $htmlbold),$PolicyName,$htmlwhite))
								$msg = ""
								$columnWidths = @("100","400")
								FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
								WriteHTMLLine 0 0 ""
							}
						}
					}
				}
			}
			ElseIf(!$?)
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Error retrieving scope options for $IPv4Scope.ScopeId"
				}
				If($Text)
				{
					Line 0 "Error retrieving scope options for $IPv4Scope.ScopeId"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Error retrieving scope options for $IPv4Scope.ScopeId"
				}
				InsertBlankLine
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 1 "None"
				}
				If($Text)
				{
					Line 3 "None"
				}
				If($HTML)
				{
					WriteHTMLLine 0 1 "None"
				}
				InsertBlankLine
			}
			
			$ScopePolicyOptions = $Null
			
			If($MSWord -or $PDF)
			{
				WriteWordLine ($xStartLevel + 2) 0 "DNS"
			}
			If($Text)
			{
				Line 2 "DNS"
			}
			If($HTML)
			{
				WriteHTMLLine ($xStartLevel + 2) 0 "DNS"
			}
			
			$ScopePolicyDNS = Get-DhcpServerv4DnsSetting -ComputerName $Script:DHCPServerName -PolicyName $Policy.Name -ScopeId $IPv4Scope.ScopeId -EA 0 
			
			If($? -and $Null -ne $ScopePolicyDNS)
			{
				If($MSWord -or $PDF)
				{
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					If($ScopePolicyDNS.DynamicUpdates -eq "Never")
					{
						$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Disabled"; }
					}
					ElseIf($ScopePolicyDNS.DynamicUpdates -eq "OnClientRequest")
					{
						$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
						$ScriptInformation += @{ Data = "Dynamically update DNS records only if requested by the DHCP clients"; Value = "Enabled"; }
					}
					ElseIf($ScopePolicyDNS.DynamicUpdates -eq "Always")
					{
						$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
						$ScriptInformation += @{ Data = "Always dynamically update DNS records"; Value = "Enabled"; }
					}
					If($ScopePolicyDNS.DeleteDnsRROnLeaseExpiry)
					{
						$ScriptInformation += @{ Data = "Discard A and PTR records when lease deleted"; Value = "Enabled"; }
					}
					Else
					{
						$ScriptInformation += @{ Data = "Discard A and PTR records when lease deleted"; Value = "Disabled"; }
					}
					If($ScopePolicyDNS.UpdateDnsRRForOlderClients)
					{
						$ScriptInformation += @{ Data = "Dynamically update DNS records for DHCP clients that do not request updates"; Value = "Enabled"; }
					}
					Else
					{
						$ScriptInformation += @{ Data = "Dynamically update DNS records for DHCP clients that do not request updates"; Value = "Disabled"; }
					}
					If($ScopePolicyDNS.DisableDnsPtrRRUpdate)
					{
						$ScriptInformation += @{ Data = "Disable dynamic updates for DNS PTR records"; Value = "Enabled"; }
					}
					Else
					{
						$ScriptInformation += @{ Data = "Disable dynamic updates for DNS PTR records"; Value = "Disabled"; }
					}
					If([string]::IsNullOrEmpty($ScopePolicyDNS.DnsSuffix))
					{
						$ScriptInformation += @{ Data = "Register DHCP clients using the following DNS suffix"; Value = "Disabled"; }
					}
					Else
					{
						$ScriptInformation += @{ Data = "Register DHCP clients using the following DNS suffix"; Value = $ScopePolicyDNS.DnsSuffix; }
					}
					If($ScopePolicyDNS.NameProtection)
					{
						$ScriptInformation += @{ Data = "Name Protection"; Value = "Enabled"; }
					}
					Else
					{
						$ScriptInformation += @{ Data = "Name Protection"; Value = "Disabled"; }
					}

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 400;
					$Table.Columns.Item(2).Width = 50;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					If($ScopePolicyDNS.DynamicUpdates -eq "Never")
					{
						Line 3 "Enable DNS dynamic updates`t`t`t`t: Disabled"
					}
					ElseIf($ScopePolicyDNS.DynamicUpdates -eq "OnClientRequest")
					{
						Line 3 "Enable DNS dynamic updates`t`t`t`t: Enabled"
						Line 3 "Dynamically update DNS records only "
						Line 3 "if requested by the DHCP clients`t`t`t: Enabled"
					}
					ElseIf($ScopePolicyDNS.DynamicUpdates -eq "Always")
					{
						Line 3 "Enable DNS dynamic updates`t`t`t`t: Enabled"
						Line 3 "Always dynamically update DNS records: Enabled"
					}
					If($ScopePolicyDNS.DeleteDnsRROnLeaseExpiry)
					{
						Line 3 "Discard A and PTR records when lease deleted`t`t: Enabled"
					}
					Else
					{
						Line 3 "Discard A and PTR records when lease deleted`t`t: Disabled"
					}
					Line 3 "Dynamically update DNS records for DHCP "
					If($ScopePolicyDNS.UpdateDnsRRForOlderClients)
					{
						Line 3 "clients that do not request updates`t`t`t: Enabled"
					}
					Else
					{
						Line 3 "clients that do not request updates`t`t`t: Disabled"
					}
					If($ScopePolicyDNS.DisableDnsPtrRRUpdate)
					{
						Line 3 "Disable dynamic updates for DNS PTR records`t`t: Enabled"
					}
					Else
					{
						Line 3 "Disable dynamic updates for DNS PTR records`t`t: Disabled"
					}
					If([string]::IsNullOrEmpty($ScopePolicyDNS.DnsSuffix))
					{
						Line 3 "Register DHCP clients using the following DNS suffix`t: Disabled"
					}
					Else
					{
						Line 3 "Register DHCP clients using the following DNS suffix`t: " $ScopePolicyDNS.DnsSuffix
					}
					If($ScopePolicyDNS.NameProtection)
					{
						Line 3 "Name Protection`t`t`t`t`t`t: Enabled"
					}
					Else
					{
						Line 3 "Name Protection`t`t`t`t`t`t: Disabled"
					}
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					If($ScopePolicyDNS.DynamicUpdates -eq "Never")
					{
						$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite)
					}
					ElseIf($ScopePolicyDNS.DynamicUpdates -eq "OnClientRequest")
					{
						$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
						$rowdata += @(,('Dynamically update DNS records only if requested by the DHCP clients',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
					}
					ElseIf($ScopePolicyDNS.DynamicUpdates -eq "Always")
					{
						$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
						$rowdata += @(,('Always dynamically update DNS records',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
					}
					If($ScopePolicyDNS.DeleteDnsRROnLeaseExpiry)
					{
						$rowdata += @(,('Discard A and PTR records when lease deleted',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('Discard A and PTR records when lease deleted',($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
					}
					If($ScopePolicyDNS.UpdateDnsRRForOlderClients)
					{
						$rowdata += @(,('Dynamically update DNS records for DHCP clients that do not request updates',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('Dynamically update DNS records for DHCP clients that do not request updates',($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
					}
					If($ScopePolicyDNS.DisableDnsPtrRRUpdate)
					{
						$rowdata += @(,('Disable dynamic updates for DNS PTR records',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('Disable dynamic updates for DNS PTR records',($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
					}
					If([string]::IsNullOrEmpty($ScopePolicyDNS.DnsSuffix))
					{
						$rowdata += @(,('Register DHCP clients using the following DNS suffix',($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('Register DHCP clients using the following DNS suffix',($htmlsilver -bor $htmlbold),$ScopePolicyDNS.DnsSuffix,$htmlwhite))
					}
					If($ScopePolicyDNS.NameProtection)
					{
						$rowdata += @(,('Name Protection',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
					}
					Else
					{
						$rowdata += @(,('Name Protection',($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
					}

					$msg = ""
					$columnWidths = @("450","50")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
					WriteHTMLLine 0 0 ""
				}
			}
			ElseIf(!$?)
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 1 "Error retrieving Policy DNS settings for Policy $Policy.Name"
				}
				If($Text)
				{
					Line 0 "Error retrieving Policy DNS settings for Policy $Policy.Name"
				}
				If($HTML)
				{
					WriteHTMLLine 0 1 "Error retrieving Policy DNS settings for Policy $Policy.Name"
				}
				InsertBlankLine
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 1 "None"
				}
				If($Text)
				{
					Line 3 "None"
				}
				If($HTML)
				{
					WriteHTMLLine 0 1 "None"
				}
				InsertBlankLine
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving scope policies for scope $($IPv4Scope.ScopeId)"
		}
		If($Text)
		{
			Line 0 "Error retrieving scope policies for scope $($IPv4Scope.ScopeId)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving scope policies for scope $($IPv4Scope.ScopeId)"
		}
		InsertBlankLine
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "None"
		}
		If($Text)
		{
			Line 2 "None"
		}
		If($HTML)
		{
			WriteHTMLLine 0 1 "None"
		}
		InsertBlankLine
	}
	$ScopePolicies = $Null

	Write-Verbose "$(Get-Date -Format G):	`t`tGetting DNS"
	If($MSWord -or $PDF)
	{
		WriteWordLine ($xStartLevel + 1) 0 "DNS"
	}
	If($Text)
	{
		Line 1 "DNS:"
	}
	If($HTML)
	{
		WriteHTMLLine ($xStartLevel + 1) 0 "DNS"
	}
	$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0
	If($? -and $Null -ne $DNSSettings)
	{
		If($DNSSettings.DeleteDnsRROnLeaseExpiry)
		{
			$DeleteDnsRROnLeaseExpiry = "Enabled"
		}
		Else
		{
			$DeleteDnsRROnLeaseExpiry = "Disabled"
		}

		If($DNSSettings.UpdateDnsRRForOlderClients)
		{
			$UpdateDnsRRForOlderClients = "Enabled"
		}
		Else
		{
			$UpdateDnsRRForOlderClients = "Disabled"
		}

		If($DNSSettings.DisableDnsPtrRRUpdate)
		{
			$DisableDnsPtrRRUpdate = "Enabled"
		}
		Else
		{
			$DisableDnsPtrRRUpdate = "Disabled"
		}

		If($DNSSettings.NameProtection)
		{
			$NameProtection = "Enabled"
		}
		Else
		{
			$NameProtection = "Disabled"
		}

		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			If($DNSSettings.DynamicUpdates -eq "Never")
			{
				$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Disabled"; }
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
			{
				$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
				$ScriptInformation += @{ Data = "Dynamically update DNS A and PTR records only if requested by the DHCP clients"; Value = "Enabled"; }
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "Always")
			{
				$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
				$ScriptInformation += @{ Data = "Always dynamically update DNS A and PTR records"; Value = "Enabled"; }
			}
			$ScriptInformation += @{ Data = "Discard A and PTR records when lease is deleted"; Value = $DeleteDnsRROnLeaseExpiry; }
			$ScriptInformation += @{ Data = "Dynamically update DNS records for DHCP clients that do not request updates"; Value = $UpdateDnsRRForOlderClients; }
			$ScriptInformation += @{ Data = "Disable dynamic updates for DNS PTR record"; Value = $DisableDnsPtrRRUpdate; }
			$ScriptInformation += @{ Data = "Name Protection"; Value = $NameProtection; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 400;
			$Table.Columns.Item(2).Width = 50;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 2 "Enable DNS dynamic updates`t`t`t: " -NoNewLine
			If($DNSSettings.DynamicUpdates -eq "Never")
			{
				Line 0 "Disabled"
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
			{
				Line 0 "Enabled"
				Line 2 "Dynamically update DNS A and PTR records only "
				Line 2 "if requested by the DHCP clients`t`t: Enabled"
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "Always")
			{
				Line 0 "Enabled"
				Line 2 "Always dynamically update DNS A and PTR records: Enabled"
			}
			Line 2 "Discard A and PTR records when lease deleted`t: " 
			Line 2 "Dynamically update DNS records for DHCP "
			Line 2 "clients that do not request updates`t`t: " $UpdateDnsRRForOlderClients
			Line 2 "Disable dynamic updates for DNS PTR records`t: " $DisableDnsPtrRRUpdate
			Line 2 "Name Protection`t`t`t`t`t: " $NameProtection
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()
			If($DNSSettings.DynamicUpdates -eq "Never")
			{
				$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite)
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
			{
				$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
				$rowdata += @(,("Dynamically update DNS A and PTR records only if requested by the DHCP clients",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "Always")
			{
				$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
				$rowdata += @(,("Always dynamically update DNS A and PTR records",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
			}
			$rowdata += @(,("Discard A and PTR records when lease deleted",($htmlsilver -bor $htmlbold),$DeleteDnsRROnLeaseExpiry,$htmlwhite))
			$rowdata += @(,('Dynamically update DNS records for DHCP clients that do not request updates',($htmlsilver -bor $htmlbold),$UpdateDnsRRForOlderClients,$htmlwhite))
			$rowdata += @(,('Disable dynamic updates for DNS PTR records',($htmlsilver -bor $htmlbold),$DisableDnsPtrRRUpdate,$htmlwhite))
			$rowdata += @(,('Name Protection',($htmlsilver -bor $htmlbold),$NameProtection,$htmlwhite))
			$msg = ""
			$columnWidths = @("450","50")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
			WriteHTMLLine 0 0 
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving DNS Settings for scope $($IPv4Scope.ScopeId)"
		}
		If($Text)
		{
			Line 0 "Error retrieving DNS Settings for scope $($IPv4Scope.ScopeId)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving DNS Settings for scope $($IPv4Scope.ScopeId)"
		}
		InsertBlankLine
	}
	$DNSSettings = $Null
	
	#next tab is Network Access Protection but I can't find anything that gives me that info
	
	#failover
	Write-Verbose "$(Get-Date -Format G):	`t`tGetting failover"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine ($xStartLevel + 1) 0 "Failover"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	}
	If($Text)
	{
		Line 1 "Failover:"
	}
	If($HTML)
	{
		WriteHTMLLine ($xStartLevel + 1) 0 "Failover"
		$rowdata = @()
	}
	
	$Failovers = $Null
	$Failovers = Get-DHCPServerV4Failover -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0

	If($? -and $Null -ne $Failovers)
	{
		ForEach($Failover in $Failovers)
		{
			Write-Verbose "$(Get-Date -Format G):	`t`tProcessing failover $($Failover.Name)"
			If($Null -ne $Failover.MaxClientLeadTime)
			{
				$MaxLeadStr = [string]::format("{0} days, {1} hours, {2} minutes", `
					$Failover.MaxClientLeadTime.Days, `
					$Failover.MaxClientLeadTime.Hours, `
					$Failover.MaxClientLeadTime.Minutes)
			}
			Else
			{
				$MaxLeadStr = ""
			}

			If($Null -ne $Failover.StateSwitchInterval)
			{
				$SwitchStr = [string]::format("{0} days, {1} hours, {2} minutes", `
					$Failover.StateSwitchInterval.Days, `
					$Failover.StateSwitchInterval.Hours, `
					$Failover.StateSwitchInterval.Minutes)
			}
			Else
			{
				$SwitchStr = "Disabled"
			}

			Switch($Failover.State)
			{
				"NoState"					{$ServerState = "No State"; $PartnerState = "No State"; Break}
				"Normal"					{$ServerState = "Normal"; $PartnerState = "Normal"; Break}
				"Init"						{$ServerState = "Initializing"; $PartnerState = "Initializing"; Break}
				"CommunicationInterrupted"	{$ServerState = "Communication Interrupted"; $PartnerState = "Communication Interrupted"; Break}
				"PartnerDown"				{$ServerState = "Normal"; $PartnerState = "Down"; Break}
				"PotentialConflict"			{$ServerState = "Potential Conflict"; $PartnerState = "Potential Conflict"; Break}
				"Startup"					{$ServerState = "Starting Up"; $PartnerState = "Starting Up"; Break}
				"ResolutionInterrupted"		{$ServerState = "Resolution Interrupted"; $PartnerState = "Resolution Interrupted"; Break}
				"ConflictDone"				{$ServerState = "Conflict Done"; $PartnerState = "Conflict Done"; Break}
				"Recover"					{$ServerState = "Recover"; $PartnerState = "Recover"; Break}
				"RecoverWait"				{$ServerState = "Recover Wait"; $PartnerState = "Recover Wait"; Break}
				"RecoverDone"				{$ServerState = "Recover Done"; $PartnerState = "Recover Done"; Break}
				Default						{$ServerState = "Unable to determine failover server state: $($Failover.State)"; $PartnerState = "Unable to determine partner server failover state: $($Failover.State)"; Break}
			}
			
			$FailoverLoadBalancePercent = (100 - $($Failover.LoadBalancePercent))

			If($MSWord -or $PDF)
			{
				$ScriptInformation += @{ Data = "Relationship name"; Value = $Failover.Name; }
				$ScriptInformation += @{ Data = "Partner Server"; Value = $Failover.PartnerServer; }
				$ScriptInformation += @{ Data = "Mode"; Value = $Failover.Mode; }
				$ScriptInformation += @{ Data = "Max Client Lead Time"; Value = $MaxLeadStr; }
				$ScriptInformation += @{ Data = "State Switchover Interval"; Value = $SwitchStr; }
				$ScriptInformation += @{ Data = "State of this Server"; Value = $ServerState; }
				$ScriptInformation += @{ Data = "State of Partner Server"; Value = $PartnerState; }
						
				If($Failover.Mode -eq "LoadBalance")
				{
					$ScriptInformation += @{ Data = "Local server"; Value = "$($Failover.LoadBalancePercent)%"; }
					$ScriptInformation += @{ Data = "Partner Server"; Value = "$($FailoverLoadBalancePercent)%"; }
				}
				Else
				{
					$ScriptInformation += @{ Data = "Role of this server"; Value = $Failover.ServerRole; }
					$ScriptInformation += @{ Data = "Addresses reserved for standby server"; Value = "$($Failover.ReservePercent)%"; }
				}
			}
			If($Text)
			{
				Line 2 "Relationship name: " $Failover.Name
				Line 2 "Partner Server`t`t`t: " $Failover.PartnerServer
				Line 2 "Mode`t`t`t`t: " $Failover.Mode
				Line 2 "Max Client Lead Time`t`t: " $MaxLeadStr
				Line 2 "State Switchover Interval`t: " $SwitchStr
				Line 2 "State of this Server`t`t: " $ServerState
				Line 2 "State of Partner Server`t`t: " $PartnerState
				If($Failover.Mode -eq "LoadBalance")
				{
					Line 2 "Local server`t`t`t: $($Failover.LoadBalancePercent)%"
					Line 2 "Partner Server`t`t`t: $($FailoverLoadBalancePercent)%"
				}
				Else
				{
					Line 2 "Role of this server`t`t: " $Failover.ServerRole
					Line 2 "Addresses reserved for standby server: $($Failover.ReservePercent)%"
				}
						
				#skip a row for spacing
				Line 0 ""
			}
			If($HTML)
			{
				$columnHeaders = @("Relationship name",($htmlsilver -bor $htmlbold),$Failover.Name,$htmlwhite)
				$rowdata += @(,('Partner Server',($htmlsilver -bor $htmlbold),$Failover.PartnerServer,$htmlwhite))
				$rowdata += @(,('Mode',($htmlsilver -bor $htmlbold),$Failover.Mode,$htmlwhite))
				$rowdata += @(,('Max Client Lead Time',($htmlsilver -bor $htmlbold),$MaxLeadStr,$htmlwhite))
				$rowdata += @(,('State Switchover Interval',($htmlsilver -bor $htmlbold),$SwitchStr,$htmlwhite))
				$rowdata += @(,('State of this Server',($htmlsilver -bor $htmlbold),$ServerState,$htmlwhite))
				$rowdata += @(,('State of Partner Server',($htmlsilver -bor $htmlbold),$PartnerState,$htmlwhite))
						
				If($Failover.Mode -eq "LoadBalance")
				{
					$rowdata += @(,('Local server',($htmlsilver -bor $htmlbold),"$($Failover.LoadBalancePercent)%",$htmlwhite))
					$rowdata += @(,('Partner Server',($htmlsilver -bor $htmlbold),"$($FailoverLoadBalancePercent)%",$htmlwhite))
				}
				Else
				{
					$rowdata += @(,('Role of this server',($htmlsilver -bor $htmlbold),$Failover.ServerRole,$htmlwhite))
					$rowdata += @(,('Addresses reserved for standby server',($htmlsilver -bor $htmlbold),"$($Failover.ReservePercent)%",$htmlwhite))
				}
			}
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "None"; Value = ""; }
		}
		If($Text)
		{
			Line 2 "None"
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @("None",($htmlsilver -bor $htmlbold),"",$htmlwhite)
		}
	}
	$Failovers = $Null

	If($MSWord -or $PDF)
	{

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
	If($HTML)
	{
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}

	Write-Verbose "$(Get-Date -Format G):	`t`tGetting Scope statistics"
	If($MSWord -or $PDF)
	{
		WriteWordLine ($xStartLevel + 1) 0 "Statistics"
	}
	If($Text)
	{
		Line 1 "Statistics"
	}
	If($HTML)
	{
		WriteHTMLLine ($xStartLevel + 1) 0 "Statistics"
	}

	$Statistics = Get-DHCPServerV4ScopeStatistics -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0

	If($? -and $Null -ne $Statistics)
	{
		GetShortStatistics $Statistics
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving scope statistics for scope $($IPv4Scope.ScopeId)"
		}
		If($Text)
		{
			Line 0 "Error retrieving scope statistics for scope $($IPv4Scope.ScopeId)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving scope statistics for scope $($IPv4Scope.ScopeId)"
		}
		InsertBlankLine
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "None"
		}
		If($Text)
		{
			Line 0 "None"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "None"
		}
		InsertBlankLine
	}
	$Statistics = $Null
}

Function ProcessIPv4MulticastScopes
{
	$CmdletName = "Get-DHCPServerV4MulticastScope"
	If(Get-Command $CmdletName -Module "DHCPServer" -EA 0)
	{
		Write-Verbose "$(Get-Date -Format G): Getting IPv4 Multicast scopes"
		$IPv4MulticastScopes = Get-DHCPServerV4MulticastScope -ComputerName $Script:DHCPServerName -EA 0

		If($? -and $Null -ne $IPv4MulticastScopes)
		{
			ForEach($IPv4MulticastScope in $IPv4MulticastScopes)
			{
				Write-Verbose "$(Get-Date -Format G): `tGetting IPv4 multicast scope data for scope $($IPv4MulticastScope.Name)"
				If($Null -ne $IPv4MulticastScope.LeaseDuration)
				{
					$DurationStr = [string]::format("{0} days, {1} hours, {2} minutes", `
						$IPv4MulticastScope.LeaseDuration.Days, `
						$IPv4MulticastScope.LeaseDuration.Hours, `
						$IPv4MulticastScope.LeaseDuration.Minutes)
				}
				Else
				{
					$DurationStr = "Unlimited"
				}
				
				If($MSWord -or $PDF)
				{
					#put each scope on a new page
					$selection.InsertNewPage()
					WriteWordLine 3 0 "Multicast Scope [$($IPv4MulticastScope.Name)]"
					WriteWordLine 4 0 "General"
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$ScriptInformation += @{ Data = "Name"; Value = $IPv4MulticastScope.Name; }
					$ScriptInformation += @{ Data = "Start IP address"; Value = $IPv4MulticastScope.StartRange; }
					$ScriptInformation += @{ Data = "End IP address"; Value = $IPv4MulticastScope.EndRange; }
					$ScriptInformation += @{ Data = "Time to live"; Value = $IPv4MulticastScope.Ttl; }
					$ScriptInformation += @{ Data = "Lease duration"; Value = $DurationStr; }
					$ScriptInformation += @{ Data = "Description"; Value = $IPv4MulticastScope.Description; }
					
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
				If($Text)
				{
					Line 0 "Multicast Scope [$($IPv4MulticastScope.Name)]"
					Line 1 "General:"
					Line 2 "Name`t`t`t: " $IPv4MulticastScope.Name
					Line 2 "Start IP address`t: " $IPv4MulticastScope.StartRange
					Line 2 "End IP address`t`t: " $IPv4MulticastScope.EndRange
					Line 2 "Time to live`t`t: " $IPv4MulticastScope.Ttl
					Line 2 "Lease duration`t`t: " $DurationStr
					Line 2 "Description`t`t: " $IPv4MulticastScope.Description
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 3 0 "Multicast Scope [$($IPv4MulticastScope.Name)]"
					WriteHTMLLine 4 0 "General"
					$rowdata = @()
					
					$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$IPv4MulticastScope.Name,$htmlwhite)
					$rowdata += @(,('Start IP address',($htmlsilver -bor $htmlbold),$IPv4MulticastScope.StartRange,$htmlwhite))
					$rowdata += @(,('End IP address',($htmlsilver -bor $htmlbold),$IPv4MulticastScope.EndRange,$htmlwhite))
					$rowdata += @(,('Time to live',($htmlsilver -bor $htmlbold),$IPv4MulticastScope.Ttl,$htmlwhite))
					$rowdata += @(,('Lease duration',($htmlsilver -bor $htmlbold),$DurationStr,$htmlwhite))
					$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$IPv4MulticastScope.Description,$htmlwhite))
					
					$msg = ""
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
					WriteHTMLLine 0 0 ""
				}

				If([string]::IsNullOrEmpty($IPv4MulticastScope.ExpiryTime))
				{
					$ExpiryTime = "Infinite"
				}
				Else
				{
					$ExpiryTime = "Multicast scope expires on $($IPv4MulticastScope.ExpiryTime)"
				}

				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "Lifetime"
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$ScriptInformation += @{ Data = "Multicast scope lifetime"; Value = $ExpiryTime; }
					
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
				If($Text)
				{
					Line 1 "Lifetime:"
					Line 2 "Multicast scope lifetime: " $ExpiryTime
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					$columnHeaders = @("Lifetime",($htmlsilver -bor $htmlbold),$ExpiryTime,$htmlwhite)
					
					$msg = ""
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
					WriteHTMLLine 0 0 ""
				}
				
				Write-Verbose "$(Get-Date -Format G): `t`tGetting exclusions"
				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "Exclusions"
				}
				If($Text)
				{
					Line 1 "Exclusions:"
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 "Exclusions"
				}
				$Exclusions = Get-DHCPServerV4MulticastExclusionRange -ComputerName $Script:DHCPServerName -Name $IPv4MulticastScope.Name -EA 0
				If($? -and $Null -ne $Exclusions)
				{
					If($MSWord -or $PDF)
					{
						[System.Collections.Hashtable[]] $ExclusionsWordTable = @()
					}
					If($Text)
					{
						Line 2 "Start IP Address`tEnd IP Address"
					}
					If($HTML)
					{
						$rowdata = @()
					}
					ForEach($Exclusion in $Exclusions)
					{
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{ 
							Start = $Exclusion.StartRange.ToString(); `
							Ending = $Exclusion.EndRange.ToString()
							}

							## Add the hash to the array
							$ExclusionsWordTable += $WordTableRowHash;
						}
						If($Text)
						{
							Line 2 $Exclusion.StartRange -NoNewLine
							Line 2 $Exclusion.EndRange 
						}
						If($HTML)
						{
							$rowdata += @(,($Exclusion.StartRange,$htmlwhite,
											$Exclusion.EndRange,$htmlwhite))
						}
					}
				}
				ElseIf(!$?)
				{
					If($MSWord -or $PDF)
					{
						$WordTableRowHash = @{ 
						Start = "Error retrieving exclusions for multicast scope"; `
						Ending = ""
						}

						## Add the hash to the array
						$ExclusionsWordTable += $WordTableRowHash;
					}
					If($Text)
					{
						Line 0 "Error retrieving exclusions for multicast scope"
						Line 0 ""
					}
					If($HTML)
					{
						$rowdata += @(,("Error retrieving exclusions for multicast scope",$htmlwhite,
										"",$htmlwhite))
					}
				}
				Else
				{
					If($MSWord -or $PDF)
					{
						$WordTableRowHash = @{ 
						Start = "None"; `
						Ending = ""
						}

						## Add the hash to the array
						$ExclusionsWordTable += $WordTableRowHash;
					}
					If($Text)
					{
						Line 2 "None"
						Line 0 ""
					}
					If($HTML)
					{
						$rowdata += @(,("None",$htmlwhite,
										"",$htmlwhite))
					}
				}
				If($MSWord -or $PDF)
				{
					## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
					If($ExclusionsWordTable.Count -gt 0)
					{
						$Table = AddWordTable -Hashtable $ExclusionsWordTable `
						-Columns Start,Ending `
						-Headers "Start IP Address","End IP Address" `
						-AutoFit $wdAutoFitContent;

						SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

						FindWordDocumentEnd
						$Table = $Null
						WriteWordLine 0 0 ""
					}
				}
				If($HTML)
				{
					$columnHeaders = @('Start IP Address',($htmlsilver -bor $htmlbold),'End IP Address',($htmlsilver -bor $htmlbold))
					$msg = ""
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
					WriteHTMLLine 0 0 ""
				}
					
				#leases
				If($IncludeLeases)
				{
					Write-Verbose "$(Get-Date -Format G): `t`tGetting leases"
					
					If($MSWord -or $PDF)
					{
						WriteWordLine 4 0 "Address Leases"
					}
					If($Text)
					{
						Line 1 "Address Leases:"
					}
					If($HTML)
					{
						WriteHTMLLine 4 0 "Address Leases"
					}
					$Leases = Get-DHCPServerV4MulticastLease -ComputerName $Script:DHCPServerName -Name $IPv4MulticastScope.Name -EA 0 | Sort-Object IPAddress
					If($? -and $Null -ne $Leases)
					{
						ForEach($Lease in $Leases)
						{
							Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing lease $($Lease.IPAddress)"
							If($Null -ne $Lease.LeaseExpiryTime)
							{
								$LeaseEndStr = [string]::format("{0} days, {1} hours, {2} minutes", `
									$Lease.LeaseExpiryTime.Days, `
									$Lease.LeaseExpiryTime.Hours, `
									$Lease.LeaseExpiryTime.Minutes)
							}
							Else
							{
								$LeaseEndStr = ""
							}

							If($Null -ne $Lease.LeaseExpiryTime)
							{
								$LeaseStartStr = [string]::format("{0} days, {1} hours, {2} minutes", `
									$Lease.LeaseStartTime.Days, `
									$Lease.LeaseStartTime.Hours, `
									$Lease.LeaseStartTime.Minutes)
							}
							Else
							{
								$LeaseStartStr = ""
							}

							If([string]::IsNullOrEmpty($Lease.LeaseExpiryTime))
							{
								$LeaseExpiryTime = "Unlimited"
							}
							Else
							{
								$LeaseExpiryTime = $LeaseEndStr
							}

							If([string]::IsNullOrEmpty($Lease.LeaseStartTime))
							{
								$LeaseStartTime = "Unlimited"
							}
							Else
							{
								$LeaseStartTime = $LeaseStartStr
							}
							
							If($MSWord -or $PDF)
							{
								[System.Collections.Hashtable[]] $ScriptInformation = @()
								$ScriptInformation += @{ Data = "Client IP address"; Value = $Lease.IPAddress; }
								$ScriptInformation += @{ Data = "Name"; Value = $Lease.HostName; }
								$ScriptInformation += @{ Data = "Lease Expiration"; Value = $LeaseExpiryTime; }
								$ScriptInformation += @{ Data = "Lease Start"; Value = $LeaseStartTime; }
								$ScriptInformation += @{ Data = "Address State"; Value = $Lease.AddressState; }
								$ScriptInformation += @{ Data = "MAC address"; Value = $Lease.ClientID; }

								$Table = AddWordTable -Hashtable $ScriptInformation `
								-Columns Data,Value `
								-List `
								-Format $wdTableGrid `
								-AutoFit $wdAutoFitFixed;

								SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

								$Table.Columns.Item(1).Width = 250;
								$Table.Columns.Item(2).Width = 250;

								$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

								FindWordDocumentEnd
								$Table = $Null
								WriteWordLine 0 0 ""
							}
							If($Text)
							{
								Line 2 "Client IP address`t: " $Lease.IPAddress
								Line 2 "Name`t`t`t: " $Lease.HostName
								Line 2 "Lease Expiration`t: " $LeaseExpiryTime
								Line 2 "Lease Start`t`t: " $LeaseStartTime
								Line 2 "Address State`t`t: " $Lease.AddressState
								Line 2 "MAC address`t: " $Lease.ClientID
								
								#skip a row for spacing
								Line 0 ""
							}
							If($HTML)
							{
								$rowdata = @()
								$columnHeaders = @("Client IP address",($htmlsilver -bor $htmlbold),$Lease.IPAddress.ToString(),$htmlwhite)
								$rowdata += @(,('Name',($htmlsilver -bor $htmlbold),$Lease.HostName,$htmlwhite))
								$rowdata += @(,('Lease Expiration',($htmlsilver -bor $htmlbold),$LeaseExpiryTime,$htmlwhite))
								$rowdata += @(,('Lease Start',($htmlsilver -bor $htmlbold),$LeaseStartTime,$htmlwhite))
								$rowdata += @(,('Address State',($htmlsilver -bor $htmlbold),$Lease.AddressState,$htmlwhite))
								$rowdata += @(,('MAC address',($htmlsilver -bor $htmlbold),$Lease.ClientID,$htmlwhite))
								
								$msg = ""
								$columnWidths = @("200","100")
								FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "300"
								InsertBlankLine
							}
						}
					}
					ElseIf(!$?)
					{
						If($MSWord -or $PDF)
						{
							WriteWordLine 0 0 "Error retrieving leases for scope"
						}
						If($Text)
						{
							Line 0 "Error retrieving leases for scope"
						}
						If($HTML)
						{
							WriteHTMLLine 0 0 "Error retrieving leases for scope"
						}
						InsertBlankLine
					}
					Else
					{
						If($MSWord -or $PDF)
						{
							WriteWordLine 0 1 "None"
						}
						If($Text)
						{
							Line 2 "None"
						}
						If($HTML)
						{
							WriteHTMLLine 0 1 "None"
						}
						InsertBlankLine
					}
					$Leases = $Null
				}
					
				Write-Verbose "$(Get-Date -Format G): `t`tGetting Multicast Scope statistics"
				
				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "Statistics"
				}
				If($Text)
				{
					Line 1 "Statistics"
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 "Statistics"
				}

				$Statistics = Get-DHCPServerV4MulticastScopeStatistics -ComputerName $Script:DHCPServerName -Name $IPv4MulticastScope.Name -EA 0

				If($? -and $Null -ne $Statistics)
				{
					GetShortStatistics $Statistics
				}
				ElseIf(!$?)
				{
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 0 "Error retrieving multicast scope statistics"
					}
					If($Text)
					{
						Line 0 "Error retrieving multicast scope statistics"
					}
					If($HTML)
					{
						WriteHTMLLine 0 0 "Error retrieving multicast scope statistics"
					}
					InsertBlankLine
				}
				Else
				{
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 1 "None"
					}
					If($Text)
					{
						Line 2 "None"
					}
					If($HTML)
					{
						WriteHTMLLine 0 1 "None"
					}
					InsertBlankLine
				}
				$Statistics = $Null
			}
		}
		ElseIf(!$?)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "Error retrieving IPv4 Multicast scopes"
			}
			If($Text)
			{
				Line 0 "Error retrieving IPv4 Multicast scopes"
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "Error retrieving IPv4 Multicast scopes"
			}
			InsertBlankLine
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "There were no IPv4 Multicast scopes"
			}
			If($Text)
			{
				Line 0 "There were no IPv4 Multicast scopes"
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "There were no IPv4 Multicast scopes"
			}
			InsertBlankLine
		}
		$IPv4MulticastScopes = $Null
	}
}

Function ProcessIPv4BOOTPTable
{
	#bootp table
	If($Null -ne $Script:BOOTPTable)
	{
		Write-Verbose "$(Get-Date -Format G):IPv4 BOOTP Table"
		
		If($MSWord -or $PDF)
		{
			$selection.InsertNewPage()
			WriteWordLine 3 0 "BOOTP Table"
			[System.Collections.Hashtable[]] $BootPWordTable = @()
		}
		If($Text)
		{
			Line 1 "BOOTP Table"
		}
		If($HTML)
		{
			WriteHTMLLine 3 0 "BOOTP Table"
			$rowdata = @()
		}
		
		ForEach($Item in $Script:BOOTPTable)
		{
			$ItemParts = $Item.Split(",")

			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				BootImage = $ItemParts[0]; `
				FileName = $ItemParts[1];
				FileServer = $ItemParts[2]
				}

				## Add the hash to the array
				$BootPWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 2 "Boot Image`t: " $ItemParts[0]
				Line 2 "File Name`t: " $ItemParts[1]
				Line 2 "FIle Server`t: " $ItemParts[2] 
				Line 0 ""
			}
			If($HTML)
			{
				$ItemParts = $Item.Split(",")
				$rowdata += @(,($ItemParts[0],$htmlwhite,
								$ItemParts[1],$htmlwhite,
								$ItemParts[2],$htmlwhite))
			}
		}

		If($MSWord -or $PDF)
		{
			## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
			If($BootPWordTable.Count -gt 0)
			{
				$Table = AddWordTable -Hashtable $BootPWordTable `
				-Columns BootImage,FileName,FileServer `
				-Headers "Boot Image","File Name","File Server" `
				-AutoFit $wdAutoFitContent;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
		}
		If($HTML)
		{
			$columnHeaders = @('Boot Image',($htmlsilver -bor $htmlbold),'File Name',($htmlsilver -bor $htmlbold),'File Server',($htmlsilver -bor $htmlbold))
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 ""
		}
	}
}

Function GetIPV6ScopeData
{
	Param([object]$IPv6Scope)

	ProcessIPv6ScopeData
}

Function ProcessIPv6ScopeData
{
	Write-Verbose "$(Get-Date -Format G): `tGetting IPv6 scope data for scope $($IPv6Scope.Name)"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Scope [$($IPv6Scope.Prefix)] $($IPv6Scope.Name)"
		WriteWordLine 4 0 "General"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Prefix"; Value = $IPv6Scope.Prefix.ToString(); }
		$ScriptInformation += @{ Data = "Preference"; Value = $IPv6Scope.Preference.ToString(); }
		$ScriptInformation += @{ Data = "Available Range"; Value = ""; }
		$ScriptInformation += @{ Data = "     Start"; Value = "$($IPv6Scope.Prefix)0:0:0:1"; }
		$ScriptInformation += @{ Data = "     End"; Value = "$($IPv6Scope.Prefix)ffff:ffff:ffff:ffff"; }
		$ScriptInformation += @{ Data = "Description"; Value = $IPv6Scope.Description; }

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
	If($Text)
	{
		Line 0 "Scope [$($IPv6Scope.Prefix)] $($IPv6Scope.Name)"
		Line 1 "General"
		Line 2 "Prefix`t`t: " $IPv6Scope.Prefix.ToString()
		Line 2 "Preference`t: " $IPv6Scope.Preference.ToString()
		Line 2 "Available Range`t: "
		Line 3 "Start`t: $($IPv6Scope.Prefix)0:0:0:1"
		Line 3 "End`t: $($IPv6Scope.Prefix)ffff:ffff:ffff:ffff"
		Line 2 "Description`t: " $IPv6Scope.Description
		Line 0 ""
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Scope [$($IPv6Scope.Prefix)] $($IPv6Scope.Name)"
		WriteHTMLLine 4 0 "General"
		$rowdata = @()
		$columnHeaders = @("Prefix",($htmlsilver -bor $htmlbold),$IPv6Scope.Prefix.ToString(),$htmlwhite)
		$rowdata += @(,('Preference',($htmlsilver -bor $htmlbold),$IPv6Scope.Preference.ToString(),$htmlwhite))
		$rowdata += @(,('Available Range',($htmlsilver -bor $htmlbold),"",$htmlwhite))
		$rowdata += @(,('     Start',($htmlsilver -bor $htmlbold),"$($IPv6Scope.Prefix)0:0:0:1",$htmlwhite))
		$rowdata += @(,('     End',($htmlsilver -bor $htmlbold),"$($IPv6Scope.Prefix)ffff:ffff:ffff:ffff",$htmlwhite))
		$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$IPv6Scope.Description,$htmlwhite))

		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}

	Write-Verbose "$(Get-Date -Format G): `t`tGetting scope DNS settings"

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "DNS"
	}
	If($Text)
	{
		Line 1 "DNS"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "DNS"
	}
	$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0
	If($? -and $Null -ne $DNSSettings)
	{
		If($DNSSettings.NameProtection)
		{
			$NameProtection = "Enabled"
		}
		Else
		{
			$NameProtection = "Disabled"
		}

		If($DNSSettings.DeleteDnsRROnLeaseExpiry)
		{
			$DeleteDnsRROnLeaseExpiry = "Enabled"
		}
		Else
		{
			$DeleteDnsRROnLeaseExpiry = "Disabled"
		}

		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			If($DNSSettings.DynamicUpdates -eq "Never")
			{
				$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Disabled"; }
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
			{
				$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
				$ScriptInformation += @{ Data = "Dynamically update DNS AAAA and PTR records only if requested by the DHCP clients"; Value = "Enabled"; }
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "Always")
			{
				$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
				$ScriptInformation += @{ Data = "Always dynamically update DNS AAAA and PTR records"; Value = "Enabled"; }
			}
			$ScriptInformation += @{ Data = "Discard AAAA and PTR records when lease is deleted"; Value = $DeleteDnsRROnLeaseExpiry; }
			$ScriptInformation += @{ Data = "Name Protection"; Value = $NameProtection; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 400;
			$Table.Columns.Item(2).Width = 50;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 2 "Enable DNS dynamic updates`t`t`t: " -NoNewLine
			If($DNSSettings.DynamicUpdates -eq "Never")
			{
				Line 0 "Disabled"
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
			{
				Line 0 "Enabled"
				Line 2 "Dynamically update DNS AAAA and PTR records only "
				Line 2 "if requested by the DHCP clients`t`t: Enabled"
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "Always")
			{
				Line 0 "Enabled"
				Line 2 "Always dynamically update DNS AAAA and PTR records: Enabled"
			}
			Line 2 "Discard AAAA and PTR records when lease deleted`t: " $DeleteDnsRROnLeaseExpiry
			Line 2 "Name Protection`t`t`t`t`t: " $NameProtection
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()
			If($DNSSettings.DynamicUpdates -eq "Never")
			{
				$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite)
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
			{
				$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
				$rowdata += @(,("Dynamically update DNS AAAA and PTR records only if requested by the DHCP clients",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "Always")
			{
				$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
				$rowdata += @(,("Always dynamically update DNS AAAA and PTR records",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
			}
			$rowdata += @(,("Discard AAAA and PTR records when lease deleted",($htmlsilver -bor $htmlbold),$DeleteDnsRROnLeaseExpiry,$htmlwhite))
			$rowdata += @(,('Name Protection',($htmlsilver -bor $htmlbold),$NameProtection,$htmlwhite))
			$msg = ""
			$columnWidths = @("450","50")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
			WriteHTMLLine 0 0 
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv6 DNS Settings for scope $($IPv6Scope.Prefix)"
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv6 DNS Settings for scope $($IPv6Scope.Prefix)"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv6 DNS Settings for scope $($IPv6Scope.Prefix)"
		}
		InsertBlankLine
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "None"
		}
		If($Text)
		{
			Line 0 "None"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "None"
		}
		InsertBlankLine
	}
	$DNSSettings = $Null
	
	Write-Verbose "$(Get-Date -Format G): `t`tGetting scope lease settings"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Lease"
	}
	If($Text)
	{
		Line 1 "Lease"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Lease"
	}
	
	$PrefStr = [string]::format("{0} days, {1} hours, {2} minutes", `
		$IPv6Scope.PreferredLifetime.Days, `
		$IPv6Scope.PreferredLifetime.Hours, `
		$IPv6Scope.PreferredLifetime.Minutes)
	
	$ValidStr = [string]::format("{0} days, {1} hours, {2} minutes", `
		$IPv6Scope.ValidLifetime.Days, `
		$IPv6Scope.ValidLifetime.Hours, `
		$IPv6Scope.ValidLifetime.Minutes)
		
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Preferred Life Time"; Value = $PrefStr; }
		$ScriptInformation += @{ Data = "Valid Life Time"; Value = $ValidStr; }

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
	If($Text)
	{
		Line 2 "Preferred Life Time`t: " $PrefStr
		Line 2 "Valid Life Time`t`t: " $ValidStr
		Line 0 ""
	}
	If($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Preferred Life Time",($htmlsilver -bor $htmlbold),$PrefStr,$htmlwhite)
		$rowdata += @(,('Valid Life Time',($htmlsilver -bor $htmlbold),$ValidStr,$htmlwhite))

		$msg = ""
		$columnWidths = @("125","200")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "325"
		WriteHTMLLine 0 0 ""
	}
	
	If($IncludeLeases)
	{
		Write-Verbose "$(Get-Date -Format G): `t`tGetting leases"
		
		If($MSWord -or $PDF)
		{
			WriteWordLine 4 0 "Address Leases"
		}
		If($Text)
		{
			Line 1 "Address Leases:"
		}
		If($HTML)
		{
			WriteHTMLLine 4 0 "Address Leases"
		}
		
		$Leases = Get-DHCPServerV6Lease -ComputerName $Script:DHCPServerName -Prefix  $IPv6Scope.Prefix -EA 0 | Sort-Object IPAddress
		If($? -and $Null -ne $Leases)
		{
			ForEach($Lease in $Leases)
			{
				Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing lease $($Lease.IPAddress)"
				If($Null -ne $Lease.LeaseExpiryTime)
				{
					$LeaseStr = [string]::format("{0} days, {1} hours, {2} minutes", `
						$Lease.LeaseExpiryTime.Day, `
						$Lease.LeaseExpiryTime.Hour, `
						$Lease.LeaseExpiryTime.Minute)
				}
				Else
				{
					$LeaseStr = ""
				}

				If($MSWord -or $PDF)
				{
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$ScriptInformation += @{ Data = "Client IPv6 address"; Value = $Lease.IPAddress; }
					$ScriptInformation += @{ Data = "Name"; Value = $Lease.HostName; }
					$ScriptInformation += @{ Data = "Lease Expiration"; Value = $LeaseStr; }
					$ScriptInformation += @{ Data = "IAID"; Value = $Lease.Iaid; }
					$ScriptInformation += @{ Data = "Type"; Value = $Lease.AddressType; }
					$ScriptInformation += @{ Data = "Unique ID"; Value = $Lease.ClientDuid; }
					$ScriptInformation += @{ Data = "Description"; Value = $Lease.Description; }
					
					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 100;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 2 "Client IPv6 address: " $Lease.IPAddress
					Line 2 "Name`t`t`t: " $Lease.HostName
					Line 2 "Lease Expiration`t: " $LeaseStr
					Line 2 "IAID`t`t`t: " $Lease.Iaid
					Line 2 "Type`t`t`t: " $Lease.AddressType
					Line 2 "Unique ID`t`t: " $Lease.ClientDuid
					Line 2 "Description`t`t: " $Lease.Description
					
					#skip a row for spacing
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					$columnHeaders = @("Client IPv6 address",($htmlsilver -bor $htmlbold),$Lease.IPAddress,$htmlwhite)
					$rowdata += @(,('Name',($htmlsilver -bor $htmlbold),$Lease.HostName,$htmlwhite))
					$rowdata += @(,('Lease Expiration',($htmlsilver -bor $htmlbold),$LeaseStr,$htmlwhite))
					$rowdata += @(,('IAID',($htmlsilver -bor $htmlbold),$Lease.Iaid,$htmlwhite))
					$rowdata += @(,('Type',($htmlsilver -bor $htmlbold),$Lease.AddressType,$htmlwhite))
					$rowdata += @(,('Unique ID',($htmlsilver -bor $htmlbold),$Lease.ClientDuid,$htmlwhite))
					$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Lease.Description,$htmlwhite))
					
					$msg = ""
					$columnWidths = @("125","250")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "375"
					WriteHTMLLine 0 0 ""
				}
			}
		}
		ElseIf(!$?)
		{
			If(MSWord -or $PDF)
			{
				WriteWordLine 0 0 "Error retrieving leases for scope $IPv6Scope.Prefix"
			}
			If($Text)
			{
				Line 0 "Error retrieving leases for scope $IPv6Scope.Prefix"
			}
			If($HTML)
			{
				WriteHTMLLine 0 0 "Error retrieving leases for scope $IPv6Scope.Prefix"
			}
			InsertBlankLine
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 1 "None"
			}
			If($Text)
			{
				Line 2 "None"
			}
			If($HTML)
			{
				WriteHTMLLine 0 1 "None"
			}
			InsertBlankLine
		}
		$Leases = $Null
	}

	Write-Verbose "$(Get-Date -Format G): `t`tGetting exclusions"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Exclusions"
	}
	If($Text)
	{
		Line 1 "Exclusions:"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Exclusions"
	}
	
	$Exclusions = Get-DHCPServerV6ExclusionRange -ComputerName $Script:DHCPServerName -Prefix  $IPv6Scope.Prefix -EA 0 | Sort-Object StartRange
	If($? -and $Null -ne $Exclusions)
	{
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ExclusionsWordTable = @()
		}
		If($HTML)
		{
			$rowdata = @()
		}
		ForEach($Exclusion in $Exclusions)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				Start = $Exclusion.StartRange.ToString(); `
				Ending = $Exclusion.EndRange.ToString()
				}

				## Add the hash to the array
				$ExclusionsWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 2 "Start IP Address`t: " $Exclusion.StartRange
				Line 2 "End IP Address`t`t: " $Exclusion.EndRange 
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata += @(,($Exclusion.StartRange,$htmlwhite,
								$Exclusion.EndRange,$htmlwhite))
			}
		}
		
		If($MSWord -or $PDF)
		{
			## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
			If($ExclusionsWordTable.Count -gt 0)
			{
				$Table = AddWordTable -Hashtable $ExclusionsWordTable `
				-Columns Start,Ending `
				-Headers "Start IP Address","End IP Address" `
				-AutoFit $wdAutoFitContent;

				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
		}
		If($HTML)
		{
			$columnHeaders = @('Start IP Address',($htmlsilver -bor $htmlbold),'End IP Address',($htmlsilver -bor $htmlbold))
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 ""
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving exclusions for scope $IPv6Scope.Prefix"
		}
		If($Text)
		{
			Line 0 "Error retrieving exclusions for scope $IPv6Scope.Prefix"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving exclusions for scope $IPv6Scope.Prefix"
		}
		InsertBlankLine
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "None"
		}
		If($Text)
		{
			Line 2 "None"
		}
		If($HTML)
		{
			WriteHTMLLine 0 1 "None"
		}
		InsertBlankLine
	}
	$Exclusions = $Null

	Write-Verbose "$(Get-Date -Format G): `t`tGetting reservations"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Reservations"
	}
	If($Text)
	{
		Line 1 "Reservations:"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Reservations"
	}
	
	$Reservations = Get-DHCPServerV6Reservation -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0 | Sort-Object Name
	If($? -and $Null -ne $Reservations)
	{
		ForEach($Reservation in $Reservations)
		{
			Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing reservation $($Reservation.Name)"
			
			If($MSWord -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Reservation name"; Value = $Reservation.Name; }
				$ScriptInformation += @{ Data = "IPv6 address"; Value = $Reservation.IPAddress; }
				$ScriptInformation += @{ Data = "DUID"; Value = $Reservation.ClientDuid; }
				$ScriptInformation += @{ Data = "IAID"; Value = $Reservation.Iaid; }
				$ScriptInformation += @{ Data = "Description"; Value = $Reservation.Description; }

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 100;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 "Reservation name: " $Reservation.Name
				Line 2 "IPv6 address: " $Reservation.IPAddress
				Line 2 "DUID`t`t: " $Reservation.ClientDuid
				Line 2 "IAID`t`t: " $Reservation.Iaid
				Line 2 "Description`t: " $Reservation.Description
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Reservation name",($htmlsilver -bor $htmlbold),$Reservation.Name,$htmlwhite)
				$rowdata += @(,('IPv6 address',($htmlsilver -bor $htmlbold),$Reservation.IPAddress.ToString(),$htmlwhite))
				$rowdata += @(,('DUID',($htmlsilver -bor $htmlbold),$Reservation.ClientDuid,$htmlwhite))
				$rowdata += @(,('IAID',($htmlsilver -bor $htmlbold),$Reservation.Iaid,$htmlwhite))
				$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Reservation.Description,$htmlwhite))
				$msg = ""
				$columnWidths = @("125","250")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "375"
				WriteHTMLLine 0 0 ""
			}

			Write-Verbose "$(Get-Date -Format G): `t`t`t`tGetting DNS settings"
			$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $Script:DHCPServerName -IPAddress $Reservation.IPAddress -EA 0
			If($? -and $Null -ne $DNSSettings)
			{
				If($DNSSettings.DeleteDnsRROnLeaseExpiry)
				{
					$DeleteDnsRROnLeaseExpiry = "Enabled"
				}
				Else
				{
					$DeleteDnsRROnLeaseExpiry = "Disabled"
				}

				If($DNSSettings.NameProtection)
				{
					$NameProtection = "Enabled"
				}
				Else
				{
					$NameProtection = "Disabled"
				}

				If($MSWord -or $PDF)
				{
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					If($DNSSettings.DynamicUpdates -eq "Never")
					{
						$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Disabled"; }
					}
					ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
					{
						$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
						$ScriptInformation += @{ Data = "Dynamically update DNS AAAA and PTR records only if requested by the DHCP clients"; Value = "Enabled"; }
					}
					ElseIf($DNSSettings.DynamicUpdates -eq "Always")
					{
						$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
						$ScriptInformation += @{ Data = "Always dynamically update DNS AAAA and PTR records"; Value = "Enabled"; }
					}
					$ScriptInformation += @{ Data = "Discard AAAA and PTR records when lease is deleted"; Value = $DeleteDnsRROnLeaseExpiry; }
					$ScriptInformation += @{ Data = "Name Protection"; Value = $NameProtection; }

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 400;
					$Table.Columns.Item(2).Width = 50;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 2 "Enable DNS dynamic updates`t`t`t: " -NoNewLine
					If($DNSSettings.DynamicUpdates -eq "Never")
					{
						Line 0 "Disabled"
					}
					ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
					{
						Line 0 "Enabled"
						Line 2 "Dynamically update DNS AAAA and PTR records only "
						Line 2 "if requested by the DHCP clients`t`t: Enabled"
					}
					ElseIf($DNSSettings.DynamicUpdates -eq "Always")
					{
						Line 0 "Enabled"
						Line 2 "Always dynamically update DNS AAAA and PTR records: Enabled"
					}
					Line 2 "Discard AAAA and PTR records when lease deleted`t: " $DeleteDnsRROnLeaseExpiry
					Line 2 "Name Protection`t`t`t`t`t: " $NameProtection
					Line 0 ""
				}
				If($HTML)
				{
					$rowdata = @()
					If($DNSSettings.DynamicUpdates -eq "Never")
					{
						$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite)
					}
					ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
					{
						$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
						$rowdata += @(,("Dynamically update DNS AAAA and PTR records only if requested by the DHCP clients",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
					}
					ElseIf($DNSSettings.DynamicUpdates -eq "Always")
					{
						$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
						$rowdata += @(,("Always dynamically update DNS AAAA and PTR records",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
					}
					$rowdata += @(,("Discard AAAA and PTR records when lease deleted",($htmlsilver -bor $htmlbold),$DeleteDnsRROnLeaseExpiry,$htmlwhite))
					$rowdata += @(,('Name Protection',($htmlsilver -bor $htmlbold),$NameProtection,$htmlwhite))
					$msg = ""
					$columnWidths = @("450","50")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
					WriteHTMLLine 0 0 
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 0 "Error to retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
				}
				If($Text)
				{
					Line 0 "Error to retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
				}
				If($HTML)
				{
					WriteHTMLLine 0 0 "Error to retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
				}
				InsertBlankLine
			}
			$DNSSettings = $Null
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving reservations for scope $IPv6Scope.Prefix"
		}
		If($Text)
		{
			Line 0 "Error retrieving reservations for scope $IPv6Scope.Prefix"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving reservations for scope $IPv6Scope.Prefix"
		}
		InsertBlankLine
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "None"
		}
		If($Text)
		{
			Line 2 "None"
		}
		If($HTML)
		{
			WriteHTMLLine 0 1 "None"
		}
		InsertBlankLine
	}
	$Reservations = $Null

	Write-Verbose "$(Get-Date -Format G): Getting IPv6 scope options"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Scope Options"
	}
	If($Text)
	{
		Line 1 "Scope Options:"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Scope Options"
	}

	$ScopeOptions = Get-DHCPServerV6OptionValue -All -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0 | Sort-Object OptionId

	If($? -and $Null -ne $ScopeOptions)
	{
		ForEach($ScopeOption in $ScopeOptions)
		{
			Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing option name $($ScopeOption.Name)"
			If([string]::IsNullOrEmpty($ScopeOption.VendorClass))
			{
				$VendorClass = "Standard" 
			}
			Else
			{
				$VendorClass = $ScopeOption.VendorClass 
			}
			
			If($MSWord -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Option Name"; Value = "$($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)"; }
				$ScriptInformation += @{ Data = "Vendor"; Value = $VendorClass; }
				$ScriptInformation += @{ Data = "Value"; Value = "$($ScopeOption.Value)"; }
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 100;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 "Option Name`t: $($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)" 
				Line 2 "Vendor`t`t: " $VendorClass
				Line 2 "Value`t`t: $($ScopeOption.Value)" 
				
				#for spacing
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Option Name",($htmlsilver -bor $htmlbold),"$($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)",$htmlwhite)
				$rowdata += @(,('Vendor',($htmlsilver -bor $htmlbold),$VendorClass,$htmlwhite))
				$rowdata += @(,('Value',($htmlsilver -bor $htmlbold),"$($ScopeOption.Value)",$htmlwhite))
				
				$msg = ""
				$columnWidths = @("100","400")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
				WriteHTMLLine 0 0 ""
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv6 scope options"
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv6 scope options"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv6 scope options"
		}
		InsertBlankLine
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "None"
		}
		If($Text)
		{
			Line 2 "None"
		}
		If($HTML)
		{
			WriteHTMLLine 0 1 "None"
		}
		InsertBlankLine
	}
	$ScopeOptions = $Null
	
	Write-Verbose "$(Get-Date -Format G): `t`tGetting Scope statistics"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Statistics"
	}
	If($Text)
	{
		Line 1 "Statistics"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Statistics"
	}

	$Statistics = Get-DHCPServerV6ScopeStatistics -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0

	If($? -and $Null -ne $Statistics)
	{
		GetShortStatistics $Statistics
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving scope statistics"
		}
		If($Text)
		{
			Line 0 "Error retrieving scope statistics"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving scope statistics"
		}
		InsertBlankLine
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "None"
		}
		If($Text)
		{
			Line 2 "None"
		}
		If($HTML)
		{
			WriteHTMLLine 0 1 "None"
		}
		InsertBlankLine
	}
	$Statistics = $Null
}

Function ProcessServerOptions
{
	#Server Options
	Write-Verbose "$(Get-Date -Format G): Getting IPv4 server options"

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 3 0 "Server Options"
	}
	If($Text)
	{
		Line 1 "Server Options"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Server Options"
	}

	$ServerOptions = Get-DHCPServerV4OptionValue -All -ComputerName $Script:DHCPServerName -EA 0 | Where-Object {$_.OptionID -ne 81} | Sort-Object OptionId

	If($? -and $Null -ne $ServerOptions)
	{
		ForEach($ServerOption in $ServerOptions)
		{
			Write-Verbose "$(Get-Date -Format G): `t`t`tProcessing option name $($ServerOption.Name)"
			If([string]::IsNullOrEmpty($ServerOption.VendorClass))
			{
				$VendorClass = "Standard"
			}
			Else
			{
				$VendorClass = $ServerOption.VendorClass
			}

			If([string]::IsNullOrEmpty($ServerOption.PolicyName))
			{
				$PolicyName = "None"
			}
			Else
			{
				$PolicyName = $ServerOption.PolicyName
			}

			If($MSWord -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()

				$ScriptInformation += @{ Data = "Option Name"; Value = "$($ServerOption.OptionId.ToString("000")) $($ServerOption.Name)"; }
				$ScriptInformation += @{ Data = "Vendor"; Value = $VendorClass; }
				$ScriptInformation += @{ Data = "Value"; Value = $ServerOption.Value[0]; }
				$ScriptInformation += @{ Data = "Policy Name"; Value = $PolicyName; }

				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 75;
				$Table.Columns.Item(2).Width = 300;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 2 "Option Name`t: $($ServerOption.OptionId.ToString("000")) $($ServerOption.Name)"
				Line 2 "Vendor`t`t: " $VendorClass
				Line 2 "Value`t`t: " $ServerOption.Value[0]
				Line 2 "Policy Name`t: " $PolicyName
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Option Name",($htmlsilver -bor $htmlbold),"$($ServerOption.OptionId.ToString("000")) $($ServerOption.Name)",$htmlwhite)
				$rowdata += @(,('Vendor',($htmlsilver -bor $htmlbold),$VendorClass,$htmlwhite))
				$rowdata += @(,('Value',($htmlsilver -bor $htmlbold),$ServerOption.Value[0],$htmlwhite))
				$rowdata += @(,('Policy Name',($htmlsilver -bor $htmlbold),$PolicyName,$htmlwhite))

				$msg = ""
				$columnWidths = @("100","400")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
				WriteHTMLLine 0 0 ""
			}
		}
	}
	ElseIf($? -and $Null -eq $ServerOptions)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no IPv4 server options"
		}
		If($Text)
		{
			Line 2 "There were no IPv4 server options"
		}
		If($HTML)
		{
			WriteHTMLLine 0 1 "There were no IPv4 server options"
		}
		InsertBlankLine
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv4 server options"
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv4 server options"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv4 server options"
		}
		InsertBlankLine
	}
	$ServerOptions = $Null
}

Function ProcessPolicies
{
	#Policies
	Write-Verbose "$(Get-Date -Format G): Getting IPv4 policies"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Policies"
	}
	If($Text)
	{
		Line 1 "Policies"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Policies"
	}

	$Policies = Get-DHCPServerV4Policy -ComputerName $Script:DHCPServerName -EA 0 | Sort-Object ProcessingOrder

	If($? -and $Null -ne $Policies)
	{
		If($MSWord -or $PDF)
		{
			ForEach($Policy in $Policies)
			{
				If($Policy.Enabled)
				{
					$PolicyEnabled = "Enabled"
				}
				Else
				{
					$PolicyEnabled = "Disabled"
				}

				$LeaseDuration = [string]::format("{0} days, {1} hours, {2} minutes", `
					$Policy.LeaseDuration.Days, `
					$Policy.LeaseDuration.Hours, `
					$Policy.LeaseDuration.Minutes)

				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "General"
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$ScriptInformation += @{ Data = "Policy Name"; Value = $Policy.Name; }
					$ScriptInformation += @{ Data = "Description"; Value = $Policy.Description; }
					$ScriptInformation += @{ Data = "Processing Order"; Value = $Policy.ProcessingOrder.ToString(); }
					$ScriptInformation += @{ Data = "Level"; Value = "Scope"; }
					$ScriptInformation += @{ Data = "State"; Value = $PolicyEnabled; }

					If($Policy.LeaseDuration.ToString() -eq "00:00:00")	#lease duration is not set
					{
						$ScriptInformation += @{ Data = "Set lease duration for the policy"; Value = "Not selected"; }
					}
					Else
					{
						#lease duration is set
						$ScriptInformation += @{ Data = "Set lease duration for the policy"; Value = "Selected"; }
						If($Policy.LeaseDuration.ToString() -eq "10675199.02:48:05.4775807") #unlimited
						{
							$ScriptInformation += @{ Data = "Lease duration for DHCP clients"; Value = "Unlimited"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = "Lease duration for DHCP clients"; Value = $LeaseDuration; }
						}
					}

					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 250;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 2 "General"
					Line 3 "Policy Name`t`t`t`t: " $Policy.Name
					Line 3 "Description`t`t`t`t: " $Policy.Description
					Line 3 "Processing Order`t`t`t: " $Policy.ProcessingOrder
					Line 3 "Level`t`t`t`t`t: Server"
					Line 3 "State`t`t`t`t`t: " $PolicyEnabled
					If($Policy.LeaseDuration.ToString() -eq "00:00:00")	#lease duration is not set
					{
						Line 3 "Set lease duration for the policy`t: Not selected" 
					}
					Else
					{
						#lease duration is set
						Line 3 "Set lease duration for the policy`t: Selected" 
						If($Policy.LeaseDuration.ToString() -eq "10675199.02:48:05.4775807") #unlimited
						{
							Line 3 "Lease duration for DHCP clients`t`t: Unlimited"
						}
						Else
						{
							Line 3 "Lease duration for DHCP clients`t`t: " $LeaseDuration
						}
					}
					Line 0 ""
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 "General"
					$rowdata = @()
					$columnHeaders = @("Policy Name",($htmlsilver -bor $htmlbold),$Policy.Name,$htmlwhite)
					$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Policy.Description,$htmlwhite))
					$rowdata += @(,('Processing Order',($htmlsilver -bor $htmlbold),$Policy.ProcessingOrder,$htmlwhite))
					$rowdata += @(,('Level',($htmlsilver -bor $htmlbold),"Server",$htmlwhite))
					$rowdata += @(,('State',($htmlsilver -bor $htmlbold),$PolicyEnabled,$htmlwhite))
					If($Policy.LeaseDuration.ToString() -eq "00:00:00")	#lease duration is not set
					{
						$rowdata += @(,('Set lease duration for the policy',($htmlsilver -bor $htmlbold),"Not selected" ,$htmlwhite))
					}
					Else
					{
						#lease duration is set
						$rowdata += @(,('Set lease duration for the policy',($htmlsilver -bor $htmlbold),"Selected",$htmlwhite))
						If($Policy.LeaseDuration.ToString() -eq "10675199.02:48:05.4775807") #unlimited
						{
							$rowdata += @(,('Lease duration for DHCP clients',($htmlsilver -bor $htmlbold),"Unlimited",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('Lease duration for DHCP clients',($htmlsilver -bor $htmlbold),$LeaseDuration,$htmlwhite))
						}
					}

					$msg = ""
					$columnWidths = @("250","250")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
					WriteHTMLLine 0 0 ""
				}

				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "Conditions"
					[System.Collections.Hashtable[]] $ScriptInformation = @()
					$ScriptInformation += @{ Data = "Condition"; Value = $Policy.Condition; }
				}
				If($Text)
				{
					Line 2 "Conditions"
					Line 3 "Condition: " $Policy.Condition
					Line 3 "Conditions                                  Operator    Value                                   "
					Line 3 "================================================================================================"
						   #123456789012345678901234567890123456789012SS1234567890SS1234567890123456789012345678901234567890
						   #Relay Agent Information - Agent Circuit Id              Default Routing and Remote Access Class
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 "Conditions"
					$rowdata = @()
					$columnHeaders = @("Condition",($htmlsilver -bor $htmlbold),$Policy.Condition,$htmlwhite)
				}
				
				If(![string]::IsNullOrEmpty($Policy.VendorClass))
				{
					For( $i = 0; $i -lt $Policy.VendorClass.Length; $i += 2 )
					{
						$xCond  = "Vendor Class"
						If($Policy.VendorClass[ $i ] -eq "EQ")
						{
							$xOper = "Equals" 
						}
						Else
						{
							$xOper = "Not Equals" 
						}
						$xValue = $Policy.VendorClass[ $i + 1 ]

						If($MSWord -or $PDF)
						{
							$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
						}
						If($Text)
						{
							Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
						}
						If($HTML)
						{
							$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
						}
					}
				}
				If(![string]::IsNullOrEmpty($Policy.UserClass))
				{
					For( $i = 0; $i -lt $Policy.UserClass.Length; $i += 2 )
					{
						$xCond  = "User Class"
						If($Policy.UserClass[ $i ] -eq "EQ")
						{
							$xOper = "Equals" 
						}
						Else
						{
							$xOper = "Not Equals" 
						}
						$xValue = $Policy.UserClass[ $i + 1 ]

						If($MSWord -or $PDF)
						{
							$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
						}
						If($Text)
						{
							Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
						}
						If($HTML)
						{
							$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
						}
					}
				}
				If(![string]::IsNullOrEmpty($Policy.MacAddress))
				{
					For( $i = 0; $i -lt $Policy.MacAddress.Length; $i += 2 )
					{
						$xCond  = "MAC Address"
						If($Policy.MacAddress[ $i ] -eq "EQ")
						{
							$xOper = "Equals" 
						}
						Else
						{
							$xOper = "Not Equals" 
						}
						$xValue = $Policy.MacAddress[ $i + 1 ]

						If($MSWord -or $PDF)
						{
							$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
						}
						If($Text)
						{
							Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
						}
						If($HTML)
						{
							$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
						}
					}
				}
				If(![string]::IsNullOrEmpty($Policy.ClientId))
				{
					For( $i = 0; $i -lt $Policy.ClientId.Length; $i += 2 )
					{
						$xCond  = "Client Identifier"
						If($Policy.ClientId[ $i ] -eq "EQ")
						{
							$xOper = "Equals" 
						}
						Else
						{
							$xOper = "Not Equals" 
						}
						$xValue = $Policy.ClientId[ $i + 1 ]

						If($MSWord -or $PDF)
						{
							$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
						}
						If($Text)
						{
							Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
						}
						If($HTML)
						{
							$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
						}
					}
				}
				If(![string]::IsNullOrEmpty($Policy.Fqdn))
				{
					For( $i = 0; $i -lt $Policy.Fqdn.Length; $i += 2 )
					{
						$xCond  = "Fully Qualified Domain Name"
						If($Policy.Fqdn[ $i ] -eq "EQ")
						{
							$xOper = "Equals" 
						}
						Else
						{
							$xOper = "Not Equals" 
						}
						$xValue = $Policy.Fqdn[ $i + 1 ]

						If($MSWord -or $PDF)
						{
							$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
						}
						If($Text)
						{
							Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
						}
						If($HTML)
						{
							$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
						}
					}
				}
				If(![string]::IsNullOrEmpty($Policy.RelayAgent))
				{
					For( $i = 0; $i -lt $Policy.RelayAgent.Length; $i += 2 )
					{
						$xCond  = "Relay Agent Information"
						If($Policy.RelayAgent[ $i ] -eq "EQ")
						{
							$xOper = "Equals" 
						}
						Else
						{
							$xOper = "Not Equals" 
						}
						$xValue = $Policy.RelayAgent[ $i + 1 ]

						If($MSWord -or $PDF)
						{
							$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
						}
						If($Text)
						{
							Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
						}
						If($HTML)
						{
							$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
						}
					}
				}
				If(![string]::IsNullOrEmpty($Policy.CircuitId))
				{
					For( $i = 0; $i -lt $Policy.CircuitId.Length; $i += 2 )
					{
						$xCond  = "Relay Agent Information - Agent Circuit Id"
						If($Policy.CircuitId[ $i ] -eq "EQ")
						{
							$xOper = "Equals" 
						}
						Else
						{
							$xOper = "Not Equals" 
						}
						$xValue = $Policy.CircuitId[ $i + 1 ]

						If($MSWord -or $PDF)
						{
							$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
						}
						If($Text)
						{
							Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
						}
						If($HTML)
						{
							$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
						}
					}
				}
				If(![string]::IsNullOrEmpty($Policy.RemoteId))
				{
					For( $i = 0; $i -lt $Policy.RemoteId.Length; $i += 2 )
					{
						$xCond  = "Relay Agent Information - Agent Remote Id"
						If($Policy.RemoteId[ $i ] -eq "EQ")
						{
							$xOper = "Equals" 
						}
						Else
						{
							$xOper = "Not Equals" 
						}
						$xValue = $Policy.RemoteId[ $i + 1 ]

						If($MSWord -or $PDF)
						{
							$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
						}
						If($Text)
						{
							Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
						}
						If($HTML)
						{
							$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
						}
					}
				}
				If(![string]::IsNullOrEmpty($Policy.SubscriberId))
				{
					For( $i = 0; $i -lt $Policy.SubscriberId.Length; $i += 2 )
					{
						$xCond  = "Relay Agent Information - Subscriber Id"
						If($Policy.SubscriberId[ $i ] -eq "EQ")
						{
							$xOper = "Equals" 
						}
						Else
						{
							$xOper = "Not Equals" 
						}
						$xValue = $Policy.SubscriberId[ $i + 1 ]

						If($MSWord -or $PDF)
						{
							$ScriptInformation += @{ Data = "Conditions: $xCond"; Value = "$xOper $xValue"; }
						}
						If($Text)
						{
							Line 3 ("{0,-42}  {1,-10}  {2,-40}" -f $xCond, $xOper, $xValue)
						}
						If($HTML)
						{
							$rowdata += @(,("Conditions: $xCond",($htmlsilver -bor $htmlbold),"$xOper $xValue",$htmlwhite))
						}
					}
				}

				If($MSWord -or $PDF)
				{
					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 250;
					$Table.Columns.Item(2).Width = 250;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
					WriteWordLine 0 0 ""
				}
				If($Text)
				{
					Line 0 ""
				}
				If($HTML)
				{
					$msg = ""
					$columnWidths = @("250","250")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
					WriteHTMLLine 0 0 ""
				}

				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "Options"
				}
				If($Text)
				{
					Line 2 "Options"
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 "Options"
				}

				$PolicyOptions = Get-DHCPServerV4OptionValue -ComputerName $Script:DHCPServerName -PolicyName $Policy.Name -EA 0 | Sort-Object OptionId

				If($? -and $Null -ne $PolicyOptions)
				{
					If($PolicyOptions -is [array] -and $PolicyOptions.Count -eq 2 -and $PolicyOptions[0].OptionId -eq 51 -and $PolicyOptions[1].OptionId -eq 81)
					{
						If($MSWord -or $PDF)
						{
							WriteWordLine 0 1 "None"
						}
						If($Text)
						{
							Line 3 "None"
						}
						If($HTML)
						{
							WriteHTMLLine 0 1 "None"
						}
					}
					Else
					{
						ForEach($PolicyOption in $PolicyOptions)
						{
							If($PolicyOption.OptionId -eq 51 -or $PolicyOption.OptionId -eq 81)
							{
								#ignore these two option IDs
								https://carlwebster.com/the-mysterious-microsoft-dhcp-option-id-81/
								https://jimswirelessworld.wordpress.com/2019/01/03/you-should-care-about-dhcp-option-51/
							}
							Else
							{
								Write-Verbose "$(Get-Date -Format G):	`t`t`tProcessing option name $($PolicyOption.Name)"
								If([string]::IsNullOrEmpty($PolicyOption.VendorClass))
								{
									$VendorClass = "Standard" 
								}
								Else
								{
									$VendorClass = $PolicyOption.VendorClass 
								}

								If([string]::IsNullOrEmpty($PolicyOption.PolicyName))
								{
									$PolicyName = "None"
								}
								Else
								{
									$PolicyName = $PolicyOption.PolicyName
								}

								If($MSWord -or $PDF)
								{
									[System.Collections.Hashtable[]] $ScriptInformation = @()
									$ScriptInformation += @{ Data = "Option Name"; Value = "$($PolicyOption.OptionId.ToString("00000")) $($PolicyOption.Name)"; }
									$ScriptInformation += @{ Data = "Vendor"; Value = $VendorClass; }
									$ScriptInformation += @{ Data = "Value"; Value = "$($PolicyOption.Value)"; }
									$ScriptInformation += @{ Data = "Policy Name"; Value = $PolicyName; }

									$Table = AddWordTable -Hashtable $ScriptInformation `
									-Columns Data,Value `
									-List `
									-Format $wdTableGrid `
									-AutoFit $wdAutoFitFixed;

									SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
									
									$Table.Columns.Item(1).Width = 75;
									$Table.Columns.Item(2).Width = 250;

									$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

									FindWordDocumentEnd
									$Table = $Null
									WriteWordLine 0 0 ""
								}
								If($Text)
								{
									Line 3 "Option Name`t: $($PolicyOption.OptionId.ToString("00000")) $($PolicyOption.Name)" 
									Line 3 "Vendor`t`t: " $VendorClass
									Line 3 "Value`t`t: $($PolicyOption.Value)" 
									Line 3 "Policy Name`t: " $PolicyName
									
									#for spacing
									Line 0 ""
								}
								If($HTML)
								{
									$rowdata = @()
									$columnHeaders = @("Option Name",($htmlsilver -bor $htmlbold),"$($PolicyOption.OptionId.ToString("00000")) $($PolicyOption.Name)",$htmlwhite)
									$rowdata += @(,('Vendor',($htmlsilver -bor $htmlbold),$VendorClass,$htmlwhite))
									$rowdata += @(,('Value',($htmlsilver -bor $htmlbold),"$($PolicyOption.Value)",$htmlwhite))
									$rowdata += @(,('Policy Name',($htmlsilver -bor $htmlbold),$PolicyName,$htmlwhite))

									$msg = ""
									$columnWidths = @("100","400")
									FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
									WriteHTMLLine 0 0 ""
								}
							}
						}
					}
				}
				ElseIf(!$?)
				{
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 0 "Error retrieving scope options for $Policy.Name"
					}
					If($Text)
					{
						Line 0 "Error retrieving Policy options for Policy $Policy.Name"
					}
					If($HTML)
					{
						WriteHTMLLine 0 0 "Error retrieving scope options for $Policy.Name"
					}
					InsertBlankLine
				}
				Else
				{
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 1 "None"
					}
					If($Text)
					{
						Line 3 "None"
					}
					If($HTML)
					{
						WriteHTMLLine 0 1 "None"
					}
					InsertBlankLine
				}
				$PolicyOptions = $Null

				If($MSWord -or $PDF)
				{
					WriteWordLine 4 0 "DNS"
				}
				If($Text)
				{
					Line 2 "DNS"
				}
				If($HTML)
				{
					WriteHTMLLine 4 0 "DNS"
				}
				
				$PolicyDNS = Get-DhcpServerv4DnsSetting -ComputerName $Script:DHCPServerName -PolicyName $Policy.Name -EA 0 
				
				If($? -and $Null -ne $PolicyDNS)
				{
					If($MSWord -or $PDF)
					{
						[System.Collections.Hashtable[]] $ScriptInformation = @()
						If($PolicyDNS.DynamicUpdates -eq "Never")
						{
							$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Disabled"; }
						}
						ElseIf($PolicyDNS.DynamicUpdates -eq "OnClientRequest")
						{
							$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
							$ScriptInformation += @{ Data = "Dynamically update DNS records only if requested by the DHCP clients"; Value = "Enabled"; }
						}
						ElseIf($PolicyDNS.DynamicUpdates -eq "Always")
						{
							$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
							$ScriptInformation += @{ Data = "Always dynamically update DNS records"; Value = "Enabled"; }
						}
						If($PolicyDNS.DeleteDnsRROnLeaseExpiry)
						{
							$ScriptInformation += @{ Data = "Discard A and PTR records when lease deleted"; Value = "Enabled"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = "Discard A and PTR records when lease deleted"; Value = "Disabled"; }
						}
						If($PolicyDNS.UpdateDnsRRForOlderClients)
						{
							$ScriptInformation += @{ Data = "Dynamically update DNS records for DHCP clients that do not request updates"; Value = "Enabled"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = "Dynamically update DNS records for DHCP clients that do not request updates"; Value = "Disabled"; }
						}
						If($PolicyDNS.DisableDnsPtrRRUpdate)
						{
							$ScriptInformation += @{ Data = "Disable dynamic updates for DNS PTR records"; Value = "Enabled"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = "Disable dynamic updates for DNS PTR records"; Value = "Disabled"; }
						}
						If([string]::IsNullOrEmpty($PolicyDNS.DnsSuffix))
						{
							$ScriptInformation += @{ Data = "Register DHCP clients using the following DNS suffix"; Value = "Disabled"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = "Register DHCP clients using the following DNS suffix"; Value = $PolicyDNS.DnsSuffix; }
						}
						If($PolicyDNS.NameProtection)
						{
							$ScriptInformation += @{ Data = "Name Protection"; Value = "Enabled"; }
						}
						Else
						{
							$ScriptInformation += @{ Data = "Name Protection"; Value = "Disabled"; }
						}

						$Table = AddWordTable -Hashtable $ScriptInformation `
						-Columns Data,Value `
						-List `
						-Format $wdTableGrid `
						-AutoFit $wdAutoFitFixed;

						SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						$Table.Columns.Item(1).Width = 400;
						$Table.Columns.Item(2).Width = 50;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

						FindWordDocumentEnd
						$Table = $Null
						WriteWordLine 0 0 ""
					}
					If($Text)
					{
						If($PolicyDNS.DynamicUpdates -eq "Never")
						{
							Line 3 "Enable DNS dynamic updates`t`t`t`t: Disabled"
						}
						ElseIf($PolicyDNS.DynamicUpdates -eq "OnClientRequest")
						{
							Line 3 "Enable DNS dynamic updates`t`t`t`t: Enabled"
							Line 3 "Dynamically update DNS records only "
							Line 3 "if requested by the DHCP clients`t`t`t: Enabled"
						}
						ElseIf($PolicyDNS.DynamicUpdates -eq "Always")
						{
							Line 3 "Enable DNS dynamic updates`t`t`t`t: Enabled"
							Line 3 "Always dynamically update DNS records: Enabled"
						}
						If($PolicyDNS.DeleteDnsRROnLeaseExpiry)
						{
							Line 3 "Discard A and PTR records when lease deleted`t`t: Enabled"
						}
						Else
						{
							Line 3 "Discard A and PTR records when lease deleted`t`t: Disabled"
						}
						Line 3 "Dynamically update DNS records for DHCP "
						If($PolicyDNS.UpdateDnsRRForOlderClients)
						{
							Line 3 "clients that do not request updates`t`t`t: Enabled"
						}
						Else
						{
							Line 3 "clients that do not request updates`t`t`t: Disabled"
						}
						If($PolicyDNS.DisableDnsPtrRRUpdate)
						{
							Line 3 "Disable dynamic updates for DNS PTR records`t`t: Enabled"
						}
						Else
						{
							Line 3 "Disable dynamic updates for DNS PTR records`t`t: Disabled"
						}
						If([string]::IsNullOrEmpty($PolicyDNS.DnsSuffix))
						{
							Line 3 "Register DHCP clients using the following DNS suffix`t: Disabled"
						}
						Else
						{
							Line 3 "Register DHCP clients using the following DNS suffix`t: " $PolicyDNS.DnsSuffix
						}
						If($PolicyDNS.NameProtection)
						{
							Line 3 "Name Protection`t`t`t`t`t`t: Enabled"
						}
						Else
						{
							Line 3 "Name Protection`t`t`t`t`t`t: Disabled"
						}
						Line 0 ""
					}
					If($HTML)
					{
						$rowdata = @()
						If($PolicyDNS.DynamicUpdates -eq "Never")
						{
							$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite)
						}
						ElseIf($PolicyDNS.DynamicUpdates -eq "OnClientRequest")
						{
							$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
							$rowdata += @(,('Dynamically update DNS records only if requested by the DHCP clients',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
						}
						ElseIf($PolicyDNS.DynamicUpdates -eq "Always")
						{
							$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
							$rowdata += @(,('Always dynamically update DNS records',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
						}
						If($PolicyDNS.DeleteDnsRROnLeaseExpiry)
						{
							$rowdata += @(,('Discard A and PTR records when lease deleted',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('Discard A and PTR records when lease deleted',($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
						}
						If($PolicyDNS.UpdateDnsRRForOlderClients)
						{
							$rowdata += @(,('Dynamically update DNS records for DHCP clients that do not request updates',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('Dynamically update DNS records for DHCP clients that do not request updates',($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
						}
						If($PolicyDNS.DisableDnsPtrRRUpdate)
						{
							$rowdata += @(,('Disable dynamic updates for DNS PTR records',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('Disable dynamic updates for DNS PTR records',($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
						}
						If([string]::IsNullOrEmpty($PolicyDNS.DnsSuffix))
						{
							$rowdata += @(,('Register DHCP clients using the following DNS suffix',($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('Register DHCP clients using the following DNS suffix',($htmlsilver -bor $htmlbold),$PolicyDNS.DnsSuffix,$htmlwhite))
						}
						If($PolicyDNS.NameProtection)
						{
							$rowdata += @(,('Name Protection',($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
						}
						Else
						{
							$rowdata += @(,('Name Protection',($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
						}
						
						$msg = ""
						$columnWidths = @("450","50")
						FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
						WriteHTMLLine 0 0 ""
					}
				}
				ElseIf(!$?)
				{
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 0 "Error retrieving Policy DNS settings for Policy $Policy.Name"
					}
					If($Text)
					{
						Line 0 "Error retrieving Policy DNS settings for Policy $Policy.Name"
					}
					If($HTML)
					{
						WriteHTMLLine 0 0 "Error retrieving Policy DNS settings for Policy $Policy.Name"
					}
					InsertBlankLine
				}
				Else
				{
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 1"None"
					}
					If($Text)
					{
						Line 3 "None"
					}
					If($HTML)
					{
						WriteHTMLLine 0 1"None"
					}
					InsertBlankLine
				}
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv4 policies"
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv4 policies"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv4 policies"
		}
		InsertBlankLine
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no IPv4 policies"
		}
		If($Text)
		{
			Line 2 "There were no IPv4 policies"
		}
		If($HTML)
		{
			WriteHTMLLine 0 1 "There were no IPv4 policies"
		}
		InsertBlankLine
	}
	$Policies = $Null
}

Function ProcessIPv4Filters
{
	#Filters
	Write-Verbose "$(Get-Date -Format G): Getting IPv4 filters"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Filters"
		WriteWordLine 4 0 "Allow"
		[System.Collections.Hashtable[]] $FiltersWordTable = @()
	}
	If($Text)
	{
		Line 1 "Filters"
		Line 2 "Allow"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Filters"
		WriteHTMLLine 4 0 "Allow"
		$rowdata = @()
	}

	Write-Verbose "$(Get-Date -Format G): `tAllow filters"
	$AllowFilters = Get-DHCPServerV4Filter -List Allow -ComputerName $Script:DHCPServerName -EA 0 | Sort-Object MacAddress

	If($? -and $Null -ne $AllowFilters)
	{
		ForEach($AllowFilter in $AllowFilters)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				MacAddress = $AllowFilter.MacAddress; `
				Description = $AllowFilter.Description
				}

				## Add the hash to the array
				$FiltersWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 3 "MAC Address`t: " $AllowFilter.MacAddress
				Line 3 "Description`t: " $AllowFilter.Description
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata += @(,($AllowFilter.MacAddress,$htmlwhite,
								$AllowFilter.Description,$htmlwhite))
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			MacAddress = "Error retrieving IPv4 allow filters"; `
			Description = ""
			}

			## Add the hash to the array
			$FiltersWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv4 allow filters"
			Line 0 ""
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 
			$rowdata += @(,("Error retrieving IPv4 allow filters",$htmlwhite,
							"",$htmlwhite))
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			MacAddress = "There were no IPv4 allow filters"; `
			Description = ""
			}

			## Add the hash to the array
			$FiltersWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 2 "There were no IPv4 allow filters"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,("There were no IPv4 allow filters",$htmlwhite,
							"",$htmlwhite))
		}
	}
	$AllowFilters = $Null

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		If($FiltersWordTable.Count -gt 0)
		{
			$Table = AddWordTable -Hashtable $FiltersWordTable `
			-Columns MacAddress,Description `
			-Headers "MAC Address","Description" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	If($HTML)
	{
		$columnHeaders = @('MAC Address',($htmlsilver -bor $htmlbold),'Description',($htmlsilver -bor $htmlbold))
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}

	Write-Verbose "$(Get-Date -Format G): `tDeny filters"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Deny"
		[System.Collections.Hashtable[]] $FiltersWordTable = @()
	}
	If($Text)
	{
		Line 2 "Deny"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Deny"
		$rowdata = @()
	}

	$DenyFilters = Get-DHCPServerV4Filter -List Deny -ComputerName $Script:DHCPServerName -EA 0 | Sort-Object MacAddress
	If($? -and $Null -ne $DenyFilters)
	{
		ForEach($DenyFilter in $DenyFilters)
		{
			If($MSWord -or $PDF)
			{
				$WordTableRowHash = @{ 
				MacAddress = $DenyFilter.MacAddress; `
				Description = $DenyFilter.Description
				}

				## Add the hash to the array
				$FiltersWordTable += $WordTableRowHash;
			}
			If($Text)
			{
				Line 3 "MAC Address`t: " $DenyFilter.MacAddress
				Line 3 "Description`t: " $DenyFilter.Description
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata += @(,($DenyFilter.MacAddress,$htmlwhite,
								$DenyFilter.Description,$htmlwhite))
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			MacAddress = "Error retrieving IPv4 deny filters"; `
			Description = ""
			}

			## Add the hash to the array
			$FiltersWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv4 deny filters"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,("Error retrieving IPv4 deny filters",$htmlwhite,
							"",$htmlwhite))
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			MacAddress = "There were no IPv4 deny filters"; `
			Description = ""
			}

			## Add the hash to the array
			$FiltersWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 2 "There were no IPv4 deny filters"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,("There were no IPv4 deny filters",$htmlwhite,
							"",$htmlwhite))
		}
	}
	$DenyFilters = $Null
	
	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		If($FiltersWordTable.Count -gt 0)
		{
			$Table = AddWordTable -Hashtable $FiltersWordTable `
			-Columns MacAddress,Description `
			-Headers "MAC Address","Description" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
	}
	If($HTML)
	{
		$columnHeaders = @('MAC Address',($htmlsilver -bor $htmlbold),'Description',($htmlsilver -bor $htmlbold))
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}

Function ProcessIPv6Properties
{
	#IPv6

	Write-Verbose "$(Get-Date -Format G): Getting IPv6 properties"
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 2 0 "IPv6"
		WriteWordLine 3 0 "Properties"
	}
	If($Text)
	{
		Line 0 "IPv6"
		Line 0 "Properties"
	}
	If($HTML)
	{
		WriteHTMLLine 2 0 "IPv6"
		WriteHTMLLine 3 0 "Properties"
	}

	Write-Verbose "$(Get-Date -Format G): `tGetting IPv6 general settings"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "General"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	}
	If($Text)
	{
		Line 1 "General"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "General"
		$rowdata = @()
	}

	If($Script:GotAuditSettings)
	{
		If($MSWord -or $PDF)
		{
			If($Script:AuditSettings.Enable)
			{
				$ScriptInformation += @{ Data = "DHCP audit logging is"; Value = "Enabled"; }
			}
			Else
			{
				$ScriptInformation += @{ Data = "DHCP audit logging is"; Value = "Disabled"; }
			}
		}
		If($Text)
		{
			If($Script:AuditSettings.Enable)
			{
				Line 2 "DHCP audit logging is enabled"
			}
			Else
			{
				Line 2 "DHCP audit logging is disabled"
			}
		}
		If($HTML)
		{
			If($Script:AuditSettings.Enable)
			{
				$columnHeaders = @("DHCP audit logging is",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
			}
			Else
			{
				$columnHeaders = @("DHCP audit logging is",($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite)
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Error retrieving audit log settings"; Value = ""; }
		}
		If($Text)
		{
			Line 0 "Error retrieving audit log settings"
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @("Error retrieving audit log settings",($htmlsilver -bor $htmlbold),,$htmlwhite)
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "There were no audit log settings"; Value = ""; }
		}
		If($Text)
		{
			Line 2 "There were no audit log settings"
			Line 0 ""
		}
		If($HTML)
		{
			$columnHeaders = @("There were no audit log settings",($htmlsilver -bor $htmlbold),"",$htmlwhite)
		}
	}

	If($MSWord -or $PDF)
	{
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
	If($HTML)
	{
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}

	#DNS settings
	Write-Verbose "$(Get-Date -Format G): `tGetting IPv6 DNS settings"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "DNS"
	}
	If($Text)
	{
		Line 1 "DNS"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "DNS"
	}

	$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $Script:DHCPServerName -EA 0
	If($? -and $Null -ne $DNSSettings)
	{
		If($DNSSettings.DeleteDnsRROnLeaseExpiry)
		{
			$DeleteDnsRROnLeaseExpiry = "Enabled"
		}
		Else
		{
			$DeleteDnsRROnLeaseExpiry = "Disabled"
		}

		If($DNSSettings.NameProtection)
		{
			$NameProtection = "Enabled"
		}
		Else
		{
			$NameProtection = "Disabled"
		}

		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			If($DNSSettings.DynamicUpdates -eq "Never")
			{
				$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Disabled"; }
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
			{
				$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
				$ScriptInformation += @{ Data = "Dynamically update DNS AAAA and PTR records only if requested by the DHCP clients"; Value = "Enabled"; }
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "Always")
			{
				$ScriptInformation += @{ Data = "Enable DNS dynamic updates"; Value = "Enabled"; }
				$ScriptInformation += @{ Data = "Always dynamically update DNS AAAA and PTR records"; Value = "Enabled"; }
			}
			$ScriptInformation += @{ Data = "Discard AAAA and PTR records when lease is deleted"; Value = $DeleteDnsRROnLeaseExpiry; }
			$ScriptInformation += @{ Data = "Name Protection"; Value = $NameProtection; }

			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 400;
			$Table.Columns.Item(2).Width = 50;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		If($Text)
		{
			Line 2 "Enable DNS dynamic updates`t`t`t: " -NoNewLine
			If($DNSSettings.DynamicUpdates -eq "Never")
			{
				Line 0 "Disabled"
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
			{
				Line 0 "Enabled"
				Line 2 "Dynamically update DNS AAAA and PTR records only "
				Line 2 "if requested by the DHCP clients`t`t: Enabled"
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "Always")
			{
				Line 0 "Enabled"
				Line 2 "Always dynamically update DNS AAAA and PTR records: Enabled"
			}
			Line 2 "Discard AAAA and PTR records when lease deleted`t: " $DeleteDnsRROnLeaseExpiry
			Line 2 "Name Protection`t`t`t`t`t: " $NameProtection
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()
			If($DNSSettings.DynamicUpdates -eq "Never")
			{
				$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite)
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
			{
				$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
				$rowdata += @(,("Dynamically update DNS AAAA and PTR records only if requested by the DHCP clients",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
			}
			ElseIf($DNSSettings.DynamicUpdates -eq "Always")
			{
				$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
				$rowdata += @(,("Always dynamically update DNS AAAA and PTR records",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
			}
			$rowdata += @(,("Discard AAAA and PTR records when lease deleted",($htmlsilver -bor $htmlbold),$DeleteDnsRROnLeaseExpiry,$htmlwhite))
			$rowdata += @(,('Name Protection',($htmlsilver -bor $htmlbold),$NameProtection,$htmlwhite))
			$msg = ""
			$columnWidths = @("450","50")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
			WriteHTMLLine 0 0 
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv6 DNS Settings for DHCP server $Script:DHCPServerName"
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv6 DNS Settings for DHCP server $Script:DHCPServerName"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv6 DNS Settings for DHCP server $Script:DHCPServerName"
		}
		InsertBlankLine
	}
	$DNSSettings = $Null

	#Advanced
	Write-Verbose "$(Get-Date -Format G): `tGetting IPv6 advanced settings"
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Advanced"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
	}
	If($Text)
	{
		Line 1 "Advanced"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Advanced"
		$rowdata = @()
	}

	If($Script:GotAuditSettings)
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Audit log file path"; Value = $Script:AuditSettings.Path; }
		}
		If($Text)
		{
			Line 2 "Audit log file path`t`t: " $Script:AuditSettings.Path
		}
		If($HTML)
		{
			$rowdata += @(,('Audit log file path',($htmlsilver -bor $htmlbold),$Script:AuditSettings.Path,$htmlwhite))
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Audit log file path"; Value = "Unable to determine"; }
		}
		If($Text)
		{
			Line 2 "Audit log file path`t`t: " "Unable to determine"
		}
		If($HTML)
		{
			$rowdata += @(,('Audit log file path',($htmlsilver -bor $htmlbold),"Unable to determine",$htmlwhite))
		}
	}
	$Script:AuditSettings = $Null

	#added 18-Jan-2016
	#get dns update credentials
	Write-Verbose "$(Get-Date -Format G): `tGetting DNS dynamic update registration credentials"
	$DNSUpdateSettings = Get-DhcpServerDnsCredential -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $DNSUpdateSettings)
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "DNS dynamic update registration credentials"; Value = ""; }
			$ScriptInformation += @{ Data = "     User name"; Value = $DNSUpdateSettings.UserName; }
			$ScriptInformation += @{ Data = "     Domain"; Value = $DNSUpdateSettings.DomainName; }
		}
		If($Text)
		{
			Line 2 "DNS dynamic update registration credentials: "
			Line 3 "User name`t: " $DNSUpdateSettings.UserName
			Line 3 "Domain`t`t: " $DNSUpdateSettings.DomainName
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,('DNS dynamic update registration credentials',($htmlsilver -bor $htmlbold),"",$htmlwhite))
			$rowdata += @(,('     User name',($htmlsilver -bor $htmlbold),$DNSUpdateSettings.UserName,$htmlwhite))
			$rowdata += @(,('     Domain',($htmlsilver -bor $htmlbold),$DNSUpdateSettings.DomainName,$htmlwhite))
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "Error retrieving DNS Update Credentials for DHCP server"; Value = $Script:DHCPServerName; }
		}
		If($Text)
		{
			Line 0 "Error retrieving DNS Update Credentials for DHCP server $Script:DHCPServerName"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,('Error retrieving DNS Update Credentials for DHCP server',($htmlsilver -bor $htmlbold),$Script:DHCPServerName,$htmlwhite))
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$ScriptInformation += @{ Data = "There were no DNS Update Credentials for DHCP server"; Value = $Script:DHCPServerName; }
		}
		If($Text)
		{
			Line 2 "There were no DNS Update Credentials for DHCP server $Script:DHCPServerName"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata += @(,('There were no DNS Update Credentials for DHCP server',($htmlsilver -bor $htmlbold),$Script:DHCPServerName,$htmlwhite))
		}
	}
	$DNSUpdateSettings = $Null

	If($MSWord -or $PDF)
	{
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
	If($HTML)
	{
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}
	
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Statistics"
		[System.Collections.Hashtable[]] $StatWordTable = @()
	}
	If($Text)
	{
		Line 1 "Statistics"
	}
	If($HTML)
	{
		WriteHTMLLine 4 0 "Statistics"
		$rowdata = @()
	}
	
	$Statistics = Get-DHCPServerV6Statistics -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $Statistics)
	{
		$UpTime = $(Get-Date) - $Statistics.ServerStartTime
		$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3} seconds", `
			$UpTime.Days, `
			$UpTime.Hours, `
			$UpTime.Minutes, `
			$UpTime.Seconds)

		[int]$InUsePercent = "{0:N0}" -f $Statistics.PercentageInUse.ToString()
		[int]$AvailablePercent = "{0:N0}" -f $Statistics.PercentageAvailable.ToString()
		$AddressesAvailable = "{0:N0}" -f $Statistics.AddressesAvailable
		$TotalAddresses = "{0:N0}" -f $Statistics.TotalAddresses

		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			Description = "Start Time"; `
			Detail = $Statistics.ServerStartTime.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Up Time"; `
			Detail = $Str
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Solicits"; `
			Detail = $Statistics.Solicits.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Advertises"; `
			Detail = $Statistics.Advertises.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Requests"; `
			Detail = $Statistics.Requests.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Replies"; `
			Detail = $Statistics.Replies.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Renews"; `
			Detail = $Statistics.Renews.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Rebinds"; `
			Detail = $Statistics.Rebinds.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Confirms"; `
			Detail = $Statistics.Confirms.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Declines"; `
			Detail = $Statistics.Declines.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Releases"; `
			Detail = $Statistics.Releases.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Total Scopes"; `
			Detail = $Statistics.TotalScopes.ToString()
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Total Addresses"; `
			Detail = $TotalAddresses
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "In Use"; `
			Detail = "$($Statistics.AddressesInUse) - $($InUsePercent)%"
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;

			$WordTableRowHash = @{ 
			Description = "Available"; `
			Detail = "$($AddressesAvailable) - $($AvailablePercent)%"
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 2 "Description" -NoNewLine
			Line 2 "Details"

			Line 2 "Start Time" -NoNewLine
			Line 2 $Statistics.ServerStartTime
			Line 2 "Up Time" -NoNewLine
			Line 3 $Str
			Line 2 "Solicits" -NoNewLine
			Line 2 $Statistics.Solicits
			Line 2 "Advertises" -NoNewLine
			Line 2 $Statistics.Advertises
			Line 2 "Requests" -NoNewLine
			Line 2 $Statistics.Requests
			Line 2 "Replies" -NoNewLine
			Line 3 $Statistics.Replies
			Line 2 "Renews" -NoNewLine
			Line 3 $Statistics.Renews
			Line 2 "Rebinds" -NoNewLine
			Line 3 $Statistics.Rebinds
			Line 2 "Confirms" -NoNewLine
			Line 2 $Statistics.Confirms
			Line 2 "Declines" -NoNewLine
			Line 2 $Statistics.Declines
			Line 2 "Releases" -NoNewLine
			Line 2 $Statistics.Releases
			Line 2 "Total Scopes" -NoNewLine
			Line 2 $Statistics.TotalScopes
			Line 2 "Total Addresses" -NoNewLine
			Line 2 $TotalAddresses
			Line 2 "In Use" -NoNewLine
			Line 3 "$($Statistics.AddressesInUse) - $($InUsePercent)%"
			Line 2 "Available" -NoNewLine
			Line 2 "$($AddressesAvailable) - $($AvailablePercent)%"
			Line 0 ""
		}
		If($HTML)
		{
			$rowdata = @()

			$rowdata += @(,("Start Time",$htmlwhite,$Statistics.ServerStartTime.ToString(),$htmlwhite))
			$rowdata += @(,("Up Time",$htmlwhite,$Str,$htmlwhite))
			$rowdata += @(,("Solicits",$htmlwhite,$Statistics.Solicits.ToString(),$htmlwhite))
			$rowdata += @(,("Advertises",$htmlwhite,$Statistics.Advertises.ToString(),$htmlwhite))
			$rowdata += @(,("Requests",$htmlwhite,$Statistics.Requests.ToString(),$htmlwhite))
			$rowdata += @(,("Replies",$htmlwhite,$Statistics.Replies.ToString(),$htmlwhite))
			$rowdata += @(,("Renews",$htmlwhite,$Statistics.Renews.ToString(),$htmlwhite))
			$rowdata += @(,("Rebinds",$htmlwhite,$Statistics.Rebinds.ToString(),$htmlwhite))
			$rowdata += @(,("Confirms",$htmlwhite,$Statistics.Confirms.ToString(),$htmlwhite))
			$rowdata += @(,("Declines",$htmlwhite,$Statistics.Declines.ToString(),$htmlwhite))
			$rowdata += @(,("Releases",$htmlwhite,$Statistics.Releases.ToString(),$htmlwhite))
			$rowdata += @(,("Total Scopes",$htmlwhite,$Statistics.TotalScopes.ToString(),$htmlwhite))
			$rowdata += @(,("Total Addresses",$htmlwhite,$TotalAddresses,$htmlwhite))
			$rowdata += @(,("In Use",$htmlwhite,"$($Statistics.AddressesInUse) - $($InUsePercent)%",$htmlwhite))
			$rowdata += @(,("Available",$htmlwhite,"$($AddressesAvailable) - $($AvailablePercent)%",$htmlwhite))
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			Description = "Error retrieving IPv6 statistics"; `
			Detail = ""
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv4 statistics"
		}
		If($HTML)
		{
			$rowdata += @(,("Error retrieving IPv6 statistics",$htmlwhite,"",$htmlwhite))
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			Description = "There were no IPv6 statistics"; `
			Detail = ""
			}

			## Add the hash to the array
			$StatWordTable += $WordTableRowHash;
		}
		If($Text)
		{
			Line 0 "There were no IPv6 statistics"
		}
		If($HTML)
		{
			$rowdata += @(,("There were no IPv6 statistics",$htmlwhite,"",$htmlwhite))
		}
	}
	$Statistics = $Null

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		If($StatWordTable.Count -gt 0)
		{
			$Table = AddWordTable -Hashtable $StatWordTable `
			-Columns Description,Detail `
			-Headers "Description","Details" `
			-AutoFit $wdAutoFitContent;

			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
	}
	If($HTML)
	{
		$columnHeaders = @('Description',($htmlsilver -bor $htmlbold),'Details',($htmlsilver -bor $htmlbold))
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}

	Write-Verbose "$(Get-Date -Format G): Getting IPv6 scopes"
	$IPv6Scopes = Get-DHCPServerV6Scope -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $IPv6Scopes)
	{
		If($MSWord -or $PDF)
		{
			$selection.InsertNewPage()
		}
		
		ForEach($IPv6Scope in $IPv6Scopes)
		{
			GetIPv6ScopeData $IPv6Scope
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv6 scopes"
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv6 scopes"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv6 scopes"
		}
		InsertBlankLine
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no IPv6 scopes"
		}
		If($Text)
		{
			Line 1 "There were no IPv6 scopes"
		}
		If($HTML)
		{
			WriteHTMLLine 0 1 "There were no IPv6 scopes"
		}
		InsertBlankLine
	}
	$IPv6Scopes = $Null

	Write-Verbose "$(Get-Date -Format G): Getting IPv6 server options"
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 3 0 "Server Options"
	}
	If($Text)
	{
		Line 0 "Server Options"
	}
	If($HTML)
	{
		WriteHTMLLine 3 0 "Server Options"
	}

	$ServerOptions = Get-DHCPServerV6OptionValue -All -ComputerName $Script:DHCPServerName -EA 0 | Sort-Object OptionId

	If($? -and $Null -ne $ServerOptions)
	{
		ForEach($ServerOption in $ServerOptions)
		{
			If([string]::IsNullOrEmpty($ServerOption.VendorClass))
			{
				$VendorClass =  "Standard"
			}
			Else
			{
				$VendorClass = $ServerOption.VendorClass
			}

			If($MSWord -or $PDF)
			{
				[System.Collections.Hashtable[]] $ScriptInformation = @()

				$ScriptInformation += @{ Data = "Option Name"; Value = "$($ServerOption.OptionId.ToString("00000")) $($ServerOption.Name)"; }
				$ScriptInformation += @{ Data = "Vendor"; Value = $VendorClass; }
				$ScriptInformation += @{ Data = "Value"; Value = $ServerOption.Value; }
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 75;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			If($Text)
			{
				Line 1 "Option Name`t: $($ServerOption.OptionId.ToString("00000")) $($ServerOption.Name)"
				Line 1 "Vendor`t`t: " $VendorClass
				Line 1 "Value`t`t: " $ServerOption.Value
				
				#for spacing
				Line 0 ""
			}
			If($HTML)
			{
				$rowdata = @()
				$columnHeaders = @("Option Name",($htmlsilver -bor $htmlbold),"$($ServerOption.OptionId.ToString("000")) $($ServerOption.Name)",$htmlwhite)
				$rowdata += @(,('Vendor',($htmlsilver -bor $htmlbold),$VendorClass,$htmlwhite))
				$rowdata += @(,('Value',($htmlsilver -bor $htmlbold),$ServerOption.Value[0],$htmlwhite))

				$msg = ""
				$columnWidths = @("100","400")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
				WriteHTMLLine 0 0 ""
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv6 server options"
		}
		If($Text)
		{
			Line 0 "Error retrieving IPv6 server options"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv6 server options"
		}
		InsertBlankLine
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no IPv6 server options"
		}
		If($Text)
		{
			Line 2 "There were no IPv6 server options"
		}
		If($HTML)
		{
			WriteHTMLLine 0 1 "There were no IPv6 server options"
		}
		InsertBlankLine
	}
	$ServerOptions = $Null
}

Function ProcessDHCPOptions
{
	Write-Verbose "$(Get-Date -Format G): Getting DHCP Options"
	
	$DHCPOptions = Get-DhcpServerV4OptionDefinition -ComputerName $Script:DHCPServerName -EA 0
	
	If($? -or $Null -ne $DHCPOptions)
	{
		Write-Verbose "$(Get-Date -Format G): `tProcessing DHCP Options"
		$DHCPOptions = $DHCPOptions | Sort-Object OptionId
	
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ItemsWordTable = @();
			$Selection.InsertNewPage()
			WriteWordLine 2 0 "DHCP Options"
		}
		If($Text)
		{
			Line 0 "DHCP Options"
			Line 0 ""
			Line 1 "OptionId  Name                                                Description                                                   Type          Vendor Class  Default Value         Multivalued"
			Line 1 "========================================================================================================================================================================================="
			       #12345678SS12345678901234567890123456789012345678901234567890SS123456789012345678901234567890123456789012345678901234567890SS123456789012SS123456789012SS12345678901234567890SS12345
		}
		If($HTML)
		{
			$rowdata = @()
			WriteHTMLLine 2 0 "DHCP Options"
		}

		ForEach($Item in $DHCPOptions)
		{
			If($MSWord -or $PDF)
			{
				$ItemsWordTable += @{ 
					OptionId     = $Item.OptionId;
					Name         = $Item.Name;
					Description  = $Item.Description;
					Type         = $Item.Type;
					VendorClass  = $Item.VendorClass;
					DefaultValue = $Item.DefaultValue;
					MultiValued  = $Item.MultiValued;
				}
			}
			If($Text)
			{
				Line 1 ( "{0,-8}  {1,-50}  {2,-60}  {3,-12}  {4,-12}  {5,-20}  {6,-8}" -f `
					$Item.OptionId, 
					$Item.Name, 
					$Item.Description, 
					$Item.Type, 
					$Item.VendorClass, 
					$( If( $null -ne $Item.DefaultValue ) { $Item.DefaultValue -join ';' } Else { ' ' } ), 
					$Item.MultiValued.ToString() 
				)
			}
			If($HTML)
			{
				$rowdata += @(,(
					$Item.OptionId,$htmlwhite,
					$Item.Name,$htmlwhite,
					$Item.Description,$htmlwhite,
					$Item.Type,$htmlwhite,
					$Item.VendorClass,$htmlwhite,
					$Item.DefaultValue.SyncRoot,$htmlwhite,
					$Item.MultiValued.ToString(),$htmlwhite
				))
			}
		}

		If($MSWord -or $PDF)
		{
			If($ItemsWordTable.Count -gt 0)
			{
				$Table = AddWordTable -Hashtable $ItemsWordTable `
				-Columns OptionId, Name, Description, Type, VendorClass, DefaultValue, MultiValued `
				-Headers  "OptionId", "Name", "Description", "Type", "Vendor Class", "Default Value", "Multivalued" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				SetWordCellFormat -Collection $Table -Size 9 -BackgroundColor $wdColorWhite
				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 55;
				$Table.Columns.Item(2).Width = 90;
				$Table.Columns.Item(3).Width = 115;
				$Table.Columns.Item(4).Width = 60;
				$Table.Columns.Item(5).Width = 60;
				$Table.Columns.Item(6).Width = 60;
				$Table.Columns.Item(7).Width = 60;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
			}
		}
		If($HTML)
		{
			$columnHeaders = @(
			'OptionId',($htmlsilver -bor $htmlBold),
			'Name',($htmlsilver -bor $htmlBold),
			'Description',($htmlsilver -bor $htmlBold),
			'Type',($htmlsilver -bor $htmlBold),
			'Vendor Class',($htmlsilver -bor $htmlBold),
			'Default Value',($htmlsilver -bor $htmlBold),
			'Multivalued',($htmlsilver -bor $htmlBold)
			)

			$msg = ""
			$columnWidths = @("60","130","180","80","85","85","80")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "700"
			WriteHTMLLine 0 0 " "
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving DHCP Options"
		}
		If($Text)
		{
			Line 0 "Error retrieving DHCP Options"
		}
		If($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving DHCP Options"
		}
		InsertBlankLine
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no DHCP Options"
		}
		If($Text)
		{
			Line 0 "There were no DHCP Options"
		}
		If($HTML)
		{
			WriteHTMLLine 0 1 "There were no DHCP Options"
		}
		InsertBlankLine
	}
}

Function ProcessHardware
{
	#V1.40 added
	Write-Verbose "$(Get-Date -Format G): Processing Hardware Information"
	If($MSWord -or $PDF)
	{
		$Script:Selection.InsertNewPage()
	}
	GetComputerWMIInfo $Script:DHCPServerName
}
#endregion

#region script setup function
Function ProcessScriptSetup
{
	$script:startTime = Get-Date
	
	#pre 1.40
	#$ComputerName = TestComputerName $ComputerName
	#$Script:DHCPServerName = $ComputerName
	
	#change for 1.40 and -AllDHCPServers
	$Script:DHCPServerNames = @()
	If($AllDHCPServers -eq $False)
	{
		Write-Verbose "$(Get-Date -Format G): Resolving computer name"
		$ComputerName = TestComputerName $ComputerName
		$Script:DHCPServerNames += $ComputerName
	}
	Else
	{
		Write-Verbose "$(Get-Date -Format G): Retrieving all DHCP servers in domain"
		$ComputerName = "All DHCP Servers"
		
		$ALLServers = Get-DHCPServerInDc -EA 0
		
		If($Null -eq $AllServers)
		{
			#oops no DHCP servers
			Write-Error "Unable to retrieve any DHCP servers. Script cannot continue"
			Exit
		}
		Else
		{
			[int]$cnt = 0
			If($AllServers -is [array])
			{
				$cnt = $AllServers.Count
				Write-Verbose "$(Get-Date -Format G): $($cnt) DHCP servers were found"
			}
			Else
			{
				$cnt = 1
				Write-Verbose "$(Get-Date -Format G): $($cnt) DHCP server was found"
			}
			
			$Script:BadDHCPErrorFile = "$($Script:pwdpath)\BadDHCPServers_$(Get-Date -f yyyy-MM-dd_HHmm) for the Domain $Script:RptDomain.txt"

			ForEach($Server in $AllServers)
			{
				$Result = TestComputerName2 $Server.DnsName
				
				If($Result -ne "BAD")
				{
					$Script:DHCPServerNames += $Result
				}
			}
			Write-Verbose "$(Get-Date -Format G): $($Script:DHCPServerNames.Count) DHCP servers will be processed"
			Write-Verbose "$(Get-Date -Format G): "
		}
	}
}
#endregion

#region script end
Function ProcessScriptEnd
{
	Write-Verbose "$(Get-Date -Format G): Script has completed"
	Write-Verbose "$(Get-Date -Format G): "
	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Verbose "$(Get-Date -Format G): Script started: $($Script:StartTime)"
	Write-Verbose "$(Get-Date -Format G): Script ended: $(Get-Date -Format G)"
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
		$runtime.Days, `
		$runtime.Hours, `
		$runtime.Minutes, `
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Verbose "$(Get-Date -Format G): Elapsed time: $($Str)"

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
		$SIFile = "$($Script:pwdpath)\DHCPInventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm) for the Domain $Script:RptDomain.txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime       : $AddDateTime" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name       : $Script:CoName" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Address    : $CompanyAddress" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Email      : $CompanyEmail" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Fax        : $CompanyFax" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Phone      : $CompanyPhone" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page         : $CoverPage" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "ComputerName       : $ComputerName" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Dev                : $Dev" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile       : $Script:DevErrorFile" 4>$Null
		}
		If($MSWord)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word FileName      : $Script:WordFileName" 4>$Null
		}
		If($HTML)
		{
			Out-File -FilePath $SIFile -Append -InputObject "HTML FileName      : $Script:HTMLFileName" 4>$Null
		}
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "PDF Filename       : $Script:PDFFileName" 4>$Null
		}
		If($Text)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Text FileName      : $Script:TextFileName" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder             : $Folder" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From               : $From" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "HW Inventory       : $Hardware" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Include Leases     : $IncludeLeases" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Include Options    : $IncludeOptions" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log                : $Log" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Report Footer      : $ReportFooter" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML       : $HTML" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF        : $PDF" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT       : $TEXT" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD       : $MSWORD" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info        : $ScriptInfo" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port          : $SmtpPort" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server        : $SmtpServer" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title              : $Script:Title" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To                 : $To" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL            : $UseSSL" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "User Name          : $UserName" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected        : $Script:RunningOS" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version       : $Host.Version" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture          : $PSCulture" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture        : $PSUICulture" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word language      : $Script:WordLanguageValue" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word version       : $Script:WordProduct" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start       : $Script:StartTime" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time       : $Str" 4>$Null
	}

	#V1.35 added
	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $true) 
		{
			try 
			{
				Stop-Transcript | Out-Null
				Write-Verbose "$(Get-Date -Format G): $Script:LogPath is ready for use"
			} 
			catch 
			{
				Write-Verbose "$(Get-Date -Format G): Transcript/log stop failed"
			}
		}
	}

	$runtime = $Null
	$Str = $Null
	$ErrorActionPreference = $SaveEAPreference
}
#endregion

#region script core
#Script begins

ProcessScriptSetup

If($AllDHCPServers -eq $False)
{
	[string]$Script:Title = "DHCP Inventory Report for Server $($Script:DHCPServerNames[0]) for the Domain $Script:RptDomain"
	SetFileNames "DHCP Inventory Report for Server $($Script:DHCPServerNames[0]) for the Domain $Script:RptDomain"
}
Else
{
	[string]$Script:Title = "DHCP Inventory Report for All DHCP Servers for the Domain $Script:RptDomain"
	SetFileNames "DHCP Inventory for All DHCP Servers for the Domain $Script:RptDomain"
}

ForEach($DHCPServer in $Script:DHCPServerNames)
{
	Write-Verbose "$(Get-Date -Format G): Processing DHCP Server: $($DHCPServer)"
	$Script:DHCPServerName = $DHCPServer
	
	ProcessServerProperties

	ProcessIPBindings

	ProcessIPv4Properties

	ProcessIPv4Statistics

	ProcessIPv4Superscopes

	ProcessIPv4Scopes

	ProcessIPv4MulticastScopes

	ProcessIPv4BOOTPTable

	ProcessServerOptions

	ProcessPolicies

	ProcessIPv4Filters

	ProcessIPv6Properties
	
	If($IncludeOptions)
	{
		ProcessDHCPOptions
	}

	If($Hardware)
	{
		ProcessHardware
	}
}
#endregion

#region finish script
Write-Verbose "$(Get-Date -Format G): "
Write-Verbose "$(Get-Date -Format G): Finishing up document"
#end of document processing

$AbstractTitle = "DHCP Inventory Report"
$SubjectTitle = "DHCP Inventory Report"

UpdateDocumentProperties $AbstractTitle $SubjectTitle

If($ReportFooter)
{
	OutputReportFooter
}

ProcessDocumentOutput

ProcessScriptEnd
#endregion