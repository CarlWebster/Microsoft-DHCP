#requires -Version 3.0
#requires -Module DHCPServer
#This File is in Unicode format.  Do not edit in an ASCII editor. Notepad++ UTF-8-BOM

#region help text

<#
.SYNOPSIS
	Creates a complete inventory of a Microsoft 2012+ DHCP server.
.DESCRIPTION
	Creates a complete inventory of a Microsoft 2012+ DHCP server using Microsoft 
	PowerShell, Word, plain text, or HTML.
	
	Creates a Word or PDF document, text or HTML file named after the DHCP server.

	Script requires at least PowerShell version 3 but runs best in version 5.
	
	Requires the DHCPServer module.
	Can be run on a DHCP server or on a Windows 8.x or Windows 10 computer with RSAT installed.
		
	Remote Server Administration Tools for Windows 8 
		http://www.microsoft.com/en-us/download/details.aspx?id=28972
		
	Remote Server Administration Tools for Windows 8.1 
		http://www.microsoft.com/en-us/download/details.aspx?id=39296
		
	Remote Server Administration Tools for Windows 10
		http://www.microsoft.com/en-us/download/details.aspx?id=45520
	
	For Windows Server 2003, 2008 and 2008 R2, use the following to export and import the 
	DHCP data:
		Export from the 2003, 2008 or 2008 R2 server:
			netsh dhcp server export C:\DHCPExport.txt all
			
			Copy the C:\DHCPExport.txt file to the 2012+ server.
			
		Import on the 2012+ server:
			netsh dhcp server import c:\DHCPExport.txt all
			
		The script can now be run on the 2012+ DHCP server to document the older DHCP 
		information.

	For Windows Server 2008 and Server 2008 R2, the 2012+ DHCP Server PowerShell cmdlets 
	can be used for the export and import.
		From the 2012+ DHCP server:
			Export-DhcpServer -ComputerName 2008R2Server.domain.tld -Leases -File 
			C:\DHCPExport.xml 
			
			Import-DhcpServer -ComputerName 2012Server.domain.tld -Leases -File 
			C:\DHCPExport.xml -BackupPath C:\dhcp\backup\ 
			
			Note: The c:\dhcp\backup path must exist before the Import-DhcpServer 
			cmdlet is run.
	
	Using netsh is much faster than using the PowerShell export and import cmdlets.
	
	Processing of IPv4 Multicast Scopes is only available with Server 2012 R2 DHCP.
	
	Word and PDF Documents include a Cover Page, Table of Contents and Footer.
	
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
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2020 at 6PM is 2020-06-01_1800.
	Output filename will be ReportName_2020-06-01_1800.docx (or .pdf or .txt).
	This parameter is disabled by default.
.PARAMETER AllDHCPServers
	The script will process all Authorized DHCP servers that are online.
	"All DHCP Servers" is used for the report title.
	This parameter is disabled by default.
	
	If both ComputerName and AllDHCPServers are used, AllDHCPServers is used.
	This parameter has an alias of ALL.
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
	The default value is contained in 
	HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated 
	on the computer running the script.
	This parameter has an alias of CN.
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
	DHCP server to run the script against.
	The computername is used for the report title.
	ComputerName can be entered as the NetBIOS name, FQDN, localhost or IP Address.
	If entered as localhost, the actual computer name is determined and used.
	If entered as an IP address, an attempt is made to determine and use the actual 
	computer name.
	
	If both ComputerName and AllDHCPServers are used, AllDHCPServers is used.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER Hardware
	Use WMI to gather hardware information on Computer System, Disks, Processor, and 
	Network Interface Cards

	This parameter may require the script be run from an elevated PowerShell session 
	using an account with permission to retrieve hardware information (i.e. Domain Admin 
	or Local Administrator).

	Selecting this parameter will add to both the time it takes to run the script and 
	size of the report.

	This parameter is disabled by default.
	This parameter has an alias of HW.
.PARAMETER IncludeLeases
	Include DHCP lease information.
	Default is to not included lease information.
.PARAMETER IncludeOptions
	Include DHCP Options information.
	Default is to not included Options information.
.PARAMETER Log
	Generates a log file for troubleshooting.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	Text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
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
.PARAMETER UserName
	Username to use for the Cover Page and Footer.
	The default value is contained in $env:username
	This parameter has an alias of UN.
	This parameter is only valid with the MSWORD and PDF output parameters.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -ComputerName DHCPServer01
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will be run remotely against DHCP server DHCPServer01.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -ComputerName localhost
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will resolve localhost to $env:computername, for example DHCPServer01.
	Script will be run remotely against DHCP server DHCPServer01 and not localhost.
	Output file name will use the server name DHCPServer01 and not localhost.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -ComputerName 192.168.1.222
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will resolve 192.168.1.222 to the DNS hostname, for example DHCPServer01.
	Script will be run remotely against DHCP server DHCPServer01 and not 192.18.1.222.
	Output file name will use the server name DHCPServer01 and not 192.168.1.222.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -PDF -ComputerName DHCPServer02
	
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will be run remotely against DHCP server DHCPServer02.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -Text -ComputerName DHCPServer02
	
	Script will be run remotely against DHCP server DHCPServer02.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -HTML -ComputerName DHCPServer02
	
	Script will be run remotely against DHCP server DHCPServer02.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -MSWord -ComputerName DHCPServer02
	
	Will use all Default values and save the document as a Word DOCX file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will be run remotely against DHCP server DHCPServer02.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -Hardware -ComputerName DHCPServer02
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will be run remotely against DHCP server DHCPServer02.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -HTML -ComputerName DHCPServer02
	
	Script will be run remotely against DHCP server DHCPServer02.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -ComputerName DHCPServer03 -IncludeLeases
	
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Script will be run remotely against DHCP server DHCPServer03.
	Output will contain DHCP lease information.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -AllDHCPServer -HTML -IncludeOptions
	
	
	The script will find all Authorized DHCP servers and will process all servers that are 
	online.
	Output will contain DHCP Options information.
.EXAMPLE
	PS C:\PSScript .\DHCP_Inventory_V1_4.ps1 -CompanyName "Carl Webster Consulting" 
	-CoverPage "Mod" -UserName "Carl Webster" -ComputerName DHCPServer01

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
	
	Script will be run remotely against DHCP server DHCPServer01.
.EXAMPLE
	PS C:\PSScript .\DHCP_Inventory_V1_4.ps1 -CN "Carl Webster Consulting" -CP "Mod" 
	-UN "Carl Webster" -ComputerName DHCPServer02 -IncludeLeases

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
	
	Script will be run remotely against DHCP server DHCPServer02.
	Output will contain DHCP lease information.
.EXAMPLE
	PS C:\PSScript .\DHCP_Inventory_V1_4.ps1 -CompanyName "Sherlock Holmes 
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
.EXAMPLE
	PS C:\PSScript .\DHCP_Inventory_V1_4.ps1 -CompanyName "Sherlock Holmes 
	Consulting"
	-CoverPage Facet 
	-UserName "Dr. Watson"
	-CompanyEmail SuperSleuth@SherlockHolmes.com

	Will use:
		Sherlock Holmes Consulting for the Company Name.
		Facet for the Cover Page format.
		Dr. Watson for the User Name.
		SuperSleuth@SherlockHolmes.com for the Company Email.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -Folder \\FileServer\ShareName
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 
	-SmtpServer mail.domain.tld
	-From XDAdmin@domain.tld 
	-To ITGroup@domain.tld	

	The script will use the email server mail.domain.tld, sending from XDAdmin@domain.tld, 
	sending to ITGroup@domain.tld.

	The script will use the default SMTP port 25 and will not use SSL.

	If the current user's credentials are not valid to send email, 
	the user will be prompted to enter valid credentials.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 
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
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 
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
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 
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
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 
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
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -Dev -ScriptInfo -Log
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Creates a text file named DHCPInventoryScriptErrors_yyyy-MM-dd_HHmm.txt that 
	contains up to the last 250 errors reported by the script.
	
	Creates a text file named DHCPInventoryScriptInfo_yyyy-MM-dd_HHmm.txt that 
	contains all the script parameters and other basic information.
	
	Creates a text file for transcript logging named 
	DHCPDocScriptTranscript_yyyy-MM-dd_HHmm.txt.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -ComputerName DHCPServer01 -Hardware
	
	Will use all Default values and add additional information for each server about its 
	hardware.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	The script will be run remotely against DHCP server DHCPServer01.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -AllDHCPServers
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	The script will find all Authorized DHCP servers and will process all servers that are 
	online.
.EXAMPLE
	PS C:\PSScript > .\DHCP_Inventory_V1_4.ps1 -ComputerName DHCPServer01 -AllDHCPServers
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Even though DHCPServer01 is specified, the script will find all Authorized DHCP servers 
	and will process all servers that are online.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word, PDF, HTML or 
	formatted text document.
.NOTES
	NAME: DHCP_Inventory_V1_4.ps1
	VERSION: 1.44
	AUTHOR: Carl Webster and Michael B. Smith
	LASTEDIT: April 28, 2020
#>

#endregion


#region script parameters
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
	[Alias("ALL")]
	[Switch]$AllDHCPServers=$False,
	
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
    
	[parameter(ParameterSetName="Word",Mandatory=$False)] 
	[parameter(ParameterSetName="PDF",Mandatory=$False)] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(Mandatory=$False)] 
	[string]$ComputerName="LocalHost",
	
	[parameter(Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(Mandatory=$False)] 
	[string]$Folder="",
	
	[parameter(Mandatory=$False)] 
	[Alias("HW")]
	[Switch]$Hardware=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$IncludeLeases=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$IncludeOptions=$False,

	[parameter(Mandatory=$False)] 
	[Switch]$Log=$False,
	
	[parameter(Mandatory=$False)] 
	[Alias("SI")]
	[Switch]$ScriptInfo=$False,
	
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

#Version 1.44 28-Apr-2020
#	Add checking for a Word version of 0, which indicates the Office installation needs repairing
#	Add Receive Side Scaling setting to Function OutputNICItem
#	Change location of the -Dev, -Log, and -ScriptInfo output files from the script folder to the -Folder location (Thanks to Guy Leech for the "suggestion")
#	Reformatted the terminating Write-Error messages to make them more visible and readable in the console
#	Remove manually checking for multiple output formats
#	Remove the SMTP parameterset and manually verify the parameters
#	Update Function SendEmail to handle anonymous unauthenticated email
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


#V1.35 added
If($Log) 
{
	#start transcript logging
	$Script:LogPath = "$($Script:pwdpath)\DHCPDocScriptTranscript_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	
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
	$Script:DevErrorFile = "$($Script:pwdpath)\DHCPInventoryScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
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
	Write-Verbose "$(Get-Date): CoName is $($Script:CoName)"
	
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[long]$wdColorGray15 = 14277081
	[long]$wdColorGray05 = 15987699 
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[int]$wdColorRed = 255
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
Else
{
	$Script:CoName = ""
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

If($TEXT)
{
	[System.Text.StringBuilder] $global:Output = New-Object System.Text.StringBuilder( 16384 )
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
	# modified 2-Aug-2018 to add ComputerOS information

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
	
	[bool]$GotComputerItems = $True
	
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

	[bool]$GotDrives = $True
	
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

	[bool]$GotProcessors = $True
	
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

		If($Nics -eq $Null ) 
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
	# modified 2-Aug-2018 to add Operating System information
	
	If($MSWord -or $PDF)
	{
		[System.Collections.Hashtable[]] $ItemInformation = @()
		$ItemInformation += @{ Data = "Manufacturer"; Value = $Item.manufacturer; }
		$ItemInformation += @{ Data = "Model"; Value = $Item.model; }
		$ItemInformation += @{ Data = "Domain"; Value = $Item.domain; }
		$ItemInformation += @{ Data = "Operating System"; Value = $OS; }
		$ItemInformation += @{ Data = "Total Ram"; Value = "$($Item.totalphysicalram) GB"; }
		$ItemInformation += @{ Data = "Physical Processors (sockets)"; Value = $Item.NumberOfProcessors; }
		$ItemInformation += @{ Data = "Logical Processors (cores w/HT)"; Value = $Item.NumberOfLogicalProcessors; }
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
		$rowdata += @(,('Operating System',($htmlsilver -bor $htmlbold),$OS,$htmlwhite))
		$rowdata += @(,('Total Ram',($htmlsilver -bor $htmlbold),"$($Item.totalphysicalram) GB",$htmlwhite))
		$rowdata += @(,('Physical Processors (sockets)',($htmlsilver -bor $htmlbold),$Item.NumberOfProcessors,$htmlwhite))
		$rowdata += @(,('Logical Processors (cores w/HT)',($htmlsilver -bor $htmlbold),$Item.NumberOfLogicalProcessors,$htmlwhite))

		$msg = ""
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
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
		[System.Collections.Hashtable[]] $DriveInformation = @()
		$DriveInformation += @{Data = "Caption"; Value = $Drive.caption; }
		$DriveInformation += @{Data = "Size"; Value = "$($drive.drivesize) GB"; }
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$DriveInformation += @{Data = "File System"; Value = $Drive.filesystem; }
		}
		$DriveInformation += @{Data = "Free Space"; Value = "$($drive.drivefreespace) GB"; }
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$DriveInformation += @{Data = "Volume Name"; Value = $Drive.volumename; }
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$DriveInformation += @{Data = "Volume is Dirty"; Value = $xVolumeDirty; }
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$DriveInformation += @{Data = "Volume Serial Number"; Value = $Drive.volumeserialnumber; }
		}
		$DriveInformation += @{Data = "Drive Type"; Value = $xDriveType; }
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
		[System.Collections.Hashtable[]] $ProcessorInformation = @()
		$ProcessorInformation += @{Data = "Name"; Value = $Processor.name; }
		$ProcessorInformation += @{Data = "Description"; Value = $Processor.description; }
		$ProcessorInformation += @{Data = "Max Clock Speed"; Value = "$($processor.maxclockspeed) MHz"; }
		If($processor.l2cachesize -gt 0)
		{
			$ProcessorInformation += @{Data = "L2 Cache Size"; Value = "$($processor.l2cachesize) KB"; }
		}
		If($processor.l3cachesize -gt 0)
		{
			$ProcessorInformation += @{Data = "L3 Cache Size"; Value = "$($processor.l3cachesize) KB"; }
		}
		If($processor.numberofcores -gt 0)
		{
			$ProcessorInformation += @{Data = "Number of Cores"; Value = $Processor.numberofcores; }
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$ProcessorInformation += @{Data = "Number of Logical Processors (cores w/HT)"; Value = $Processor.numberoflogicalprocessors; }
		}
		$ProcessorInformation += @{Data = "Availability"; Value = $xAvailability; }
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

	$xIPAddress = @()
	ForEach($IPAddress in $Nic.ipaddress)
	{
		$xIPAddress += "$($IPAddress)"
	}

	$xIPSubnet = @()
	ForEach($IPSubnet in $Nic.ipsubnet)
	{
		$xIPSubnet += "$($IPSubnet)"
	}

	If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
		$xnicdnsdomainsuffixsearchorder = @()
		ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
		{
			$xnicdnsdomainsuffixsearchorder += "$($DNSDomain)"
		}
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
		[System.Collections.Hashtable[]] $NicInformation = @()
		$NicInformation += @{Data = "Name"; Value = $ThisNic.Name; }
		If($ThisNic.Name -ne $nic.description)
		{
			$NicInformation += @{Data = "Description"; Value = $Nic.description; }
		}
		$NicInformation += @{Data = "Connection ID"; Value = $ThisNic.NetConnectionID; }
		If(validObject $Nic Manufacturer)
		{
			$NicInformation += @{Data = "Manufacturer"; Value = $Nic.manufacturer; }
		}
		$NicInformation += @{Data = "Availability"; Value = $xAvailability; }
		$NicInformation += @{Data = "Allow the computer to turn off this device to save power"; Value = $PowerSaving; }
		$NicInformation += @{Data = "Receive Side Scaling"; Value = $RSSEnabled; }
		$NicInformation += @{Data = "Physical Address"; Value = $Nic.macaddress; }
		If($xIPAddress.Count -gt 1)
		{
			$NicInformation += @{Data = "IP Address"; Value = $xIPAddress[0]; }
			$NicInformation += @{Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }
			$NicInformation += @{Data = "Subnet Mask"; Value = $xIPSubnet[0]; }
			$cnt = -1
			ForEach($tmp in $xIPAddress)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += @{Data = "IP Address"; Value = $tmp; }
					$NicInformation += @{Data = "Subnet Mask"; Value = $xIPSubnet[$cnt]; }
				}
			}
		}
		Else
		{
			$NicInformation += @{Data = "IP Address"; Value = $xIPAddress; }
			$NicInformation += @{Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }
			$NicInformation += @{Data = "Subnet Mask"; Value = $xIPSubnet; }
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$NicInformation += @{Data = "DHCP Enabled"; Value = $Nic.dhcpenabled; }
			$NicInformation += @{Data = "DHCP Lease Obtained"; Value = $dhcpleaseobtaineddate; }
			$NicInformation += @{Data = "DHCP Lease Expires"; Value = $dhcpleaseexpiresdate; }
			$NicInformation += @{Data = "DHCP Server"; Value = $Nic.dhcpserver; }
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$NicInformation += @{Data = "DNS Domain"; Value = $Nic.dnsdomain; }
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$NicInformation += @{Data = "DNS Search Suffixes"; Value = $xnicdnsdomainsuffixsearchorder[0]; }
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += @{Data = ""; Value = $tmp; }
				}
			}
		}
		$NicInformation += @{Data = "DNS WINS Enabled"; Value = $xdnsenabledforwinsresolution; }
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$NicInformation += @{Data = "DNS Servers"; Value = $xnicdnsserversearchorder[0]; }
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += @{Data = ""; Value = $tmp; }
				}
			}
		}
		$NicInformation += @{Data = "NetBIOS Setting"; Value = $xTcpipNetbiosOptions; }
		$NicInformation += @{Data = "WINS: Enabled LMHosts"; Value = $xwinsenablelmhostslookup; }
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$NicInformation += @{Data = "Host Lookup File"; Value = $Nic.winshostlookupfile; }
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$NicInformation += @{Data = "Primary Server"; Value = $Nic.winsprimaryserver; }
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$NicInformation += @{Data = "Secondary Server"; Value = $Nic.winssecondaryserver; }
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$NicInformation += @{Data = "Scope ID"; Value = $Nic.winsscopeid; }
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
		Line 2 "Physical Address`t: " $nic.macaddress
		Line 2 "Receive Side Scaling`t: " $RSSEnabled
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
#			'fr-'	{ 'Sommaire Automatique 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break } #changed 10-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			# fix in 1.42 thanks to Johan Kallio 'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
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
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n"
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}) -ne $Null
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

Function ValidateWordTableValues 
{
	Param([int]$Rows, [string]$DHCPServer, [string]$DoingWhat)
	
	If(($Rows -lt 1) -or ($Rows -gt 32767))
	{
		Write-Host "`n`n`t`tUh Oh! Something bad unexpected happened" -ForegroundColor Red
		Write-Host "`n`n`t`tThe Word variable Rows is an unexpected value of $Rows" -ForegroundColor Red
		Write-Host "`n`n`t`tThe DHCP server being processed is $DHCPServer" -ForegroundColor Red
		Write-Host "`n`n`t`tOther info: $DoingWhat" -ForegroundColor Red
		Write-Host "`n`n`t`tPlease email webster@carlwebster.com with this information (a screenshot would be helpful)`n`n" -ForegroundColor Red
		Write-Error "
		`n`n
		`t`t
		Script cannot continue.
		`n`n
		"
		ProcessScriptEnd
		AbortScript
	}
	Else
	{
		Return
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
					$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}).Id
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
	$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}).Id
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

	Write-Output $Global:Output.ToString() | Out-File $Script:Filename1 4>$Null
}

Function SaveandCloseHTMLDocument
{
	Out-File -FilePath $Script:FileName1 -Append -InputObject "<p></p></body></html>" 4>$Null
}

Function SetFileName1andFileName2
{
	Param([string]$OutputFileName)
	
	#set $Script:Filename1 and $Script:Filename2 with no file extension
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

Function ShowScriptOptions
{
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): AddDateTime     : $($AddDateTime)"
	If($MSWord -or $PDF)
	{
		Write-Verbose "$(Get-Date): Company Name    : $($Script:CoName)"
		Write-Verbose "$(Get-Date): Company Address : $($CompanyAddress)"
		Write-Verbose "$(Get-Date): Company Email   : $($CompanyEmail)"
		Write-Verbose "$(Get-Date): Company Fax     : $($CompanyFax)"
		Write-Verbose "$(Get-Date): Company Phone   : $($CompanyPhone)"
		Write-Verbose "$(Get-Date): Cover Page      : $($CoverPage)"
	}
	Write-Verbose "$(Get-Date): ComputerName    : $($ComputerName)"
	Write-Verbose "$(Get-Date): Dev             : $($Dev)"
	If($Dev)
	{
		Write-Verbose "$(Get-Date): DevErrorFile    : $($Script:DevErrorFile)"
	}
	Write-Verbose "$(Get-Date): Filename1       : $($Script:Filename1)"
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Filename2       : $($Script:Filename2)"
	}
	Write-Verbose "$(Get-Date): Folder          : $($Folder)"
	Write-Verbose "$(Get-Date): From            : $($From)"
	Write-Verbose "$(Get-Date): HW Inventory    : $($Hardware)"
	Write-Verbose "$(Get-Date): Include Leases  : $($IncludeLeases)"
	Write-Verbose "$(Get-Date): Include Options : $($IncludeOptions)"
	Write-Verbose "$(Get-Date): Log             : $($Log)"
	Write-Verbose "$(Get-Date): Save As HTML    : $($HTML)"
	Write-Verbose "$(Get-Date): Save As PDF     : $($PDF)"
	Write-Verbose "$(Get-Date): Save As TEXT    : $($TEXT)"
	Write-Verbose "$(Get-Date): Save As WORD    : $($MSWORD)"
	Write-Verbose "$(Get-Date): ScriptInfo      : $($ScriptInfo)"
	Write-Verbose "$(Get-Date): Smtp Port       : $($SmtpPort)"
	Write-Verbose "$(Get-Date): Smtp Server     : $($SmtpServer)"
	Write-Verbose "$(Get-Date): Title           : $($Script:Title)"
	Write-Verbose "$(Get-Date): To              : $($To)"
	Write-Verbose "$(Get-Date): Use SSL         : $($UseSSL)"
	Write-Verbose "$(Get-Date): User Name       : $($UserName)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): OS Detected     : $($Script:RunningOS)"
	Write-Verbose "$(Get-Date): PoSH version    : $($Host.Version)"
	Write-Verbose "$(Get-Date): PSCulture       : $($PSCulture)"
	Write-Verbose "$(Get-Date): PSUICulture     : $($PSUICulture)"
	If($MSWORD -or $PDF)
	{
		Write-Verbose "$(Get-Date): Word language   : $($Script:WordLanguageValue)"
		Write-Verbose "$(Get-Date): Word version    : $($Script:WordProduct)"
	}
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): Script start    : $($Script:StartTime)"
	Write-Verbose "$(Get-Date): "
	Write-Verbose "$(Get-Date): "

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
			Remove-Variable -Name word -Scope Global 4>$Null
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
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
	ElseIf($Text)
	{
		Line 1 $txt
		Line 0 ""
	}
	ElseIf($HTML)
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
	ElseIf($Text)
	{
		Line 0 ""
	}
	ElseIf($HTML)
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
		Write-Verbose "$(Get-Date): Testing to see if $($CName) is online and reachable"
		If(Test-Connection -ComputerName $CName -quiet)
		{
			Write-Verbose "$(Get-Date): Server $($CName) is online."
		}
		Else
		{
			Write-Verbose "$(Get-Date): Computer $($CName) is offline"
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
		Write-Verbose "$(Get-Date): Computer name has been renamed from localhost to $($CName)"
		Write-Verbose "$(Get-Date): Testing to see if $($CName) is a DHCP Server"
		$results = Get-DHCPServerVersion -ComputerName $CName -EA 0
		If($? -and $Null -ne $results)
		{
			#the computer is a dhcp server
			Write-Verbose "$(Get-Date): Computer $($CName) is a DHCP Server"
			Return $CName
		}
		ElseIf(!$? -or $Null -eq $results)
		{
			#the computer is not a dhcp server
			Write-Verbose "$(Get-Date): Computer $($CName) is not a DHCP Server"
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
			Write-Verbose "$(Get-Date): Computer name has been renamed from $($ip) to $($CName)"
			Write-Verbose "$(Get-Date): Testing to see if $($CName) is a DHCP Server"
			$results = Get-DHCPServerVersion -ComputerName $CName -EA 0
			If($? -and $Null -ne $results)
			{
				#the computer is a dhcp server
				Write-Verbose "$(Get-Date): Computer $($CName) is a DHCP Server"
				Return $CName
			}
			ElseIf(!$? -or $Null -eq $results)
			{
				#the computer is not a dhcp server
				Write-Verbose "$(Get-Date): Computer $($CName) is not a DHCP Server"
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
		Write-Verbose "$(Get-Date): Testing to see if $($CName) is a DHCP Server"
		$results = Get-DHCPServerVersion -ComputerName $CName -EA 0
		If($? -and $Null -ne $results)
		{
			#the computer is a dhcp server
			Write-Verbose "$(Get-Date): Computer $($CName) is a DHCP Server"
			Return $CName
		}
		ElseIf(!$? -or $Null -eq $results)
		{
			#the computer is not a dhcp server
			Write-Verbose "$(Get-Date): Computer $($CName) is not a DHCP Server"
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
		Write-Verbose "$(Get-Date): Testing to see if $($CName) is online and reachable"
		If(Test-Connection -ComputerName $CName -quiet)
		{
			Write-Verbose "$(Get-Date): Server $($CName) is online."
		}
		Else
		{
			Write-Verbose "$(Get-Date): Computer $($CName) is offline"
			Write-Output "$(Get-Date): Computer $($CName) is offline" | Out-File $Script:BadDHCPErrorFile -Append 4>$Null
			Return "BAD"
		}
	}

	#if computer name is localhost, get actual computer name
	If($CName -eq "localhost")
	{
		$CName = $env:ComputerName
		Write-Verbose "$(Get-Date): Computer name has been renamed from localhost to $($CName)"
		Write-Verbose "$(Get-Date): Testing to see if $($CName) is a DHCP Server"
		$results = Get-DHCPServerVersion -ComputerName $CName -EA 0
		If($? -and $Null -ne $results)
		{
			#the computer is a dhcp server
			Write-Verbose "$(Get-Date): Computer $($CName) is a DHCP Server"
			Return $CName
		}
		ElseIf(!$? -or $Null -eq $results)
		{
			#the computer is not a dhcp server
			Write-Verbose "$(Get-Date): Computer $($CName) is not a DHCP Server"
			Write-Output "$(Get-Date): Computer $($CName) is not a DHCP Server" | Out-File $Script:BadDHCPErrorFile -Append 4>$Null
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
			Write-Verbose "$(Get-Date): Computer name has been renamed from $($ip) to $($CName)"
			Write-Verbose "$(Get-Date): Testing to see if $($CName) is a DHCP Server"
			$results = Get-DHCPServerVersion -ComputerName $CName -EA 0
			If($? -and $Null -ne $results)
			{
				#the computer is a dhcp server
				Write-Verbose "$(Get-Date): Computer $($CName) is a DHCP Server"
				Return $CName
			}
			ElseIf(!$? -or $Null -eq $results)
			{
				#the computer is not a dhcp server
				Write-Verbose "$(Get-Date): Computer $($CName) is not a DHCP Server"
				Write-Output "$(Get-Date): Computer $($CName) is not a DHCP Server" | Out-File $Script:BadDHCPErrorFile -Append 4>$Null
				Return "BAD"
			}
		}
		Else
		{
			Write-Verbose "$(Get-Date): Unable to resolve $($CName) to a hostname"
			Write-Output "$(Get-Date): Unable to resolve $($CName) to a hostname" | Out-File $Script:BadDHCPErrorFile -Append 4>$Null
			Return "BAD"
		}
	}
	Else
	{
		Write-Verbose "$(Get-Date): Testing to see if $($CName) is a DHCP Server"
		$results = Get-DHCPServerVersion -ComputerName $CName -EA 0
		If($? -and $Null -ne $results)
		{
			#the computer is a dhcp server
			Write-Verbose "$(Get-Date): Computer $($CName) is a DHCP Server"
			Return $CName
		}
		ElseIf(!$? -or $Null -eq $results)
		{
			#the computer is not a dhcp server
			Write-Verbose "$(Get-Date): Computer $($CName) is not a DHCP Server"
			Write-Output "$(Get-Date): Computer $($CName) is not a DHCP Server" | Out-File $Script:BadDHCPErrorFile -Append 4>$Null
			Return "BAD"
		}
	}

	Write-Verbose "$(Get-Date): "
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
	This function formats an HTML line
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

	0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	InsertBlankLine

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
# re-implemented/re-based by Michael B. Smith
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

	## FIXME - long story short, this function was wrong and had been wrong for a long time. 
	## The function generated invalid HTML, and ignored fontname and fontsize parameters. I fixed
	## those items, but that made the report unreadable, because all of the formatting had been based
	## on this function not working properly.

	## here is a typical H1 previously generated:
	## <h1>///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\<font face='Calibri' color='#000000' size='1'></h1></font>

	## fixing the function generated this (unreadably small):
	## <h1><font face='Calibri' color='#000000' size='1'>///&nbsp;&nbsp;Forest Information&nbsp;&nbsp;\\\</font></h1>

	## So I took all the fixes out. This routine now generates valid HTML, but the fontName, fontSize,
	## and options parameters are ignored; so the routine generates equivalent output as before. I took
	## the fixes out instead of fixing all the call sites, because there are 225 call sites! If you are
	## willing to update all the call sites, you can easily re-instate the fixes. They have only been
	## commented out with '##' below.

	## if( [String]::IsNullOrEmpty( $fontName ) )
	## {
	##	$fontName = 'Calibri'
	## }
	## if( $fontSize -le 0 )
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
		[Bool] $ital = $options -band $htmlitalics
		[Bool] $bold = $options -band $htmlBold
		## $color = $global:htmlColor[ $options -band 0xffffc ]

		## ## build the HTML output string
##		$HTMLBody = ''
##		if( $ital ) { $HTMLBody += '<i>' }
##		if( $bold ) { $HTMLBody += '<b>' } 
		if( $ital ) { $null = $sb.Append( '<i>' ) }
		if( $bold ) { $null = $sb.Append( '<b>' ) } 

		switch( $style )
		{
			1 { $HTMLOpen = '<h1>'; $HTMLClose = '</h1>'; Break }
			2 { $HTMLOpen = '<h2>'; $HTMLClose = '</h2>'; Break }
			3 { $HTMLOpen = '<h3>'; $HTMLClose = '</h3>'; Break }
			4 { $HTMLOpen = '<h4>'; $HTMLClose = '</h4>'; Break }
			Default { $HTMLOpen = ''; $HTMLClose = ''; Break }
		}

		## $HTMLBody += $HTMLOpen
		$null = $sb.Append( $HTMLOpen )

		## if($HTMLClose -eq '')
		## {
		##	$HTMLBody += "<br><font face='" + $fontName + "' " + "color='" + $color + "' size='"  + $fontSize + "'>"
		## }
		## else
		## {
		##	$HTMLBody += "<font face='" + $fontName + "' " + "color='" + $color + "' size='"  + $fontSize + "'>"
		## }
		
##		while( $tabs -gt 0 )
##		{ 
##			$output += '&nbsp;&nbsp;&nbsp;&nbsp;'
##			$tabs--
##		}
		## output the rest of the parameters.
##		$output += $name + $value
		## $HTMLBody += $output
		$null = $sb.Append( ( '&nbsp;&nbsp;&nbsp;&nbsp;' * $tabs ) + $name + $value )

		## $HTMLBody += '</font>'
##		if( $HTMLClose -eq '' ) { $HTMLBody += '<br>'     }
##		else                    { $HTMLBody += $HTMLClose }

##		if( $ital ) { $HTMLBody += '</i>' }
##		if( $bold ) { $HTMLBody += '</b>' } 

##		if( $HTMLClose -eq '' ) { $HTMLBody += '<br />' }

		if( $HTMLClose -eq '' ) { $null = $sb.Append( '<br>' )     }
		else                    { $null = $sb.Append( $HTMLClose ) }

		if( $ital ) { $null = $sb.Append( '</i>' ) }
		if( $bold ) { $null = $sb.Append( '</b>' ) } 

		if( $HTMLClose -eq '' ) { $null = $sb.Append( '<br />' ) }
	}
	##$HTMLBody += $crlf
	$null = $sb.AppendLine( '' )

##	Out-File -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null
	Out-File -FilePath $Script:FileName1 -Append -InputObject $sb.ToString() 4>$Null
}
#endregion

#region HTML table functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable function
# Created by Ken Avram
# modified by Jake Rutski
# re-implemented by Michael B. Smith for v2.23. Also made the documentation match reality.
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
	## Use StringBuilder - MBS
	## In the normal case, tables are only a few dozen cells. But in the case
	## of Sites, OUs, and Users - there may be many hundreds of thousands of 
	## cells. Using normal strings is too slow.

	## if( $ExtraSpecialVerbose )
	## {
	##	$global:rowInfo1 = $rowInfo
	## }
<#
	if( $SuperVerbose )
	{
		wv "AddHTMLTable: fontName '$fontName', fontsize $fontSize, colCount $colCount, rowCount $rowCount"
		if( $null -ne $rowInfo -and $rowInfo.Count -gt 0 )
		{
			wv "AddHTMLTable: rowInfo has $( $rowInfo.Count ) elements"
			if( $ExtraSpecialVerbose )
			{
				wv "AddHTMLTable: rowInfo length $( $rowInfo.Length )"
				for( $ii = 0; $ii -lt $rowInfo.Length; $ii++ )
				{
					$row = $rowInfo[ $ii ]
					wv "AddHTMLTable: index $ii, type $( $row.GetType().FullName ), length $( $row.Length )"
					for( $yyy = 0; $yyy -lt $row.Length; $yyy++ )
					{
						wv "AddHTMLTable: index $ii, yyy = $yyy, val = '$( $row[ $yyy ] )'"
					}
					wv "AddHTMLTable: done"
				}
			}
		}
		else
		{
			wv "AddHTMLTable: rowInfo is empty"
		}
		if( $null -ne $fixedInfo -and $fixedInfo.Count -gt 0 )
		{
			wv "AddHTMLTable: fixedInfo has $( $fixedInfo.Count ) elements"
		}
		else
		{
			wv "AddHTMLTable: fixedInfo is empty"
		}
	}
#>

	##$htmlbody = ''
	[System.Text.StringBuilder] $sb = New-Object System.Text.StringBuilder( 8192 )

	if( $rowInfo -and $rowInfo.Length -lt $rowCount )
	{
##		$oldCount = $rowCount
		$rowCount = $rowInfo.Length
##		if( $SuperVerbose )
##		{
##			wv "AddHTMLTable: updated rowCount to $rowCount from $oldCount, based on rowInfo.Length"
##		}
	}

	for( $rowCountIndex = 0; $rowCountIndex -lt $rowCount; $rowCountIndex++ )
	{
		$null = $sb.AppendLine( '<tr>' )
		## $htmlbody += '<tr>'
		## $htmlbody += $crlf - make the HTML readable

		## each row of rowInfo is an array
		## each row consists of tuples: an item of text followed by an item of formatting data
<#		
		$row = $rowInfo[ $rowCountIndex ]
		if( $ExtraSpecialVerbose )
		{
			wv "!!!!! AddHTMLTable: rowCountIndex = $rowCountIndex, row.Length = $( $row.Length ), row gettype = $( $row.GetType().FullName )"
			wv "!!!!! AddHTMLTable: colCount $colCount"
			wv "!!!!! AddHTMLTable: row[0].Length $( $row[0].Length )"
			wv "!!!!! AddHTMLTable: row[0].GetType $( $row[0].GetType().FullName )"
			$subRow = $row
			if( $subRow -is [Array] -and $subRow[ 0 ] -is [Array] )
			{
				$subRow = $subRow[ 0 ]
				wv "!!!!! AddHTMLTable: deref subRow.Length $( $subRow.Length ), subRow.GetType $( $subRow.GetType().FullName )"
			}

			for( $columnIndex = 0; $columnIndex -lt $subRow.Length; $columnIndex += 2 )
			{
				$item = $subRow[ $columnIndex ]
				wv "!!!!! AddHTMLTable: item.GetType $( $item.GetType().FullName )"
				## if( !( $item -is [String] ) -and $item -is [Array] )
##				if( $item -is [Array] -and $item[ 0 ] -is [Array] )				
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
		if( $subRow -is [Array] -and $subRow[ 0 ] -is [Array] )
		{
			$subRow = $subRow[ 0 ]
			## wv "***** AddHTMLTable: deref rowCountIndex $rowCountIndex, subRow.Length $( $subRow.Length ), subRow.GetType $( $subRow.GetType().FullName )"
		}

		$subRowLength = $subRow.Length
		for( $columnIndex = 0; $columnIndex -lt $colCount; $columnIndex += 2 )
		{
			$item = if( $columnIndex -lt $subRowLength ) { $subRow[ $columnIndex ] } else { 0 }
			## if( !( $item -is [String] ) -and $item -is [Array] )
##			if( $item -is [Array] -and $item[ 0 ] -is [Array] )
##			{
##				$item = $item[ 0 ]
##			}

			$text   = if( $item ) { $item.ToString() } else { '' }
			$format = if( ( $columnIndex + 1 ) -lt $subRowLength ) { $subRow[ $columnIndex + 1 ] } else { 0 }
			## item, text, and format ALWAYS have values, even if empty values
			$color  = $global:htmlColor[ $format -band 0xffffc ]
			[Bool] $bold = $format -band $htmlBold
			[Bool] $ital = $format -band $htmlitalics
<#			
			if( $ExtraSpecialVerbose )
			{
				wv "***** columnIndex $columnIndex, subRow.Length $( $subRow.Length ), item GetType $( $item.GetType().FullName ), item '$item'"
				wv "***** format $format, color $color, text '$text'"
				wv "***** format gettype $( $format.GetType().Fullname ), text gettype $( $text.GetType().Fullname )"
			}
#>

			if( $null -eq $fixedInfo -or $fixedInfo.Length -eq 0 )
			{
				$null = $sb.Append( "<td style=""background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>" )
				##$htmlbody += "<td style=""background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>"
			}
			else
			{
				$null = $sb.Append( "<td style=""width:$( $fixedInfo[ $columnIndex / 2 ] ); background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>" )
				##$htmlbody += "<td style=""width:$( $fixedInfo[ $columnIndex / 2 ] ); background-color:$( $color )""><font face='$( $fontName )' size='$( $fontSize )'>"
			}

			##if( $bold ) { $htmlbody += '<b>' }
			##if( $ital ) { $htmlbody += '<i>' }
			if( $bold ) { $null = $sb.Append( '<b>' ) }
			if( $ital ) { $null = $sb.Append( '<i>' ) }

			if( $text -eq ' ' -or $text.length -eq 0)
			{
				##$htmlbody += '&nbsp;&nbsp;&nbsp;'
				$null = $sb.Append( '&nbsp;&nbsp;&nbsp;' )
			}
			else
			{
				for ($inx = 0; $inx -lt $text.length; $inx++ )
				{
					if( $text[ $inx ] -eq ' ' )
					{
						##$htmlbody += '&nbsp;'
						$null = $sb.Append( '&nbsp;' )
					}
					else
					{
						break
					}
				}
				##$htmlbody += $text
				$null = $sb.Append( $text )
			}

##			if( $bold ) { $htmlbody += '</b>' }
##			if( $ital ) { $htmlbody += '</i>' }
			if( $bold ) { $null = $sb.Append( '</b>' ) }
			if( $ital ) { $null = $sb.Append( '</i>' ) }

			$null = $sb.AppendLine( '</font></td>' )
##			$htmlbody += '</font></td>'
##			$htmlbody += $crlf
		}

		$null = $sb.AppendLine( '</tr>' )
##		$htmlbody += '</tr>'
##		$htmlbody += $crlf
	}

##	if( $ExtraSpecialVerbose )
##	{
##		$global:rowInfo = $rowInfo
##		wv "!!!!! AddHTMLTable: HTML = '$htmlbody'"
##	}

##	Out-File -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
	Out-File -FilePath $Script:FileName1 -Append -InputObject $sb.ToString() 4>$Null 
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

	## FIXME - the help text for this function is wacky wrong - MBS
	## FIXME - Use StringBuilder - MBS - this only builds the table header - benefit relatively small
<#
	if( $SuperVerbose )
	{
		wv "FormatHTMLTable: fontname '$fontname', size $fontSize, tableheader '$tableheader'"
		wv "FormatHTMLTable: noborder $noborder, noheadcols $noheadcols"
		if( $rowarray -and $rowarray.count -gt 0 )
		{
			wv "FormatHTMLTable: rowarray has $( $rowarray.count ) elements"
		}
		else
		{
			wv "FormatHTMLTable: rowarray is empty"
		}
		if( $columnarray -and $columnarray.count -gt 0 )
		{
			wv "FormatHTMLTable: columnarray has $( $columnarray.count ) elements"
		}
		else
		{
			wv "FormatHTMLTable: columnarray is empty"
		}
		if( $fixedwidth -and $fixedwidth.count -gt 0 )
		{
			wv "FormatHTMLTable: fixedwidth has $( $fixedwidth.count ) elements"
		}
		else
		{
			wv "FormatHTMLTable: fixedwidth is empty"
		}
	}
#>

	$HTMLBody = "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>" + $crlf

	If( $null -eq $columnArray -or $columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If( $null -ne $rowArray )
	{
		$NumRows = $rowArray.length + 1
	}
	Else
	{
		$NumRows = 1
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

	if( $columnArray -and $columnArray.Length -gt 0 )
	{
		$HTMLBody += '<tr>' + $crlf

		for( $columnIndex = 0; $columnIndex -lt $NumCols; $columnindex += 2 )
		{
			$val = $columnArray[ $columnIndex + 1 ]
			$tmp = $global:htmlColor[ $val -band 0xffffc ]
			[Bool] $bold = $val -band $htmlBold
			[Bool] $ital = $val -band $htmlitalics

			if( $null -eq $fixedWidth -or $fixedWidth.Length -eq 0 )
			{
				$HTMLBody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			else
			{
				$HTMLBody += "<td style=""width:$($fixedWidth[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			if( $bold ) { $HTMLBody += '<b>' }
			if( $ital ) { $HTMLBody += '<i>' }

			$array = $columnArray[ $columnIndex ]
			if( $array )
			{
				if( $array -eq ' ' -or $array.Length -eq 0 )
				{
					$HTMLBody += '&nbsp;&nbsp;&nbsp;'
				}
				else
				{
					for( $i = 0; $i -lt $array.Length; $i += 2 )
					{
						if( $array[ $i ] -eq ' ' )
						{
							$HTMLBody += '&nbsp;'
						}
						else
						{
							break
						}
					}
					$HTMLBody += $array
				}
			}
			else
			{
				$HTMLBody += '&nbsp;&nbsp;&nbsp;'
			}
			
			if( $bold ) { $HTMLBody += '</b>' }
			if( $ital ) { $HTMLBody += '</i>' }
		}

		$HTMLBody += '</font></td>'
		$HTMLBody += $crlf
	}

	$HTMLBody += '</tr>' + $crlf

	Out-File -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
	$HTMLBody = ''

	##$rowindex = 2
	If( $rowArray )
	{
<#
		if( $ExtraSpecialVerbose )
		{
			wv "***** FormatHTMLTable: rowarray length $( $rowArray.Length )"
			for( $ii = 0; $ii -lt $rowArray.Length; $ii++ )
			{
				$row = $rowArray[ $ii ]
				wv "***** FormatHTMLTable: index $ii, type $( $row.GetType().FullName ), length $( $row.Length )"
				for( $yyy = 0; $yyy -lt $row.Length; $yyy++ )
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

	Out-File -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 4>$Null 
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
}#endregion

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
				} ## end foreach
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
				} ## end foreach

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
		[Parameter()] [AllowNull()] $BackgroundColor = $Null,
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
				} # end foreach
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

#region email function
Function SendEmail
{
	Param([array]$Attachments)
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

#region DHCP script functions
Function ProcessServerProperties
{
	Write-Verbose "$(Get-Date): Server Properties and Configuration"
	Write-Verbose "$(Get-Date): "

	Write-Verbose "$(Get-Date): Getting DHCP server information"
	
	$tmp = $Script:DHCPServerName.Split(".")
	$NetBIOSName = $tmp[0]
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "DHCP Server Information: " $NetBIOSName
		WriteWordLine 0 0 "Server name: " $Script:DHCPServerName
	}
	ElseIf($Text)
	{
		Line 0 "DHCP Server Information: " $NetBIOSName
		Line 1 "Server name`t: " $Script:DHCPServerName
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "DHCP Server Information: " $NetBIOSName
		WriteHTMLLine 0 0 "Server name: " $Script:DHCPServerName
	}

	$DHCPDB = Get-DHCPServerDatabase -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $DHCPDB)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Database path: " $DHCPDB.FileName.SubString(0,($DHCPDB.FileName.LastIndexOf('\')))
			WriteWordLine 0 0 "Backup path: " $DHCPDB.BackupPath
		}
		ElseIf($Text)
		{
			Line 1 "Database path`t: " $DHCPDB.FileName.SubString(0,($DHCPDB.FileName.LastIndexOf('\')))
			Line 1 "Backup path`t: " $DHCPDB.BackupPath
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Database path: " $DHCPDB.FileName.SubString(0,($DHCPDB.FileName.LastIndexOf('\')))
			WriteHTMLLine 0 0 "Backup path: " $DHCPDB.BackupPath
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving DHCP Server Database information"
		}
		ElseIf($Text)
		{
			Line 0 "Error retrieving DHCP Server Database information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving DHCP Server Database information"
		}
	}

	InsertBlankLine

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
				WriteWordLine 0 0 "DHCP server is authorized"
			}
			ElseIf($Text)
			{
				Line 1 "DHCP server is authorized"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 "DHCP server is authorized"
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "DHCP server is not authorized"
			}
			ElseIf($Text)
			{
				Line 1 "DHCP server is not authorized"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 "DHCP server is not authorized"
			}
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving DHCP Server setting information"
		}
		ElseIf($Text)
		{
			Line 0 "Error retrieving DHCP Server setting information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving DHCP Server setting information"
		}
	}
	InsertBlankLine
	[gc]::collect() 
}

Function ProcessIPv4Bindings
{
	Write-Verbose "$(Get-Date): `tGetting IPv4 bindings"
	$IPv4Bindings = Get-DHCPServerV4Binding -ComputerName $Script:DHCPServerName -EA 0 | Sort-Object IPAddress

	If($? -and $Null -ne $IPv4Bindings)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Connections and server bindings"
		}
		ElseIf($Text)
		{
			Line 0 "Connections and server bindings"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Connections and server bindings"
		}
		
		ForEach($IPv4Binding in $IPv4Bindings)
		{
			If($IPv4Binding.BindingState)
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 1 "Enabled " -NoNewLine
				}
				ElseIf($Text)
				{
					Line 1 "Enabled " -NoNewLine
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 0 1 "Enabled " 
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 1 "Disabled " -NoNewLine
				}
				ElseIf($Text)
				{
					Line 1 "Disabled " -NoNewLine
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 0 1 "Disabled " 
				}
			}
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "$($IPv4Binding.IPAddress) $($IPv4Binding.InterfaceAlias)"
			}
			ElseIf($Text)
			{
				Line 0 "$($IPv4Binding.IPAddress) $($IPv4Binding.InterfaceAlias)"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 1 "$($IPv4Binding.IPAddress) $($IPv4Binding.InterfaceAlias)"
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv4 server bindings"
		}
		ElseIf($Text)
		{
			Line 0 "Error retrieving IPv4 server bindings"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv4 server bindings"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no IPv4 server bindings"
		}
		ElseIf($Text)
		{
			Line 1 "There were no IPv4 server bindings"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "There were no IPv4 server bindings"
		}
	}
	$IPv4Bindings = $Null
	[gc]::collect() 

	InsertBlankLine
}

Function ProcessIPv6Bindings
{
	Write-Verbose "$(Get-Date): `tGetting IPv6 bindings"
	$IPv6Bindings = Get-DHCPServerV6Binding -ComputerName $Script:DHCPServerName -EA 0 | Sort-Object IPAddress

	If($? -and $Null -ne $IPv6Bindings)
	{
		WriteWordLine 0 0 "Connections and server bindings:"
		ForEach($IPv6Binding in $IPv6Bindings)
		{
			If($IPv6Binding.BindingState)
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 1 "Enabled " -NoNewLine
				}
				ElseIf($Text)
				{
					Line 1 "Enabled " -NoNewLine
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 0 1 "Enabled " 
				}
			}
			Else
			{
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 1 "Disabled " -NoNewLine
				}
				ElseIf($Text)
				{
					Line 1 "Disabled " -NoNewLine
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 0 1 "Disabled " 
				}
			}
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "$($IPv6Binding.IPAddress) $($IPv6Binding.InterfaceAlias)"
			}
			ElseIf($Text)
			{
				Line 0 "$($IPv6Binding.IPAddress) $($IPv6Binding.InterfaceAlias)"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 1 "$($IPv6Binding.IPAddress) $($IPv6Binding.InterfaceAlias)"
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv6 server bindings"
		}
		ElseIf($Text)
		{
			Line 0 "Error retrieving IPv6 server bindings"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv6 server bindings"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no IPv6 server bindings"
		}
		ElseIf($Text)
		{
			Line 1 "There were no IPv6 server bindings"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "There were no IPv6 server bindings"
		}
	}
	$IPv6Bindings = $Null
	[gc]::collect() 
}

Function ProcessIPv4Properties
{
	Write-Verbose "$(Get-Date): Getting IPv4 properties"
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 2 0 "IPv4"
		WriteWordLine 3 0 "Properties"
	}
	ElseIf($Text)
	{
		Line 0 ""
		Line 0 "IPv4"
		Line 0 "Properties"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "IPv4"
		WriteHTMLLine 3 0 "Properties"
	}

	Write-Verbose "$(Get-Date): `tGetting IPv4 general settings"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "General"
	}
	ElseIf($Text)
	{
		Line 1 "General"
	}
	ElseIf($HTML)
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
				WriteWordLine 0 1 "DHCP audit logging is enabled"
			}
			ElseIf($Text)
			{
				Line 2 "DHCP audit logging is enabled"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 1 "DHCP audit logging is enabled"
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 1 "DHCP audit logging is disabled"
			}
			ElseIf($Text)
			{
				Line 2 "DHCP audit logging is disabled"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 1 "DHCP audit logging is disabled"
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving audit log settings"
		}
		ElseIf($Text)
		{
			Line 0 "Error retrieving audit log settings"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving audit log settings"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "There were no audit log settings"
		}
		ElseIf($Text)
		{
			Line 0 "There were no audit log settings"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "There were no audit log settings"
		}
	}
	[gc]::collect() 

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
			WriteWordLine 0 1 "Show the BOOTP table folder is enabled"
		}
		ElseIf($Text)
		{
			Line 2 "Show the BOOTP table folder is enabled"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "Show the BOOTP table folder is enabled"
		}
	}

	#DNS settings
	Write-Verbose "$(Get-Date): `tGetting IPv4 DNS settings"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "DNS"
	}
	ElseIf($Text)
	{
		Line 1 "DNS"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "DNS"
	}

	$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $Script:DHCPServerName -EA 0
	If($? -and $Null -ne $DNSSettings)
	{
		GetDNSSettings $DNSSettings "A"
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv4 DNS Settings for DHCP server $Script:DHCPServerName"
		}
		ElseIf($Text)
		{
			Line 0 "Error retrieving IPv4 DNS Settings for DHCP server $Script:DHCPServerName"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv4 DNS Settings for DHCP server $Script:DHCPServerName"
		}
	}
	$DNSSettings = $Null
	[gc]::collect() 

	#now back to some server settings
	Write-Verbose "$(Get-Date): `tGetting IPv4 NAP settings"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Network Access Protection"
	}
	ElseIf($Text)
	{
		Line 1 "Network Access Protection"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Network Access Protection"
	}

	If($Script:GotServerSettings)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "Network Access Protection is " -NoNewLine
			If($Script:ServerSettings.NapEnabled)
			{
				WriteWordLine 0 0 "Enabled on all scopes"
			}
			Else
			{
				WriteWordLine 0 0 "Disabled on all scopes"
			}
			WriteWordLine 0 1 "DHCP server behavior when NPS is unreachable: " -NoNewLine
			Switch($Script:ServerSettings.NpsUnreachableAction)
			{
				"Full"			{WriteWordLine 0 0 "Full Access"}
				"Restricted"	{WriteWordLine 0 0 "Restricted Access"}
				"NoAccess"		{WriteWordLine 0 0 "Drop Client Packet"}
				Default			{WriteWordLine 0 0 "Unable to determine NPS unreachable action: $($Script:ServerSettings.NpsUnreachableAction)"}
			}
		}
		ElseIf($Text)
		{
			Line 2 "Network Access Protection is " -NoNewLine
			If($Script:ServerSettings.NapEnabled)
			{
				Line 0 "Enabled on all scopes"
			}
			Else
			{
				Line 0 "Disabled on all scopes"
			}
			Line 2 "DHCP server behavior when NPS is unreachable: " -NoNewLine
			Switch($Script:ServerSettings.NpsUnreachableAction)
			{
				"Full"			{Line 0 "Full Access"}
				"Restricted"	{Line 0 "Restricted Access"}
				"NoAccess"		{Line 0 "Drop Client Packet"}
				Default			{Line 0 "Unable to determine NPS unreachable action: $($Script:ServerSettings.NpsUnreachableAction)"}
			}
		}
		ElseIf($HTML)
		{
			If($Script:ServerSettings.NapEnabled)
			{
				WriteHTMLLine 0 1 "Network Access Protection is Enabled on all scopes"
			}
			Else
			{
				WriteHTMLLine 0 1 "Network Access Protection is Disabled on all scopes"
			}
			
			$tmp = ""
			Switch($Script:ServerSettings.NpsUnreachableAction)
			{
				"Full"			{$tmp = "Full Access"}
				"Restricted"	{$tmp = "Restricted Access"}
				"NoAccess"		{$tmp = "Drop Client Packet"}
				Default			{$tmp = "Unable to determine NPS unreachable action: $($Script:ServerSettings.NpsUnreachableAction)"}
			}
			WriteHTMLLine 0 1 "DHCP server behavior when NPS is unreachable: $tmp" 
		}
	}

	#filters
	Write-Verbose "$(Get-Date): `tGetting IPv4 filters"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Filters"
	}
	ElseIf($Text)
	{
		Line 1 "Filters"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Filters"
	}

	$MACFilters = Get-DHCPServerV4FilterList -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $MACFilters)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "Enable Allow list: " -NoNewLine
			If($MACFilters.Allow)
			{
				WriteWordLine 0 0 "Enabled"
			}
			Else
			{
				WriteWordLine 0 0 "Disabled"
			}
			WriteWordLine 0 1 "Enable Deny list: " -NoNewLine
			If($MACFilters.Deny)
			{
				WriteWordLine 0 0 "Enabled"
			}
			Else
			{
				WriteWordLine 0 0 "Disabled"
			}
		}
		ElseIf($Text)
		{
			Line 2 "Enable Allow list`t: " -NoNewLine
			If($MACFilters.Allow)
			{
				Line "Enabled"
			}
			Else
			{
				Line 0 "Disabled"
			}
			Line 2 "Enable Deny list`t: " -NoNewLine
			If($MACFilters.Deny)
			{
				Line 0 "Enabled"
			}
			Else
			{
				Line 0 "Disabled"
			}
		}
		ElseIf($HTML)
		{
			If($MACFilters.Allow)
			{
				WriteHTMLLine 0 1 "Enable Allow list: Enabled"
			}
			Else
			{
				WriteHTMLLine 0 1 "Enable Allow list: Disabled"
			}
			If($MACFilters.Deny)
			{
				WriteHTMLLine 0 1 "Enable Deny list: Enabled"
			}
			Else
			{
				WriteHTMLLine 0 1 "Enable Deny list: Disabled"
			}
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving MAC filters for DHCP server $Script:DHCPServerName"
		}
		ElseIf($Text)
		{
			Line 0 "Error retrieving MAC filters for DHCP server $Script:DHCPServerName"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving MAC filters for DHCP server $Script:DHCPServerName"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no MAC filters for DHCP server $Script:DHCPServerName"
		}
		ElseIf($Text)
		{
			Line 2 "There were no MAC filters for DHCP server $Script:DHCPServerName"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "There were no MAC filters for DHCP server $Script:DHCPServerName"
		}
	}
	$MACFilters = $Null
	[gc]::collect() 

	#failover
	Write-Verbose "$(Get-Date): `tGetting IPv4 Failover"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Failover"
	}
	ElseIf($Text)
	{
		Line 1 "Failover"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Failover"
	}

	$Failovers = Get-DHCPServerV4Failover -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $Failovers)
	{
		If($MSWord -or $PDF)
		{
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If($Failovers -is [array])
			{
				[int]$Rows = ($Failovers.Count * 8) - 1
				#subtract the very last row used for spacing
			}
			Else
			{
				[int]$Rows = 7
			}
			ValidateWordTableValues $Rows $Script:DHCPServerName "Function ProcessIPv4Properties Failovers"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = $wdLineStyleNone
			$table.Borders.OutsideLineStyle = $wdLineStyleNone
			[int]$xRow = 0
			ForEach($Failover in $Failovers)
			{
				Write-Verbose "$(Get-Date): `t`tProcessing failover $($Failover.Name)"
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Relationship name"
				$Table.Cell($xRow,2).Range.Text = $Failover.Name
						
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "State of the server"
				Switch($Failover.State)
				{
					"NoState" 
					{
						$Table.Cell($xRow,2).Range.Text = "No State"
					}
					"Normal" 
					{
						$Table.Cell($xRow,2).Range.Text = "Normal"
					}
					"Init" 
					{
						$Table.Cell($xRow,2).Range.Text = "Initializing"
					}
					"CommunicationInterrupted" 
					{
						$Table.Cell($xRow,2).Range.Text = "Communication Interrupted"
					}
					"PartnerDown" 
					{
						$Table.Cell($xRow,2).Range.Text = "Normal"
					}
					"PotentialConflict" 
					{
						$Table.Cell($xRow,2).Range.Text = "Potential Conflict"
					}
					"Startup" 
					{
						$Table.Cell($xRow,2).Range.Text = "Starting Up"
					}
					"ResolutionInterrupted" 
					{
						$Table.Cell($xRow,2).Range.Text = "Resolution Interrupted"
					}
					"ConflictDone" 
					{
						$Table.Cell($xRow,2).Range.Text = "Conflict Done"
					}
					"Recover" 
					{
						$Table.Cell($xRow,2).Range.Text = "Recover"
					}
					"RecoverWait" 
					{
						$Table.Cell($xRow,2).Range.Text = "Recover Wait"
					}
					"RecoverDone" 
					{
						$Table.Cell($xRow,2).Range.Text = "Recover Done"
					}
					Default 
					{
						$Table.Cell($xRow,2).Range.Text = "Unable to determine server state"
					}
				}
						
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Partner Server"
				$Table.Cell($xRow,2).Range.Text = $Failover.PartnerServer
						
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Mode"
				$Table.Cell($xRow,2).Range.Text = $Failover.Mode
						
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Message Authentication"
				If($Failover.EnableAuth)
				{
					$Table.Cell($xRow,2).Range.Text = "Enabled"
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = "Disabled"
				}
						
				If($Failover.Mode -eq "LoadBalance")
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "Local server"
					$Table.Cell($xRow,2).Range.Text = "$($Failover.LoadBalancePercent)%"
						
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "Partner Server"
					$tmp = (100 - $($Failover.LoadBalancePercent))
					$Table.Cell($xRow,2).Range.Text = "$($tmp)%"
					$tmp = $Null
				}
				Else
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "Role of this server"
					$Table.Cell($xRow,2).Range.Text = $Failover.ServerRole
						
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "Addresses reserved for standby server"
					$Table.Cell($xRow,2).Range.Text = "$($Failover.ReservePercent)%"
				}
						
				#skip a row for spacing
				$xRow++
			}
			$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
			$table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
		ElseIf($Text)
		{
			ForEach($Failover in $Failovers)
			{
				Write-Verbose "$(Get-Date): `t`tProcessing failover $($Failover.Name)"
				Line 2 "Relationship name: " $Failover.Name
						
				Line 2 "State of the server`t: " -NoNewLine
				Switch($Failover.State)
				{
					"NoState" 
					{
						Line 0 "No State"
					}
					"Normal" 
					{
						Line 0 "Normal"
					}
					"Init" 
					{
						Line 0 "Initializing"
					}
					"CommunicationInterrupted" 
					{
						Line 0 "Communication Interrupted"
					}
					"PartnerDown" 
					{
						Line 0 "Normal"
					}
					"PotentialConflict" 
					{
						Line 0 "Potential Conflict"
					}
					"Startup" 
					{
						Line 0 "Starting Up"
					}
					"ResolutionInterrupted" 
					{
						Line 0 "Resolution Interrupted"
					}
					"ConflictDone" 
					{
						Line 0 "Conflict Done"
					}
					"Recover" 
					{
						Line 0 "Recover"
					}
					"RecoverWait" 
					{
						Line 0 "Recover Wait"
					}
					"RecoverDone" 
					{
						Line 0 "Recover Done"
					}
					Default 
					{
						Line 0 "Unable to determine server state"
					}
				}
						
				Line 2 "Partner Server`t`t: " $Failover.PartnerServer
				Line 2 "Mode`t`t`t: " $Failover.Mode
				Line 2 "Message Authentication`t: " -NoNewLine
				If($Failover.EnableAuth)
				{
					Line 0 "Enabled"
				}
				Else
				{
					Line 0 "Disabled"
				}
						
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
				}
						
				#skip a row for spacing
				Line 0 ""
			}
		}
		ElseIf($HTML)
		{
			ForEach($Failover in $Failovers)
			{
				Write-Verbose "$(Get-Date): `t`tProcessing failover $($Failover.Name)"
				$rowdata = @()
				$columnHeaders = @("Relationship name",($htmlsilver -bor $htmlbold),$Failover.Name,$htmlwhite)
						
				$tmp = ""
				Switch($Failover.State)
				{
					"NoState"					{$tmp = "No State"}
					"Normal"					{$tmp = "Normal"}
					"Init"						{$tmp = "Initializing"}
					"CommunicationInterrupted"	{$tmp = "Communication Interrupted"}
					"PartnerDown"				{$tmp = "Normal"}
					"PotentialConflict"			{$tmp = "Potential Conflict"}
					"Startup"					{$tmp = "Starting Up"}
					"ResolutionInterrupted"		{$tmp = "Resolution Interrupted"}
					"ConflictDone"				{$tmp = "Conflict Done"}
					"Recover"					{$tmp = "Recover"}
					"RecoverWait"				{$tmp = "Recover Wait"}
					"RecoverDone"				{$tmp = "Recover Done"}
					Default						{$tmp = "Unable to determine server state"}
				}
				$rowdata += @(,('State of the server',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$rowdata += @(,('Partner Server',($htmlsilver -bor $htmlbold),$Failover.PartnerServer,$htmlwhite))
				$rowdata += @(,('Mode',($htmlsilver -bor $htmlbold),$Failover.Mode,$htmlwhite))
						
				$tmp = ""
				If($Failover.EnableAuth)
				{
					$tmp = "Enabled"
				}
				Else
				{
					$tmp = "Disabled"
				}
				$rowdata += @(,('Message Authentication',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
						
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
				$msg = ""
				$columnWidths = @("200","100")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "300"
				InsertBlankLine
			}
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There was no Failover configured for DHCP server $Script:DHCPServerName"
		}
		ElseIf($Text)
		{
			Line 2 "There was no Failover configured for DHCP server $Script:DHCPServerName"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "There was no Failover configured for DHCP server $Script:DHCPServerName"
		}
	}
	$Failovers = $Null
	[gc]::collect() 

	#Advanced
	Write-Verbose "$(Get-Date): `tGetting IPv4 advanced settings"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Advanced"
	}
	ElseIf($Text)
	{
		Line 1 "Advanced"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Advanced"
	}

	If($Script:GotServerSettings)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "Conflict detection attempts`t: " $Script:ServerSettings.ConflictDetectionAttempts
		}
		ElseIf($Text)
		{
			Line 2 "Conflict detection attempts`t: " $Script:ServerSettings.ConflictDetectionAttempts
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "Conflict detection attempts: " $Script:ServerSettings.ConflictDetectionAttempts
		}
	}

	If($Script:GotAuditSettings)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "Audit log file path`t`t: " $Script:AuditSettings.Path
		}
		ElseIf($Text)
		{
			Line 2 "Audit log file path`t`t: " $Script:AuditSettings.Path
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "Audit log file path: " $Script:AuditSettings.Path
		}
	}

	#added 18-Jan-2016
	#get dns update credentials
	Write-Verbose "$(Get-Date): `tGetting DNS dynamic update registration credentials"
	$DNSUpdateSettings = Get-DhcpServerDnsCredential -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $DNSUpdateSettings)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "DNS dynamic update registration credentials: "
			WriteWordLine 0 2 "User name`t: " $DNSUpdateSettings.UserName
			WriteWordLine 0 2 "Domain`t`t: " $DNSUpdateSettings.DomainName
		}
		ElseIf($Text)
		{
			Line 2 "DNS dynamic update registration credentials: "
			Line 3 "User name`t: " $DNSUpdateSettings.UserName
			Line 3 "Domain`t`t: " $DNSUpdateSettings.DomainName
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "DNS dynamic update registration credentials: "
			WriteHTMLLine 0 2 "User name: " $DNSUpdateSettings.UserName
			WriteHTMLLine 0 2 "Domain: " $DNSUpdateSettings.DomainName
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving DNS Update Credentials for DHCP server $Script:DHCPServerName"
		}
		ElseIf($Text)
		{
			Line 0 "Error retrieving DNS Update Credentials for DHCP server $Script:DHCPServerName"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving DNS Update Credentials for DHCP server $Script:DHCPServerName"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no DNS Update Credentials for DHCP server $Script:DHCPServerName"
		}
		ElseIf($Text)
		{
			Line 2 "There were no DNS Update Credentials for DHCP server $Script:DHCPServerName"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "There were no DNS Update Credentials for DHCP server $Script:DHCPServerName"
		}
	}
	$DNSUpdateSettings = $Null
	[gc]::collect() 
}

Function ProcessIPv4Statistics
{
	Write-Verbose "$(Get-Date): Getting IPv4 Statistics"
	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Statistics"
	}
	ElseIf($Text)
	{
		Line 1 "Statistics"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 4 0 "Statistics"
	}
	[gc]::collect() 

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
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			[int]$Rows = 16
			Write-Verbose "$(Get-Date): `tAdd IPv4 statistics table to doc"
			ValidateWordTableValues $Rows $Script:DHCPServerName "Function ProcessIPv4Statistics Statistics"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = $wdLineStyleSingle
			$table.Borders.OutsideLineStyle = $wdLineStyleSingle
			$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell(1,1).Range.Font.Bold = $True
			$Table.Cell(1,1).Range.Text = "Description"
			$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell(1,2).Range.Font.Bold = $True
			$Table.Cell(1,2).Range.Text = "Details"

			$Table.Cell(2,1).Range.Text = "Start Time"
			$Table.Cell(2,2).Range.Text = $Statistics.ServerStartTime.ToString()
			$Table.Cell(3,1).Range.Text = "Up Time"
			$Table.Cell(3,2).Range.Text = $Str
			$Table.Cell(4,1).Range.Text = "Discovers"
			$Table.Cell(4,2).Range.Text = $Statistics.Discovers.ToString()
			$Table.Cell(5,1).Range.Text = "Offers"
			$Table.Cell(5,2).Range.Text = $Statistics.Offers.ToString()
			$Table.Cell(6,1).Range.Text = "Delayed Offers"
			$Table.Cell(6,2).Range.Text = $Statistics.DelayedOffers.ToString()
			$Table.Cell(7,1).Range.Text = "Requests"
			$Table.Cell(7,2).Range.Text = $Statistics.Requests.ToString()
			$Table.Cell(8,1).Range.Text = "Acks"
			$Table.Cell(8,2).Range.Text = $Statistics.Acks.ToString()
			$Table.Cell(9,1).Range.Text = "Nacks"
			$Table.Cell(9,2).Range.Text = $Statistics.Naks.ToString()
			$Table.Cell(10,1).Range.Text = "Declines"
			$Table.Cell(10,2).Range.Text = $Statistics.Declines.ToString()
			$Table.Cell(11,1).Range.Text = "Releases"
			$Table.Cell(11,2).Range.Text = $Statistics.Releases.ToString()
			$Table.Cell(12,1).Range.Text = "Total Scopes"
			$Table.Cell(12,2).Range.Text = $Statistics.TotalScopes.ToString()
			$Table.Cell(13,1).Range.Text = "Scopes with delay configured"
			$Table.Cell(13,2).Range.Text = $Statistics.ScopesWithDelayConfigured.ToString()
			$Table.Cell(14,1).Range.Text = "Total Addresses"
			$Table.Cell(14,2).Range.Text = "{0:N0}" -f $Statistics.TotalAddresses.ToString()
			$Table.Cell(15,1).Range.Text = "In Use"
			$Table.Cell(15,2).Range.Text = "$($Statistics.AddressesInUse) ($($InUsePercent))%"
			$Table.Cell(16,1).Range.Text = "Available"
			$Table.Cell(16,2).Range.Text = "{0:N0}" -f "$($Statistics.AddressesAvailable) ($($AvailablePercent))%"

			$table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
			Write-Verbose "$(Get-Date): `tFinished IPv4 statistics table"
			Write-Verbose "$(Get-Date): "
		}
		ElseIf($Text)
		{
			Line 2 "Description" -NoNewLine
			Line 3 "Details"

			Line 2 "Start Time:" -NoNewLine
			Line 3 $Statistics.ServerStartTime
			Line 2 "Up Time:" -NoNewLine
			Line 3 $Str
			Line 2 "Discovers:" -NoNewLine
			Line 3 $Statistics.Discovers
			Line 2 "Offers:" -NoNewLine
			Line 4 $Statistics.Offers
			Line 2 "Delayed Offers:" -NoNewLine
			Line 3 $Statistics.DelayedOffers
			Line 2 "Requests:" -NoNewLine
			Line 3 $Statistics.Requests
			Line 2 "Acks:" -NoNewLine
			Line 4 $Statistics.Acks
			Line 2 "Nacks:" -NoNewLine
			Line 4 $Statistics.Naks
			Line 2 "Declines:" -NoNewLine
			Line 3 $Statistics.Declines
			Line 2 "Releases:" -NoNewLine
			Line 3 $Statistics.Releases
			Line 2 "Total Scopes:" -NoNewLine
			Line 3 $Statistics.TotalScopes
			Line 2 "Scopes w/delay configured:" -NoNewLine
			Line 1 $Statistics.ScopesWithDelayConfigured
			Line 2 "Total Addresses:" -NoNewLine
			$tmp = "{0:N0}" -f $Statistics.TotalAddresses
			Line 2 $tmp
			Line 2 "In Use:" -NoNewLine
			Line 4 "$($Statistics.AddressesInUse) ($($InUsePercent))%"
			Line 2 "Available:" -NoNewLine
			$tmp = "{0:N0}" -f $Statistics.AddressesAvailable
			Line 3 "$($tmp) ($($AvailablePercent))%"
		}
		ElseIf($HTML)
		{
			Write-Verbose "$(Get-Date): `tAdd IPv4 statistics table to doc"
			$rowdata = @()

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
			$rowdata += @(,("In Use",$htmlwhite,"$($Statistics.AddressesInUse) ($($InUsePercent))%",$htmlwhite))
			$tmp = "{0:N0}" -f "$($Statistics.AddressesAvailable) ($($AvailablePercent))%"
			$rowdata += @(,("Available",$htmlwhite,$tmp,$htmlwhite))

			$columnHeaders = @('Description',($htmlsilver -bor $htmlbold),'Details',($htmlsilver -bor $htmlbold))
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			InsertBlankLine

			Write-Verbose "$(Get-Date): `tFinished IPv4 statistics table"
			Write-Verbose "$(Get-Date): "
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv4 statistics"
		}
		ElseIf($Text)
		{
			Line 0 "Error retrieving IPv4 statistics"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv4 statistics"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "There were no IPv4 statistics"
		}
		ElseIf($Text)
		{
			Line 0 "There were no IPv4 statistics"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "There were no IPv4 statistics"
		}
	}
	$Statistics = $Null
	[gc]::collect() 
}

Function ProcessIPv4Superscopes
{
	Write-Verbose "$(Get-Date): Getting IPv4 Superscopes"
	$IPv4Superscopes = Get-DHCPServerV4Superscope -ComputerName $Script:DHCPServerName -EA 0

	If($? -and $Null -ne $IPv4Superscopes)
	{
		ForEach($IPv4Superscope in $IPv4Superscopes)
		{
			If(![string]::IsNullOrEmpty($IPv4Superscope.SuperscopeName))
			{
				If($MSWord -or $PDF)
				{
					#put each superscope on a new page
					$selection.InsertNewPage()
					Write-Verbose "$(Get-Date): `tGetting IPv4 superscope data for scope $($IPv4Superscope.SuperscopeName)"
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
						Write-Verbose "$(Get-Date): Processing scope id $($xScopeId) for Superscope $($IPv4Superscope.SuperscopeName)"
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
				ElseIf($Text)
				{
					Write-Verbose "$(Get-Date): `tGetting IPv4 superscope data for scope $($IPv4Superscope.SuperscopeName)"
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
						Write-Verbose "$(Get-Date): Processing scope id $($xScopeId) for Superscope $($IPv4Superscope.SuperscopeName)"
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
				ElseIf($HTML)
				{
					Write-Verbose "$(Get-Date): `tGetting IPv4 superscope data for scope $($IPv4Superscope.SuperscopeName)"
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
						Write-Verbose "$(Get-Date): Processing scope id $($xScopeId) for Superscope $($IPv4Superscope.SuperscopeName)"
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
		ElseIf($Text)
		{
			Line 0 "Error retrieving IPv4 Superscopes"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv4 Superscopes"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no IPv4 Superscopes"
		}
		ElseIf($Text)
		{
			Line 2 "There were no IPv4 Superscopes"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "There were no IPv4 Superscopes"
		}
	}
	$IPv4Superscopes = $Null
	[gc]::collect() 
}

Function GetIPv4ScopeData
{
	Param([object]$IPv4Scope, [int]$xStartLevel)
	
	Write-Verbose "$(Get-Date): `tBuild array of Allow/Deny filters"
	$Filters = Get-DHCPServerV4Filter -ComputerName $Script:DHCPServerName -EA 0

	If($MSWord -or $PDF)
	{
		GetIPv4ScopeData_WordPDF $xStartLevel $Filters
	}
	ElseIf($Text)
	{
		GetIPv4ScopeData_Text $xStartLevel $Filters
	}
	ElseIf($HTML)
	{
		GetIPv4ScopeData_HTML $xStartLevel $Filters
	}
}

Function GetIPv4ScopeData_WordPDF
{
	Param([int] $xStartLevel, [object]$filters)
	#put each scope on a new page
	$selection.InsertNewPage()
	Write-Verbose "$(Get-Date): `tGetting IPv4 scope data for scope $($IPv4Scope.Name)"
	WriteWordLine $xStartLevel 0 "Scope [$($IPv4Scope.ScopeId)] $($IPv4Scope.Name)"
	WriteWordLine ($xStartLevel + 1) 0 "Address Pool"
	$TableRange = $doc.Application.Selection.Range
	[int]$Columns = 2
	[int]$Rows = 5
	ValidateWordTableValues $Rows $Script:DHCPServerName "Function GetIPv4ScopeData_WordPDF Scope [$($IPv4Scope.ScopeId)] $($IPv4Scope.Name)"
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = $myHash.Word_TableGrid
	$table.Borders.InsideLineStyle = $wdLineStyleNone
	$table.Borders.OutsideLineStyle = $wdLineStyleNone
	$Table.Cell(1,1).Range.Text = "Start IP Address"
	$Table.Cell(1,2).Range.Text = $IPv4Scope.StartRange.ToString()
	$Table.Cell(2,1).Range.Text = "End IP Address"
	$Table.Cell(2,2).Range.Text = $IPv4Scope.EndRange.ToString()
	$Table.Cell(3,1).Range.Text = "Subnet Mask"
	$Table.Cell(3,2).Range.Text = $IPv4Scope.SubnetMask.ToString()
	$Table.Cell(4,1).Range.Text = "Lease duration"
	If($IPv4Scope.LeaseDuration -eq "00:00:00")
	{
		$Table.Cell(4,2).Range.Text = "Unlimited"
	}
	Else
	{
		$Str = [string]::format("{0} days, {1} hours, {2} minutes", `
			$IPv4Scope.LeaseDuration.Days, `
			$IPv4Scope.LeaseDuration.Hours, `
			$IPv4Scope.LeaseDuration.Minutes)

		$Table.Cell(4,2).Range.Text = $Str
	}
	$Table.Cell(5,1).Range.Text = "Description"
	$Table.Cell(5,2).Range.Text = $IPv4Scope.Description

	$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
	$table.AutoFitBehavior($wdAutoFitContent)

	#return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
	$TableRange = $Null
	$Table = $Null

	If($IncludeLeases)
	{
		Write-Verbose "$(Get-Date):	`t`tGetting leases"
		
		WriteWordLine ($xStartLevel + 1) 0 "Address Leases"
		$Leases = Get-DHCPServerV4Lease -ComputerName $Script:DHCPServerName -ScopeId  $IPv4Scope.ScopeId -EA 0 | Sort-Object IPAddress
		If($? -and $Null -ne $Leases)
		{
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If($Leases -is [array])
			{
				[int]$Rows = ($Leases.Count * 11) - 1
				#subtract the very last row used for spacing
			}
			Else
			{
				[int]$Rows = 10
			}
			ValidateWordTableValues $Rows $Script:DHCPServerName "Function GetIPv4ScopeData_WordPDF Getting leases"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = $wdLineStyleNone
			$table.Borders.OutsideLineStyle = $wdLineStyleNone
			[int]$xRow = 0
			ForEach($Lease in $Leases)
			{
				Write-Verbose "$(Get-Date):	`t`t`tProcessing lease $($Lease.IPAddress)"
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

				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Client IP address"
				$Table.Cell($xRow,2).Range.Text = $Lease.IPAddress.ToString()
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Name"
				$Table.Cell($xRow,2).Range.Text = $Lease.HostName
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Lease Expiration"
				If([string]::IsNullOrEmpty($Lease.LeaseExpiryTime))
				{
					If($Lease.AddressState -eq "ActiveReservation")
					{
						$Table.Cell($xRow,2).Range.Text = "Reservation (active)"
					}
					Else
					{
						$Table.Cell($xRow,2).Range.Text = "Reservation (inactive)"
					}
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $LeaseStr
				}
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Type"
				$Table.Cell($xRow,2).Range.Text = $Lease.ClientType
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Unique ID"
				$Table.Cell($xRow,2).Range.Text = $Lease.ClientID
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Description"
				$Table.Cell($xRow,2).Range.Text = $Lease.Description

				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Network Access Protection"
				$Table.Cell($xRow,2).Range.Text = $Lease.NapStatus

				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Probation Expiration"
				
				If([string]::IsNullOrEmpty($Lease.ProbationEnds))
				{
					$Table.Cell($xRow,2).Range.Text = "N/A"
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $ProbationStr
				}
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Filter"
				
				#$Filters | ForEach-Object { $index = $Null }{ If( $_.MacAddress -eq $Lease.ClientID ) { $index = $_ } }
				
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
					$Table.Cell($xRow,2).Range.Text = $Index.List
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = "<None>"
				}

				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Policy"
				
				If([string]::IsNullOrEmpty($Lease.PolicyName))
				{
					$Table.Cell($xRow,2).Range.Text = "<None>"
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $Lease.PolicyName
				}
				
				#skip a row for spacing
				$xRow++
			}
			$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
			$table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
		ElseIf(!$?)
		{
			WriteWordLine 0 0 "Error retrieving leases for scope $IPv4Scope.ScopeId"
		}
		Else
		{
			WriteWordLine 0 0 "<None>"
		}
		$Leases = $Null
	}

	Write-Verbose "$(Get-Date):	`t`tGetting exclusions"
	WriteWordLine ($xStartLevel + 1) 0 "Exclusions"
	$Exclusions = Get-DHCPServerV4ExclusionRange -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object StartRange
	If($? -and $Null -ne $Exclusions)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($Exclusions -is [array])
		{
			[int]$Rows = $Exclusions.Count + 1
		}
		Else
		{
			[int]$Rows = 2
		}
		ValidateWordTableValues $Rows $Script:DHCPServerName "Function GetIPv4ScopeData_WordPDF Getting exclusions"
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = $wdLineStyleNone
		$table.Borders.OutsideLineStyle = $wdLineStyleNone
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "Start IP Address"
		$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,2).Range.Font.Bold = $True
		$Table.Cell(1,2).Range.Text = "End IP Address"
		[int]$xRow = 1
		ForEach($Exclusion in $Exclusions)
		{
			$xRow++
			$Table.Cell($xRow,1).Range.Text = $Exclusion.StartRange.ToString()
			$Table.Cell($xRow,2).Range.Text = $Exclusion.EndRange.ToString() 
		}
		$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
		$table.AutoFitBehavior($wdAutoFitContent)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving exclusions for scope $IPv4Scope.ScopeId"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$Exclusions = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting reservations"
	WriteWordLine ($xStartLevel + 1) 0 "Reservations"
	$Reservations = Get-DHCPServerV4Reservation -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object IPAddress
	If($? -and $Null -ne $Reservations)
	{
		ForEach($Reservation in $Reservations)
		{
			Write-Verbose "$(Get-Date):	`t`t`tProcessing reservation $($Reservation.Name)"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If([string]::IsNullOrEmpty($Reservation.Description))
			{
				[int]$Rows = 4
			}
			Else
			{
				[int]$Rows = 5
			}
			ValidateWordTableValues $Rows $Script:DHCPServerName "Function GetIPv4ScopeData_WordPDF Getting reservations"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = $wdLineStyleNone
			$table.Borders.OutsideLineStyle = $wdLineStyleNone
			$Table.Cell(1,1).Range.Text = "Reservation name"
			$Table.Cell(1,2).Range.Text = $Reservation.Name
			$Table.Cell(2,1).Range.Text = "IP address"
			$Table.Cell(2,2).Range.Text = $Reservation.IPAddress.ToString()
			$Table.Cell(3,1).Range.Text = "MAC address"
			$Table.Cell(3,2).Range.Text = $Reservation.ClientId
			$Table.Cell(4,1).Range.Text = "Supported types"
			$Table.Cell(4,2).Range.Text = $Reservation.Type
			If(![string]::IsNullOrEmpty($Reservation.Description))
			{
				$Table.Cell(5,1).Range.Text = "Description"
				$Table.Cell(5,2).Range.Text = $Reservation.Description
			}
			$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
			$table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null

			Write-Verbose "$(Get-Date):	`t`t`t`tGetting DNS settings"
			$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $Script:DHCPServerName -IPAddress $Reservation.IPAddress -EA 0
			If($? -and $Null -ne $DNSSettings)
			{
				GetDNSSettings $DNSSettings "A"
			}
			Else
			{
				WriteWordLine 0 0 "Error retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
			}
			$DNSSettings = $Null
			WriteWordLine 0 0 ""
		}
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving reservations for scope $IPv4Scope.ScopeId"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$Reservations = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting scope options"
	WriteWordLine ($xStartLevel + 1) 0 "Scope Options"
	$ScopeOptions = Get-DHCPServerV4OptionValue -All -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object OptionId

	If($? -and $Null -ne $ScopeOptions)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($ScopeOptions -is [array])
		{
			[int]$Rows = (($ScopeOptions.Count * 5) - 5) - 1
			#subtract option 51
			#subtract the very last row used for spacing
		}
		Else
		{
			If($ScopeOptions.OptionId -eq 51)
			{
				#if the only scope option is 51, ignore it
				[int]$Rows = 0
				WriteWordLine 0 1 "<None>"
			}
			Else
			{
				#id the only scope option is not 51, add rows for it
				[int]$Rows = 4
			}
		}
		
		If($Rows -gt 0)
		{
			ValidateWordTableValues $Rows $Script:DHCPServerName "Function GetIPv4ScopeData_WordPDF ScopeOptions"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = $wdLineStyleNone
			$table.Borders.OutsideLineStyle = $wdLineStyleNone
			[int]$xRow = 0
			ForEach($ScopeOption in $ScopeOptions)
			{
				If($ScopeOption.OptionId -eq 51 -or $ScopeOption.OptionId -eq 81)
				{
					#ignore these two option IDs
					https://carlwebster.com/the-mysterious-microsoft-dhcp-option-id-81/
					https://jimswirelessworld.wordpress.com/2019/01/03/you-should-care-about-dhcp-option-51/
				}
				Else
				{
					Write-Verbose "$(Get-Date):	`t`t`tProcessing option name $($ScopeOption.Name)"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "Option Name"
					$Table.Cell($xRow,2).Range.Text = "$($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)" 
					
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "Vendor"
					If([string]::IsNullOrEmpty($ScopeOption.VendorClass))
					{
						$Table.Cell($xRow,2).Range.Text = "Standard" 
					}
					Else
					{
						$Table.Cell($xRow,2).Range.Text = $ScopeOption.VendorClass 
					}
					
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "Value"
					$Table.Cell($xRow,2).Range.Text = "$($ScopeOption.Value)" 
					
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "Policy Name"
					
					If([string]::IsNullOrEmpty($ScopeOption.PolicyName))
					{
						$Table.Cell($xRow,2).Range.Text = "<None>"
					}
					Else
					{
						$Table.Cell($xRow,2).Range.Text = $ScopeOption.PolicyName
					}
				
					#for spacing
					$xRow++
				}
			}
			$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
			$table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving scope options for $IPv4Scope.ScopeId"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$ScopeOptions = $Null
	
	Write-Verbose "$(Get-Date):	`t`tGetting policies"
	WriteWordLine ($xStartLevel + 1) 0 "Policies"
	$ScopePolicies = Get-DHCPServerV4Policy -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object ProcessingOrder

	If($? -and $Null -ne $ScopePolicies)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($ScopePolicies -is [array])
		{
			[int]$Rows = $ScopePolicies.Count * 6
		}
		Else
		{
			[int]$Rows = 6
		}
		ValidateWordTableValues $Rows $Script:DHCPServerName "Function GetIPv4ScopeData_WordPDF Getting policies"
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = $wdLineStyleNone
		$table.Borders.OutsideLineStyle = $wdLineStyleNone
		[int]$xRow = 0
		ForEach($Policy in $ScopePolicies)
		{
			Write-Verbose "$(Get-Date):	`t`t`tProcessing policy name $($Policy.Name)"
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Policy Name"
			$Table.Cell($xRow,2).Range.Text = $Policy.Name
			
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Description"
			$Table.Cell($xRow,2).Range.Text = $Policy.Description

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Processing Order"
			$Table.Cell($xRow,2).Range.Text = $Policy.ProcessingOrder.ToString()

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Level"
			$Table.Cell($xRow,2).Range.Text = "Scope"

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Address Range"
			
			$IPRange = Get-DHCPServerV4PolicyIPRange -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -Name $Policy.Name -EA 0

			If($? -and $Null -ne $IPRange)
			{
				$Table.Cell($xRow,2).Range.Text = "$($IPRange.StartRange) - $($IPRange.EndRange)"
			}
			Else
			{
				$Table.Cell($xRow,2).Range.Text = "<None>"
			}

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "State"
			If($Policy.Enabled)
			{
				$Table.Cell($xRow,2).Range.Text = "Enabled"
			}
			Else
			{
				$Table.Cell($xRow,2).Range.Text = "Disabled"
			}
		}
		$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
		$table.AutoFitBehavior($wdAutoFitContent)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving scope policies"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$ScopePolicies = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting DNS"
	WriteWordLine ($xStartLevel + 1) 0 "DNS"
	$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0
	If($? -and $Null -ne $DNSSettings)
	{
		GetDNSSettings $DNSSettings "A"
	}
	Else
	{
		WriteWordLine 0 0 "Error retrieving DNS Settings for scope $($IPv4Scope.ScopeId)"
	}
	$DNSSettings = $Null
	
	#next tab is Network Access Protection but I can't find anything that gives me that info
	
	#failover
	Write-Verbose "$(Get-Date):	`t`tGetting failover"
	WriteWordLine ($xStartLevel + 1) 0 "Failover"
	
	$Failovers = $Null
	$Failovers = Get-DHCPServerV4Failover -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0

	If($? -and $Null -ne $Failovers)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($Failovers -is [array])
		{
			[int]$Rows = ($Failovers.Count * 10) - 1
			#subtract the very last row used for spacing
		}
		Else
		{
			[int]$Rows = 9
		}
		ValidateWordTableValues $Rows $Script:DHCPServerName "Function GetIPv4ScopeData_WordPDF Getting failover"
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = $wdLineStyleNone
		$table.Borders.OutsideLineStyle = $wdLineStyleNone
		[int]$xRow = 0
		ForEach($Failover in $Failovers)
		{
			Write-Verbose "$(Get-Date):	`t`tProcessing failover $($Failover.Name)"
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Relationship name"
			$Table.Cell($xRow,2).Range.Text = $Failover.Name
					
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Partner Server"
			$Table.Cell($xRow,2).Range.Text = $Failover.PartnerServer
					
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Mode"
			$Table.Cell($xRow,2).Range.Text = $Failover.Mode
					
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

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Max Client Lead Time"
			$Table.Cell($xRow,2).Range.Text = $MaxLeadStr
					
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

			$xRow++
			$Table.Cell($xRow,1).Range.Text = "State Switchover Interval"
			$Table.Cell($xRow,2).Range.Text = $SwitchStr
					
			Switch($Failover.State)
			{
				"NoState" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "No State"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "No State"
				}
				"Normal" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Normal"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Normal"
				}
				"Init" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Initializing"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Initializing"
				}
				"CommunicationInterrupted" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Communication Interrupted"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Communication Interrupted"
				}
				"PartnerDown" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Normal"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Down"
				}
				"PotentialConflict" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Potential Conflict"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Potential Conflict"
				}
				"Startup" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Starting Up"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Starting Up"
				}
				"ResolutionInterrupted" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Resolution Interrupted"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Resolution Interrupted"
				}
				"ConflictDone" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Conflict Done"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Conflict Done"
				}
				"Recover" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Recover"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Recover"
				}
				"RecoverWait" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Recover Wait"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Recover Wait"
				}
				"RecoverDone" 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Recover Done"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Recover Done"
				}
				Default 
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of this Server"
					$Table.Cell($xRow,2).Range.Text = "Unable to determine server state"
					$xRow++
					$Table.Cell($xRow,1).Range.Text = "State of Partner Server"
					$Table.Cell($xRow,2).Range.Text = "Unable to determine server state"
				}
			}
					
			If($Failover.Mode -eq "LoadBalance")
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Local server"
				$Table.Cell($xRow,2).Range.Text = "$($Failover.LoadBalancePercent)%"
					
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Partner Server"
				$tmp = (100 - $($Failover.LoadBalancePercent))
				$Table.Cell($xRow,2).Range.Text = "$($tmp)%"
				$tmp = $Null
			}
			Else
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Role of this server"
				$Table.Cell($xRow,2).Range.Text = $Failover.ServerRole
					
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Addresses reserved for standby server"
				$Table.Cell($xRow,2).Range.Text = "$($Failover.ReservePercent)%"
			}
					
			#skip a row for spacing
			$xRow++
		}
		$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
		$table.AutoFitBehavior($wdAutoFitContent)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$Failovers = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting Scope statistics"
	WriteWordLine ($xStartLevel + 1) 0 "Statistics"

	$Statistics = Get-DHCPServerV4ScopeStatistics -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0

	If($? -and $Null -ne $Statistics)
	{
		GetShortStatistics $Statistics
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving scope statistics"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$Statistics = $Null
}

Function GetIPv4ScopeData_HTML
{
	Param([int] $xStartLevel, [object]$filters)

	Write-Verbose "$(Get-Date): `tGetting IPv4 scope data for scope $($IPv4Scope.Name)"
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
	InsertBlankLine

	If($IncludeLeases)
	{
		Write-Verbose "$(Get-Date):	`t`tGetting leases"
		
		WriteHTMLLine ($xStartLevel + 1) 0 "Address Leases"
		$Leases = Get-DHCPServerV4Lease -ComputerName $Script:DHCPServerName -ScopeId  $IPv4Scope.ScopeId -EA 0 | Sort-Object IPAddress
		If($? -and $Null -ne $Leases)
		{
			ForEach($Lease in $Leases)
			{
				Write-Verbose "$(Get-Date):	`t`t`tProcessing lease $($Lease.IPAddress)"
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
				$rowdata += @(,('Network Access Protection',($htmlsilver -bor $htmlbold),$Lease.NapStatus,$htmlwhite))

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
				InsertBlankLine
			}
		}
		ElseIf(!$?)
		{
			WriteHTMLLine 0 0 "Error retrieving leases for scope $IPv4Scope.ScopeId"
		}
		Else
		{
			WriteHTMLLine 0 0 "None"
		}
		$Leases = $Null
	}

	Write-Verbose "$(Get-Date):	`t`tGetting exclusions"
	WriteHTMLLine ($xStartLevel + 1) 0 "Exclusions"
	$Exclusions = Get-DHCPServerV4ExclusionRange -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object StartRange
	If($? -and $Null -ne $Exclusions)
	{
		$rowdata = @()
		ForEach($Exclusion in $Exclusions)
		{
			$rowdata += @(,($Exclusion.StartRange.ToString(),$htmlwhite,
							$Exclusion.EndRange.ToString(),$htmlwhite))
		}

		$columnHeaders = @('Start IP Address',($htmlsilver -bor $htmlbold),'End IP Address',($htmlsilver -bor $htmlbold))
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		InsertBlankLine
	}
	ElseIf(!$?)
	{
		WriteHTMLLine 0 0 "Error retrieving exclusions for scope $IPv4Scope.ScopeId"
	}
	Else
	{
		WriteHTMLLine 0 1 "None"
	}
	$Exclusions = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting reservations"
	WriteHTMLLine ($xStartLevel + 1) 0 "Reservations"
	$Reservations = Get-DHCPServerV4Reservation -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object IPAddress
	If($? -and $Null -ne $Reservations)
	{
		ForEach($Reservation in $Reservations)
		{
			Write-Verbose "$(Get-Date):	`t`t`tProcessing reservation $($Reservation.Name)"
			$rowdata = @()
			$columnHeaders = @("Reservation name",($htmlsilver -bor $htmlbold),$Reservation.Name,$htmlwhite)
			$rowdata += @(,('IP address',($htmlsilver -bor $htmlbold),$Reservation.IPAddress.ToString(),$htmlwhite))
			$rowdata += @(,('MAC address',($htmlsilver -bor $htmlbold),$Reservation.ClientId,$htmlwhite))
			$rowdata += @(,('Supported types',($htmlsilver -bor $htmlbold),$Reservation.Type,$htmlwhite))
			If(![string]::IsNullOrEmpty($Reservation.Description))
			{
				$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Reservation.Description,$htmlwhite))
			}
			$msg = ""
			$columnWidths = @("150","200")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
			InsertBlankLine

			Write-Verbose "$(Get-Date):	`t`t`t`tGetting DNS settings"
			$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $Script:DHCPServerName -IPAddress $Reservation.IPAddress -EA 0
			If($? -and $Null -ne $DNSSettings)
			{
				GetDNSSettings $DNSSettings "A"
			}
			Else
			{
				WriteHTMLLine 0 0 "Error retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
			}
			$DNSSettings = $Null
			InsertBlankLine
		}
	}
	ElseIf(!$?)
	{
		WriteHTMLLine 0 0 "Error retrieving reservations for scope $IPv4Scope.ScopeId"
	}
	Else
	{
		WriteHTMLLine 0 1 "None"
	}
	$Reservations = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting scope options"
	WriteHTMLLine ($xStartLevel + 1) 0 "Scope Options"
	$ScopeOptions = Get-DHCPServerV4OptionValue -All -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object OptionId

	If($? -and $Null -ne $ScopeOptions)
	{
		If($ScopeOptions -is [array])
		{
			#do nothing for html
		}
		Else
		{
			If($ScopeOptions.OptionId -eq 51)
			{
				#if the only scope option is 51, ignore it
				WriteHTMLLine 0 1 "None"
			}
			Else
			{
				#do nothing for html
			}
		}
		
		ForEach($ScopeOption in $ScopeOptions)
		{
			If($ScopeOption.OptionId -eq 51 -or $ScopeOption.OptionId -eq 81)
			{
				#ignore these two option IDs
				https://carlwebster.com/the-mysterious-microsoft-dhcp-option-id-81/
				https://jimswirelessworld.wordpress.com/2019/01/03/you-should-care-about-dhcp-option-51/
			}
			Else
			{
				Write-Verbose "$(Get-Date):	`t`t`tProcessing option name $($ScopeOption.Name)"
				$rowdata = @()
				$columnHeaders = @("Option Name",($htmlsilver -bor $htmlbold),"$($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)",$htmlwhite)
				
				$tmp = ""
				If([string]::IsNullOrEmpty($ScopeOption.VendorClass))
				{
					$tmp = "Standard" 
				}
				Else
				{
					$tmp = $ScopeOption.VendorClass 
				}
				$rowdata += @(,('Vendor',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$rowdata += @(,('Value',($htmlsilver -bor $htmlbold),"$($ScopeOption.Value)",$htmlwhite))
				
				$tmp = ""
				If([string]::IsNullOrEmpty($ScopeOption.PolicyName))
				{
					$tmp = "None"
				}
				Else
				{
					$tmp = $ScopeOption.PolicyName
				}
				$rowdata += @(,('Policy Name',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			
				$msg = ""
				$columnWidths = @("150","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
				InsertBlankLine
			}
		}
	}
	ElseIf(!$?)
	{
		WriteHTMLLine 0 0 "Error retrieving scope options for $IPv4Scope.ScopeId"
	}
	Else
	{
		WriteHTMLLine 0 1 "None"
	}
	$ScopeOptions = $Null
	
	Write-Verbose "$(Get-Date):	`t`tGetting policies"
	WriteHTMLLine ($xStartLevel + 1) 0 "Policies"
	$ScopePolicies = Get-DHCPServerV4Policy -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object ProcessingOrder

	If($? -and $Null -ne $ScopePolicies)
	{
		ForEach($Policy in $ScopePolicies)
		{
			Write-Verbose "$(Get-Date):	`t`t`tProcessing policy name $($Policy.Name)"
			$rowdata = @()
			$columnHeaders = @("Policy Name",($htmlsilver -bor $htmlbold),$Policy.Name,$htmlwhite)
			$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Policy.Description,$htmlwhite))
			$rowdata += @(,('Processing Order',($htmlsilver -bor $htmlbold),$Policy.ProcessingOrder.ToString(),$htmlwhite))
			$rowdata += @(,('Level',($htmlsilver -bor $htmlbold),"Scope",$htmlwhite))

			$tmp = ""
			$IPRange = Get-DHCPServerV4PolicyIPRange -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -Name $Policy.Name -EA 0
			If($? -and $Null -ne $IPRange)
			{
				$tmp = "$($IPRange.StartRange) - $($IPRange.EndRange)"
			}
			Else
			{
				$tmp = "None"
			}
			$rowdata += @(,('Address Range',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))

			$tmp = ""
			If($Policy.Enabled)
			{
				$tmp = "Enabled"
			}
			Else
			{
				$tmp = "Disabled"
			}
			$rowdata += @(,('State',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))

			$msg = ""
			$columnWidths = @("150","200")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
			InsertBlankLine
		}
	}
	ElseIf(!$?)
	{
		WriteHTMLLine 0 0 "Error retrieving scope policies"
	}
	Else
	{
		WriteHTMLLine 0 1 "None"
	}
	$ScopePolicies = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting DNS"
	WriteHTMLLine ($xStartLevel + 1) 0 "DNS"
	$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0
	If($? -and $Null -ne $DNSSettings)
	{
		GetDNSSettings $DNSSettings "A"
	}
	Else
	{
		WriteHTMLLine 0 0 "Error retrieving DNS Settings for scope $($IPv4Scope.ScopeId)"
	}
	$DNSSettings = $Null
	
	#next tab is Network Access Protection but I can't find anything that gives me that info
	
	#failover
	Write-Verbose "$(Get-Date):	`t`tGetting failover"
	WriteHTMLLine ($xStartLevel + 1) 0 "Failover"
	
	$Failovers = Get-DHCPServerV4Failover -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0
	
	If($? -and $Null -ne $Failovers)
	{
		ForEach($Failover in $Failovers)
		{
			Write-Verbose "$(Get-Date):	`t`tProcessing failover $($Failover.Name)"
			$rowdata = @()
			$columnHeaders = @("Relationship name",($htmlsilver -bor $htmlbold),$Failover.Name,$htmlwhite)
			$rowdata += @(,('Partner Server',($htmlsilver -bor $htmlbold),$Failover.PartnerServer,$htmlwhite))
			$rowdata += @(,('Mode',($htmlsilver -bor $htmlbold),$Failover.Mode,$htmlwhite))
					
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
			$rowdata += @(,('Max Client Lead Time',($htmlsilver -bor $htmlbold),$MaxLeadStr,$htmlwhite))
					
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
			$rowdata += @(,('State Switchover Interval',($htmlsilver -bor $htmlbold),$SwitchStr,$htmlwhite))
					
			Switch($Failover.State)
			{
				"NoState" 
				{
					$rowdata += @(,('State of this Server',($htmlsilver -bor $htmlbold),"No State",$htmlwhite))
					$rowdata += @(,('State of Partner Server',($htmlsilver -bor $htmlbold),"No State",$htmlwhite))
				}
				"Normal" 
				{
					$rowdata += @(,('State of this Server',($htmlsilver -bor $htmlbold),"Normal",$htmlwhite))
					$rowdata += @(,('State of Partner Server',($htmlsilver -bor $htmlbold),"Normal",$htmlwhite))
				}
				"Init" 
				{
					$rowdata += @(,('State of this Server',($htmlsilver -bor $htmlbold),"Initializing",$htmlwhite))
					$rowdata += @(,('State of Partner Server',($htmlsilver -bor $htmlbold),"Initializing",$htmlwhite))
				}
				"CommunicationInterrupted" 
				{
					$rowdata += @(,('State of this Server',($htmlsilver -bor $htmlbold),"Communication Interrupted",$htmlwhite))
					$rowdata += @(,('State of Partner Server',($htmlsilver -bor $htmlbold),"Communication Interrupted",$htmlwhite))
				}
				"PartnerDown" 
				{
					$rowdata += @(,('State of this Server',($htmlsilver -bor $htmlbold),"Normal",$htmlwhite))
					$rowdata += @(,('State of Partner Server',($htmlsilver -bor $htmlbold),"Down",$htmlwhite))
				}
				"PotentialConflict" 
				{
					$rowdata += @(,('State of this Server',($htmlsilver -bor $htmlbold),"Potential Conflict",$htmlwhite))
					$rowdata += @(,('State of Partner Server',($htmlsilver -bor $htmlbold),"Potential Conflict",$htmlwhite))
				}
				"Startup" 
				{
					$rowdata += @(,('State of this Server',($htmlsilver -bor $htmlbold),"Starting Up",$htmlwhite))
					$rowdata += @(,('State of Partner Server',($htmlsilver -bor $htmlbold),"Starting Up",$htmlwhite))
				}
				"ResolutionInterrupted" 
				{
					$rowdata += @(,('State of this Server',($htmlsilver -bor $htmlbold),"Resolution Interrupted",$htmlwhite))
					$rowdata += @(,('State of Partner Server',($htmlsilver -bor $htmlbold),"Resolution Interrupted",$htmlwhite))
				}
				"ConflictDone" 
				{
					$rowdata += @(,('State of this Server',($htmlsilver -bor $htmlbold),"Conflict Done",$htmlwhite))
					$rowdata += @(,('State of Partner Server',($htmlsilver -bor $htmlbold),"Conflict Done",$htmlwhite))
				}
				"Recover" 
				{
					$rowdata += @(,('State of this Server',($htmlsilver -bor $htmlbold),"Recover",$htmlwhite))
					$rowdata += @(,('State of Partner Server',($htmlsilver -bor $htmlbold),"Recover",$htmlwhite))
				}
				"RecoverWait" 
				{
					$rowdata += @(,('State of this Server',($htmlsilver -bor $htmlbold),"Recover Wait",$htmlwhite))
					$rowdata += @(,('State of Partner Server',($htmlsilver -bor $htmlbold),"Recover Wait",$htmlwhite))
				}
				"RecoverDone" 
				{
					$rowdata += @(,('State of this Server',($htmlsilver -bor $htmlbold),"Recover Done",$htmlwhite))
					$rowdata += @(,('State of Partner Server',($htmlsilver -bor $htmlbold),"Recover Done",$htmlwhite))
				}
				Default 
				{
					$rowdata += @(,('State of this Server',($htmlsilver -bor $htmlbold),"Unable to determine server state",$htmlwhite))
					$rowdata += @(,('State of Partner Server',($htmlsilver -bor $htmlbold),"Unable to determine server state",$htmlwhite))
				}
			}
					
			If($Failover.Mode -eq "LoadBalance")
			{
				$rowdata += @(,('Local server',($htmlsilver -bor $htmlbold),"$($Failover.LoadBalancePercent)%",$htmlwhite))
					
				$tmp = (100 - $($Failover.LoadBalancePercent))
				$rowdata += @(,('Partner Server',($htmlsilver -bor $htmlbold),"$($tmp)%",$htmlwhite))
			}
			Else
			{
				$rowdata += @(,('Role of this server',($htmlsilver -bor $htmlbold),$Failover.ServerRole,$htmlwhite))
					
				$rowdata += @(,('Addresses reserved for standby server',($htmlsilver -bor $htmlbold),"$($Failover.ReservePercent)%",$htmlwhite))
			}

			$msg = ""
			$columnWidths = @("200","100")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "300"
			InsertBlankLine
		}
	}
	Else
	{
		WriteHTMLLine 0 1 "None"
	}
	$Failovers = $Null

	Write-Verbose "$(Get-Date):	`t`tGetting Scope statistics"
	WriteHTMLLine ($xStartLevel + 1) 0 "Statistics"

	$Statistics = Get-DHCPServerV4ScopeStatistics -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0

	If($? -and $Null -ne $Statistics)
	{
		GetShortStatistics $Statistics
	}
	ElseIf(!$?)
	{
		WriteHTMLLine 0 0 "Error retrieving scope statistics"
	}
	Else
	{
		WriteHTMLLine 0 1 "None"
	}
	$Statistics = $Null
}

Function GetIPv4ScopeData_Text
{
	Param([int] $xStartLevel, [object]$filters)

	Write-Verbose "$(Get-Date): `tGetting IPv4 scope data for scope $($IPv4Scope.Name)"
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
	If(![string]::IsNullOrEmpty($IPv4Scope.Description))
	{
		Line 2 "Description`t`t: " $IPv4Scope.Description
	}

	If($IncludeLeases)
	{
		Write-Verbose "$(Get-Date): `t`tGetting leases"
		
		Line 1 "Address Leases:"
		$Leases = Get-DHCPServerV4Lease -ComputerName $Script:DHCPServerName -ScopeId  $IPv4Scope.ScopeId -EA 0 | Sort-Object IPAddress
		If($? -and $Null -ne $Leases)
		{
			ForEach($Lease in $Leases)
			{
				Write-Verbose "$(Get-Date): `t`t`tProcessing lease $($Lease.IPAddress)"
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
				If(![string]::IsNullOrEmpty($Lease.Description))
				{
					Line 2 "Description`t`t`t: " $Lease.Description
				}
				Line 2 "Network Access Protection`t: " $Lease.NapStatus
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
				
				#$Filters | ForEach-Object { $index = $Null }{ If( $_.MacAddress -eq $Lease.ClientID ) { $index = $_ } }
				
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
					Line 0 "<None>"
				}
				Line 2 "Policy`t`t`t`t: " -NoNewLine
				If([string]::IsNullOrEmpty($Lease.PolicyName))
				{
					Line 0 "<None>"
				}
				Else
				{
					Line 0 $Lease.PolicyName
				}
				
				#skip a row for spacing
				Line 0 ""
			}
		}
		ElseIf(!$?)
		{
			Line 0 "Error retrieving leases for scope $IPv4Scope.ScopeId"
		}
		Else
		{
			Line 2 "<None>"
		}
		$Leases = $Null
	}

	Write-Verbose "$(Get-Date): `t`tGetting exclusions"
	Line 1 "Exclusions:"
	$Exclusions = Get-DHCPServerV4ExclusionRange -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object StartRange
	If($? -and $Null -ne $Exclusions)
	{
		ForEach($Exclusion in $Exclusions)
		{
			Line 2 "Start IP Address`t: " $Exclusion.StartRange.ToString()
			Line 2 "End IP Address`t`t: " $Exclusion.EndRange.ToString() 
			Line 0 ""
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving exclusions for scope $IPv4Scope.ScopeId"
	}
	Else
	{
		Line 2 "<None>"
	}
	$Exclusions = $Null
	
	Write-Verbose "$(Get-Date): `t`tGetting reservations"
	Line 1 "Reservations:"
	$Reservations = Get-DHCPServerV4Reservation -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object IPAddress
	If($? -and $Null -ne $Reservations)
	{
		ForEach($Reservation in $Reservations)
		{
			Write-Verbose "$(Get-Date): `t`t`tProcessing reservation $($Reservation.Name)"
			Line 2 "Reservation name`t: " $Reservation.Name
			Line 2 "IP address`t`t: " $Reservation.IPAddress
			Line 2 "MAC address`t`t: " $Reservation.ClientId
			Line 2 "Supported types`t`t: " $Reservation.Type
			If(![string]::IsNullOrEmpty($Reservation.Description))
			{
				Line 2 "Description`t`t: " $Reservation.Description
			}

			Write-Verbose "$(Get-Date): `t`t`t`tGetting DNS settings"
			$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $Script:DHCPServerName -IPAddress $Reservation.IPAddress -EA 0
			If($? -and $Null -ne $DNSSettings)
			{
				GetDNSSettings $DNSSettings "A"
			}
			Else
			{
				Line 0 "Error retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
			}
			$DNSSettings = $Null
			Line 0 ""
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving reservations for scope $IPv4Scope.ScopeId"
	}
	Else
	{
		Line 2 "<None>"
	}
	$Reservations = $Null

	Write-Verbose "$(Get-Date): `t`tGetting scope options"
	Line 1 "Scope Options:"
	$ScopeOptions = Get-DHCPServerV4OptionValue -All -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object OptionId

	If($? -and $Null -ne $ScopeOptions)
	{
		If($ScopeOptions -is [array])
		{
			#do nothing for text
		}
		Else
		{
			If($ScopeOptions.OptionId -eq 51)
			{
				#if the only scope option is 51, ignore it
				Line 2 "<None>"
			}
			Else
			{
				#do nothing for text
			}
		}

		ForEach($ScopeOption in $ScopeOptions)
		{
			If($ScopeOption.OptionId -eq 51 -or $ScopeOption.OptionId -eq 81)
			{
				#ignore these two option IDs
				https://carlwebster.com/the-mysterious-microsoft-dhcp-option-id-81/
				https://jimswirelessworld.wordpress.com/2019/01/03/you-should-care-about-dhcp-option-51/
			}
			Else
			{
				Write-Verbose "$(Get-Date): `t`t`tProcessing option name $($ScopeOption.Name)"
				Line 2 "Option Name`t: $($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)" 
				Line 2 "Vendor`t`t: " -NoNewLine
				If([string]::IsNullOrEmpty($ScopeOption.VendorClass))
				{
					Line 0 "Standard" 
				}
				Else
				{
					Line 0 $ScopeOption.VendorClass 
				}
				Line 2 "Value`t`t: $($ScopeOption.Value)" 
				Line 2 "Policy Name`t: " -NoNewLine
				
				If([string]::IsNullOrEmpty($ScopeOption.PolicyName))
				{
					Line 0 "<None>"
				}
				Else
				{
					Line 0 $ScopeOption.PolicyName
				}
			
				#for spacing
				Line 0 ""
			}
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving scope options for $IPv4Scope.ScopeId"
	}
	Else
	{
		Line 2 "<None>"
	}
	$ScopeOptions = $Null
	
	Write-Verbose "$(Get-Date): `t`tGetting policies"
	Line 1 "Policies:"
	$ScopePolicies = Get-DHCPServerV4Policy -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0 | Sort-Object ProcessingOrder

	If($? -and $Null -ne $ScopePolicies)
	{
		ForEach($Policy in $ScopePolicies)
		{
			Write-Verbose "$(Get-Date): `t`t`tProcessing policy name $($Policy.Name)"
			Line 2 "Policy Name`t`t: " $Policy.Name
			If(![string]::IsNullOrEmpty($Policy.Description))
			{
				Line 2 "Description`t`t: " $Policy.Description
			}
			Line 2 "Processing Order`t: " $Policy.ProcessingOrder
			Line 2 "Level`t`t`t: Scope"
			Line 2 "Address Range`t: " -NoNewLine
			
			$IPRange = Get-DHCPServerV4PolicyIPRange -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -Name $Policy.Name -EA 0

			If($? -and $Null -ne $IPRange)
			{
				Line 0 "$($IPRange.StartRange) - $($IPRange.EndRange)"
			}
			Else
			{
				Line 0 "<None>"
			}
			Line 2 "State`t`t`t: " -NoNewLine
			If($Policy.Enabled)
			{
				Line 0 "Enabled"
			}
			Else
			{
				Line 0 "Disabled"
			}
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving scope policies"
	}
	Else
	{
		Line 2 "<None>"
	}
	$ScopePolicies = $Null

	Write-Verbose "$(Get-Date): `t`tGetting DNS"
	Line 1 "DNS:"
	$DNSSettings = Get-DHCPServerV4DnsSetting -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0
	If($? -and $Null -ne $DNSSettings)
	{
		GetDNSSettings $DNSSettings "A"
	}
	Else
	{
		Line 0 "Error retrieving DNS Settings for scope $($IPv4Scope.ScopeId)"
	}
	$DNSSettings = $Null
	
	#next tab is Network Access Protection but I can't find anything that gives me that info
	
	#failover
	Write-Verbose "$(Get-Date): `t`tGetting Failover"
	Line 1 "Failover:"
	
	$Failovers = Get-DHCPServerV4Failover -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0

	If($? -and $Null -ne $Failovers)
	{
		ForEach($Failover in $Failovers)
		{
			Write-Verbose "$(Get-Date): `t`tProcessing failover $($Failover.Name)"
			Line 2 "Relationship name: " $Failover.Name
			Line 2 "Partner Server`t`t`t: " $Failover.PartnerServer
			Line 2 "Mode`t`t`t`t: " $Failover.Mode
					
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

			Line 2 "Max Client Lead Time`t`t: " $MaxLeadStr
					
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

			Line 2 "State Switchover Interval`t: " $SwitchStr
					
			Switch($Failover.State)
			{
				"NoState" 
				{
					Line 2 "State of this Server`t`t: No State"
					Line 2 "State of Partner Server`t`t: No State"
				}
				"Normal" 
				{
					Line 2 "State of this Server`t`t: Normal"
					Line 2 "State of Partner Server`t`t: Normal"
				}
				"Init" 
				{
					Line 2 "State of this Server`t`t: Initializing"
					Line 2 "State of Partner Server`t`t: Initializing"
				}
				"CommunicationInterrupted" 
				{
					Line 2 "State of this Server`t`t: Communication Interrupted"
					Line 2 "State of Partner Server`t`t: Communication Interrupted"
				}
				"PartnerDown" 
				{
					Line 2 "State of this Server`t`t: Normal"
					Line 2 "State of Partner Server`t`t: Down"
				}
				"PotentialConflict" 
				{
					Line 2 "State of this Server`t`t: Potential Conflict"
					Line 2 "State of Partner Server`t`t: Potential Conflict"
				}
				"Startup" 
				{
					Line 2 "State of this Server`t`t: Starting Up"
					Line 2 "State of Partner Server`t`t: Starting Up"
				}
				"ResolutionInterrupted" 
				{
					Line 2 "State of this Server`t`t: Resolution Interrupted"
					Line 2 "State of Partner Server`t`t: Resolution Interrupted"
				}
				"ConflictDone" 
				{
					Line 2 "State of this Server`t`t: Conflict Done"
					Line 2 "State of Partner Server`t`t: Conflict Done"
				}
				"Recover" 
				{
					Line 2 "State of this Server`t`t: Recover"
					Line 2 "State of Partner Server`t`t: Recover"
				}
				"RecoverWait" 
				{
					Line 2 "State of this Server`t`t: Recover Wait"
					Line 2 "State of Partner Server`t`t: Recover Wait"
				}
				"RecoverDone" 
				{
					Line 2 "State of this Server`t`t: Recover Done"
					Line 2 "State of Partner Server`t`t: Recover Done"
				}
				Default 
				{
					Line 2 "State of this Server`t`t: Unable to determine Server: state"
					Line 2 "State of Partner Server`t`t: Unable to determine Server: state"
				}
			}
					
			If($Failover.Mode -eq "LoadBalance")
			{
				Line 2 "Local server`t`t`t: $($Failover.LoadBalancePercent)%"
				Line 2 "Partner Server`t`t`t: " -NoNewLine
				$tmp = (100 - $($Failover.LoadBalancePercent))
				Line 0 "$($tmp)%"
				$tmp = $Null
			}
			Else
			{
				Line 2 "Role of this server`t`t: " $Failover.ServerRole
				Line 2 "Addresses reserved for standby server: $($Failover.ReservePercent)%"
			}
					
			#skip a row for spacing
			Line 0 ""
		}
	}
	Else
	{
		Line 2 "<None>"
	}
	$Failovers = $Null

	Write-Verbose "$(Get-Date): `t`tGetting Scope statistics"
	Line 1 "Statistics:"

	$Statistics = Get-DHCPServerV4ScopeStatistics -ComputerName $Script:DHCPServerName -ScopeId $IPv4Scope.ScopeId -EA 0

	If($? -and $Null -ne $Statistics)
	{
		GetShortStatistics $Statistics
	}
	ElseIf(!$?)
	{
		Line "Error retrieving scope statistics"
	}
	Else
	{
		Line 2 "<None>"
	}
	$Statistics = $Null
}

Function GetIPv6ScopeData_WordPDF
{
	Write-Verbose "$(Get-Date): `tGetting IPv6 scope data for scope $($IPv6Scope.Name)"
	WriteWordLine 3 0 "Scope [$($IPv6Scope.Prefix)] $($IPv6Scope.Name)"
	WriteWordLine 4 0 "General"
	$TableRange = $doc.Application.Selection.Range
	[int]$Columns = 2
	If(![string]::IsNullOrEmpty($IPv6Scope.Description))
	{
		[int]$Rows = 6
	}
	Else
	{
		[int]$Rows = 5
	}
	ValidateWordTableValues $Rows $Script:DHCPServerName "Function GetIPv6ScopeData_WordPDF Getting IPv6 scope data for scope $($IPv6Scope.Name)"
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = $myHash.Word_TableGrid
	$table.Borders.InsideLineStyle = $wdLineStyleNone
	$table.Borders.OutsideLineStyle = $wdLineStyleNone
	$Table.Cell(1,1).Range.Text = "Prefix"
	$Table.Cell(1,2).Range.Text = $IPv6Scope.Prefix
	$Table.Cell(2,1).Range.Text = "Preference"
	$Table.Cell(2,2).Range.Text = $IPv6Scope.Preference
	$Table.Cell(3,1).Range.Text = "Available Range"
	$Table.Cell(3,2).Range.Text = ""
	$Table.Cell(4,1).Range.Text = "`tStart"
	$Table.Cell(4,2).Range.Text = "$($IPv6Scope.Prefix)0:0:0:1"
	$Table.Cell(5,1).Range.Text = "`tEnd"
	$Table.Cell(5,2).Range.Text = "$($IPv6Scope.Prefix)ffff:ffff:ffff:ffff"
	If(![string]::IsNullOrEmpty($IPv6Scope.Description))
	{
		$Table.Cell(6,1).Range.Text = "Description"
		$Table.Cell(6,2).Range.Text = $IPv6Scope.Description
	}
	$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
	$table.AutoFitBehavior($wdAutoFitContent)

	#return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
	$TableRange = $Null
	$Table = $Null

	Write-Verbose "$(Get-Date): `t`tGetting scope DNS settings"
	WriteWordLine 4 0 "DNS"
	$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0
	If($? -and $Null -ne $DNSSettings)
	{
		GetDNSSettings $DNSSettings "AAAA"
	}
	Else
	{
		WriteWordLine 0 0 "Error retrieving IPv6 DNS Settings for scope $($IPv6Scope.Prefix)"
	}
	$DNSSettings = $Null
	
	Write-Verbose "$(Get-Date): `t`tGetting scope lease settings"
	WriteWordLine 4 0 "Lease"
	
	$PrefStr = [string]::format("{0} days, {1} hours, {2} minutes", `
		$IPv6Scope.PreferredLifetime.Days, `
		$IPv6Scope.PreferredLifetime.Hours, `
		$IPv6Scope.PreferredLifetime.Minutes)
	
	$ValidStr = [string]::format("{0} days, {1} hours, {2} minutes", `
		$IPv6Scope.ValidLifetime.Days, `
		$IPv6Scope.ValidLifetime.Hours, `
		$IPv6Scope.ValidLifetime.Minutes)
	$TableRange = $doc.Application.Selection.Range
	[int]$Columns = 2
	[int]$Rows = 2
	$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
	$table.Style = $myHash.Word_TableGrid
	$table.Borders.InsideLineStyle = $wdLineStyleNone
	$table.Borders.OutsideLineStyle = $wdLineStyleNone
	$Table.Cell(1,1).Range.Text = "Preferred life time"
	$Table.Cell(1,2).Range.Text = $PrefStr
	$Table.Cell(2,1).Range.Text = "Valid life time"
	$Table.Cell(2,2).Range.Text = $ValidStr

	$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
	$table.AutoFitBehavior($wdAutoFitContent)

	#return focus back to document
	$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

	#move to the end of the current document
	$selection.EndKey($wdStory,$wdMove) | Out-Null
	$TableRange = $Null
	$Table = $Null
	
	If($IncludeLeases)
	{
		Write-Verbose "$(Get-Date): `t`tGetting leases"
		WriteWordLine 4 0 "Address Leases"
		$Leases = Get-DHCPServerV6Lease -ComputerName $Script:DHCPServerName -Prefix  $IPv6Scope.Prefix -EA 0 | Sort-Object IPAddress
		If($? -and $Null -ne $Leases)
		{
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If($Leases -is [array])
			{
				[int]$Rows = $Leases.Count * 8
			}
			Else
			{
				[int]$Rows = 7
			}
			ValidateWordTableValues $Rows $Script:DHCPServerName "Function GetIPv6ScopeData_WordPDF Getting leases"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = $wdLineStyleNone
			$table.Borders.OutsideLineStyle = $wdLineStyleNone
			[int]$xRow = 0
			ForEach($Lease in $Leases)
			{
				Write-Verbose "$(Get-Date): `t`t`tProcessing lease $($Lease.IPAddress)"
				$xRow++
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

				$Table.Cell($xRow,1).Range.Text = "Client IPv6 address"
				$Table.Cell($xRow,2).Range.Text = $Lease.IPAddress
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Name"
				$Table.Cell($xRow,2).Range.Text = $Lease.HostName
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Lease Expiration"
				$Table.Cell($xRow,2).Range.Text = $LeaseStr
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "IAID"
				$Table.Cell($xRow,2).Range.Text = $Lease.Iaid
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Type"
				$Table.Cell($xRow,2).Range.Text = $Lease.AddressType
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Unique ID"
				$Table.Cell($xRow,2).Range.Text = $Lease.ClientDuid
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Description"
				$Table.Cell($xRow,2).Range.Text = $Lease.Description
				
				#skip a row for spacing
				$xRow++
			}
			$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
			$table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
		ElseIf(!$?)
		{
			WriteWordLine 0 0 "Error retrieving leases for scope $IPv6Scope.Prefix"
		}
		Else
		{
			WriteWordLine 0 1 "<None>"
		}
		$Leases = $Null
	}

	Write-Verbose "$(Get-Date): `t`tGetting exclusions"
	WriteWordLine 4 0 "Exclusions"
	$Exclusions = Get-DHCPServerV6ExclusionRange -ComputerName $Script:DHCPServerName -Prefix  $IPv6Scope.Prefix -EA 0 | Sort-Object StartRange
	If($? -and $Null -ne $Exclusions)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($Exclusions -is [array])
		{
			[int]$Rows = $Exclusions.Count + 1
		}
		Else
		{
			[int]$Rows = 2
		}
		ValidateWordTableValues $Rows $Script:DHCPServerName "Function GetIPv6ScopeData_WordPDF Getting exclusions"
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = $wdLineStyleNone
		$table.Borders.OutsideLineStyle = $wdLineStyleNone
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "Start IP Address"
		$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,2).Range.Font.Bold = $True
		$Table.Cell(1,2).Range.Text = "End IP Address"
		[int]$xRow = 1
		ForEach($Exclusion in $Exclusions)
		{
			$xRow++
			$Table.Cell($xRow,1).Range.Text = $Exclusion.StartRange
			$Table.Cell($xRow,2).Range.Text = $Exclusion.EndRange 
		}
		$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
		$table.AutoFitBehavior($wdAutoFitContent)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving exclusions for scope $IPv6Scope.Prefix"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$Exclusions = $Null

	Write-Verbose "$(Get-Date): `t`tGetting reservations"
	WriteWordLine 4 0 "Reservations"
	$Reservations = Get-DHCPServerV6Reservation -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0 | Sort-Object IPAddress
	If($? -and $Null -ne $Reservations)
	{
		ForEach($Reservation in $Reservations)
		{
			Write-Verbose "$(Get-Date): `t`t`tProcessing reservation $($Reservation.Name)"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If([string]::IsNullOrEmpty($Reservation.Description))
			{
				[int]$Rows = 4
			}
			Else
			{
				[int]$Rows = 5
			}
			ValidateWordTableValues $Rows $Script:DHCPServerName "Function GetIPv6ScopeData_WordPDF Processing reservation $($Reservation.Name)"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = $wdLineStyleNone
			$table.Borders.OutsideLineStyle = $wdLineStyleNone
			$Table.Cell(1,1).Range.Text = "Reservation name"
			$Table.Cell(1,2).Range.Text = $Reservation.Name
			$Table.Cell(2,1).Range.Text = "IPv6 address"
			$Table.Cell(2,2).Range.Text = $Reservation.IPAddress
			$Table.Cell(3,1).Range.Text = "DUID"
			$Table.Cell(3,2).Range.Text = $Reservation.ClientDuid
			$Table.Cell(4,1).Range.Text = "IAID"
			$Table.Cell(4,2).Range.Text = $Reservation.Iaid
			If(![string]::IsNullOrEmpty($Reservation.Description))
			{
				$Table.Cell(5,1).Range.Text = "Description"
				$Table.Cell(5,2).Range.Text = $Reservation.Description
			}
			$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
			$table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null

			Write-Verbose "$(Get-Date): `t`t`t`tGetting DNS settings"
			$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $Script:DHCPServerName -IPAddress $Reservation.IPAddress -EA 0
			If($? -and $Null -ne $DNSSettings)
			{
				GetDNSSettings $DNSSettings "AAAA"
			}
			Else
			{
				WriteWordLine 0 0 "Error to retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
			}
			$DNSSettings = $Null
		}
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving reservations for scope $IPv6Scope.Prefix"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$Reservations = $Null

	Write-Verbose "$(Get-Date): Getting IPv6 scope options"
	WriteWordLine 4 0 "Scope Options"
	$ScopeOptions = Get-DHCPServerV6OptionValue -All -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0 | Sort-Object OptionId

	If($? -and $Null -ne $ScopeOptions)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		If($ScopeOptions -is [array])
		{
			[int]$Rows = $ScopeOptions.Count * 4
		}
		Else
		{
			[int]$Rows = 3
		}
		ValidateWordTableValues $Rows $Script:DHCPServerName "Function GetIPv6ScopeData_WordPDF Getting IPv6 scope options"
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = $wdLineStyleNone
		$table.Borders.OutsideLineStyle = $wdLineStyleNone
		[int]$xRow = 0
		ForEach($ScopeOption in $ScopeOptions)
		{
			Write-Verbose "$(Get-Date): `t`t`tProcessing option name $($ScopeOption.Name)"
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Option Name"
			$Table.Cell($xRow,2).Range.Text = "$($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)" 
			
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Vendor"
			If([string]::IsNullOrEmpty($ScopeOption.VendorClass))
			{
				$Table.Cell($xRow,2).Range.Text = "Standard" 
			}
			Else
			{
				$Table.Cell($xRow,2).Range.Text = $ScopeOption.VendorClass 
			}
			
			$xRow++
			$Table.Cell($xRow,1).Range.Text = "Value"
			$Table.Cell($xRow,2).Range.Text = "$($ScopeOption.Value)" 
			
			#for spacing
			$xRow++
		}
		$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
		$table.AutoFitBehavior($wdAutoFitContent)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving IPv6 scope options"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$ScopeOptions = $Null
	
	Write-Verbose "$(Get-Date): `t`tGetting Scope statistics"
	WriteWordLine 4 0 "Statistics"

	$Statistics = Get-DHCPServerV6ScopeStatistics -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0

	If($? -and $Null -ne $Statistics)
	{
		GetShortStatistics $Statistics
	}
	ElseIf(!$?)
	{
		WriteWordLine 0 0 "Error retrieving scope statistics"
	}
	Else
	{
		WriteWordLine 0 1 "<None>"
	}
	$Statistics = $Null
}

Function GetIPv6ScopeData_HTML
{
	Write-Verbose "$(Get-Date): `tGetting IPv6 scope data for scope $($IPv6Scope.Name)"
	WriteHTMLLine 3 0 "Scope [$($IPv6Scope.Prefix)] $($IPv6Scope.Name)"
	WriteHTMLLine 4 0 "General"
	$rowdata = @()
	$columnHeaders = @("Prefix",($htmlsilver -bor $htmlbold),$IPv6Scope.Prefix,$htmlwhite)
	$rowdata += @(,('Preference',($htmlsilver -bor $htmlbold),$IPv6Scope.Preference,$htmlwhite))
	$rowdata += @(,('Available Range',($htmlsilver -bor $htmlbold),"",$htmlwhite))
	$rowdata += @(,('     Start',($htmlsilver -bor $htmlbold),"$($IPv6Scope.Prefix)0:0:0:1",$htmlwhite))
	$rowdata += @(,('     End',($htmlsilver -bor $htmlbold),"$($IPv6Scope.Prefix)ffff:ffff:ffff:ffff",$htmlwhite))
	If(![string]::IsNullOrEmpty($IPv6Scope.Description))
	{
		$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$IPv6Scope.Description,$htmlwhite))
	}

	$msg = ""
	$columnWidths = @("200","100")
	FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "300"
	InsertBlankLine

	Write-Verbose "$(Get-Date): `t`tGetting scope DNS settings"
	WriteHTMLLine 4 0 "DNS"
	$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0
	If($? -and $Null -ne $DNSSettings)
	{
		GetDNSSettings $DNSSettings "AAAA"
	}
	Else
	{
		WriteHTMLLine 0 0 "Error retrieving IPv6 DNS Settings for scope $($IPv6Scope.Prefix)"
	}
	$DNSSettings = $Null
	
	Write-Verbose "$(Get-Date): `t`tGetting scope lease settings"
	WriteHTMLLine 4 0 "Lease"
	
	$PrefStr = [string]::format("{0} days, {1} hours, {2} minutes", `
		$IPv6Scope.PreferredLifetime.Days, `
		$IPv6Scope.PreferredLifetime.Hours, `
		$IPv6Scope.PreferredLifetime.Minutes)
	
	$ValidStr = [string]::format("{0} days, {1} hours, {2} minutes", `
		$IPv6Scope.ValidLifetime.Days, `
		$IPv6Scope.ValidLifetime.Hours, `
		$IPv6Scope.ValidLifetime.Minutes)
	$rowdata = @()
	$columnHeaders = @("Preferred life time",($htmlsilver -bor $htmlbold),$PrefStr,$htmlwhite)
	$rowdata += @(,('Valid life time',($htmlsilver -bor $htmlbold),$ValidStr,$htmlwhite))

	$msg = ""
	$columnWidths = @("200","100")
	FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "300"
	InsertBlankLine
	
	If($IncludeLeases)
	{
		Write-Verbose "$(Get-Date): `t`tGetting leases"
		WriteHTMLLine 4 0 "Address Leases"
		$Leases = Get-DHCPServerV6Lease -ComputerName $Script:DHCPServerName -Prefix  $IPv6Scope.Prefix -EA 0 | Sort-Object IPAddress
		If($? -and $Null -ne $Leases)
		{
			ForEach($Lease in $Leases)
			{
				Write-Verbose "$(Get-Date): `t`t`tProcessing lease $($Lease.IPAddress)"
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

				$rowdata = @()
				$columnHeaders = @("Client IPv6 address",($htmlsilver -bor $htmlbold),$Lease.IPAddress,$htmlwhite)
				$rowdata += @(,('Name',($htmlsilver -bor $htmlbold),$Lease.HostName,$htmlwhite))
				$rowdata += @(,('Lease Expiration',($htmlsilver -bor $htmlbold),$LeaseStr,$htmlwhite))
				$rowdata += @(,('IAID',($htmlsilver -bor $htmlbold),$Lease.Iaid,$htmlwhite))
				$rowdata += @(,('Type',($htmlsilver -bor $htmlbold),$Lease.AddressType,$htmlwhite))
				$rowdata += @(,('Unique ID',($htmlsilver -bor $htmlbold),$Lease.ClientDuid,$htmlwhite))
				$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Lease.Description,$htmlwhite))
				
				$msg = ""
				$columnWidths = @("200","100")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "300"
				InsertBlankLine
			}
		}
		ElseIf(!$?)
		{
			WriteHTMLLine 0 0 "Error retrieving leases for scope $IPv6Scope.Prefix"
		}
		Else
		{
			WriteHTMLLine 0 1 "None"
		}
		$Leases = $Null
	}

	Write-Verbose "$(Get-Date): `t`tGetting exclusions"
	WriteHTMLLine 4 0 "Exclusions"
	$Exclusions = Get-DHCPServerV6ExclusionRange -ComputerName $Script:DHCPServerName -Prefix  $IPv6Scope.Prefix -EA 0 | Sort-Object StartRange
	If($? -and $Null -ne $Exclusions)
	{
		$rowdata = @()
		ForEach($Exclusion in $Exclusions)
		{
			$rowdata += @(,($Exclusion.StartRange,$htmlwhite,
							$Exclusion.EndRange,$htmlwhite))
		}

		$columnHeaders = @('Start IP Address',($htmlsilver -bor $htmlbold),'End IP Address',($htmlsilver -bor $htmlbold))
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		InsertBlankLine
	}
	ElseIf(!$?)
	{
		WriteHTMLLine 0 0 "Error retrieving exclusions for scope $IPv6Scope.Prefix"
	}
	Else
	{
		WriteHTMLLine 0 1 "None"
	}
	$Exclusions = $Null

	Write-Verbose "$(Get-Date): `t`tGetting reservations"
	WriteHTMLLine 4 0 "Reservations"
	$Reservations = Get-DHCPServerV6Reservation -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0 | Sort-Object IPAddress
	If($? -and $Null -ne $Reservations)
	{
		ForEach($Reservation in $Reservations)
		{
			Write-Verbose "$(Get-Date): `t`t`tProcessing reservation $($Reservation.Name)"
			$rowdata = @()
			$columnHeaders = @("Reservation name",($htmlsilver -bor $htmlbold),$Reservation.Name,$htmlwhite)
			$rowdata += @(,('IPv6 address',($htmlsilver -bor $htmlbold),$Reservation.IPAddress.ToString(),$htmlwhite))
			$rowdata += @(,('DUID',($htmlsilver -bor $htmlbold),$Reservation.ClientDuid,$htmlwhite))
			$rowdata += @(,('IAID',($htmlsilver -bor $htmlbold),$Reservation.Iaid,$htmlwhite))
			If(![string]::IsNullOrEmpty($Reservation.Description))
			{
				$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Reservation.Description,$htmlwhite))
			}
			$msg = ""
			$columnWidths = @("200","100")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "300"
			InsertBlankLine

			Write-Verbose "$(Get-Date): `t`t`t`tGetting DNS settings"
			$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $Script:DHCPServerName -IPAddress $Reservation.IPAddress -EA 0
			If($? -and $Null -ne $DNSSettings)
			{
				GetDNSSettings $DNSSettings "AAAA"
			}
			Else
			{
				WriteHTMLLine 0 0 "Error to retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
			}
			$DNSSettings = $Null
		}
	}
	ElseIf(!$?)
	{
		WriteHTMLLine 0 0 "Error retrieving reservations for scope $IPv6Scope.Prefix"
	}
	Else
	{
		WriteHTMLLine 0 1 "None"
	}
	$Reservations = $Null

	Write-Verbose "$(Get-Date): Getting IPv6 scope options"
	WriteHTMLLine 4 0 "Scope Options"
	$ScopeOptions = Get-DHCPServerV6OptionValue -All -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0 | Sort-Object OptionId

	If($? -and $Null -ne $ScopeOptions)
	{
		ForEach($ScopeOption in $ScopeOptions)
		{
			Write-Verbose "$(Get-Date): `t`t`tProcessing option name $($ScopeOption.Name)"
			$rowdata = @()
			$columnHeaders = @("Option Name",($htmlsilver -bor $htmlbold),"$($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)",$htmlwhite)
			
			$tmp = ""
			If([string]::IsNullOrEmpty($ScopeOption.VendorClass))
			{
				$tmp = "Standard" 
			}
			Else
			{
				$tmp = $ScopeOption.VendorClass 
			}
			$rowdata += @(,('Vendor',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			$rowdata += @(,('Value',($htmlsilver -bor $htmlbold),"$($ScopeOption.Value)",$htmlwhite))
			
			$msg = ""
			$columnWidths = @("200","100")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "300"
			InsertBlankLine
		}
	}
	ElseIf(!$?)
	{
		WriteHTMLLine 0 0 "Error retrieving IPv6 scope options"
	}
	Else
	{
		WriteHTMLLine 0 1 "None"
	}
	$ScopeOptions = $Null
	
	Write-Verbose "$(Get-Date): `t`tGetting Scope statistics"
	WriteHTMLLine 4 0 "Statistics"

	$Statistics = Get-DHCPServerV6ScopeStatistics -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0

	If($? -and $Null -ne $Statistics)
	{
		GetShortStatistics $Statistics
	}
	ElseIf(!$?)
	{
		WriteHTMLLine 0 0 "Error retrieving scope statistics"
	}
	Else
	{
		WriteHTMLLine 0 1 "None"
	}
	$Statistics = $Null
}

Function GetIPv6ScopeData_Text
{
	Write-Verbose "$(Get-Date): `tGetting IPv6 scope data for scope $($IPv6Scope.Name)"
	Line 0 "Scope [$($IPv6Scope.Prefix)] $($IPv6Scope.Name)"
	Line 1 "General"
	Line 2 "Prefix`t`t: " $IPv6Scope.Prefix
	Line 2 "Preference`t: " $IPv6Scope.Preference
	Line 2 "Available Range`t: "
	Line 3 "Start`t: $($IPv6Scope.Prefix)0:0:0:1"
	Line 3 "End`t: $($IPv6Scope.Prefix)ffff:ffff:ffff:ffff"
	If(![string]::IsNullOrEmpty($IPv6Scope.Description))
	{
		Line 2 "Description`t: " $IPv6Scope.Description
	}

	Write-Verbose "$(Get-Date): `t`tGetting scope DNS settings"
	Line 1 "DNS"
	$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0
	If($? -and $Null -ne $DNSSettings)
	{
		GetDNSSettings $DNSSettings "AAAA"
	}
	Else
	{
		Line 0 "Error retrieving IPv6 DNS Settings for scope $($IPv6Scope.Prefix)"
	}
	$DNSSettings = $Null
	
	Write-Verbose "$(Get-Date): `t`tGetting scope lease settings"
	Line 1 "Lease"
	
	$PrefStr = [string]::format("{0} days, {1} hours, {2} minutes", `
		$IPv6Scope.PreferredLifetime.Days, `
		$IPv6Scope.PreferredLifetime.Hours, `
		$IPv6Scope.PreferredLifetime.Minutes)
	
	$ValidStr = [string]::format("{0} days, {1} hours, {2} minutes", `
		$IPv6Scope.ValidLifetime.Days, `
		$IPv6Scope.ValidLifetime.Hours, `
		$IPv6Scope.ValidLifetime.Minutes)
	Line 2 "Preferred life time`t: " $PrefStr
	Line 2 "Valid life time`t`t: " $ValidStr

	If($IncludeLeases)
	{
		Write-Verbose "$(Get-Date): `t`tGetting leases"
		Line 1 "Address Leases:"
		$Leases = Get-DHCPServerV6Lease -ComputerName $Script:DHCPServerName -Prefix  $IPv6Scope.Prefix -EA 0 | Sort-Object IPAddress
		If($? -and $Null -ne $Leases)
		{
			ForEach($Lease in $Leases)
			{
				Write-Verbose "$(Get-Date): `t`t`tProcessing lease $($Lease.IPAddress)"
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

				Line 2 "Client IPv6 address: " $Lease.IPAddress
				Line 2 "Name`t`t`t: " $Lease.HostName
				Line 2 "Lease Expiration`t: " $LeaseStr
				Line 2 "IAID`t`t`t: " $Lease.Iaid
				Line 2 "Type`t`t`t: " $Lease.AddressType
				Line 2 "Unique ID`t`t: " $Lease.ClientDuid
				If(![string]::IsNullOrEmpty($Lease.Description))
				{
					Line 2 "Description`t`t: " $Lease.Description
				}
				
				#skip a row for spacing
				Line 0 ""
			}
		}
		ElseIf(!$?)
		{
			Line 0 "Error retrieving leases for scope $IPv6Scope.Prefix"
		}
		Else
		{
			Line 2 "<None>"
		}
		$Leases = $Null
	}

	Write-Verbose "$(Get-Date): `t`tGetting exclusions"
	Line 1 "Exclusions:"
	$Exclusions = Get-DHCPServerV6ExclusionRange -ComputerName $Script:DHCPServerName -Prefix  $IPv6Scope.Prefix -EA 0 | Sort-Object StartRange
	If($? -and $Null -ne $Exclusions)
	{
		ForEach($Exclusion in $Exclusions)
		{
			Line 2 "Start IP Address`t: " $Exclusion.StartRange
			Line 2 "End IP Address`t`t: " $Exclusion.EndRange 
			Line 0 ""
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving exclusions for scope $IPv6Scope.Prefix"
	}
	Else
	{
		Line 2 "<None>"
	}
	$Exclusions = $Null

	Write-Verbose "$(Get-Date): `t`tGetting reservations"
	Line 1 "Reservations:"
	$Reservations = Get-DHCPServerV6Reservation -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0 | Sort-Object IPAddress
	If($? -and $Null -ne $Reservations)
	{
		ForEach($Reservation in $Reservations)
		{
			Write-Verbose "$(Get-Date): `t`t`tProcessing reservation $($Reservation.Name)"
			Line 2 "Reservation name: " $Reservation.Name
			Line 2 "IPv6 address: " $Reservation.IPAddress
			Line 2 "DUID`t`t: " $Reservation.ClientDuid
			Line 2 "IAID`t`t: " $Reservation.Iaid
			If(![string]::IsNullOrEmpty($Reservation.Description))
			{
				Line 2 "Description`t: " $Reservation.Description
			}

			Write-Verbose "$(Get-Date): `t`t`t`tGetting DNS settings"
			$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $Script:DHCPServerName -IPAddress $Reservation.IPAddress -EA 0
			If($? -and $Null -ne $DNSSettings)
			{
				GetDNSSettings $DNSSettings "AAAA"
			}
			Else
			{
				Line 0 "Error to retrieving DNS Settings for reserved IP address $Reservation.IPAddress"
			}
			$DNSSettings = $Null
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving reservations for scope $IPv6Scope.Prefix"
	}
	Else
	{
		Line 2 "<None>"
	}
	$Reservations = $Null

	Write-Verbose "$(Get-Date): Getting IPv6 scope options"
	Line 1 "Scope Options:"
	$ScopeOptions = Get-DHCPServerV6OptionValue -All -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0 | Sort-Object OptionId

	If($? -and $Null -ne $ScopeOptions)
	{
		ForEach($ScopeOption in $ScopeOptions)
		{
			Write-Verbose "$(Get-Date): `t`t`tProcessing option name $($ScopeOption.Name)"
			Line 2 "Option Name`t: $($ScopeOption.OptionId.ToString("00000")) $($ScopeOption.Name)" 
			Line 2 "Vendor`t`t: " -NoNewLine
			If([string]::IsNullOrEmpty($ScopeOption.VendorClass))
			{
				Line 0 "Standard" 
			}
			Else
			{
				Line 0 $ScopeOption.VendorClass 
			}
			Line 2 "Value`t`t: $($ScopeOption.Value)" 
			
			#for spacing
			Line 0 ""
		}
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving IPv6 scope options"
	}
	Else
	{
		Line 2 "<None>"
	}
	$ScopeOptions = $Null
	
	Write-Verbose "$(Get-Date): `t`tGetting Scope statistics"
	Line 1 "Statistics:"

	$Statistics = Get-DHCPServerV6ScopeStatistics -ComputerName $Script:DHCPServerName -Prefix $IPv6Scope.Prefix -EA 0

	If($? -and $Null -ne $Statistics)
	{
		GetShortStatistics $Statistics
	}
	ElseIf(!$?)
	{
		Line 0 "Error retrieving scope statistics"
	}
	Else
	{
		Line 2 "<None>"
	}
	$Statistics = $Null
}

Function GetIPV6ScopeData
{
	Param([object]$IPv6Scope)

	If($MSWord -or $PDF)
	{
		GetIPv6ScopeData_WordPDF
	}
	ElseIf($Text)
	{
		GetIPv6ScopeData_Text
	}
	ElseIf($HTML)
	{
		GetIPv6ScopeData_HTML
	}
}

Function GetDNSSettings
{
	Param([object]$DNSSettings, [string]$As)
	
	If($MSWord -or $PDF)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		[int]$Rows = 4
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = $wdLineStyleNone
		$table.Borders.OutsideLineStyle = $wdLineStyleNone
		[int]$xxRow = 1
		$Table.Cell($xxRow,1).Range.Text = "Enable DNS dynamic updates"
		If($DNSSettings.DynamicUpdates -eq "Never")
		{
			$Table.Cell($xxRow,2).Range.Text =  "Disabled"
		}
		ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
		{
			$Table.Cell($xxRow,2).Range.Text =  "Enabled"
			$xxRow++
			$Table.Cell($xxRow,1).Range.Text =  "Dynamically update DNS $($As) and PTR records only if requested by the DHCP clients"
			$Table.Cell($xxRow,2).Range.Text =  "Enabled"
		}
		ElseIf($DNSSettings.DynamicUpdates -eq "Always")
		{
			$Table.Cell($xxRow,2).Range.Text =  "Enabled"
			$xxRow++
			$Table.Cell($xxRow,1).Range.Text =  "Always dynamically update DNS $($As) and PTR records"
			$Table.Cell($xxRow,2).Range.Text =  "Enabled"
		}
		$xxRow++
		$Table.Cell($xxRow,1).Range.Text = "Discard $($As) and PTR records when lease is deleted"
		If($DNSSettings.DeleteDnsRROnLeaseExpiry)
		{
			$Table.Cell($xxRow,2).Range.Text = "Enabled"
		}
		Else
		{
			$Table.Cell($xxRow,2).Range.Text = "Disabled"
		}
		$xxRow++
		$Table.Cell($xxRow,1).Range.Text = "Name Protection"
		If($DNSSettings.NameProtection)
		{
			$Table.Cell($xxRow,2).Range.Text = "Enabled"
		}
		Else
		{
			$Table.Cell($xxRow,2).Range.Text = "Disabled"
		}

		$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
		$table.AutoFitBehavior($wdAutoFitContent)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null
	}
	ElseIf($Text)
	{
		Line 2 "Enable DNS dynamic updates`t`t`t: " -NoNewLine
		If($DNSSettings.DynamicUpdates -eq "Never")
		{
			Line 0 "Disabled"
		}
		ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
		{
			Line 0 "Enabled"
			Line 2 "Dynamically update DNS $($As) & PTR records only if requested by the DHCP clients"
		}
		ElseIf($DNSSettings.DynamicUpdates -eq "Always")
		{
			Line 0 "Enabled"
			Line 2 "Always dynamically update DNS $($As) & PTR records"
		}
		Line 2 "Discard $($As) & PTR records when lease deleted`t: " -NoNewLine
		If($DNSSettings.DeleteDnsRROnLeaseExpiry)
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
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		If($DNSSettings.DynamicUpdates -eq "Never")
		{
			$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite)
		}
		ElseIf($DNSSettings.DynamicUpdates -eq "OnClientRequest")
		{
			$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
			$rowdata += @(,("Dynamically update DNS $($As) & PTR records only if requested by the DHCP clients",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
		}
		ElseIf($DNSSettings.DynamicUpdates -eq "Always")
		{
			$columnHeaders = @("Enable DNS dynamic updates",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
			$rowdata += @(,("Always dynamically update DNS $($As) & PTR records",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
		}
		If($DNSSettings.DeleteDnsRROnLeaseExpiry)
		{
			$rowdata += @(,("Discard $($As) & PTR records when lease deleted",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite))
		}
		Else
		{
			$rowdata += @(,("Discard $($As) & PTR records when lease deleted",($htmlsilver -bor $htmlbold),"Disabled",$htmlwhite))
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
		$columnWidths = @("250","100")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
		InsertBlankLine
	}
}

Function GetShortStatistics
{
	Param([object]$Statistics)
	
	If($MSWord -or $PDF)
	{
		$TableRange = $doc.Application.Selection.Range
		[int]$Columns = 2
		[int]$Rows = 4
		$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
		$table.Style = $myHash.Word_TableGrid
		$table.Borders.InsideLineStyle = $wdLineStyleSingle
		$table.Borders.OutsideLineStyle = $wdLineStyleSingle
		$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,1).Range.Font.Bold = $True
		$Table.Cell(1,1).Range.Text = "Description"
		$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
		$Table.Cell(1,2).Range.Font.Bold = $True
		$Table.Cell(1,2).Range.Text = "Details"

		$Table.Cell(2,1).Range.Text = "Total Addresses"
		[decimal]$TotalAddresses = "{0:N0}" -f ($Statistics.AddressesFree + $Statistics.AddressesInUse)
		$Table.Cell(2,2).Range.Text = $TotalAddresses.ToString()
		$Table.Cell(3,1).Range.Text = "In Use"
		[int]$InUsePercent = "{0:N0}" -f $Statistics.PercentageInUse.ToString()
		$Table.Cell(3,2).Range.Text = "$($Statistics.AddressesInUse) ($($InUsePercent))%"
		$Table.Cell(4,1).Range.Text = "Available"
		If($TotalAddresses -ne 0)
		{
			[int]$AvailablePercent = "{0:N0}" -f (($Statistics.AddressesFree / $TotalAddresses) * 100)
		}
		Else
		{
			[int]$AvailablePercent = 0
		}
		$Table.Cell(4,2).Range.Text = "$($Statistics.AddressesFree) ($($AvailablePercent))%"

		$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
		$table.AutoFitBehavior($wdAutoFitContent)

		#return focus back to document
		$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

		#move to the end of the current document
		$selection.EndKey($wdStory,$wdMove) | Out-Null
		$TableRange = $Null
		$Table = $Null
	}
	ElseIf($Text)
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
		Line 0 "$($tmp) ($($InUsePercent))%"
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
		Line 0 "$($tmp) ($($AvailablePercent))%"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()

		[decimal]$TotalAddresses = "{0:N0}" -f ($Statistics.AddressesFree + $Statistics.AddressesInUse)
		$rowdata += @(,("Total Addresses",$htmlwhite,
						$TotalAddresses.ToString(),$htmlwhite))

		[int]$InUsePercent = "{0:N0}" -f $Statistics.PercentageInUse.ToString()
		$rowdata += @(,("In Use",$htmlwhite,
						"$($Statistics.AddressesInUse) ($($InUsePercent))%",$htmlwhite))

		If($TotalAddresses -ne 0)
		{
			[int]$AvailablePercent = "{0:N0}" -f (($Statistics.AddressesFree / $TotalAddresses) * 100)
		}
		Else
		{
			[int]$AvailablePercent = 0
		}
		$rowdata += @(,("Available",$htmlwhite,
						"$($Statistics.AddressesFree) ($($AvailablePercent))%",$htmlwhite))

		$columnHeaders = @('Description',($htmlsilver -bor $htmlbold),'Details',($htmlsilver -bor $htmlbold))
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		InsertBlankLine
	}
}

Function ProcessIPv4Scopes
{
	Write-Verbose "$(Get-Date): Getting IPv4 scopes"
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
	[gc]::collect() 
}

Function ProcessIPv4MulticastScopes
{
	$CmdletName = "Get-DHCPServerV4MulticastScope"
	If(Get-Command $CmdletName -Module "DHCPServer" -EA 0)
	{
		Write-Verbose "$(Get-Date): Getting IPv4 Multicast scopes"
		$IPv4MulticastScopes = Get-DHCPServerV4MulticastScope -ComputerName $Script:DHCPServerName -EA 0

		If($? -and $Null -ne $IPv4MulticastScopes)
		{
			ForEach($IPv4MulticastScope in $IPv4MulticastScopes)
			{
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
					Write-Verbose "$(Get-Date): `tGetting IPv4 multicast scope data for scope $($IPv4MulticastScope.Name)"
					WriteWordLine 3 0 "Multicast Scope [$($IPv4MulticastScope.Name)]"
					WriteWordLine 4 0 "General"
					$TableRange = $doc.Application.Selection.Range
					[int]$Columns = 2
					[int]$Rows = 6
					$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
					$table.Style = $myHash.Word_TableGrid
					$table.Borders.InsideLineStyle = $wdLineStyleNone
					$table.Borders.OutsideLineStyle = $wdLineStyleNone
					$Table.Cell(1,1).Range.Text = "Name"
					$Table.Cell(1,2).Range.Text = $IPv4MulticastScope.Name
					$Table.Cell(2,1).Range.Text = "Start IP address"
					$Table.Cell(2,2).Range.Text = $IPv4MulticastScope.StartRange
					$Table.Cell(3,1).Range.Text = "End IP address"
					$Table.Cell(3,2).Range.Text = $IPv4MulticastScope.EndRange
					$Table.Cell(4,1).Range.Text = "Time to live"
					$Table.Cell(4,2).Range.Text = $IPv4MulticastScope.Ttl
					$Table.Cell(5,1).Range.Text = "Lease duration"
					$Table.Cell(5,2).Range.Text = $DurationStr
					$Table.Cell(6,1).Range.Text = "Description"
					$Table.Cell(6,2).Range.Text = $IPv4MulticastScope.Description
					
					$table.AutoFitBehavior($wdAutoFitContent)

					#return focus back to document
					$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

					#move to the end of the current document
					$selection.EndKey($wdStory,$wdMove) | Out-Null
					$TableRange = $Null
					$Table = $Null

					WriteWordLine 4 0 "Lifetime"
					WriteWordLine 0 1 "Multicast scope lifetime: " -NoNewLine
					If([string]::IsNullOrEmpty($IPv4MulticastScope.ExpiryTime))
					{
						WriteWordLine 0 0 "Infinite"
					}
					Else
					{
						WriteWordLine 0 0 "Multicast scope expires on $($IPv4MulticastScope.ExpiryTime)"
					}
					
					Write-Verbose "$(Get-Date): `t`tGetting exclusions"
					WriteWordLine 4 0 "Exclusions"
					$Exclusions = Get-DHCPServerV4MulticastExclusionRange -ComputerName $Script:DHCPServerName -Name $IPv4MulticastScope.Name -EA 0
					If($? -and $Null -ne $Exclusions)
					{
						$TableRange = $doc.Application.Selection.Range
						[int]$Columns = 2
						If($Exclusions -is [array])
						{
							[int]$Rows = $Exclusions.Count + 1
						}
						Else
						{
							[int]$Rows = 2
						}
						ValidateWordTableValues $Rows $Script:DHCPServerName "Function ProcessIPv4MulticastScopes Getting exclusions"
						$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
						$table.Style = $myHash.Word_TableGrid
						$table.Borders.InsideLineStyle = $wdLineStyleNone
						$table.Borders.OutsideLineStyle = $wdLineStyleNone
						$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell(1,1).Range.Font.Bold = $True
						$Table.Cell(1,1).Range.Text = "Start IP Address"
						$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell(1,2).Range.Font.Bold = $True
						$Table.Cell(1,2).Range.Text = "End IP Address"
						[int]$xRow = 1
						ForEach($Exclusion in $Exclusions)
						{
							$xRow++
							$Table.Cell($xRow,1).Range.Text = $Exclusion.StartRange
							$Table.Cell($xRow,2).Range.Text = $Exclusion.EndRange 
						}
						$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
						$table.AutoFitBehavior($wdAutoFitContent)

						#return focus back to document
						$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

						#move to the end of the current document
						$selection.EndKey($wdStory,$wdMove) | Out-Null
						$TableRange = $Null
						$Table = $Null
					}
					ElseIf(!$?)
					{
						WriteWordLine 0 0 "Error retrieving exclusions for multicast scope"
					}
					Else
					{
						WriteWordLine 0 1 "<None>"
					}
					
					#leases
					If($IncludeLeases)
					{
						Write-Verbose "$(Get-Date): `t`tGetting leases"
						
						WriteWordLine 4 0 "Address Leases"
						$Leases = Get-DHCPServerV4MulticastLease -ComputerName $Script:DHCPServerName -Name $IPv4MulticastScope.Name -EA 0 | Sort-Object IPAddress
						If($? -and $Null -ne $Leases)
						{
							$TableRange = $doc.Application.Selection.Range
							[int]$Columns = 2
							If($Leases -is [array])
							{
								[int]$Rows = ($Leases.Count * 7) - 1
								#subtract the very last row used for spacing
							}
							Else
							{
								[int]$Rows = 6
							}
							ValidateWordTableValues $Rows $Script:DHCPServerName "Function ProcessIPv4MulticastScopes Getting leases"
							$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
							$table.Style = $myHash.Word_TableGrid
							$table.Borders.InsideLineStyle = $wdLineStyleNone
							$table.Borders.OutsideLineStyle = $wdLineStyleNone
							[int]$xRow = 0
							ForEach($Lease in $Leases)
							{
								Write-Verbose "$(Get-Date): `t`t`tProcessing lease $($Lease.IPAddress)"
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

								$xRow++
								$Table.Cell($xRow,1).Range.Text = "Client IP address"
								$Table.Cell($xRow,2).Range.Text = $Lease.IPAddress
								
								$xRow++
								$Table.Cell($xRow,1).Range.Text = "Name"
								$Table.Cell($xRow,2).Range.Text = $Lease.HostName
								
								$xRow++
								$Table.Cell($xRow,1).Range.Text = "Lease Expiration"
								If([string]::IsNullOrEmpty($Lease.LeaseExpiryTime))
								{
									$Table.Cell($xRow,2).Range.Text = "Unlimited"
								}
								Else
								{
									$Table.Cell($xRow,2).Range.Text = $LeaseEndStr
								}
								
								$xRow++
								$Table.Cell($xRow,1).Range.Text = "Lease Start"
								If([string]::IsNullOrEmpty($Lease.LeaseStartTime))
								{
									$Table.Cell($xRow,2).Range.Text = "Unlimited"
								}
								Else
								{
									$Table.Cell($xRow,2).Range.Text = $LeaseStartStr
								}
								
								$xRow++
								$Table.Cell($xRow,1).Range.Text = "Address State"
								$Table.Cell($xRow,2).Range.Text = $Lease.AddressState
								
								$xRow++
								$Table.Cell($xRow,1).Range.Text = "MAC address"
								$Table.Cell($xRow,2).Range.Text = $Lease.ClientID
								
								#skip a row for spacing
								$xRow++
							}
							$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
							$table.AutoFitBehavior($wdAutoFitContent)

							#return focus back to document
							$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

							#move to the end of the current document
							$selection.EndKey($wdStory,$wdMove) | Out-Null
							$TableRange = $Null
							$Table = $Null
						}
						ElseIf(!$?)
						{
							WriteWordLine 0 0 "Error retrieving leases for scope"
						}
						Else
						{
							WriteWordLine 0 1 "<None>"
						}
						$Leases = $Null
					}
					
					Write-Verbose "$(Get-Date): `t`tGetting Multicast Scope statistics"
					WriteWordLine 4 0 "Statistics"

					$Statistics = Get-DHCPServerV4MulticastScopeStatistics -ComputerName $Script:DHCPServerName -Name $IPv4MulticastScope.Name -EA 0

					If($? -and $Null -ne $Statistics)
					{
						GetShortStatistics $Statistics
					}
					ElseIf(!$?)
					{
						WriteWordLine 0 0 "Error retrieving multicast scope statistics"
					}
					Else
					{
						WriteWordLine 0 1 "<None>"
					}
					$Statistics = $Null
				}
				ElseIf($Text)
				{
					Write-Verbose "$(Get-Date): `tGetting IPv4 multicast scope data for scope $($IPv4MulticastScope.Name)"
					Line 0 "Multicast Scope [$($IPv4MulticastScope.Name)]"
					Line 1 "General:"
					Line 2 "Name`t`t`t: " $IPv4MulticastScope.Name
					Line 2 "Start IP address`t: " $IPv4MulticastScope.StartRange
					Line 2 "End IP address`t`t: " $IPv4MulticastScope.EndRange
					Line 2 "Time to live`t`t: " $IPv4MulticastScope.Ttl
					Line 2 "Lease duration`t`t: " $DurationStr
					If(![string]::IsNullOrEmpty($IPv4MulticastScope.Description))
					{
						Line 2 "Description`t`t: " $IPv4MulticastScope.Description
					}
					
					Line 1 "Lifetime:"
					Line 2 "Multicast scope lifetime: " -NoNewLine
					If([string]::IsNullOrEmpty($IPv4MulticastScope.ExpiryTime))
					{
						Line 0 "Infinite"
					}
					Else
					{
						Line 0 "Multicast scope expires on $($IPv4MulticastScope.ExpiryTime)"
					}
					
					Write-Verbose "$(Get-Date): `t`tGetting exclusions"
					Line 1 "Exclusions:"
					$Exclusions = Get-DHCPServerV4MulticastExclusionRange -ComputerName $Script:DHCPServerName -Name $IPv4MulticastScope.Name -EA 0
					If($? -and $Null -ne $Exclusions)
					{
						Line 2 "Start IP Address`tEnd IP Address"
						ForEach($Exclusion in $Exclusions)
						{
							Line 2 $Exclusion.StartRange -NoNewLine
							Line 2 $Exclusion.EndRange 
						}
					}
					ElseIf(!$?)
					{
						Line 0 "Error retrieving exclusions for multicast scope"
					}
					Else
					{
						Line 2 "<None>"
					}
					
					#leases
					If($IncludeLeases)
					{
						Write-Verbose "$(Get-Date): `t`tGetting leases"
						
						Line 1 "Address Leases:"
						$Leases = Get-DHCPServerV4MulticastLease -ComputerName $Script:DHCPServerName -Name $IPv4MulticastScope.Name -EA 0 | Sort-Object IPAddress
						If($? -and $Null -ne $Leases)
						{
							ForEach($Lease in $Leases)
							{
								Write-Verbose "$(Get-Date): `t`t`tProcessing lease $($Lease.IPAddress)"
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

								Line 2 "Client IP address`t: " $Lease.IPAddress
								Line 2 "Name`t`t`t: " $Lease.HostName
								Line 2 "Lease Expiration`t: " -NoNewLine
								If([string]::IsNullOrEmpty($Lease.LeaseExpiryTime))
								{
									Line 0 "Unlimited"
								}
								Else
								{
									Line 0 $LeaseEndStr
								}
								
								Line 2 "Lease Start`t`t: " -NoNewLine
								If([string]::IsNullOrEmpty($Lease.LeaseStartTime))
								{
									Line 0 "Unlimited"
								}
								Else
								{
									Line 0 $LeaseStartStr
								}
								
								Line 2 "Address State`t`t: " $Lease.AddressState
								Line 2 "MAC address`t: " $Lease.ClientID
								
								#skip a row for spacing
								Line 0 ""
							}
						}
						ElseIf(!$?)
						{
							Line 0 "Error retrieving leases for scope"
						}
						Else
						{
							Line 2 "<None>"
						}
						$Leases = $Null
					}
					
					Write-Verbose "$(Get-Date): `t`tGetting Multicast Scope statistics"
					Line 1 "Statistics:"

					$Statistics = Get-DHCPServerV4MulticastScopeStatistics -ComputerName $Script:DHCPServerName -Name $IPv4MulticastScope.Name -EA 0

					If($? -and $Null -ne $Statistics)
					{
						GetShortStatistics $Statistics
					}
					ElseIf(!$?)
					{
						Line 0 "Error retrieving multicast scope statistics"
					}
					Else
					{
						Line 2 "<None>"
					}
					$Statistics = $Null
				}
				ElseIf($HTML)
				{
					Write-Verbose "$(Get-Date): `tGetting IPv4 multicast scope data for scope $($IPv4MulticastScope.Name)"
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
					$columnWidths = @("200","100")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -NoBorder $True -tablewidth "300"
					InsertBlankLine

					WriteHTMLLine 4 0 "Lifetime"
					If([string]::IsNullOrEmpty($IPv4MulticastScope.ExpiryTime))
					{
						WriteHTMLLine 0 1 "Multicast scope lifetime: Infinite"
					}
					Else
					{
						WriteHTMLLine 0 1 "Multicast scope lifetime: Multicast scope expires on $($IPv4MulticastScope.ExpiryTime)"
					}
					
					Write-Verbose "$(Get-Date): `t`tGetting exclusions"
					WriteHTMLLine 4 0 "Exclusions"
					$Exclusions = Get-DHCPServerV4MulticastExclusionRange -ComputerName $Script:DHCPServerName -Name $IPv4MulticastScope.Name -EA 0
					If($? -and $Null -ne $Exclusions)
					{
						$rowdata = @()
						$TableRange = $doc.Application.Selection.Range
						ForEach($Exclusion in $Exclusions)
						{
							$rowdata += @(,($Exclusion.StartRange,$htmlwhite,
											$Exclusion.EndRange,$htmlwhite))
						}
						$columnHeaders = @('Start IP Address',($htmlsilver -bor $htmlbold),'End IP Address',($htmlsilver -bor $htmlbold))
						$msg = ""
						FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
						InsertBlankLine
					}
					ElseIf(!$?)
					{
						WriteHTMLLine 0 0 "Error retrieving exclusions for multicast scope"
					}
					Else
					{
						WriteHTMLLine 0 1 "None"
					}
					
					#leases
					If($IncludeLeases)
					{
						Write-Verbose "$(Get-Date): `t`tGetting leases"
						
						WriteHTMLLine 4 0 "Address Leases"
						$Leases = Get-DHCPServerV4MulticastLease -ComputerName $Script:DHCPServerName -Name $IPv4MulticastScope.Name -EA 0 | Sort-Object IPAddress
						If($? -and $Null -ne $Leases)
						{
							ForEach($Lease in $Leases)
							{
								Write-Verbose "$(Get-Date): `t`t`tProcessing lease $($Lease.IPAddress)"
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

								$rowdata = @()
								$columnHeaders = @("Client IP address",($htmlsilver -bor $htmlbold),$Lease.IPAddress.ToString(),$htmlwhite)
								
								$rowdata += @(,('Name',($htmlsilver -bor $htmlbold),$Lease.HostName,$htmlwhite))
								
								$tmp = ""
								If([string]::IsNullOrEmpty($Lease.LeaseExpiryTime))
								{
									$tmp = "Unlimited"
								}
								Else
								{
									$tmp = $LeaseEndStr
								}
								$rowdata += @(,('Lease Expiration',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
								
								$tmp = ""
								If([string]::IsNullOrEmpty($Lease.LeaseStartTime))
								{
									$tmp = "Unlimited"
								}
								Else
								{
									$tmp = $LeaseStartStr
								}
								$rowdata += @(,('Lease Start',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
								$rowdata += @(,('Address State',($htmlsilver -bor $htmlbold),$Lease.AddressState,$htmlwhite))
								$rowdata += @(,('MAC address',($htmlsilver -bor $htmlbold),$Lease.ClientID,$htmlwhite))
								
								$msg = ""
								$columnWidths = @("200","100")
								FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "300"
								InsertBlankLine
							}
						}
						ElseIf(!$?)
						{
							WriteHTMLLine 0 0 "Error retrieving leases for scope"
						}
						Else
						{
							WriteHTMLLine 0 1 "None"
						}
						$Leases = $Null
					}
					
					Write-Verbose "$(Get-Date): `t`tGetting Multicast Scope statistics"
					WriteHTMLLine 4 0 "Statistics"

					$Statistics = Get-DHCPServerV4MulticastScopeStatistics -ComputerName $Script:DHCPServerName -Name $IPv4MulticastScope.Name -EA 0

					If($? -and $Null -ne $Statistics)
					{
						GetShortStatistics $Statistics
					}
					ElseIf(!$?)
					{
						WriteHTMLLine 0 0 "Error retrieving multicast scope statistics"
					}
					Else
					{
						WriteHTMLLine 0 1 "None"
					}
					$Statistics = $Null
				}
			}
		}
		ElseIf(!$?)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "Error retrieving IPv4 Multicast scopes"
			}
			ElseIf($Text)
			{
				Line 0 "Error retrieving IPv4 Multicast scopes"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 "Error retrieving IPv4 Multicast scopes"
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 1 "There were no IPv4 Multicast scopes"
			}
			ElseIf($Text)
			{
				Line 2 "There were no IPv4 Multicast scopes"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 1 "There were no IPv4 Multicast scopes"
			}
		}
		$IPv4MulticastScopes = $Null
	}
	[gc]::collect() 
}

Function ProcessIPv4BOOTPTable
{
	#bootp table
	If($Null -ne $Script:BOOTPTable)
	{
		Write-Verbose "$(Get-Date):IPv4 BOOTP Table"
		
		If($MSWord -or $PDF)
		{
			$selection.InsertNewPage()
			WriteWordLine 3 0 "BOOTP Table"
			
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 3
			If($Script:BOOTPTable -is [array])
			{
				[int]$Rows = $Script:BOOTPTable.Count + 1
			}
			Else
			{
				[int]$Rows = 2
			}
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = $wdLineStyleNone
			$table.Borders.OutsideLineStyle = $wdLineStyleNone
			$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell(1,1).Range.Font.Bold = $True
			$Table.Cell(1,1).Range.Text = "Boot Image"
			$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell(1,2).Range.Font.Bold = $True
			$Table.Cell(1,2).Range.Text = "File Name"
			$Table.Cell(1,3).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell(1,3).Range.Font.Bold = $True
			$Table.Cell(1,3).Range.Text = "File Server"
			[int]$xRow = 1
			ForEach($Item in $Script:BOOTPTable)
			{
				$xRow++
				$ItemParts = $Item.Split(",")
				$Table.Cell($xRow,1).Range.Text = $ItemParts[0]
				$Table.Cell($xRow,2).Range.Text = $ItemParts[1] 
				$Table.Cell($xRow,3).Range.Text = $ItemParts[2] 
			}
			
			$table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
		ElseIf($Text)
		{
			Line 1 "BOOTP Table"
			
			ForEach($Item in $Script:BOOTPTable)
			{
				$ItemParts = $Item.Split(",")
				Line 2 "Boot Image`t: " $ItemParts[0]
				Line 2 "File Name`t: " $ItemParts[1]
				Line 2 "FIle Server`t: " $ItemParts[2] 
				Line 0 ""
			}
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 3 0 "BOOTP Table"
			$rowdata = @()
			
			ForEach($Item in $Script:BOOTPTable)
			{
				$ItemParts = $Item.Split(",")
				$rowdata += @(,($ItemParts[0],$htmlwhite,
								$ItemParts[1],$htmlwhite,
								$ItemParts[2],$htmlwhite))
			}
			$columnHeaders = @('Boot Image',($htmlsilver -bor $htmlbold),'File Name',($htmlsilver -bor $htmlbold),'File Server',($htmlsilver -bor $htmlbold))
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			InsertBlankLine
		}
	}
	[gc]::collect() 
}

Function ProcessServerOptions
{
	#Server Options
	Write-Verbose "$(Get-Date): Getting IPv4 server options"

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 3 0 "Server Options"
	}
	ElseIf($Text)
	{
		Line 1 "Server Options"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "Server Options"
	}

	$ServerOptions = Get-DHCPServerV4OptionValue -All -ComputerName $Script:DHCPServerName -EA 0 | Where-Object {$_.OptionID -ne 81} | Sort-Object OptionId

	If($? -and $Null -ne $ServerOptions)
	{
		If($MSWord -or $PDF)
		{
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If($ServerOptions -is [array])
			{
				[int]$Rows = ($ServerOptions.Count * 5)
			}
			Else
			{
				[int]$Rows = 4
			}
			ValidateWordTableValues $Rows $Script:DHCPServerName "Function ProcessServerOptions Server Options"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = $wdLineStyleNone
			$table.Borders.OutsideLineStyle = $wdLineStyleNone
			[int]$xRow = 0
			ForEach($ServerOption in $ServerOptions)
			{
				Write-Verbose "$(Get-Date): `t`t`tProcessing option name $($ServerOption.Name)"
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Option Name"
				$Table.Cell($xRow,2).Range.Text = "$($ServerOption.OptionId.ToString("000")) $($ServerOption.Name)"
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Vendor"
				If([string]::IsNullOrEmpty($ServerOption.VendorClass))
				{
					$Table.Cell($xRow,2).Range.Text = "Standard"
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $ServerOption.VendorClass
				}
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Value"
				$Table.Cell($xRow,2).Range.Text = $ServerOption.Value[0]

				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Policy Name"
				If([string]::IsNullOrEmpty($ServerOption.PolicyName))
				{
					$Table.Cell($xRow,2).Range.Text = "None"
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $ServerOption.PolicyName
				}
				#for spacing
				$xRow++
			}
			$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
			$table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
		ElseIf($Text)
		{
			ForEach($ServerOption in $ServerOptions)
			{
				Write-Verbose "$(Get-Date): `t`t`tProcessing option name $($ServerOption.Name)"
				Line 2 "Option Name`t: $($ServerOption.OptionId.ToString("000")) $($ServerOption.Name)"
				Line 2 "Vendor`t`t: " -NoNewLine
				If([string]::IsNullOrEmpty($ServerOption.VendorClass))
				{
					Line 0 "Standard"
				}
				Else
				{
					Line 0 $ServerOption.VendorClass
				}
				
				Line 2 "Value`t`t: " $ServerOption.Value[0]
				Line 2 "Policy Name`t: " -NoNewLine
				If([string]::IsNullOrEmpty($ServerOption.PolicyName))
				{
					Line 0 "<None>"
				}
				Else
				{
					Line 0 $ServerOption.PolicyName
				}
				#for spacing
				Line 0 ""
			}
		}
		ElseIf($HTML)
		{
			ForEach($ServerOption in $ServerOptions)
			{
				Write-Verbose "$(Get-Date): `t`t`tProcessing option name $($ServerOption.Name)"
				$rowdata = @()
				$columnHeaders = @("Option Name",($htmlsilver -bor $htmlbold),"$($ServerOption.OptionId.ToString("000")) $($ServerOption.Name)",$htmlwhite)
				
				$tmp = ""
				If([string]::IsNullOrEmpty($ServerOption.VendorClass))
				{
					$tmp = "Standard"
				}
				Else
				{
					$tmp = $ServerOption.VendorClass
				}
				$rowdata += @(,('Vendor',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$rowdata += @(,('Value',($htmlsilver -bor $htmlbold),$ServerOption.Value[0],$htmlwhite))

				$tmp = ""
				If([string]::IsNullOrEmpty($ServerOption.PolicyName))
				{
					$tmp = "None"
				}
				Else
				{
					$tmp = $ServerOption.PolicyName
				}
				$rowdata += @(,('Policy Name',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))

				$msg = ""
				$columnWidths = @("150","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
				InsertBlankLine
			}
		}
	}
	ElseIf($? -and $Null -eq $ServerOptions)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no IPv4 server options"
		}
		ElseIf($Text)
		{
			Line 2 "There were no IPv4 server options"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "There were no IPv4 server options"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv4 server options"
		}
		ElseIf($Text)
		{
			Line 0 "Error retrieving IPv4 server options"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv4 server options"
		}
	}
	$ServerOptions = $Null
	[gc]::collect() 
}

Function ProcessPolicies
{
	#Policies
	Write-Verbose "$(Get-Date): Getting IPv4 policies"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Policies"
	}
	ElseIf($Text)
	{
		Line 1 "Policies"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "Policies"
	}

	$Policies = Get-DHCPServerV4Policy -ComputerName $Script:DHCPServerName -EA 0 | Sort-Object ProcessingOrder

	If($? -and $Null -ne $Policies)
	{
		If($MSWord -or $PDF)
		{
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If($Policies -is [array])
			{
				[int]$Rows = $Policies.Count * 6
			}
			Else
			{
				[int]$Rows = 5
			}
			ValidateWordTableValues $Rows $Script:DHCPServerName "Function ProcessPolicies Policies"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = $wdLineStyleNone
			$table.Borders.OutsideLineStyle = $wdLineStyleNone
			[int]$xRow = 0
			ForEach($Policy in $Policies)
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Policy Name"
				$Table.Cell($xRow,2).Range.Text = $Policy.Name
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Description"
				$Table.Cell($xRow,2).Range.Text = $Policy.Description

				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Processing Order"
				$Table.Cell($xRow,2).Range.Text = $Policy.ProcessingOrder

				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Level"
				$Table.Cell($xRow,2).Range.Text = "Server"

				$xRow++
				$Table.Cell($xRow,1).Range.Text = "State"
				If($Policy.Enabled)
				{
					$Table.Cell($xRow,2).Range.Text = "Enabled"
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = "Disabled"
				}
				#for spacing
				$xRow++
				$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
				$table.AutoFitBehavior($wdAutoFitContent)

				#return focus back to document
				$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

				#move to the end of the current document
				$selection.EndKey($wdStory,$wdMove) | Out-Null
				$TableRange = $Null
				$Table = $Null
			}
		}
		ElseIf($Text)
		{
			ForEach($Policy in $Policies)
			{
				Line 2 "Policy Name`t`t: " $Policy.Name
				If(![string]::IsNullOrEmpty($Policy.Description))
				{
					Line 2 "Description`t`t: " $Policy.Description
				}
				Line 2 "Processing Order`t: " $Policy.ProcessingOrder
				Line 2 "Level`t`t`t: Server"
				Line 2 "State`t`t`t: " -NoNewLine
				If($Policy.Enabled)
				{
					Line 0 "Enabled"
				}
				Else
				{
					Line 0 "Disabled"
				}
				#for spacing
				Line 0 ""
			}
		}
		ElseIf($HTML)
		{
			ForEach($Policy in $Policies)
			{
				$rowdata = @()
				$columnHeaders = @("Policy Name",($htmlsilver -bor $htmlbold),$Policy.Name,$htmlwhite)
				$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Policy.Description,$htmlwhite))
				$rowdata += @(,('Processing Order',($htmlsilver -bor $htmlbold),$Policy.ProcessingOrder,$htmlwhite))
				$rowdata += @(,('Level',($htmlsilver -bor $htmlbold),"Server",$htmlwhite))

				$tmp = ""
				If($Policy.Enabled)
				{
					$tmp = "Enabled"
				}
				Else
				{
					$tmp = "Disabled"
				}
				$rowdata += @(,('State',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))

				$msg = ""
				$columnWidths = @("150","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
				InsertBlankLine
			}	
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv4 policies"
		}
		ElseIf($Text)
		{
			Line 0 "Error retrieving IPv4 policies"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv4 policies"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no IPv4 policies"
		}
		ElseIf($Text)
		{
			Line 2 "There were no IPv4 policies"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "There were no IPv4 policies"
		}
	}
	$Policies = $Null
	[gc]::collect() 
}

Function ProcessIPv4Filters
{
	#Filters
	Write-Verbose "$(Get-Date): Getting IPv4 filters"
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Filters"
	}
	ElseIf($Text)
	{
		Line 1 "Filters"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "Filters"
	}

	Write-Verbose "$(Get-Date): `tAllow filters"
	$AllowFilters = Get-DHCPServerV4Filter -List Allow -ComputerName $Script:DHCPServerName -EA 0 | Sort-Object MacAddress

	If($? -and $Null -ne $AllowFilters)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 4 0 "Allow"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If($AllowFilters -is [array])
			{
				[int]$Rows = $AllowFilters.Count + 1
			}
			Else
			{
				[int]$Rows = 2
			}
			ValidateWordTableValues $Rows $Script:DHCPServerName "Function ProcessIPv4Filters Allow filters"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = $wdLineStyleNone
			$table.Borders.OutsideLineStyle = $wdLineStyleNone
			$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell(1,1).Range.Font.Bold = $True
			$Table.Cell(1,1).Range.Text = "MAC Address"
			$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell(1,2).Range.Font.Bold = $True
			$Table.Cell(1,2).Range.Text = "Description"
			[int]$xRow = 1
			ForEach($AllowFilter in $AllowFilters)
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = $AllowFilter.MacAddress
				$Table.Cell($xRow,2).Range.Text = $AllowFilter.Description
			}
			$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
			$table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
		ElseIf($Text)
		{
			Line 2 "Allow"
			ForEach($AllowFilter in $AllowFilters)
			{
				Line 3 "MAC Address`t: " $AllowFilter.MacAddress
				Line 3 "Description`t: " $AllowFilter.Description
				Line 0 ""
			}
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 4 0 "Allow"
			$rowdata = @()
			ForEach($AllowFilter in $AllowFilters)
			{
				$rowdata += @(,($AllowFilter.MacAddress,$htmlwhite,
								$AllowFilter.Description,$htmlwhite))
			}
			$columnHeaders = @('MAC Address',($htmlsilver -bor $htmlbold),'Description',($htmlsilver -bor $htmlbold))
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			InsertBlankLine
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv4 allow filters"
		}
		ElseIf($Text)
		{
			Line 0 "Error retrieving IPv4 allow filters"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv4 allow filters"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no IPv4 allow filters"
		}
		ElseIf($Text)
		{
			Line 2 "There were no IPv4 allow filters"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "There were no IPv4 allow filters"
		}
	}
	$AllowFilters = $Null
	[gc]::collect() 

	Write-Verbose "$(Get-Date): `tDeny filters"
	$DenyFilters = Get-DHCPServerV4Filter -List Deny -ComputerName $Script:DHCPServerName -EA 0 | Sort-Object MacAddress
	If($? -and $Null -ne $DenyFilters)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 4 0 "Deny"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If($DenyFilters -is [array])
			{
				[int]$Rows = $DenyFilters.Count + 1
			}
			Else
			{
				[int]$Rows = 2
			}
			ValidateWordTableValues $Rows $Script:DHCPServerName "Function ProcessIPv4Filters Deny filters"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = $wdLineStyleNone
			$table.Borders.OutsideLineStyle = $wdLineStyleNone
			$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell(1,1).Range.Font.Bold = $True
			$Table.Cell(1,1).Range.Text = "MAC Address"
			$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell(1,2).Range.Font.Bold = $True
			$Table.Cell(1,2).Range.Text = "Description"
			[int]$xRow = 1

			ForEach($DenyFilter in $DenyFilters)
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = $DenyFilter.MacAddress
				$Table.Cell($xRow,2).Range.Text = $DenyFilter.Description
			}
			$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
			$table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
		ElseIf($Text)
		{
			Line 2 "Deny"
			ForEach($DenyFilter in $DenyFilters)
			{
				Line 3 "MAC Address`t: " $DenyFilter.MacAddress
				Line 3 "Description`t: " $DenyFilter.Description
				Line 0 ""
			}
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 4 0 "Deny"
			$rowdata = @()
			ForEach($DenyFilter in $DenyFilters)
			{
				$rowdata += @(,($DenyFilter.MacAddress,$htmlwhite,
								$DenyFilter.Description,$htmlwhite))
			}
			$columnHeaders = @('MAC Address',($htmlsilver -bor $htmlbold),'Description',($htmlsilver -bor $htmlbold))
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			InsertBlankLine
		}
	}
	ElseIf(!$?)
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 "Error retrieving IPv4 deny filters"
		}
		ElseIf($Text)
		{
			Line 0 "Error retrieving IPv4 deny filters"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv4 deny filters"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no IPv4 deny filters"
		}
		ElseIf($Text)
		{
			Line 2 "There were no IPv4 deny filters"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "There were no IPv4 deny filters"
		}
	}
	$DenyFilters = $Null
	[gc]::collect() 
}

Function ProcessIPv6Properties
{
	#IPv6

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 2 0 "IPv6"
		WriteWordLine 3 0 "Properties"

		Write-Verbose "$(Get-Date): Getting IPv6 properties"
		Write-Verbose "$(Get-Date): `tGetting IPv6 general settings"
		WriteWordLine 4 0 "General"

		If($Script:GotAuditSettings)
		{
			If($Script:AuditSettings.Enable)
			{
				WriteWordLine 0 1 "DHCP audit logging is enabled"
			}
			Else
			{
				WriteWordLine 0 1 "DHCP audit logging is disabled"
			}
		}
		ElseIf(!$?)
		{
			WriteWordLine 0 0 "Error retrieving audit log settings"
		}
		Else
		{
			WriteWordLine 0 1 "There were no audit log settings"
		}

		#DNS settings
		Write-Verbose "$(Get-Date): `tGetting IPv6 DNS settings"
		WriteWordLine 4 0 "DNS"
		$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $Script:DHCPServerName -EA 0
		If($? -and $Null -ne $DNSSettings)
		{
			GetDNSSettings $DNSSettings "AAAA"
		}
		Else
		{
			WriteWordLine 0 0 "Error retrieving IPv6 DNS Settings for DHCP server $Script:DHCPServerName"
		}
		$DNSSettings = $Null

		#Advanced
		Write-Verbose "$(Get-Date): `tGetting IPv6 advanced settings"
		WriteWordLine 4 0 "Advanced"
		If($Script:GotAuditSettings)
		{
			WriteWordLine 0 1 "Audit log file path " $Script:AuditSettings.Path
		}
		$Script:AuditSettings = $Null

		#added 18-Jan-2016
		#get dns update credentials
		Write-Verbose "$(Get-Date): `tGetting DNS dynamic update registration credentials"
		$DNSUpdateSettings = Get-DhcpServerDnsCredential -ComputerName $Script:DHCPServerName -EA 0

		If($? -and $Null -ne $DNSUpdateSettings)
		{
			WriteWordLine 0 1 "DNS dynamic update registration credentials: "
			WriteWordLine 0 2 "User name: " $DNSUpdateSettings.UserName
			WriteWordLine 0 2 "Domain: " $DNSUpdateSettings.DomainName
		}
		ElseIf(!$?)
		{
			WriteWordLine 0 0 "Error retrieving DNS Update Credentials for DHCP server $Script:DHCPServerName"
		}
		Else
		{
			WriteWordLine 0 1 "There were no DNS Update Credentials for DHCP server $Script:DHCPServerName"
		}
		$DNSUpdateSettings = $Null
		[gc]::collect() 
		
		WriteWordLine 4 0 "Statistics"
		$Statistics = Get-DHCPServerV6Statistics -ComputerName $Script:DHCPServerName -EA 0

		If($? -and $Null -ne $Statistics)
		{
			$UpTime = $(Get-Date) - $Statistics.ServerStartTime
			$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3} seconds", `
				$UpTime.Days, `
				$UpTime.Hours, `
				$UpTime.Minutes, `
				$UpTime.Seconds)

			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			[int]$Rows = 16
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = $wdLineStyleSingle
			$table.Borders.OutsideLineStyle = $wdLineStyleSingle
			$Table.Cell(1,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell(1,1).Range.Font.Bold = $True
			$Table.Cell(1,1).Range.Text = "Description"
			$Table.Cell(1,2).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell(1,2).Range.Font.Bold = $True
			$Table.Cell(1,2).Range.Text = "Details"

			$Table.Cell(2,1).Range.Text = "Start Time"
			$Table.Cell(2,2).Range.Text = $Statistics.ServerStartTime.ToString()
			$Table.Cell(3,1).Range.Text = "Up Time"
			$Table.Cell(3,2).Range.Text = $Str
			$Table.Cell(4,1).Range.Text = "Solicits"
			$Table.Cell(4,2).Range.Text = $Statistics.Solicits.ToString()
			$Table.Cell(5,1).Range.Text = "Advertises"
			$Table.Cell(5,2).Range.Text = $Statistics.Advertises.ToString()
			$Table.Cell(6,1).Range.Text = "Requests"
			$Table.Cell(6,2).Range.Text = $Statistics.Requests.ToString()
			$Table.Cell(7,1).Range.Text = "Replies"
			$Table.Cell(7,2).Range.Text = $Statistics.Replies.ToString()
			$Table.Cell(8,1).Range.Text = "Renews"
			$Table.Cell(8,2).Range.Text = $Statistics.Renews.ToString()
			$Table.Cell(9,1).Range.Text = "Rebinds"
			$Table.Cell(9,2).Range.Text = $Statistics.Rebinds.ToString()
			$Table.Cell(10,1).Range.Text = "Confirms"
			$Table.Cell(10,2).Range.Text = $Statistics.Confirms.ToString()
			$Table.Cell(11,1).Range.Text = "Declines"
			$Table.Cell(11,2).Range.Text = $Statistics.Declines.ToString()
			$Table.Cell(12,1).Range.Text = "Releases"
			$Table.Cell(12,2).Range.Text = $Statistics.Releases.ToString()
			$Table.Cell(13,1).Range.Text = "Total Scopes"
			$Table.Cell(13,2).Range.Text = $Statistics.TotalScopes.ToString()
			$Table.Cell(14,1).Range.Text = "Total Addresses"
			$tmp = "{0:N0}" -f $Statistics.TotalAddresses
			$Table.Cell(14,2).Range.Text = $tmp
			$Table.Cell(15,1).Range.Text = "In Use"
			[int]$InUsePercent = "{0:N0}" -f $Statistics.PercentageInUse.ToString()
			$Table.Cell(15,2).Range.Text = "$($Statistics.AddressesInUse) ($($InUsePercent)%)"
			$Table.Cell(16,1).Range.Text = "Available"
			[int]$AvailablePercent = "{0:N0}" -f $Statistics.PercentageAvailable.ToString()
			$tmp = "{0:N0}" -f $Statistics.AddressesAvailable
			$Table.Cell(16,2).Range.Text = "$($tmp) ($($AvailablePercent)%)"

			$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustNone)
			$table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
		ElseIf(!$?)
		{
			WriteWordLine 0 0 "Error retrieving IPv6 statistics"
		}
		Else
		{
			WriteWordLine 0 0 "There were no IPv6 statistics"
		}
		$Statistics = $Null

		Write-Verbose "$(Get-Date): Getting IPv6 scopes"
		$IPv6Scopes = Get-DHCPServerV6Scope -ComputerName $Script:DHCPServerName -EA 0

		If($? -and $Null -ne $IPv6Scopes)
		{
			$selection.InsertNewPage()
			ForEach($IPv6Scope in $IPv6Scopes)
			{
				GetIPv6ScopeData $IPv6Scope
			}
		}
		ElseIf(!$?)
		{
			WriteWordLine 0 0 "Error retrieving IPv6 scopes"
		}
		Else
		{
			WriteWordLine 0 1 "There were no IPv6 scopes"
		}
		$IPv6Scopes = $Null

		Write-Verbose "$(Get-Date): Getting IPv6 server options"
		$selection.InsertNewPage()
		WriteWordLine 3 0 "Server Options"

		$ServerOptions = Get-DHCPServerV6OptionValue -All -ComputerName $Script:DHCPServerName -EA 0 | Sort-Object OptionId

		If($? -and $Null -ne $ServerOptions)
		{
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			If($ServerOptions -is [array])
			{
				[int]$Rows = $ServerOptions.Count * 4
			}
			Else
			{
				[int]$Rows = 3
			}
			ValidateWordTableValues $Rows $Script:DHCPServerName "Function ProcessIPv6Properties Server Options"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = $wdLineStyleNone
			$table.Borders.OutsideLineStyle = $wdLineStyleNone
			[int]$xRow = 0
			ForEach($ServerOption in $ServerOptions)
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Option Name"
				$Table.Cell($xRow,2).Range.Text = "$($ServerOption.OptionId.ToString("00000")) $($ServerOption.Name)"
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Vendor"
				If([string]::IsNullOrEmpty($ServerOption.VendorClass))
				{
					$Table.Cell($xRow,2).Range.Text =  "Standard"
				}
				Else
				{
					$Table.Cell($xRow,2).Range.Text = $ServerOption.VendorClass
				}
				
				$xRow++
				$Table.Cell($xRow,1).Range.Text = "Value"
				$Table.Cell($xRow,2).Range.Text = $ServerOption.Value
				
				#for spacing
				$xRow++
			}
			$table.AutoFitBehavior($wdAutoFitContent)

			#return focus back to document
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$TableRange = $Null
			$Table = $Null
		}
		ElseIf(!$?)
		{
			WriteWordLine 0 0 "Error retrieving IPv6 server options"
		}
		Else
		{
			WriteWordLine 0 1 "There were no IPv6 server options"
		}
		$ServerOptions = $Null
	}
	ElseIf($Text)
	{
		Line 0 "IPv6"
		Line 0 "Properties"

		Write-Verbose "$(Get-Date): Getting IPv6 properties"
		Write-Verbose "$(Get-Date): `tGetting IPv6 general settings"
		Line 1 "General"

		If($Script:GotAuditSettings)
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
		ElseIf(!$?)
		{
			Line 0 "Error retrieving audit log settings"
		}
		Else
		{
			Line 2 "There were no audit log settings"
		}

		#DNS settings
		Write-Verbose "$(Get-Date): `tGetting IPv6 DNS settings"
		Line 1 "DNS"
		$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $Script:DHCPServerName -EA 0
		If($? -and $Null -ne $DNSSettings)
		{
			GetDNSSettings $DNSSettings "AAAA"
		}
		Else
		{
			Line 0 "Error retrieving IPv6 DNS Settings for DHCP server $Script:DHCPServerName"
		}
		$DNSSettings = $Null

		#Advanced
		Write-Verbose "$(Get-Date): `tGetting IPv6 advanced settings"
		Line 1 "Advanced"
		If($Script:GotAuditSettings)
		{
			Line 2 "Audit log file path " $Script:AuditSettings.Path
		}
		$Script:AuditSettings = $Null

		#added 18-Jan-2016
		#get dns update credentials
		Write-Verbose "$(Get-Date): `tGetting DNS dynamic update registration credentials"
		$DNSUpdateSettings = Get-DhcpServerDnsCredential -ComputerName $Script:DHCPServerName -EA 0

		If($? -and $Null -ne $DNSUpdateSettings)
		{
			Line 2 "DNS dynamic update registration credentials: "
			Line 3 "User name`t: " $DNSUpdateSettings.UserName
			Line 3 "Domain`t`t: " $DNSUpdateSettings.DomainName
		}
		ElseIf(!$?)
		{
			Line 0 "Error retrieving DNS Update Credentials for DHCP server $Script:DHCPServerName"
		}
		Else
		{
			Line 2 "There were no DNS Update Credentials for DHCP server $Script:DHCPServerName"
		}
		$DNSUpdateSettings = $Null
		[gc]::collect() 
		
		Line 1 "Statistics"
		$Statistics = Get-DHCPServerV6Statistics -ComputerName $Script:DHCPServerName -EA 0

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

			Line 2 "Description" -NoNewLine
			Line 2 "Details"

			Line 2 "Start Time: " -NoNewLine
			Line 2 $Statistics.ServerStartTime
			Line 2 "Up Time: " -NoNewLine
			Line 2 $Str
			Line 2 "Solicits: " -NoNewLine
			Line 2 $Statistics.Solicits
			Line 2 "Advertises: " -NoNewLine
			Line 2 $Statistics.Advertises
			Line 2 "Requests: " -NoNewLine
			Line 2 $Statistics.Requests
			Line 2 "Replies: " -NoNewLine
			Line 2 $Statistics.Replies
			Line 2 "Renews: " -NoNewLine
			Line 2 $Statistics.Renews
			Line 2 "Rebinds: " -NoNewLine
			Line 2 $Statistics.Rebinds
			Line 2 "Confirms: " -NoNewLine
			Line 2 $Statistics.Confirms
			Line 2 "Declines: " -NoNewLine
			Line 2 $Statistics.Declines
			Line 2 "Releases: " -NoNewLine
			Line 2 $Statistics.Releases
			Line 2 "Total Scopes: " -NoNewLine
			Line 2 $Statistics.TotalScopes
			Line 2 "Total Addresses: " -NoNewLine
			$tmp = "{0:N0}" -f $Statistics.TotalAddresses
			Line 1 $tmp
			Line 2 "In Use: " -NoNewLine
			Line 2 "$($Statistics.AddressesInUse) ($($InUsePercent)%)"
			Line 2 "Available: " -NoNewLine
			$tmp = "{0:N0}" -f $Statistics.AddressesAvailable 
			Line 2 "$($tmp) ($($AvailablePercent)%)"
		}
		ElseIf(!$?)
		{
			Line 0 "Error retrieving IPv6 statistics"
		}
		Else
		{
			Line 0 "There were no IPv6 statistics"
		}
		
		$Statistics = $Null

		Write-Verbose "$(Get-Date): Getting IPv6 scopes"
		$IPv6Scopes = Get-DHCPServerV6Scope -ComputerName $Script:DHCPServerName -EA 0

		If($? -and $Null -ne $IPv6Scopes)
		{
			ForEach($IPv6Scope in $IPv6Scopes)
			{
				GetIPv6ScopeData $IPv6Scope
			}
		}
		ElseIf(!$?)
		{
			Line 0 "Error retrieving IPv6 scopes"
		}
		Else
		{
			Line 1 "There were no IPv6 scopes"
		}
		$IPv6Scopes = $Null

		Write-Verbose "$(Get-Date): Getting IPv6 server options"
		Line 0 "Server Options"

		$ServerOptions = Get-DHCPServerV6OptionValue -All -ComputerName $Script:DHCPServerName -EA 0 | Sort-Object OptionId

		If($? -and $Null -ne $ServerOptions)
		{
			ForEach($ServerOption in $ServerOptions)
			{
				Line 1 "Option Name`t: $($ServerOption.OptionId.ToString("00000")) $($ServerOption.Name)"
				Line 1 "Vendor`t`t: " -NoNewLine
				If([string]::IsNullOrEmpty($ServerOption.VendorClass))
				{
					Line 0 "Standard"
				}
				Else
				{
					Line 0 $ServerOption.VendorClass
				}
				Line 1 "Value`t`t: " $ServerOption.Value
				
				#for spacing
				Line 0 ""
			}
		}
		ElseIf(!$?)
		{
			Line 0 "Error retrieving IPv6 server options"
		}
		Else
		{
			Line 2 "There were no IPv6 server options"
		}
		$ServerOptions = $Null
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 2 0 "IPv6"
		WriteHTMLLine 3 0 "Properties"

		Write-Verbose "$(Get-Date): Getting IPv6 properties"
		Write-Verbose "$(Get-Date): `tGetting IPv6 general settings"
		WriteHTMLLine 4 0 "General"

		If($Script:GotAuditSettings)
		{
			If($Script:AuditSettings.Enable)
			{
				WriteHTMLLine 0 1 "DHCP audit logging is enabled"
			}
			Else
			{
				WriteHTMLLine 0 1 "DHCP audit logging is disabled"
			}
		}
		ElseIf(!$?)
		{
			WriteHTMLLine 0 0 "Error retrieving audit log settings"
		}
		Else
		{
			WriteHTMLLine 0 1 "There were no audit log settings"
		}

		#DNS settings
		Write-Verbose "$(Get-Date): `tGetting IPv6 DNS settings"
		WriteHTMLLine 4 0 "DNS"
		$DNSSettings = Get-DHCPServerV6DnsSetting -ComputerName $Script:DHCPServerName -EA 0
		If($? -and $Null -ne $DNSSettings)
		{
			GetDNSSettings $DNSSettings "AAAA"
		}
		Else
		{
			WriteHTMLLine 0 0 "Error retrieving IPv6 DNS Settings for DHCP server $Script:DHCPServerName"
		}
		$DNSSettings = $Null

		#Advanced
		Write-Verbose "$(Get-Date): `tGetting IPv6 advanced settings"
		WriteHTMLLine 4 0 "Advanced"
		If($Script:GotAuditSettings)
		{
			WriteHTMLLine 0 1 "Audit log file path " $Script:AuditSettings.Path
		}
		$Script:AuditSettings = $Null

		#added 18-Jan-2016
		#get dns update credentials
		Write-Verbose "$(Get-Date): `tGetting DNS dynamic update registration credentials"
		$DNSUpdateSettings = Get-DhcpServerDnsCredential -ComputerName $Script:DHCPServerName -EA 0

		If($? -and $Null -ne $DNSUpdateSettings)
		{
			WriteHTMLLine 0 1 "DNS dynamic update registration credentials: "
			WriteHTMLLine 0 2 "User name: " $DNSUpdateSettings.UserName
			WriteHTMLLine 0 2 "Domain: " $DNSUpdateSettings.DomainName
		}
		ElseIf(!$?)
		{
			WriteHTMLLine 0 0 "Error retrieving DNS Update Credentials for DHCP server $Script:DHCPServerName"
		}
		Else
		{
			WriteHTMLLine 0 1 "There were no DNS Update Credentials for DHCP server $Script:DHCPServerName"
		}
		$DNSUpdateSettings = $Null
		[gc]::collect() 
		
		WriteHTMLLine 4 0 "Statistics"
		$Statistics = Get-DHCPServerV6Statistics -ComputerName $Script:DHCPServerName -EA 0

		If($? -and $Null -ne $Statistics)
		{
			$UpTime = $(Get-Date) - $Statistics.ServerStartTime
			$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3} seconds", `
				$UpTime.Days, `
				$UpTime.Hours, `
				$UpTime.Minutes, `
				$UpTime.Seconds)

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
			$tmp = "{0:N0}" -f $Statistics.TotalAddresses.ToString()
			$rowdata += @(,("Total Addresses",$htmlwhite,$tmp,$htmlwhite))
			$rowdata += @(,("In Use",$htmlwhite,"$($Statistics.AddressesInUse) ($($InUsePercent))%",$htmlwhite))
			$tmp = "{0:N0}" -f "$($Statistics.AddressesAvailable) ($($AvailablePercent))%"
			$rowdata += @(,("Available",$htmlwhite,$tmp,$htmlwhite))

			$columnHeaders = @('Description',($htmlsilver -bor $htmlbold),'Details',($htmlsilver -bor $htmlbold))
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			InsertBlankLine
		}
		ElseIf(!$?)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv6 statistics"
		}
		Else
		{
			WriteHTMLLine 0 0 "There were no IPv6 statistics"
		}
		$Statistics = $Null

		Write-Verbose "$(Get-Date): Getting IPv6 scopes"
		$IPv6Scopes = Get-DHCPServerV6Scope -ComputerName $Script:DHCPServerName -EA 0

		If($? -and $Null -ne $IPv6Scopes)
		{
			$selection.InsertNewPage()
			ForEach($IPv6Scope in $IPv6Scopes)
			{
				GetIPv6ScopeData $IPv6Scope
			}
		}
		ElseIf(!$?)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv6 scopes"
		}
		Else
		{
			WriteHTMLLine 0 1 "There were no IPv6 scopes"
		}
		$IPv6Scopes = $Null

		Write-Verbose "$(Get-Date): Getting IPv6 server options"
		WriteHTMLLine 3 0 "Server Options"

		$ServerOptions = Get-DHCPServerV6OptionValue -All -ComputerName $Script:DHCPServerName -EA 0 | Sort-Object OptionId

		If($? -and $Null -ne $ServerOptions)
		{
			ForEach($ServerOption in $ServerOptions)
			{
				$rowdata = @()
				$columnHeaders = @("Option Name",($htmlsilver -bor $htmlbold),"$($ServerOption.OptionId.ToString("000")) $($ServerOption.Name)",$htmlwhite)
				
				$tmp = ""
				If([string]::IsNullOrEmpty($ServerOption.VendorClass))
				{
					$tmp = "Standard"
				}
				Else
				{
					$tmp = $ServerOption.VendorClass
				}
				$rowdata += @(,('Vendor',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$rowdata += @(,('Value',($htmlsilver -bor $htmlbold),$ServerOption.Value[0],$htmlwhite))

				$msg = ""
				$columnWidths = @("150","200")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "350"
				InsertBlankLine
			}
		}
		ElseIf(!$?)
		{
			WriteHTMLLine 0 0 "Error retrieving IPv6 server options"
		}
		Else
		{
			WriteHTMLLine 0 1 "There were no IPv6 server options"
		}
		$ServerOptions = $Null
	}
}

Function ProcessDHCPOptions
{
	Write-Verbose "$(Get-Date): Getting DHCP Options"
	
	$DHCPOptions = Get-DhcpServerV4OptionDefinition -ComputerName $Script:DHCPServerName -EA 0
	
	If($? -or $Null -ne $DHCPOptions)
	{
		Write-Verbose "$(Get-Date): `tProcessing DHCP Options"
		$DHCPOptions = $DHCPOptions | Sort-Object OptionId
	
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ItemsWordTable = @();
			$Selection.InsertNewPage()
			WriteWordLine 2 0 "DHCP Options"
		}
		ElseIf($Text)
		{
			Line 0 "DHCP Options"
			Line 0 ""
			Line 1 "OptionId  Name                                                Description                                                   Type          Vendor Class  Default Value         Multivalued"
			Line 1 "========================================================================================================================================================================================="
			       #12345678SS12345678901234567890123456789012345678901234567890SS123456789012345678901234567890123456789012345678901234567890SS123456789012SS123456789012SS12345678901234567890SS12345
		}
		ElseIf($HTML)
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
			ElseIf($Text)
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
			ElseIf($HTML)
			{
				#$tmpDV = $Item.DefaultValue.SyncRoot
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
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns OptionId, Name, Description, Type, VendorClass, DefaultValue, MultiValued `
			-Headers  "OptionId", "Name", "Description", "Type", "Vendor Class", "Default Value", "Multivalued" `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			SetWordCellFormat -Collection $Table -Size 9

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
		ElseIf($HTML)
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
		ElseIf($Text)
		{
			Line 0 "Error retrieving DHCP Options"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 0 "Error retrieving DHCP Options"
		}
	}
	Else
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "There were no DHCP Options"
		}
		ElseIf($Text)
		{
			Line 0 "There were no DHCP Options"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "There were no DHCP Options"
		}
	}
}

Function ProcessHardware
{
	#V1.40 added
	If($Hardware)
	{
		Write-Verbose "$(Get-Date): Processing Hardware Information"
		If($MSWord -or $PDF)
		{
			$Script:Selection.InsertNewPage()
		}
		GetComputerWMIInfo $Script:DHCPServerName
	}
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
		Write-Verbose "$(Get-Date): Resolving computer name"
		$ComputerName = TestComputerName $ComputerName
		$Script:DHCPServerNames += $ComputerName
	}
	Else
	{
		Write-Verbose "$(Get-Date): Retrieving all DHCP servers in domain"
		$ComputerName = "All DHCP Servers"
		
		$ALLServers = Get-DHCPServerInDc -EA 0
		
		If($Null -eq $AllServers)
		{
			#oops no DHCP servers
			Write-Error "Unable to retrieve any DHCP servers.  Script cannot continue"
			Exit
		}
		Else
		{
			[int]$cnt = 0
			If($AllServers -is [array])
			{
				$cnt = $AllServers.Count
				Write-Verbose "$(Get-Date): $($cnt) DHCP servers were found"
			}
			Else
			{
				$cnt = 1
				Write-Verbose "$(Get-Date): $($cnt) DHCP server was found"
			}
			
			$Script:BadDHCPErrorFile = "$($Script:pwdpath)\BadDHCPServers_$(Get-Date -f yyyy-MM-dd_HHmm).txt"

			ForEach($Server in $AllServers)
			{
				$Result = TestComputerName2 $Server.DnsName
				
				If($Result -ne "BAD")
				{
					$Script:DHCPServerNames += $Result
				}
			}
			Write-Verbose "$(Get-Date): $($Script:DHCPServerNames.Count) DHCP servers will be processed"
			Write-Verbose "$(Get-Date): "
		}
	}
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
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
		$runtime.Days, `
		$runtime.Hours, `
		$runtime.Minutes, `
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
		$SIFile = "$($Script:pwdpath)\DHCPInventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime       : $($AddDateTime)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name       : $($Script:CoName)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Address    : $($CompanyAddress)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Email      : $($CompanyEmail)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Fax        : $($CompanyFax)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Company Phone      : $($CompanyPhone)" 4>$Null		
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page         : $($CoverPage)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "ComputerName       : $($ComputerName)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Dev                : $($Dev)" 4>$Null
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile       : $($Script:DevErrorFile)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Filename1          : $($Script:FileName1)" 4>$Null
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Filename2          : $($Script:FileName2)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder             : $($Folder)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "From               : $($From)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "HW Inventory       : $($Hardware)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Include Leases     : $($IncludeLeases)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Include Options    : $($IncludeOptions)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Log                : $($Log)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML       : $($HTML)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF        : $($PDF)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT       : $($TEXT)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD       : $($MSWORD)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script Info        : $($ScriptInfo)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Port          : $($SmtpPort)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Smtp Server        : $($SmtpServer)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Title              : $($Script:Title)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "To                 : $($To)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Use SSL            : $($UseSSL)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "User Name          : $($UserName)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected        : $($Script:RunningOS)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version       : $($Host.Version)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture          : $($PSCulture)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture        : $($PSUICulture)" 4>$Null
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word language      : $($Script:WordLanguageValue)" 4>$Null
			Out-File -FilePath $SIFile -Append -InputObject "Word version       : $($Script:WordProduct)" 4>$Null
		}
		Out-File -FilePath $SIFile -Append -InputObject "" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Script start       : $($Script:StartTime)" 4>$Null
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time       : $($Str)" 4>$Null
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
				Write-Verbose "$(Get-Date): $Script:LogPath is ready for use"
			} 
			catch 
			{
				Write-Verbose "$(Get-Date): Transcript/log stop failed"
			}
		}
	}

	$runtime = $Null
	$Str = $Null
	$ErrorActionPreference = $SaveEAPreference
			
	Write-Host "                                                                                    " -BackgroundColor Black -ForegroundColor White
	Write-Host "               This FREE script was brought to you by Conversant Group              " -BackgroundColor Black -ForegroundColor White
	Write-Host "We design, build, and manage infrastructure for a secure, dependable user experience" -BackgroundColor Black -ForegroundColor White
	Write-Host "                       Visit our website conversantgroup.com                        " -BackgroundColor Black -ForegroundColor White
	Write-Host "                                                                                    " -BackgroundColor Black -ForegroundColor White
}
#endregion

#region script core
#Script begins

ProcessScriptSetup

If($AllDHCPServers -eq $False)
{
	[string]$Script:Title = "DHCP Inventory Report for Server $($Script:DHCPServerNames[0])"
	SetFileName1andFileName2 "$($Script:DHCPServerNames[0])_DHCP_Inventory"
}
Else
{
	[string]$Script:Title = "DHCP Inventory Report for All DHCP Servers"
	SetFileName1andFileName2 "DHCP Inventory for All DHCP Servers"
}

ForEach($DHCPServer in $Script:DHCPServerNames)
{
	Write-Verbose "$(Get-Date): Processing DHCP Server: $($DHCPServer)"
	$Script:DHCPServerName = $DHCPServer
	
	ProcessServerProperties

	ProcessIPv4Bindings

	ProcessIPv6Bindings

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

	ProcessHardware
}
#endregion

#region finish script
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): Finishing up document"
#end of document processing

$AbstractTitle = "DHCP Inventory"
$SubjectTitle = "DHCP Inventory"
UpdateDocumentProperties $AbstractTitle $SubjectTitle

ProcessDocumentOutput

ProcessScriptEnd
#endregion