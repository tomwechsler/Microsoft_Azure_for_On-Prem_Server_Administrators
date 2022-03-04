
#Declaration of Variables
$objFormattedDate=get-date -f "dd-MM-yyyy HH:mm:ss"
$objTxtDate=get-date -f "ddMMyyyHHmmss"
$objHotfixDate = (get-date).AddDays(-7).ToString('dd/MM/yy HH:mm:ss tt')
$objHost=$env:computername
$objHTML=$null


#Start of HTML Document format
$objHTML=	"<html>"
$objHTML+=	"<head>"
$objHTML+=	"<Title><h1>" + $objHost + "</h1></Title>"
$objHTML+=	"<Style>"
$objHTML+=	" table{  border: 1px solid black;}`
			 td {  border-bottom: 1px solid #ddd;text-align: left;font-size:12}`
			 .label {  border-bottom: 1px solid #ddd;text-align: left;font-weight: bold;color:blue;font-size:18}`
			 th {  border-bottom: 1px solid #ddd;text-align: left;font-size:15}`
"
$objHTML+=	"</Style>"
$objHTML+=	"</head>"
$objHTML+=	"<body>"
$objHTML+=	"<h1><b>"+ $objHost+"_"+$objFormattedDate+"</h1></b><th>"






###################################################################################### SECTION::Operation System Information

#Get Wmi Information into Array
$objOsInfo = gwmi -class Win32_OperatingSystem

#Set the Table and first header
$objHTML+=	"<table width=100%>"
$objHTML+= "<tr> <br> </tr>" 
$objHTML+= "<tr>"
$objHTML+= "<th class=""label""> Computer Information  </th>"
$objHTML+= "</tr>"

#Begin the information dump
$objHTML+=	"<tr><td><b> Caption </b></td>"
$objHTML+=	"<td>" + $objOsInfo.Caption + "</td></tr>"
$objHTML+=	"<tr><td><b> Service Pack </b></td>"
$objHTML+=	"<td" + $objOsInfo.CSDVersion + "</td></tr>"
$objHTML+=	"<tr><td><b> Architecture </b></td>"
$objHTML+=	"<td>" + $objOsInfo.OSArchitecture + "</td></tr>"
$objHTML+=	"<tr><td ><b> WindowsDirectory </b></td>"
$objHTML+=	"<td>" + $objOsInfo.WindowsDirectory + "</td></tr>"
$objHTML+=	"<tr><td colspan=1><b> NumberOfProcesses </b></td>"
$objHTML+=	"<td colspan=1>" + $objOsInfo.NumberOfProcesses + "</td></tr>"
$objHTML+=	"<tr><td><b> TotalVisibleMemorySize </b></td>"
$objHTML+=	"<td>" + [math]::Round($objOsInfo.TotalVisibleMemorySize/1024/1024,2) + "GB</td></tr>"
$objHTML+=	"<tr><td><b> FreePhysicalMemory </b></td>"
$objHTML+=	"<td>" + [math]::Round($objOsInfo.FreePhysicalMemory/1024/1024,2) + "GB</td></tr>"
$objHTML+=	"<tr><td><b> TotalVirtualMemorySize </b></td>"
$objHTML+=	"<td>" + [math]::Round($objOsInfo.TotalVirtualMemorySize/1024/1024) + "GB</td></tr>"
$objHTML+=	"<tr><td><b> FreeVirtualMemory </b></td>"
$objHTML+=	"<td>" + [math]::Round($objOsInfo.FreeVirtualMemory/1024/1024,2) + "GB</td></tr>"
$objHTML+=	"<tr><td><b> InstallDate </b></td>"
$objHTML+=	"<td>" + $objOsInfo.InstallDate.substring(0,8) + "</td></tr>"
$objHTML+=	"<tr><td><b> LastBootUpTime </b></td>"
$objHTML+=	"<td>" + $objOsInfo.LastBootUpTime.substring(0,4) +"-"+ $objOsInfo.LastBootUpTime.substring(4,2) +"-"+ `
$objOsInfo.LastBootUpTime.substring(6,2) +" "+ $objOsInfo.LastBootUpTime.substring(8,2) +":"+ $objOsInfo.LastBootUpTime.substring(10,2) +":"+$objOsInfo.LastBootUpTime.substring(12,2) + "</td></tr>"
$objHTML+=	"</table>"
#End Of Operating System Information






###################################################################################### SECTION::Disks

#Get Wmi Information into Array
$objDiskInfo = gwmi -Query "Select * from Win32_LogicalDisk where DriveType = '3'" 

#Set the Table and first header
$objHTML+=	"<table width=100%>"
$objHTML+= "<tr> <br> </tr>" 
$objHTML+= "<tr>"
$objHTML+= "<th class=""label""> Disk Information  </th>"
$objHTML+= "</tr>"

#Set Headers
$objHTML+= "<tr>"
$objHTML+= "<th><b> Disk  </b></th>"
$objHTML+= "<th><b> Description  </b></th>"
$objHTML+= "<th><b> Size  </b></th>"
$objHTML+= "<th><b> FreeSpace  </b></th>"

#Loop for each item in Array
Foreach ( $objDisk in $objDiskInfo)
{
	#Check if Disk is low on space, if yes, mark font as red
	IF($objDisk.FreeSpace/1024/1024/1024 -le 10 )
	{
		$objHTML+=	"<tr style=""color:red"">"
	}
	Else 
	{
		$objHTML+=	"<tr>"
	}
	
	#Dump information
	$objHTML+=	"<td>" +	$objDisk.DeviceID				 	+ "</td>"
	$objHTML+=	"<td>" +	$objDisk.VolumeName					+ "</td>"
	$objHTML+=	"<td>" +	[math]::Round($objDisk.Size/1024/1024/1024,2)		+ "GB</td>"
	$objHTML+=	"<td>" +	[math]::Round($objDisk.FreeSpace/1024/1024/1024,2)	+ "GB</td>"
	$objHTML+=	"</tr>"
}
$objHTML+=	"</table>"
#END Logical Disk Information






###################################################################################### SECTION::Services

#Get Wmi Information into Array
$objServiceInfo = gwmi -Query "Select * from win32_service" 

#Set the Table and first header
$objHTML+=	"<table width=100%>"
$objHTML+= "<tr> <br> </tr>" 
$objHTML+= "<tr>"
$objHTML+= "<th class=""label""> Problematic Services  </th>"
$objHTML+= "</tr>"

#Set Headers
$objHTML+= "<tr>"
$objHTML+= "<th><b> Name  </b></th>"
$objHTML+= "<th><b> DisplayName  </b></th>"
$objHTML+= "<th><b> StartMode  </b></th>"
$objHTML+= "<th><b> State  </b></th>"
$objHTML+= "<th><b> ProcessId  </b></th>"
$objHTML+= "<th><b> Path  </b></th>"

#Loop for each item in Array
Foreach ( $objService in $objServiceInfo)
{
	#Check if Disk is low on space, if yes, mark font as red
	IF( ($objService.StartMode -eq "Auto" -or  $objService.StartMode -like "*Delayd*") -and $objService.State -ne "Running"  )
	{
		$objHTML+=	"<tr style=""color:red"">"
			
			#Move this content below the else statement to print all info
			$objHTML+=	"<td>" +	$objService.Name				 	+ "</td>"
			$objHTML+=	"<td>" +	$objService.DisplayName				+ "</td>"
			$objHTML+=	"<td>" +	$objService.StartMode				+ "</td>"
			$objHTML+=	"<td>" +	$objService.State					+ "</td>"
			$objHTML+=	"<td>" +	$objService.ProcessId				+ "</td>"
			$objHTML+=	"<td>" +	$objService.PathName				+ "</td>"
			$objHTML+=	"</tr>"
	}
	Else 
	{
		#Do Nothing.Remove below hash and follow above steps to print everything
		#$objHTML+=	"<tr>"
	}
	

}
$objHTML+=	"</table>"
#END Services Information






###################################################################################### SECTION::Processe=>Top CPU

#Get Wmi Information into Array
$objTopProcessCPU = Get-Process | Sort-Object  cpu -Descending | select -First 10

#Set the Table and first header
$objHTML+=	"<table width=100%>"
$objHTML+= "<tr> <br> </tr>" 
$objHTML+= "<tr>"
$objHTML+= "<th class=""label""> Top 10 Processes by CPU  </th>"
$objHTML+= "</tr>"

#Set Headers
$objHTML+= "<tr>"
$objHTML+= "<th><b> ProcessName  </b></th>"
$objHTML+= "<th><b> SessionId  </b></th>"
$objHTML+= "<th><b> CPU  </b></th>"
$objHTML+= "<th><b> PriorityClass  </b></th>"
$objHTML+= "<th><b> HandleCount  </b></th>"
$objHTML+= "<th><b> Path  </b></th>"

#Loop for each item in Array
Foreach ( $objProc in $objTopProcessCPU)
{
	$objHTML+=	"<tr>"
	$objHTML+=	"<td>" +	$objProc.ProcessName		+ "</td>"
	$objHTML+=	"<td>" +	$objProc.SessionId			+ "</td>"
	$objHTML+=	"<td>" +	$objProc.CPU				+ "</td>"
	$objHTML+=	"<td>" +	$objProc.PriorityClass		+ "</td>"
	$objHTML+=	"<td>" +	$objProc.HandleCount		+ "</td>"
	$objHTML+=	"<td>" +	$objProc.Path				+ "</td>"
	$objHTML+=	"</tr>"

}
$objHTML+=	"</table>"
#END SProcess Information -> CPU





###################################################################################### SECTION::Processe=>Top Memory

#Get Wmi Information into Array
$objTopProcessMem = Get-Process | Sort-Object  WorkingSet -Descending | select -First 10

#Set the Table and first header
$objHTML+=	"<table width=100%>"
$objHTML+= "<tr> <br> </tr>" 
$objHTML+= "<tr>"
$objHTML+= "<th class=""label""> Top 10 Processes by Memory  </th>"
$objHTML+= "</tr>"

#Set Headers
$objHTML+= "<tr>"
$objHTML+= "<th><b> ProcessName  </b></th>"
$objHTML+= "<th><b> SessionId  </b></th>"
$objHTML+= "<th><b> WorkingSet  </b></th>"
$objHTML+= "<th><b> VirtualMemorySize  </b></th>"
$objHTML+= "<th><b> PagedMemorySize  </b></th>"
$objHTML+= "<th><b> PrivateMemorySize  </b></th>"
$objHTML+= "<th><b> PagedSystemMemorySize  </b></th>"
$objHTML+= "<th><b> NonpagedSystemMemorySize  </b></th>"

#Loop for each item in Array
Foreach ( $objProc in $objTopProcessMem)
{
	$objHTML+=	"<tr>"
	$objHTML+=	"<td>" +	$objProc.ProcessName		 								+ "</td>"
	$objHTML+=	"<td>" +	$objProc.SessionId											+ "</td>"
	$objHTML+=	"<td>" +	[math]::Round($objProc.WorkingSet64/1024/1024,2)			+ "MB</td>"
	$objHTML+=	"<td>" +	[math]::Round($objProc.VirtualMemorySize64/1024/1024,2)		+ "MB</td>"
	$objHTML+=	"<td>" +	[math]::Round($objProc.PagedMemorySize/1024/1024,2)			+ "MB</td>"
	$objHTML+=	"<td>" +	[math]::Round($objProc.PrivateMemorySize/1024/1024,2)		+ "MB</td>"
	$objHTML+=	"<td>" +	[math]::Round($objProc.PagedSystemMemorySize/1024/1024,2)	+ "MB</td>"
	$objHTML+=	"<td>" +	[math]::Round($objProc.NonpagedSystemMemorySize/1024/1024,2)+ "MB</td>"
	$objHTML+=	"</tr>"

}
$objHTML+=	"</table>"
#END Process Information -> Memory





###################################################################################### SECTION::Patches

#Get Wmi Information into Array
$objHotfix = gwmi -query "Select * from win32_QuickfixEngineering where InstalledOn > '$objHotfixDate'"

#Set the Table and first header
$objHTML+=	"<table width=100%>"
$objHTML+= "<tr> <br> </tr>" 
$objHTML+= "<tr>"
$objHTML+= "<th class=""label""> Patches Installed in the Last 7 Days  </th>"
$objHTML+= "</tr>"

#Set Headers
$objHTML+= "<tr>"
$objHTML+= "<th><b> Name  </b></th>"
$objHTML+= "<th><b> HotFixID  </b></th>"
$objHTML+= "<th><b> Caption  </b></th>"
$objHTML+= "<th><b> InstallDate  </b></th>"
$objHTML+= "<th><b> InstalledBy  </b></th>"

#Loop for each item in Array
Foreach ( $objPatch in $objHotfix)
{
	$objHTML+=	"<tr>"
	$objHTML+=	"<td>" +	$objPatch.Name			+ "</td>"
	$objHTML+=	"<td>" +	$objPatch.HotFixID		+ "</td>"
	$objHTML+=	"<td>" +	$objPatch.Caption		+ "</td>"
	$objHTML+=	"<td>" +	$objPatch.InstallDate	+ "</td>"
	$objHTML+=	"<td>" +	$objPatch.InstalledBy	+ "</td>"
	$objHTML+=	"</tr>"

}
$objHTML+=	"</table>"
#END Last Hotfixes




###################################################################################### SECTION::Event Log System =>Error

#Get Wmi Information into Array
$objLogSysError = Get-Eventlog -LogName System -EntryType Error -Newest 10

#Set the Table and first header
$objHTML+=	"<table width=100%>"
$objHTML+= "<tr> <br> </tr>" 
$objHTML+= "<tr>"
$objHTML+= "<th class=""label""> Last 10 System Log Errors  </th>"
$objHTML+= "</tr>"

#Set Headers
$objHTML+= "<tr>"
$objHTML+= "<th><b> EventID  </b></th>"
$objHTML+= "<th><b> InstanceID  </b></th>"
$objHTML+= "<th><b> Time  </b></th>"
$objHTML+= "<th><b> Source  </b></th>"
$objHTML+= "<th><b> Message  </b></th>"

#Loop for each item in Array
Foreach ( $objEventError in $objLogSysError)
{
	$objHTML+=	"<tr>"
	$objHTML+=	"<td>" +	$objEventError.EventID		+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.InstanceID	+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Time			+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Source		+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Message		+ "</td>"
	$objHTML+=	"</tr>"

}
$objHTML+=	"</table>"
#END Event Log => System Error





###################################################################################### SECTION::Event Log System =>Warning

#Get Wmi Information into Array
$objLogSysError = Get-Eventlog -LogName System -EntryType Warning -Newest 10

#Set the Table and first header
$objHTML+=	"<table width=100%>"
$objHTML+= "<tr> <br> </tr>" 
$objHTML+= "<tr>"
$objHTML+= "<th class=""label""> Last 10 System Log Warnings  </th>"
$objHTML+= "</tr>"

#Set Headers
$objHTML+= "<tr>"
$objHTML+= "<th><b> EventID  </b></th>"
$objHTML+= "<th><b> InstanceID  </b></th>"
$objHTML+= "<th><b> Time  </b></th>"
$objHTML+= "<th><b> Source  </b></th>"
$objHTML+= "<th><b> Message  </b></th>"
$objHTML+= "</tr>"

#Loop for each item in Array
Foreach ( $objEventError in $objLogSysError)
{
	$objHTML+=	"<tr>"
	$objHTML+=	"<td>" +	$objEventError.EventID		+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.InstanceID	+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Time			+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Source		+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Message		+ "</td>"
	$objHTML+=	"</tr>"

}
$objHTML+=	"</table>"
#END Event Log => System Warning





###################################################################################### SECTION::Event Log Application =>Error

#Get Wmi Information into Array
$objLogSysError = Get-Eventlog -LogName Application -EntryType Error -Newest 10

#Set the Table and first header
$objHTML+=	"<table width=100%>"
$objHTML+= "<tr> <br> </tr>" 
$objHTML+= "<tr>"
$objHTML+= "<th class=""label""> Last 10 Application Log Error  </th>"
$objHTML+= "</tr>"

#Set Headers
$objHTML+= "<tr>"
$objHTML+= "<th><b> EventID  </b></th>"
$objHTML+= "<th><b> InstanceID  </b></th>"
$objHTML+= "<th><b> Time  </b></th>"
$objHTML+= "<th><b> Source  </b></th>"
$objHTML+= "<th><b> Message  </b></th>"
$objHTML+= "</tr>"

#Loop for each item in Array
Foreach ( $objEventError in $objLogSysError)
{
	$objHTML+=	"<tr>"
	$objHTML+=	"<td>" +	$objEventError.EventID		+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.InstanceID	+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Time			+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Source		+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Message		+ "</td>"
	$objHTML+=	"</tr>"

}
$objHTML+=	"</table>"
#END Event Log => Application Error





###################################################################################### SECTION::Event Log Application =>Warning

#Get Wmi Information into Array
$objLogSysError = Get-Eventlog -LogName Application -EntryType Warning -Newest 10

#Set the Table and first header
$objHTML+=	"<table width=100%>"
$objHTML+= "<tr> <br> </tr>" 
$objHTML+= "<tr>"
$objHTML+= "<th class=""label""> Last 10 Application Log Warning  </th>"
$objHTML+= "</tr>"

#Set Headers
$objHTML+= "<tr>"
$objHTML+= "<th><b> EventID  </b></th>"
$objHTML+= "<th><b> InstanceID  </b></th>"
$objHTML+= "<th><b> Time  </b></th>"
$objHTML+= "<th><b> Source  </b></th>"
$objHTML+= "<th><b> Message  </b></th>"
$objHTML+= "</tr>"

#Loop for each item in Array
Foreach ( $objEventError in $objLogSysError)
{
	$objHTML+=	"<tr>"
	$objHTML+=	"<td>" +	$objEventError.EventID		+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.InstanceID	+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Time			+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Source		+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Message		+ "</td>"
	$objHTML+=	"</tr>"

}
$objHTML+=	"</table>"
#END Event Log => Application Warning





###################################################################################### SECTION::Event Log Application =>Warning

#Get Wmi Information into Array
$objLogSysError = Get-Eventlog -LogName Security -EntryType FailureAudit -Newest 10

#Set the Table and first header
$objHTML+=	"<table width=100%>"
$objHTML+= "<tr> <br> </tr>" 
$objHTML+= "<tr>"
$objHTML+= "<th class=""label""> Last 10 Security Audit Failures  </th>"
$objHTML+= "</tr>"

#Set Headers
$objHTML+= "<tr>"
$objHTML+= "<th><b> EventID  </b></th>"
$objHTML+= "<th><b> InstanceID  </b></th>"
$objHTML+= "<th><b> Time  </b></th>"
$objHTML+= "<th><b> Source  </b></th>"
$objHTML+= "<th><b> Message  </b></th>"
$objHTML+= "</tr>"

#Loop for each item in Array
Foreach ( $objEventError in $objLogSysError)
{
	$objHTML+=	"<tr>"
	$objHTML+=	"<td>" +	$objEventError.EventID		+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.InstanceID	+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Time			+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Source		+ "</td>"
	$objHTML+=	"<td>" +	$objEventError.Message		+ "</td>"
	$objHTML+=	"</tr>"

}
$objHTML+=	"</table>"
#END Event Log => Application Warning


#End of HTML
$objHTML+=	"</body>"
$objHTML+=	"</html>"

#Write File
$objHTML | out-file $PSScriptRoot\$objHost_$objTxtDate".html"

#Set HTML Object to Null
$objHTML=$null
