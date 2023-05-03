# Straits SnipeIT Asset Reporting Script
# Version 1.0.3
# Changes - 27/8/2021
# Report Last Seen and Current Logged in User
# Changes for OS Information Column
# Correction to RAM information (showing installed ram size)
# - Parsad

function SnipeData()
{
	#need admin rights 
	if (-not(Get-ItemProperty -path "HKLM:\SOFTWARE\AssetInstallation" -Name Installed))
	{
		Write-Host "Installation"
		Set-ItemProperty -Path "HKLM:\SOFTWARE\AssetInstallation" -Name Installed -Value "1"
		Install-PackageProvider -Name "NuGet" -Force
		# One time only install: (requires an admin PowerShell window)
		Install-Module SnipeitPS -Force
		#Check for updates occasionally:
		Update-Module SnipeitPS
		Write-Host "Property Created and Module Installed"
		SnipeData
	}
	else
	{
		#Must import and set to every session.
		Import-Module SnipeitPS
		#set base url and the apikey for snipe-it module in powershell
		#API Key User autosync
		Set-SnipeitInfo -URL 'http://itasset.mydomain.local' -apiKey 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiYTg0YzUwNGI1ZDMyNTRjMWUxMjkwZTY4ZjRkZWZkYzBkZGQ2Y2ZmMWJkNWIwYWQ0Mjk2YWYwMzc5M2NiMTQzZDYyNWZjMTBmMjg5ZDBmZGQiLCJpYXQiOjE2Mjk2OTAyNzEsIm5iZiI6MTYyOTY5MDI3MSwiZXhwIjoyMjYwODQyMjcwLCJzdWIiOiIzMzkiLCJzY29wZXMiOltdfQ.Xl8mbIP3iY5JOInCS8rVvT9YdSvgzFtYP7dsrM9Yy7n0yuRNqW2B8ya_E0LxMnEbliPskzYptx1J_8w_jYlh6rwmwZsVB8I0maiyUvXIm4Y3LFJTF1hwoqJ_KiWgE3nWurO90WrZagsbeHYS2Hx6EQUuy2JJTzlM_mvvB_OACDoJ8YO9LCY7ix-GXNdiPimE4NFqgveS_IjIxwzQwNOx0Jcb1KqRw08jXqjhGcVxjytpYRBoEt0LjcpK-Mp5CMOwvFtErNNfkI09TxiE1IaQI2m8LI70WXNpwmBmbbTOyvEtr-kAA6ekQHPQQH6_afABWlTSPsfm1-8WJof3OyvnaVylJpxtUCCpQhzO42VxEA456il9hw1AjFaUkK-8plyq_QwWMAnTr16L5fghv4zV7qqFKaL4_osB84f7O298bbCmQnIO0e8qrkUNpItUXSonl4Pq-aOfFhrhtRqAerwvLbjTqH-q8tNCTrZcSakw16cDI-WN4MN1sP6PSo7T75_upRxAF4ljP9iNxhgRwXXw8vNbyf94IzX7SVqPKe1bfz8kjTLXALI2y1rMpry4ARoB24QkVFnI4Mp_T3cvVi-u4X6G0Trp9ZGJ5flddcdU0fJRJQVa0zxoRQzmij8kQH58mdaUfsO6lhWuTIIRkeaH30zxbmeTgwCk3aYJlYK1MkY'
		#Get PC Specs and Details
		$CPUInfo = (gwmi win32_ComputerSystem).name #Get CPU Information
		$OSInfo = Get-WmiObject Win32_OperatingSystem #Get OS Information
		$GetRam=Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | Foreach {"{0:N2}" -f ([math]::round(($_.Sum / 1GB),2))}
		$MacAddress=(Get-WmiObject Win32_NetworkAdapterConfiguration | where {$_.ipenabled -EQ $true}).Macaddress | select-object -first 1
		$CPU="CPU   :"+(Get-WMIObject win32_Processor).Name
		$RAM="RAM   :"+$GetRam+" GB."
		$IpAddress="IP Address   :"+(Test-Connection $CPUInfo -count 1).IPv4Address.IPAddressToString
		$DISKTOTAL = Get-CimInstance win32_logicaldisk | where caption -eq "C:" | foreach-object {write " $('{0:N2}' -f ($_.Size/1gb)) GB "}
		$DISKFREE = Get-CimInstance win32_logicaldisk | where caption -eq "C:" | foreach-object {write " $('{0:N2}' -f ($_.FreeSpace/1gb)) GB /"}
		$TotalDisk= "Available Disk Memory :"+$DISKFREE+$DISKTOTAL
		
		#get os version
		$get_version=(Get-WmiObject Win32_OperatingSystem).Version
		switch($get_version){
			'10.0.19043'{$get_version="21H1"}
			'10.0.19042'{$get_version="20H2"}
			'10.0.19041'{$get_version="2004"}
			'10.0.18363'{$get_version="1909"}
			'10.0.18362'{$get_version="1903"}
			'10.0.17763'{$get_version="1809"}
			'10.0.17134'{$get_version="1803"}
			'10.0.16299'{$get_version="1709"}
			'10.0.15063'{$get_version="1703"}
			'10.0.14393'{$get_version="1607"}
			'10.0.10586'{$get_version="1511"}
			}
			$Version     = $get_version
			$OSBuild     = "{0}.{1}" -f (Get-ItemPropertyValue "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name CurrentBuild), (Get-ItemPropertyValue "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name UBR)
			$Edition     = Get-ItemPropertyValue "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ProductName
			
			$CurrentUser=(Get-CimInstance -ClassName Win32_ComputerSystem).Username
			$CrtDate=(get-date).ToString('D')
			$CrtTime=(get-date).ToString('T')
			$LastSeen=$CrtDate+" | "+$CrtTime
			
			#Gathering Mutiple Data into Single
			$information= $CPU+"`n"+$RAM+"`n"+$IpAddress+"`n"+$TotalDisk+"`n"
			
			#Get Serial ID and name from the pc to get Asset ID from snipe-it 
			$ID=(Get-SnipeitAsset -serial (Get-WmiObject Win32_BIOS).SerialNumber).id
			$SerialNum=(Get-WmiObject Win32_BIOS).SerialNumber
			$Name=iex hostname
			$Model=(Get-ComputerInfo).CsModel
			$M_ID=(Get-SnipeitModel -search "$Model").id #Get model id from snipe-it to be assigned if does not exists.
			
			#check if asset serialnumber exists.
			if((Get-SnipeitAsset -serial (Get-WmiObject Win32_BIOS).SerialNumber).serial -eq (Get-WmiObject Win32_BIOS).SerialNumber )
			{
				#Update the asset
				Set-SnipeitAsset -id $ID -name " $Name" -customfields  @{"_snipeit_mac_address_1"="$MacAddress";"_snipeit_hardware_info_2"="$information";"_snipeit_window_version_3"="$Version";"_snipeit_os_build_5"="$OSBuild";"_snipeit_windows_edition_6"="$Edition";"_snipeit_current_signed_in_user_7"="$CurrentUser";"_snipeit_last_seen_8"="$LastSeen"}
				Write-Host "Updated."
				}
			else
			{
					#Create asset
					New-SnipeitAsset  -serial $SerialNum -asset_tag "$Name" -status_id 7 -model_id $M_ID -name "$Name" -customfields @{ "_snipeit_mac_address_1"="$MacAddress";"_snipeit_hardware_info_2"="$information";"_snipeit_window_version_3"="$Version";"_snipeit_os_build_5"="$OSBuild";"_snipeit_windows_edition_6"="$Edition";"_snipeit_current_signed_in_user_7"="$CurrentUser";"_snipeit_last_seen_8"="$LastSeen"}
					Write-Host "Created."
				}
	}
}



if(Test-Path HKLM:\SOFTWARE\AssetInstallation)
{
	SnipeData
}
else
{
	#Creates a new folder at HKLM:\SOFTWARE 
	New-Item -Path "HKLM:\SOFTWARE" -Name AssetInstallation
	SnipeData
}