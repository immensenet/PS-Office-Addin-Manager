function Get-UserFromSID
{
	[CmdletBinding()]
	param
	(
		[string]$SID
	)
	
	$UPN = Get-ItemProperty -ErrorAction SilentlyContinue -Path (Join-Path "HKLM:\SOFTWARE\Microsoft\IdentityStore\LogonCache\B16898C6-A148-4967-9171-64D755DA8520\Sid2Name" $SID) -Name IdentityName | Select-Object -ExpandProperty IdentityName
	if ($UPN)
	{
		return $UPN
	}
	else
	{
		try
		{
			return (New-Object System.Security.Principal.SecurityIdentifier $SID).Translate([System.Security.Principal.NTAccount]).Value
		}
		catch
		{
			Write-Error "Unable to convert SID: $($MostRecentLogonEvent.UserSid) to Username"
		}
	}
}
function Get-OfficeRegistryPath
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$UserRegistryMountPoint,
		[Parameter(Mandatory = $true)]
		[ValidateSet('Outlook', 'Excel', 'Word', 'PowerPoint', 'OneNote', 'Project', 'Publisher')]
		[string]$OfficeProduct,
		[string]$ProgID = ""
	)
	
	$OfficeInfo = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
	$OfficeVersion = [System.Version]($OfficeInfo | Get-ItemPropertyValue -Name "VersionToReport")
	$ShortVersion = ($OfficeVersion.Major.ToString() + "." + $OfficeVersion.Minor.ToString())
	$WOW6432Shim = if ($OfficeInfo.Platform -like "x86" -and $env:PROCESSOR_ARCHITECTURE -notlike "x86") { "\WOW6432Node" }
	else { "" }
	$MachineAddinClassPath = "HKLM:\SOFTWARE\Classes\$WOW6432Shim\CLSID"
	$UserClassPath = "$UserRegistryMountPoint\Software\Classes\$WOW6432Shim\CLSID"
	
	$RetVal = [pscustomobject]@{
		MachineClassPath			   = $MachineAddinClassPath
		UserClassPath				   = $UserClassPath
		UserOfficeCompatAddinCleanLoad = "$UserRegistryMountPoint\Software\Microsoft\OfficeCompat\$OfficeProduct\AddinCleanLoad"
		UserOfficeCompatAddinUsage	   = "$UserRegistryMountPoint\Software\Microsoft\OfficeCompat\$OfficeProduct\AddinUsage"
		UserAddinResiliencyDisabledItemsPath = "$UserRegistryMountPoint\Software\Microsoft\Office\$($ShortVersion)\$OfficeProduct\Resiliency\DisabledItems"
		UserAddinResiliencyCrashingAddinListPath = "$UserRegistryMountPoint\Software\Microsoft\Office\$($ShortVersion)\$OfficeProduct\Resiliency\CrashingAddinList"
		UserAddinResiliencyDoNotDisableList = "$UserRegistryMountPoint\Software\Microsoft\Office\$($ShortVersion)\$OfficeProduct\Resiliency\DoNotDisableAddinList"
		UserAddinResiliencyNotificationReminderAddinData = "$UserRegistryMountPoint\Software\Microsoft\Office\$($ShortVersion)\$OfficeProduct\Resiliency\NotificationReminderAddinData"
		UserAddinLoadTimes			   = "$UserRegistryMountPoint\Software\Microsoft\Office\$($ShortVersion)\$OfficeProduct\AddInLoadTimes"
		UserAddinPerformanceLog	       = "$UserRegistryMountPoint\Software\Microsoft\Office\$($ShortVersion)\$OfficeProduct\Addins"
		ManagedAddins				   = "$UserRegistryMountPoint\Software\Policies\Microsoft\Office\$($ShortVersion)\$OfficeProduct\Resiliency\AddinList"
	}
	$AddinPaths = @("HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\$WOW6432Shim\Microsoft\Office\$OfficeProduct\Addins", "HKLM:\SOFTWARE\$WOW6432Shim\Microsoft\Office\$OfficeProduct\Addins", "$UserRegistryMountPoint\Software\Microsoft\Office\$OfficeProduct\Addins")
	if (![string]::IsNullOrWhiteSpace($ProgID))
	{
		$AddinPaths = $AddinPaths | ForEach-Object{ Join-Path $_ $ProgID } | Where-Object{ Test-Path $_ }
		$RetVal.UserAddinLoadTimes = Join-Path $RetVal.UserAddinLoadTimes $ProgID
		$RetVal.UserAddinPerformanceLog = Join-Path $RetVal.UserAddinPerformanceLog $ProgID
	}
	$RetVal | Add-Member -NotePropertyName AddinPaths -NotePropertyValue $AddinPaths
	$RetVal
}
function Get-ClassAssemblyPath
{
	[CmdletBinding()]
	param
	(
		[string]$ProgID,
		[string]$SID
	)
	
	$ProgIDSearchPaths = @()
	$ProgIDSearchPaths += "registry::HKEY_USERS\$($SID)\Software\Classes"
	$ProgIDSearchPaths += "HKLM:\Software\Classes"
	$ProgIDSearchPaths += "HKLM:\Software\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes"
	$Platform = Get-Item -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" | ForEach-Object{ $_.GetValue("Platform") }
	$WOW6432Node = if ($Platform -like "x86" -and $env:PROCESSOR_ARCHITECTURE -notlike "x86") { "\WOW6432Node" }
	else { "" }
	
	foreach ($SearchPath in $ProgIDSearchPaths)
	{
		$CurrentVersion = Get-Item (Join-Path $SearchPath "$ProgID\CurVer") -ErrorAction SilentlyContinue | ForEach-Object{ $_.GetValue("") }
		if ($CurrentVersion)
		{
			$ProgID = $CurrentVersion
		}
		$CLSID = Get-Item (Join-Path $SearchPath "$ProgID\CLSID") -ErrorAction SilentlyContinue | ForEach-Object{ $_.GetValue("") }
		if ($CLSID)
		{
			$RetVal = Get-Item (Join-Path $SearchPath "$WOW6432Node\CLSID\$CLSID\InProcServer32") -ErrorAction SilentlyContinue | ForEach-Object{ $_.GetValue("") }
			if ($RetVal)
			{
				return $RetVal;
			}
		}
	}
}

function Get-VSTOAssemblyPath
{
	[CmdletBinding()]
	param
	(
		[string]$VSTOManifest
	)
	
	[uri]$ManifestUri = $VSTOManifest.Split("|") | Select-Object -First 1
	$ManifestPath = $ManifestUri.LocalPath
	[xml]$ManifestXML = Get-Content $ManifestPath
	$AssemblyIdentity = $ManifestXML.assembly.dependency.dependentAssembly.assemblyIdentity
	$AddinFolderPath = Split-Path $ManifestPath -Parent
	$AddinDLLPath = Join-Path $AddinFolderPath $AssemblyIdentity.name
	if ((Test-Path $AddinDLLPath))
	{
		$ManifestPath
	}
}

function Get-UserProfile
{
	[CmdletBinding()]
	param
	(
		[string]$UserName,
		[string]$UserSID,
		[switch]$MostRecentlyLoggedOn
	)
	
	$LoggedOnUser = Get-WmiObject -ClassName Win32_Process -Filter "Name = 'Explorer.exe'" | ForEach-Object {
		$Owner = $_.GetOwner()
		[pscustomobject]@{
			Domain = $Owner.Domain
			Username = $Owner.User
			SID	     = $_.GetOwnerSid().sid
		} 
	} | Select-Object -Unique -First 1
	if ($LoggedOnUser)
	{
		$UserSID = $LoggedOnUser.SID
	}
	
		# Regex pattern for SIDs
		$PatternSID = 'S-1-5-21-\d+-\d+\-\d+\-\d+$'
		$RegProfileList = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*' | Where-Object { $_.PSChildName -match $PatternSID }
		$ProfileList = $RegProfileList | Where-Object {
			[string]::IsNullOrWhiteSpace($UserSID) -or $_.PSChildName -like $UserSID
		} | ForEach-Object {
		$UserProfile = $_
		$SID = $_.PSChildName
		$UserDir = Split-Path -Path $UserProfile.ProfileImagePath -Leaf
		$UserHivePath = Join-Path $UserProfile.ProfileImagePath "ntuser.dat"
		$ProfileObj = [PSCustomObject]@{ }
		$ProfileObj | Add-Member -NotePropertyName SID -NotePropertyValue $SID
		$ProfileObj | Add-Member -NotePropertyName UserDir -NotePropertyValue $UserDir
		$ProfileObj | Add-Member -NotePropertyName LastUse -NotePropertyValue (Get-Item -Path $UserProfile.ProfileImagePath | Select-Object -ExpandProperty LastWriteTime -First 1)
		$ProfileObj | Add-Member -NotePropertyName UserHivePath -NotePropertyValue $UserHivePath
		$ProfileObj | Add-Member -NotePropertyName UserName -NotePropertyValue (Get-UserFromSID $SID)
		$ProfileObj | Add-Member -NotePropertyName PSComputerName -NotePropertyValue $env:COMPUTERNAME
		$ProfileObj
	}
	if ($MostRecentlyLoggedOn)
	{
		$ProfileList | Sort-Object LastUse -Descending | Select-Object -First 1
	}
	else
	{
		$ProfileList | Sort-Object LastUse -Descending
	}
}

function Get-ManagedAddin
{
	[CmdletBinding()]
	param
	(
		$UserRegistryMountPoint,
		[Parameter(Mandatory = $true)]
		[ValidateSet('Outlook', 'Excel', 'Word', 'PowerPoint', 'OneNote', 'Project', 'Publisher')]
		[string]$OfficeProduct
	)
	$OfficeRegistryPaths = Get-OfficeRegistryPath -UserRegistryMountPoint $UserRegistryMountPoint -OfficeProduct $OfficeProduct
	Get-Item $OfficeRegistryPaths.ManagedAddins -ErrorAction SilentlyContinue | ForEach-Object {
		$ManagedAddinKey = $_
		$ManagedAddins = $_.GetValueNames()
		$ManagedAddins | ForEach-Object{
			$ProgID = $_
			$ManagedBehaviorID = $ManagedAddinKey.GetValue($_)
			$RetObj = [pscustomobject]@{ }
			$RetObj | Add-Member -NotePropertyName "ProgID" -NotePropertyValue $ProgID
			$ManagedBehavior = switch ($ManagedBehaviorID.ToString())
			{
				"0" { "AlwaysDisabled" }
				"1" { "AlwaysEnabled" }
				"2" { "UserDefined" }
			}
			$RetObj | Add-Member -NotePropertyName "ManagedBehavior" -NotePropertyValue $ManagedBehavior
			$RetObj
		}
	}
}
function Get-ResiliencyCrashingItem
{
	[CmdletBinding()]
	param
	(
		$UserRegistryMountPoint,
		[Parameter(Mandatory = $true)]
		[ValidateSet('Outlook', 'Excel', 'Word', 'PowerPoint', 'OneNote', 'Project', 'Publisher')]
		[string]$OfficeProduct,
		[string]$DLLPath
	)
	$OfficeRegistryPaths = Get-OfficeRegistryPath -UserRegistryMountPoint $UserRegistryMountPoint -OfficeProduct $OfficeProduct
	$Item = Get-Item $OfficeRegistryPaths.UserAddinResiliencyCrashingAddinListPath -ErrorAction SilentlyContinue
	if ($Item)
	{
		$Item | Select-Object -ExpandProperty property | ForEach-Object {
			$PropertyName = $_
			$Bytes = $Item.GetValue($PropertyName)
			if ($Bytes)
			{
				Get-ResiliencyCrashingDataFromByteArray -Bytes $Bytes
			}
		}
	}
}

function Get-ResiliencyCrashingDataFromByteArray
{
	param ([byte[]]$Bytes)
	$RetVal = New-Object psobject
	try
	{
		$stream = New-Object System.IO.MemoryStream -ArgumentList @( , $Bytes)
		$reader = New-Object System.IO.BinaryReader -ArgumentList @($stream, [System.Text.Encoding]::Unicode)
		$pathLength = $reader.ReadInt32();
		$pathBytes = $reader.ReadBytes($pathLength);
		$path = [System.Text.Encoding]::Unicode.GetString($pathBytes);
		$RetVal | Add-Member -NotePropertyName DLLPath -NotePropertyValue $path.Replace("`0", "")
		$RetVal | Add-Member -NotePropertyName RegistryProperty -NotePropertyValue $ItemProperty
		
		if (!$DLLpath -or (($DLLPath -and $RetVal.DllPath -like $DLLPath)))
		{
			Write-Output $RetVal
		}
	}
	catch
	{
		Write-Error $_
	}
	finally
	{
		$reader.Close()
		$stream.Close()
	}
}

function Get-ResiliencyDisabledDataFromByteArray
{
	param ([byte[]]$Bytes)
	$RetVal = New-Object psobject
	try
	{
		$stream = New-Object System.IO.MemoryStream -ArgumentList @( , $Bytes)
		$reader = New-Object System.IO.BinaryReader -ArgumentList @($stream, [System.Text.Encoding]::Unicode)
		$type = $reader.ReadInt32();
		$pathLength = $reader.ReadInt32();
		$nameLength = $reader.ReadInt32();
		$pathBytes = $reader.ReadBytes($pathLength);
		$RetVal | Add-Member -NotePropertyName Type -NotePropertyValue $type
		$path = $null
		if ($type -band 0x40000000)
		{
			for ($i = 0; $i -lt 5; $i++) { $reader.ReadInt32() | Out-Null }
		}
		$path = [System.Text.Encoding]::Unicode.GetString($pathBytes);
		$RetVal | Add-Member -NotePropertyName DLLPath -NotePropertyValue $path.Replace("`0", "")
		$nameBytes = $reader.ReadBytes($nameLength);
		$name = $null
		$name = [System.Text.Encoding]::Unicode.GetString($nameBytes);
		$RetVal | Add-Member -NotePropertyName ProgID -NotePropertyValue $name.Replace("`0", "")		
		$RetVal
	}
	catch
	{
		Write-Error $_
	}
	finally
	{
		$reader.Close()
		$stream.Close()
	}
}

function Get-ResiliencyDisabledItem
{
	[CmdletBinding()]
	param
	(
		$UserRegistryMountPoint,
		$DLLPath,
		$ProgID,
		[Parameter(Mandatory = $true)]
		[ValidateSet('Outlook', 'Excel', 'Word', 'PowerPoint', 'OneNote', 'Project', 'Publisher')]
		[string]$OfficeProduct
	)
	
	$OfficeRegistryPaths = Get-OfficeRegistryPath -UserRegistryMountPoint $UserRegistryMountPoint -OfficeProduct $OfficeProduct
	$Item = Get-Item $OfficeRegistryPaths.UserAddinResiliencyDisabledItemsPath -ErrorAction SilentlyContinue
	$Item | Select-Object -ExpandProperty property | ForEach-Object {
		$PropertyName = $_
		$Bytes = $Item.GetValue($PropertyName)
		if ($Bytes)
		{
			$Retval = Get-ResiliencyDisabledDataFromByteArray -Bytes $Bytes
			$ItemProperty = $Item | Get-ItemProperty -Name $PropertyName
            $ItemPropertyPipelineable = [pscustomobject]@{"Name"=$PropertyName;"Path"=$ItemProperty.PSPath}            
            $RetVal | Add-Member -NotePropertyName RegistryProperty -NotePropertyValue $ItemPropertyPipelineable

			if (!$DLLpath -and !$ProgID)
			{
				Write-Output $RetVal
			}
            elseif($RetVal.ProgID -and $ProgID -and ($ProgID -like $RetVal.ProgID))
            {
				Write-Output $RetVal
            }
            elseif(!$ProgID -and $DLLPath -and ($RetVal.DllPath -like $DLLPath))
            {
				Write-Output $RetVal
            }
		}
	}
}

function Get-AddinLoadTime
{
	[CmdletBinding()]
	param
	(
		$UserRegistryMountPoint,
		$ProgID,
		[Parameter(Mandatory = $true)]
		[ValidateSet('Outlook', 'Excel', 'Word', 'PowerPoint', 'OneNote', 'Project', 'Publisher')]
		[string]$OfficeProduct
	)
	
	$OfficeRegistryPaths = Get-OfficeRegistryPath -UserRegistryMountPoint $UserRegistryMountPoint -OfficeProduct $OfficeProduct
	Get-Item $OfficeRegistryPaths.UserAddinLoadTimes -ErrorAction SilentlyContinue | ForEach-Object {
		$LoadTimeRegValue = $_.GetValue($ProgID)
		if ($LoadtimeRegvalue)
		{
			try
			{
				[byte[]]$Bytes = $LoadTimeRegValue
				$stream = New-Object System.IO.MemoryStream -ArgumentList @( , $Bytes)
				$reader = New-Object System.IO.BinaryReader -ArgumentList @($stream, [System.Text.Encoding]::Unicode)
				$null = $reader.ReadInt32()
				$reader.ReadInt32() | Out-Null
				$LoadTime = $reader.ReadInt32()
				$LoadTime
			}
			catch
			{
				Write-Error $_
			}
			finally
			{
				$reader.Close()
				$stream.Close()
			}
		}
	}
}

function Get-MergedRegistryValue
{
	[CmdletBinding()]
	param
	(
		[string[]]$Paths
	)
	
	$HashTable = @{ }
	$Paths | ForEach-Object{
		$Path = $_
		
		$Item = Get-Item $Path
		foreach ($Property in $Item.Property)
		{
			$HashTable."$Property" = $Item.GetValue($Property)
		}
	}
	[pscustomobject]$HashTable
}

function Open-RegistryHive
{
	[CmdletBinding()]
	param
	(
		[string]$HivePath,
		[string]$PSMountPoint
	)
	$CmdMountPoint = ([uri]$PSMountPoint).LocalPath.Substring(1)
	reg load $CmdMountPoint $HivePath | Out-Null
}

function Close-RegistryHive
{
	[CmdletBinding()]
	param
	(
		[string]$PSMountPoint
	)
	$CmdMountPoint = ([uri]$PSMountPoint).LocalPath.Substring(1)
	reg unload $CmdMountPoint | Out-Null
}

Function Get-FriendlyDisabledReason
{
	param ([int]$ReasonCode)
	switch ($ReasonCode)
	{
		0x00000001 { "Boot load (LoadBehavior = 3)" }
		0x00000002 { "Demand load (LoadBehavior = 9)" }
		0x00000003 { "Crash" }
		0x00000004 { "Handling FolderSwitch event" }
		0x00000005 { "Handling BeforeFolderSwitch event" }
		0x00000006 { "Item Open" }
		0x00000007 { "Iteration Count" }
		0x00000008 { "Shutdown" }
		0x00000009 { "Crash, but not disabled because add-in is in the allow list" }
		0x0000000A { "Crash, but not disabled because user selected no in disable dialog" }
	}
}

function Get-OfficeAddin
{
	[CmdletBinding(DefaultParameterSetName = 'AutoDetectUser')]
	[OutputType([array])]
	param
	(
		[Parameter(ParameterSetName = 'SpecifiedUser')]
		[ValidatePattern('S-1-\d+\-\d+\-\d+\-\d+\-\d+\-\d+$')]
		[string]$UserSID = $null,
		[Parameter(ParameterSetName = 'AutoDetectUser')]
		[switch]$AutoDetectUser,
		[Parameter(Mandatory = $true)]
		[ValidateSet('Outlook', 'Excel', 'Word', 'PowerPoint', 'OneNote', 'Project', 'Publisher')]
		[string]$OfficeProduct,
		[string]$ProgID = $null
	)
	
	$UserProfile = Get-UserProfile -UserSID $UserSID -MostRecentlyLoggedOn:$AutoDetectUser
	
	$PSUserRegistryMountPoint = "Registry::HKEY_USERS\$($UserProfile.SID)"
	$OfficeRegistryPaths = Get-OfficeRegistryPath -UserRegistryMountPoint $PSUserRegistryMountPoint -OfficeProduct $OfficeProduct
	if (!(Test-Path $OfficeRegistryPaths.UserClassPath))
	{
		$UnloadHive = $true
		Open-RegistryHive -HivePath $UserProfile.UserHivePath -MountPoint $PSUserRegistryMountPoint
	}
	
	
	Write-Verbose "Getting $OfficeProduct Addins for $($UserProfile.Username)..."
	$ManagedAddins = Get-ManagedAddin -UserRegistryMountPoint $PSUserRegistryMountPoint -OfficeProduct $OfficeProduct
	$DoNotDisableList = Get-Item $OfficeRegistryPaths.UserAddinResiliencyDoNotDisableList -ErrorAction SilentlyContinue
	$AddinRegistryList = Get-ChildItem -Path $OfficeRegistryPaths.AddinPaths
	if (![string]::IsNullOrWhiteSpace($ProgID))
	{
		$AddinRegistryList = $AddinRegistryList | Where-Object{ $_.PsChildName -like $ProgID }
	}
	$AddinList = $AddinRegistryList | Group-Object PsChildName | ForEach-Object {
		$Addin = Get-MergedRegistryValue $_.Group.PsPath
		$RetObj = [pscustomobject]@{ }
		$RetObj | Add-Member -NotePropertyName Computer -NotePropertyValue $env:COMPUTERNAME
		$RetObj | Add-Member -NotePropertyName UserName -NotePropertyValue $UserProfile.UserName
		$RetObj | Add-Member -NotePropertyName OfficeProduct -NotePropertyValue $OfficeProduct
		$RetObj | Add-Member -NotePropertyName UserSID -NotePropertyValue $UserProfile.SID
		$RetObj | Add-Member -NotePropertyName ProgID -NotePropertyValue $_.Name
		$FriendlyName = $null
		if (!$Addin.FriendlyName)
		{
			Write-Warning "Unable to find FriendlyName for $($_.Name)"
			$FriendlyName = $_.Name
		}
		else
		{
			$FriendlyName = $Addin.FriendlyName
		}
		$RetObj | Add-Member -NotePropertyName FriendlyName -NotePropertyValue $FriendlyName
		Write-Verbose ($RetObj.ProgID + " - " + $FriendlyName)
		$Location = if ($Addin.Manifest -and $Addin.Manifest.EndsWith("|vstolocal"))
		{
			# VSTO
			$Addin.Manifest
		}
		elseif ($Addin.FileName)
		{
			# Native
			$Addin.FileName
		}
		else
		{
			# Managed
			Get-ClassAssemblyPath -ProgID $RetObj.ProgID -SID $UserProfile.SID
		}
		if (!$Location)
		{
			Write-Warning "Unable to find DLL for $($FriendlyName)"
			return
		}
		$ManagedEntry = $ManagedAddins | Where-Object{ $_.ProgID -eq $RetObj.ProgID }
		$ManagedState = if ($MangedEntry)
		{
			$ManagedEntry.ManagedBehavior
		}
		else
		{
			"Unmanaged"
		}
		$RetObj | Add-Member -NotePropertyName Location -NotePropertyValue $Location
		$RetObj | Add-Member -NotePropertyName Managed -NotePropertyValue $ManagedState
		$DisabledItemUserRegData = Get-ResiliencyDisabledItem -UserRegistryMountPoint $PSUserRegistryMountPoint -DLLPath $Location -OfficeProduct $OfficeProduct -ProgID $RetObj.ProgID
		$RetObj | Add-Member -NotePropertyName InDisabledList -NotePropertyValue ($null -ne $DisabledItemUserRegData)
		$CrashingItemUserRegData = Get-ResiliencyCrashingItem -UserRegistryMountPoint $PSUserRegistryMountPoint -OfficeProduct $OfficeProduct -DLLPath $Location
		$RetObj | Add-Member -NotePropertyName InCrashingList -NotePropertyValue ($null -ne $CrashingItemUserRegData)
		$RetObj | Add-Member -NotePropertyName LoadTime -NotePropertyValue (Get-AddinLoadTime -UserRegistryMountPoint $PSUserRegistryMountPoint -ProgID $RetObj.ProgID -OfficeProduct $OfficeProduct)
		$RetObj | Add-Member -NotePropertyName InDoNotDisableList -NotePropertyValue ($DoNotDisableList.Property -contains $RetObj.ProgID)
		$RetObj | Add-Member -NotePropertyName Loaded -NotePropertyValue (($Addin.Loadbehavior -band 1) -ne 0)
		$RetObj | Add-Member -NotePropertyName LoadAtStartup -NotePropertyValue (($Addin.Loadbehavior -band 2) -ne 0)
		$RetObj | Add-Member -NotePropertyName LoadOnDemand -NotePropertyValue (($Addin.Loadbehavior -band 8) -ne 0)
		$RetObj | Add-Member -NotePropertyName LoadOnce -NotePropertyValue (($Addin.Loadbehavior -band 16) -ne 0)
		if ($RetObj.InDisabledList)
		{
			
		}
		
		$RetObj
	}
	Write-Output $AddinList | Where-Object{ [string]::IsNullOrEmpty($_.FriendlyName) -eq $false }
	[gc]::Collect()
	IF ($UnloadHive)
	{
		Close-RegistryHive -MountPoint $PSUserRegistryMountPoint
	}
}

function Set-OfficeAddin
{
	[CmdletBinding(SupportsShouldProcess)]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[ValidateSet('Outlook', 'Excel', 'Word', 'PowerPoint', 'OneNote', 'Project', 'Publisher')]
		[string]$OfficeProduct,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$ProgID,
		[Parameter(Mandatory = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[string]$UserSID,
		[boolean]$DoNotDisableList = $null,
		[switch]$ClearResiliencyData,
		[ValidateSet('Unmanaged', 'AlwaysEnabled', 'AlwaysDisabled', 'UserDefined')]
		[string]$ManagedBehavior = $null,
		[boolean]$LoadAtStartup,
		[switch]$Force
	)
	Begin
	{
		if (-not $PSBoundParameters.ContainsKey('Verbose'))
		{
			$VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference')
		}
		if (-not $PSBoundParameters.ContainsKey('Confirm'))
		{
			$ConfirmPreference = $PSCmdlet.SessionState.PSVariable.GetValue('ConfirmPreference')
		}
		if (-not $PSBoundParameters.ContainsKey('WhatIf'))
		{
			$WhatIfPreference = $PSCmdlet.SessionState.PSVariable.GetValue('WhatIfPreference')
		}
		Write-Verbose ('[{0}] Confirm={1} ConfirmPreference={2} WhatIf={3} WhatIfPreference={4}' -f $MyInvocation.MyCommand, $Confirm, $ConfirmPreference, $WhatIf, $WhatIfPreference)
	}
	PROCESS
	{
		$UserProfile = Get-UserProfile -UserSID $UserSID -MostRecentlyLoggedOn:$AutoDetectUser
		$PSUserRegistryMountPoint = "Registry::HKEY_USERS\$($UserProfile.SID)"
		$OfficeRegistryPaths = Get-OfficeRegistryPath -UserRegistryMountPoint $PSUserRegistryMountPoint -OfficeProduct $OfficeProduct -ProgID $ProgID
		if ($null -ne $PSItem)
		{
			$AddinCustomObj = $PSItem
		}
		else
		{
			$AddinCustomObj = Get-OfficeAddin -UserSID $UserSID -OfficeProduct $OfficeProduct -ProgID $ProgID
		}
		if (!(Test-Path $OfficeRegistryPaths.UserClassPath))
		{
			$UnloadHive = $true
			Open-RegistryHive -HivePath $UserProfile.UserHivePath -MountPoint $PSUserRegistryMountPoint
		}
		if ($null -ne $DoNotDisableList)
		{
			Write-Verbose "Adding $ProgID to UserAddinResiliencyDoNotDisableList"
			if ($DoNotDisableList -and ($PSCmdlet.ShouldProcess("ShouldProcess?") -or $Force))
			{
				New-Item -Path $OfficeRegistryPaths.UserAddinResiliencyDoNotDisableList -Force | Out-Null
				New-ItemProperty -Path $OfficeRegistryPaths.UserAddinResiliencyDoNotDisableList -Name $ProgID -Value 1 -Force | Out-Null
			}
			else
			{
				Remove-ItemProperty -Path $OfficeRegistryPaths.UserAddinResiliencyDoNotDisableList -Name $ProgID -ErrorAction SilentlyContinue
			}
		}
		
		if ($ClearResiliencyData)
		{
			Write-Verbose "Clearing resiliency data"
			$DisabledItem = Get-ResiliencyDisabledItem -DLLPath $AddinCustomObj.Location -ProgID $ProgID -UserRegistryMountPoint $PSUserRegistryMountPoint -OfficeProduct $OfficeProduct
			if ($DisabledItem -and ($PSCmdlet.ShouldProcess("ShouldProcess?") -or $Force))
			{
				$DisabledItem.RegistryProperty | Remove-ItemProperty
			}
			Write-Verbose "Clearing UserAddinCrashLogs"
			$CrashingItem = Get-ResiliencyCrashingItem -DLLPath $AddinCustomObj.Location -UserRegistryMountPoint $PSUserRegistryMountPoint -OfficeProduct $OfficeProduct
			if ($CrashingItem -and ($PSCmdlet.ShouldProcess("ShouldProcess?") -or $Force))
			{
				Remove-ItemProperty -Path $CrashingItem.RegistryProperty
			}
			Write-Verbose "Clearing UserAddinPerformanceLog"
			if ($PSCmdlet.ShouldProcess("ShouldProcess?") -or $Force)
			{
				New-Item -Path $OfficeRegistryPaths.UserAddinPerformanceLog -Force | Out-Null
			}
			$AddinCleanLoadItems = Get-Item $OfficeRegistryPaths.UserOfficeCompatAddinCleanLoad
			Write-Verbose "Clearing CleanLoadItems"
			if ($PSCmdlet.ShouldProcess("ShouldProcess?") -or $Force)
			{
				$AddinCleanLoadItems.GetValueNames() | Where-Object{ $_ -like $AddinCustomObj.Location } | ForEach-Object{
					Remove-ItemProperty -Path $AddinCleanLoadItems -Name $_
				}
			}
			Write-Verbose "Clearing UserAddinUsageLogs"
			if ($PSCmdlet.ShouldProcess("ShouldProcess?") -or $Force)
			{
				$AddinUsageItems = Get-Item $OfficeRegistryPaths.UserOfficeCompatAddinUsage
				$AddinUsageItems.GetValueNames() | Where-Object{ $_ -like $AddinCustomObj.Location } | ForEach-Object{
					Remove-ItemProperty -Path $AddinUsageItems -Name $_
				}
			}
			Write-Verbose "Clearing UserAddinResiliencyNotificationReminderAddinData"
			if ($PSCmdlet.ShouldProcess("ShouldProcess?") -or $Force)
			{
				Remove-ItemProperty -Path $OfficeRegistryPaths.UserAddinResiliencyNotificationReminderAddinData -Name "$ProgID*" -Force -ErrorAction SilentlyContinue
			}
			
			if ($null -ne $LoadAtStartup)
			{
				foreach ($AddinPath in $OfficeRegistryPaths.AddinPaths)
				{
					
					$LoadBehavior = Get-Item -Path $AddinPath | ForEach-Object{ $_.GetValue("LoadBehavior") }
					if ($null -ne $LoadBehavior)
					{
						
						$LoadBehavior = [int]$LoadBehavior
						if ($LoadAtStartup)
						{
							$LoadBehavior = $LoadBehavior -bor 0x03
						}
						else
						{
							$LoadBehavior = $LoadBehavior -band (-bnot 0x00000003)
						}
						Write-Verbose "Setting Loadbehavior to $LoadBehavior for $AddinPath"
						if ($PSCmdlet.ShouldProcess("ShouldProcess?") -or $Force)
						{
							Set-ItemProperty -Path $AddinPath -Name LoadBehavior -Value $LoadBehavior -Force
						}
					}
				}
			}
			if ($ManagedBehavior)
			{
				$ProductReleaseIds = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name ProductReleaseIds
				if ($ManagedBehavior -ne "Unmanaged")
				{
					if ($ProductReleaseIds -notlike "*O365ProPlusRetail*")
					{
						Write-Warning "Office ProPlus not detected. Managed behavior will likely not function as only ProPlus processes Group Policies"
					}
					$DesiredValue = switch ($ManagedBehavior)
					{
						'AlwaysDisabled' { 0 }
						'AlwaysEnabled' { 1 }
						'UserDefined' { 2 }
					}
					Write-Verbose "Setting ManagedState of $ProgID to $($DesiredValue.ToString())"
					if ($PSCmdlet.ShouldProcess("ShouldProcess?") -or $Force)
					{
						if (!(Test-Path $OfficeRegistryPaths.ManagedAddins))
						{
							New-Item $OfficeRegistryPaths.ManagedAddins -Force | Out-Null
						}
						New-ItemProperty -Path $OfficeRegistryPaths.ManagedAddins -Name $ProgID -Value $DesiredValue.ToString() -Force | Out-Null
						
						
					}
					else
					{
						Remove-ItemProperty -Path $OfficeRegistryPaths.ManagedAddins -Name $ProgID -ErrorAction SilentlyContinue
					}
				}
			}
			[gc]::Collect()
			if ($UnloadHive)
			{
				Close-RegistryHive -MountPoint $PSUserRegistryMountPoint
			}
		}
		else
		{
			Write-Verbose ('[{0}] Confirm={1} ConfirmPreference={2} WhatIf={3} WhatIfPreference={4}' -f $MyInvocation.MyCommand, $Confirm, $ConfirmPreference, $WhatIf, $WhatIfPreference)
		}
	}
}