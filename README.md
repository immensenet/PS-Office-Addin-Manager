# PS-Office-Addin-Manager
PowerShell Module for managing Office Addins, specifically those disabled by resiliency

## Installation

```
Install-Module OfficeAddinManager
```
## Usage

### Get-OfficeAddin

#### Get Outlook Addins
```
Get-OfficeAddin -OfficeProduct Outlook
```
```
UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Outlook
ProgID             : Microsoft.VbaAddinForOutlook.1
FriendlyName       : Microsoft VBA for Outlook Addin
Location           : outlvba.dll
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           : 0
InDoNotDisableList : False
Loaded             : True
LoadAtStartup      : False
LoadOnDemand       : True
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Outlook
ProgID             : OneNote.OutlookAddin
FriendlyName       : OneNote Notes about Outlook Items
Location           : C:\Program Files (x86)\Microsoft Office\root\Office16\ONBttnOL.dll
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           : 16
InDoNotDisableList : False
Loaded             : True
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Outlook
ProgID             : OscAddin.Connect
FriendlyName       : Outlook Social Connector 2016
Location           : C:\Program Files (x86)\Microsoft Office\root\Office16\SOCIALCONNECTOR.DLL
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           : 47
InDoNotDisableList : False
Loaded             : True
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Outlook
ProgID             : UmOutlookAddin.FormRegionAddin
FriendlyName       : Microsoft Exchange Add-in
Location           : C:\Program Files (x86)\Microsoft Office\root\Office16\ADDINS\UmOutlookAddin.dll
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           : 47
InDoNotDisableList : False
Loaded             : True
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Outlook
ProgID             : Apple.DAV.Addin
FriendlyName       : iCloud Outlook Add-in
Location           : C:\Program Files (x86)\Common Files\Apple\Internet Services\APLZOD32.dll
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           : 141
InDoNotDisableList : False
Loaded             : True
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Outlook
ProgID             : MimecastServicesForOutlook.AddinModule
FriendlyName       : Mimecast for Outlook
Location           : C:\Program Files (x86)\Mimecast\Mimecast Outlook Add-In\adxloader.dll
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           : 281
InDoNotDisableList : True
Loaded             : True
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Outlook
ProgID             : OutlookAddin.OutlAddin
FriendlyName       : Foxit PDF Creator COM Add-in
Location           : C:\Program Files (x86)\Foxit Software\Foxit PhantomPDF\plugins\Creator\x86\OutLookAddin_x86.dll
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           : 62
InDoNotDisableList : False
Loaded             : True
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Outlook
ProgID             : AccessAddin.DC
FriendlyName       : Microsoft Access Outlook Add-in for Data Collection and Publishing
Location           : C:\Program Files (x86)\Microsoft Office\root\Office16\ADDINS\ACCOLK.DLL
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           :
InDoNotDisableList : False
Loaded             : False
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Outlook
ProgID             : ColleagueImport.ColleagueImportAddin
FriendlyName       : Microsoft SharePoint Server Colleague Import Add-in
Location           : C:\Program Files (x86)\Microsoft Office\root\Office16\ADDINS\ColleagueImport.dll
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           : 16
InDoNotDisableList : False
Loaded             : True
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Outlook
ProgID             : G2MAddin.OutlookAddin
FriendlyName       : GoToMeeting Outlook COM Addin
Location           : C:\Users\dkattan\AppData\Local\GoToMeeting\16786\G2MOutlookAddin.dll
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           : 15
InDoNotDisableList : False
Loaded             : False
LoadAtStartup      : False
LoadOnDemand       : False
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Outlook
ProgID             : TeamsAddin.FastConnect
FriendlyName       : Microsoft Teams Meeting Add-in for Microsoft Office
Location           : C:\Users\dkattan\AppData\Local\Microsoft\TeamsMeetingAddin\1.0.20031.2\x86\Microsoft.Teams.AddinLo
                     ader.dll
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           : 78
InDoNotDisableList : False
Loaded             : True
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False
```

#### Get Excel Addins
```
Get-OfficeAddin -OfficeProduct Outlook
```
```
UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Excel
ProgID             : ExcelPlugInShell.PowerMapConnect
FriendlyName       : Microsoft Power Map for Excel
Location           : C:\Program Files (x86)\Microsoft Office\root\Office16\ADDINS\Power Map Excel Add-in\EXCELPLUGINSHELL.DLL
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           :
InDoNotDisableList : False
Loaded             : False
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Excel
ProgID             : MicrosoftDataStreamerforExcel
FriendlyName       : Microsoft Data Streamer for Excel
Location           : C:\Program Files (x86)\Microsoft Office\Root\Office16\ADDINS\EduWorks Data Streamer Add-In\MicrosoftDataStreamerforExcel.vsto|vstolocal
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           :
InDoNotDisableList : False
Loaded             : False
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Excel
ProgID             : NativeShim.InquireConnector.1
FriendlyName       : Inquire
Location           : C:\Program Files (x86)\Microsoft Office\root\Office16\DCF\NativeShim.dll
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           :
InDoNotDisableList : False
Loaded             : False
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Excel
ProgID             : ExcelAddin.ExcelAddinPH
FriendlyName       : Foxit PDF Creator COM Add-in
Location           : C:\Program Files (x86)\Foxit Software\Foxit PhantomPDF\plugins\Creator\x86\FPC_ExcelAddin_x86.dll
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           : 953
InDoNotDisableList : False
Loaded             : False
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Excel
ProgID             : AdHocReportingExcelClientLib.AdHocReportingExcelClientAddIn.1
FriendlyName       : Microsoft Power View for Excel
Location           : C:\Program Files (x86)\Microsoft Office\root\Office16\ADDINS\Power View Excel Add-in\AdHocReportingExcelClient.dll
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           :
InDoNotDisableList : False
Loaded             : False
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False

UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Excel
ProgID             : PowerPivotExcelClientAddIn.NativeEntry.1
FriendlyName       : Microsoft Power Pivot for Excel
Location           : C:\Program Files (x86)\Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in\PowerPivotExcelClientAddIn.dll
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           :
InDoNotDisableList : False
Loaded             : False
LoadAtStartup      : True
LoadOnDemand       : False
LoadOnce           : False
```

#### Get Specific Outlook Addin
```
Get-OfficeAddin -OfficeProduct Outlook -ProgID G2MAddin.OutlookAddin
```
```
UserName           : DESKTOP-8HG67X\Darren
OfficeProduct      : Outlook
ProgID             : G2MAddin.OutlookAddin
FriendlyName       : GoToMeeting Outlook COM Addin
Location           : C:\Users\dkattan\AppData\Local\GoToMeeting\16786\G2MOutlookAddin.dll
Managed            : Unmanaged
InDisabledList     : False
InCrashingList     : False
LoadTime           : 15
InDoNotDisableList : False
Loaded             : False
LoadAtStartup      : False
LoadOnDemand       : False
LoadOnce           : False
```

### Set-OfficeAddin
#### Set Addin to be always enabled and clear resiliency data like crash counters
```
Get-OfficeAddin -OfficeProduct Outlook -ProgID G2MAddin.OutlookAddin | Set-OfficeAddin -ManagedBehavior AlwaysEnabled -ClearResiliencyData
```
