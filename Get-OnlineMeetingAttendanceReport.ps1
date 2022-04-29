param($Timer)
$ErrorActionPreference = 'Continue'

#Import Modules
Import-Module Microsoft.Graph.Authentication -SkipEditionCheck
Import-Module Microsoft.Graph.Calendar -SkipEditionCheck
Import-Module Microsoft.Graph.Users -SkipEditionCheck
Import-Module Microsoft.Graph.CloudCommunications -SkipEditionCheck
Import-Module PnP.PowerShell -SkipEditionCheck

#Declare Variables from environment variables
$TenantId = $ENV:TenantId #Ex: dfeeef05-adf7-4200-adf4-0034f2aaadd1
$AppId = $ENV:AppId #Ex: dfeeef05-adf7-4200-adf4-0034f2aaadd1
$ReportInDays = $ENV:ReportInDays #Ex: 7
$OrganizerUPN = $ENV:OrganizerUPN #Ex: MeganB@contoso.com,AdeleV@contoso.com
$SPOFolder = $ENV:SPOFolder #Ex: Shared Documents/Training/Attendance Reports
$SPOSiteURL = $ENV:SPOSiteURL #Ex: https://contoso.sharepoint.com/sites/Training
$CertificateThumbprint = $ENV:CertificateThumbprint #Ex: 8cc5edbc1cd9f61120fc6fdbefe1d8cd8f2759d5
$TempFolder = $ENV:Temp #This is a default environment variable, no change needed

#Connect to Microsoft Graph to pull meeting attendance reports
Connect-MgGraph -ClientId $AppID -TenantId $TenantID -CertificateThumbprint $CertificateThumbprint
If (!(Get-MgContext)) {
    Write-Error "Connection to Microsoft Graph was unsuccessful, exiting now..."
    return
}

#Connect to PnP, in order to move file to a SharePoint site
Connect-PnPOnline $SPOSiteURL -ClientId $AppId -Tenant $TenantId -Thumbprint $CertificateThumbprint

#Get matching events
[System.DateTime]$EventStartDate = (Get-Date).AddDays(-$ReportInDays)
ForEach ($Organizer in $OrganizerUPN) {
    
    #Get matching events, based on the start datetime of the meeting. If $ReportInDays is set to 1, then meetings from yesterday and today will return
    $MatchingEvents = Get-MgUserEvent -UserId $Organizer -Filter "Start/datetime ge `'$EventStartDate`'"
    Foreach ($Event in $MatchingEvents) {

        #Get Organizer information
        $OrganizerUser = Get-MgUser -UserId $Organizer

        #Get Attendance Report for each event
        $i = 0
        $JoinWebUrl = $Event.OnlineMeeting.JoinUrl
        $OnlineMeeting = Get-MgUserOnlineMeeting -UserId $OrganizerUser.Id -Filter "joinWebUrl eq `'$JoinWebUrl`'"
        $AttendanceReports = Get-MgUserOnlineMeetingAttendanceReport -UserId $OrganizerUser.Id -OnlineMeetingId $OnlineMeeting.Id
        Foreach ($AttendanceReport in $AttendanceReports) {
            
            #Meetings can have multiple attendance reports, usually if the same meeting is started and stopped more than once
            $AttendanceRecords = Get-MgUserOnlineMeetingAttendanceReportAttendanceRecord -UserId $OrganizerUser.Id -MeetingAttendanceReportId $AttendanceReport.Id -OnlineMeetingId $OnlineMeeting.Id
            $AttendanceRecordTable = @()
            Foreach ($AttendanceRecord in $AttendanceRecords) {
                
                #Lookup at the users in the attendance reports to provide additional information. Some information is not available for external attendees
                $Attendee = Get-MgUser -UserId $Attendancerecord.Id
                $AttendanceRecord | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $Attendee.UserPrincipalName
                $AttendanceRecordTable += $AttendanceRecord
                Remove-Variable Attendee
            }
            #Since there can be multiple attendance reports for the same meeting, the names need to be unique
            If ($i -gt 0) {
                #CSV output name
                $CSVName = "$((Get-Date $Event.start.datetime -f s).replace(':','-'))_$($Event.Subject)_$i.csv"
            }
            Else {
                #CSV output location
                $CSVName = "$((Get-Date $Event.start.datetime -f s).replace(':','-'))_$($Event.Subject).csv"
            }
            #Export to a CSV in a temp location in the Azure Function, so that it can be uploaded to SharePoint Online
            $AttendanceRecordTable | Select-Object id, UserPrincipalName, @{Name = "DisplayName"; Expression = { $_.Identity.DisplayName } }, EmailAddress, Role, TotalAttendanceInSeconds, @{Name = "TotalAttendanceInMinutes"; Expression = { $_.TotalAttendanceInSeconds / 60 } }, @{Name = "JoinDateTime"; Expression = { $_.AttendanceIntervals.JoinDateTime } }, @{Name = "LeaveDateTime"; Expression = { $_.AttendanceIntervals.LeaveDateTime } } | Export-CSV -NoTypeInformation -Path "$TempFolder\$CSVName"
            Add-PnPFile -Path "$TempFolder\$CSVName" -Folder $SPOFolder
            $i++
        }
        Remove-Variable CSVName, JoinWebUrl, OnlineMeeting, AttendanceReport, AttendanceRecords
    }
}

Disconnect-MgGraph
Disconnect-PnPOnline