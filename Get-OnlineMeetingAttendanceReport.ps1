[CmdletBinding()]
param (
    [Parameter (Mandatory = $true,HelpMessage = 'Enter the Tenant ID')]
    [ValidateNotNullorEmpty()]
    [string]$TenantID,
    [Parameter (Mandatory = $true,HelpMessage = 'Enter the Azure App ID')]
    [ValidateNotNullorEmpty()]
    [string]$AppID,
    [Parameter (Mandatory = $true,HelpMessage = 'Enter the UPN of the organizer(s)')]
    [ValidateNotNullorEmpty()]
    [string[]]$OrganizerUPN,
    [Parameter (Mandatory = $true,HelpMessage = 'Enter the number of days to search for meetings organized by the listed organizer')]
    [ValidateNotNullorEmpty()]
    [int]$ReportInDays,
    [Parameter (Mandatory = $true,HelpMessage = 'Enter the thumbprint of the certificiate used to authenticate as the Azure App')]
    [ValidateNotNullorEmpty()]
    [string]$CertificateThumbprint
)

#Connect to Microsoft Graph
Connect-MgGraph -ClientId $AppID -TenantId $TenantID -CertificateThumbprint $CertificateThumbprint
If(!(Get-MgContext)) {
    Write-Error "Connection to Microsoft Graph was unsuccessful, exiting now..."
    return
}

#Get matching events
[System.DateTime]$EventStartDate = (Get-Date).AddDays(-$ReportInDays)
ForEach ($Organizer in $OrganizerUPN) {
    $MatchingEvents = Get-MgUserEvent -UserId $Organizer -Filter "Start/datetime ge `'$EventStartDate`'"
    Foreach ($Event in $MatchingEvents) {

        #Get Organizer information
        $OrganizerUser = Get-MgUser -UserId $Organizer

        #Get Attendance Report for each event
        $i = 0
        $JoinWebUrl = $Event.OnlineMeeting.JoinUrl
        $OnlineMeeting = Get-MgUserOnlineMeeting -UserId $OrganizerUser.Id -Filter "joinWebUrl eq `'$JoinWebUrl`'"
        $AttendanceReports = Get-MgUserOnlineMeetingAttendanceReport -UserId $OrganizerUser.Id -OnlineMeetingId $OnlineMeeting.Id
        Foreach($AttendanceReport in $AttendanceReports) {
            $AttendanceRecords = Get-MgUserOnlineMeetingAttendanceReportAttendanceRecord -UserId $OrganizerUser.Id -MeetingAttendanceReportId $AttendanceReport.Id -OnlineMeetingId $OnlineMeeting.Id
            $AttendanceRecordTable = @()
            Foreach($AttendanceRecord in $AttendanceRecords) {
                $Attendee = Get-MgUser -UserId $Attendancerecord.Id
                $AttendanceRecord | Add-Member -MemberType NoteProperty -Name "UserPrincipalName" -Value $Attendee.UserPrincipalName
                $AttendanceRecordTable += $AttendanceRecord
                #$AttendanceRecord | Select-Object id,@{Name="DisplayName";Expression={$_.Identity.DisplayName}},EmailAddress,Role,TotalAttendanceInSeconds,@{Name="JoinDateTime";Expression={$_.AttendanceIntervals.JoinDateTime}},@{Name="LeaveDateTime";Expression={$_.AttendanceIntervals.LeaveDateTime}} | Export-CSV -NoTypeInformation -Path $CSVName
                Remove-Variable Attendee
            }
            If($i -gt 0) {
                #CSV output location
                $CSVName = "$((Get-Date $Event.start.datetime -f s).replace(':','-'))_$($Event.Subject)_$i.csv"
            } Else {
                #CSV output location
                $CSVName = "$((Get-Date $Event.start.datetime -f s).replace(':','-'))_$($Event.Subject).csv"
            }
            $AttendanceRecordTable | Select-Object id,UserPrincipalName,@{Name="DisplayName";Expression={$_.Identity.DisplayName}},EmailAddress,Role,TotalAttendanceInSeconds,@{Name="TotalAttendanceInMinutes";Expression={$_.TotalAttendanceInSeconds / 60}},@{Name="JoinDateTime";Expression={$_.AttendanceIntervals.JoinDateTime}},@{Name="LeaveDateTime";Expression={$_.AttendanceIntervals.LeaveDateTime}} | Export-CSV -NoTypeInformation -Path "c:\temp\$CSVName"
            $i++
        }
        
        Remove-Variable CSVName,JoinWebUrl,OnlineMeeting,AttendanceReport,AttendanceRecords
    }
}

Disconnect-MgGraph
