#SETTINGS
$recipient = "" #put your e-mail address here

function Add_to_log ($log_record) {
    $log_file_path = "$PSScriptRoot\SendAppointment.log"
    if ((Get-Item $log_file_path).Length / 1KB -gt 100) {
        Remove-Item $log_file_path 
        Add-Content -Path $log_file_path -Value "$(Get-Date) - Removed previous log file"
    }
    Add-Content -Path $log_file_path -Value "$(Get-Date) - $log_record"
}

function SendNotification ($start, $end) {
    $now = Get-Date
    $time_span = (New-timespan -start $now -end $start).TotalMinutes
    if (($time_span -lt 15) -and ($time_span -gt 0)) {
        $Mail = $Outlook.CreateItem(0)
        $Mail.To = $recipient
        $Mail.Subject = "Work event from $(($start).ToShortTimeString()) to $(($end).ToShortTimeString())"
        $Mail.Body = "Work event from $(($start).ToShortTimeString()) to $(($end).ToShortTimeString())"
        $Mail.Send()
    }
}

$current_user = Get-WmiObject -Class win32_computersystem | Select-Object -ExpandProperty username #Check if the PC is unlocked, otherwise stop the script
$process = get-process logonui -ea silentlycontinue
if (($current_user -and $process) -eq $false) {
    Break
}

$outlook = New-Object -ComObject Outlook.Application
$mapi = $outlook.GetNamespace('MAPI')
$calendars = $mapi.GetDefaultFolder(9)

# Sort all calendar items by start time and just grab the subject and start time.
$near_appointments = $calendars.items #| select conversationtopic,start,duration,end,isrecurring,meetingstatus,recurrencestate | sort start -Descending

#$_ | Export-Csv -Path appointments.csv -NoTypeInformation -Delimiter ";"
$single_appointments = @()
$recurrence_appointments = @()
$near_appointments | ForEach-Object {
    if ($_.meetingstatus -ne 5 -and $_.meetingstatus -ne 7) {
        if ($_.IsRecurring -eq $false) {
            #collect single appointmens
            $single_appointments += $_
        }
        else {
            #collect recurrence appointmens
            $recurrence_appointments += $_
        }
    }
}

$today_single_appointments = @()#filter out single appointments
$single_appointments | ForEach-Object {
    if ((($_.start).Date) -eq ((Get-Date).Date)) {
        $today_single_appointments += $_
    }
}

$today_recurrence_appointments = @()#process recurring appointments
foreach ($recurrence_appointment in $recurrence_appointments) {
    $recurrenceappointment_start_time = $recurrence_appointment.start.TimeOfDay
    $recurrence = $recurrence_appointment.GetRecurrencePattern() #find recurrence properties

    $occurrences = @()
    for ($i = 0; $i -lt $recurrence.occurrences; $i++) {
        switch ($recurrence.RecurrenceType) {
            0 { $occurrence = $recurrence.PatternStartDate.AddDays($i * $recurrence.Interval) }
            1 { $occurrence = $recurrence.PatternStartDate.AddDays($i * $recurrence.Interval * 7) }
            2 { $occurrence = $recurrence.PatternStartDate.AddMonths($i * $recurrence.Interval) }
            3 { $occurrence = $recurrence.PatternStartDate.AddMonths($i * $recurrence.Interval) }
            5 { $occurrence = $recurrence.PatternStartDate.AddYears($i * $recurrence.Interval) }
            6 { $occurrence = $recurrence.PatternStartDate.AddYears($i * $recurrence.Interval) }
            Default { Add_to_log("Wrong recurrence") }
        }
        $occurrences += $occurrence
    }

    foreach ($occurrence in $occurrences) {
        $occurrence = $occurrence.AddSeconds($recurrenceappointment_start_time.TotalSeconds)
        $single_recurrence_appointment = @{
            Subject          = $recurrence_appointment.Subject
            recurrence_start = $occurrence
        }
        if ($single_recurrence_appointment.RecurrenceStart.Date -eq ((Get-Date).Date)) {
            $today_recurrence_appointments += $single_recurrence_appointment
        }
    }
      
}

Add_to_log("Found $($today_recurrence_appointments.Count) recurring and $($today_single_appointments.Count) single appointments")
foreach ($today_recurrence_appointment in $today_recurrence_appointments) {
    $start = $today_recurrence_appointment.RecurrenceStart
    $end = $today_recurrence_appointment.end
    SendNotification -start $start -end $end
}

foreach ($today_single_appointment in $today_single_appointments) {
    $start = $today_single_appointment.start
    $end = $today_single_appointment.end
    SendNotification -start $start -end $end
}
