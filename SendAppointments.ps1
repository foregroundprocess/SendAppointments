#SETTINGS
$recipient = "" #put your e-mail address here

function AddToLog ($LogRecord) {
    $LogFilePath = "$PSScriptRoot\SendAppointment.log"
    if ((Get-Item $LogFilePath).Length / 1KB -gt 100) {
        Remove-Item $LogFilePath 
        Add-Content -Path $LogFilePath -Value "$(Get-Date) - Removed previous log file"
    }
    Add-Content -Path $LogFilePath -Value "$(Get-Date) - $LogRecord"
}

function SendNotification ($Start, $End) {
    $now = Get-Date
    $TimeSpan = (New-TimeSpan -Start $now -End $Start).TotalMinutes
    if (($TimeSpan -lt 15) -and ($TimeSpan -gt 0)) {
        $Mail = $Outlook.CreateItem(0)
        $Mail.To = $recipient
        $Mail.Subject = "Work event from $(($Start).ToShortTimeString()) to $(($End).ToShortTimeString())"
        $Mail.Body = "Work event from $(($Start).ToShortTimeString()) to $(($End).ToShortTimeString())"
        $Mail.Send()
    }
}

$currentuser = Get-WmiObject -Class win32_computersystem | Select-Object -ExpandProperty username #Check if the PC is unlocked, otherwise stop the script
$process = get-process logonui -ea silentlycontinue
if (($currentuser -and $process) -eq $false) {
    Break
}

$outlook = New-Object -ComObject Outlook.Application
$mapi = $outlook.GetNamespace('MAPI')
$calendars = $mapi.GetDefaultFolder(9)

# Sort all calendar items by start time and just grab the subject and start time.
$nearappointments = $calendars.items #| select conversationtopic,start,duration,end,isrecurring,meetingstatus,recurrencestate | sort Start -Descending

#$nearappointments | Export-Csv -Path appointments.csv -NoTypeInformation -Delimiter ";"
$singleappointments = @()
$recurrenceappointments = @()
$nearappointments | ForEach-Object {
    if ($_.meetingstatus -ne 5 -and $_.meetingstatus -ne 7) {
        if ($_.IsRecurring -eq $false) {
            #collect single appointmens
            $singleappointments += $_
        }
        else {
            #collect recurrence appointmens
            $recurrenceappointments += $_
        }
    }
}

$todaysingleappointments = @()#filter out single appointments
$singleappointments | ForEach-Object {
    if ((($_.Start).Date) -eq ((Get-Date).Date)) {
        $todaysingleappointments += $_
    }
}

$todayrecurrenceappointments = @()#process recurring appointments
foreach ($recurrenceappointment in $recurrenceappointments) {
    $recurrenceappointment_StartTime = $recurrenceappointment.Start.TimeOfDay
    $recurrence = $recurrenceappointment.GetRecurrencePattern() #find recurrence properties

    $Occurrences = @()
    switch ($recurrence.RecurrenceType) {
        0 {
            for ($i = 0; $i -lt $recurrence.Occurrences; $i++) {
                $Occurrence = $recurrence.PatternStartDate.AddDays($i * $recurrence.Interval)
                $Occurrences += $Occurrence
            }
        }
        1 {
            for ($i = 0; $i -lt $recurrence.Occurrences; $i++) {
                $Occurrence = $recurrence.PatternStartDate.AddDays($i * $recurrence.Interval * 7)
                $Occurrences += $Occurrence
            } 
        }
        2 {
            for ($i = 0; $i -lt $recurrence.Occurrences; $i++) {
                $Occurrence = $recurrence.PatternStartDate.AddMonths($i * $recurrence.Interval)
                $Occurrences += $Occurrence
            } 
        }
        3 {
            for ($i = 0; $i -lt $recurrence.Occurrences; $i++) {
                $Occurrence = $recurrence.PatternStartDate.AddMonths($i * $recurrence.Interval)
                $Occurrences += $Occurrence
            } 
        }
        5 {
            for ($i = 0; $i -lt $recurrence.Occurrences; $i++) {
                $Occurrence = $recurrence.PatternStartDate.AddYears($i * $recurrence.Interval)
                $Occurrences += $Occurrence
            } 
        }
        6 {
            for ($i = 0; $i -lt $recurrence.Occurrences; $i++) {
                $Occurrence = $recurrence.PatternStartDate.AddYears($i * $recurrence.Interval)
                $Occurrences += $Occurrence
            } 
        }
        Default {AddToLog("Wrong recurrence")}
    }

    foreach ($Occurrence in $Occurrences) {
        $Occurrence = $Occurrence.AddSeconds($recurrenceappointment_StartTime.TotalSeconds)
        $singlerecurrenceappointment = @{
            Subject         = $recurrenceappointment.Subject
            RecurrenceStart = $Occurrence
        }
        if ($singlerecurrenceappointment.RecurrenceStart.Date -eq ((Get-Date).Date)) {
            $todayrecurrenceappointments += $singlerecurrenceappointment
        }
    }
      
}

AddToLog("Found $($todayrecurrenceappointments.Count) recurring and $($todaysingleappointments.Count) single appointments")
foreach ($todayrecurrenceappointment in $todayrecurrenceappointments) {
    $Start = $todayrecurrenceappointment.RecurrenceStart
    $End = $todayrecurrenceappointment.End
    SendNotification -Start $Start -End $End
    #Write-host "1 - $($todaysingleappointment.Subject) to $End"
}

foreach ($todaysingleappointment in $todaysingleappointments) {
    $Start = $todaysingleappointment.Start
    $End = $todaysingleappointment.End
    SendNotification -Start $Start -End $End
    #Write-host "2 - $($todaysingleappointment.Subject) to $End"
}