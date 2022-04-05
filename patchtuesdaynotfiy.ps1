## set variables
$emailfrom = "example@site.com"
$emailto = "example@site.com"
$smtpserver = "server"
$emailsubject = "NOTICE: UPDATE SCHEDULE FOR SERVICE INTERRUPTIONS"

## configure html email template
$emailtemplate = "
<html>
    <style>
        H1
            {background:yellow;margin-bottom:15.0pt;margin-left:0in;font-size:26.0pt;font-family:'Calibri Light',sans-serif;letter-spacing:.25pt;text-align:center;}
        Body
            {margin-top:12.0pt;
            margin-bottom:12.0pt;
            font-size:14.0pt;
            font-family:'Calibri',sans-serif;}

        h2{margin-top:12.0pt;
        margin-bottom:0pt;
            font-size:18.0pt;
            font-family:'Calibri',sans-serif;
            font-weight:bold;
            color:#ff0000;}	
    </style>

    <body>
        <H1>Servers Service Interruption</H1>
        NOTICE FOR SERVER INTERRUPTIONS. Please expect the following outages.
        <h2>SERVICES AFFECTED: </h2><br><br>
        SCHEDULE:<BR>
        Tuesday_2nd : SERVERS<BR>
        Wednesday_2nd : SERVERS<BR>
        Thursday_2nd : SERVERS<BR>
        Friday_2nd : SERVERS<BR>
        Monday_2nd : SERVERS<BR><BR>

        Tuesday_3rd : SERVERS<BR>
        Wednesday_3rd : SERVERS<BR>
        Thursday_3rd : SERVERS<BR>
        Friday_3rd : SERVERS<BR>
        Monday_3rd : SERVERS<BR><BR>

        Tuesday_4th : SERVERS<BR>
        Wednesday_4th : SERVERS<BR>
        Thursday_4th : SERVERS<BR>
        Friday_4th : SERVERS<BR>
        Monday_4th : SERVERS<BR><BR>

        Tuesday_5th : SERVERS<BR>
        Wednesday_5th : SERVERS<BR>
        Thursday_5th : SERVERS<BR>
        Friday_5th : SERVERS<BR>
        Monday_5th : SERVERS<BR><BR>
    </body>
</html>
"

## begin functions for date calculation
function get-patchtuesday {
param([datetime]$date)
    switch ($date.DayOfWeek){
        "Sunday"    {$patchTuesday = $date.AddDays(9); break} 
        "Monday"    {$patchTuesday = $date.AddDays(8); break} 
        "Tuesday"   {$patchTuesday = $date.AddDays(7); break} 
        "Wednesday" {$patchTuesday = $date.AddDays(13); break} 
        "Thursday"  {$patchTuesday = $date.AddDays(12); break} 
        "Friday"    {$patchTuesday = $date.AddDays(11); break} 
        "Saturday"  {$patchTuesday = $date.AddDays(10); break} 
     }
	$patchTuesday
}
## find patch tuesday for the current month
$currPT = Get-PatchTuesday (Get-Date -Day 1)


## set varaibles for email template for each 2nd/rd/4th/5th weekday in the month
$WeekDays = @(
    'Tuesday',
    'Wednesday',
    'Thursday',
    'Friday',
    'Saturday',
    'Sunday',
    'Monday'
)

$count = 0
## create the variables and replace in email template with date
Do {
    foreach ($weekday in $weekdays) {
        $value = get-date ($currPT.adddays(+($count))) -Format MM/dd/yy
        New-Variable -Name "$($weekday)_2nd" -Value $value -Force
        $emailtemplate = $emailtemplate.Replace("$($weekday)_2nd","Second $weekday $((Get-Variable -Name "$($weekday)_2nd").Value)")
        $count++
    }
}
while ($count -gt 0 -and $count -lt 6)
while ($count -gt 6 -and $count -lt 13){
    foreach ($weekday in $weekdays) {
        $value = get-date ($currPT.adddays(+($count))) -Format MM/dd/yy
        New-Variable -Name "$($weekday)_3rd" -Value $value -Force
        $emailtemplate = $emailtemplate.Replace("$($weekday)_3rd","Third $weekday $((Get-Variable -Name "$($weekday)_3rd").Value)")
        $count++
    }
}
While ($count -gt 13 -and $count -lt 20){
    foreach ($weekday in $weekdays) {
        $value = get-date ($currPT.adddays(+($count))) -Format MM/dd/yy
        New-Variable -Name "$($weekday)_4th" -Value $value -Force
        $emailtemplate = $emailtemplate.Replace("$($weekday)_4th","Fourth $weekday $((Get-Variable -Name "$($weekday)_4th").Value)")
        $count++
    }
}
While ($count -gt 20 -and $count -lt 26){
    foreach ($weekday in $weekdays) {
        $value = get-date ($currPT.adddays(+($count))) -Format MM/dd/yy
        New-Variable -Name "$($weekday)_5th" -Value $value -Force
        $emailtemplate = $emailtemplate.Replace("$($weekday)_5th","Fifth $weekday $((Get-Variable -Name "$($weekday)_5th").Value)")
        $count++
    }
}


## send email
Send-MailMessage -Subject $emailsubject -SmtpServer $smtpserver -Body $Emailtemplate -To $emailto -From $emailfrom -BodyAsHtml
