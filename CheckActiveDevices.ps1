Import-Module ImportExcel

#Email parameters
$subject="Devices Status Report"
$body="This is a test of the CheckActiveDevices script."
$from="noreply@mdlogistics.com"
$server="email.mdlogistics.com"

#Array containing the email addresses you wish to send the output to
$emailAddresses = @('sgyll@mdlogistics.com')

#Path where the file will be saved
$outputFile = ('C:\Temp\Device Status Report.xlsx')

#Initiates the arrays that will be populated below, then used to export to seperate sheets in the Excel file
$finalShire = @()
$finalGondor = @()
$finalMordor = @()
$finalRivendell = @()

#You can edit this sqlite query depending on what you want to search
$query = "SELECT name, type, description, manufacturer, model, ip_address, location FROM devices WHERE location LIKE '%Rivendell%' AND name LIKE '%Workstation%'"

#Gets data from SpiceWorks based on $query above
$spiceImport = getSpiceData -query $query

#Prints the query to the console
Write-Host 'Getting device data using the following query:'
Write-Host -ForegroundColor Yellow "$query"

#For each device in the list pulled from SpiceWorks
foreach ($device in $spiceImport)
{
    $name = $device.name

    $testnet = ''
    #Tries to ping the device, then stores results in variable $testnet
    try
    {
        $testnet = Test-Connection $device.ip_address -Count 2 -ErrorAction Stop   
    } 
    #If unable to ping the computer, it displays that it is not connected.
    catch
    {
        Write-Host "$name is not connected" -ForegroundColor Red
        $device | Add-Member -MemberType NoteProperty -Name Status -Value OFFLINE    
    }

    #If the ping was successful, then...
    if($testnet)
    {
        Write-Host "$name is online" -ForegroundColor Green
        $device | Add-Member -MemberType NoteProperty -Name Status -Value ONLINE
    }

    #checks the location member of $device to match it to the corresponding location, then adds it to the appropriate array
    #If a location cannot be determined, the $device is processed as default, thereby adding it to $finalUnknown
    switch -Wildcard ($device.location)
    {
        '*Shire*'
        {
            $finalShire += $device
        }

        '*Gondor*'
        {
            $finalGondor += $device
        }

        '*Mordor*'
        {
            $finalMordor += $device
        }

        '*Rivendell*'
        {
            $finalRivendell += $device
        }

        default
        {
            $finalUnknown += $device
        }
    } #end switch($devicelocation)
} #end foreach ($device in $deviceImport)

#if $finalShire is not empty, output to excel with conditional formatting, else export the blank sheet anyway ###Could probably be cleaned up a lot
if($finalShire.Count -ne 0)
{
    $finalShire | Export-Excel $outputFile -WorkSheetname 'The Shire' -AutoSize -BoldTopRow -AutoFilter -ConditionalText $(
        New-ConditionalText OFFLINE darkred -BackgroundColor pink
        New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
        )
}
else
{
    $finalShire | Export-Excel $outputFile -WorkSheetname 'The Shire'
}

if($finalGondor.Count -ne 0)
{
    $finalGondor | Export-Excel $outputFile -WorkSheetname 'Kingdom of Gondor' -AutoSize -BoldTopRow -AutoFilter -ConditionalText $(
        New-ConditionalText OFFLINE darkred -BackgroundColor pink
        New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
        )
}
else
{
    $finalGondor | Export-Excel $outputFile -WorkSheetname 'Kingdom of Gondor'
}

if($finalMordor.Count -ne 0)
{
    $finalMordor | Export-Excel $outputFile -WorkSheetname 'Realm of Evil, Mordor' -AutoSize -BoldTopRow -AutoFilter -ConditionalText $(
        New-ConditionalText OFFLINE darkred -BackgroundColor pink
        New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
        )
}
else
{
    $finalMordor | Export-Excel $outputFile -WorkSheetname 'Realm of Evil, Mordor'
}

if($finalRivendell.Count -ne 0)
{
    $finalRivendell | Export-Excel $outputFile -WorkSheetname 'Elf city of Rivendell' -AutoSize -BoldTopRow -AutoFilter -ConditionalText $(
        New-ConditionalText OFFLINE darkred -BackgroundColor pink
        New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
        )
}
else
{
    $finalRivendell | Export-Excel $outputFile -WorkSheetname 'Elf city of Rivendell'
}

if($finalUnknown.Count -ne 0)
{
    $finalUnknown | Export-Excel $outputFile -WorkSheetname 'Unknown' -AutoSize -BoldTopRow -AutoFilter -ConditionalText $(
        New-ConditionalText OFFLINE darkred -BackgroundColor pink
        New-ConditionalText ONLINE darkgreen -BackgroundColor lightgreen
        )
}
else
{
    $finalUnknown | Export-Excel $outputFile -WorkSheetname 'Unknown'
}

#send email to each email address in $emailAddresses
<#
foreach ($email in $emailAddresses)
{
    Send-MailMessage -To $email -From $from -Body $body -SmtpServer $server -Subject $subject -Attachments $outputFile
}
#>
