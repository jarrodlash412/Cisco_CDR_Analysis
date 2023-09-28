<#
.SYNOPSIS
This script analyzes Cisco Call Manager CDR data to produce reports on both device and phone number usage.

.DESCRIPTION
The script contains functions to process CDR records and then writes the analysis to an Excel file. It focuses on two primary aspects:

1. Device Analysis (`GetCiscoCDRperDevice` function):
    - Analyzes call data records for devices.
    - Gathers stats for each device, including inbound/outbound call details, user details, and timestamps for the first and last call.
    - Outputs the analysis to an "Device Analysis" tab in the Excel sheet.

2. Phone Number Analysis (`GetCiscoCDRperPhoneNumber` function):
    - Analyzes call data records based on phone numbers.
    - Processes both originating and destination phone numbers, along with the associated users.
    - Additional details of this function's output need to be filled in, as the function's logic is incomplete in the provided script.

.NOTE
The script also contains a utility function, `Write-DataToExcel`, that is responsible for writing data to the specified Excel file and worksheet.
The `write-Errorlog` function logs errors to a specified file.


.LASTEDIT
20230929

#>

$date = get-date -Format "MM/dd/yyyy HH:mm"
$tabFreeze = "Device Analysis"

Function Write-DataToExcel {
    param ($filelocation, $details, $tabname, $tabcolor)
    Write-Host "Writing to file: $filelocation Tab: $tabname"
    $excelpackage = Open-ExcelPackage -Path $filelocation 
    $ws = Add-Worksheet -ExcelPackage $excelpackage -WorksheetName $tabname 
    $ws.Workbook.Worksheets[$ws.index].TabColor = $tabcolor
    if ($tabFreeze.Contains($tabname)) {
        $details | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter -FreezePane 2, 3
    }
    else {
        $details | Export-Excel -ExcelPackage $excelpackage -WorksheetName $ws -AutoSize -AutoFilter -FreezeTopRow
    }
    Clear-Variable details 
    Clear-Variable filelocation
    Clear-Variable tabname
    Clear-Variable TabColor
}

Function write-Errorlog {
    param ($logfile, $errordata, $msgData)
    $errordetail = '"' + $date + '","' + $msgData + '","' + $errordata + '"'
    Write-Host $errordetail
    $errordetail |  Out-File -FilePath $logname -Append 
    Clear-Variable errordetail, msgData
}

function GetCiscoCDRperDevice {
    param ($filelocation, $data )
    $starttime = get-date
    Write-Host "Running Device Analysis..."
    $device_stats = @{}
    $time_format = 'yyyy-MM-dd HH:mm:ss.fff'
    $row_counter = 1
    $baseDate = Get-Date "1970-01-01 00:00:00Z"

    # Loop through the CDR records
    foreach ($row in $data) {
        # Get the device name, call duration, and user
        $orig_device = $row.origDeviceName
        $dest_device = $row.destDeviceName
        $duration = [int]$row.duration
        $baseDate = Get-Date "1970-01-01 00:00:00Z"
        $timestamp = $baseDate.AddSeconds($row.dateTimeOrigination)
        $user = $null
        $partition = $null

        # Determine the user based on whether the call was inbound or outbound
        if ($row.origNodeId -eq $row.destNodeId) {
            $user = $row.callingPartyUnicodeLoginUserID
            $partition = $row.callingPartyNumberPartition
        }
        elseif ($row.origDeviceName -eq $orig_device) {
            $user = $row.callingPartyUnicodeLoginUserID
            $partition = $row.originalCalledPartyNumberPartition
        }
        elseif ($row.destDeviceName -eq $dest_device) {
            $user = $row.finalCalledPartyUnicodeLoginUserID
            $partition = $row.originalCalledPartyNumberPartition
        }

        # Check if the device already exists in the stats hashtable
        if ($device_stats.ContainsKey($orig_device) -eq $false) {
            $device_stats[$orig_device] = @{
                'inbound_calls'    = 0
                'inbound_minutes'  = 0
                'outbound_calls'   = 0
                'outbound_minutes' = 0
                'total_calls'      = 0
                'total_minutes'    = 0
                'user'             = $user
                'partition'        = $partition
                'first_call'       = $timestamp
                'last_call'        = $timestamp
            }
        }
        else {
            # Update the first and last call dates if necessary
            if ($timestamp -lt $device_stats[$orig_device]['first_call']) {
                $device_stats[$orig_device]['first_call'] = $timestamp
            }
            if ($timestamp -gt $device_stats[$orig_device]['last_call']) {
                $device_stats[$orig_device]['last_call'] = $timestamp
            }
        }

        # Calculate the call direction and update the stats
        if ($row.origNodeId -eq $row.destNodeId) {
            $device_stats[$orig_device]['outbound_calls'] += 1
            $device_stats[$orig_device]['outbound_minutes'] += $duration
            $device_stats[$orig_device]['total_calls'] += 1
            $device_stats[$orig_device]['total_minutes'] += $duration
        }
        else {
            $device_stats[$orig_device]['inbound_calls'] += 1
            $device_stats[$orig_device]['inbound_minutes'] += $duration
            $device_stats[$orig_device]['total_calls'] += 1
            $device_stats[$orig_device]['total_minutes'] += $duration
        }
    }

    # Convert hashtable to an array for sorting
    $results = @()
    foreach ($device in $device_stats.Keys) {
        $user = $device_stats[$device]['user']
        if ($user -eq '\ ') {
            $user = $null
        }
        $results += [PSCustomObject]@{
            'Row'              = $row_counter++
            'Device'           = $device
            'User'             = $user
            'Partition'        = $partition
            'Inbound Calls'    = $device_stats[$device]['inbound_calls']
            'Inbound Minutes'  = [timespan]::FromSeconds($device_stats[$device]['inbound_minutes']).ToString('mm\:ss')
            'Outbound Calls'   = $device_stats[$device]['outbound_calls']
            'Outbound Minutes' = [timespan]::FromSeconds($device_stats[$device]['outbound_minutes']).ToString('mm\:ss')
            'Total Calls'      = $device_stats[$device]['total_calls']
            'Total Minutes'    = [timespan]::FromSeconds($device_stats[$device]['total_minutes']).ToString('mm\:ss')
            'First Call'       = $device_stats[$device]['first_call'].ToString($time_format)
            'Last Call'        = $device_stats[$device]['last_call'].ToString($time_format)
        }
    }

    # Sort
    $results = $results | Sort-Object -Descending 'Total Calls', 'Total Minutes'

    if ($results.count -gt 0) {
        $Details = @()
        Foreach ($result in $results) {
            $detail = New-Object PSObject
            $detail | Add-Member NoteProperty -Name "Device" -Value $result.'Device'
            $detail | Add-Member NoteProperty -Name "User" -Value $result.'User'
            $detail | Add-Member NoteProperty -Name "Partition" -Value $result.'Partion'
            $detail | Add-Member NoteProperty -Name "Inbound Calls" -Value $result.'Inbound Calls'
            $detail | Add-Member NoteProperty -Name "Outbound Calls" -Value $result.'Outbound Calls'
            $detail | Add-Member NoteProperty -Name "Inbound Minutes" -Value $result.'Inbound Minutes'
            $detail | Add-Member NoteProperty -Name "Outbound Minutes" -Value $result.'Outbound Minutes'
            $detail | Add-Member NoteProperty -Name "Total Calls" -Value $result.'Total Calls'
            $detail | Add-Member NoteProperty -Name "Total Minutes" -Value $result.'Total Minutes'
            $detail | Add-Member NoteProperty -Name "First Call" -Value $result.'First Call'
            $detail | Add-Member NoteProperty -Name "Last Call" -Value $result.'Last Call'

            $Details += $detail
        }
    }
    else { $details = "No data to display" }

    $Details | Export-Excel -Path $filelocation -WorksheetName "Device Analysis" -AutoFilter -AutoSize
    $excel = Open-ExcelPackage -Path $filelocation 
    $Green = "Green"
    $Green = [System.Drawing.Color]::$green 
    $excel.Workbook.Worksheets[1].TabColor = $Green  
    Close-ExcelPackage -ExcelPackage $excel
    Clear-Variable details
    Clear-Variable excel
    Clear-Variable green
}

function GetCiscoCDRperPhoneNumber {
    param ($filelocation, $data )
    $starttime = get-date
    Write-Host "Running Phone Number Analysis..."
    $phone_stats = @{}
    $time_format = 'yyyy-MM-dd HH:mm:ss.fff'
    $row_counter = 1
    $baseDate = Get-Date "1970-01-01 00:00:00Z"

    # Loop through the CDR records
    foreach ($row in $data) {
        # Get the phone numbers
        $orig_number = $row.originalCalledPartyNumber
        $dest_number = $row.finalCalledPartyNumber
        $duration = [int]$row.duration
        $timestamp = $baseDate.AddSeconds($row.dateTimeOrigination)
        $user = $null
        $partition = $null

        # Determine the user based on whether the call was inbound or outbound
        if ($row.origNodeId -eq $row.destNodeId) {
            $user = $row.callingPartyUnicodeLoginUserID
        }
        elseif ($row.origDeviceName -eq $orig_number) {
            $user = $row.callingPartyUnicodeLoginUserID
        }
        elseif ($row.destDeviceName -eq $dest_number) {
            $user = $row.finalCalledPartyUnicodeLoginUserID
            $partition = $row.originalCalledPartyNumberPartition
        }

        # Check if the phone number already exists in the stats hashtable
        if ($phone_stats.ContainsKey($orig_number) -eq $false) {
            $phone_stats[$orig_number] = @{
                'inbound_calls'    = 0
                'inbound_minutes'  = 0
                'outbound_calls'   = 0
                'outbound_minutes' = 0
                'total_calls'      = 0
                'total_minutes'    = 0
                'user'             = $user
                'partition'        = $partition
                'first_call'       = $timestamp
                'last_call'        = $timestamp
            }
        }
        else {
            # Update the first and last call dates if necessary
            if ($timestamp -lt $phone_stats[$orig_number]['first_call']) {
                $phone_stats[$orig_number]['first_call'] = $timestamp
            }
            if ($timestamp -gt $phone_stats[$orig_number]['last_call']) {
                $phone_stats[$orig_number]['last_call'] = $timestamp
            }
        }

        # Calculate the call direction and update the stats
        if ($row.origNodeId -eq $row.destNodeId) {
            $phone_stats[$orig_number]['outbound_calls'] += 1
            $phone_stats[$orig_number]['outbound_minutes'] += $duration
            $phone_stats[$orig_number]['total_calls'] += 1
            $phone_stats[$orig_number]['total_minutes'] += $duration
        }
        else {
            $phone_stats[$orig_number]['inbound_calls'] += 1
            $phone_stats[$orig_number]['inbound_minutes'] += $duration
            $phone_stats[$orig_number]['total_calls'] += 1
            $phone_stats[$orig_number]['total_minutes'] += $duration
        }
    }

    # Convert hashtable to an array for sorting
    $results = @()
    foreach ($phone in $phone_stats.Keys) {
        $user = $phone_stats[$phone]['user']
        if ($user -eq '\ ') {
            $user = $null
        }
        $results += [PSCustomObject]@{
            'Row'              = $row_counter++
            'Phone Number'     = "'" + $phone
            'User'             = $user
            'Partition'        = $partition
            'Inbound Calls'    = $phone_stats[$phone]['inbound_calls']
            'Inbound Minutes'  = [timespan]::FromSeconds($phone_stats[$phone]['inbound_minutes']).ToString('mm\:ss')
            'Outbound Calls'   = $phone_stats[$phone]['outbound_calls']
            'Outbound Minutes' = [timespan]::FromSeconds($phone_stats[$phone]['outbound_minutes']).ToString('mm\:ss')
            'Total Calls'      = $phone_stats[$phone]['total_calls']
            'Total Minutes'    = [timespan]::FromSeconds($phone_stats[$phone]['total_minutes']).ToString('mm\:ss')
            'First Call'       = $phone_stats[$phone]['first_call'].ToString($time_format)
            'Last Call'        = $phone_stats[$phone]['last_call'].ToString($time_format)
        }
    }
       # Sort
       $results = $results | Sort-Object -Descending 'Total Calls', 'Total Minutes'

       if ($results.count -gt 0) {
           $Details = @()
           Foreach ($result in $results) {
               $detail = New-Object PSObject
               $detail | Add-Member NoteProperty -Name "Phone Number" -Value $result.'Phone Number'
               $detail | Add-Member NoteProperty -Name "User" -Value $result.'User'
               $detail | Add-Member NoteProperty -Name "Partition" -Value $result.'Partion'
               $detail | Add-Member NoteProperty -Name "Inbound Calls" -Value $result.'Inbound Calls'
               $detail | Add-Member NoteProperty -Name "Outbound Calls" -Value $result.'Outbound Calls'
               $detail | Add-Member NoteProperty -Name "Inbound Minutes" -Value $result.'Inbound Minutes'
               $detail | Add-Member NoteProperty -Name "Outbound Minutes" -Value $result.'Outbound Minutes'
               $detail | Add-Member NoteProperty -Name "Total Calls" -Value $result.'Total Calls'
               $detail | Add-Member NoteProperty -Name "Total Minutes" -Value $result.'Total Minutes'
               $detail | Add-Member NoteProperty -Name "First Call" -Value $result.'First Call'
               $detail | Add-Member NoteProperty -Name "Last Call" -Value $result.'Last Call'
   
               $Details += $detail
           }
       }
       else { $details = "No data to display" }
    $tabname = "Phone Number Analysis"
    $tabcolor = "Purple"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}


function GetDeviceToPhoneNumberMapping {
    param ($filelocation, $data)
    Write-Host "Generating device-to-phone mapping..."

    $device_to_phone = @{}

    foreach ($row in $data) {
        $dest_device = $row.destDeviceName
        $called_number = "#" + $row.finalCalledPartyNumber

        if (-not [string]::IsNullOrEmpty($called_number)) {
            if (-not $device_to_phone.ContainsKey($dest_device)) {
                $device_to_phone[$dest_device] = @{}
            }

            $device_to_phone[$dest_device][$called_number] = $true
        }
    }

    $results = @()
    foreach ($device in $device_to_phone.Keys) {
        $results += [PSCustomObject]@{
            'Device'        = $device
            'Phone Numbers' = ($device_to_phone[$device].Keys -join ', ')
            'Count'         = $device_to_phone[$device].Count
        }
    }

    # Sort
    $results = $results | Sort-Object -Descending 'Count', 'Device', 'Phone Numbers'

    if ($results.count -gt 0) {
        $Details = @()
        Foreach ($result in $results) {
            $detail = New-Object PSObject
            $detail | Add-Member NoteProperty -Name "Device" -Value $result.'Device'
            $detail | Add-Member NoteProperty -Name "Count" -Value $result.'Count'
            $detail | Add-Member NoteProperty -Name "Phone Numbers" -Value $result.'Phone Numbers'
            $Details += $detail
        }
    }
    else { $details = "No data to display" }

    $tabname = "Device to Phone Mapping"
    $tabcolor = "Blue"
    Write-DataToExcel $filelocation $Details $tabname $tabcolor
}





Function ProgressBar {
    param([int]$percent)
    Write-Progress -Activity "Processing records..." -PercentComplete $percent -Status "$percent% complete"
}

Clear-Host
Write-Host "This will create an Excel Spreadsheet."

$filename = Read-Host "Enter path to Cisco CDR file to process"

if ($filename) {
    $fileInfo = Get-Item $filename
    $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
    Write-Host "Input file size: $fileSizeMB MB"

    $dirlocation = Read-Host "Enter location to store report (i.e. c:\scriptout or press enter for this directory)"
    if (-not $dirlocation) {
        $dirlocation = (Get-Location).Path
    }   
    $directory = $dirlocation + "\Cisco_CDR_Reports"

    try { Resolve-Path -Path $directory -ErrorAction Stop }
    catch {
        try { new-item -path $directory -itemtype "Directory" -ErrorAction Stop }
        catch {
            $logfile, $errordata, $msgData
            $date = get-date -Format "MM/dd/yyyy HH:mm"
            $errordetail = $date + ", Error creating directory. ," + $directory + "," + $error[0].exception.message 
            Write-Host $errordetail
        }
    }

    Import-Module ImportExcel

    # Performance metrics start
    $startTime = Get-Date
    $startMemory = (Get-Process -Id $PID).WorkingSet / 1MB

    # Fetch System Info
    $cpu = Get-WmiObject -Class Win32_Processor
    $ram = Get-WmiObject -Class Win32_ComputerSystem
    $drive = Get-WmiObject -Class Win32_DiskDrive 

    Write-Host "CPU: $($cpu.Name) running at $($cpu.MaxClockSpeed) MHz."
    Write-Host "RAM: $($ram.TotalPhysicalMemory / 1GB) GB."

    Write-Host "Reading input file... " $filename
    
    # Read the CDR file
    $data = Import-Csv $filename

    # Performance metrics after file read
    $endTime = Get-Date
    $endMemory = (Get-Process -Id $PID).WorkingSet / 1MB
    $elapsedTime = [math]::Round(($endTime - $startTime).TotalSeconds)
    Write-Host "Reading completed in $elapsedTime seconds. Memory used: $($endMemory - $startMemory) MB."

    $filedate = Get-Date -Format "MM-dd-yyyy.HH.mm.ss"
    $filelocation = $directory + "\CiscoCDRAnalysis-" + $filedate + ".xlsx"
    Write-Host "Saving to: " $filelocation

    # Start Timer
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Execute the functions
    GetCiscoCDRperDevice $filelocation $data
    GetDeviceToPhoneNumberMapping $filelocation $data
    GetCiscoCDRperPhoneNumber $filelocation $data

    # Stop Timer
    $stopwatch.Stop()

    # Output Results
    Write-Host "`nScript Execution Time: $($stopwatch.Elapsed.TotalSeconds) seconds."
}
