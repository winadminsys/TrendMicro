# List interfaces
#$Adapters = .\tcpdump.exe -D

push-location

# Get adapters with IPEnabled
$Adapters = @()
$Adapters = Get-NetAdapter -IncludeHidden | ? {$_.Status -eq "up" -and $_.MacAddress -ne ""} #| select Name, InterfaceDescription,DeviceID, DeviceName | ft
"Found : " + $Adapters.Count + " adapters `r`n"

$IPS = Get-NetIPAddress

Foreach ($Adapter in $Adapters)
{
    $Details = New-object PSObject

    "Processing ... `r`n" + $Adapter.DeviceName + " -- " + $Adapter.InterfaceDescription + " -- " + $Adapter.Name
    $AdapterID = $Adapter.DeviceName

    $ExePath = "C:\temp\tcpdump.exe"
    $Args = "-i "+ $AdapterID +" -nn -v -s 1500 -c 1 ether[20:2] == 0x2000"

    # Start process
    $Process = New-Object System.Diagnostics.ProcessStartInfo
    $Process.FileName = $ExePath
    $Process.RedirectStandardError = $true
    $Process.RedirectStandardOutput = $true
    $Process.UseShellExecute = $false
    $Process.Arguments =  $Args
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $Process
    $p.Start() | Out-Null
    "Checking CDP Packets ... `r`n"
    #$p.WaitForExit()
    $Timeout = new-timespan -Seconds 70
	$sw = [diagnostics.stopwatch]::StartNew()

    # Check elapsed time
    $TimedOut = $false
    While ( $Process = get-process | where {$_.ProcessName -eq "tcpdump"} )
    {
        start-sleep -seconds 2

        If ($sw.elapsed -gt $timeout)
        {
            # Kill current process
            "Time Out .... !!!!"
            Stop-Process $Process.Id
            $TimedOut = $true
            break
        }
    }


    If ($TimedOut)
    {
        "No data link found ! `r`n"
    }
    Else
    {

        $AssignedIp = $IPS | where {$_.InterfaceIndex -eq $Adapter.InterfaceIndex}
        
        $Details | Add-Member -Name "InterfaceIndex" -MemberType NoteProperty -Value $Adapter.InterfaceIndex
        $Details | Add-Member -Name "ID" -MemberType NoteProperty -Value $Adapter.DeviceName
        $Details | Add-Member -Name "Name" -MemberType NoteProperty -Value $Adapter.Name
        $Details | Add-Member -Name  "Description" -MemberType NoteProperty -Value $Adapter.InterfaceDescription
        $Details | Add-Member -Name "IPAddress" -MemberType NoteProperty -Value $AssignedIp.IPAddress
        $Details | Add-Member -Name "MACAddress" -MemberType NoteProperty -Value $Adapter.MACAddress
       
        $stdout = $p.StandardOutput.ReadToEnd()
        $stderr = $p.StandardError.ReadToEnd()
        #Write-Host "exit code: " + $p.ExitCode

        $NetworkInfos = $stdout.Split(“`n”)

        #$NetworkInfos
	
	    If ($NetworkInfos.Length -gt 2)
        {
		    #Clear-Host
		    #$NetworkInfos

            # Select properties you want to show
	        $Properties = @("Device-ID", "Platform", "VTP Management", "Management Addresses", "Port-ID", "Capability", "VLAN ID", "Duplex")

	        Foreach ($Info in $NetworkInfos)
	        {
	            Foreach ($Property in $Properties)
	            {
				    #If ($Info -like "*$Property*")
	                If ($Info -match $Property)
	                {
	                    [String]$Item = $Info | Select-String -AllMatches -Pattern $Property
	                    $Pos = $Item.LastIndexOf(":")
	                    $Right = $Item.Substring($Pos+1).Replace("'","").Trim()
	                    #"--- $Property : $Right"
	                    $Details | Add-Member -Name $Property -MemberType NoteProperty -Value $Right
	                }
	            }

	        }

            # Show details
            $Details
	    }
	    Else
	    {
		    "No data link found ! `r`n"
	    }
    }
}

pop-location
