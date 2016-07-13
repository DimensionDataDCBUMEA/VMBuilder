param (
    $FileName = $PSScriptRoot + "\sample.xlsx",
    $username = "caas-username-goes-here",
	$password = "caas-password-goes-here",
	$adminPass = "virtual-machine-admin-password-goes-here"
)

CLEAR

$ErrorActionPreference = 'stop'

function IsNumeric ($Value) {
    return $Value -match "^-?[0-9]\d*(\.\d+)?$"
}


function CheckMachine ($mchinename,$sleeptime) {

    # Sleep while the machine is still building...

    $startTime = (Get-Date)

    do
        { 
            Start-Sleep -Seconds $sleeptime
            $myserver = Get-CaasServer -name $mchinename
            $serverState = $myserver.state
            Write-Host "." -NoNewLine
            #Write-Host "  --   Current state: " -nonewline
            #Write-host $serverState -ForegroundColor Red
        }
    while ($serverState -ne "NORMAL")

    $endTime = (Get-Date)
    $ts = New-TimeSpan -Start $startTime -End $endTime
    $duration = $ts.totalseconds
    Write-Host " took " -NoNewLine
    Write-Host $duration -NoNewLine -ForegroundColor Green
    Write-Host " seconds"

}

Write-Host "Installing Excel Reader ..."

# Load EPPlus
$DLLPath = $PSScriptRoot + "\EPPlus\EPPlus.dll"
[Reflection.Assembly]::LoadFile($DLLPath) | Out-Null

$ExcelPackage = New-Object OfficeOpenXml.ExcelPackage -ArgumentList $FileName

Write-Host "Selecting Master List Worksheet ..."
$Worksheet = $ExcelPackage.Workbook.Worksheets[1]   # SELECT the Master List Sheet

$startCol = $Worksheet.Dimension.Start.Address
$endCol = $Worksheet.Dimension.End.Address


# now I have to figure out the startcolumn name and the endcolumn name
for ($i=0;$i -lt $startCol.Length;$i++) {
    # must check this char to see if it's numeric
    $char = $startCol.Substring($i,1)
    $numericCheck = IsNumeric($Char)
    if ($numericCheck) {
        $i = $i--
        $myCheck = $startCol
        $startCol = $myCheck.Substring(0,$i)
        $startRow = $myCheck.Substring($i,$myCheck.Length-$i)
        break
    }
}

# now I have to figure out the startcolumn name and the endcolumn name
for ($i=0;$i -lt $endCol.Length;$i++) {
    # must check this char to see if it's numeric
    $char = $endCol.Substring($i,1)
    $numericCheck = IsNumeric($Char)
    if ($numericCheck) {
        $i = $i--
        $myCheck = $endCol
        $endCol = $myCheck.Substring(0,$i)
        $endRow = $myCheck.Substring($i,$myCheck.Length-$i)
        break
    }
}


$mypassword = ConvertTo-SecureString –String $password –AsPlainText -Force
$credential = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $Username, $mypassword

New-CaasConnection -ApiCredentials $Credential -Region Africa_AF -Vendor DimensionData -Name TestBuilder | Out-null
Set-CaasActiveConnection -Name TestBuilder | out-null

    # Now we loop through all the rows of server records
    for ($i=2;$i -lt $Worksheet.Dimension.Rows+1;$i++) {

        $startPos = $startCol+$i
        $endPos = $endCol+$i

        $row = $Worksheet.Cells[$startPos+":"+$endPos].Value

        # right so now I have each row element in a $row array

        $nodename = $row[0,0]
        $machinename = $row[0,1]
        $ipaddress = $row[0,2]
        $templatename = $row[0,3]
        $networkdomain = $row[0,4]
        $vlanname = $row[0,5]
        $primarydns = $row[0,6]
        $secondarydns = $row[0,7]
        $timezone = $row[0,8]
        $vcpu = $row[0,9]
        $vcputype = $row[0,10]
        $vram = $row[0,11]
        $powerstate = $row[0,12]

        if ($vcputype -match 'HIGH Performance') { $vcpuType = "HIGHPERFORMANCE" }
        if ($vcputype -match 'Standard') { $vcpuType = "STANDARD" }

        $disksize = (0..5)
        $disktype = (0..5)

        $disksize[0] = $row[0,13]
        $disktype[0] = $row[0,14]
        $disksize[1] = $row[0,15]
        $disktype[1] = $row[0,16]
        $disksize[2] = $row[0,17]
        $disktype[2] = $row[0,18]
        $disksize[3] = $row[0,19]
        $disktype[3] = $row[0,20]
        $disksize[4] = $row[0,21]
        $disktype[4] = $row[0,22]
        $disksize[5] = $row[0,23]
        $disktype[5] = $row[0,24]

        Write-Host ""

        Write-host "Building ... " -NoNewline
        Write-Host $machinename -ForegroundColor Green -NoNewline
        Write-Host " on nodeID: " -NoNewline
        Write-Host $nodename -ForegroundColor Green

        Write-Host ""

        Write-Host "  -- vCPU: " -NoNewline
        Write-Host $vcpu -ForegroundColor Green
        Write-Host "  -- vCPU Type: " -NoNewline
        Write-Host $vcputype -ForegroundColor Green
        Write-Host "  -- vRAM: " -NoNewline
        Write-Host $vram -ForegroundColor Green
        Write-Host "  -- Primary DNS: " -NoNewline
        Write-Host $primarydns -ForegroundColor Green
        Write-Host "  -- Secondary DNS: " -NoNewline
        Write-Host $secondarydns -ForegroundColor Green
        Write-Host "  -- IP Address: " -NoNewline
        Write-Host $ipaddress -ForegroundColor Green

        Write-Host "  -- Network Domain: " -NoNewline
        Write-Host $networkdomain -ForegroundColor Green
        Write-Host "  -- VLAN: " -NoNewline
        Write-Host $vlanname -ForegroundColor Green

        Write-Host "  -- Powerstate: " -NoNewline
        Write-Host $powerstate -ForegroundColor Green

        Write-Host "  -- Template: " -NoNewline
        Write-Host $templatename -ForegroundColor Green

        for ($disk=0;$disk -lt 6;$disk++) {

            if ($disktype[$disk] -match "High Performance") { $disktype[$disk] = "HIGHPERFORMANCE" }
            if ($disktype[$disk] -match "Standard") { $disktype[$disk] = "STANDARD" }
            if ($disktype[$disk] -match "Economy") { $disktype[$disk] = "ECONOMY" }

            $dsksize = $disksize[$disk]
            $dsktype = $disktype[$disk]


            if ($dsksize) {

                Write-Host "  -- Disk $disk " -NoNewline
                Write-Host "$dsksize GB ($dsktype)" -ForegroundColor Green

            }

        }

        # lets figure out the powerstate stuff now....
        if ($powerstate -eq 'Off') { $powerstate = $false } else { $powerstate = $true }

        $serverImage = Get-CaasOsImage -Name $templatename -DataCenterId $nodename
        $networkDomain = Get-CaasNetworkDomain -NetworkDomainName $networkdomain
        $vlan = Get-CaasVlan -Name $vlanname

        # Hard code South African time zone into the servers....
        $timezone = 140

        $serverDetails = New-CaasServerDetails -Name $machineName -AdminPassword $adminPass -IsStarted $powerState -NetworkDomain $networkdomain -PrimaryVlan $vlan -PrimaryDns $primarydns -SecondaryDns $secondarydns -ServerImage $serverImage -CpuSpeed $vcputype -CpuCount $vcpu -MemoryGb $vram -MicrosoftTimeZone $timezone -CpuCoresPerSocket 1
        $serverDetails.PrivateIp = $ipaddress

        $server = New-CaasServer -ServerDetails $serverDetails 

        # Sleep while the machine is still building...
        Write-Host ""
        Write-Host "  -- Building Server " -NoNewLine
        CheckMachine $machineName 10

        # Now that the machine is built - I need to check the size of the disk(0) and resize if necessary. Only if the param is larger than the actual
        # Then I need to check the Tier for the disk
        # After that is completed - I need to add each of the other disks specified in the excel file.

        $myServer = Get-CaaSServer -name $machineName

        $myDiskSize = $myServer.Disk.SizeGB
        $myDiskType = $myServer.Disk.Speed

        # first we expand the disk if necessary...
        if ($myDiskSize -lt $disksize[0]) {
            # expand the disk to $disksize[0]
            Resize-CaasServerDisk -NewSizeInGB $diskSize[0] -ScsiId 0 -Server $myServer | Out-Null
            # Sleep while the machine is still building...
            Write-Host "  -- Resizing Boot disk " -NoNewLine
            CheckMachine $machineName 10
        }

        # then we storage vmotion if necessary...
        if ($myDiskType -notmatch $disktype[0]) {
            # storage vmotion the disk
            Set-CaasServerDiskSpeed -ScsiId 0 -Server $myServer -Speed $diskType[0] | Out-Null
            # Sleep while the machine is still building...
            Write-Host "  -- Changing Boot disk tier " -NoNewline
            #CheckMachine $machineName 20
        }

        # now we need to loop through all the "Other" drives and provision them if necessary on this machine.
        for ($disk=1;$disk -lt 6;$disk++) {

            $dsksize = $disksize[$disk]
            $dsktype = $disktype[$disk]

            if ($dsksize) {

                # Provision this disk now.
                Add-CaaSServerDisk -Server $myServer -SizeinGB $dsksize -Speed $dsktype | out-null
                # Sleep while the machine is still building...
                Write-Host "  -- Adding disk... $disk " -NoNewline
                CheckMachine $machineName 20

            }

        }

        # finally - lets enable monitoring on this server - just in case......
        # but essentials monitoring of course.
        Enable-CaaSServerMonitoring -server $myServer -ServicePlan ESSENTIALS | Out-Null
        Write-Host "  -- Enabling Essentials Monitoring " -NoNewline
        CheckMachine $machineName 5

        # It would be a good idea to now figure out how to remote execute commands on the machine...
        # Cisco Client VPN is a good option here... probably the only option actually!

    }

Remove-CaasConnection -Name TestBuilder
Write-Host ""
Write-Host "Process complete ..."


