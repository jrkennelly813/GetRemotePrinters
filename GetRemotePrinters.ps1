Import-Module ActiveDirectory

Write-Host "This program is used to find shared printers mapped to remote computers" -ForegroundColor Green
Write-Host "For Domain, Campus, and target OU Input, Please enter exactly as seen in AD.." -ForegroundColor Yellow

# Credntials passed for PSSession
$Creds = Get-Credential

# Iterate through specific sections of AD to get computers
Function GetComps {
    # The user inputs here need to match what shows in AD in order to work properly
    $domain = Read-Host "Enter the domain (PCC-Domain or EDU-Domain)" 
    $site = Read-Host "Enter the campus" 
    $topOU = Read-Host "Enter the Target OU for the computers you wish to search" 

    $sb = "OU={0},OU={1},OU=PCC,DC={2},DC=pima,DC=edu" -f $topOU,$site,$domain

    $comps = Get-ADComputer -Filter * -SearchBase $sb

    return $comps.Name
}

$comps = GetComps

Write-Host "...Retrieving data..." -foregroundColor Yellow

# Hash Table holding results
$data = @{}

# Iterate through each passed connection computer to retrive printer info
foreach ($comp in $comps) {
    #Enable-PSRemoting -Force
    if (Test-Connection $comp -Count 1 -EA SilentlyContinue) {
        $session = New-PSSession -ComputerName $comp 
        if ($session -ne $null) {
            Enter-PSSession -ComputerName $comp -Credential $creds -EA SilentlyContinue
            -Wait 2
            $printers = Invoke-Command -ComputerName $comp -Credential $creds -ScriptBlock {Get-WMIObject Win32_Printer}
            if ($printers.ShareName -ne $null) {
                $key = $comp
                $value = [string]$printers.ShareName

                $data.Add($key, $value)
            }
            else {
                $key = $comp
                $value = "No Shared Printers Found"

                $data.Add($key, $value)
            }
        }
        else {
            Write-Host "Cannot Connect to " $comp
        }
    }
}

Write-Host "..Data has been successfully parsed.." -ForegroundColor Cyan

$fileName = Read-Host "Please enter a name for your file" 

# set a path to save the CSV
$location = Resolve-Path ~/Desktop/ 
$path = [string]$location + [string]$fileName + ".csv"

# Export the data collected to a CSV
$data.GetEnumerator() | Select-Object -Property key,value | Export-Csv -NoTypeInformation -Path $path

Write-Host 'Your file was saved to the location' -Separator " -> " $path -ForegroundColor yellow
Write-Host "All Processes Completed You May Now Exit." -ForegroundColor Green
