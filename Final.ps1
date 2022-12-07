# Powershell Final #
# Wyatt Tilque #
# Find Information about a set of computers

#Gets date for path name
$date=get-date -Format "MM-dd-yyyy"

#Header for formating
$header = "-----------------"


#Asks for name of person running
$name=Read-host -Prompt "Who is running this script?"

#Asks user for outout folder location and makes it if it's not there
$userpath=Read-host -Prompt "What folder do you want the Reports in? (Make sure you put an \ at the end of the path)"
if (-not (test-path '$userpath' -pathtype container)) #Tests if the directory is there and if not makes it
    {
      New-Item -Path $userpath -ItemType directory
    }

#Allows the user to import the CSV file
$usercsv=Read-host -Prompt "Where exactly is the CSV file?"
$cns = import-csv -path $usercsv

#For each loop so it goes through each computer
foreach ($cn in $cns)
    {
    #Loads the Logins for the computers into the loop
    $username=$cn.Username
    $password=ConvertTo-SecureString $cn.Password -AsPlainText -Force
    $userpass= New-Object System.Management.Automation.PSCredential($username,$password)


    #Makes the path to output the file to
    $path = $userpath + $cn.Computers + "-INV-" + $date +".txt"
    
    #Writes the name to the file
    "Report by " + "$name" | out-file $path
    "$header" | out-file $path -append

    #Writes the path to the file to the host
    write-host "Reports outputed to" "$path"

    #Write Hostname to Report
    $cn.Computers | out-file $path -append
    " " | out-file $path -Append

    #Gets the OS info
    "Windows Info" | out-file $path -append
    Invoke-Command -ComputerName $cn.Computers -Credential $userpass -ScriptBlock{Get-CimInstance -Class Win32_OperatingSystem} | Select-Object -Property Version | out-file $path -Append
    
    #Gets the Processor info
    "Processor Info" | out-file $path -append
    Invoke-Command -ComputerName $cn.Computers -Credential $userpass -ScriptBlock {Get-CimInstance -Class Win32_Processor} | Select-Object -Property Name,NumberOfLogicalProcessors,Manufacturer | out-file $path -Append
    
    #Ram Info
    "Installed Ram in GB" | out-file $path -Append
    "$header" | out-file $path -Append
    Invoke-Command -ComputerName $cn.Computers -Credential $userpass -ScriptBlock {(Get-CimInstance -Class Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1gb} | out-file $path -Append #Checks the computers installed memory and converts it GB
    " " | out-file $path -Append
    
    #Roles and Features and checks if the machine is a server or workstation
    $serws = Invoke-Command -ComputerName $cn.Computers -Credential $userpass -ScriptBlock {Get-CimInstance -ClassName Win32_OperatingSystem} #Gets OS info from the computer
    "Installed Roles and Features" | out-file $path -append
    "$header" | out-file $path -Append
    if ($serws.ProductType -eq 1) { #checks if the computer is a server or not
        "This is Not a server" | out-file $path -append
        } else {
         Invoke-Command -ComputerName $cn.Computers -Credential $userpass -ScriptBlock {Get-WindowsFeature | Where-Object {$_.Installed -match $True}} | Select-Object -Property Name | out-file $path -append
        }
     " " | out-file $path -append

    # Test Whether the Remote Desktop Protocol is Configures
    $rdp = Invoke-Command -ComputerName $cn.Computers -Credential $userpass -ScriptBlock {Get-ItemProperty -path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -name "fDenyTSConnections"} #Grabs a registry key to check Remote desktop protocol
    if ($rdp.fDenyTSConnections -eq 0){ #Tests if it's on or not
        "Remote Desktop Protocol is Enabled" | out-file $path -append
        }else{
        "Remote Desktop Protocol is Disabled" | out-file $path -append
        }
    " " | out-file $path -Append
    #Warnings and Errors for System and Application Event Log
    "Last 5 Warnings in System Event Log" | out-file $path -Append
    "$header" | out-file $path -Append
    Invoke-Command -ComputerName $cn.Computers -Credential $userpass -ScriptBlock {Get-EventLog -LogName System -EntryType Warning -Newest 5} | out-file $path -Append
    
    "Last 5 Errors in System Event Log" | out-file $path -Append
    "$header" | out-file $path -Append
    Invoke-Command -ComputerName $cn.Computers -Credential $userpass -ScriptBlock {Get-EventLog -LogName System -EntryType Error -Newest 5} | out-file $path -Append
    
    "Last 5 Warnings in Application Event Log" | out-file $path -Append
    "$header" | out-file $path -Append
    Invoke-Command -ComputerName $cn.Computers -Credential $userpass -ScriptBlock {Get-EventLog -LogName Application -EntryType Warning -Newest 5} | out-file $path -Append
    
    "Last 5 Errors in Application Event Log" | out-file $path -Append
    "$header" | out-file $path -Append
    Invoke-Command -ComputerName $cn.Computers -Credential $userpass -ScriptBlock {Get-EventLog -LogName Application -EntryType Error -Newest 5} | out-file $path -Append
    }