#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# References
#   http://www.analistadeti.com/print-server-gerar-evento-de-impressao-event-viewer/
#   https://gallery.technet.microsoft.com/Script-to-generate-print-84bdcf69
#   http://www.thomasmaurer.ch/2011/04/powershell-run-mysql-querys-with-powershell/
#   https://blogs.technet.microsoft.com/wincat/2011/08/25/trigger-a-powershell-script-from-a-windows-event/
#   http://dev.mysql.com/downloads/connector/net/
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# Function to connect and query MySQL
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Function Run-MySQLQuery {
<#
.SYNOPSIS
   run-MySQLQuery
    
.DESCRIPTION
   By default, this script will:
    - Will open a MySQL Connection
	- Will Send a Command to a MySQL Server
	- Will close the MySQL Connection
	This function uses the MySQL .NET Connector or MySQL.Data.dll file
     
.PARAMETER ConnectionString
    Adds the MySQL Connection String for the specific MySQL Server
     
.PARAMETER Query
 
    The MySQL Query which should be send to the MySQL Server
	
.EXAMPLE
    C:\PS> run-MySQLQuery -ConnectionString "Server=localhost;Uid=root;Pwd=p@ssword;database=project;" -Query "SELECT * FROM firsttest" 
    
    Description
    -----------
    This command run the MySQL Query "SELECT * FROM firsttest" 
	to the MySQL Server "localhost" with the Credentials User: Root and password: p@ssword and selects the database project
         
.EXAMPLE
    C:\PS> run-MySQLQuery -ConnectionString "Server=localhost;Uid=root;Pwd=p@ssword;database=project;" -Query "UPDATE firsttest SET firstname='Thomas' WHERE Firstname like 'PAUL'" 
    
    Description
    -----------
    This command run the MySQL Query "UPDATE project.firsttest SET firstname='Thomas' WHERE Firstname like 'PAUL'" 
	to the MySQL Server "localhost" with the Credentials User: Root and password: p@ssword
	
.EXAMPLE
    C:\PS> run-MySQLQuery -ConnectionString "Server=localhost;Uid=root;Pwd=p@ssword;" -Query "UPDATE project.firsttest SET firstname='Thomas' WHERE Firstname like 'PAUL'" 
    
    Description
    -----------
    This command run the MySQL Query "UPDATE project.firsttest SET firstname='Thomas' WHERE Firstname like 'PAUL'" 
	to the MySQL Server "localhost" with the Credentials User: Root and password: p@ssword and selects the database project
    
#>
	Param(
        [Parameter(
            Mandatory = $true,
            ParameterSetName = '',
            ValueFromPipeline = $true)]
            [string]$query,   
		[Parameter(
            Mandatory = $true,
            ParameterSetName = '',
            ValueFromPipeline = $true)]
            [string]$connectionString
        )
	Begin {
		Write-Verbose "Starting Begin Section"		
    }
	Process {
		Write-Verbose "Starting Process Section"
		try {
			# load MySQL driver and create connection
			Write-Verbose "Create Database Connection"
			# You could also could use a direct Link to the DLL File
			# $mySQLDataDLL = "C:\scripts\mysql\MySQL.Data.dll"
			# [void][system.reflection.Assembly]::LoadFrom($mySQLDataDLL)
			[void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")
			$connection = New-Object MySql.Data.MySqlClient.MySqlConnection
			$connection.ConnectionString = $ConnectionString
			Write-Verbose "Open Database Connection"
			$connection.Open()
			
			# Run MySQL Querys
			Write-Verbose "Run MySQL Querys"
			$command = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connection)
			$dataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($command)
			$dataSet = New-Object System.Data.DataSet
			$recordCount = $dataAdapter.Fill($dataSet, "data")
			#$dataSet.Tables["data"] | Format-Table
            return $dataSet.Tables[“data”]
		}		
		catch {
			Write-Host "Could not run MySQL Query" $Error[0]	
		}	
		Finally {
			Write-Verbose "Close Connection"
			$connection.Close()
		}
    }
	End {
		Write-Verbose "Starting End Section"
	}
}


#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# Function to escape special chars
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
function EscapeSpecialChars([string]$string)
{
    $string = $string.Replace("`\","`\`\")
    $string = $string.Replace("'","`\'")
    return $string
}


#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# Variables = Parameter received
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
$EventID = $args[0]          # EventID
$TimeCreated = [datetime]$args[1]
$TimeCreated = $TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")  # TimeCreated
$JobID     = $args[2]        # JobID
$FileName  = $args[3]        # FileName
$User      = $args[4]        # User
$Client    = $args[5]        # Client
$Printer   = $args[6]        # Printer
$Address   = $args[7]        # Address
$JobBytes  = $args[8]        # JobBytes
$PageCount = $args[9]        # PageCount


#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# Read Log EventID 805
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
$StartDate = (Get-Date -Date $TimeCreated) - (New-Timespan -Second 1) 
$EndDate = (Get-Date -Date $TimeCreated) + (New-Timespan -Second 5) 

$ServerHostname = "localhost"
$PrintLog805 = Get-WinEvent -ErrorAction SilentlyContinue -ComputerName $ServerHostname -FilterHashTable @{ProviderName="Microsoft-Windows-PrintService"; StartTime=$StartDate; EndTime=$EndDate; ID=805;}

if ($UserName -gt ""){
    $DirectorySearcher = New-Object System.DirectoryServices.DirectorySearcher 
    $LdapFilter = "(&(objectClass=user)(samAccountName=${User}))" 
    $DirectorySearcher.Filter = $LdapFilter 
    $UserEntry = [adsi]"$($DirectorySearcher.FindOne().Path)" 
    $ADName = $UserEntry.displayName 
} 

$PrintNumberOfCopies = $PrintLog805 | Where-Object {$_.Message -like "Processando trabalho $JobID." -and $_.TimeCreated -ge $StartDate -and $_.TimeCreated -le $EndDate } 
	
if (($PrintNumberOfCopies | Measure-Object).Count -eq 1) {
    # retrieve the remaining fields from the event log contents 
    $logxml = [xml]$PrintNumberOfCopies.ToXml() 
    $NumberOfCopies = $logxml.Event.UserData.RenderJobDiag.Copies 
    # some flawed printer drivers always report 0 copies for every print job; output a warning so this can be investigated further and set copies to be 1 in this case as a guess of what the actual number of copies was 
    if ($NumberOfCopies -eq 0) { 
        $NumberOfCopies = 1 
    } 
} 
# otherwise, either no or more than 1 matching event log ID 805 record was found 
#   both cases are unusual error conditions so report the error but continue on, assuming one copy was printed 
else {
    $NumberOfCopies = 1 
} 	
# Calculate total pages printed 
$TotalPages = [int]$PageCount * [int]$NumberOfCopies 

#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# Save printlog to a file
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
$Data = Get-Date -f "yyyy.MM.dd"
$OutputFile="c:\printlog\LogEventID307_$Data.txt"

# If file log not exists, create a new one with headers
if (!(Test-Path $OutputFile)){
   Out-File  -filepath $OutputFile -inputobject "EventID ; JobID; TimeCreated ; FileName ; User ; Client ; Printer ; Address ; JobBytes ; PageCount ; NumberOfCopies ; TotalPages ; NumberOfCopies ; TotalPages"
   Out-File  -filepath $OutputFile -inputobject "$EventID ; $JobID; $TimeCreated ; `"$FileName`" ; $User ; $Client ; $Printer ; $Address ; $JobBytes ; $PageCount ; $NumberOfCopies ; $TotalPages ; $NumberOfCopies ; $TotalPages" -Append
}
else{
   Out-File  -filepath $OutputFile -inputobject "$EventID ; $JobID; $TimeCreated ; `"$FileName`" ; $User ; $Client ; $Printer ; $Address ; $JobBytes ; $PageCount ; $NumberOfCopies ; $TotalPages ; $NumberOfCopies ; $TotalPages" -Append
}

#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# MySQL Connection
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
$DBAddress="192.168.0.1"
$DBUser="printlog"
$DBPassword="printlog"
$DBName="printlog"
$DSN = "Server=$DBAddress;Uid=$DBUser;Pwd=$DBPassword;Database=$DBName;"

#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
# Executa a inserção dos dados no banco MySQL
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
$FileName = EscapeSpecialChars $FileName
$Client = EscapeSpecialChars $Client
$Printer = EscapeSpecialChars $Printer
$sql_query_insert = "INSERT INTO printlog (eventid, jobid, user, client, timecreated, filename, printer, address, jobbytes, pagecount, numberofcopies, totalpages) VALUES ('$EventID', '$JobID', '$User', '$Client', '$TimeCreated', '$FileName', '$Printer', '$Address', '$JobBytes', '$PageCount', '$NumberOfCopies', '$TotalPages');";
run-MySQLQuery -ConnectionString $DSN -Query $sql_query_insert
