#-----------------------------------------------------------------
# Importieren der Notwendigen Klassen
#-----------------------------------------------------------------
$mypath = Split-Path $MyInvocation.MyCommand.Path -Parent
write-host $mypath
import-module -name "$mypath\include\Automation.ps1" -Force -Verbose

#-----------------------------------------------------------------
# XMLWriter Instanzieren -- wird sp√§ter verwendet um die XML 
# zu schreiben
#-----------------------------------------------------------------

$importXML       = "$mypath\Template.xlsx"
$AutoQServiceUrl = "net.tcp://demo01:8080/PrinTaurusService"
$AutoQUser       = "autoq"
$AutoQPassword   = "Start123!"
$PrintServers    = @("Demo01")


$ExcelObj = New-Object -comobject Excel.Application
$ExcelWorkBook = $ExcelObj.Workbooks.Open($importXML)

$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("Assignments")
$rowcount=$ExcelWorkSheet.UsedRange.Rows.Count

# Lese in Array um Duplikate zu entfernen
$queues = @()
for($i=2;$i -le $rowcount;$i++) {
    $queues+=$ExcelWorkSheet.Columns.Item(1).Rows.Item($i).Text
}
$queues | select -Unique

foreach ($queuename in $queues) {
  
    
    #-----------------------------------------------------------------
    # Vergabe der Variablen f√ºr neuen PMCPort
    #-----------------------------------------------------------------
    #$newPort = [PrinTaurusAutomationPortTCPIP]::new()    # TCPIP Port 
    $newPort = [PrinTaurusAutomationPortPMC]::new()         # PMC Port

    $newPort.portName = $queuename + "_Port"
    #$newPort.hostName = "test"
    $newPort.PMC_CommandLine = """C:\Program Files\AKI\PrinTaurus PortMonitor\data\__Plugins\OneQueueRouter\OneQueueRouterClient.exe"" -l None -p ""%p"" -i %i -d ""%d"""
    #-----------------------------------------------------------------
    # Vergabe der Variablen f√ºr neue Druckerwarteschlangen
    #-----------------------------------------------------------------
    $newQueue = [PrinTaurusAutomationQueue]::new()
    $newQueue.queueName                                     = $queuename
    $newQueue.portName                                      = "PMC:" + $queuename + "_Port"
    $newQueue.driverName                                    = "Ghostscript PDF"
    #$newQueue.printProcessor                                = "WinPrint:::RAW"
    #$newQueue.comment                                       = "TestKommentar"
    #$newQueue.location                                      = "TestLocation"
    #$newQueue.shareName                                     = ""
    #$newQueue.availability                                  = ""
    #$newQueue.publishINAD                                   = ""
    #$newQueue.defaultPriority                               = 1
    #$newQueue.deleteSpoolfiles                              = ""
    #$newQueue.deleteSaveSpoolFiles                          = ""
    #$newQueue.HaltFailedSpoolfiles                          = ""
    #$newQueue.PrintFirstInSpooler                           = 1
    #$newQueue.KeepSpoolfilesAfterPrinting                   = 1
    #$newQueue.RawPrinting                                   = 1
    #$newQueue.LogQueueStopping                              = 1
    #$newQueue.EnableBidi                                    = 1
    #$newQueue.BranchOfficeDirectPrinting                    = 1                        
    #$newQueue.EMFDespoolingSetting                          = 1
    <#



    #-----------------------------------------------------------------
    # Schreiben der XML
    # initialisieren der Automation
    # erstellen des globalen Jobs
    # erstellen von unterjobs
    #-----------------------------------------------------------------
    #>
    


    foreach ($server in $PrintServers) {
    write-host "Create $queuename for server $server"
    if (!(Test-Path "$mypath\output")) {New-Item -Path "$mypath\output" -ItemType Directory}
    $outputXML = "$mypath\output\PrinTaurusHealthcareAutomation_" + $queuename + "_" + $server +".xml"
    $XmlWriter = New-Object System.XMl.XmlTextWriter($outputXML,$Null)

        $automation = [NewAutomation]::new()         # Instanzieren
        $automation.executeDateTime = (Get-Date -UFormat "%Y-%m-%dT00:00:00")
        
        $automation.server = $server
        $automation.onErrorFinal = "Proceed"
        $automation.StartXML($xmlWriter)            # Automation XML Start
        $automation.jobName = "CreateQueue_" + $queuename
            $automation.Job($xmlWriter)
            # Create PMCPort
                $newPort.sequenzeNumber = 1                # Kontrolle der Sequenz
                $newPort.Create($xmlwriter)                 #  Erstelle neuen Port
            # Create Queue
                $newQueue.sequenzeNumber = 2                #   Z√§hle Sequenz hoch
                $newQueue.Create($xmlWriter)                #  Erstelle neue Queue
                
            $automation.EndJob($XmlWriter)

        $automation.EndXML($xmlWriter)              #       Beende das XML
        
    }
}

function runAutoQ($outputXML) {

#-----------------------------------------------------------------
# AutoQ f¸r alle Files in Ordner ausf¸hren
#-----------------------------------------------------------------
        $AutoQ = [AutoQ]::new()
        $AutoQ.User = $AutoQUser
        $AutoQ.Password = $AutoQPassword
        $AutoQ.autoqPath = "$mypath\AutoQ.exe"
        $AutoQ.Outputfile = $outputXML
        $AutoQ.serviceuri = $AutoQServiceUrl
        $AutoQ.run() # Dieser Befehl f¸hrt AutoQ.exe aus
        }

foreach ($file in (Get-ChildItem -Path "$mypath\output")) {
runAutoQ $file.FullName
}

#-----------------------------------------------------------------
# Aufr√§umen der Variablen
#-----------------------------------------------------------------
$PrintServers = $Null
$AutoQ = $null
$AutoQUser     = $Null
$AutoQPassword = $Null
$mypath = $Null
$newport = $Null
$newQueue = $Null
$automation = $null
$newSAPDevice = $Null
$queues = $Null

$ExcelObj.Workbooks.Close()
$ExcelObj.Quit()    

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelWorkBook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelObj)

Remove-Variable -Name ExcelObj

Remove-Module Automation
#>
