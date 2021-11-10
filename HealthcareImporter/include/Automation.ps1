Push-Location $(Split-Path $Script:MyInvocation.MyCommand.Path)
import-module -name ".\AutomationPrintQueue.ps1" -Force -Verbose
import-module -name ".\AutomationPort.ps1" -Force -Verbose
import-module -name ".\AutomationSAP.ps1" -Force -Verbose
#justatest
Class PrinTaurusAutomation
{      
    [void] StartXML($xmlWriter) {
    # Create The Document
    #$XmlWriter = New-Object System.XMl.XmlTextWriter("D:\Skripte\AutoQ\XMLCreator\include\test.xml",$Null)

    # Set The Formatting
    $xmlWriter.Formatting = "Indented"
    $xmlWriter.Indentation = "4"

    # Write the XML Decleration
    $xmlWriter.WriteStartDocument()
 
    # Set the XSL
    #$XSLPropText = "type='text/xsl' href='style.xsl'"
    #$xmlWriter.WriteProcessingInstruction("xml-stylesheet", $XSLPropText)
 
    # Write Root Element
    $xmlWriter.WriteStartElement("Jobs")
    $xmlWriter.WriteAttributeString("xmlns", "http://www.aki-gmbh.com")
    $xmlWriter.WriteAttributeString("xsi:schemaLocation", "http://www.aki-gmbh.com automate.xsd")
    $xmlWriter.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")

    }

    [void] EndXML($xmlWriter) {

    # $xmlWriter.WriteEndElement() # <-- Closing JobQueueElement
    # $xmlWriter.WriteEndElement() # <-- Closing JobsElement
    # End the XML Document
    $xmlWriter.WriteEndDocument()
 
    # Finish The Document
    $xmlWriter.Finalize
    $xmlWriter.Flush()
    $xmlWriter.Close()
    }
}

class NewAutomation : PrinTaurusAutomation {
[string] 
$executeDateTime
[string] 
$server
[string] 
$jobName
[string] 
$onErrorFinal
    
    [void] Job($xmlwriter) {

    $xmlWriter.WriteStartElement("Job")
    $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
    $xmlWriter.WriteAttributeString("Date", $this.executeDateTime)
    $xmlWriter.WriteAttributeString("Server", $this.server)
    $xmlWriter.WriteAttributeString("Name", $this.jobName)
    $xmlWriter.WriteAttributeString("OnError",$this.onErrorFinal)

    }

    [void] ERPJob($xmlwriter) {

    $xmlWriter.WriteStartElement("ErpJob")
    $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
    $xmlWriter.WriteAttributeString("Date", $this.executeDateTime)
    $xmlWriter.WriteAttributeString("Server", $this.server)
    $xmlWriter.WriteAttributeString("Name", $this.jobName)
    $xmlWriter.WriteAttributeString("OnError",$this.onErrorFinal)

    }

    [void] EndJob($xmlwriter) {

    $xmlWriter.WriteEndElement()
    }
    
    [void] EndErpJob($xmlwriter) {

    $xmlWriter.WriteEndElement()
    }
    }

class AutoQ : PrinTaurusAutomation {
    [string]
    $user
    [string]
    $password
    [string]
    $outputFile
    [string]
    $autoqPath
    [string]
    $serviceuri

    [void] run() {
        if ($this.serviceuri) {

        if (($this.user -ne "") -and ($this.password -ne "") -and ($this.outputFile -ne "") -and ($this.autoqPath)) {
            try {
                & $this.autoqPath @("-f",$this.outputfile,"-u", $this.user,"-p", $this.password, "-su", $this.serviceuri)
            } catch { 
            }
            }
        }

        if (!$this.serviceuri) {

        if (($this.user -ne "") -and ($this.password -ne "") -and ($this.outputFile -ne "") -and ($this.autoqPath)) {
            try {
                & $this.autoqPath @("-f",$this.outputfile,"-u", $this.user,"-p", $this.password)
               
            } catch { }
            }
        }
           
            
            
    }
}
