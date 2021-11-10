Class PrinTaurusAutomationQueue
{
    [Guid]
    hidden $ID = (New-Guid).Guid
    
    [string]
    $queueName

    [string]
    $portName
    
    #[ValidateRange(0,100)]
    [int]
    $sequenzeNumber = 1
    
    #[ValidatePattern('^[a-z]')]
    #[ValidateLength(3,15)]
    [string]
    $driverName

    [string]
    $templateName

    [string]
    $printProcessor = "WinPrint:::RAW"

    [string]
    $comment

    [string]
    $location

    [string]
    $shareName

    [string]
    $availability = "Everytime"

    [bool]
    $publishINAD

    [int]
    $defaultPriority = 1

    [int]
    $deleteSpoolfiles

    [int]
    $deleteSaveSpoolFiles

    [bool] 
    $HaltFailedSpoolfiles

    [bool] 
    $PrintFirstInSpooler

    [bool] 
    $KeepSpoolfilesAfterPrinting

    [bool] 
    $RawPrinting

    [bool] 
    $LogQueueStopping

    [bool] 
    $EnableBidi

    [bool] 
    $BranchOfficeDirectPrinting

    [bool]
    $EMFDespoolingSetting
    
    [void] Create($xmlWriter) {
        
        $xmlWriter.WriteStartElement("CreateQueue")
        $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber)
        $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
        $xmlWriter.WriteElementString("Name",$this.queueName)
        $xmlWriter.WriteElementString("Driver",$this.driverName)
            $xmlWriter.WriteStartElement("Ports")
            $xmlWriter.WriteElementString("PortName",$this.portName)
            $xmlWriter.WriteEndElement()
        $xmlWriter.WriteElementString("Availability",$this.availability)
        $xmlWriter.WriteElementString("DefaultPriority",$this.defaultPriority)
        if ($this.deleteSaveSpoolFiles -and $this.deleteSaveSpoolFiles -ne "") {$xmlWriter.WriteElementString("DeleteSaveSpoolFiles",$this.deleteSaveSpoolFiles)}
        if ($this.deleteSpoolfiles -and $this.deleteSpoolfiles -ne "") {$xmlWriter.WriteElementString("DeleteSpoolfiles",$this.deleteSpoolfiles)}
        $xmlWriter.WriteElementString("PrintProcessor",$this.printProcessor)
        if ($this.comment -and $this.comment -ne "") {$xmlWriter.WriteElementString("Comment",$this.comment)}
        if ($this.location -and $this.location -ne "") {$xmlWriter.WriteElementString("Location",$this.location)}
        if ($this.shareName -and $this.shareName -ne "") {$xmlWriter.WriteElementString("Sharename",$this.shareName)}
        $xmlWriter.WriteElementString("PublishInAd",$this.publishINAD.toString().ToLower())
        if ($this.templateName -and $this.templateName -ne "") {$xmlWriter.WriteElementString("TemplateName",$this.templateName)}
            $xmlWriter.WriteStartElement("QueueAttributes")
            $xmlWriter.WriteElementString("HaltFailedSpoolfiles",$this.HaltFailedSpoolfiles.toString().ToLower())
            $xmlWriter.WriteElementString("PrintFirstInSpooler",$this.PrintFirstInSpooler.toString().ToLower())
            $xmlWriter.WriteElementString("KeepSpoolfilesAfterPrinting",$this.KeepSpoolfilesAfterPrinting.toString().ToLower())
            $xmlWriter.WriteElementString("RawPrinting",$this.RawPrinting.toString().ToLower())
            $xmlWriter.WriteElementString("LogQueueStopping",$this.LogQueueStopping.toString().ToLower())
            $xmlWriter.WriteElementString("EnableBidi",$this.EnableBidi.toString().ToLower())
            $xmlWriter.WriteElementString("BranchOfficeDirectPrinting",$this.BranchOfficeDirectPrinting.toString().ToLower())
            $xmlWriter.WriteElementString("EMFDespoolingSetting",$this.EMFDespoolingSetting.toString().ToLower())
            $xmlWriter.WriteEndElement()
        $xmlWriter.WriteEndElement() # <-- Closing CreateQueue
            
        if ($this.templateName -ne "" -and $this.publishINAD) { 
            $xmlWriter.WriteStartElement("ChangeQueue")
                $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber + 1)
                $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
                    $xmlWriter.WriteElementString("Name",$this.queueName)
                    $xmlWriter.WriteElementString("PublishInAd","false")
                $xmlWriter.WriteEndElement() # <-- Closing ChangeQueue
                $xmlWriter.WriteStartElement("ChangeQueue")
                $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber + 2)
                $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
                    $xmlWriter.WriteElementString("Name",$this.queueName)
                    $xmlWriter.WriteElementString("PublishInAd","true")
            $xmlWriter.WriteEndElement() # <-- Closing ChangeQueue
        }
    }

        <#[void] Change($xmlWriter) {
        
        $xmlWriter.WriteStartElement("CreateQueue")
        $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber)
        $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
        $xmlWriter.WriteElementString("Name",$this.queueName)
        $xmlWriter.WriteElementString("Driver",$this.driverName)
            $xmlWriter.WriteStartElement("Ports")
            $xmlWriter.WriteElementString("PortName",$this.portName)
            $xmlWriter.WriteEndElement()
        $xmlWriter.WriteElementString("Availability",$this.availability)
        $xmlWriter.WriteElementString("DefaultPriority",$this.defaultPriority)
        if ($this.deleteSaveSpoolFiles -and $this.deleteSaveSpoolFiles -ne "") {$xmlWriter.WriteElementString("DeleteSaveSpoolFiles",$this.deleteSaveSpoolFiles)}
        if ($this.deleteSpoolfiles -and $this.deleteSpoolfiles -ne "") {$xmlWriter.WriteElementString("DeleteSpoolfiles",$this.deleteSpoolfiles)}
        $xmlWriter.WriteElementString("PrintProcessor",$this.printProcessor)
        if ($this.comment -and $this.comment -ne "") {$xmlWriter.WriteElementString("Comment",$this.comment)}
        if ($this.location -and $this.location -ne "") {$xmlWriter.WriteElementString("Location",$this.location)}
        if ($this.shareName -and $this.shareName -ne "") {$xmlWriter.WriteElementString("Sharename",$this.shareName)}
        $xmlWriter.WriteElementString("PublishInAd",$this.publishINAD.toString().ToLower())
        if ($this.templateName -and $this.templateName -ne "") {$xmlWriter.WriteElementString("TemplateName",$this.templateName)}
            $xmlWriter.WriteStartElement("QueueAttributes")
            $xmlWriter.WriteElementString("HaltFailedSpoolfiles",$this.HaltFailedSpoolfiles.toString().ToLower())
            $xmlWriter.WriteElementString("PrintFirstInSpooler",$this.PrintFirstInSpooler.toString().ToLower())
            $xmlWriter.WriteElementString("KeepSpoolfilesAfterPrinting",$this.KeepSpoolfilesAfterPrinting.toString().ToLower())
            $xmlWriter.WriteElementString("RawPrinting",$this.RawPrinting.toString().ToLower())
            $xmlWriter.WriteElementString("LogQueueStopping",$this.LogQueueStopping.toString().ToLower())
            $xmlWriter.WriteElementString("EnableBidi",$this.EnableBidi.toString().ToLower())
            $xmlWriter.WriteElementString("BranchOfficeDirectPrinting",$this.BranchOfficeDirectPrinting.toString().ToLower())
            $xmlWriter.WriteElementString("EMFDespoolingSetting",$this.EMFDespoolingSetting.toString().ToLower())
            $xmlWriter.WriteEndElement()
        $xmlWriter.WriteEndElement() # <-- Closing CreateQueue
            
        if ($this.templateName -ne "" -and $this.publishINAD) { 
            $xmlWriter.WriteStartElement("ChangeQueue")
                $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber + 1)
                $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
                    $xmlWriter.WriteElementString("Name",$this.queueName)
                    $xmlWriter.WriteElementString("PublishInAd","false")
                $xmlWriter.WriteEndElement() # <-- Closing ChangeQueue
                $xmlWriter.WriteStartElement("ChangeQueue")
                $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber + 2)
                $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
                    $xmlWriter.WriteElementString("Name",$this.queueName)
                    $xmlWriter.WriteElementString("PublishInAd","true")
            $xmlWriter.WriteEndElement() # <-- Closing ChangeQueue
        }
    }#>

        [void] Delete($xmlWriter) {
        
        $xmlWriter.WriteStartElement("DeleteQueue")
        $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber)
            $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
            $xmlWriter.WriteElementString("Name",$this.DeleteQueueName)
            $xmlWriter.WriteElementString("ForceDelete",$this.ForceDelete)
        $xmlWriter.WriteEndElement()
        }
    }



