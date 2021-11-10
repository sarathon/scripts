Class PrinTaurusAutomationSAP {
    #[ValidateRange(0,100)]
    [int]
    $sequenzeNumber = 1

    [string] 
    $LongName 
    
    [string]
    $ShortName

    [string] 
    $DeviceType

    [string] 
    $Model = ""

    [string] 
    $Location = ""

    [string] 
    $Message = ""

    [string] 
    $ProsName

    [string] 
    $SpoolServer

    [string] 
    $Loms 

    [bool] 
    $Disabled = 0
    
    [bool] 
    $SkipSanity = 0
}

Class PrinTaurusAutomationSAPDevice : PrinTaurusAutomationSAP {
    [void] Create($xmlWriter) {
    $xmlWriter.WriteStartElement("CreateErpDevice")
    $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber)
    $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
    if($this.ERPSkipSanity) {$xmlWriter.WriteAttributeString("SkipSanityChecks", "true")}
        $xmlWriter.WriteElementString("Name",$this.LongName)
        $xmlWriter.WriteElementString("Destination",$this.ShortName)
        $xmlWriter.WriteElementString("Type",$this.DeviceType)
       if(!($this.Model -eq "")) { $xmlWriter.WriteElementString("Model",$this.Model) }
       if(!($this.Location -eq "")) { $xmlWriter.WriteElementString("Location",$this.Location)}
       if(!($this.Message -eq "")) { $xmlWriter.WriteElementString("Message",$this.Message)}
           $xmlWriter.WriteStartElement("Method")
                $xmlWriter.WriteStartElement("E")
                $xmlWriter.WriteElementString("ProsName",$this.ProsName)
                $xmlWriter.WriteElementString("SpoolServer",$this.SpoolServer)
                $xmlWriter.WriteElementString("Loms",$this.Loms)
                $xmlWriter.WriteEndElement()
           $xmlWriter.WriteEndElement()
    $xmlWriter.WriteEndElement()
    }

    [void] Change($xmlWriter) {
    $xmlWriter.WriteStartElement("ChangeErpDevice")
    $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber)
    $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
    if($this.ERPSkipSanity) {$xmlWriter.WriteAttributeString("SkipSanityChecks", "true")}
       $xmlWriter.WriteElementString("Destination",$this.ShortName)
       if(!($this.DeviceType -eq "")) { $xmlWriter.WriteElementString("Type",$this.DeviceType) }
       if(!($this.Model -eq "")) { $xmlWriter.WriteElementString("Model",$this.Model) }
       if(!($this.Location -eq "")) { $xmlWriter.WriteElementString("Location",$this.Location)}
       if(!($this.Message -eq "")) { $xmlWriter.WriteElementString("Message",$this.Message)}
           $xmlWriter.WriteStartElement("Method")
                $xmlWriter.WriteStartElement("E")
                $xmlWriter.WriteElementString("ProsName",$this.ProsName)
                $xmlWriter.WriteElementString("SpoolServer",$this.SpoolServer)
                $xmlWriter.WriteElementString("Loms",$this.Loms)
    $xmlWriter.WriteEndElement()
    }

    [void] Disable($xmlWriter) {
    $xmlWriter.WriteStartElement("ChangeErpDevice")
    $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber)
    $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
       $xmlWriter.WriteElementString("Destination",$this.ShortName)
       $xmlWriter.WriteElementString("Disabled",1)
    $xmlWriter.WriteEndElement()
    }

    [void] Enable($xmlWriter) {
    $xmlWriter.WriteStartElement("ChangeErpDevice")
    $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber)
    $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
       $xmlWriter.WriteElementString("Destination",$this.ShortName)
       $xmlWriter.WriteElementString("Disabled",0)
    $xmlWriter.WriteEndElement()
    }

    [void] Delete($xmlWriter) {
        $xmlWriter.WriteStartElement("DeleteErpDevice")
    $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber)
    $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
        $xmlWriter.WriteElementString("Destination",$this.ShortName)
    $xmlWriter.WriteEndElement()
    }
}