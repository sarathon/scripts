Class PrinTaurusAutomationPort {
    [ValidateNotNullOrEmpty()]
    [string]
    $hostName
    [ValidateNotNullOrEmpty()]
    [string]
    $portName
    [int]
    $port = 9100
    [bool]
    $SnmpEnable = 0
    
    [string]
    $SnmpCommunity = "public"

    [int]
    $SnmpIndex = 1

    #[ValidateRange(0,100)]
    [int]
    $sequenzeNumber = 1
}

Class PrinTaurusAutomationPortTCPIP : PrinTaurusAutomationPort {
    [void] Create($xmlWriter) {
 
        $xmlWriter.WriteStartElement("CreateTcpIpPort")
        $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber)
        $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
            $xmlWriter.WriteElementString("Name",$this.portName)
            $xmlWriter.WriteElementString("Hostname",$this.hostName)
            $xmlWriter.WriteElementString("Port",$this.port)
            if ($this.SnmpEnable) {
                    $xmlWriter.WriteStartElement("Snmp")
                        $xmlWriter.WriteElementString("SnmpEnable",$this.SnmpEnable.toString().ToLower())
                        $xmlWriter.WriteElementString("SnmpIndex",$this.SnmpIndex)
                        $xmlWriter.WriteElementString("SnmpCommunity",$this.SnmpCommunity)
                    $xmlWriter.WriteEndElement()
                } 
             else
            {
                $xmlWriter.WriteStartElement("Snmp")
                    $xmlWriter.WriteElementString("SnmpEnable",$this.SnmpEnable.toString().ToLower())
                    $xmlWriter.WriteElementString("SnmpIndex","")
                    $xmlWriter.WriteElementString("SnmpCommunity","")
                $xmlWriter.WriteEndElement()
            }
        $xmlWriter.WriteEndElement()
    }
    
    <#
    [void] Change($xmlWriter) {

        $xmlWriter.WriteStartElement("CreateTcpIpPort")
        $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber)
        $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
            $xmlWriter.WriteElementString("Name",$this.portName)
            $xmlWriter.WriteElementString("Hostname",$this.hostName)
            $xmlWriter.WriteElementString("Port",$this.port)
      if ($this.SnmpEnable) {
            $xmlWriter.WriteStartElement("Snmp")
                $xmlWriter.WriteElementString("SnmpEnable",$this.SnmpEnable.toString().ToLower())
                $xmlWriter.WriteElementString("SnmpIndex",$this.snmpIndex)
                $xmlWriter.WriteElementString("SnmpCommunity",$this.SnmpCommunity)
            $xmlWriter.WriteEndElement()
        } 
        else
        {
            $xmlWriter.WriteStartElement("Snmp")
                $xmlWriter.WriteElementString("SnmpEnable",$this.SnmpEnable.toString().ToLower())
                $xmlWriter.WriteElementString("SnmpIndex","")
                $xmlWriter.WriteElementString("SnmpCommunity","")
            $xmlWriter.WriteEndElement()
        }
        $xmlWriter.WriteEndElement()
    
        }
        #>

    [void] Delete($xmlWriter) {
            $xmlWriter.WriteStartElement("DeletePrinTaurusPort")
            $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber)
            $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
                $xmlWriter.WriteElementString("Name",$this.PortName)
            $xmlWriter.WriteEndElement()
        }
    }

Class PrinTaurusAutomationPortPMC : PrinTaurusAutomationPort {
        [ValidateNotNullOrEmpty()]
        [string]
        $PMC_CommandLine
        [ValidateNotNullOrEmpty()]
        [string]
        $PMC_OnErrorCommand
        [string]
        $PMC_PostPrintCommand
        [string]
        $PMC_PrePrintCommand

        [bool]
        $PMC_IppPort = 0
        [bool]
        $PMC_SSL = 0

        [bool]
        $PMC_FollowPrintPort = 0

        [string]
        $PMC_DestinationQueue

        [bool]
        $PMC_PJL = 0

        [bool]
        $PMC_SaveSpoolFile = 0

        [void] Create($xmlWriter) {
            $pmcPortName = "PMC:" + $this.portName
            $xmlWriter.WriteStartElement("CreatePrinTaurusPort")
            $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber)
            $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
            $xmlWriter.WriteElementString("Name",$pmcPortName)
            $xmlWriter.WriteStartElement("DestinationInfo")             
                if ($this.PMC_CommandLine.Length -ne 0) {   

                $xmlWriter.WriteStartElement("ProgramCall")
                    $xmlWriter.WriteElementString("CommandLine", $this.PMC_CommandLine)
                    if ($this.PMC_PostPrintCommand.Length -gt 0 -or $this.PMC_OnErrorCommand.Length -gt 0) { 
                            $xmlWriter.WriteStartElement("OtherCommands>")
                            if ( $this.PMC_PostPrintCommand.Length -ne 0) { $xmlWriter.WriteElementString("PostPrintCommand",$this.PMC_PostPrintCommand) }
                            if ( $this.PMC_OnErrorCommand.Length -ne 0) {  $xmlWriter.WriteElementString("OnErrorCommand",$this.PMC_OnErrorCommand) }
                            $xmlWriter.WriteEndElement()
                            }
                    $xmlWriter.WriteElementString("SaveSpoolFile", $this.PMC_SaveSpoolFile.ToString().ToLower())
                $xmlWriter.WriteEndElement() 

                } elseif ($this.PMC_IppPort) {
                    $xmlWriter.WriteStartElement("NetworkPort")
                        $xmlWriter.WriteElementString("Hostname",$this.hostName) 
                        $xmlWriter.WriteElementString("SSL",$this.PMC_SSL.ssl.toString().ToLower())

                    $xmlWriter.WriteEndElement()

                } elseif ($this.PMC_FollowPrintPort) {
                    $xmlWriter.WriteStartElement("FollowPrintQueue")
                    $xmlWriter.WriteEndElement()

                } elseif ($this.PMC_DestinationQueue.Length -ne 0) {
                    $xmlWriter.WriteStartElement("QueueRedirection")
                    $xmlWriter.WriteElementString("DestinationQueue",$this.PMC_DestinationQueue)
                    if ($this.PMC_PrePrintCommand -ne "" -or $this.PMC_PostPrintCommand -ne "" -or $this.PMC_OnErrorCommand -ne "") { 
                    $xmlWriter.WriteStartElement("Commands")
                        if ($this.PMC_PrePrintCommand.Length -ne 0) { $xmlWriter.WriteElementString("PrePrintCommand",$this.PMC_PrePrintCommand)}
                           
                    $xmlWriter.WriteEndElement()
                }
                $xmlWriter.WriteElementString("SaveSpoolFile", $this.PMC_SaveSpoolFile.ToString().ToLower()) 
                    $xmlWriter.WriteEndElement()
            
                } else {

            $xmlWriter.WriteStartElement("NetworkPort")
                $xmlWriter.WriteElementString("Hostname",$this.hostName)
                $xmlWriter.WriteElementString("Port",$this.port)
                $xmlWriter.WriteElementString("PJL",$this.PMC_PJL.toString().ToLower())
                if ($this.PMC_PrePrintCommand.Length -ne 0 -or ($this.PMC_PostPrintCommand.Length -ne 0) -or ($this.PMC_OnErrorCommand.Length -ne 0)) { 
                    $xmlWriter.WriteStartElement("Commands")
                        if ($this.PMC_PrePrintCommand.Length -ne 0) { $xmlWriter.WriteElementString("PrePrintCommand",$this.PMC_PrePrintCommand)}
                           if ($this.PMC_PostPrintCommand.Length -gt 0 -or $this.PMC_OnErrorCommand.Length -gt 0) { 
                            $xmlWriter.WriteStartElement("OtherCommands>")
                            if ( $this.PMC_PostPrintCommand.Length -ne 0) { $xmlWriter.WriteElementString("PostPrintCommand",$this.PMC_PostPrintCommand) }
                            if ( $this.PMC_OnErrorCommand.Length -ne 0) {  $xmlWriter.WriteElementString("OnErrorCommand",$this.PMC_OnErrorCommand) }
                            $xmlWriter.WriteEndElement()
                            }
                    $xmlWriter.WriteEndElement()
                }
                $xmlWriter.WriteElementString("SaveSpoolFile", $this.PMC_SaveSpoolFile.ToString().ToLower())

            $xmlWriter.WriteEndElement()

            }
            
            $xmlWriter.WriteEndElement()
         $xmlWriter.WriteEndElement()
        }
        
        <#
        [void] Change($xmlWriter) {
    
            $xmlWriter.WriteStartElement("CreateTcpIpPort")
            $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber)
            $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
                $xmlWriter.WriteElementString("Name",$this.portName)
                $xmlWriter.WriteElementString("Hostname",$this.hostName)
                $xmlWriter.WriteElementString("Port",$this.port)
          if ($this.SnmpEnable) {
                $xmlWriter.WriteStartElement("Snmp")
                    $xmlWriter.WriteElementString("SnmpEnable",$this.SnmpEnable.toString().ToLower())
                    $xmlWriter.WriteElementString("SnmpIndex",$this.snmpIndex)
                    $xmlWriter.WriteElementString("SnmpCommunity",$this.SnmpCommunity)
                $xmlWriter.WriteEndElement()
            } 
            else
            {
                $xmlWriter.WriteStartElement("Snmp")
                    $xmlWriter.WriteElementString("SnmpEnable",$this.SnmpEnable.toString().ToLower())
                    $xmlWriter.WriteElementString("SnmpIndex","")
                    $xmlWriter.WriteElementString("SnmpCommunity","")
                $xmlWriter.WriteEndElement()
            }
            $xmlWriter.WriteEndElement()
        
            }
            #>
    
        [void] Delete($xmlWriter) {
            $pmcPortName = "PMC:" + $this.portName
                $xmlWriter.WriteStartElement("DeletePrinTaurusPort")
                $xmlWriter.WriteAttributeString("Sequence", $this.sequenzeNumber)
                $xmlWriter.WriteAttributeString("ReferenceID", (Get-Random -Minimum 1000000 -Maximum 9999999))
                    $xmlWriter.WriteElementString("Name",$pmcPortName)
                $xmlWriter.WriteEndElement()
            }
        }
    
    