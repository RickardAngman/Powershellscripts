function SEB_date_to_OFX_date {
    param (
        $datum
    )
    $shortdate=$datum -replace "-" ,""
    return $shortdate
}

function GetFirstTransactionDate {
    param (
        $FileName
    )
    [datetime]$toreturn = [datetime]::Today
    [datetime]$datetimeDate= [datetime]::Today
    Import-Csv -Path $FileName -Header Datum,Text,Belopp -Delimiter ";" | ForEach-Object  { 
        try {
            [datetime]$datetimeDate= [datetime]::ParseExact($_.Datum,'yyyy-MM-dd',$null)
        }
        catch {
            
        }
        if ($toreturn -gt $datetimeDate) {
            Write-Debug ($toreturn.ToString() + " is greater than " + $datetimeDate)
            $toreturn=$datetimeDate
            Write-Debug  ("Setting new toreturn to " + $toreturn)
        }
    }
    return SEB_date_to_OFX_date $toreturn.ToShortDateString()
}

function GetLastTransactionDate {
    param (
        $FileName
    )
    [datetime]$toreturn = "2000-01-01"
    [datetime]$datetimeDate= "2000-01-01"
    Import-Csv -Path $FileName -Header Datum,Text,Belopp -Delimiter ";" | ForEach-Object  { 
        try {
            [datetime]$datetimeDate= [datetime]::ParseExact($_.Datum,'yyyy-MM-dd',$null)
        }
        catch {
            
        }   
        if ($toreturn -lt $datetimeDate) {
            Write-Debug ($toreturn.ToString() + " is less than " + $datetimeDate)
            $toreturn=$datetimeDate
            Write-Debug ("Setting new toreturn to " + $toreturn)
        }
    }
    return SEB_date_to_OFX_date $toreturn.ToShortDateString()    
}
$dir = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$csvFileName=$dir+"\"+"kontoutdrag.csv"
$ofxFileName=$dir+"\"+"ofxForImport.ofx"
$dateForFirstTransaction=GetFirstTransactionDate($csvFileName)
$dateForLastTransaction=GetLastTransactionDate($csvFileName)
# Set The Formatting
$xmlsettings = New-Object System.Xml.XmlWriterSettings
$xmlsettings.Indent = $true
$xmlsettings.IndentChars = "  "
# Set the File Name Create The Document
$XmlWriter = [System.XML.XmlWriter]::Create($ofxFileName, $xmlsettings)
$xmlWriter.WriteStartDocument()
$xmlwriter.WriteRaw('<?OFX OFXHEADER="200" VERSION="202" SECURITY="NONE" OLDFILEUID="NONE" NEWFILEUID="NONE"?>')
$xmlWriter.WriteStartElement("OFX")
    $xmlWriter.WriteStartElement("SIGNONMSGSRSV1") 
        $XmlWriter.WriteStartElement("SONRS")
            $XmlWriter.WriteStartElement("STATUS")
                $XmlWriter.WriteElementString("CODE",0)
                $xmlwriter.WriteElementString("SEVERITY","INFO")
                $xmlwriter.WriteEndElement() # <-- stänger STATUS
            
            #$XmlWriter.WriteElementString("DTSERVER",$dateForFirstTransaction) # Hitta datum för första transaktion och skapa datumet - läge för en funktion
            $XmlWriter.WriteElementString("DTSERVER",$dateForLastTransaction)
            $XmlWriter.WriteElementString("LANGUAGE","ENG")
            $xmlwriter.WriteEndElement() # <-- stänger SONRS
        $xmlwriter.WriteEndElement() # <-- stänger SIGNONMSGSRSV1
    $XmlWriter.WriteStartElement("BANKMSGSRSV1")
        $xmlwriter.WriteStartElement("STMTTRNRS")
            $XmlWriter.WriteElementString("TRNUID",0)
            $XmlWriter.WriteStartElement("STATUS")
                $XmlWriter.WriteElementString("CODE",0)
                $xmlwriter.WriteElementString("SEVERITY","INFO")
                $xmlwriter.WriteEndElement() # <-- stänger STATUS
            $XmlWriter.WriteStartElement("STMTRS")
                $XmlWriter.WriteElementString("CURDEF","SEK")
                $XmlWriter.WriteStartElement("BANKACCTFROM")
                    $XmlWriter.WriteElementString("BANKID","iCOFX")
                    $XmlWriter.WriteElementString("ACCTID","1234567890")
                    $XmlWriter.WriteElementString("ACCTTYPE","SAVINGS")
                    $xmlwriter.WriteEndElement() # <-- stänger BANKACCTFROM
                $XmlWriter.WriteStartElement("BANKTRANLIST")
                    $XmlWriter.WriteElementString("DTSTART",$dateForFirstTransaction) # Hitta datum
                    
                    $XmlWriter.WriteElementString("DTEND",$dateForLastTransaction) # Hitta datum
################ Loopa igenom alla transaktioner
                    Import-Csv -Path $csvFileName -Header Datum,Text,Belopp -Delimiter ";" | ForEach-Object  { 
                        $decBelopp=0
                        try {
                            [decimal]$decBelopp= [System.Convert]::ToDecimal($_.Belopp,[cultureinfo]::GetCultureInfo('sv-SE'))    
                            }
                        catch {
                        }
                        if ($decBelopp -ne 0) {
                            if ($decBelopp -lt 0) {
                                $TRNTYPE="DEBIT"
                            } else {
                                $TRNTYPE="CREDIT"
                            }
                            $DTPOSTED=SEB_date_to_OFX_date -datum $_.Datum
                            $FITID=New-Guid 
                            $XmlWriter.WriteStartElement("STMTTRN")
                            $XmlWriter.WriteElementString("TRNTYPE",$TRNTYPE) # Beroende på minus eller positivt
                            $XmlWriter.WriteElementString("DTPOSTED",$DTPOSTED) # Hitta datum
                            $XmlWriter.WriteElementString("TRNAMT",$decBelopp.ToString()) # Beloppet, debit har minus och credit har plus
                            $XmlWriter.WriteElementString("FITID",$FITID) # Unik identfierare, kanske använda New-Guid?
                            $XmlWriter.WriteElementString("NAME",$_.Text.trim()) # Från texten
                            $XmlWriter.WriteElementString("MEMO",$_.Text.trim()) # Från texten
                            $xmlwriter.WriteEndElement() # <-- stänger STMTTRN
                            }
                        }
################ Stänger loopen av alla transaktioner
                    $xmlwriter.WriteEndElement() # <-- stänger BANKTRANLIST
                    $xmlwriter.WriteStartElement("LEDGERBAL")
                        $xmlwriter.WriteElementString("BALAMT",",00")
                        $xmlwriter.WriteElementString("DTASOF",$dateForLastTransaction) # Något slags datum
                        $xmlwriter.WriteEndElement() # <-- stänger LEDGERBAL
            $xmlwriter.WriteEndElement() # <-- stänger STMTTRNRS
        $xmlwriter.WriteEndElement() # <-- stänger BANKMSGSRSV1
$xmlWriter.WriteEndElement() # <-- End <Root> 
# End, Finalize and close the XML Document
$xmlWriter.WriteEndDocument()
$xmlWriter.Flush()
$xmlWriter.Close()