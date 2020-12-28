function SEB_date_to_OFX_date {
    # SEB uses yyyy-MM-dd but OFX likes yyyyMMdd
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
    Import-Csv -Path $FileName -Header Datum,Text,Belopp -Delimiter "," | ForEach-Object  { 
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
    Import-Csv -Path $FileName -Header Datum,Text,Belopp -Delimiter "," | ForEach-Object  { 
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
function AddTransactionDatesToFilename {
    # Add transactions dates to the filename given to this function.
    param (
        $FileName
    )
    Write-Host "BaseName"(Get-Item $FileName).BaseName
    Write-Host "Extension"(Get-Item $FileName).Extension
    $NewFileName = (Get-Item $FileName).BaseName + "-" + $dateForFirstTransaction + "-" + $dateForLastTransaction + (Get-Item $FileName).Extension 
    Write-Host "FileName"$FileName
    Write-Host "NewFileName:"$NewFileName
    Rename-Item -Path $FileName -NewName $NewFileName
}

$dir = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
$excelFileName=$dir+"\"+"kontoutdrag.xlsx" # The supposed name that SEB exports
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
                    $XmlWriter.WriteElementString("ACCTID","53543329987") # 2234567890 = Sparkontot | 1234567890 = Lönekontot | 53540293744 = Gemensamt kortkonto | 53543329987 = Gemensamt sparkonto
                    $XmlWriter.WriteElementString("ACCTTYPE","SAVINGS")
                    $xmlwriter.WriteEndElement() # <-- stänger BANKACCTFROM
                $XmlWriter.WriteStartElement("BANKTRANLIST")
                    $XmlWriter.WriteElementString("DTSTART",$dateForFirstTransaction) # Hitta datum
                    $XmlWriter.WriteElementString("DTEND",$dateForLastTransaction) # Hitta datum
################ Loopa igenom alla transaktioner
                    Import-Csv -Path $csvFileName -Header Datum,Text,Belopp -Delimiter "," | ForEach-Object {
                        [decimal]$decBelopp=0
                        [decimal]$decBeloppsvSE=0
                        [decimal]$decBeloppenUs=0
                        try {
                            [decimal]$decBeloppsvSE= [System.Convert]::ToDecimal($_.Belopp,[cultureinfo]::GetCultureInfo('sv-SE')) 
                            }
                        catch {
                        }
                        try {
                            [decimal]$decBeloppenUs= [System.Convert]::ToDecimal($_.Belopp,[cultureinfo]::GetCultureInfo('en-US'))
                            }
                        catch {
                        }
                        if ([Math]::Abs($decBeloppsvSE) -gt [Math]::Abs($decBeloppenUs)) {
                                $decBelopp = $decBeloppsvSE
                            } else {
                                $decBelopp = $decBeloppenUs
                        }
                        Write-Host $decBelopp.ToDecimal([cultureinfo]::GetCultureInfo('sv-SE'))
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
                            $XmlWriter.WriteElementString("TRNAMT",$decBelopp.ToString()) # Beloppet, debit har minus och credit har plus. Need to use ToString to get the right decimal sign for matching MsMoney.
                            $XmlWriter.WriteElementString("FITID",$FITID) # Unik identfierare, kanske använda New-Guid?
                            [string]$LeftText = $_.Text
                            $pos = $LeftText.IndexOf("/")
                            if ($pos -ne -1) {
                                $LeftText = $LeftText.Substring(0, $pos)
                            } 
                            $XmlWriter.WriteElementString("NAME",$LeftText.trim()) # Från texten
                            $XmlWriter.WriteElementString("MEMO",$LeftText.trim()) # Från texten
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

# Archive the input files to make the directory ready for next export from SEB
Add-Type -AssemblyName PresentationFramework
$continue = [System.Windows.MessageBox]::Show('Please check ofx-file before archive of input files. Do you want to archive', 'Confirmation', 'YesNo');
if ($continue -eq 'Yes') {
    # The user said "yes"
    AddTransactionDatesToFilename $excelFileName
    AddTransactionDatesToFilename $csvFileName
    Write-Output "Renamed the files"
} else {
    # The user said "no"
    Write-Output "Terminating process..."
    Start-Sleep 1
}
