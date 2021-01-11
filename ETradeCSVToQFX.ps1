[CmdletBinding(RemotingCapability='None')]
param(
    [Parameter(Position=0)]
    [ValidateNotNullOrEmpty()]
    [string]
    ${CSVFilePath},

    [Parameter(Position=1)]
    [ValidateNotNullOrEmpty()]
    [string]
    ${TemplateOFXFilePath},

    [Parameter(Position=2)]
    [ValidateNotNullOrEmpty()]
    [string]
    ${OFXFilePath})
begin
{
    [hashtable]$TranslateTransType = @{
            'Bill Pay' = 'PAYMENT'
            'Check' = 'PAYMENT' # Reusing PAYMENT
            'Deposit' = 'DEP'
            'Direct Debit' = 'DIRECTDEBIT'
            'Direct Deposit' = 'DIRECTDEP'
            'Interest' = 'INT'
            'Transfer In' = 'XFER'
            'Transfer Out' = 'XFER'
        }

    function ConvertTo-ETradeQFX {
        [CmdletBinding(RemotingCapability='None')]
        param(
            [Parameter(Position=0)]
            [ValidateNotNullOrEmpty()]
            [string]
            ${CSVFilePath},
        
            [Parameter(Position=1)]
            [ValidateNotNullOrEmpty()]
            [string]
            ${TemplateOFXFilePath})
        begin
        {
            [array]$TransactionXML = @()
            [string]$DtStartXML = '99999999999999'
            [string]$DtEndXML = '00000000000000'
            [string]$BalAmtXML = '0.00'
        }
        process
        {
            [int32]$FitID = 2678000

            $TransactionXML = `
                Get-Content -Path $CSVFilePath `
                    | Select-Object -Skip 4 `
                    | ConvertFrom-Csv `
                    | Select-Object -Skip 1 `
                    | ForEach-Object -Process {
                            [string]$TransactionDateXML = "20000101"
                            [string]$FitIDXML = "20000101"
                            if ($_.TransactionDate -match '(\d+)/(\d+)/(\d\d)')
                            {
                                [string]$Month = $matches[1]
                                if ($Month.Length -lt 2)
                                {
                                    $Month = '0' + $Month
                                }
                                [string]$Day = $matches[2]
                                if ($Day.Length -lt 2)
                                {
                                    $Day = '0' + $Day
                                }
                                $TransactionDateXML = "20$( $matches[3] )$( $Month )$( $Day )"
                                $FitIDXML = "20$( $matches[3] )$( $Day )$( $Month )"
                            }
                            [string]$NameXML = $_.Description
                            if ($NameXML.Length -gt 32)
                            {
                                $NameXML = $NameXML.Substring(0,32)
                            }

                            if ($DtStartXML.CompareTo( $TransactionDateXML ) -gt 0)
                            {
                                $DtStartXML = $TransactionDateXML
                            }
                            if ($DtEndXML.CompareTo( $TransactionDateXML ) -lt 0)
                            {
                                $DtEndXML = $TransactionDateXML
                                $BalAmtXML = $_.Balance -replace '(\-*\d+\.\d\d)\d*', '$1' # Remove unnecessary decimal digits
                            }
                            
                            '<STMTTRN>'
                            "<TRNTYPE>$( $TranslateTransType[$_.TransactionType] )"
                            "<DTPOSTED>$( $TransactionDateXML )160000" # Convert MM/DD/YY to 'YYYYMMDD160000'
                            "<TRNAMT>$( $_.Amount -replace '(\-*\d+\.\d\d)\d*', '$1' )" # Remove unnecessary decimal digits
                            "<FITID>$( $FitIDXML )$( $FitID )" # Convert MM/DD/YY to 'YYYYDDMMFitID'
                            "<NAME>$( $NameXML )"
                            "<MEMO>$( $_.TransactionType )-$( $NameXML )"
                            '</STMTTRN>'
                
                            $FitID -= 1000                
                        }
        }
        end
        {
            $InStream = New-Object -TypeName System.IO.StreamReader -ArgumentList (Resolve-Path $TemplateOFXFilePath)
            [string]$InLine = ''
            while ($null -ne ($InLine = $InStream.ReadLine()) -and ($InLine -notmatch '<DTSTART>\d+'))
            {
                $InLine.ToString()
            }
            "<DTSTART>$( $DtStartXML )040000"
            while ($null -ne ($InLine = $InStream.ReadLine()) -and ($InLine -notmatch '<DTEND>\d+'))
            {
                $InLine.ToString()
            }
            "<DTEND>$( $DtEndXML )035959"
            while ($null -ne ($InLine = $InStream.ReadLine()) -and ($InLine -notmatch '<STMTTRN>'))
            {
                $InLine.ToString()
            }

            while ($null -ne ($InLine = $InStream.ReadLine()) -and ($InLine -notmatch '</BANKTRANLIST>'))
            { }
            $TransactionXML

            $InLine.ToString()
            while ($null -ne ($InLine = $InStream.ReadLine()) -and ($InLine -notmatch '<BALAMT>\d+.\d+'))
            {
                $InLine.ToString()
            }
            "<BALAMT>$BalAmtXML"
            while ($null -ne ($InLine = $InStream.ReadLine()) -and ($InLine -notmatch '<DTASOF>\d+'))
            {
                $InLine.ToString()
            }
            "<DTASOF>$( $DtEndXML )040000"
            while (-not $InStream.EndOfStream -and $null -ne ($InLine = $InStream.ReadLine()))
            {
                $InLine.ToString()
            }
        }
    }
}
process
{
}
end
{
    ConvertTo-ETradeQFX $CSVFilePath $TemplateOFXFilePath | Out-File -Path $OFXFilePath -Encoding ascii
}
