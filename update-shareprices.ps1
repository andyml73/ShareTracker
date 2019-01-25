<#  
    Use powershell to update shares spreadsheet with live values.
    Values are pulled from Yahoo finance
#>


Function Write-Log([string]$logentry){
    [string]$now = get-date
    Add-Content $logfile -Value ($now+":"+$logentry)
}


$logfile = "C:\Users\S232480\Documents\_Scripts\PowerShell\update-shareprices.log"
Set-Content $logfile " "
$file = "C:\Users\S232480\OneDrive - Atos\MyDocs\shares.xlsx"
$sheetName = "FT-export"

$browser = New-Object System.Net.WebClient
$browser.Proxy.Credentials =[System.Net.CredentialCache]::DefaultNetworkCredentials


# Get exchange rate (required for Atos shares)
$exprice = $exdate = $extime = "NA"

$expage = $browser.DownloadString("https://finance.yahoo.com/quote/GBPEUR%3DX?p=GBPEUR%3DX")
$search = $expage | select-string -Pattern "EUR.,.regularMarketPrice.:{.raw.:([\d\.\,]+),"
if ($search.matches.Success){$exprice = $search.matches.groups[1].Value}

# Get date and time also
$search2 = $expage | select-string -Pattern "U.S. Markets.,.time.:.(\d\d\d\d\-\d\d\-\d\d)T(\d\d\:\d\d)\:"
if ($search2.matches.Success){
    $exdate = $search2.matches.groups[1].Value
    $extime = $search2.matches.groups[2].Value
}

$now = get-date
$datetime = $now.ToString("dd/MM/yyyy hh:mm")
#$datetime = $exdate + " " + $extime

#Create an instance of Excel.Application and Open Excel file
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
#$objExcel.Visible=$false # See updates happening?
start-sleep -Milliseconds 2000 # Issue with one drive and versioning?


#Count max row
$rowMax = ($sheet.UsedRange.Rows).count # Includes cells that are formatted (or have formatting removed?!)


#Define Excel column positions
$cName   = 1
$cSymbol = 15
$cPrice  = 16
$cPence  = 17
$cQty    = 20

# Put search patterns into a dictionary, different search patterns are required for shares, funds.
$patterns = @{
    1=".currentPrice.:{.raw.:([\d\.]+),";
    2="bid.:{.raw.:(.*?),.fmt";
    3="GBp.,.regularMarketPrice.:{.raw.:([\d\.]+),"
}
$pattype = 3


# a loop to store data in Excel
for ($i=2; $i -le $rowMax-1; $i++){
    $symbol = $sheet.Cells.Item($i,$cSymbol).text
    $qty = "0"
    $pence = "0"

    if ($symbol){ 
        $name = $sheet.Cells.Item($i,$cName).text
        write-log "`nProcessing $name"

        start-sleep -Milliseconds 500
        $URL = "https://finance.yahoo.com/quote/"+$symbol+"/?p="+$symbol
        $page = $browser.DownloadString($URL)

        #store downloaded page, useful for troubleshooting search difficulties
        $download = "C:\Users\S232480\Documents\_Scripts\PowerShell\downloads\" + $symbol +".htm"
        sc $download $page

        #Use optimised search string
        if ($symbol -match "(ATO|LLOY)"){
            $pattype = 1
            $searchstr = $patterns[$pattype]
            write-log "SEARCH pattern 1: $searchstr"
        }
        elseif ($symbol -match "(HMWO|IGUS|VERX|VGOV)"){
            $pattype = 2
            $searchstr = $patterns[$pattype]
            write-log "SEARCH pattern 2: $searchstr"
        }
        else {
            $pattype = 3
            $searchstr = $patterns[$pattype]
            write-log "SEARCH pattern 3: $searchstr"
            $page = $page.split("£")[1]   # Needed to locate valid data amongst the noise
        }

        # Perform search
        $search = $page | select-string -Pattern $patterns[$pattype]

        $price = $price_num = "0"
        if ($search.matches.Success){$price = $search.matches.groups[1].Value}
        else {
            write-log "Search failed"
        }

        # Only update spreadsheet if we have valid data
        $success = 0
        [double]$prevprice = $sheet.cells.item($i,$cPrice).text
        if ($price -match '^[\d\.]+$' -and $price -gt 0){
            $price_num = $price/1      # convert str to number
            write-log "+ve number returned"
	        #Also checking for a n% swing due to problems retrieving the price
            $lowerp = $prevprice *.9
            $upperp = $prevprice * 1.1
            write-log "$lowerp < $price_num > $upperp ?"
            if ( ($price_num -gt $lowerp) -and ($price_num -lt $upperp) ){
                write-log "with-in n%, updating sheet"
                $success = 1
                $sheet.Cells.Item($i,$cPrice) = $price_num
            }
        }

        $qty = $sheet.Cells.Item($i,$cQty).text        # for share-prices.csv
        $pence = $sheet.Cells.Item($i,$cPence).text    # for share-prices.csv
        $price_prnt = [math]::round($price_num,2)      # two decimal places for writing to screen
        if ($success -eq 0){
            write-host "$pattype $price_prnt`t$name ..prev: $prevprice ** NOT updated **" -ForegroundColor Red
            write-log "$price_num`t$name ..prev: $prevprice ** NOT updated **"
            }
        else {
            if ($prevprice -gt $price_num){  #some colour formatting
                write-host "$price_prnt" -ForegroundColor Yellow -NoNewline; write-host "`t$name ..prev: $prevprice"}
            elseif ($prevprice -lt $price_num){
                write-host "$price_prnt" -ForegroundColor Green -NoNewline; write-host "`t$name ..prev: $prevprice"}
            else {
                write-host "$price_prnt`t$name ..prev: $prevprice"}

            write-log "$price_num`t$name ..prev: $prevprice"
        }

        # Also adding data to a separate CSV file to allow for historical analysis
        $addstring = $datetime+","+$symbol+","+$pence+","+$qty
        $newstring = $addstring.replace("_","") # remove any underscores
        Add-Content -path "C:\Users\S232480\Documents\My Docs\share-prices.csv" $newstring
    }

    if ($i -gt 35){break} # Added because rowMax was 85 (due to cells that were formatted?)
}


#store the exchange rate details in spreadsheet
$sheet.Cells.Item(28,2) = $exprice
$sheet.Cells.Item(28,3) = $exdate
$sheet.Cells.Item(28,4) = $extime

#show new total (r25,c18)
$newtotal = $sheet.cells.item(25,18).text
write-host -F Black ("`nTotal: "+$newtotal)

#close excel file?
$workbook.save()
#$objExcel.quit()

<#
Todo 

Yahoo site returns XML that contains way too much data, i.e. other companies and other currencies,
Consequently the search string sometimes returns the wrong price!?

** Update, included the page.split operation to fix. 

#>