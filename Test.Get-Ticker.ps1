# Converted from bash to Powershell reference: https://github.com/alexanderepstein/Bash-Snippets/blob/master/stocks/stocks

function Get-Ticker([string]$CompanyName, [string]$Ticker) {
    $Query = ""
    if ($CompanyName) {
        $Query = $CompanyName.Replace(' ', '+')
    }
    elseif ($Ticker) {
        $Query = $Ticker.Replace(' ', '+')
    }
     

    $response = Invoke-WebRequest "http://d.yimg.com/autoc.finance.yahoo.com/autoc?query=$Query&region=1&lang=en%22"
    $data = $response.Content | ConvertFrom-Json
    
    try
    {
        $Symbol = $data.ResultSet.Result[0].symbol    

        $data = Invoke-WebRequest "https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=$($Symbol)&apikey=KPCCCRJVMOGN9L6T"
        $stockInfo = $data.Content | ConvertFrom-Json

        Write-Verbose $stockInfo.'Global Quote'
        try
        {
            $symbol = $stockInfo.'Global Quote'.'01. symbol'             #: FB
            $open = $stockInfo.'Global Quote'.'02. open'               #: 161.5000
            $high = $stockInfo.'Global Quote'.'03. high'               #: 163.7800
            $low  = $stockInfo.'Global Quote'.'04. low'                #: 161.2400
            $latestPrice = $stockInfo.'Global Quote'.'05. price'              #: 163.0800
            $volume =  $stockInfo.'Global Quote'.'06. volume'             #: 7629474
            $lastUpdated = $stockInfo.'Global Quote'.'07. latest trading day' #: 2019-03-20
            $close = $stockInfo.'Global Quote'.'08. previous close'     #: 161.5700
            $priceChange = $stockInfo.'Global Quote'.'09. change'             #: 1.5100
            $priceChangePercentage = $stockInfo.'Global Quote'.'10. change percent'     #: 0.9346%

            $output = New-Object -TypeName PSObject
            $output | Add-Member -MemberType NoteProperty -Name ExchangeName -Value $symbol
            $output | Add-Member -MemberType NoteProperty -Name LatestPrice -Value $latestPrice
            $output | Add-Member -MemberType NoteProperty -Name Open -Value $open
            $output | Add-Member -MemberType NoteProperty -Name High -Value $high
            $output | Add-Member -MemberType NoteProperty -Name Low -Value $low
            $output | Add-Member -MemberType NoteProperty -Name Close -Value $close
            $output | Add-Member -MemberType NoteProperty -Name priceChange -Value $priceChange
            $output | Add-Member -MemberType NoteProperty -Name PriceChangePercentage -Value $priceChangePercentage
            $output | Add-Member -MemberType NoteProperty -Name Volume -Value $volume
            $output | Add-Member -MemberType NoteProperty -Name lastUpdated -Value $lastUpdated

        }
        catch [DivideByZeroException]
        {
            Write-Host "Divide by zero exception"
        }

        finally
        {
            $output
        }

    }

    catch [System.Net.WebException],[System.Exception]
    {
        Write-Host "Other exception"
    }
    finally
    {
        #Write-Host "cleaning up ..."
               
    }    
}
cls

$tickerList = "DOW,FB,AMZN"

$output1 = @()
"DOW,FB,AMZN".Split(",") | ForEach {
    write-host $_ 
    $output1 += Get-Ticker -Ticker dow $_
 }

$output1 | Out-GridView
