#Sourcing Data from Alpha Vantage Api

 

#"https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=" + $ticker + "&apikey="

 

 

##parse json function - chop up Json query and extract most recent price and stock symbol

 

function Parse-Json($Stock) {

 

$thing = $stock  | convertfrom-json

 

foreach ($obj in $thing.psobject.properties)

{

    $parm = @{Time = $obj.Name}

 

    $obj.value.psobject.properties | ForEach-Object {

        $parm[$PSItem.name] = $PSItem.value

    }

 

    [pscustomobject]$parm

 

}

 

}

 

#Get-stockprice - pass stock symbol to receive price

function Get-StockPrice
{
  param
  (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [Object[]]
    $Ticker 
  )
  
  process
  {
    $Ticker | ForEach-Object {

      $element = $_
      $url = "https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=" + $element + "&apikey="

      $stock = Invoke-RestMethod -Uri $url | ConvertTo-Json

    Parse-Json -Stock $stock | select "01. symbol","05. price"
    }
  }
}

 
function Get-StockPrice
{
  param
  (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [string[]]
    $ticker 
  )
  
  process
  {
    ForEach-Object {
    
    $url = "https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=" + $ticker + "&apikey="

    $stock = Invoke-RestMethod -Uri $url | ConvertTo-Json

    Parse-Json -Stock $stock | select "01. symbol","05. price"

    }
  }
}

 

#Post-StockPrice -> send price and symbol to sql database

 

function Post-StockPrice($lists){

 

    foreach($list in $lists){

        $date = Get-Date | Out-String

        $ticker = Get-StockPrice -tickers $list

        $symbol = $ticker.'01. symbol'

        $price = $ticker.'05. price'

        $conn=new-object System.Data.SqlClient.SQLConnection

        $ConnectionString = "Server=;Database=StockPrice;Integrated Security=True;Connect Timeout=5"

        $conn.ConnectionString=$ConnectionString

        $conn.Open()

        $commandText = "insert Table_2 Values('" + $symbol + "','" + $price + "','" + $date + "')"

        $command = $conn.CreateCommand()

        $command.CommandText = $commandText

        $command.ExecuteNonQuery()

 

        $conn.Close()

    }

}

 
 

#Import to Sql

 

#$conn=new-object System.Data.SqlClient.SQLConnection

#$ConnectionString = "Server=DESKTOP-BJPH0E1;Database=StockPrice;Integrated Security=True;Connect Timeout=5"

#$conn.ConnectionString=$ConnectionString

#$conn.Open()

#$commandText = "insert Table_2 Values('" + $symbol + "','" + $price + "')"

#$command = $conn.CreateCommand()

#$command.CommandText = $commandText

#$command.ExecuteNonQuery()

 

$conn.Close()

 

 

 

######################################################################################################################################################

 