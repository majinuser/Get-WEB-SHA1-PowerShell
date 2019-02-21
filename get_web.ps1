$ConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;data source=D:\Object\wget\web\script\DB.accdb;"
$sqlstr = 'SELECT * FROM web_sit where url_status="1"'
$conn = New-Object -ComObject ADODB.Connection
$conn.Open($ConnStr)
$rs = $conn.Execute($sqlstr)
 
$results = While ($rs.EOF -eq $false)
{
    $CustomObject = New-Object -TypeName PSObject
    $rs.Fields | ForEach-Object -Process {
        $CustomObject | Add-Member -MemberType NoteProperty -Name $_.Name -Value $_.Value
    }
    $CustomObject
    $rs.MoveNext()
}
#$results | Out-String
$rs.close()
$conn.close()
for($i=0;$i -le $results.Count-1; $i=$i+1)
{
$url_add=$results[$i].url_add
$url_id=$results[$i].url_id
cmd /c "wget.exe -O $url_id $url_add"
$hash=get-filehash -path $url_id -Algorithm sha1 | Out-String
$sha=$hash.split(" ",[StringSplitOptions]::RemoveEmptyEntries)[7]

$objOleDbConnection = New-Object "System.Data.OleDb.OleDbConnection"
$objOleDbCommand = New-Object "System.Data.OleDb.OleDbCommand"
$objOleDbAdapter = New-Object "System.Data.OleDb.OleDbDataAdapter"
$objDataTable = New-Object "System.Data.DataTable"
$objOleDbConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;data source=D:\Object\wget\web\script\DB.accdb;"
$objOleDbConnection.Open()
$objOleDbConnection.State
$objOleDbCommand.Connection = $objOleDbConnection
$objOleDbCommand.CommandText = "insert into web_tab (url_id,SHA1) values ('$url_id','$sha')"
$objOleDbCommand.CommandText
$objOleDbCommand.ExecuteNonQuery()
$objOleDbConnection.Close()
$objOleDbConnection.State
}
