

Measure-Command {

    for($i=1; $i -le 10; $i++) {

        Invoke-SQL "rtix-sqtt-01" "ttrack2000" "select top 10 * from IDT_Articles"

    }

    for($i=1; $i -le 10; $i++) {

        Invoke-SQL "rtix-sqap-01" "chart" "select top 10 * from ApplicationLoginFeatures"

    }

}



function Invoke-SQL {
    param(
        [string] $dataSource = "",
        [string] $database = "ttrack2000",
        [string] $sqlCommand = $(throw "Please specify a query.")
      )

    $connectionString = "Data Source=$dataSource; " +
            "Integrated Security=SSPI; " +
            "Initial Catalog=$database"

    $connection = new-object system.data.SqlClient.SQLConnection($connectionString)
    $command = new-object system.data.sqlclient.sqlcommand($sqlCommand,$connection)
    $connection.Open()

    $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $command
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataSet) | Out-Null

    $connection.Close()
    $dataSet.Tables.Count

}