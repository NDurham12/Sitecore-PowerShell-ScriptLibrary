function Import-Excel {
    <#
        .SYNOPSIS
            Returns array of all excel rows. 
        
        .EXAMPLE
            Returns array values of excel rows
            
            PS master:\> Import-Excel -FileName C:\MyExcelFile.xls

        .NOTES
            Nona Durham
			http://www.codeproject.com/Articles/670082/Use-Excel-in-PowerShell-without-a-full-version-of
            
    #>
    param (
    [string]$FileName
    )

    if ($FileName -eq "") {
        throw "Please provide path to the Excel file"
        Exit
    }
    
    if (-not (Test-Path $FileName)) {
        throw "Path '$FileName' does not exist."
        Exit
    }

    $connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='$FileName';Extended Properties='Excel 12.0 Xml;HDR=YES';"
    $conn = New-Object System.Data.OleDb.OleDbConnection($connectionString)
    $conn.open()
    
    $dt = $conn.GetSchema("Tables")
    $worksheets = $dt | select Table_name
    
    if ($worksheets.count -gt 1) {
       
        $selection = New-Object System.Collections.Specialized.OrderedDictionary
       
        
        foreach($sheet in $worksheets ) {
            $selection.Add($($sheet.Table_name),$($sheet.Table_name))  
        }
        
        $result = Read-Variable -Parameters `
                    @{ Name = "selectedWorksheet"; Title="Available Worksheets"; Options=$selection; Editor="radio"; Tooltip="Select worksheet to import";} `
                     -Description "Multiple worksheets found, please select the worksheet to import" `
                    -Title "Workseet Selector" -Width 500 -Height 600 -OkButtonName "Import" -CancelButtonName "Abort" -ShowHints
    }
    else {
        $selectedWorksheet = $worksheets[0].Table_Name
    }
    
    if ($result -ne "ok") {
        exit
    }
        
    write-log -log info " the worksheet selected was $selectedWorksheet"
    $query = 'select * from ['+$selectedWorksheet+']';
    
    $cmd = New-Object System.Data.OleDb.OleDbCommand($query,$conn) 
    $dataAdapter = New-Object System.Data.OleDb.OleDbDataAdapter($cmd) 
    $dataTable = New-Object System.Data.DataTable 

    $dataAdapter.fill($dataTable)
    $conn.close()

    $myDataRow ="";
    $columnArray =@();
    foreach($col in $dataTable.Columns) {
        $columnArray += $col.toString();
    }

    $returnObject = @();
    foreach($rows in $dataTable.Rows)
    {
        $i=0;
        $rowObject = @{};
        foreach($columns in $rows.ItemArray){
            $rowObject += @{$columnArray[$i]=$columns.toString()};
            $i++;
        } 

        $returnObject += new-object PSObject -Property $rowObject;
    }

    return $returnObject;
} 
