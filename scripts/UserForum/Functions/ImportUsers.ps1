Import-Function -Name Convert-CSVMediaItemToArray #Converts file from stream to CSV array
Import-Function -Name Download-ExcelMediaItem #Downloads file to server, imports data into array and removed file from server
Import-Function -Name Import-UsersFromList #Creates users from array

$result = Read-Variable -Width 500 -Height 380 -OkButtonName "Import" -CancelButtonName "Cancel" `
        -Parameters @{ Name = "FilesToUpload"; Title="Upload Lists Wizard"; Source="DataSource=/sitecore/media library/UserForumImports&DatabaseName=master"; editor="treelist"} `
        -Title "Import users Wizard" -Description "Please select the file you'd like to import" `
       
if ($result -ne "ok") { exit }

Write-log "$($FilesToUpload.Count) files selected"

foreach($file in $filestoupload)
{
    switch -wildcard ($file.extension) {
        "csv" {
            write-log "Processing CSV File $file"
            $data = Convert-CSVMediaItemToArray $file 
        }
        "xls*" { 
            write-log "Processing Excel File $file"
            $data = Download-ExcelMediaItem $file 
        }
    }
    
    $data = $data | where-object { $_."Email Address" -ne "" -and $_.email -ne "" }
    
    write-log "Importing Users..."
    Import-UsersFromList $data    
}