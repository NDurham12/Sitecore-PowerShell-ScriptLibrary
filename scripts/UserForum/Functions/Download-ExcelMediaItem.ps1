Import-Function -Name Import-Excel

function Download-ExcelMediaItem{
    <#
        .SYNOPSIS
            Downloads Excel Media Item and returns array of all excel rows. References Import-Excel Function
            
        .EXAMPLE
            Returns array values of excel rows
            
            PS master:\> ExcelMediaItem $Item

        .NOTES
            Nona Durham
            
    #>
    param ([Sitecore.Data.Items.MediaItem]$MediaItem)
    $dateTime = Get-Date -format yyyyMMdd_hhmmssff
    $dataFolder = [Sitecore.Configuration.Settings]::DataFolder
    $fileName = "$dataFolder\$dateTime.$($MediaItem.Name).$($mediaItem.Extension)"  #unique file name to save file as 
    
    $stream = ($mediaItem).GetMediaStream()                         #get stream
    $buffer = new-object Byte[] $($stream.Length)                   #create buffer
    $loaded = $stream.Read($buffer, 0, $stream.Length)              #read buffer
    $stream.Close();                                                #close stream
    
    $output = [System.IO.File]::WriteAllBytes($fileName,$buffer)    #write file to filesystem
    $data =  Import-Excel $fileName                                 #Import data from excel into array
    Remove-Item $fileName                                           #delete file from filesystem
    return $data
}