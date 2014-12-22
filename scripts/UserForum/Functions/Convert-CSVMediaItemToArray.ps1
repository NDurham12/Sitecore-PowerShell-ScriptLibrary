function Convert-CSVMediaItemToArray{
    <#
        .SYNOPSIS
            Downloads CSV Media Item and returns array of all rows.
            
        .EXAMPLE
            Returns array values of excel rows
            
            PS master:\> Convert-CSVMediaItemToArray $Item

        .NOTES
            Nona Durham
            
    #>
    param ([Sitecore.Data.Items.MediaItem]$MediaItem)
    $stream = $mediaItem.GetMediaStream()
    $buffer = new-object Byte[] $($stream.Length)
    $read = $stream.Read($buffer, 0, $stream.Length)
    $stream.Close();
    $enc =  [System.Text.Encoding]::ASCII
    $output = $enc.GetString($buffer)
    
    return convertfrom-csv $output
}