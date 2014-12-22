function Import-UsersFromList{
    param ($Users)
    $usersImported = 0
    $OutputLog = @()
    
    write-output "There are $($Users.count) to import"
    
    foreach($userInfo in $Users){
        $indexCt += 1
        $userInfo
        $email = ($userInfo."email address" + $userInfo."email" + $userInfo."emailaddress" + "").trim()  #handles multiple columsn for email address
        $userFullname = $userInfo."FullName" + $userInfo."First Name" + $userInfo."FirstName" + " " + $userInfo."Last Name" + $userInfo."LastName"
        
        write-log "processing $email $userFullname"
        Write-Progress -PercentComplete ($indexCt / $Users.count) -Activity "Importing Users" -CurrentOperation "Importing: $email" -Status "Importing"
        
        if ($email.length -eq 0 -or $email -eq $null) { 
            $OutputLog += @{status = "Error"; LogInfo = "Email is blank for user $userFullname"; username = $userFullname; email = $email}
            continue
        }
        
        $ValidEmail = $email -match "^[a-zA-Z0-9_-]+(?:\.[a-zA-Z0-9_-]+)*@(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]*[a-zA-Z0-9])?\.)+[a-zA-Z0-9](?:[a-zA-Z0-9-]*[a-zA-Z0-9])?$"
        
        if (-not $ValidEmail) {
           $OutputLog += @{status = "Error"; LogInfo = "Email Address `"$email`" Is NOT valid"; username = $userFullname; email = $email }
            continue
        }
        
        $userName = "extranet\$email"
        $founduser = Get-User -id $userName -ErrorAction SilentlyContinue
         
        if ($foundUser){
            $OutputLog += @{status = "Warning"; LogInfo = "Skipping because user $userName found"; username = $userFullname; email = $email}
        }
        else 
        {
            $comment = "User Forum " + $userInfo.Company + $userInfo.Comment + ""
            $usersImported += 1
            $OutputLog += @{status = "Info"; LogInfo ="Creating User $userFullname $email $comment"; username = $userFullname; email = $email}
            
            $newUser = New-User -id $userName -enabled -password djwu*27dj -email $email  -comment $comment
            
            try {
                set-user -id $userName -FullName $userFullname -ErrorAction SilentlyContinue    
            }
            catch{
             
            }
            
            add-rolemember -identity extranet\UserForumWHUsers -members $userName
        }
    }

    Write-output "Total Users Imported: $usersImported"
    Write-log  "Total Users Imported: $usersImported"
    
    $props = @{
        InfoTitle = "User Forum WH Users"
        InfoDescription = "Lists the users imported for the User Forum."
        PageSize = 200
    }
    
    $OutputLog | Show-ListView @props -Property @{Label="Email"; Expression={$_.Email} }, @{Label="Info"; Expression={$_.LogInfo} }, @{Label="Status"; Expression={$_.Status} }
    
    close-window
}