
Connect-AzureAD
Connect-ExchangeOnline 

$loop = $true
$forwardEnabled = $false
$autoReplyEnabled = $false
$forwardEmail = ""
$mailboxTooLarge = $false
$autoReplyMessage = ""
do {
write-host "___________                     ____ ___                    " -ForegroundColor Blue
write-host "\__    ___/__________  _____   |    |   \______ ___________ " -ForegroundColor Blue
write-host "  |    |_/ __ \_  __ \/     \  |    |   /  ___// __ \_  __ \" -ForegroundColor Blue
write-host "  |    |\  ___/|  | \/  Y Y  \ |    |  /\___ \\  ___/|  | \/" -ForegroundColor Blue
write-host "  |____| \___  >__|  |__|_|  / |______//____  >\___  >__|   " -ForegroundColor Blue
write-host "             \/            \/               \/     \/       " -ForegroundColor Blue


$termedEmail = read-host "[PROMPT] Please enter the UPN of the user you would like to terminate"
$mailbox = Get-Mailbox -Identity $termedEmail -ErrorAction SilentlyContinue
try{
    $termedUser = Get-AzureADUser -ObjectId $termedEmail
}
catch {
    write-host "User not found. Please try again" -ForegroundColor Red
   
    Continue
}
write-host "Display Name $($termedUser.DisplayName)" 
write-host "Email $($termedUser.UserPrincipalName)"
write-host "Job Title $($termedUser.JobTitle)"
write-host "Department $($termedUser.Department)"

switch(read-host "Is this the user you would like to terminate? (Y/N)") {
"y" {
    $confirmed = $false
    do {
    switch($reconfirm = read-host "Please re-enter the address of the user to confirm termination") {
        $termedEmail {
            $confirmed = $true
        }
        "c"{ 
            write-host "Termination cancelled" -ForegroundColor Red
            Exit
        }
        default {
            write-host "The email did not match. Try again or enter c to cancel" -ForegroundColor Red
        }
    }
    } until ($confirmed-eq $true)
    $forwardConfirmed = $false
    do {
    switch(read-host "Does this address need to forward to another user? (Y/N)") {
        "y" {
            $forwardEnabled = $true
            do {
                $forwardEmail = read-host "Please enter the email address to forward to"
                write-host $forwardEmail
                switch(read-host "Is this the correct email address? (Y/N)") {
                    "y" {
                        $forwardConfirmed = $true
                    }
                    "n" {
                        write-host "Please try again" -ForegroundColor Red
                    }
                    default {
                        write-host "Invalid response. Please try again." -ForegroundColor Red
                    }
                }


            } until ($forwardConfirmed -eq $true)

            
        }
        "n" {
            $forwardConfirmed = $true
        }
        default {
            write-host "Invalid response. Please try again." -ForegroundColor Red
        }
    }

    } until ($forwardConfirmed -eq $true)
    $autoReplyConfirmed = $false
    do {
    switch(read-host "Does this user need to have an auto-reply enabled? (Y/N)") {
        "y" {
            $autoReplyEnabled = $true
            do {
                $autoReplyMessage = read-host "Please enter the message you would like to send"
                write-host $autoReplyMessage
                switch(read-host "Is this the correct message? (Y/N)") {
                    "y" {
                        $autoReplyConfirmed = $true
                    }
                    "n" {
                        write-host "Please try again" -ForegroundColor Red
                    }
                    default {
                        write-host "Invalid response. Please try again." -ForegroundColor Red
                    }
                }
            } until ($autoReplyConfirmed -eq $true)
           
        }
        "n" {
            $autoReplyConfirmed = $true
        }
        default {
            write-host "Invalid response. Please try again." -ForegroundColor Red
        }
    }
    } until ($autoReplyConfirmed -eq $true )

    write-host "Blocking User Sign In ...." -NoNewLine 
    try {
        Set-AzureADUser -ObjectId $termedEmail -AccountEnabled $false
    }
    catch {
        write-host "  --FAILED" -ForegroundColor Red
        write-host $_.Exception.Message -ForegroundColor Red
        read-host "Press enter to exit the script"
        Exit
    }
    write-host "  --COMPMETE" -ForegroundColor Green 
    write-host "Siging out of all sessions...." -NoNewLine
    try {
        Get-AzureADUser -ObjectId $termedEmail | Revoke-AzureADUserAllRefreshToken
    }
    catch {
        write-host "  --FAILED" -ForegroundColor Red
        write-host $_.Exception.Message -ForegroundColor Red
        read-host "Press enter to exit the script"
        Exit
    }
    write-host "  --COMPLETE" -ForegroundColor Green
    if ($forwardEnabled -eq $true) {
        write-host "Enabling forwarding to $forwardEmail" -NoNewLine
        try {
            Set-Mailbox -Identity $termedEmail -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $forwardEmail
        }
        catch {
            write-host "  --FAILED" -ForegroundColor Red
            write-host $_.Exception.Message -ForegroundColor Red
            read-host "Press enter to exit the script"
            Exit
        }
        write-host "  --COMPLETE" -ForegroundColor Green
    }
    if ($autoReplyEnabled -eq $true) {
        write-host "Enabling auto-reply" -NoNewLine
        try {
            Set-MailboxAutoReplyConfiguration -Identity $termedEmail -AutoReplyState Enabled -InternalMessage "$autoReplyMessage" -ExternalMessage "$autoReplyMessage" -ExternalAudience All
        }
        catch {
            write-host "  --FAILED" -ForegroundColor Red
            write-host $_.Exception.Message -ForegroundColor Red
            read-host "Press enter to exit the script"
            Exit
        }
        write-host "  --COMPLETE" -ForegroundColor Green
    }
    write-host "Converting to shared mailbox ..." -NoNewLine
    try {
        $stats = Get-EXOMailboxStatistics -Identity $termedEmail
        if($stats.TotalItemSize.Value.ToGB() -lt 50) {
            Set-Mailbox -Identity $termedEmail -Type Shared
            
            write-host "  --COMPLETE" -ForegroundColor Green
        }
        else {
            write-host " --FAILED" -ForegroundColor Red
            write-host "Mailbox is too large to convert to shared mailbox" -ForegroundColor Red
            $mailboxTooLarge = $true
          
        }
    }
    catch {
        write-host " --FAILED" -ForegroundColor Red
        write-host $_.Exception.Message -ForegroundColor Red
        read-host "Press enter to exit the script"
        Exit
    }
    
    if($mailboxTooLarge -eq $true) {
        write-host "Can not remove license because mailbox is too large to convert to shared mailbox" -ForegroundColor Red
       
    }
    else {
        write-host "Removing all Licenses from user ..." -NoNewLine
    try {
        $Skus = Get-AzureADUser -ObjectId $termedEmail | Select -ExpandProperty AssignedLicenses | Select SkuID
        if($Skus -is [array])
    {
        $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        for ($i=0; $i -lt $Skus.Count; $i++) {
            $licenses.RemoveLicenses +=  (Get-AzureADSubscribedSku | Where-Object -Property SkuID -Value $Skus[$i].SkuId -EQ).SkuID   
        }
        Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $licenses
    } else {
        $licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $licenses.RemoveLicenses =  (Get-AzureADSubscribedSku | Where-Object -Property SkuID -Value $Skus.SkuId -EQ).SkuID
        Set-AzureADUserLicense -ObjectId $termedEmail -AssignedLicenses $licenses
    }
    }
    catch {
        write-host "  --FAILED" -ForegroundColor Red
        write-host $_.Exception.Message -ForegroundColor Red
        read-host "Press enter to exit the script"
        Exit
    }
    write-host "  --COMPLETE" -ForegroundColor Green
    }
    write-host "User termination is complete" -ForegroundColor Green
    read-host "Press any key to exit the script"
    $loop = $false
   
}
"n" {
    write-host "Starting script from start" -ForegroundColor Green
Continue
}
default {

 write-host "[ERROR] Invalid response. Please try again." -ForegroundColor Red
}
}

} until ($loop -eq $false) 
