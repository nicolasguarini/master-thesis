#versione 1.0 - 11/03/2024 Alessandro Ghirardi
param(
[string]$TenantUser,
[string]$TenantPassword,
[string]$adsyncUser,
[string]$adsyncPassword,
[string]$server,
[string]$adsyncServer,
[string]$serverExc,
[string]$domain,
[string]$disableDate,
[string]$userToDismiss,
[string]$OOOUser,
[string]$LogPath
)


if ([Environment]::Is64BitProcess -eq $false) {
    write-warning "Y'arg Matey, we're off to 64-bit land....."
    if ($myInvocation.Line) {
        write-host "1-Rilevato ambiente 32bit. Mi rilancio a 64bit "
        #write "Rilevato ambiente 32bit. Mi rilancio a 64bit " (Get-Date -Format dd-MM-yyyy_HH-mm-ss).ToString() ".<br>" | Out-File -Encoding "UTF8" -FilePath $logPath -append 
        &"$env:WINDIR\sysnative\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile $myInvocation.Line
    }else{
        write-host "2-Rilevato ambiente 32bit. Mi rilancio a 64bit "
        #write "Rilevato ambiente 32bit. Mi rilancio a 64bit " (Get-Date -Format dd-MM-yyyy_HH-mm-ss).ToString() ".<br>" | Out-File -Encoding "UTF8" -FilePath $logPath -append 
        write-host "args: " "$($myInvocation.InvocationName)" """$TenantUser""" """$TenantPassword""" """$adsyncUser""" """$adsyncPassword""" """$server""" """$adsyncServer""" """$serverExc""" """$domain""" """$DisableDate""" """$userToDismiss""" """$OOOUser""" """$LogPath"""
        &"$env:WINDIR\sysnative\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile -file "$($myInvocation.InvocationName)" """$TenantUser""" """$TenantPassword""" """$adsyncUser""" """$adsyncPassword""" """$server""" """$adsyncServer""" """$serverExc""" """$domain""" """$DisableDate""" """$userToDismiss""" """$OOOUser""" """$LogPath"""
    }
exit $lastexitcode
}

try{
    . "C:\SRPData\Script\SRP_Functions.ps1"
    $result = @()
    [string]$EXCAppId = "66e98152-6394-4d37-a49e-d02071538907"
    [string]$GraphAppId = "66e98152-6394-4d37-a49e-d02071538907"
    [string]$CertificateThumbprint = "7221B087BEAF8801FEA510948C899710CE5304DF"
    [string]$Organization = "ircaeu.onmicrosoft.com"
    [string]$tenantID="d6f90f31-7902-44ea-8438-bb423e4c7324"
    $cert = Get-ChildItem Cert:\LocalMachine\My\$CertificateThumbprint

    <#Attività da Eseguire
    1. Reset Password
    2. Rimozione utenti dal gruppo della VPN (vpngroup)
    3. Diabiled AD User
    4. Check ADM - Disable
    5. Hide From GAL
    6. Impostazione OOO

    #>

    #AD*************************************************************************************#
    $result += connectToDomainController $Server $LogPath   
    $DCPSSession=$result[-1].returnvalue 
   
    #1. Reset Password
    $Password = -join ((33..126) | Get-Random -Count 24 | ForEach-Object { [char]$_ })
    $result += SetAdUserPassword $domain $userToDismiss $password $logPath
    
    #2. Rimozione utenti dal gruppo della VPN 
    if(($domain -like "*@ircagroup.com")-or($domain -like "*@ircadistribuzione.com")-or($domain -like "*@irca.eu")-or($domain -like "*@ircacorp.com"))
    {
        $result += RemoveUserGroup "vpngroup" $userToDismiss $logPath
    }
    elseif(($domain -like "*@dobla.com")-or($domain -like "*@dobla.nl"))
    {
        $result += RemoveUserGroup "GL-SG-VPNAccess" $userToDismiss $logPath
    }

    #3. Disable AD account
    $result += disableADUser $domain $userToDismiss $logPath

    #4. Check ADM - Disable
    $ADMCheck += irca_CheckADM $domain $userToDismiss $logPath
    $result += $ADMCheck

    if($admcheck.Status -eq "KO")
    {
        $result += disableADUser $domain $admcheck.returnvalue
    }

    $CheckEmail = $null
    $CheckEmail = Get-ADUser $userToDismiss -Properties EmailAddress| select EmailAddress

    if(($checkEmail -ne $null)-and($checkEmail -ne ""))
    {
        if(($domain -like "*@ircagroup.com")-or($domain -like "*@ircadistribuzione.com")-or($domain -like "*@irca.eu")-or($domain -like "*@ircacorp.com"))
        {
            #5. Hide From GAL
            $result += connectToExchangeServer $serverExc $logPath
            $ExchSession=$result[-1].returnvalue
            $result += setRemoteMailboxHiddenAttribute $domain $server $userToDismiss "true" $logPath
            $result += RemoveSessions $ExchSession $LogPath
        }

        $result += connectToO365viaApp $EXCAppId $CertificateThumbprint $Organization $logPath
        if(($domain -like "*@ircagroup.com")-or($domain -like "*@ircadistribuzione.com")-or($domain -like "*@irca.eu")-or($domain -like "*@ircacorp.com"))
        {
            $companyOOO="IRCA S.p.A."
        }
        elseif(($domain -like "*@dobla.com")-or($domain -like "*@dobla.nl"))
        {
            $companyOOO="Dobla BV"
        }
        else
        {
            $companyOOO="Cesarin S.p.A."
        }

        if($OOOUser -ne "")
        {
            #O365*******************************************************************************#
            #6. Impostazione OOO
            $OOOSentTo=Get-ADUser $OOOUser -Properties EmailAddress,DisplayName| select EmailAddress, DisplayName
            $OOOActual=Get-ADUser $userToDismiss -Properties EmailAddress,DisplayName| select EmailAddress, DisplayName

            $message=
            "<P>
	            Please be informed that <strong>"+"$($OOOActual.DisplayName)"+"</strong> is no longer member of $($companyOOO) staff.
		        <li>In case, please contact <strong>"+"$($OOOSentTo.DisplayName)"+" - "+"$($OOOSentTo.EmailAddress)"+" </strong>.<br><br></li>
	
                Best Regards.<br>
	            $($companyOOO)<br>
            </p>"

            #Recupera UPN dell'account
            $userToDismiss=Get-ADUser $userToDismiss |Select-Object UserPrincipalName -ExpandProperty UserPrincipalName

            $result += setMailboxOutOfOffice $userToDismiss $message $message $logPath
        }
    }

    #Mail
    $SmtpServer="172.22.21.16"
    $from="NoReply_SRP@irca.eu"
    $to="HelpDesk@elmec.it"
    $subject="User to dismiss $($userToDismiss)"
    $BodyHTML = "
        <font face=""verdana"">
            <p>  
                Hi,<br><br>

                The account in subject has been dismissed!<br>
                Please, plan the ticket for the deletion date (1 month)<br>
                Thank you!<br><br>

                Kind regards<br>
                Team Elmec<br>
                <hr>
                This is an automatic email, don't reply to it.<br>
                For any issue, please contact <strong>helpdesk@elmec.it</strong>
            </p>
        </font>"
    $cc=""
    $credential=$null

    $result += SendEmail $From $to $subject $BodyHTML $cc $SmtpServer $Credential $logPath
    
    $result += RemoveSessions_O365_EXC

    $formatResult = formatResult $result $LogPath

    if(!$result.Status.Contains("KO")){
        write "SCRIPT COMPLETATO CON SUCCESSO<br>" | Out-File -Encoding "UTF8" -FilePath $logPath -append
    }
    else{
        write "RILEVATO ERRORE DURANTE ESECUZIONE<br>" | Out-File -Encoding "UTF8" -FilePath $logPath -append        
    }        
}
catch{
    $formatResult = formatResult $result $LogPath
    $("RILEVATO ERRORE DURANTE ESECUZIONE <br>" + $_.Exception.Message) | out-file -Encoding "UTF8" -FilePath $logPath -append   
}