
Connect-ExchangeOnline
#if($null -eq $aadToken){
#
#$context = [Microsoft.Azure.Commands.Common.Authentication.Abstractions.AzureRmProfileProvider]::Instance.Profile.DefaultContext
#$aadToken = [Microsoft.Azure.Commands.Common.Authentication.AzureSession]::Instance.AuthenticationFactory.Authenticate($context.Account, $context.Environment, $context.Tenant.Id.ToString(), $null, [Microsoft.Azure.Commands.Common.Authentication.ShowDialog]::Never, $null, "https://graph.windows.net").AccessToken
#Connect-MsolService -AdGraphAccessToken $aadToken
#Connect-AzureAD -AadAccessToken $aadToken -AccountId $context.Account.Id -TenantId $context.tenant.id
#Connect-ExchangeOnline -
#}


$block = {
    param($UPN)
    if(){
    }
    return $retval
}

$MaxThreads = 100
$RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
$RunspacePool.Open()
$jobs = New-Object System.Collections.ArrayList
$i = 1
foreach ($u in $thelist) {
    $Params = @{
        upn = $u.UserPrincipalName
    }
    $PowerShell = [powershell]::Create()
    $PowerShell.RunspacePool = $RunspacePool
    [void]$PowerShell.AddScript($block)
    [void]$powerShell.AddParameters($Params)
    $Handle = $PowerShell.BeginInvoke()

    $temp = "" | Select PowerShell, returnval, handle
    $temp.PowerShell = $PowerShell
    $temp.returnval = $returnval
    $temp.handle = $handle

    [void]$jobs.Add($temp)

    Write-host ("Available Runspaces in RunspacePool: {0}" -f $RunspacePool.GetAvailableRunspaces())
    write-host ("Remaining Jobs: {0}" -f @($jobs | Where {
                $_.handle.iscompleted -ne 'Completed'
            }).Count)
    #$status = [string]$i +'/' + $thelist.count
    #Write-Progress -Activity "Running" -Status $status -PercentComplete ($i /$thelist.Count *100)   
    #$i++
}

$return = $jobs | ForEach {
    $_.powershell.EndInvoke($_.handle)
    $_.PowerShell.Dispose()
}
$jobs.clear()

$return.rows