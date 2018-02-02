# Set-ExecutionPolicy RemoteSigned to run script, or copy paste the code into powershell directly

if([IntPtr]::Size * 8 -ne 32) {
    #Powershell (x86) is needed to properly to connect to (32bit) quickbooks
    Write-Host "Powershell is running 64bit process"
    Write-Host "Relaunching as 32-bit powershell session"
    C:\Windows\SysWOW64\WindowsPowerShell\v1.0\PowerShell.exe -File $MyInvocation.MyCommand.Path
} else {
    try {
        $qb = New-Object -ComObject QBFC13.QBSessionManager
    } catch {
        Write-Warning -Message "Could not import ComObject because $($_.Exception.Message)"
    }
    

    # Will try to access open quickbooks file if blank.
    $qbFile = ""

    # Change app name to something unique if you're doing more than testing
    $qbAppName = "Test QBFC Request"

    # you may need to allow access in quickbooks
    $qb.OpenConnection($qbFile, $qbAppName)

    $qb.BeginSession("", 0)

    $reqMsg = $qb.CreateMsgSetRequest("US", 6, 0)

    $reqMsg.AppendInventoryAdjustmentQueryRq()

    $resMsg = $qb.DoRequests($reqMsg)

    $qb.EndSession()
    $qb.CloseConnection()

    $QBXML = $resMsg

    try {
        $QBXMLMsgRq = $QBXML.ResponseList
    } catch {
        Write-Warning -Message "Could not get InvAdjustQuery because $($_.Exception.Message)"
    }

    $InvAdjQueryRes = $QBXMLMsgRq.GetAt(0)
    
    for ($i = 0; $i -lt $InvAdjQueryRes.Detail.Count; $i++) {
        Write-Host "Inventory Adjustment # $($i)"
            try {
                $InvAdjRet = $InvAdjQueryRes.Detail.GetAt($i);
                $TxnID = $InvAdjRet.TxnID.GetValue()
                try {
                    $memo = $InvAdjRet.Memo.GetValue()
                } catch {
                    # Create Memo if Null
                    $memo = "no memo"
                    #Write-Warning -Message "Could not get Memo because $($_.Exception.Message)"  
                }
            } catch {
                Write-Warning -Message "Could not get Inventory Adjustment because $($_.Exception.Message)"
            }
        Write-Host $TxnID
        Write-Host $memo
    }
}
