# パラメータ設定
# $i : CSV ファイルのパス
Param (
    $i
)

# Credential
$UserCredential = Get-Credential

# Exchange Online セッション設定
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session

# CSV ファイルを読み込み、実行
import-csv -Encoding Default $i | Foreach-Object {
    $args = @{
        Identity=$_."AccountID"
        #ForwardingSmtpAddress=$_."ForwardingSmtpAddress"
        DeliverToMailboxAndForward=$true
    }

    if($_."ForwardingSmtpAddress" -ne "") {
        Set-Mailbox @args -ForwardingSmtpAddress $_."ForwardingSmtpAddress"
    }
    else {
        Set-Mailbox @args -ForwardingSmtpAddress $null
    }
} 
