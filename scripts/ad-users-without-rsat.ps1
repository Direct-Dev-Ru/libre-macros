# Полный скрипт: пользователи из OU через LDAPS
# Ниже — самодостаточный скрипт на Windows PowerShell 5.1 (ADSI / System.DirectoryServices).
# OU задаётся путём name_1/name_2/... (сверху внiz: от корня домена к вложенной OU).

#Requires -Version 5.1
<#
.SYNOPSIS
  Выгрузка пользователей из указанной OU через LDAPS.
.PARAMETER DomainDns
  DNS-имя домена, например: example.com
.PARAMETER OuPath
  Путь к OU через слэш: Sales/Employees/Active
.PARAMETER LdapServer
  FQDN контроллера или балансировщика LDAPS, например: dc01.example.com
.PARAMETER LdapPort
  Порт LDAPS (обычно 636)
.PARAMETER OutputCsv
  Куда сохранить CSV
.PARAMETER IgnoreCertErrors
  Игнорировать ошибки SSL-сертификата (только для тестов!)
#>
param(
    [Parameter(Mandatory = $true)]
    [string] $DomainDns,                 # example.com
    [Parameter(Mandatory = $true)]
    [string] $OuPath,                    # name_1/name_2/name_3
    [Parameter(Mandatory = $true)]
    [string] $LdapServer,                # dc01.example.com
    [int] $LdapPort = 636,
    [string] $OutputCsv = ".\ad_users.csv",
    [switch] $IgnoreCertErrors,
    [System.Management.Automation.PSCredential] $Credential
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
function ConvertTo-DomainDn {
    param([string] $DnsName)
    ($DnsName.Split('.') | ForEach-Object { "DC=$_" }) -join ','
}
function ConvertTo-OuDn {
    param(
        [string] $OuPath,
        [string] $DomainDn
    )
    $parts = $OuPath.Trim('/').Split('/', [System.StringSplitOptions]::RemoveEmptyEntries)
    if ($parts.Count -eq 0) {
        throw "OuPath пустой. Пример: Sales/Employees/Active"
    }
    # name_1/name_2 -> OU=name_2,OU=name_1,DC=...
    $ouParts = for ($i = $parts.Count - 1; $i -ge 0; $i--) {
        "OU=$($parts[$i])"
    }
    ($ouParts + $DomainDn) -join ','
}
function Enable-LdapsCertBypass {
    # Только для лаборатории / самоподписанных сертификатов
    if (-not ("TrustAllCerts" -as [type])) {
        Add-Type @"
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
public static class TrustAllCerts {
    public static bool Validator(
        object sender,
        X509Certificate certificate,
        X509Chain chain,
        SslPolicyErrors sslPolicyErrors) {
        return true;
    }
    public static void Enable() {
        ServicePointManager.ServerCertificateValidationCallback =
            new RemoteCertificateValidationCallback(Validator);
    }
}
"@
    }
    [TrustAllCerts]::Enable()
}
function Get-LdapProperty {
    param(
        [System.DirectoryServices.ResultPropertyCollection] $Props,
        [string] $Name
    )
    if (-not $Props.Contains($Name)) { return $null }
    $val = $Props[$Name]
    if ($null -eq $val -or $val.Count -eq 0) { return $null }
    if ($val.Count -eq 1) { return [string]$val[0] }
    return ($val | ForEach-Object { [string]$_ }) -join '; '
}
# --- подготовка ---
if ($IgnoreCertErrors) {
    Write-Warning "IgnoreCertErrors включён — не используйте в production."
    Enable-LdapsCertBypass
}
$domainDn = ConvertTo-DomainDn -DnsName $DomainDns
$ouDn     = ConvertTo-OuDn -OuPath $OuPath -DomainDn $domainDn
# SearchRoot = конкретная OU; Subtree = OU + все вложенные OU
$ldapPath = "LDAPS://${LdapServer}:${LdapPort}/${ouDn}"
Write-Host "Domain DN : $domainDn"
Write-Host "OU DN     : $ouDn"
Write-Host "LDAP path : $ldapPath"
# --- подключение ---
if ($Credential) {
    $user = $Credential.UserName
    $pass = $Credential.GetNetworkCredential().Password
    $searchRoot = New-Object System.DirectoryServices.DirectoryEntry(
        $ldapPath, $user, $pass,
        [System.DirectoryServices.AuthenticationTypes]::Secure
    )
} else {
    $searchRoot = New-Object System.DirectoryServices.DirectoryEntry(
        $ldapPath,
        $null, $null,
        [System.DirectoryServices.AuthenticationTypes]::Secure `
        -bor [System.DirectoryServices.AuthenticationTypes]::Sealing
    )
}
# Проверка, что OU доступна
$null = $searchRoot.RefreshCache()
# --- поиск ---
$searcher = New-Object System.DirectoryServices.DirectorySearcher
$searcher.SearchRoot  = $searchRoot
$searcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
# только пользователи (не компьютеры и т.п.)
$searcher.Filter = "(&(objectCategory=person)(objectClass=user))"
# атрибуты для выгрузки — допишите свои
$attrs = @(
    'distinguishedName',
    'sAMAccountName',
    'userPrincipalName',
    'displayName',
    'givenName',
    'sn',
    'mail',
    'title',
    'department',
    'company',
    'telephoneNumber',
    'mobile',
    'enabled',              # может не вернуться через ADSI — см. ниже
    'userAccountControl',
    'whenCreated',
    'whenChanged',
    'lastLogonTimestamp',
    'memberOf'
)
foreach ($a in $attrs) {
    [void]$searcher.PropertiesToLoad.Add($a)
}
# постраничная выборка (важно для больших OU)
$searcher.PageSize = 1000
$results = $searcher.FindAll()
Write-Host "Найдено записей: $($results.Count)"
# --- формирование таблицы ---
$rows = foreach ($r in $results) {
    $p = $r.Properties
    $uac = Get-LdapProperty $p 'userAccountControl'
    $enabled = $null
    if ($null -ne $uac) {
        # ACCOUNTDISABLE = 0x2
        $enabled = (-bnot ([int]$uac -band 2))
    }
    [pscustomobject]@{
        distinguishedName  = Get-LdapProperty $p 'distinguishedName'
        sAMAccountName     = Get-LdapProperty $p 'sAMAccountName'
        userPrincipalName  = Get-LdapProperty $p 'userPrincipalName'
        displayName        = Get-LdapProperty $p 'displayName'
        givenName          = Get-LdapProperty $p 'givenName'
        sn                 = Get-LdapProperty $p 'sn'
        mail               = Get-LdapProperty $p 'mail'
        title              = Get-LdapProperty $p 'title'
        department         = Get-LdapProperty $p 'department'
        company            = Get-LdapProperty $p 'company'
        telephoneNumber    = Get-LdapProperty $p 'telephoneNumber'
        mobile             = Get-LdapProperty $p 'mobile'
        enabled            = $enabled
        userAccountControl = $uac
        whenCreated        = Get-LdapProperty $p 'whenCreated'
        whenChanged        = Get-LdapProperty $p 'whenChanged'
        lastLogonTimestamp = Get-LdapProperty $p 'lastLogonTimestamp'
        memberOf           = Get-LdapProperty $p 'memberOf'
    }
}
$rows |
    Sort-Object sAMAccountName |
    Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
Write-Host "Сохранено: $OutputCsv"
# освобождение COM-объектов
$results.Dispose()
$searcher.Dispose()
$searchRoot.Dispose()


# Как запускать
# Интерактивный ввод пароля
# $cred = Get-Credential   # DOMAIN\user или user@example.com
# .\Export-AdUsersFromOu.ps1 `
#     -DomainDns   "example.com" `
#     -OuPath      "Company/Sales/Employees" `
#     -LdapServer  "dc01.example.com" `
#     -LdapPort    636 `
#     -Credential  $cred `
#     -OutputCsv   "C:\Temp\sales_users.csv"
# Если сертификат LDAPS самоподписанный (только для теста):

# .\Export-AdUsersFromOu.ps1 ... -IgnoreCertErrors
# Как формируется OU
# OuPath	DN
# Sales
# OU=Sales,DC=example,DC=com
# Company/Sales/Active
# OU=Active,OU=Sales,OU=Company,DC=example,DC=com
# Порядок: первый сегмент — ближе к корню домена, последний — целевая OU.

# Частые проблемы с LDAPS
# «The server is not operational»

# неверный FQDN сервера;
# порт 636 закрыт;
# на DC не настроен сертификат для LDAPS;
# OU-путь не совпадает с реальными именами (регистр обычно не важен).
# Ошибка SSL / certificate

# в production — установить доверенный сертификат на DC;
# в тесте — -IgnoreCertErrors.
# «Access is denied»

# нужна учётная запись с правами чтения на эту OU (-Credential).
# Пустой результат

# пользователи могут быть во вложенных OU — скрипт ищет с Subtree (включая дочерние OU);
# если нужны только прямые дети OU, замените:
# $searcher.SearchScope = [System.DirectoryServices.SearchScope]::OneLevel
# lastLogonTimestamp

# приходит как Windows FILETIME; для читаемой даты:
# if ($ts = Get-LdapProperty $p 'lastLogonTimestamp') {
#     [datetime]::FromFileTime([int64]$ts)
# }
# Альтернатива: если RSAT уже есть
# При установленном модуле ActiveDirectory и LDAPS на DC проще так:

# $ouDn = "OU=Active,OU=Sales,OU=Company,DC=example,DC=com"
# Get-ADUser -LDAPFilter "(objectCategory=person)" `
#     -SearchBase $ouDn `
#     -SearchScope Subtree `
#     -Properties displayName, mail, department |
#     Select-Object SamAccountName, DisplayName, Mail, Department |
#     Export-Csv users.csv -NoTypeInformation -Encoding UTF8
# Но для чистого ADSI + LDAPS без RSAT достаточно первого скрипта.

# Если нужно — могу дописать вариант с рекурсивным разрешением OU по CN (когда имена OU не уникальны и путь name_1/name_2 может быть неоднозначным).