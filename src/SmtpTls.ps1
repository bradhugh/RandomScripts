[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]
    $Hostname,

    [Parameter(Mandatory=$false)]
    [int]
    $Port = 587
)

$ErrorActionPreference = "Stop"

$reader = $null
$writer = $null
$ssl = $null
$tcp = $null

try {
  $tcp = new-object System.Net.Sockets.TcpClient
    $tcp.Connect($hostname, $port)
    $encoding = New-Object System.Text.Utf8Encoding($false)
    $reader = New-Object System.IO.StreamReader($tcp.GetStream(), $encoding, $false, 1024, $true)
    $writer = New-Object System.IO.StreamWriter($tcp.GetStream(), $encoding, 1024, $true)
    $writer.AutoFlush = $true
    $reader.ReadLine()
    $writer.WriteLine("EHLO mail.contoso.com")

    $line = $null
    do
    {
      $line = $reader.ReadLine()
      Write-Output $line
    } while (-not $line.StartsWith("250 "))

    $writer.WriteLine("STARTTLS")
    $reader.ReadLine()
    $reader.Close()
    $writer.Close()

    $ssl = new-object System.Net.Security.SslStream($tcp.GetStream())
    $ssl.AuthenticateAsClient($hostname)

    $ssl | Format-List

    $reader = New-Object System.IO.StreamReader($ssl, $encoding, $false, 1024, $true)
    $writer = New-Object System.IO.StreamWriter($ssl, $encoding, 1024, $true)
    $writer.AutoFlush = $true

    $writer.WriteLine("EHLO mail.contoso.com")

    $line = $null
    do
    {
      $line = $reader.ReadLine()
      Write-Output $line
    } while (-not $line.StartsWith("250 "))
}
finally {
  if ($reader ) { $reader.Close() }
  if ($writer) { $writer.Close() }
  if ($ssl) { $ssl.Close() }
  if ($tcp) { $tcp.Close() }
}