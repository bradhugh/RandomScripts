[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]
    $Server,

    [Parameter(Mandatory=$true)]
    [string]
    $SmtpStream,

    [Parameter(Mandatory=$false)]
    [int]
    $Port = 25
)

$tcpClient = $null
$stream = $null
$writer = $null

try {
    $tcpClient = New-Object System.Net.Sockets.TcpClient
    $tcpClient.Connect($Server, $Port)

    $stream = $tcpClient.GetStream()

    $writer = New-Object System.IO.StreamWriter($stream)
    $writer.Write($SmtpStream)
    $writer.Flush()
}
finally {
    if ($writer) {
        $writer.Close()
    }

    if ($stream) {
        $stream.Close()
    }

    if ($tcpClient) {
        $tcpClient.Close()
    }
}
