# MIT License
# 
# Copyright (c) 2018 Brad Hughes
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

# This script prints a list of IMAP folders in a mailbox

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    $Server,
    [Parameter(Mandatory=$true)]
    $Username,
    [Parameter(Mandatory=$true)]
    [SecureString]$Password,
    [Parameter(Mandatory=$false)]
    $Port = 0,
    [Parameter(Mandatory=$false)]
    $Ssl = $true,
    [Parameter(Mandatory=$false)]
    [switch]$HashFolderNames = $false
)

$ErrorActionPreference = "Stop"

$script:CommandCounter = 1

Function ConvertFrom-SecureString
{
    param(
        [securestring]$Secure
    )

    $bstr = $null
    try {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
        $unsecure = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)

        return $unsecure
    }
    finally {
        if ($bstr) {
            [System.Runtime.InteropServices.Marshal]::FreeBSTR($bstr)
        }
    }
}

Function Connect-ImapServer
{
    param(
        [Parameter(Mandatory=$true)]
        $Server,
        [Parameter(Mandatory=$false)]
        $Port = 0,
        [Parameter(Mandatory=$false)]
        $Ssl = $true
    )

    $tcpClient = New-Object System.Net.Sockets.TcpClient
    $tcpClient.Connect($Server, $Port)

    if ($Ssl) {
        $stream = New-Object System.Net.Security.SslStream($tcpClient.GetStream())
        $stream.AuthenticateAsClient($Server)
    }
    else {
        $stream = $tcpClient.GetStream()
    }

    return New-Object PSObject -Property @{
        Client = $tcpClient
        Stream = $stream
        Reader = New-Object System.IO.StreamReader($stream)
        Writer = New-Object System.IO.StreamWriter($stream)
    }
}

Function Send-Command
{
    param(
        [Parameter(Mandatory=$true)]
        [System.IO.StreamWriter]$Writer,

        [Parameter(Mandatory=$true)]
        [string]$CommandText,

        [Parameter(Mandatory=$true)]
        [int]$Counter
    )

    $Writer.WriteLine("A$Counter $CommandText")

    $Writer.Flush()
}

Function Receive-Response
{
    param(
        [Parameter(Mandatory=$true)]
        [System.IO.StreamReader]$Reader,

        [Parameter(Mandatory=$false)]
        [string]$DonePrefix = $Null
    )

    $resp = $null

    do
    {
        $resp = $Reader.ReadLine()
        Write-Output $resp
    }
    while ($DonePrefix -and -not $resp.StartsWith("$DonePrefix "))

    if ($resp -notmatch "^.+? OK")
    {
        throw "ERROR response was received from IMAP server: '$resp'"
    }
}

Function Execute-Command
{
    param(
        [Parameter(Mandatory=$true)]
        [System.IO.StreamWriter]$Writer,

        [Parameter(Mandatory=$true)]
        [System.IO.StreamReader]$Reader,

        [Parameter(Mandatory=$true)]
        [string]$CommandText
    )

    Send-Command -CommandText $CommandText -Writer $Writer -Counter:$script:CommandCounter
    $resp = Receive-Response -Reader $Reader -DonePrefix "A$($script:CommandCounter)"

    $script:CommandCounter++

    return $resp
}

#
# Main Script Body Start
#

# Handle the defaults values for the port
if ($Port -eq 0)
{
    $Port = if ($Ssl) { 993 } else { 143 }
}

$connection = $null

try
{
    $connection = Connect-ImapServer -Server $Server -Port $Port -Ssl $Ssl
    $resp = Receive-Response -Reader $connection.Reader

    $plainPassword = ConvertFrom-SecureString -Secure $Password

    $resp = Execute-Command -CommandText "LOGIN `"$Username`" `"$plainPassword`"" -Writer $connection.Writer -Reader $connection.Reader
    $resp = Execute-Command -CommandText "LIST `"`" `"*`"" -Writer $connection.Writer -Reader $connection.Reader

    if ($HashFolderNames)
    {
        $sha1 = new-Object System.Security.Cryptography.SHA1CryptoServiceProvider
        $exp = New-Object System.Text.RegularExpressions.Regex('^\* LIST\s+\(.+?\)\s+"."\s+"(?<path>.+)"$')
        $resp | Where-Object { $_[0] -eq '*' } | Foreach-Object {
            $match = $exp.Match($_);
            if ($match.Success) {

                $path = $match.Groups["path"].Value
                $pathBytes = [System.Text.Encoding]::UTF8.GetBytes($path.ToLowerInvariant())
                $hashBytes = $sha1.ComputeHash($pathBytes)
                $hashHex = New-Object System.Text.StringBuilder
                $hashBytes | ForEach-Object { $hashHex.AppendFormat("{0:X}", $_) | Out-Null }

                $hashRecord = New-Object PSObject -Property @{
                    FolderPath = $path
                    FolderHash = $hashHex.ToString()
                }

                Write-Output $hashRecord
            }
        }
    }
    else
    {
        Write-Output $resp
    }
}
finally
{
    if ($connection) {
        $connection.Stream.Close()
        $connection.Client.Close()
    }
}



