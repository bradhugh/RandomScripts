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
    $Password,
    [Parameter(Mandatory=$false)]
    $Port = 0,
    [Parameter(Mandatory=$false)]
    $Ssl = $true
)

$ErrorActionPreference = "Stop"

$script:CommandCounter = 1

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

    do
    {
        $resp = $Reader.ReadLine()
        Write-Output $resp
    }
    while ($DonePrefix -and -not $resp.StartsWith("$DonePrefix "))
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

    $resp = Execute-Command -CommandText "LOGIN `"$($Username)`" `"$($Password)`"" -Writer $connection.Writer -Reader $connection.Reader
    $resp = Execute-Command -CommandText "LIST `"`" `"*`"" -Writer $connection.Writer -Reader $connection.Reader
    Write-Output $resp
}
finally
{
    if ($connection) {
        $connection.Stream.Close()
        $connection.Client.Close()
    }
}


