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

# This Script exports message trace results to a CSV for a set of recipients

param(
    [Parameter(Mandatory=$false)] $Username = "tenantadmin@contoso.onmicrosoft.com",
    [Parameter(Mandatory=$false)] $PasswordFile = "C:\temp\password.txt",
    
    [Parameter(Mandatory=$false)] $OutputFile = ("C:\temp\MessageTrace_{0:yyyy-MM-dd}.csv" -f [DateTime]::Now),
    
    [Parameter(Mandatory=$false)] $StartDate = ([DateTime]::Today.AddDays(-3)),
    [Parameter(Mandatory=$false)] $EndDate = ([DateTime]::Today.AddDays(1)),
    
    # Add recipients here (comma seperated)
    [Parameter(Mandatory=$false)] $Recipients = @(
        "recip1@contoso.onmicrosoft.com",
        "recip2@contoso.onmicrosoft.com"),

    [Parameter(Mandatory=$false)] $PageSize = 1000
)

$ErrorActionPreference = "Stop"

Function Connect-O365PowerShell
{
    $o365PowerShellUrl = "https://outlook.office365.com/PowerShell-LiveID"
    $pass = ConvertTo-SecureString -AsPlainText -Force (Get-Content $PasswordFile)
    $credential = New-Object PSCredential($Username, $pass)
    $session = New-PSSession -ConnectionUri $o365PowerShellUrl `
                    -ConfigurationName Microsoft.Exchange `
                    -Credential $credential `
                    -Authentication Basic

    Import-PSSession $session -AllowClobber
    return $session
}

try
{
    $null = Connect-O365PowerShell

    $fi = New-Object System.IO.FileInfo($OutputFile)
    $processingFileName = "$($fi.DirectoryName)\$($fi.BaseName)_Processing.csv"
    $messages = $null
    $page = 1

    do
    {
        # Run the Get-MessageTrace cmdlet to get the next page of results
        $messages = Get-MessageTrace `
                        -RecipientAddress $recipients `
                        -Page $page -PageSize $PageSize `
                        -StartDate $StartDate `
                        -EndDate $EndDate

        $messages | Export-Csv -NoTypeInformation -Path $processingFileName -Append

        $page++
    }
    while ($messages)

    # flip the processing file to the final output
    Rename-Item $processingFileName $OutputFile 
}
Finally
{
    # TODO: This session handling could use improvement
    Get-PSSession | Remove-PSSession
}

