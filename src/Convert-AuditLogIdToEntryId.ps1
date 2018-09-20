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

# This script converts ID values found in Unified Audit logs to EntryId values

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$AuditLogId,

    [Parameter(Mandatory=$false)]
    [Switch]$SkipJsonEscape = $false
)

# The string could be JSON escaped, let's decode it
if (-not $SkipJsonEscape)
{
    # First make sure we remove any quotes
    $AuditLogId = $AuditLogId.Replace("`"", "")
    $AuditLogId = ConvertFrom-Json "`"$AuditLogId`""
}

$bytes = [Convert]::FromBase64String($AuditLogId)

# Strip off the first and last bytes. These are not part of the EntryId
$bytes = $bytes[1..($bytes.Length - 2)]

$sbHex = New-Object System.Text.StringBuilder($bytes.Length * 2)
foreach ($byte in $bytes)
{
    $null = $sbHex.AppendFormat("{0:X2}", $byte)
}

return $sbHex.ToString()