<#
 .SYNOPSIS
    This script is used for removing and re-adding duplicate Mailbox Permissions.

 .DESCRIPTION
    This script can be used in the case that a single user is added to multiple ACEs for a mailbox when viewed with Get-MailboxPermission

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
    OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER
 
 .EXAMPLE
    Repair-DuplicateMailboxPermission.ps1 -SharedMailbox shared@contoso.com

    This will remove and re-add permission for any user that shows up in the list twice.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
param (
  [Parameter(Mandatory = $true)]
  [ValidateNotNull()]
  [string]
  $SharedMailbox
)

$ErrorActionPreference = "Stop"

Function Select-AccessRights([array]$AccessRights) {
  return $AccessRights | ForEach-Object {
    $_.ToString().Split(", ".ToCharArray(), [StringSplitOptions]::RemoveEmptyEntries) 
  } | Sort-Object -Unique
}

Function Select-InheritanceType([array]$InheritanceType) {
  [string]$effective = $null
  foreach ($current in $InheritanceType) {
    if ([string]::IsNullOrEmpty($effective)) {
      $effective = $current
    } else {
      switch ($effective) {
        "All" {} # Do nothing as you can't get a higher value than All
        "SelfAndChildren" {
          switch ($current) {
            "All" { $effective = "All" } # All is highest
            "Children" {} # SelfAndChildren already includes Children
            "Descendents" { $effective = "All" } # If we add Descendents to SelfAndChildren, this make All
            "None" {} # SelfAndChildren already includes None/Self
            "SelfAndChildren" {} # SelfAndChildren is already the effective
            default { throw "Unexpected InheritanceType $current" }
          }
        }
        "Descendents" {
          switch ($current) {
            "All" { $effective = "All" } # All is highest
            "Children" {} # Descendents already includes Children
            "Descendents" {} # Descendents is already the effective
            "None" { $effective = "All" } # If we add None/Self to Descendents, this makes all
            "SelfAndChildren" { $effective = "All" } # If we add SelfAndChildren to Descendents, this make All
            default { throw "Unexpected InheritanceType $current" }
          }
        }
        "Children" {
          switch ($current) {
            "All" { $effective = "All" } # All is highest
            "Children" {} # Children is already the effective
            "Descendents" { $effective = "Descendents" } # Descendents is inclusive of Children
            "None" { $effective = "SelfAndChildren" } # If we add None/Self to Children, this makes SelfAndChildren
            "SelfAndChildren" { $effective = "SelfAndChildren" } # Children includes Children
            default { throw "Unexpected InheritanceType $current" }
          }
        }
        "None" {
          switch ($current) {
            "All" { $effective = "All"; } # All is highest
            "Children" { $effective = "SelfAndChildren" } # None/Self plus children is SelfAndChildren
            "Descendents" { $effective = "All" } # None/Self plus Descendents is All
            "None" { } # None/Self is already the effective
            "SelfAndChildren" { } # SelfAndChildren already included None/Self
            default { throw "Unexpected InheritanceType $current" }
          }
        }
        default { throw "Unexpected InheritanceType $effective" }
      }
    }
  }

  return $effective
}

### MAIN SCRIPT START ###

# Get the permissions
$acl = Get-MailboxPermission -Identity $SharedMailbox `
  # We cant do anything about inherited permissions
  | Where-Object { $_.IsInherited -eq $false } `
  # We are not touching deny permissions
  | Where-Object { $_.Deny -eq $false } `
  # Any permissions that are unresolved SIDs we wouldn't be able to re-add
  | Where-Object { $_.User -notlike "S-1-*" }

# Group them
$grouped = $acl `
  | Select-Object -Property *,@{ Name="UserName"; Expression={ $_.User.ToString() } } `
  | Group-Object "UserName"

# Find the user names that need repairing
$toRepair = $grouped | Where-Object {
  $_.Count -gt 1 -and $_.Name -notlike "NT AUTHORITY\*"
}

# Do the remove and add
foreach ($entry in $toRepair) {
  $user = Get-User -Identity $entry.Name
  if (-not $user) {
    throw "Could not find user '$($entry.Name)'"
  }

  if ($user.Count -gt 1) {
    throw "Found more than one user matching '$user'"
  }

  if ($PSCmdlet.ShouldProcess("Removing and re-adding permissions for User $($entry.Name) to $SharedMailbox")) {

    # Remove all of the existing ACEs for this user
    foreach ($ace in $entry.Group) {
      Remove-MailboxPermission -Identity $SharedMailbox -User $user.Sid -AccessRights $ace.AccessRights -InheritanceType $ace.InheritanceType -Confirm:$false -ErrorAction:SilentlyContinue
    }

    # Add the new single ACE
    $rights = Select-AccessRights $entry.Group.AccessRights
    $inheritanceType = Select-InheritanceType $entry.Group.InheritanceType
    $rightsAdded = Add-MailboxPermission -Identity $SharedMailbox -User $user.Sid -AccessRights $rights -InheritanceType $inheritanceType

    Write-Output $rightsAdded
  }
}