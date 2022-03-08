<#
    List-DLMembers.ps1
    
    This script lists all members of a distribution list including all 
    child groups members recursively. It create a CSV file with all
    groups of the given distribution group.

    Parameter: Distribution group Alias or Name or Email Address, can't be Dynamic 
	Distribution Group.	If want to pull all distro groups in the org, just run the 
	script without param.

    Example:
    .\List-DLMembers.ps1 -DLName "NA-Sales"
#>
param($DLname)


<#

Function    : Expand-Group
Parameter   : Distribution Group Name
Description : This function populates all the members of the
given distribution group to a global variable named $Global:Users 

#>
Function Expand-Group ($GroupName)
{
    $members = Get-DistributionGroupMember -Identity $GroupName

    foreach($member in $members)
    {
        $RecipientType = (Get-Recipient $member.Alias).RecipientType
        
        $member.Name + "`t`t`t" + $RecipientType
        if ($RecipientType -like "*DynamicDistributionGroup*")
        {
            # Found an dynamic child group - calling Expand-Dynamic-Group to expand the group
            Expand-Dynamic-Group -GroupName $member.Alias
        }
		elseif ($RecipientType -like "*DistributionGroup*")
		{
            # Found an child group - calling myself to expand the group
            Expand-Group -GroupName $member.Alias
		}
		else
		{
			# Create a PSCustomObject of the current member
			$MemberObject = [PSCustomObject] @{
				Name = "$($member.Name)" 
				Title = "$($member.Title)" 
				Department = "$($member.Department)" 
				Email = "$($member.PrimarySMTPAddress)" 
				Country = "$($member.CountryOrRegion)" 
				Memberof = "$GroupName"
			}

			# Store the member object to Users array
			$global:users += $MemberObject
		}
    }
}
<#
    End of the Function
#>


<#

Function    : Expand-Dynamic-Group
Parameter   : Distribution Group Name
Description : This function populates all the members of the
given dynamic distribution group to a global variable named $Global:Users 

#>
Function Expand-Dynamic-Group ($GroupName)
{
	$members = Get-Recipient -RecipientPreviewFilter (get-dynamicdistributiongroup $GroupName).RecipientFilter -OrganizationalUnit $GroupName.RecipientContainer

    foreach($member in $members)
    {
        $RecipientType = (Get-Recipient $member.Alias).RecipientType
        
        $member.Name + "`t`t`t" + $RecipientType

		# Create a PSCustomObject of the current member
		$MemberObject = [PSCustomObject] @{
			Name = "$($member.Name)" 
			Title = "$($member.Title)" 
			Department = "$($member.Department)" 
			Email = "$($member.PrimarySMTPAddress)" 
			Country = "$($member.CountryOrRegion)" 
			Memberof = "$GroupName"
		}

		# Store the member object to Users array
		$global:users += $MemberObject
    }
}
<#
    End of the Function
#>


<#
	THE SCRIPT STARTS HERE
#>

# Create a Global Array Variable to store all DL member objects
$global:users = @()

if([string]$DLname -ne "")
{
	# Create a Global Array Variable to store all DL member objects
	$global:users = @()

	# Call the Expand-Group function to populate the all DL members
	Expand-Group -GroupName $DLname

	#Store the member objects to a CSV file
	$filename = ".\$DLName-Members.csv"
	$global:users | Export-Csv $filename
}else{
	$confirmation = Read-Host "Populate all distribution groups... Are you Sure You Want To Proceed? (Y/N)"
	if ($confirmation -eq 'y') 
	{
		#Get all distribution groups
		$distrogroups = Get-DistributionGroup -ResultSize Unlimited

		Foreach($distrogroup in $distrogroups)
		{
			# Create a Global Array Variable to store all DL member objects
			$global:users = @()

			"Distro: " + $distrogroup.Name
		
			# Call the Expand-Group function to populate the all DL members
			Expand-Group -GroupName $distrogroup.Name

			#Store the member objects to a CSV file
			$filename = ".\$distrogroup-Members.csv"
			$global:users | Export-Csv $filename
		}
	}
}


<#
	END OF THE SCRIPT
#>