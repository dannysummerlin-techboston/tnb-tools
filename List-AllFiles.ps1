<#
 .Synopsis
  List all files in arbitrarily large SharePoint document libraries

 .Description
  Export a CSV file detailing the name, file type, file size, relative URL, creation, and modification details.

 .Parameter username
  Username for SharePoint access

 .Parameter password
  Password for SharePoint access

 .Parameter siteUrl
  The URL for the SharePoint site you'd like to list out

 .Parameter listName
  The name of the list to use for exporting file details, defaults to Documents

 .Parameter outputFile
  A file name for the final data to be gathered into
#>
Function List-AllFiles {
	[CmdletBinding()]
	Param(
		[string]$username,
		[string]$password,
		[string]$siteUrl,
		[string]$listName = "Documents",
		[string]$outputFile = "SharePoint Files List.csv",
		[integer]$pageSize = 4000
	)

	if (!(Get-Module -ListAvailable -Name "Pnp.PowerShell")) {
		Install-Module -name Pnp.PowerShell -Force -AcceptLicense
	}
	Import-Module Pnp.PowerShell

	#Connect to SharePoint Online site
	[pscredential]$credentials = New-Object System.Management.Automation.PSCredential ($username, (ConvertTo-SecureString $password -AsPlainText -Force))
	Connect-PnPOnline $siteUrl -Credentials $credentials
	$global:totalCounter = 0
	$global:fileCounter = 0

	#Get all Documents from the document library
	$List  = Get-PnPList -Identity $listName
	Write-host "Connected, beginning"
	Get-PnPListItem -List $listName -PageSize $pageSize -Fields Author, Editor, Created, File_x0020_Type, File_x0020_Size -ScriptBlock  {
		Param($rawItems)

		$Results = New-Object System.Collections.ArrayList
		$items = $rawItems | Where {$_.FileSystemObjectType -eq "File"}
		Write-Progress -PercentComplete (($global:totalCounter / $List.ItemCount) * 100) -Activity "Getting Documents from Library" -CurrentOperation "Getting partial List"
		$global:totalCounter += $items.Count
		$itemCounter = 0
		Foreach ($Item in $items) {
			Write-Progress -PercentComplete (($itemCounter / $items.Count) * 100) -Activity "Getting Documents from Library" -CurrentOperation "Processing list items"
			$o = New-Object PSObject -Property ([ordered]@{
				Name              = $Item["FileLeafRef"]
				FileType          = $Item["File_x0020_Type"]
				FileSize          = $Item["File_x0020_Size"]
				RelativeURL       = $Item["FileRef"]
				CreatedByEmail    = $Item["Author"].LookupValue
				CreatedOn         = $Item["Created"]
				Modified          = $Item["Modified"]
				ModifiedByEmail   = $Item["Editor"].LookupValue
			})
			$Results.Add($o) > $null
			$itemCounter++ > $null
		}
		#Export the results to temporary CSV
		$global:fileCounter++ > $null
		$Results | Export-Csv -Path "List-$($global:fileCounter).csv" -NoTypeInformation
		# may need to rate limit
		Start-Sleep -Seconds 2
	} > $null
	# combine csv files into final output
	$getFirstLine = $true
	Get-ChildItem "List-*.csv" | % {
		try {
			$lines = Get-Content $_
			if($getFirstLine) { $linesToWrite = $lines }
			else { $linesToWrite = ($lines | Select -Skip 1) }
			$getFirstLine = $false
			Add-Content $outputFile $linesToWrite
			Remove-Item $_
		} catch {
			Write-Error $_
		}
	}
	Write-Information "Document Library Inventory Exported to CSVs Successfully!"
}
