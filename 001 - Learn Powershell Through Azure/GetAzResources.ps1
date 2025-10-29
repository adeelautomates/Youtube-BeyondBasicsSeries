Connect-AzAccount # Connect to Any Subscription when AuthN as yourself. Be sure you have global reader for this to work.

$subs = (Get-AzSubscription).name # collect all subs
$dataCollected = New-Object System.Collections.Generic.List[pscustomobject]
foreach($sub in $subs){ # Loop through each sub
    set-azcontext $sub | Out-Null
    $resources = get-azresource | select-object Name, ResourceGroupName, Location, Tags # collect all of its resources
    foreach($resource in $resources){ # Loop through each resource
        $dataCollected.Add([pscustomobject]@{ # store data in list as it loops
            Subscription = $sub
            ResourceGroup = $resource.ResourceGroupName
            Name = $resource.Name
            Region = $resource.Location
            Tags = (($resource.tags.GetEnumerator() | ForEach-Object {"'$($_.Key) : $($_.Value)'"}) -join ', ')
        })
    }
}
$dataCollected | Export-Csv "C:\test\Resources.csv" # set the folder you want to export to and provide a fileName