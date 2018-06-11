Function Resolve-VsComnToolsPath
{
	[CmdletBinding()]
	Param (
		$vsComnToolsPath = ($env:VS140COMNTOOLS, $env:VS120COMNTOOLS, $env:VS110COMNTOOLS, $env:VS100COMNTOOLS, $env:VS90COMNTOOLS)
	)

	$latestVsComnToolsPath = $vsComnToolsPath | Where-Object { $_ } | Select-Object -First 1

	if (-not $latestVsComnToolsPath)
	{
		Throw New-Object System.InvalidOperationException -ArgumentList "Visual Studio is not installed."
	}

	Write-Output (Resolve-Path $latestVsComnToolsPath)
}

Function Resolve-TfsPath
{
	<#
	.SYNOPSIS
	Returns the path to tf.exe.
	#>

	$latestVsComnToolsPath = Resolve-VsComnToolsPath
	$tfPath = Join-Path $latestVsComnToolsPath "..\IDE\tf.exe" -Resolve

	Write-Output $tfPath
}

Function ToSwitch
{
	<#
	.SYNOPSIS
	Returns the name of a switch variable if it is enabled.
	#>

	Param ([System.Management.Automation.PSVariable[]]$switchVar)

	$switchVar | ? { $_.Value } | % { $_.Name }
}

Function Invoke-TfsCommand
{
	<#
	.SYNOPSIS
	Invokes tf.exe.
	.DESCRIPTION
	http://msdn.microsoft.com/en-us/library/z51z7zy0.aspx
	.EXAMPLE
	PS C:\> Invoke-TfsCommand checkin "C:\file.html" ("recursive", "noprompt") @{comment = "My checkin comment."}
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		$command,

		[Parameter(Position=1)]
		$parameter,

		[Parameter(Position=2)]
		$switch,

		[Parameter(Position=3)]
		[hashtable]
		$option,

		[Parameter()]
		[PSCredential]
		$credential,

		[Parameter()]
		$tfPath = (Resolve-TfsPath)
	)

	Function Get-TfsCommandArgs
	{
		Param ($masked)

		Write-Output $command

		if ($parameter) { $parameter | % { Write-Output $_ } }
		if ($switch) { $switch | % { Write-Output ("/{0}" -f $_) } }
		if ($option) { $option.GetEnumerator() | % { Write-Output ("/{0}:{1}" -f $_.Name, $_.Value) } }

		if ($credential)
		{
			if ($masked) { $password = "****" } else { $password = $credential.GetNetworkCredential().Password}
			Write-Output ('/login:"{0},{1}"' -f $credential.UserName, $password)
		}
	}

	$line = (Get-TfsCommandArgs -Masked $true) -join " "
	Write-Verbose $line

	return & $tfPath (Get-TfsCommandArgs) 2>&1
}

Function Invoke-TfsHelp
{
	<#
	.SYNOPSIS
	Invokes tf.exe help operation.
	.DESCRIPTION
	http://msdn.microsoft.com/en-us/library/dhaa6tz1.aspx
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		$commandname
	)
		
	return Invoke-TfsCommand help $commandname
}

Function Invoke-TfsDir
{
	<#
	.SYNOPSIS
	Invokes tf.exe dir operation.
	.DESCRIPTION
	http://msdn.microsoft.com/en-us/library/6320xzye.aspx
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		$itemspec,

		[Parameter()]
		$credential,

		[Parameter()]
		[Switch]
		$recursive
	)
		
	return Invoke-TfsCommand dir $itemspec (ToSwitch (gv recursive)) -Credential $credential
}

Function Invoke-TfsGetLatest
{
	<#
	.SYNOPSIS
	Invokes tf.exe get operation.
	.DESCRIPTION
	http://msdn.microsoft.com/en-us/library/fx7sdeyf.aspx
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		$itemspec,

		[Parameter()]
		$credential,

		[Parameter()]
		[Switch]
		$force,

		[Parameter()]
		[Switch]
		$recursive = $true,

		[Parameter()]
		[Switch]
		$noprompt = $true
	)
	
	return Invoke-TfsCommand get $itemspec (ToSwitch (gv force), (gv recursive), (gv noprompt)) -Credential $credential
}

Function Invoke-TfsCheckout
{
	<#
	.SYNOPSIS
	Invokes tf.exe checkout operation.
	.DESCRIPTION
	http://msdn.microsoft.com/en-us/library/1yft8zkw.aspx
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		$itemspec,

		[Parameter()]
		$credential,

		[Parameter()]
		[Switch]
		$recursive,

		[Parameter()]
		[Switch]
		$noprompt = $true
	)
	
	return Invoke-TfsCommand checkout $itemspec (ToSwitch (gv recursive), (gv noprompt)) -Credential $credential
}

Function Invoke-TfsCheckin
{
	<#
	.SYNOPSIS
	Invokes tf.exe checkin operation.
	.DESCRIPTION
	http://msdn.microsoft.com/en-us/library/c327ca1z.aspx
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		$itemspec,

		[Parameter()]
		$credential,

		[Parameter(Mandatory=$true)]
		$comment,

		[Parameter()]
		[Switch]
		$recursive,

		[Parameter()]
		[Switch]
		$noprompt = $true
	)

	return Invoke-TfsCommand checkin $itemspec (ToSwitch (gv recursive), (gv noprompt)) @{comment = $comment} -Credential $credential
}

Function Invoke-TfsUndo
{
	<#
	.SYNOPSIS
	Invokes tf.exe undo operation.
	.DESCRIPTION
	http://msdn.microsoft.com/en-us/library/c72skhw4.aspx
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		$itemspec,

		[Parameter()]
		$credential,

		[Parameter()]
		[Switch]
		$recursive,

		[Parameter()]
		[Switch]
		$noprompt = $true
	)
	
	return Invoke-TfsCommand undo $itemspec (ToSwitch (gv recursive), (gv noprompt)) -Credential $credential
}

Function Invoke-TfsDelete
{
	<#
	.SYNOPSIS
	Invokes tf.delete operation.
	.DESCRIPTION
	http://msdn.microsoft.com/en-us/library/k45zb450.aspx
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		$itemspec,

		[Parameter()]
		$credential,

		[Parameter()]
		[Switch]
		$recursive,

		[Parameter()]
		[Switch]
		$noprompt = $true
	)
	
	return Invoke-TfsCommand delete $itemspec (ToSwitch (gv recursive), (gv noprompt)) -Credential $credential
}

Function Invoke-TfsAdd
{
	<#
	.SYNOPSIS
	Invokes tf.exe add operation.
	.DESCRIPTION
	http://msdn.microsoft.com/en-us/library/f9yw4ea0.aspx
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		$itemspec,

		[Parameter()]
		$credential,

		[Parameter()]
		[Switch]
		$recursive,

		[Parameter()]
		[Switch]
		$noprompt = $true
	)
	
	return Invoke-TfsCommand add $itemspec (ToSwitch (gv recursive), (gv noprompt)) -Credential $credential
}

Function Invoke-TfsLabel
{
	<#
	.SYNOPSIS
	Invokes tf.exe label operation.
	.DESCRIPTION
	http://msdn.microsoft.com/en-us/library/9ew32kd1.aspx
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		$labelname,

		[Parameter(Mandatory=$true,Position=1)]
		$itemspec,

		[Parameter()]
		$credential,

		[Parameter()]
		$comment = "Automated label from $env:COMPUTERNAME",

		[Parameter()]
		[Switch]
		$recursive
	)
	
	return Invoke-TfsCommand label ($labelname, $itemspec) (ToSwitch (gv recursive)) @{comment = $comment} -Credential $credential
}

Function Invoke-TfsHistory
{
	<#
	.SYNOPSIS
	Invokes tf.exe history operation.
	.DESCRIPTION
	https://msdn.microsoft.com/en-us/library/yxtbh4yh.aspx
	#>

	[CmdletBinding(DefaultParameterSetName="ByTake")]
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		$itemspec,

		[parameter(mandatory=$true, position=1)]
		$user,

		[parameter(mandatory=$true, ParameterSetName="ByTake")]
		$take,

		[parameter(mandatory=$true, ParameterSetName="ByChangeset")]
		$changeSet,

		[parameter(mandatory=$true, position=2, ParameterSetName="ByDate")]
		$startDate,
		
		[parameter(mandatory=$true, position=3, ParameterSetName="ByDate")]
		$endDate,

		$format = "Detailed",

		[Parameter()]
		$credential,

		[Parameter()]
		[Switch]
		$recursive = $true
	)
	
	$option = @{user = $user; format = $format}

	if ($take) {
		$option["stopafter"] = $take
	} elseif ($changeSet) {
		$option["version"] = "C$changeSet~C$changeSet"
	} elseif ($startDate -and $endDate) {
		$option["version"] = "D$startDate~D$endDate"
	}

	return Invoke-TfsCommand history $itemspec (ToSwitch (gv recursive)) $option -Credential $credential
}

Function Invoke-TfsMerge
{
	<#
	.SYNOPSIS
	Invokes tf.exe merge operation.
	.DESCRIPTION
	https://msdn.microsoft.com/en-us/library/bd6dxhfy.aspx
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		$source,

		[parameter(mandatory=$true, position=1)]
		$destination,

		[parameter(mandatory=$true, position=2)]
		$changeSet,

		$format = "Detailed",

		[Parameter()]
		$credential,

		[Parameter()]
		[Switch]
		$recursive = $true,

		[Parameter()]
		[Switch]
		$discard,

		[Parameter()]
		[Switch]
		$preview,

		[Parameter()]
		[Switch]
		$force
	)
	
	$option = @{version = "C$changeSet~C$changeSet"; format = $format}

	return Invoke-TfsCommand merge ($source, $destination) (ToSwitch (gv recursive), (gv discard), (gv preview), (gv force)) $option -Credential $credential
}

Function Invoke-TfsMergeCandidate
{
	<#
	.SYNOPSIS
	Invokes tf.exe merge operation with candidate switch.
	.DESCRIPTION
	https://msdn.microsoft.com/en-us/library/bd6dxhfy.aspx
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true, Position=0)]
		$source,

		[parameter(mandatory=$true, position=1)]
		$destination,

		$format = "Detailed",

		[Parameter()]
		$credential,

		[Parameter()]
		[Switch]
		$recursive = $true
	)
	
	$candidate = $true
	$option = @{format = $format}

	return Invoke-TfsCommand merge ($source, $destination) (ToSwitch (gv recursive), (gv candidate)) $option -Credential $credential
}

Function Test-TfsPath
{
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		$itemspec,

		[Parameter()]
		$credential,

		[Parameter()]
		[Switch]
		$recursive
	)

	$result = Invoke-TfsDir -itemspec $itemspec -credential $credential -recursive:$recursive

	return $LastExitCode -eq 0
}

Function Invoke-MsBuild
{
	<#
	.SYNOPSIS
	Invokes MSBuild.exe.
	#>

	[CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true,Position=0)]
		$solutionPath,

		[Parameter(Position=1)]
		$target = "Rebuild",

		[Parameter(Position=2)]
		$configuration = "Debug",

		[Parameter(Position=3)]
		$msBuildPath = (
			( Join-Path ${env:ProgramFiles(x86)} "MSBuild\12.0\bin\amd64\MSBuild.exe" ),
			( Join-Path $env:windir "Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" )
		)
	)

	Write-Host "Build Solution: $solutionPath"
	
	$path = $msBuildPath | ? { Test-Path $_ } | select -f 1
	& $path $solutionPath "/target:$target" "/property:Configuration=$configuration"

	if ($LastExitCode -ne 0) { throw "MSBuild.exe error." }
}

Export-ModuleMember -function Resolve-TfsPath
Export-ModuleMember -function Invoke-TfsCommand
Export-ModuleMember -function Invoke-TfsHelp
Export-ModuleMember -function Invoke-TfsDir
Export-ModuleMember -function Invoke-TfsGetLatest
Export-ModuleMember -function Invoke-TfsCheckout
Export-ModuleMember -function Invoke-TfsCheckin
Export-ModuleMember -function Invoke-TfsUndo
Export-ModuleMember -function Invoke-TfsAdd
Export-ModuleMember -function Invoke-TfsDelete
Export-ModuleMember -function Invoke-TfsLabel
Export-ModuleMember -function Invoke-TfsHistory
Export-ModuleMember -function Invoke-TfsMerge
Export-ModuleMember -function Invoke-TfsMergeCandidate
Export-ModuleMember -function Test-TfsPath
Export-ModuleMember -function Invoke-MsBuild
