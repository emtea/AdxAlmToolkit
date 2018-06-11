function ConvertTo-PSObject
{
	[CmdletBinding()]
	param (
		[parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		[hashtable]
		$property
	)

	process
	{
		return New-Object PSObject -Property $property
	}
}

function ConvertTo-Hashtable
{
	[CmdletBinding(DefaultParameterSetName="ByCollection")]
	param (
		[parameter(Position=0, Mandatory=$true, ParameterSetName="ByCollection")]
		$collection,

		[parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName="ByPipeline")]
		$item,

		$keyProperty,

		$valueProperty
	)
	
	begin {
		$table = @{}
	}

	process {
		if ($collection) {
			$collection | ConvertTo-Hashtable -keyProperty $keyProperty -valueProperty $valueProperty
		} else {
			if ($keyProperty -is [scriptblock]) {
				$key = & $keyProperty $item
			} elseif ($keyProperty) {
				$key = $item.$keyProperty
			} else {
				$key = $item
			}

			if ($valueProperty -is [scriptblock]) {
				$value = & $valueProperty $item
			} elseif ($valueProperty) {
				$value = $item.$valueProperty
			} else {
				$value = $item
			}

			$table.Add($key, $value)
		}
	}

	end {
		Write-Output $table
	}
}

function ConvertTo-IndexedPair
{
	[CmdletBinding(DefaultParameterSetName="ByCollection")]
	param (
		[parameter(Position=0, Mandatory=$true, ParameterSetName="ByCollection")]
		$collection,

		[parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName="ByPipeline")]
		$item
	)
	
	begin {
		$index = 0;
	}

	process {
		if ($collection) {
			$collection | ConvertTo-IndexedPair
		} else {
			@{ Key = $index++; Value = $item } | ConvertTo-PSObject
		}
	}
}

function Invoke-Using
{
	[CmdletBinding()]
	param (
		[parameter(Position=0, Mandatory=$true)]
		[System.IDisposable]
		$obj,

		[parameter(Position=1, Mandatory=$true)]
		[scriptblock]
		$action
	)

	process
	{
		try
		{
			& $action
		}
		finally
		{
			if ($obj -is [System.IDisposable])
			{
				$obj.Dispose()
			}
		}
	}
}

function Invoke-Timed
{
	[CmdletBinding()]
	param (
		[parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		[scriptblock]
		$action,

		[scriptblock]
		$onStop = { param([System.TimeSpan]$elapsed) "Elapsed: {0}" -f $elapsed }
	)

	begin
	{
		$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
	}

	process
	{
		& $action
	}

	end
	{
		$stopwatch.Stop()
		& $onStop $stopwatch.Elapsed
	}
}

function Test-Any()
{
	begin
	{
		$any = $false
	}

	process
	{
		$any = $true
	}

	end
	{
		Write-Output $any
	}
}

function Invoke-WithProgressBar
{
	[CmdletBinding(DefaultParameterSetName="ByCollection")]
	param (
		[parameter(Position=0, Mandatory=$true, ParameterSetName="ByCollection")]
		$collection,

		[parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName="ByPipeline")]
		[alias("key", "k", "i", "n")]
		$index,

		[parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName="ByPipeline")]
		[alias("v")]
		$value,

		[parameter(Mandatory=$true, ParameterSetName="ByPipeline")]
		[int]
		$length,

		$id = 0,

		$activity = "Processing...",

		[scriptblock]
		$onStatus = { param($v, $i, $p) "[{0:0}%] {1} of {2}" -f $p, ($i + 1), $length },

		[scriptblock]
		$onMessage = { param($v, $i, $p) $v }
	)

	process {
		if ($collection) {
			$collection | ConvertTo-IndexedPair | Invoke-WithProgressBar -length $collection.Length -id $id -activity $activity -onStatus $onStatus -onMessage $onMessage
		} else {
			if ($length -gt 0) {
				$percentage = ($index / $length) * 100
				$status = (& $onStatus -v $value -i $index -p $percentage)
				$message = (& $onMessage -v $value -i $index -p $percentage)
				Write-Progress -Id $id -Activity $activity -percentComplete $percentage -Status $status -CurrentOperation $message
			}
			Write-Output $value
		}
	}

	end {
		if ($length -gt 0) {
			Write-Progress -Id $id -Activity $activity -Completed
		}
	}
}

Export-ModuleMember -function ConvertTo-PSObject
Export-ModuleMember -function ConvertTo-Hashtable
Export-ModuleMember -function ConvertTo-IndexedPair
Export-ModuleMember -function Invoke-Using
Export-ModuleMember -function Invoke-Timed
Export-ModuleMember -function Test-Any
Export-ModuleMember -function Invoke-WithProgressBar
