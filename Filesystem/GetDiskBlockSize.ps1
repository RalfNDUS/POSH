# Get disk block size and starting offset

Get-WmiObject Win32_DiskPartition  | select-object `
	@{Name="Name"; Expression={$_.name}},
	@{Name="Index"; Expression={$_.index}},
	@{Name="Size"; Expression={[long]($_.size/1mb)}},
	@{Name="BlockSize"; Expression={$_.BlockSize}},
	@{Name="StartingOffset"; Expression={$_.StartingOffset}} | `
	sort Name  | ft -Autosize
	