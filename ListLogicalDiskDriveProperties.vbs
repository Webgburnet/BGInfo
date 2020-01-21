' List Logical Disk Drive Properties


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDisks = objWMIService.ExecQuery _
("Select * from Win32_LogicalDisk")

For each objDisk in colDisks
Echo "Compressed: " & objDisk.Compressed
Echo "Description: " & objDisk.Description
Echo "DeviceID: " & objDisk.DeviceID
Echo "DriveType: " & objDisk.DriveType
Echo "FileSystem: " & objDisk.FileSystem
Echo "FreeSpace: " & objDisk.FreeSpace
Echo "MediaType: " & objDisk.MediaType
Echo "Name: " & objDisk.Name
Echo "QuotasDisabled: " & objDisk.QuotasDisabled
Echo "QuotasIncomplete: " & objDisk.QuotasIncomplete
Echo "QuotasRebuilding: " & objDisk.QuotasRebuilding
Echo "Size: " & objDisk.Size
Echo "SupportsDiskQuotas: " & objDisk.SupportsDiskQuotas
Echo "SupportsFileBasedCompression: " & objDisk.SupportsFileBasedCompression
Echo "SystemName: " & objDisk.SystemName
Echo "VolumeDirty: " & objDisk.VolumeDirty
Echo "VolumeName: " & objDisk.VolumeName
Echo "VolumeSerialNumber: " & objDisk.VolumeSerialNumber
Next