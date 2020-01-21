strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2" )
Set colDisks = objWMIService.ExecQuery _
    ("Select * from Win32_LogicalDisk Where DriveType = 3" )
For Each objDisk in colDisks
    intFreeSpace = objDisk.FreeSpace
    intTotalSpace = objDisk.Size
	intFreeSpacekB = intFreeSpace/1024
	intFreeSpaceMB = intFreeSpacekB/1024
	intFreeSpaceGB = intFreeSpaceMB/1024
	intTotalSpacekB = intTotalSpace/1024
	intTotalSpaceMB = intTotalSpacekB/1024
	intTotalSpaceGB = intTotalSpaceMB/1024
    pctFreeSpace = intFreeSpace / intTotalSpace
    Echo objDisk.DeviceID & "\ " & FormatNumber(intFreeSpaceGB,"2") & " / " & FormatNumber(intTotalSpaceGB,"2") & " GB " & objDisk.FileSystem & " (" & FormatPercent(pctFreeSpace) & ")"
Next 