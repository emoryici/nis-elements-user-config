' VB script to copy from default config folder: C:\ProgramData\Laboratory Imaging\Platform\default
' to current user folder C:\ProgramData\Laboratory Imaging\Platform\[USER]
'
Set wshShell = CreateObject("WScript.Shell")
strName = UCase( wshShell.ExpandEnvironmentStrings( "%USERNAME%" ) )

Dim StdOut : Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(1)

Set fso = CreateObject("Scripting.FileSystemObject")


msgboxStr = "Are you sure you want to set NIS-Elements layout to default? " & strName
intAnswer = _
    Msgbox(msgboxStr, _
        vbYesNo, "Set default layout")
If intAnswer = vbYes Then
    
    strDest = "C:\ProgramData\" & Chr(34) & "Laboratory Imaging" & Chr(34) &  "\Platform\" & strName
    strSource = "C:\ProgramData\" & Chr(34) & "Laboratory Imaging" & Chr(34) &  "\Platform\default"

    strLogDest = "C:\ProgramData\" & Chr(34) & "Laboratory Imaging" & Chr(34) &  "\Platform\"
    strParams = "/e /log:" & strLogDest ' e for entire folder structure, and log file

    tmpDest = "C:\tmp\temp2"
    tmpSource = "C:\tmp\temp1"
    'If Not fso.FolderExists( strRoot & strName ) Then
    '   
    cmd = "robocopy.exe " & tmpSource & " " & tmpDest & " /e /log:C:\tmp\robocopy.log"
    status = wshShell.Run(cmd, 0, True)
    StdOut.Write "robocopy status: " & CStr(status)

End If




' robocopy status codes
' Code	Meaning
' 0	No errors occurred and no files were copied.
' 1	One of more files were copied successfully.
' 2	Extra files or directories were detected.  Examine the log file for more information.
' 4	Mismatched files or directories were detected.  Examine the log file for more information.
' 8	Some files or directories could not be copied and the retry limit was exceeded.
' 16	Robocopy did not copy any files.  Check the command line parameters and verify that Robocopy has enough rights to write to the destination folder.

' example directory
'C:\ProgramData\Laboratory Imaging\Platform\JOHNSM3