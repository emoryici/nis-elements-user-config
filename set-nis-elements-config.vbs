' Emory Integrated Cellular Imaging
' 11/2018

' VB script to copy from default config folder: C:\ProgramData\Laboratory Imaging\Platform\default
' to current user folder C:\ProgramData\Laboratory Imaging\Platform\[USER]
'
Set wshShell = CreateObject("WScript.Shell")
strUsername = UCase( wshShell.ExpandEnvironmentStrings( "%USERNAME%" ) ) ' Nikon folders created in upper case

' StdOut print to console only works with CScript.exe
Dim StdOut : Set StdOut = CreateObject("Scripting.FileSystemObject").GetStandardStream(1)

' currently unused - for checking folder status - robocopy creates automatically if not present
'Set fso = CreateObject("Scripting.FileSystemObject")

msgboxStr = "Are you sure you want to set NIS-Elements layout to default? " & strUsername
intAnswer = _
    Msgbox(msgboxStr, _
        vbYesNo, "Set default layout")
If intAnswer = vbYes Then
    
    ' Chr(34) is " required for spaces in cmd line 
    'C:\ProgramData\Laboratory Imaging\Platform
    strDest = """C:/ProgramData/Laboratory Imaging/Platform/" & strUsername & "/"""
    strSource = """C:/ProgramData/Laboratory Imaging/Platform/default/"""
    strLogFile = """C:/ProgramData/Laboratory Imaging/Platform/" & strUsername & "_rc.txt"""
    ' /is for same files, and /it for different - both means all
    ' e for entire folder structure, /v verbose and /log file
    strParams = "/mir /e /v /log:" & strLogFile

    'If Not fso.FolderExists( strDest ) Then
    '   ...
    'End If
    
    cmd = "robocopy.exe " & strSource & " " & strDest & " " & strParams
    StdOut.Write "robocopy command: " & cmd & vbNewLine

    status = wshShell.Run(cmd, 0, True)
    ' see below for status codes returned by robocopy
    If status > 1 Then
        StdOut.Write "robocopy status: " & CStr(status) & vbNewLine
        StdOut.Write "copy incomplete - please examine log file: " & strLogFile
    Else
        StdOut.Write "robocopy status: " & CStr(status) & vbNewLine
        StdOut.Write "copy completed without issue"
    End If

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
