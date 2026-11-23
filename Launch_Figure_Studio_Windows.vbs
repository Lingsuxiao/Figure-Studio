Option Explicit

' Windows silent double-click launcher for MATLAB Figure Layout Tool
' Pure ASCII contents to avoid GBK/UTF-8 encoding issues in Windows Script Host.

Dim fso, shell, appDir, matlabExe, matlabCmd, cmdLine

Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

appDir = fso.GetParentFolderName(WScript.ScriptFullName)

If Not fso.FileExists(fso.BuildPath(appDir, "FigureLayoutApp.m")) Then
    MsgBox "Cannot find FigureLayoutApp.m in:" & vbCrLf & appDir, vbCritical, "Figure Layout Tool"
    WScript.Quit 1
End If

matlabExe = FindMatlabExe(fso, shell)
If matlabExe = "" Then
    MsgBox "MATLAB (matlab.exe) was not found automatically." & vbCrLf & _
           "Please open MATLAB manually and run:" & vbCrLf & _
           "run('" & appDir & "\start_app.m')", vbExclamation, "Figure Layout Tool"
    WScript.Quit 1
End If

Dim safeDir
safeDir = Replace(appDir, "'", "''")

matlabCmd = "try; addpath('" & safeDir & "'); addpath('" & safeDir & "\core'); " & _
            "app = FigureLayoutApp(true); waitfor(app.UIFigure); " & _
            "catch ME; disp(getReport(ME)); end; exit;"

' Window style 0 = hidden (no black console window, no MATLAB Desktop)
cmdLine = """" & matlabExe & """ -nodesktop -nosplash -r """ & matlabCmd & """"
shell.Run cmdLine, 0, False

WScript.Quit 0

Function FindMatlabExe(fso, shell)
    FindMatlabExe = ""

    ' 1. Try PATH via where matlab.exe
    On Error Resume Next
    Dim execObj, line
    Set execObj = shell.Exec("cmd /c where matlab.exe")
    If Err.Number = 0 Then
        Do While Not execObj.StdOut.AtEndOfStream
            line = Trim(execObj.StdOut.ReadLine)
            If line <> "" And fso.FileExists(line) Then
                FindMatlabExe = line
                Exit Function
            End If
        Loop
    End If
    Err.Clear
    On Error GoTo 0

    ' 2. Fallback: scan common MATLAB installation directories (pick latest release)
    Dim baseDirs, baseDir, rootFolder, subFolder, candidate, bestPath, bestName
    baseDirs = Array( _
        "C:\Program Files\MATLAB", _
        "D:\Program Files\MATLAB", _
        "C:\MATLAB", _
        "D:\MATLAB", _
        "E:\MATLAB" _
    )
    bestPath = ""
    bestName = ""

    For Each baseDir In baseDirs
        If fso.FolderExists(baseDir) Then
            Set rootFolder = fso.GetFolder(baseDir)
            For Each subFolder In rootFolder.SubFolders
                candidate = fso.BuildPath(subFolder.Path, "bin\matlab.exe")
                If fso.FileExists(candidate) Then
                    If StrComp(subFolder.Name, bestName, vbTextCompare) > 0 Then
                        bestName = subFolder.Name
                        bestPath = candidate
                    End If
                End If
            Next
        End If
    Next

    FindMatlabExe = bestPath
End Function
