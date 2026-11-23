Option Explicit

' Creates a Windows Desktop shortcut with the custom frosted-glass app_icon.ico.
' 100% pure ASCII script & target filename so it works on all Windows locales (GBK/UTF-8).

Dim fso, shell, appDir, vbsTarget, iconPath, desktopDir, lnkPath, sc

Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

appDir     = fso.GetParentFolderName(WScript.ScriptFullName)
vbsTarget  = fso.BuildPath(appDir, "Launch_Figure_Studio_Windows.vbs")
iconPath   = fso.BuildPath(appDir, "assets\app_icon.ico")
desktopDir = shell.SpecialFolders("Desktop")
lnkPath    = fso.BuildPath(desktopDir, "Figure Studio.lnk")

If Not fso.FileExists(vbsTarget) Then
    MsgBox "Cannot find launcher script:" & vbCrLf & vbsTarget, vbCritical, "Figure Studio"
    WScript.Quit 1
End If

Set sc = shell.CreateShortcut(lnkPath)
sc.TargetPath       = "wscript.exe"
sc.Arguments        = """" & vbsTarget & """"
sc.WorkingDirectory = appDir
sc.Description      = "MATLAB Scientific Figure Layout Tool"
If fso.FileExists(iconPath) Then
    sc.IconLocation = iconPath & ",0"
End If
sc.Save

MsgBox "Desktop shortcut created successfully:" & vbCrLf & lnkPath, vbInformation, "Figure Studio"
WScript.Quit 0
