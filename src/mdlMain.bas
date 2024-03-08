Attribute VB_Name = "mdlMain"
Public Declare Function SHGetSpecialFolderPath _
   Lib "shell32.dll" _
   Alias "SHGetSpecialFolderPathA" _
   (ByVal hWnd As Long, _
   ByVal lpszPath As String, _
   ByVal nFolder As Integer, _
   ByVal fCreate As Boolean) As Boolean
   
Public Function Registry_Read(Key_Path, Key_Name) As Variant
On Error Resume Next
    Dim Registry As Object
    Set Registry = CreateObject("WScript.Shell")

Registry_Read = Registry.RegRead(Key_Path & Key_Name)
End Function

Public Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(FileName) And vbDirectory) = 0

ErrorHandler:
    FileExists = False
End Function

Public Function getA3SoftwareInstallPath(name As String) As String
    getA3CD1Path = Registry_Read("HKEY_LOCAL_MACHINE\SOFTWARE\A3 Ltd.\" & name & "\", "InstallPath")
End Function

Public Function isFramework2Installed() As Boolean
    Dim myReturn As Boolean
    myReturn = False
    
    If Registry_Read("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727\", "Install") = 1 Then
        myReturn = True
    End If
    
    isFramework2Installed = myReturn
End Function

Public Function isA3SoftwareInstalled(name As String) As Boolean
    Dim myReturn As Boolean
    myReturn = False
    
    If Not (getA3SoftwareInstallPath(name) = "") Then
        myReturn = True
    End If
    
    isA3CD1Installed = myReturn
End Function

Sub Main()
    On Error Resume Next
    
    'reset the progress bar
    frmMdl.shape.Width = 0
    frmMdl.Show
    
    'force redraw and execute flash player installation
    DoEvents
        ExecCmd ("MsiExec.exe /i " & Chr(34) & "install_flash_player_active_x.msi" & Chr(34) & "/q REBOOT=ReallySuppress")
    
    'force redraw and update progress bar
    DoEvents
        frmMdl.shape.Width = 75
        frmMdl.lblProgress.Caption = "33%"
    
    'force redraw and see if net is installed
    DoEvents
        If mdlNet.isFramework2Installed = False Then
            ExecCmd ("Dotnetfx.exe /q:a /c:" & Chr(34) & "install /q" & Chr(34))
        End If
    
    'force redraw and launch installer
    DoEvents
        frmMdl.shape.Width = 150
        frmMdl.lblProgress.Caption = "75%"
    
        ExecCmd ("setup.exe")
        frmMdl.Hide
    End
End Sub

