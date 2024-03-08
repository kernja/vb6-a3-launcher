Attribute VB_Name = "mdlNet"
Private Function Registry_Read(Key_Path, Key_Name) As Variant
On Error Resume Next
    Dim Registry As Object
    Set Registry = CreateObject("WScript.Shell")

    Registry_Read = Registry.RegRead(Key_Path & Key_Name)
End Function

Public Function isFramework2Installed() As Boolean
    Dim myReturn As Boolean
    myReturn = False
    
    If Registry_Read("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727\", "Install") = 1 Then
        myReturn = True
    End If
    
    isFramework2Installed = myReturn
End Function
