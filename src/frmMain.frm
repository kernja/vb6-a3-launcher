VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adventures of Woody and Captain Ion CD 1"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11385
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":25CA
   ScaleHeight     =   502
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   759
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image btnCancel 
      Height          =   315
      Left            =   5805
      MousePointer    =   10  'Up Arrow
      Top             =   6825
      Width           =   1335
   End
   Begin VB.Image btnInstall 
      Height          =   315
      Left            =   4215
      MousePointer    =   10  'Up Arrow
      Top             =   6825
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const softwareName As String = "A3CD1"
Private Const installedBG As String = "bgInstalled.bmp"

Private Sub Form_Load()
    If mdlMain.isA3SoftwareInstalled(softwareName) = True Then
        Me.Picture = LoadPicture(App.path & installedBG)
    End If
End Sub

Private Sub btnCancel_Click()
    End
End Sub

Private Sub btnInstall_Click()
    Dim path As String
     
    If mdlMain.isA3SoftwareInstalled(softwareName) = False Then
        Me.Hide
        mdlMain.Main
    Else
        Me.Hide
        path = mdlMain.getA3SoftwareInstallPath(softwareName)
        path = path & "\" & softwareName & ".exe"
        ExecCmd (path)
        End
    End If
    
End Sub

