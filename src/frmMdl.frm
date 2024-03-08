VERSION 5.00
Begin VB.Form frmMdl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Installation"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmMdl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   58
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
   Begin VB.Shape shape 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMdl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
