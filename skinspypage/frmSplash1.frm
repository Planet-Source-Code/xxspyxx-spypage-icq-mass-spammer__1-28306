VERSION 5.00
Begin VB.Form frmSplash1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3615
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   1320
      Picture         =   "frmSplash1.frx":000C
      ScaleHeight     =   1335
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   960
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "[ SpyPage v1.0 ]"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmSplash1.frx":2368
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   " [ http://es.webgrp.net ]"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      MouseIcon       =   "frmSplash1.frx":24BA
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_Click()
On Error Resume Next
gotoweb
End Sub


Private Sub Label2_Click()
On Error Resume Next
FRMTEST.Visible = True
Unload Me
End Sub


