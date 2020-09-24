VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{F2BF180A-9B19-4644-B5BD-A9238B839A24}#1.0#0"; "JSSKIN.ocx"
Begin VB.Form FRMTEST 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "SpyPage v1.0"
   ClientHeight    =   8355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   Icon            =   "FRMTEST.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin SpyPageControl.JSBORDER JSBORDER2 
      Align           =   4  'Align Right
      Height          =   7845
      Left            =   6735
      TabIndex        =   41
      Top             =   405
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   13838
      BORDERTYPE      =   2
   End
   Begin SpyPageControl.JSBORDER JSBORDER1 
      Align           =   2  'Align Bottom
      Height          =   105
      Left            =   0
      TabIndex        =   40
      Top             =   8250
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   185
      BORDERTYPE      =   3
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2280
      Picture         =   "FRMTEST.frx":1CFA
      ScaleHeight     =   375
      ScaleWidth      =   2175
      TabIndex        =   39
      Top             =   720
      Width           =   2175
   End
   Begin SpyPageControl.JSBORDER JSBORDER3 
      Align           =   3  'Align Left
      Height          =   7845
      Left            =   0
      TabIndex        =   38
      Top             =   405
      Width           =   105
      _ExtentX        =   185
      _ExtentY        =   13838
      BORDERTYPE      =   1
   End
   Begin SpyPageControl.JSCAPTION JSCAPTION1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   37
      Top             =   0
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   714
      SHOWONTOP       =   -1  'True
      Style           =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   6240
   End
   Begin VB.TextBox TextUin 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   35
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Skin:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   25
      Top             =   7440
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "Whistler"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Luna"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Options:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   3840
      TabIndex        =   20
      Top             =   4440
      Width           =   2775
      Begin VB.CommandButton cmdMultiRemove 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton open 
         Caption         =   "Load List"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   23
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdMultiAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   22
         ToolTipText     =   "Add the above account to the list"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtinputproxy 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   480
         TabIndex        =   21
         Text            =   "Proxy"
         ToolTipText     =   "Username to be added"
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.TextBox txtProxy 
      Height          =   285
      Left            =   1080
      TabIndex        =   19
      Top             =   9000
      Width           =   1575
   End
   Begin VB.TextBox txtNewProxy 
      Height          =   285
      Left            =   1080
      TabIndex        =   18
      Top             =   9360
      Width           =   1575
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1080
      TabIndex        =   17
      Top             =   9720
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Proxy List:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   3840
      TabIndex        =   11
      Top             =   1560
      Width           =   2775
      Begin VB.CommandButton cmdMultiSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         ToolTipText     =   "Save the list"
         Top             =   2040
         Width           =   615
      End
      Begin VB.ListBox lstProxys 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1425
         ItemData        =   "FRMTEST.frx":2AB8
         Left            =   480
         List            =   "FRMTEST.frx":2ABA
         MultiSelect     =   2  'Extended
         TabIndex        =   12
         ToolTipText     =   "Username List"
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblMultiAccounts 
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Number of Proxys:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label19 
         BackColor       =   &H00000000&
         Caption         =   "Proxys:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
   End
   Begin MSWinsockLib.Winsock SockPager 
      Left            =   240
      Top             =   8640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "UIN List:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   3495
      Begin VB.CommandButton Command6 
         Caption         =   "Add UIN"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox UIN 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   31
         Text            =   "UIN #"
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   30
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Load List"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   29
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ListBox UINlist 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1230
         ItemData        =   "FRMTEST.frx":2ABC
         Left            =   240
         List            =   "FRMTEST.frx":2ABE
         TabIndex        =   28
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label uinnumber 
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   34
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "UIN's:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   33
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Send Message to ICQ UIN(s):"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2550
      End
   End
   Begin VB.TextBox TextMail 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   3495
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox TextMessage 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   240
      MaxLength       =   450
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CommandButton BtnSend 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "&Send"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MaskColor       =   &H00400000&
      TabIndex        =   1
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox TextSubject 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      MaxLength       =   30
      TabIndex        =   0
      Top             =   5040
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Current UIN:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   36
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "From Email:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label LabelStatus 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   6960
      Width           =   3495
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   5400
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   4800
      Width           =   720
   End
End
Attribute VB_Name = "FRMTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cMessage As String
Dim cSubject As String
Dim cMail As String
Dim cUin As String
Dim i As Integer
Dim NextLine As String



Private Sub cmdMultiAdd_Click()
On Error Resume Next
lstProxys.AddItem txtinputproxy.Text
lblMultiAccounts.Caption = lstProxys.ListCount
End Sub

Private Sub Command1_Click()
On Error Resume Next
Me.JSCAPTION1.Path = App.Path & "\PW.jss"
Me.JSBORDER1.Path = App.Path & "\PW.jss"
Me.JSBORDER2.Path = App.Path & "\PW.jss"
Me.JSBORDER3.Path = App.Path & "\PW.jss"
Me.JSCAPTION1.REDRAW
End Sub

Private Sub Command2_Click()
On Error Resume Next
Me.JSCAPTION1.Path = App.Path & "\luna.jss"
Me.JSBORDER1.Path = App.Path & "\luna.jss"
Me.JSBORDER2.Path = App.Path & "\luna.jss"
Me.JSBORDER3.Path = App.Path & "\luna.jss"
Me.JSCAPTION1.REDRAW
End Sub



Private Sub Command3_Click()
On Error Resume Next
  Dim Uins
 Uins = FreeFile
 Open "uins.txt" For Input As #Uins
 While Not EOF(1)
 Line Input #Uins, NextLine
 UINlist.AddItem NextLine
 Wend
 Close #Uins
 uinnumber.Caption = UINlist.ListCount
End Sub

Private Sub Command4_Click()
On Error Resume Next
UINlist.RemoveItem UINlist.ListIndex
uinnumber.Caption = UINlist.ListCount
End Sub



Private Sub Command6_Click()
On Error Resume Next
UINlist.AddItem UIN.Text
uinnumber.Caption = UINlist.ListCount
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.JSCAPTION1.Path = App.Path & "\pw.jss"
Me.JSBORDER1.Path = App.Path & "\pw.jss"
Me.JSBORDER2.Path = App.Path & "\pw.jss"
Me.JSBORDER3.Path = App.Path & "\pw.jss"

i = 0

On Error Resume Next
  Dim proxys
  
 proxys = FreeFile
 Open "proxys.txt" For Input As #proxys
 While Not EOF(1)
 Line Input #proxys, NextLine
 lstProxys.AddItem NextLine
 Wend
 Close #proxys
 lblMultiAccounts.Caption = lstProxys.ListCount
 
  Dim UIN
 UIN = FreeFile
 Open "uins.txt" For Input As #UIN
 While Not EOF(1)
 Line Input #UIN, NextLine
 UINlist.AddItem NextLine
 Wend
 Close #UIN
 uinnumber.Caption = UINlist.ListCount
   
   
   
   SockPager.Close
txtProxy.Visible = False
txtNewProxy.Visible = False
txtPort.Visible = False
End Sub

Private Sub BtnExit_Click()
On Error Resume Next
   End
End Sub

Private Sub BtnSend_Click()
   On Error Resume Next
   
   Dim cSend As String
   Dim cData As String

   Timer1.Interval = 15000 ' 5000 = 5 Seconds
Timer1.Enabled = True
   
   If Not IsNumeric(TextUin.Text) Then
      LabelStatus.Caption = "Please wait 15 Seconds...."
         
      TextUin.SetFocus
      Exit Sub
   End If
         
   If Trim(TextMessage.Text) = "" Then
      MsgBox "Don't Allow Blank Messages"
         
      TextMessage.SetFocus
      Exit Sub
   End If



   LabelStatus.Caption = "Connecting to Proxy..."
   
   
   SockPager.Close
      
   
   
   cSubject = ChangeSpaces(TextSubject.Text)
   cMessage = ChangeSpaces(TextMessage.Text)
   cMail = ChangeSpaces(TextMail.Text)
   cUin = ChangeSpaces(TextUin.Text)
   
   cData = "from=anonymous&fromemail=" & cMail & "&subject=" & cSubject & "&body=" & cMessage & "&to=" & Trim(cUin) & "&Send=" & """"

   cSend = "POST http://wwp.icq.com/scripts/WWPMsg.dll HTTP/1.0" & vbCrLf
   cSend = cSend & "Referer: http://wwp.icq.com" & vbCrLf
   cSend = cSend & "User-Agent: Mozilla/4.06 (Win95; I)" & vbCrLf
   cSend = cSend & "Proxy-Connection: Keep-Alive" & vbCrLf
   cSend = cSend & "Host: wwp.icq.com:80" & vbCrLf
   cSend = cSend & "Content-type: application/x-www-form-urlencoded" & vbCrLf
   cSend = cSend & "Content-length: " & Len(cData) & vbCrLf
   cSend = cSend & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf & vbCrLf
   cSend = cSend & cData & vbCrLf & vbCrLf & vbCrLf & vbCrLf
   
   SockPager.Tag = cSend

Dim SP() As String
Dim iProxy As Integer

iProxy = Int(Rnd * lstProxys.ListCount)
SP = Split(lstProxys.List(iProxy), ":")
txtNewProxy.Text = SP(0)
txtPort.Text = SP(1)



SockPager.Connect txtNewProxy.Text, txtPort.Text


End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next


   
   SockPager.Close
   
  
   End
End Sub

Private Sub open_Click()
On Error Resume Next
  Dim proxys
  
 proxys = FreeFile
 Open "proxys.txt" For Input As #proxys
 While Not EOF(1)
 Line Input #proxys, NextLine
 lstProxys.AddItem NextLine
 Wend
 Close #proxys
 lblMultiAccounts.Caption = lstProxys.ListCount
End Sub

Private Sub SockPager_Connect()
   On Error Resume Next
   
   
   LabelStatus.Caption = "Sending..."
  
   SockPager.SendData SockPager.Tag
End Sub

Private Sub SockPager_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   On Error Resume Next
   LabelStatus.Caption = "Error..."
   
   SockPager.Tag = ""
End Sub

Private Sub SockPager_SendComplete()
   On Error Resume Next
   LabelStatus.Caption = "Sent..."
   
   SockPager.Tag = ""
End Sub

Private Function ChangeSpaces(cString As String) As String
   On Error Resume Next
  
   Dim cChar As String
   Dim cReturn As String
  
   Dim nLoop As Long
  
  
   cReturn = ""
  
   For nLoop = 1 To Len(cString)
       cChar = Mid(cString, nLoop, 1)
      
       If cChar = " " Then
          cChar = "+"
       End If
      
       cReturn = cReturn + cChar
   Next
  
   ChangeSpaces = cReturn
End Function

Private Sub Winsock1_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub



Private Sub cmdMultiRefresh_Click()
On Error Resume Next
lblMultiAccounts.Caption = lstProxys.ListCount
End Sub

Private Sub cmdMultiRemove_Click()
On Error Resume Next
lstProxys.RemoveItem lstProxys.ListIndex
lblMultiAccounts.Caption = lstProxys.ListCount
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
LabelStatus.Caption = "Connected..."


 If i < UINlist.ListCount Then
 TextUin.Text = UINlist.List(i)
 SockPager.Close
 SockPager.Connect txtNewProxy.Text, txtPort.Text
 i = i + 1
 Else
 TextUin.Text = "End of Listbox"
 Timer1.Enabled = False
 End If
 End Sub
