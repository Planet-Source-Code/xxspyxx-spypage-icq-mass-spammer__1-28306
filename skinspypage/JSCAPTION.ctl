VERSION 5.00
Begin VB.UserControl JSCAPTION 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3060
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   MouseIcon       =   "JSCAPTION.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   26
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   204
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Stretch         =   -1  'True
      Top             =   75
      Width           =   240
   End
End
Attribute VB_Name = "JSCAPTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private pb As PropertyBag

Private JS_path As String

Private FRMontop As New clsOnTop

Private JS_TOPLEFT As clsBitmap

Private JS_TOPMID As clsBitmap

Private JS_TOPRIGHT As clsBitmap

Private JS_CLOSE As clsBitmap

Private JS_MAX As clsBitmap

Private JS_MIN As clsBitmap

Private JS_BONTOP As clsBitmap

Private JS_DRAGOK As Boolean

Private JS_ONTOP As Boolean

Private JS_XOFFSET As Integer

Private JS_YOFFSET As Integer

Private JS_SHOWONTOP As Boolean

Private JS_FROMTOP As Integer

Private JS_FROMRIGHT As Integer

Private JS_ICONSPACE As Integer

Private JS_SHOWICON As Boolean

Private JS_CONTROLBOX As Boolean

Enum ACTION

    jsclose = 0

    jsmin = 1

    jsmax = 2

    jsontop = 3

End Enum

Private JS_DOWHAT As ACTION

Private JS_DOACTION As Boolean

Private JS_BORDERSTYLE As JS_BORDER

Enum JS_BORDER

    FIXED = 0

    SIZABLE = 1

End Enum

Public Property Get ControlBox() As Boolean

ControlBox = JS_CONTROLBOX

End Property

Public Property Let ControlBox(newvalue As Boolean)

JS_CONTROLBOX = newvalue

PropertyChanged "ControlBox"

End Property

Public Property Let Movable(newvalue As Boolean)

JS_DRAGOK = newvalue

PropertyChanged "Movable"

End Property

Public Property Get Movable() As Boolean

Movable = JS_DRAGOK

End Property

Private Sub FormDrag()

If JS_DRAGOK = True Then

    ReleaseCapture

    SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0

End If

End Sub

Private Sub DOSKIN()

If Ambient.UserMode = True Then

  Dim varTemp As Variant

   Dim byteArr() As Byte

 On Error Resume Next

Set JS_TOPLEFT = New clsBitmap

Set JS_TOPMID = New clsBitmap

Set JS_TOPRIGHT = New clsBitmap

Set JS_CLOSE = New clsBitmap

Set JS_MAX = New clsBitmap

Set JS_MIN = New clsBitmap

Set JS_BONTOP = New clsBitmap

   Set pb = New PropertyBag

   Open JS_path For Binary As #1

   Get #1, , varTemp

   Close #1

   byteArr = varTemp

   pb.Contents = byteArr

With pb

JS_TOPLEFT.LoadResource .ReadProperty("TOPLEFT")

JS_TOPMID.LoadResource .ReadProperty("TOPMID")

JS_TOPRIGHT.LoadResource .ReadProperty("TOPRIGHT")

JS_CLOSE.LoadResource .ReadProperty("CLOSE")

If UserControl.Parent.WindowState = 2 Then

    JS_MAX.LoadResource .ReadProperty("RES1")

Else

    JS_MAX.LoadResource .ReadProperty("MAX")

End If

JS_MIN.LoadResource .ReadProperty("MIN")

If JS_ONTOP = True Then

    JS_BONTOP.LoadResource .ReadProperty("ONTOP3")

Else

        JS_BONTOP.LoadResource .ReadProperty("ONTOP1")

End If

JS_XOFFSET = .ReadProperty("XOFFSET")

JS_YOFFSET = .ReadProperty("YOFFSET")

JS_FROMRIGHT = .ReadProperty("FROMRIGHT")

JS_FROMTOP = .ReadProperty("FROMTOP")

JS_ICONSPACE = .ReadProperty("ICONSPACE")

UserControl.ForeColor = .ReadProperty("FORECOLOR")

End With

UserControl.Height = (JS_TOPLEFT.Height) * Screen.TwipsPerPixelY

For i = 0 To UserControl.ScaleWidth

    BitBlt UserControl.hdc, JS_TOPMID.Width * i, 0, JS_TOPMID.Width, JS_TOPMID.Height, JS_TOPMID.hdc, 0, 0, SRCCOPY

Next i

BitBlt UserControl.hdc, 0, 0, JS_TOPLEFT.Width, JS_TOPLEFT.Height, JS_TOPLEFT.hdc, 0, 0, SRCCOPY

BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_TOPRIGHT.Width, 0, JS_TOPRIGHT.Width, JS_TOPRIGHT.Height, JS_TOPRIGHT.hdc, 0, 0, SRCCOPY

If JS_SHOWICON = True Then

 Image1.Picture = UserControl.Parent.Icon

End If

If JS_CONTROLBOX = True Then

    BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_CLOSE.Width, JS_CLOSE.Height, JS_CLOSE.hdc, 0, 0, SRCCOPY

    If JS_BORDERSTYLE = SIZABLE Then

        BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_MAX.Width, JS_MAX.Height, JS_MAX.hdc, 0, 0, SRCCOPY

        BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_MIN.Width - JS_ICONSPACE - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_MIN.Width, JS_MIN.Height, JS_MIN.hdc, 0, 0, SRCCOPY

    End If

End If

If JS_SHOWONTOP = True Then

        BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_BONTOP.Width - JS_ICONSPACE - JS_MIN.Width - JS_ICONSPACE - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_BONTOP.Width, JS_BONTOP.Height, JS_BONTOP.hdc, 0, 0, SRCCOPY

End If

UserControl.CurrentX = JS_XOFFSET

UserControl.CurrentY = JS_YOFFSET

UserControl.Print UserControl.Parent.Caption

Set JS_TOPLEFT = Nothing

Set JS_TOPMID = Nothing

Set JS_TOPRIGHT = Nothing

End If

End Sub

Public Property Get ONTOP() As Boolean

ONTOP = JS_ONTOP

End Property

Public Property Let ONTOP(newvalue As Boolean)

JS_ONTOP = newvalue

If JS_ONTOP = True Then

    FRMontop.MakeTopMost UserControl.Parent.hWnd

ElseIf JS_ONTOP = False Then

    FRMontop.MakeNormal UserControl.Parent.hWnd

End If

PropertyChanged "ONTOP"

End Property

Public Property Get Path() As String

Path = JS_path

End Property

Public Property Let Path(NewPath As String)

JS_path = NewPath

PropertyChanged "Path"

DOSKIN

End Property

Public Property Get ShowIcon() As Boolean

ShowIcon = JS_SHOWICON

End Property

Public Property Let ShowIcon(newvalue As Boolean)

JS_SHOWICON = newvalue

PropertyChanged "ShowIcon"

End Property

Public Property Get SHOWONTOP() As Boolean

SHOWONTOP = JS_SHOWONTOP

End Property

Public Property Let SHOWONTOP(newvalue As Boolean)

JS_SHOWONTOP = newvalue

PropertyChanged "SHOWONTOP"

End Property

Public Property Get Style() As JS_BORDER

Style = JS_BORDERSTYLE

End Property

Public Property Let Style(newstyle As JS_BORDER)

JS_BORDERSTYLE = newstyle

PropertyChanged "Style"

End Property

Public Function REDRAW()

UserControl.Refresh

End Function

Private Sub UserControl_Click()

If JS_CONTROLBOX = True Then

    If JS_DOACTION = True Then

        If JS_DOWHAT = jsclose Then

            Unload UserControl.Parent

        ElseIf JS_DOWHAT = jsontop Then

            If JS_ONTOP = True Then

                FRMontop.MakeNormal UserControl.Parent.hWnd

                JS_ONTOP = False

            JS_BONTOP.LoadResource pb.ReadProperty("ONTOP1")

            BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_BONTOP.Width - JS_ICONSPACE - JS_MIN.Width - JS_ICONSPACE - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_BONTOP.Width, JS_BONTOP.Height, JS_BONTOP.hdc, 0, 0, SRCCOPY

            ElseIf JS_ONTOP = False Then

                FRMontop.MakeTopMost UserControl.Parent.hWnd

                JS_ONTOP = True

                JS_BONTOP.LoadResource pb.ReadProperty("ONTOP3")

                BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_BONTOP.Width - JS_ICONSPACE - JS_MIN.Width - JS_ICONSPACE - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_BONTOP.Width, JS_BONTOP.Height, JS_BONTOP.hdc, 0, 0, SRCCOPY

            End If

        ElseIf JS_DOWHAT = jsmax Then

        If UserControl.Parent.WindowState = 2 Then

            UserControl.Parent.WindowState = 0

        Else

        UserControl.Parent.WindowState = 2

        End If

     ElseIf JS_DOWHAT = jsmin Then

         UserControl.Parent.WindowState = 1

     End If

    End If

End If

End Sub

Private Sub UserControl_DblClick()

If JS_DOACTION = False Then

If UserControl.Parent.WindowState = 2 Then

    UserControl.Parent.WindowState = 0

Else

    UserControl.Parent.WindowState = 2

End If

End If

End Sub

Private Sub UserControl_Initialize()

Set FRMontop = New clsOnTop

End Sub

Private Sub UserControl_InitProperties()

    Set UserControl.Font = Ambient.Font

    UserControl.Parent.Controls(UserControl.Ambient.DisplayName).Align = vbAlignTop

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If JS_DOACTION = True Then

If JS_CONTROLBOX = True Then

 If JS_DOWHAT = jsclose Then

    JS_CLOSE.LoadResource pb.ReadProperty("CLOSE2")

    BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_CLOSE.Width, JS_CLOSE.Height, JS_CLOSE.hdc, 0, 0, SRCCOPY

 ElseIf JS_DOWHAT = jsmax Then

 If UserControl.Parent.WindowState = 2 Then

    JS_MAX.LoadResource pb.ReadProperty("RES3")

    BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_MAX.Width, JS_MAX.Height, JS_MAX.hdc, 0, 0, SRCCOPY

 Else

     JS_MAX.LoadResource pb.ReadProperty("MAX2")

     BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_MAX.Width, JS_MAX.Height, JS_MAX.hdc, 0, 0, SRCCOPY

 End If

 ElseIf JS_DOWHAT = jsmin Then

    JS_MIN.LoadResource pb.ReadProperty("MIN2")

    BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_MIN.Width - JS_ICONSPACE - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_MIN.Width, JS_MIN.Height, JS_MIN.hdc, 0, 0, SRCCOPY

 End If

End If

Else

FormDrag

End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If JS_CONTROLBOX = True Then

If y > JS_FROMTOP And y < (JS_CLOSE.Height + JS_FROMTOP) Then

    If x > UserControl.ScaleWidth - JS_CLOSE.Width - JS_FROMRIGHT And x < UserControl.ScaleWidth - JS_FROMRIGHT Then

        JS_CLOSE.LoadResource pb.ReadProperty("CLOSE3")

        BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_CLOSE.Width, JS_CLOSE.Height, JS_CLOSE.hdc, 0, 0, SRCCOPY

        JS_DOWHAT = jsclose

        JS_DOACTION = True

    ElseIf x > UserControl.ScaleWidth - JS_CLOSE.Width - JS_ICONSPACE - JS_MAX.Width - JS_FROMRIGHT And x < UserControl.ScaleWidth - JS_CLOSE.Width - JS_ICONSPACE - JS_FROMRIGHT Then

        If JS_BORDERSTYLE = SIZABLE Then

         If UserControl.Parent.WindowState = 2 Then

            JS_MAX.LoadResource pb.ReadProperty("RES2")

            BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_MAX.Width, JS_MAX.Height, JS_MAX.hdc, 0, 0, SRCCOPY

            JS_DOWHAT = jsmax

            JS_DOACTION = True

         Else

            JS_MAX.LoadResource pb.ReadProperty("MAX3")

            BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_MAX.Width, JS_MAX.Height, JS_MAX.hdc, 0, 0, SRCCOPY

            JS_DOWHAT = jsmax

            JS_DOACTION = True

         End If

        End If

    ElseIf x > UserControl.ScaleWidth - JS_MIN.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_ICONSPACE - JS_MAX.Width - JS_FROMRIGHT And x < UserControl.ScaleWidth - JS_CLOSE.Width - JS_MIN.Width - JS_ICONSPACE - JS_ICONSPACE - JS_FROMRIGHT Then

        If JS_BORDERSTYLE = SIZABLE Then

            JS_MIN.LoadResource pb.ReadProperty("MIN3")

            BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_MIN.Width - JS_ICONSPACE - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_MIN.Width, JS_MIN.Height, JS_MIN.hdc, 0, 0, SRCCOPY

            JS_DOWHAT = jsmin

            JS_DOACTION = True

        End If

    ElseIf x > UserControl.ScaleWidth - JS_BONTOP.Width - JS_ICONSPACE - JS_MIN.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_ICONSPACE - JS_MAX.Width - JS_FROMRIGHT And x < UserControl.ScaleWidth - JS_ICONSPACE - JS_MIN.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_ICONSPACE - JS_MAX.Width - JS_FROMRIGHT Then

       If JS_SHOWONTOP = True Then

            JS_DOWHAT = jsontop

            JS_DOACTION = True

        End If

    Else

        JS_CLOSE.LoadResource pb.ReadProperty("CLOSE")

        BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_CLOSE.Width, JS_CLOSE.Height, JS_CLOSE.hdc, 0, 0, SRCCOPY

        If JS_BORDERSTYLE = SIZABLE Then

                 If UserControl.Parent.WindowState = 2 Then

            JS_MAX.LoadResource pb.ReadProperty("RES1")

 Else

             JS_MAX.LoadResource pb.ReadProperty("MAX")

 End If

             BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_MAX.Width, JS_MAX.Height, JS_MAX.hdc, 0, 0, SRCCOPY

            JS_MIN.LoadResource pb.ReadProperty("MIN")

            BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_MIN.Width - JS_ICONSPACE - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_MIN.Width, JS_MIN.Height, JS_MIN.hdc, 0, 0, SRCCOPY

        End If

        JS_DOACTION = False

    End If

End If

End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If JS_CONTROLBOX = True Then

If JS_DOACTION = True Then

 If JS_DOWHAT = jsclose Then

    JS_CLOSE.LoadResource pb.ReadProperty("CLOSE")

    BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_CLOSE.Width, JS_CLOSE.Height, JS_CLOSE.hdc, 0, 0, SRCCOPY

 ElseIf JS_DOWHAT = jsmax Then

    JS_MAX.LoadResource pb.ReadProperty("MAX")

    BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_MAX.Width, JS_MAX.Height, JS_MAX.hdc, 0, 0, SRCCOPY

 ElseIf JS_DOWHAT = jsmin Then

    JS_MIN.LoadResource pb.ReadProperty("MIN")

    BitBlt UserControl.hdc, UserControl.ScaleWidth - JS_MIN.Width - JS_ICONSPACE - JS_MAX.Width - JS_ICONSPACE - JS_CLOSE.Width - JS_FROMRIGHT, JS_FROMTOP, JS_MIN.Width, JS_MIN.Height, JS_MIN.hdc, 0, 0, SRCCOPY

 End If

End If

End If

End Sub

Private Sub UserControl_Paint()

DOSKIN

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

JS_path = PropBag.ReadProperty("Path", "")

JS_DRAGOK = PropBag.ReadProperty("Movable", True)

JS_BORDERSTYLE = PropBag.ReadProperty("Style", 1)

JS_SHOWICON = PropBag.ReadProperty("ShowIcon", True)

JS_CONTROLBOX = PropBag.ReadProperty("ControlBox", True)

JS_ONTOP = PropBag.ReadProperty("ONTOP", False)

JS_SHOWONTOP = PropBag.ReadProperty("SHOWONTOP", False)

    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)

If JS_BORDERSTYLE = FIXED Then

        JS_RESIZE = False

    Else

        JS_RESIZE = True

    End If

If JS_DRAGOK = True Then

    UserControl.MousePointer = 99

Else

    UserControl.MousePointer = 0

End If

If JS_ONTOP = True Then

    FRMontop.MakeTopMost UserControl.Parent.hWnd

ElseIf JS_ONTOP = False Then

    FRMontop.MakeNormal UserControl.Parent.hWnd

End If

End Sub

Private Sub UserControl_Resize()

DOSKIN

If Ambient.UserMode = False Then

UserControl.Height = 400

End If

Image1.Move 8, 5

If JS_SHOWICON = True Then

 Image1.Refresh

End If

End Sub

Private Sub UserControl_Terminate()

Set JS_TOPLEFT = Nothing

Set JS_TOPMID = Nothing

Set JS_TOPRIGHT = Nothing

Set JS_CLOSE = Nothing

Set JS_MAX = Nothing

Set JS_MIN = Nothing

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

PropBag.WriteProperty "Path", JS_path, ""

PropBag.WriteProperty "Movable", JS_DRAGOK, True

PropBag.WriteProperty "ShowIcon", JS_SHOWICON, True

PropBag.WriteProperty "ControlBox", JS_CONTROLBOX, True

PropBag.WriteProperty "ONTOP", JS_ONTOP, False

PropBag.WriteProperty "SHOWONTOP", JS_SHOWONTOP, False

PropBag.WriteProperty "Style", JS_BORDERSTYLE, 1

    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)

End Sub

Public Property Get Font() As Font



    Set Font = UserControl.Font

End Property

Public Property Set Font(ByVal New_Font As Font)

    Set UserControl.Font = New_Font

    PropertyChanged "Font"

    DOSKIN

End Property

