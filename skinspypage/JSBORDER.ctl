VERSION 5.00
Begin VB.UserControl JSBORDER 
   Alignable       =   -1  'True
   BackColor       =   &H00DC7E5A&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   199
End
Attribute VB_Name = "JSBORDER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private JS_BMP1 As clsBitmap

Private JS_BMP2 As clsBitmap

Private JS_BMP3 As clsBitmap

Private pbl As PropertyBag

Enum BTYPE

    Left = 1

    Right = 2

    bottom = 3

End Enum

Private BORDERSTYLES As BTYPE

Private RESIZEHOW As Integer

Private JS_path As String

Public Property Let BORDERTYPE(NewBordertype As BTYPE)

BORDERSTYLES = NewBordertype

If BORDERSTYLES = Left Then

    UserControl.Parent.Controls(UserControl.Ambient.DisplayName).Align = vbAlignLeft

ElseIf BORDERSTYLES = Right Then

    UserControl.Parent.Controls(UserControl.Ambient.DisplayName).Align = vbAlignRight

ElseIf BORDERSTYLES = bottom Then

    UserControl.Parent.Controls(UserControl.Ambient.DisplayName).Align = vbAlignBottom

End If

PropertyChanged "BORDERTYPE"

End Property

Public Property Get BORDERTYPE() As BTYPE

BORDERTYPE = BORDERSTYLES

End Property

Private Sub DOSKIN()

If Ambient.UserMode = True Then

  Dim varTemp As Variant

   Dim byteArr() As Byte

 On Error Resume Next

        Set JS_BMP1 = New clsBitmap

        Set JS_BMP2 = New clsBitmap

        If BORDERSTYLES = bottom Then

            Set JS_BMP3 = New clsBitmap

        End If

   Set pbl = New PropertyBag

   Open JS_path For Binary As #1

   Get #1, , varTemp

   Close #1

   byteArr = varTemp

   pbl.Contents = byteArr

With pbl

    If BORDERSTYLES = 1 Then

        JS_BMP1.LoadResource .ReadProperty("LEFTTOP")

        JS_BMP2.LoadResource .ReadProperty("LEFTMID")

    ElseIf BORDERSTYLES = 2 Then

        JS_BMP1.LoadResource .ReadProperty("RIGHTTOP")

        JS_BMP2.LoadResource .ReadProperty("RIGHTMID")

    ElseIf BORDERSTYLES = 3 Then

        JS_BMP1.LoadResource .ReadProperty("LEFTBOT")

        JS_BMP2.LoadResource .ReadProperty("RIGHTBOT")

        JS_BMP3.LoadResource .ReadProperty("BOTTOM")

    End If

End With

If BORDERSTYLES = Left Then

    UserControl.Width = (JS_BMP1.Width) * Screen.TwipsPerPixelX

ElseIf BORDERSTYLES = Right Then

    UserControl.Width = (JS_BMP1.Width) * Screen.TwipsPerPixelX

ElseIf BORDERSTYLES = bottom Then

    UserControl.Height = (JS_BMP3.Height) * Screen.TwipsPerPixelY

End If

    If BORDERSTYLES = Left Then

        For z = 0 To UserControl.ScaleHeight

            BitBlt UserControl.hdc, 0, JS_BMP2.Height * z, JS_BMP2.Width, JS_BMP2.Height, JS_BMP2.hdc, 0, 0, SRCCOPY

        Next z

        BitBlt UserControl.hdc, 0, 0, JS_BMP1.Width, JS_BMP1.Height, JS_BMP1.hdc, 0, 0, SRCCOPY

    ElseIf BORDERSTYLES = Right Then

        For n = 0 To UserControl.ScaleHeight

            BitBlt UserControl.hdc, 0, JS_BMP2.Height * n, JS_BMP2.Width, JS_BMP2.Height, JS_BMP2.hdc, 0, 0, SRCCOPY

        Next n

        BitBlt UserControl.hdc, 0, 0, JS_BMP1.Width, JS_BMP1.Height, JS_BMP1.hdc, 0, 0, SRCCOPY

    ElseIf BORDERSTYLES = bottom Then

        For i = 0 To UserControl.ScaleWidth

            BitBlt UserControl.hdc, JS_BMP3.Width * i, 0, JS_BMP3.Width, JS_BMP3.Height, JS_BMP3.hdc, 0, 0, SRCCOPY

        Next i

        BitBlt UserControl.hdc, 0, 0, JS_BMP1.Width, JS_BMP1.Height, JS_BMP1.hdc, 0, 0, SRCCOPY

        BitBlt UserControl.hdc, UserControl.ScaleWidth - (JS_BMP2.Width), 0, JS_BMP2.Width, JS_BMP2.Height, JS_BMP2.hdc, 0, 0, SRCCOPY

    End If

Set JS_BMP1 = Nothing

Set JS_BMP2 = Nothing

Set JS_BMP3 = Nothing

Set pbl = Nothing

End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If JS_RESIZE = True Then

If BORDERSTYLES = Left Then

    ReleaseCapture

    SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTLEFT, 0

ElseIf BORDERSTYLES = Right Then

    ReleaseCapture

    SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTRIGHT, 0

ElseIf BORDERSTYLES = bottom Then

    If RESIZEHOW = 0 Then

        ReleaseCapture

        SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMLEFT, 0

    ElseIf RESIZEHOW = 1 Then

            ReleaseCapture

            SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTBOTTOM, 0

    ElseIf RESIZEHOW = 2 Then

            ReleaseCapture

            SendMessage UserControl.Parent.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0

    End If

End If

End If

End Sub

Public Property Get Path() As String

Path = JS_path

End Property

Public Property Let Path(NewPath As String)

JS_path = NewPath

PropertyChanged "Path"

DOSKIN

End Property

Public Function REDRAW()

UserControl.Refresh

End Function

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If JS_RESIZE = True Then

    If BORDERSTYLES = bottom Then

        If x >= 0 And x <= 10 Then

            RESIZEHOW = 0

            UserControl.MousePointer = 6

        ElseIf x >= UserControl.ScaleWidth - 10 And x <= UserControl.ScaleWidth Then

            RESIZEHOW = 2

            UserControl.MousePointer = 8

        Else

            RESIZEHOW = 1

            UserControl.MousePointer = 7

        End If

    Else: UserControl.MousePointer = 9

    End If

End If

End Sub

Private Sub UserControl_Paint()

DOSKIN

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

BORDERSTYLES = PropBag.ReadProperty("BORDERTYPE", 0)

JS_path = PropBag.ReadProperty("Path", "")

End Sub

Private Sub UserControl_Resize()

If Ambient.UserMode = False Then

    If BORDERSTYLES = Left Then

        UserControl.Width = 100

    ElseIf BORDERSTYLES = Right Then

        UserControl.Width = 100

    ElseIf BORDERSTYLES = bottom Then

        UserControl.Height = 100

    End If

End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

PropBag.WriteProperty "BORDERTYPE", BORDERSTYLES, 0

PropBag.WriteProperty "Path", JS_path, ""

End Sub

