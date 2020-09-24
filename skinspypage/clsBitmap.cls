VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDC As Long

Private mBitmap As Long

Private mOldBitmap As Long

Private mWidth As Long

Private mHeight As Long

Public Function LoadResource(ResourceID As Long) As Boolean

    ClearAll

    mBitmap = ResourceID

    If (mBitmap <> API_NULL_HANDLE) Then

        LoadResource = LoadBitmapIntoDC

    End If

End Function

Private Function LoadBitmapIntoDC() As Boolean

Dim ScreenDC As Long

Dim BitmapData As BITMAP

    ScreenDC = GetDC(API_NULL_HANDLE)

    mDC = CreateCompatibleDC(ScreenDC)

    ReleaseDC API_NULL_HANDLE, ScreenDC

    If (mDC <> API_NULL_HANDLE) Then

        mOldBitmap = SelectObject(mDC, mBitmap)

        GDIGetObject mBitmap, Len(BitmapData), BitmapData

        mWidth = BitmapData.bmWidth

        mHeight = BitmapData.bmHeight

        LoadBitmapIntoDC = True

    End If

End Function

Property Get Width() As Long

    Width = mWidth

End Property

Property Get Height() As Long

    Height = mHeight

End Property

Property Get hdc() As Long

    hdc = mDC

End Property

Private Sub ClearAll()

    If (mDC <> API_NULL_HANDLE) Then

        If (mBitmap <> API_NULL_HANDLE) Then

            SelectObject mDC, mOldBitmap

            DeleteObject mBitmap

            mBitmap = API_NULL_HANDLE

        End If

        DeleteObject mDC

        mDC = API_NULL_HANDLE

    End If

End Sub

Private Sub Class_Terminate()

    ClearAll

End Sub

