Attribute VB_Name = "Module1"
Option Explicit

Public InSound As Boolean
Public ResourceSkinNum As Integer
Public WhichDates As Integer
Public NewDateTime As Date

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

Public Sub FormDrag(TheForm As Form)
   ReleaseCapture
   SendMessage TheForm.hWnd, &HA1, 2, 0&
End Sub

Public Sub MakeFormRounded(Obj As Object, Radius As Long)
    Dim hRgn As Long
    hRgn = CreateRoundRectRgn(0, 0, Obj.ScaleWidth, Obj.ScaleHeight, Obj.ScaleWidth / Radius, Obj.ScaleHeight / Radius)
    SetWindowRgn Obj.hWnd, hRgn, True
    Call DeleteObject(hRgn)
End Sub

Public Sub PlaySound(Flavor As Integer)
   Const SYNC = 1
   Dim temp As String
   If InSound Then
      If Flavor = 60 Then
         temp = "Hover.wav"
      ElseIf Flavor = 61 Then
         temp = "Clique.wav"
      End If
      sndPlaySound ByVal App.Path & "\" & temp, SYNC
   End If
End Sub
Public Sub SkinButtons(F1 As Object)
   Dim L4 As Long

   On Error GoTo ProcedureError
   For L4 = 0 To F1.Button.ubound
      'Assumes control name is "Button" throughout project
      Set F1.Button(L4).SkinUp = LoadResPicture(ResourceSkinNum, vbResBitmap)
      Set F1.Button(L4).SkinOver = LoadResPicture(ResourceSkinNum + 1, vbResBitmap)
      'Set F1.Button(L4).SkinDown = LoadResPicture(ResourceSkinNum + 2, vbResBitmap)
      F1.Button(L4).TransparentColor = &HFF00FF 'Magenta
   Next
ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox("Public Sub SkinButtons") = vbRetry Then Resume Next

End Sub
Public Sub OutlineControl(C As Control, Frm As Object)

   Dim X As Integer

   Const C1& = &H808080
   Const C2& = &HC0C0C0
   Const C3& = vbWhite

'Left border

For X = 1 To 3
   Frm.Line (C.Left - X, C.Top - 3)-(C.Left - X, C.Top + C.Height + 2), C1, BF
   Frm.Line (C.Left + C.Width + X - 1, C.Top - 2)-(C.Left + C.Width + X, C.Top + C.Height - 1), C3, BF
Next X

For X = 4 To 6
   Frm.Line (C.Left - X, C.Top - 6)-(C.Left - X, C.Top + C.Height + 5), C2, BF
   Frm.Line (C.Left + C.Width + X - 1, C.Top - 5)-(C.Left + C.Width + X, C.Top + C.Height + 5), C2, BF
   Frm.Line (C.Left - 5, C.Top - X)-(C.Left + C.Width + 5, C.Top - X), C2, BF
Next X

For X = 7 To 9
   Frm.Line (C.Left - X, C.Top - 7)-(C.Left - X, C.Top + C.Height + 8), C3, BF
   Frm.Line (C.Left + C.Width + X - 1, C.Top - 8)-(C.Left + C.Width + X, C.Top + C.Height + 7), C1, BF
Next X

'Top Border
Frm.Line (C.Left - 2, C.Top - 1)-(C.Left + C.Width, C.Top - 1), C1, BF
Frm.Line (C.Left - 2, C.Top - 2)-(C.Left + C.Width + 1, C.Top - 2), C1, BF
Frm.Line (C.Left - 2, C.Top - 3)-(C.Left + C.Width + 2, C.Top - 3), C1, BF

Frm.Line (C.Left - 9, C.Top - 7)-(C.Left + C.Width + 7, C.Top - 7), C3, BF
Frm.Line (C.Left - 9, C.Top - 8)-(C.Left + C.Width + 8, C.Top - 8), C3, BF
Frm.Line (C.Left - 9, C.Top - 9)-(C.Left + C.Width + 9, C.Top - 9), C3, BF

'Bottom border
Frm.Line (C.Left, C.Top + C.Height)-(C.Left + C.Width + 2, C.Top + C.Height), C3, BF
Frm.Line (C.Left - 1, C.Top + C.Height + 1)-(C.Left + C.Width + 2, C.Top + C.Height + 1), C3, BF
Frm.Line (C.Left - 2, C.Top + C.Height + 2)-(C.Left + C.Width + 2, C.Top + C.Height + 2), C3, BF

For X = 3 To 5
    Frm.Line (C.Left - 5, C.Top + C.Height + X)-(C.Left + C.Width + 5, C.Top + C.Height + X), C2, BF
Next X

'Where borders connect
Frm.Line (C.Left - 6, C.Top + C.Height + 6)-(C.Left + C.Width + 9, C.Top + C.Height + 6), C1, BF
Frm.Line (C.Left - 7, C.Top + C.Height + 7)-(C.Left + C.Width + 9, C.Top + C.Height + 7), C1, BF
Frm.Line (C.Left - 8, C.Top + C.Height + 8)-(C.Left + C.Width + 9, C.Top + C.Height + 8), C1, BF

End Sub
Public Function ErrMsgBox(Msg As String) As Integer
    ErrMsgBox = MsgBox("Error: " & Err.Number & ". " & Err.Description, vbRetryCancel + vbCritical, Msg)
End Function
