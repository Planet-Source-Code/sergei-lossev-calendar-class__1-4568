Attribute VB_Name = "basDropDownCalendar"
Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Sub GetDate(oWhere As Object, ByRef sDateValue As String)
    Dim uPoint As POINTAPI
    Dim a As Long

    uPoint.X = oWhere.Left / Screen.TwipsPerPixelX
    uPoint.Y = (oWhere.Top + oWhere.Height) / Screen.TwipsPerPixelY

    a = ClientToScreen(oWhere.Parent.hwnd, uPoint)
    
    If IsDate(sDateValue) Then
        frmDropDownCalendar.SelectedDate = CDate(sDateValue)
    Else
        frmDropDownCalendar.SelectedDate = Now
    End If
        
    frmDropDownCalendar.Left = uPoint.X * Screen.TwipsPerPixelX
    frmDropDownCalendar.Top = uPoint.Y * Screen.TwipsPerPixelY

    frmDropDownCalendar.Show vbModal

    DoEvents

    sDateValue = Format(frmDropDownCalendar.SelectedDate, "DD Mmm YYYY")
End Sub

