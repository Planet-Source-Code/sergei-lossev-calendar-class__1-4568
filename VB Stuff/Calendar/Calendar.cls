VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' *****************   clsCalendar   *******************
'
' Code by Sergei Lossev
' slossev@uwaterlo.ca
'
' I don't care how you decide to use this code as long as you use
' it for a good cause. :o)
'
'
' I tried to make the class structure as generic as possible. Some of the code
' in here may appear redundant (and may be it is), but it was included anyway
' for implementing future addons
'
' Sorry, guys, no comments. But comments take time, and time is money. :o)

Option Explicit

'****************************** C O N S T A N T S ******************************
Private Const m_iCELL_SPACING = 60

Private Const m_iCOLOR_SEL_RECT = vbRed

Private Const m_iCOLOR_NORMAL_BACKGROUND = vbWhite
Private Const m_iCOLOR_SELECTED_BACK = vbBlue

Private Const m_iCOLOR_NORMAL_FORE = vbBlack
Private Const m_iCOLOR_SELECTED_FORE = vbWhite

Private Const m_iHEADER_HEIGHT = 240


Private Const m_iCOLOR_LIGHT_GRAY = &H8000000F
Private Const m_iCOLOR_MEDIUM_GRAY = &H808080
Private Const m_iCOLOR_DARK_GRAY = &H404040

'********************************** T Y P E S **********************************
'****************************** V A R I A B L E S ******************************
Private m_iCellWidth As Integer
Private m_iCellHeight As Integer

Private WithEvents m_picContainer As PictureBox
Attribute m_picContainer.VB_VarHelpID = -1

Private m_iLeftMargin As Integer
Private m_iTopMargin As Integer

Private m_dtPrevClicked As Date

Private m_lHeaderForeColor As Long
Private m_lHeaderBackColor As Long
Private m_bHeaderBold As Boolean

Private m_iShownMonth As Integer
Private m_iShownYear As Integer

Private m_bMultiSelect As Boolean

Private m_colSelectedDates As clsDayInfoCol

Private m_colDays As clsDayInfoCol

'************************************ EVENTS ***********************************
Public Event DateClicked(ByVal Button As Integer, ByVal dtDateClicked As Date, iPos As Integer)


'************************************ A P I ************************************
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_CONTROL = &H11
Private Const VK_SHIFT = &H10

'****************************** M A I N   C O D E ******************************
Private Function KeyIsPressed(lKey As Long) As Boolean
    KeyIsPressed = (GetKeyState(lKey) And 128)
End Function

Public Property Let ShownMonth(iNewValue As Integer)
    m_iShownMonth = iNewValue
    RefreshMonthData
End Property

Public Property Let ShownYear(iNewValue As Integer)
    m_iShownYear = iNewValue
    RefreshMonthData
End Property

Public Property Get ShownMonth() As Integer
    ShownMonth = m_iShownMonth
End Property

Public Property Get ShownYear() As Integer
    ShownYear = m_iShownYear
End Property

Public Property Set PictureBox(picAny As PictureBox)
    Set m_picContainer = picAny
    m_picContainer.AutoRedraw = True
    m_picContainer.BackColor = m_iCOLOR_NORMAL_BACKGROUND
End Property

Private Sub Class_Initialize()
    Set m_colSelectedDates = New clsDayInfoCol
    m_iCellWidth = Screen.TwipsPerPixelX * 20
    m_iCellHeight = Screen.TwipsPerPixelY * 15
    m_iLeftMargin = 150
    m_iTopMargin = 360 + m_iCellHeight
    
    m_bMultiSelect = True
    
    m_lHeaderForeColor = vbBlack
    m_lHeaderBackColor = m_iCOLOR_LIGHT_GRAY
    m_bHeaderBold = False
     
    Set m_colDays = New clsDayInfoCol
    
    m_iShownMonth = Month(Now)
    m_iShownYear = Year(Now)
    
    RefreshMonthData
    
'    SelectedDate = Now

    
   
End Sub

Public Sub RefreshCalendar()
    Dim iPos As Integer
    Dim iCount As Integer
      
    If m_picContainer Is Nothing Then Exit Sub
      
    DrawMainHeader m_iShownMonth, m_iShownYear
    DrawWeeklyHeaders
 
    Dim iDayCount As Integer
    Dim dtTemp As Date
    Dim iStartPos As Integer
       
    dtTemp = CDate("1 " & MonthName(m_iShownMonth) & " " & m_iShownYear)
    iDayCount = DateDiff("d", dtTemp, DateAdd("m", 1, dtTemp))
    iStartPos = Weekday(dtTemp, vbSunday) - 1
        
    For iPos = 0 To iStartPos - 1
        DrawDay iPos, ""
    Next iPos
        
    Dim oDay As clsDayInfo
        
    For iCount = 1 To m_colDays.Count
        Set oDay = m_colDays.Item(iCount)
        iPos = Day(oDay.DateTime) + iStartPos - 1
        DrawDay iPos, iPos - iStartPos + 1, oDay.ForeColor, oDay.BackColor, oDay.Bold, oDay.RectColor, oDay.Selected
    Next iCount
            
    For iPos = iStartPos + iDayCount To 41
        DrawDay iPos, ""
    Next iPos
End Sub

Private Sub DrawMainHeader(iMonth As Integer, iYear As Integer)
    Dim iOldTopMargin As Integer
    'Dim ubutton As utypDay
    
    iOldTopMargin = m_iTopMargin
    
    'Filled rect
    m_picContainer.Line (0, 0 + 0 * Screen.TwipsPerPixelY)- _
            (m_picContainer.Width, _
            m_iHEADER_HEIGHT + Screen.TwipsPerPixelY), m_lHeaderBackColor, BF
            
    'Black rect
    m_picContainer.Line (-1 * Screen.TwipsPerPixelY, -1 * Screen.TwipsPerPixelY)- _
            (m_picContainer.Width - 5 * Screen.TwipsPerPixelY, _
            m_iHEADER_HEIGHT + Screen.TwipsPerPixelY), vbBlack, B
    
    'Top Horiz light
    m_picContainer.Line (0 * Screen.TwipsPerPixelY, 0 * Screen.TwipsPerPixelY)- _
            (m_picContainer.Width - 5 * Screen.TwipsPerPixelY, _
            0), vbWhite
            
    'Left Vert light
    m_picContainer.Line (0 * Screen.TwipsPerPixelY, 0 * Screen.TwipsPerPixelY)- _
            (0 * Screen.TwipsPerPixelY, _
            m_iHEADER_HEIGHT + 1 * Screen.TwipsPerPixelY), vbWhite
    
    'Bottom Horiz Dark
    m_picContainer.Line (0 * Screen.TwipsPerPixelY, m_iHEADER_HEIGHT - 0 * Screen.TwipsPerPixelY)- _
            (m_picContainer.Width - 6 * Screen.TwipsPerPixelY, _
            m_iHEADER_HEIGHT - 0 * Screen.TwipsPerPixelY), m_iCOLOR_MEDIUM_GRAY
            
    'Right Vert Dark
    m_picContainer.Line (m_picContainer.Width - 6 * Screen.TwipsPerPixelY, 1 * Screen.TwipsPerPixelY)- _
            (m_picContainer.Width - 6 * Screen.TwipsPerPixelY, _
            m_iHEADER_HEIGHT + 1 * Screen.TwipsPerPixelY), m_iCOLOR_MEDIUM_GRAY
    
    With m_picContainer
        
        Dim sTemp As String
        Dim iTextWidth As Integer
        Dim iTextHeight As Integer
        
        
        .ForeColor = m_lHeaderForeColor
        .FontBold = m_bHeaderBold
        
        sTemp = MonthName(iMonth, True) & " " & iYear
        
        iTextHeight = .TextHeight(sTemp)
        iTextWidth = .TextWidth(sTemp)
        
        .CurrentX = (m_picContainer.Width - iTextWidth) \ 2
        
        .CurrentY = (m_iHEADER_HEIGHT - iTextHeight) \ 2
        
        m_picContainer.Print sTemp
    End With
    
    m_iTopMargin = iOldTopMargin
End Sub

Private Sub DrawWeeklyHeaders()
    Dim iOldTopMargin As Integer
    
    iOldTopMargin = m_iTopMargin
    
    m_iTopMargin = m_iTopMargin - m_iCellHeight - m_iCELL_SPACING
    
    m_picContainer.Line (m_iLeftMargin, m_iTopMargin + m_iCellHeight + Screen.TwipsPerPixelY)- _
            (m_iLeftMargin + 7 * (m_iCellWidth + m_iCELL_SPACING) - m_iCELL_SPACING, _
            m_iTopMargin + m_iCellHeight + Screen.TwipsPerPixelY), 0
    
    
    DrawDay 0, "S", m_iCOLOR_NORMAL_FORE, RGB(255, 255, 255), True, m_iCOLOR_NORMAL_BACKGROUND, False
    DrawDay 1, "M", m_iCOLOR_NORMAL_FORE, RGB(255, 255, 255), True, m_iCOLOR_NORMAL_BACKGROUND, False
    DrawDay 2, "T", m_iCOLOR_NORMAL_FORE, RGB(255, 255, 255), True, m_iCOLOR_NORMAL_BACKGROUND, False
    DrawDay 3, "W", m_iCOLOR_NORMAL_FORE, RGB(255, 255, 255), True, m_iCOLOR_NORMAL_BACKGROUND, False
    DrawDay 4, "T", m_iCOLOR_NORMAL_FORE, RGB(255, 255, 255), True, m_iCOLOR_NORMAL_BACKGROUND, False
    DrawDay 5, "F", m_iCOLOR_NORMAL_FORE, RGB(255, 255, 255), True, m_iCOLOR_NORMAL_BACKGROUND, False
    DrawDay 6, "S", m_iCOLOR_NORMAL_FORE, RGB(255, 255, 255), True, m_iCOLOR_NORMAL_BACKGROUND, False
    
    m_iTopMargin = iOldTopMargin
End Sub

Private Function GetPositonFromXY(iX As Single, iY As Single) As Integer
    Dim iPosX As Single
    Dim iPosY As Single
    
    Dim iRow As Integer
    Dim iCol As Integer
    
    
    iPosX = iX - m_iLeftMargin
    iPosY = iY - m_iTopMargin
    
    If iPosX < 0 Or iPosY < 0 Then
        GetPositonFromXY = -1
        Exit Function
    End If
    
    iRow = iPosY \ (m_iCellHeight + m_iCELL_SPACING)
    
    If (iRow > 5) Or (iRow < 0) Then
        GetPositonFromXY = -1
        Exit Function
    End If
    
    iCol = iPosX \ (m_iCellWidth + m_iCELL_SPACING)
    
    If (iCol > 6) Or (iRow < 0) Then
        GetPositonFromXY = -1
        Exit Function
    End If
    
    GetPositonFromXY = iRow * 7 + iCol
    
End Function

Private Sub Class_Terminate()
    Set m_colDays = Nothing
    Set m_colSelectedDates = Nothing
End Sub

Private Sub m_picContainer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iPos As Integer
    iPos = GetPositonFromXY(X, Y)

    If iPos = -1 Or Button <> 1 Then Exit Sub

    DoEvents

    Dim iDayCount As Integer
    Dim dtTemp As Date
    Dim iStartPos As Integer
       
    dtTemp = CDate("1 " & MonthName(m_iShownMonth) & " " & m_iShownYear)
    iDayCount = DateDiff("d", dtTemp, DateAdd("m", 1, dtTemp))
    iStartPos = Weekday(dtTemp, vbSunday) - 1
        
    If (iPos < iStartPos) Or (iPos >= iStartPos + iDayCount) Then Exit Sub

    dtTemp = CDate(iPos - iStartPos + 1 & " " & MonthName(m_iShownMonth) & " " & m_iShownYear)
    
    Dim iCount As Integer
    
    
    If Not m_bMultiSelect Then
        RefreshMonthData
        m_colSelectedDates.Clear
        SelectDay dtTemp
    Else
        
        
        With m_colDays
        
            If Not KeyIsPressed(VK_CONTROL) And Not KeyIsPressed(VK_SHIFT) Then
                    RefreshMonthData
                    m_colSelectedDates.Clear
                    SelectDay dtTemp
                    
            ElseIf Not KeyIsPressed(VK_CONTROL) And KeyIsPressed(VK_SHIFT) Then
                    If m_colSelectedDates.Count = 0 Then
                        SelectDay dtTemp
                    Else
                        RefreshMonthData
                        If m_dtPrevClicked <= dtTemp Then
                            For iCount = 0 To DateDiff("d", m_dtPrevClicked, dtTemp)
                                SelectDay DateAdd("d", m_dtPrevClicked, iCount)
                            Next iCount
                        ElseIf m_dtPrevClicked > dtTemp Then
                            For iCount = 0 To DateDiff("d", dtTemp, m_dtPrevClicked)
                                SelectDay DateAdd("d", m_dtPrevClicked, -iCount)
                            Next iCount
                        End If
                    End If
                
            ElseIf KeyIsPressed(VK_CONTROL) And Not KeyIsPressed(VK_SHIFT) Then
                If m_colSelectedDates.Exists(FormatKey(dtTemp)) Then
                    UnSelectDay dtTemp
                Else
                    SelectDay dtTemp
                End If
            End If
        End With
        
    End If
    RefreshCalendar
    
    m_dtPrevClicked = dtTemp
    
    RaiseEvent DateClicked(Button, dtTemp, iPos)
    
End Sub

Private Sub DrawDay(iLocation As Integer, _
                    sLabel As String, _
                    Optional lForeColor As Long = vbBlack, _
                    Optional lBackColor As Long = vbWhite, _
                    Optional bBold As Boolean = False, _
                    Optional lSelRectColor As Long = vbWhite, _
                    Optional bSelect As Boolean = False)

    ' +++
    '  iLocation from 0 to 41
    ' +++
    If m_picContainer Is Nothing Then Exit Sub
    
    Dim iRow As Integer
    Dim iCol As Integer
    
    iRow = iLocation \ 7
    iCol = iLocation - iRow * 7
    
    Dim iCurX As Integer
    Dim iCurY As Integer
    
    Dim iTextWidth As Integer
    Dim iTextHeight As Integer
    
    With m_picContainer
        
        iCurX = m_iLeftMargin + (m_iCellWidth + m_iCELL_SPACING) * iCol
        iCurY = m_iTopMargin + (m_iCellHeight + m_iCELL_SPACING) * iRow
                
        m_picContainer.Line (iCurX, iCurY)-(iCurX + m_iCellWidth, iCurY + m_iCellHeight), vbWhite, BF
        m_picContainer.Line (iCurX + 30, iCurY + 30)-(iCurX - 30 + m_iCellWidth, iCurY - 30 + m_iCellHeight), lBackColor, BF
        
        m_picContainer.Line (iCurX, iCurY)-(iCurX + m_iCellWidth, iCurY + m_iCellHeight), lSelRectColor, B
        
        .FontBold = bBold
        .ForeColor = lForeColor
        
        iTextHeight = .TextHeight(sLabel)
        iTextWidth = .TextWidth(sLabel)
        
        .CurrentX = iCurX + Int((m_iCellWidth - iTextWidth) / 2)
        
        .CurrentY = iCurY + Int((m_iCellHeight - iTextHeight) / 2)
        
        m_picContainer.Print sLabel
    End With
    
    
End Sub

Public Property Get MonthDataCol() As clsDayInfoCol
    Set MonthDataCol = m_colDays
End Property

Public Property Set MonthDataCol(ByVal vNewValue As clsDayInfoCol)
    Set m_colDays = vNewValue
End Property


Public Sub RefreshMonthData()
    Dim dtTemp As Date
    Dim iDayCount As Integer
    Dim iCount As Integer
    
    dtTemp = CDate("1 " & MonthName(m_iShownMonth) & " " & m_iShownYear)
    iDayCount = DateDiff("d", dtTemp, DateAdd("m", 1, dtTemp))
    
    m_colDays.Clear
    m_colSelectedDates.Clear
    For iCount = 0 To iDayCount - 1
        m_colDays.Add DateAdd("d", iCount, dtTemp)
    Next iCount
End Sub

Public Property Get MultiSelect() As Boolean
    MultiSelect = m_bMultiSelect
End Property

Public Property Let MultiSelect(ByVal bNewValue As Boolean)
    m_bMultiSelect = bNewValue
End Property

Public Function FormatKey(dtDate As Date) As String
    FormatKey = Format(dtDate, "DD Mmm YYYY")
End Function

Public Property Get SelectedDays() As clsDayInfoCol
    Set SelectedDays = m_colSelectedDates
End Property

Public Property Set SelectedDays(ByVal oNewValue As clsDayInfoCol)
    Set m_colSelectedDates = oNewValue
End Property

Public Sub SelectDay(dtDate As Date)
    If Month(dtDate) <> m_iShownMonth Or Year(dtDate) <> m_iShownYear Then Exit Sub
    m_colDays.Item(FormatKey(dtDate)).Selected = True
    If Not m_colSelectedDates.Exists(FormatKey(dtDate)) Then
        m_colSelectedDates.Add dtDate
    End If
End Sub

Public Sub UnSelectDay(dtDate As Date)
    m_colDays.Item(FormatKey(dtDate)).Selected = False
    If m_colSelectedDates.Exists(FormatKey(dtDate)) Then
        m_colSelectedDates.Remove FormatKey(dtDate)
    End If
End Sub
