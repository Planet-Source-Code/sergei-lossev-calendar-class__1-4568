VERSION 5.00
Begin VB.Form frmDropDownCalendar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2910
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2910
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   2910
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ComboBox cboYears 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   915
   End
   Begin VB.ComboBox cboMonths 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   1830
   End
   Begin VB.PictureBox picCalendar 
      Height          =   2355
      Left            =   60
      ScaleHeight     =   2295
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frmDropDownCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_oCalendar As clsCalendar
Attribute m_oCalendar.VB_VarHelpID = -1
Private m_dtSelectedDate As Date

Private Sub FillCombos()
    Dim iCount As Integer
    
    For iCount = 1990 To 2100
        cboYears.AddItem iCount
    Next iCount
    
    For iCount = 1 To 12
        cboMonths.AddItem MonthName(iCount)
        
    Next iCount
End Sub

Private Sub cboMonths_Click()
    ' change the current month on the calendar
    m_oCalendar.ShownMonth = cboMonths.ListIndex + 1
    m_oCalendar.RefreshCalendar
End Sub

Private Sub cboYears_Click()
    ' change the current year on the calendar
    m_oCalendar.ShownYear = cboYears.List(cboYears.ListIndex)
    m_oCalendar.RefreshCalendar
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    Set m_oCalendar = New clsCalendar
    Set m_oCalendar.PictureBox = picCalendar
    
    'Turning off multiselect
    m_oCalendar.MultiSelect = False
    
    FillCombos
    
    cboYears.ListIndex = Year(Now) - 1990
    cboMonths.ListIndex = Month(Now) - 1
    
    'selecting current day
    m_oCalendar.SelectDay Now

    'refreshing the calendar
    m_oCalendar.RefreshCalendar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_oCalendar = Nothing
End Sub

Private Sub m_oCalendar_DateClicked(ByVal Button As Integer, ByVal dtDateClicked As Date, iPos As Integer)
    SelectedDate = dtDateClicked
    DoEvents
'    Sleep 100
    Me.Hide
End Sub

Public Property Get SelectedDate() As Date
    SelectedDate = m_dtSelectedDate
End Property

Public Property Let SelectedDate(ByVal vNewValue As Date)
    m_dtSelectedDate = vNewValue
    
    m_oCalendar.ShownMonth = Month(m_dtSelectedDate)
    m_oCalendar.ShownYear = Year(m_dtSelectedDate)
    
    cboYears.ListIndex = Year(m_dtSelectedDate) - CInt(cboYears.List(0))
    cboMonths.ListIndex = Month(m_dtSelectedDate) - 1
    
    m_oCalendar.SelectedDays.Clear
    m_oCalendar.SelectDay vNewValue
    m_oCalendar.RefreshCalendar
End Property
