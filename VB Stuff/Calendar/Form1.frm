VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboDropDownCalendar 
      Height          =   315
      Left            =   765
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4320
      Width           =   1770
   End
   Begin VB.ListBox lstDates 
      Height          =   2985
      Left            =   3825
      TabIndex        =   3
      Top             =   720
      Width           =   2805
   End
   Begin VB.ComboBox cboYear 
      Height          =   315
      Left            =   2655
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   1005
   End
   Begin VB.ComboBox cboMonth 
      Height          =   315
      Left            =   765
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   1860
   End
   Begin VB.PictureBox picCalendar 
      Height          =   2625
      Left            =   765
      ScaleHeight     =   2565
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Click on the combobox to drop down a calendar"
      Height          =   465
      Left            =   2655
      TabIndex        =   6
      Top             =   4320
      Width           =   1950
   End
   Begin VB.Label Label1 
      Caption         =   "Use SHIFT and CTRL to multi select"
      Height          =   330
      Left            =   765
      TabIndex        =   4
      Top             =   3825
      Width           =   3930
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' make sure you declare the class using WithEvents
Private WithEvents m_oCalendar As clsCalendar
Attribute m_oCalendar.VB_VarHelpID = -1



Private Sub cboDropDownCalendar_DropDown()
    Dim sTemp As String
    sTemp = cboDropDownCalendar.List(0)
    
    GetDate cboDropDownCalendar, sTemp
    
    cboDropDownCalendar.Clear
    cboDropDownCalendar.AddItem sTemp
    cboDropDownCalendar.ListIndex = 0

End Sub

Private Sub cboMonth_Click()
    ' change the current month on the calendar
    m_oCalendar.ShownMonth = cboMonth.ListIndex + 1
    m_oCalendar.RefreshCalendar
End Sub

Private Sub cboYear_Click()
    ' change the current year on the calendar
    m_oCalendar.ShownYear = cboYear.List(cboYear.ListIndex)
    m_oCalendar.RefreshCalendar
End Sub

Private Sub Form_Load()
    'creating a calendar class
    Set m_oCalendar = New clsCalendar
    'attaching it to a pic box
    Set m_oCalendar.PictureBox = picCalendar
    
    ' Loading the dropdown form
    Load frmDropDownCalendar
    
    Dim iCount As Integer
    
    For iCount = 1990 To 2030
        cboYear.AddItem iCount
    Next iCount
    
    For iCount = 1 To 12
        cboMonth.AddItem MonthName(iCount)
    Next iCount
    
    cboYear.ListIndex = Year(Now) - 1990
    cboMonth.ListIndex = Month(Now) - 1
    
    'selecting current day
    m_oCalendar.SelectDay Now

    'refreshing the calendar
    m_oCalendar.RefreshCalendar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'cleaning up
    Set m_oCalendar = Nothing
    Unload frmDropDownCalendar
End Sub

Private Sub m_oCalendar_DateClicked(ByVal Button As Integer, ByVal dtDateClicked As Date, iPos As Integer)
    Dim iCount As Integer
    lstDates.Clear
    'put all the selected dates in a list box
    For iCount = 1 To m_oCalendar.SelectedDays.Count
        lstDates.AddItem Format(m_oCalendar.SelectedDays.Item(iCount).DateTime, "dd Mmm YYYY")
    Next iCount
End Sub
