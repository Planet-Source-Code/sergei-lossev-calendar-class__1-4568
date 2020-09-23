VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDayInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

  
Option Explicit
'****************************** C O N S T A N T S ******************************
'********************************** T Y P E S **********************************
'****************************** V A R I A B L E S ******************************
Private m_dtDateTime As Date
Private m_bSelected As Boolean
Private m_lBackColor As Long
Private m_lForeColor As Long
Private m_lRectColor As Long
Private m_bBold As Boolean

'************************************ A P I ************************************
'****************************** M A I N   C O D E ******************************
Public Property Get DateTime() As Date
    DateTime = m_dtDateTime
End Property

Public Property Let DateTime(ByVal bNewValue As Date)
    m_dtDateTime = bNewValue
End Property

Public Property Get Bold() As Boolean
    Bold = m_bBold
End Property

Public Property Let Bold(ByVal bNewValue As Boolean)
    m_bBold = bNewValue
End Property

Public Property Get Selected() As Boolean
    Selected = m_bSelected
End Property

Public Property Let Selected(ByVal bNewValue As Boolean)
    m_bSelected = bNewValue
    If m_bSelected Then
        BackColor = vbBlue
        ForeColor = vbWhite
        RectColor = vbBlue
    Else
        BackColor = vbWhite
        ForeColor = vbBlack
        RectColor = vbWhite
    End If
End Property
Public Property Get RectColor() As Long
    RectColor = m_lRectColor
End Property

Public Property Let RectColor(ByVal lNewValue As Long)
    m_lRectColor = lNewValue
End Property

Public Property Get ForeColor() As Long
    ForeColor = m_lForeColor
End Property

Public Property Let ForeColor(ByVal lNewValue As Long)
    m_lForeColor = lNewValue
End Property

Public Property Get BackColor() As Long
    BackColor = m_lBackColor
End Property

Public Property Let BackColor(ByVal lNewValue As Long)
    m_lBackColor = lNewValue
End Property
