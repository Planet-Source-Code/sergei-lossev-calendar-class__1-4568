VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDayInfoCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsDATAJobTime"
Attribute VB_Ext_KEY = "Member0" ,"clsDATAJobTime"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable to hold collection
Private m_colDays As Collection

Public Function Add(dtDateTime As Date, _
                    Optional bSelected As Boolean = False, _
                    Optional lBackColor As Long = vbWhite, _
                    Optional lForeColor As Long = vbBlack, _
                    Optional lRectColor As Long = vbWhite, _
                    Optional bBold As Boolean = False) As clsDayInfo
                    'Optional sKey As String) As clsDayInfo
    'create a new object
    Dim oDay As clsDayInfo
    Set oDay = New clsDayInfo


    'set the properties passed into the method
    With oDay
        .DateTime = dtDateTime
        .Selected = bSelected
        .BackColor = lBackColor
        .ForeColor = lForeColor
        .RectColor = lRectColor
        .Bold = bBold
        
        If bSelected Then
            .BackColor = vbBlue
            .ForeColor = vbWhite
            '.RectColor = vbBlue
        End If
        
    End With
    
    
    
    'If Len(sKey) = 0 Then
    '    m_colDays.Add oDay
    'Else
        m_colDays.Add oDay, FormatKey(dtDateTime)  'sKey
    'End If

    'return the object created
    Set Add = oDay
    Set oDay = Nothing

End Function

Public Property Get Item(vIndexKey As Variant) As clsDayInfo
Attribute Item.VB_UserMemId = 0
  Set Item = m_colDays(vIndexKey)
End Property

Public Property Get Count() As Long
    Count = m_colDays.Count
End Property

Public Sub Remove(vIndexKey As Variant)
    m_colDays.Remove vIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = m_colDays.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set m_colDays = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set m_colDays = Nothing
End Sub

Public Function Exists(vKey As Variant) As Boolean
    Dim oTemp As clsDayInfo
    On Error Resume Next
    Set oTemp = m_colDays.Item(vKey)
    
    Exists = (Err = 0)
End Function

Public Sub Clear()
    Set m_colDays = New Collection
End Sub

Public Function FormatKey(dtDate As Date) As String
    FormatKey = Format(dtDate, "DD Mmm YYYY")
End Function
