VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTrackerFieldModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTrackerFieldModel
' Author    : tian.a.liu
' Date      : 3/31/2017 16:50
' Version   : 0.0.1
' Purpose   : Tracker field model
'---------------------------------------------------------------------------------------
Option Explicit

Private m_objVisibleRange As Range
Private m_boolVisible As Boolean
Private m_intIndex As Integer
Private m_boolEnableUpdate As Boolean
Private m_boolRequired As Boolean
Private m_strQueryItems As String
Private m_objFiledModel As clsDBFieldModel
Private m_strFieldColume As Long
Private m_strFieldColumeA As String
Private m_intOrderNumber As Integer
Private m_strOrderDirection As String

Public Property Get VisibleRange() As Range
    Set VisibleRange = m_objVisibleRange
End Property

Public Property Get Visible() As Boolean
    Visible = m_boolVisible
End Property

Public Property Get Index() As Integer
    Index = m_intIndex
End Property

Public Property Get EnableUpdate() As Boolean
    EnableUpdate = m_boolEnableUpdate
End Property

Public Property Get Required() As Boolean
    Required = m_boolRequired
End Property

Public Property Get QueryItems() As String
    QueryItems = m_strQueryItems
End Property

Public Property Get FieldModel() As clsDBFieldModel
    Set FieldModel = m_objFiledModel
End Property

Public Property Get FieldColume() As Long
     FieldColume = m_strFieldColume
End Property

Public Property Get FieldColumeA() As String
    FieldColumeA = m_strFieldColumeA
End Property

Public Property Get OrderNumber() As Integer
    OrderNumber = m_intOrderNumber
End Property

Public Property Get OrderDirection() As String
    OrderDirection = m_strOrderDirection
End Property

Public Sub Bind(ByVal Index As Long, ByVal objVisibleRange As Range, ByVal objFieldRange As Range)
    Set objVisibleRange = objVisibleRange
    
    m_strFieldColume = Index 'objFieldColumeRange.Value
    If Index > 0 Then m_strFieldColumeA = Split(objVisibleRange.Worksheet.Cells(1, m_strFieldColume).Address, "$")(1)
    
    m_boolVisible = IIf(Trim(objVisibleRange.Value & "" <> ""), True, False)
    m_intIndex = Val(objVisibleRange.Value & "")
    m_boolEnableUpdate = IIf(Trim(objVisibleRange.Offset(0, 1).Value & "" <> ""), True, False)
    m_boolRequired = IIf(Trim(objVisibleRange.Offset(0, 2).Value & "" <> ""), True, False)
    m_strQueryItems = objVisibleRange.Offset(0, 3).Value
    If objVisibleRange.Offset(0, 4).Value & "" <> "" Then
        Dim aryOrder() As String
        aryOrder = Split(objVisibleRange.Offset(0, 4).Value, ",")
        If UBound(aryOrder) > 0 Then
            m_intOrderNumber = Val(aryOrder(0))
            m_strOrderDirection = aryOrder(1)
        Else
            m_intOrderNumber = Val(objVisibleRange.Offset(0, 4).Value)
            m_strOrderDirection = "ASC"
        End If
    End If
    
    Set m_objFiledModel = New clsDBFieldModel
    Call m_objFiledModel.Bind(objFieldRange)
End Sub

