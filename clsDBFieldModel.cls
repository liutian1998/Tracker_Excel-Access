VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBFieldModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsDBFieldModel
' Author    : tian.a.liu
' Date      : 4/1/2017 13:38
' Version   : 0.0.2
' Purpose   : Database field model
'---------------------------------------------------------------------------------------

Option Explicit
Private m_strKey As String

Private m_strFieldName As String
Private m_strContent As String
Private m_enmDataType As fldDataType
Private m_boolPrimaryKey As String
Private m_intFieldSize As String
Private m_strDescription As String
Private m_objFieldColumeRange As Range
Private m_intWidth As Integer
Private m_objFormatRange As Range

Public Enum fldDataType
    fldTypeText = 0
    fldTypeLong = 1
    fldTypeDateTime = 2
    fldTypeCurrency = 3
    fldTypeFilePath = 4
End Enum

Public Property Get Key() As String
    Key = m_strKey
End Property

Public Property Get FieldName() As String
    FieldName = m_strFieldName
End Property

Public Property Get Content() As String
    Content = m_strContent
End Property

Public Property Get DataType() As fldDataType
    DataType = m_enmDataType
End Property

Public Property Get PrimaryKey() As Boolean
    PrimaryKey = m_boolPrimaryKey
End Property

Public Property Get FieldSize() As Integer
    FieldSize = Val(m_intFieldSize)
End Property

Public Property Get Description() As String
    Description = m_strDescription
End Property

Public Property Get FieldColumeRange() As Range
    Set FieldColumeRange = m_objFieldColumeRange
End Property

Public Property Get Width() As Integer
    Width = m_intWidth
End Property

Public Property Get FormatRange() As Range
    Set FormatRange = m_objFormatRange
End Property

Public Sub Bind(ByVal objFieldColumeRange As Range)
    Set m_objFieldColumeRange = objFieldColumeRange
    
    m_strKey = objFieldColumeRange.Value
    m_strFieldName = objFieldColumeRange.Offset(0, 1).Value
    m_strContent = objFieldColumeRange.Offset(0, 2).Value
    m_intWidth = Val(objFieldColumeRange.Offset(0, 3).Value & "")
    
    Select Case UCase(objFieldColumeRange.Offset(0, 4).Value)
    Case "TEXT"
        m_enmDataType = fldTypeText
    Case "LONG"
        m_enmDataType = fldTypeLong
    Case "DATETIME"
        m_enmDataType = fldTypeDateTime
    Case "CURRENCY"
        m_enmDataType = fldTypeCurrency
    Case "FILEPATH"
        m_enmDataType = fldTypeFilePath
    Case Else
        m_enmDataType = fldTypeText
    End Select
    
    m_boolPrimaryKey = IIf(objFieldColumeRange.Offset(0, 5).Value & "" <> "", True, False)
    m_intFieldSize = objFieldColumeRange.Offset(0, 6).Value
    m_strDescription = objFieldColumeRange.Offset(0, 7).Value
    
    Set m_objFormatRange = objFieldColumeRange.Offset(0, 8)
End Sub
