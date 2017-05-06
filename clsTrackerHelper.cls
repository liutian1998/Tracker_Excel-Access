VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTrackerHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTrackerHelper
' Author    : tian.a.liu
' Date      : 4/10/2017 17:34
' Version   : 0.0.4
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private m_objConn As ADODB.Connection
Private m_objConfigSheet As Worksheet
Private m_strConfigTableName As String

Private Const SEARCH_TYPE_CHAR_TEXT_EQUEL As String = "$"
Private Const SEARCH_TYPE_CHAR_TEXT_CONTAIN As String = "|"
Private Const SEARCH_TYPE_CHAR_NUM_CONTAIN As String = "|"
Private Const SPLIT_CHAR_TEXT_CONTAIN As String = "|"
Private Const SPLIT_CHAR_NUM_CONTAIN As String = "|"

Public Property Get ConfigTableName() As String
    If m_strConfigTableName <> "" Then
        ConfigTableName = m_strConfigTableName
    End If
End Property

Public Property Let ConfigTableName(NewValue As String)
    m_strConfigTableName = NewValue
End Property

Private Sub Class_Terminate()
On Error Resume Next
    If Not m_objConn Is Nothing Then
        If m_objConn.State <> adStateClosed Then
            m_objConn.Close
        End If
        Set m_objConn = Nothing
    End If
End Sub

Public Property Get DBPassword() As String
    DBPassword = m_objConfigSheet.Range("B2").Value
End Property

'DB路径
Public Property Get DBPath() As String
    DBPath = m_objConfigSheet.Range("B1").Value
End Property

'Connection 对象
Public Property Get Connection() As ADODB.Connection
    Set Connection = m_objConn
End Property

Public Function GetOrderBySrting(ByVal strTableName As String) As String
    Dim objFind As Range
    
    Set objFind = m_objConfigSheet.Range("1:1").Find(What:=strTableName, LookAt:=True)
    
    If Not objFind Is Nothing Then
        If objFind.Offset(0, 2).Value <> "" Then
            GetOrderBySrting = objFind.Offset(0, 2).Value
        End If
    End If
End Function

Public Function GetOrderSrting(dicViews As Scripting.Dictionary) As String
    Dim i As Long
    Dim colOrder As Collection, colItem As Collection
    Dim strTmp As String
    
    Set colOrder = New Collection
    For i = 0 To dicViews.Count - 1
        If dicViews.Items(i).OrderNumber > 0 Then
            Set colItem = New Collection
            colItem.Add dicViews.Items(i).OrderNumber, "Number"
            colItem.Add dicViews.Items(i).OrderDirection, "Direction"
            colItem.Add dicViews.Items(i).FieldModel.FieldName, "FieldName"
            colOrder.Add colItem
        End If
    Next i
    
    ReDim aryIndex(1 To colOrder.Count) As Long
    
    For i = 1 To colOrder.Count
        aryIndex(i) = colOrder(i)(1)
    Next i
    
    Call nsort(aryIndex)
    
    strTmp = ""
    For i = UBound(aryIndex) To LBound(aryIndex) Step -1
        If colOrder(aryIndex(i))(3) <> "" Then
            strTmp = strTmp & colOrder(aryIndex(i))(3) & " " & colOrder(aryIndex(i))(2) & ","
        End If
    Next i
    If strTmp <> "" Then strTmp = Left(strTmp, Len(strTmp) - 1)
    GetOrderSrting = strTmp
End Function

'sort array desc
Private Sub nsort(ByRef arrtosort() As Long)
    Dim i As Long, j As Long, tmp As Long
    For i = LBound(arrtosort) To UBound(arrtosort)
        j = (UBound(arrtosort) - i) * Rnd + i
        If i <> j Then
            tmp = arrtosort(i)
            arrtosort(i) = arrtosort(j)
            arrtosort(j) = tmp
        End If
    Next i
End Sub

'Primary Key
Public Property Get PrimaryKeyField(ByRef dicViews As Dictionary) As clsTrackerFieldModel
    Dim objItem As clsTrackerFieldModel
    Dim i As Long
    For i = 0 To dicViews.Count - 1
        Set objItem = dicViews.Items(i)
        If objItem.FieldModel.PrimaryKey = True Then
            Set PrimaryKeyField = objItem
            Exit For
        End If
    Next i
End Property

Private Function ExsitSheet(strSheetName As String) As Boolean
    Dim objSht As Worksheet
    For Each objSht In ThisWorkbook.Worksheets
        If UCase(objSht.Name) = UCase(strSheetName) Then
            ExsitSheet = True
            Exit Function
        End If
    Next objSht
    
    ExsitSheet = False
End Function

'Connect DB and create connetion object
Public Function ConnectDB(strConfigSheetName As String, ByRef strOutput As String) As Boolean
    Dim strConn As String
    
On Error GoTo ConnectDB_Error
    
    If Not ExsitSheet(strConfigSheetName) Then
        strOutput = "Can not find database config sheet"
        ConnectDB = False
        Exit Function
    End If
    
    Set m_objConfigSheet = ThisWorkbook.Worksheets(strConfigSheetName)
    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBPath & _
        ";Jet OLEDB:Database Password=" & DBPassword & ""
        
    Set m_objConn = New ADODB.Connection
    m_objConn.ConnectionString = strConn
    m_objConn.Open
    
    ConnectDB = True
    On Error GoTo 0
    Exit Function
    
ConnectDB_Error:
    strOutput = "Error " & Err.Number & " (" & Err.Description & ") in procedure ConnectDB of Class Module clsTrackerDB"
    ConnectDB = False
End Function

'Get Table Name about join or else
Public Function GetTableNameEx(ByVal strTableName As String) As String
    Dim objFind As Range
    
    Set objFind = m_objConfigSheet.Range("1:1").Find(What:=strTableName, LookAt:=True)
    
    If objFind Is Nothing Then
        GetTableNameEx = strTableName
    Else
        If objFind.Offset(0, 1).Value <> "" Then
            GetTableNameEx = objFind.Offset(0, 1).Value
        Else
            GetTableNameEx = objFind.Value
        End If
    End If
    
End Function

'Query all data
Public Function QueryAll(ByVal strTableName As String, ByVal strTrackerSheetName As String, _
    ByVal strCopyAddress As String, ByRef strOutput As String, Optional lngTop As Long = 200) As Boolean
On Error GoTo QueryAll_Error

    Dim dicViews As Dictionary
    Dim strMsg As String
    
    If Not GetViewModels(strTableName, strTrackerSheetName, dicViews, strMsg) Then
        strOutput = strMsg
        QueryAll = False
        Exit Function
    End If
    
    Dim objRs As ADODB.Recordset
    Dim objTrackerSheet As Worksheet
    Dim strSQL As String
    
    Set objTrackerSheet = ThisWorkbook.Worksheets(strTrackerSheetName)
    
    strSQL = "SELECT " & IIf(lngTop > 0, " TOP " & lngTop, "") & vbCrLf
    strSQL = strSQL & GetSelectSQLStringOfQuery(dicViews) & vbCrLf
    strSQL = strSQL & "FROM " & GetRealTableName(strTableName) & vbCrLf
    strSQL = strSQL & "WHERE 1=1 " & GetConditionSQLStringOfQuery(objTrackerSheet, dicViews) & vbCrLf
    Dim strOrderBy As String
    
    strOrderBy = GetOrderSrting(dicViews)
    If strOrderBy <> "" Then
        strSQL = strSQL & "ORDER BY " & strOrderBy
    Else
        strSQL = strSQL & "ORDER BY " & PrimaryKeyField(dicViews).FieldModel.FieldName
    End If
    
    Set objRs = New ADODB.Recordset
    objRs.Open strSQL, m_objConn, adOpenKeyset, adLockReadOnly
    
    If objRs.RecordCount = 0 Then
        strOutput = "Result count is 0"
    Else
        strOutput = "Result count is " & objRs.RecordCount
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        
        objTrackerSheet.Range(strCopyAddress).CopyFromRecordset objRs
        SetContent m_objConfigSheet.Name, strTableName, strTrackerSheetName, Val(Split(strCopyAddress, "$")(2)) - 1
        objTrackerSheet.Range(strCopyAddress).Select
        
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    
    QueryAll = True
    Set objRs = Nothing
    
    On Error GoTo 0
    Exit Function

QueryAll_Error:
    strOutput = "Error " & Err.Number & " (" & Err.Description & ") in procedure QueryAll of Class Module clsTrackerHelper"
    QueryAll = False
End Function

'Get table name between table and tableex
Private Function GetRealTableName(strTableName As String) As String
    Dim strName As String
    
    If m_strConfigTableName <> "" Then
        GetRealTableName = strTableName
    Else
        strName = GetTableNameEx(strTableName)
        If strName <> "" Then
            GetRealTableName = strName
        Else
            GetRealTableName = strTableName
        End If
    End If
End Function

'get SQL string fo query data
Private Function GetSelectSQLStringOfQuery(ByRef dicViews As Dictionary) As String
    Dim i As Long
    Dim objViewItem As clsTrackerFieldModel
    Dim strSQL As String
    
    strSQL = ""
    For i = 1 To dicViews.Count
        Set objViewItem = dicViews.Items(i - 1)
        If objViewItem.Visible Then
            strSQL = strSQL & IIf(objViewItem.FieldModel.FieldName & "" <> "", objViewItem.FieldModel.FieldName, "NULL") & "," & vbCrLf
        End If
    Next i
    
    If strSQL <> "" Then If Right(strSQL, 2) = vbCrLf Then strSQL = Left(strSQL, Len(strSQL) - 2)
    If strSQL <> "" Then If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1)
    GetSelectSQLStringOfQuery = strSQL
End Function

'check data required
Public Function CheckRequired(ByRef objTrackerSheet As Worksheet, ByRef dicViews As Dictionary, _
    lngTrackFirstRow As Long, lngTrackLastRow As Long, ByRef strOutput As String) As Boolean
    
On Error GoTo CheckRequired_Error
    Dim i As Long, j As Long
    Dim objViewItem As clsTrackerFieldModel
    
    For i = lngTrackFirstRow To lngTrackLastRow
        For j = 1 To dicViews.Count
            Set objViewItem = dicViews.Items(j - 1)
            If objViewItem.Required Then
                If objTrackerSheet.Range(objViewItem.FieldColumeA & i).Value & "" = "" Then
                    strOutput = objViewItem.FieldColumeA & i & " is empty, can not submit"
                    CheckRequired = False
                    Exit Function
                End If
            End If
        Next j
    Next i
    
    CheckRequired = True
    On Error GoTo 0
    Exit Function

CheckRequired_Error:
    CheckRequired = False
    strOutput = "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckRequired of Class Module clsTrackerHelper"
End Function

'check data content length is avalible
Public Function CheckTextLength(ByRef objTrackerSheet As Worksheet, ByRef dicViews As Dictionary, _
    lngTrackFirstRow As Long, lngTrackLastRow As Long, ByRef strOutput As String) As Boolean
    
On Error GoTo CheckTextLength_Error
    Dim i As Long, j As Long
    Dim objViewItem As clsTrackerFieldModel
    
    For i = lngTrackFirstRow To lngTrackLastRow
        For j = 1 To dicViews.Count
            Set objViewItem = dicViews.Items(j - 1)
            If objViewItem.FieldModel.DataType = fldTypeText And objViewItem.EnableUpdate Then
                If Len(objTrackerSheet.Range(objViewItem.FieldColumeA & i).Value & "") > objViewItem.FieldModel.FieldSize Then
                    strOutput = objViewItem.FieldColumeA & i & "content length is over " & objViewItem.FieldModel.FieldSize & ", can not submit"
                    CheckTextLength = False
                    Exit Function
                End If
            End If
        Next j
    Next i
    
    CheckTextLength = True
    On Error GoTo 0
    Exit Function

CheckTextLength_Error:
    CheckTextLength = False
    strOutput = "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckTextLength of Class Module clsTrackerHelper"
End Function

'update all data
Public Function UpdateALL(strTableName As String, strTrackerSheetName As String, _
    lngTrackFirstRow As Long, lngTrackLastRow As Long, ByRef strOutput As String, Optional blnEnableInsert = False) As Boolean
    
On Error GoTo UpdateALL_Error
    Dim dicViews As Dictionary
    Dim strMsg As String
    Dim objTrackerSheet As Worksheet
    
    If m_objConn.State <> adStateOpen Then
        strOutput = "Connection is not open"
        UpdateALL = False
        Exit Function
    End If
    
    Set objTrackerSheet = ThisWorkbook.Worksheets(strTrackerSheetName)
    
    If Not GetViewModels(strTableName, strTrackerSheetName, dicViews, strMsg) Then
        strOutput = strMsg
        UpdateALL = False
        Exit Function
    End If
    
    Dim strSQL As String
    Dim i As Long
    Dim objViewItem As clsTrackerFieldModel
    
    If Not CheckRequired(objTrackerSheet, dicViews, lngTrackFirstRow, lngTrackLastRow, strOutput) Then
        UpdateALL = False
        Exit Function
    End If
    
    If Not CheckTextLength(objTrackerSheet, dicViews, lngTrackFirstRow, lngTrackLastRow, strOutput) Then
        UpdateALL = False
        Exit Function
    End If
    
    Dim objPrimKeyModel As clsTrackerFieldModel
    
    Set objPrimKeyModel = PrimaryKeyField(dicViews)
    m_objConn.BeginTrans
    
    For i = lngTrackFirstRow To lngTrackLastRow
        If objTrackerSheet.Range(objPrimKeyModel.FieldColumeA & i).Value & "" <> "" Then
            strSQL = "UPDATE " & strTableName & vbCrLf
            strSQL = strSQL & "SET " & GetConditionSQLStringOfUpdate(strTrackerSheetName, dicViews, i) & vbCrLf
            
            If objPrimKeyModel.FieldModel.DataType = fldTypeLong Then
                strSQL = strSQL & "Where " & objPrimKeyModel.FieldModel.FieldName & " =" & _
                    objTrackerSheet.Range(objPrimKeyModel.FieldColumeA & i).Value & vbCrLf
            Else
                strSQL = strSQL & "Where " & objPrimKeyModel.FieldModel.FieldName & " ='" & _
                    objTrackerSheet.Range(objPrimKeyModel.FieldColumeA & i).Value & "'" & vbCrLf
            End If
            m_objConn.Execute strSQL
        Else
            If blnEnableInsert Then
                strSQL = "INSERT INTO " & strTableName & vbCrLf
                strSQL = strSQL & GetInsertContentSQLString(objTrackerSheet, i, dicViews)
                m_objConn.Execute strSQL
            End If
        End If
    Next i
    
    m_objConn.CommitTrans
    
    UpdateALL = True
    On Error GoTo 0
    Exit Function
UpdateALL_Error:
    m_objConn.RollbackTrans
    strOutput = "Error " & Err.Number & " (" & Err.Description & ") in procedure UpdateALL of Class Module clsTrackerDB"
    UpdateALL = False
End Function

'get view model
Public Function GetViewModels(ByVal strTableName As String, ByVal strTrackerSheetName As String _
    , ByRef dicFields As Dictionary, ByRef strOutput As String, Optional ByVal strConfigSheetName As String = "") As Boolean
    Dim objTableNameRange As Range
    Dim objTrackerNameRange As Range
    
    If strConfigSheetName <> "" Then
        Set m_objConfigSheet = ThisWorkbook.Worksheets(strConfigSheetName)
    End If
    
    Set objTableNameRange = m_objConfigSheet.Range("1:1").Find(What:=strTableName, LookAt:=True)
    
    If objTableNameRange Is Nothing Then
        strOutput = "Can not find Table " & strTableName & " in config"
        GetViewModels = False
        Exit Function
    End If
    
    Set objTrackerNameRange = m_objConfigSheet.Range(objTableNameRange.Address & ":" & _
        m_objConfigSheet.Cells(1, m_objConfigSheet.Columns.Count).Address).Find _
        (What:=strTrackerSheetName, LookAt:=True)
    
    If objTrackerNameRange Is Nothing Then
        strOutput = "Can not find Tracker " & strTrackerSheetName & " in config"
        GetViewModels = False
        Exit Function
    End If
    
    Dim dicData As Dictionary
    Dim lngLastRow As Long
    Dim objTmpViewModel As clsTrackerFieldModel
    Dim i As Long, j As Long
    
    lngLastRow = m_objConfigSheet.Cells(m_objConfigSheet.Rows.Count, _
        objTrackerNameRange.Column).End(xlUp).Row
        
    Set dicData = New Dictionary
    j = 0
    For i = 3 To lngLastRow
        Set objTmpViewModel = New clsTrackerFieldModel
        If objTrackerNameRange.Offset(i - 1, 0).Value & "" <> "" Then
            j = j + 1
            objTmpViewModel.Bind j, objTrackerNameRange.Offset(i - 1, 0), objTableNameRange.Offset(i - 1, 0)
        Else
            objTmpViewModel.Bind 0, objTrackerNameRange.Offset(i - 1, 0), objTableNameRange.Offset(i - 1, 0)
        End If
        
        dicData.Add objTmpViewModel.FieldModel.Key, objTmpViewModel
    Next i
    
    Set dicFields = dicData
    GetViewModels = True
End Function

'get SQL string for set content when update
Private Function GetConditionSQLStringOfUpdate(ByVal strTrackerSheetName As String, ByRef dicViews As Dictionary, ByVal lngRow As Long)
    Dim i As Long
    Dim strSQL As String
    Dim objViewItem As clsTrackerFieldModel
    Dim objTrackerSheet As Worksheet
    
    Set objTrackerSheet = ThisWorkbook.Worksheets(strTrackerSheetName)
    
    For i = 1 To dicViews.Count
        Set objViewItem = dicViews.Items(i - 1)
        If Not objViewItem.EnableUpdate Then
            GoTo Continue
        End If
        
        If Trim(objViewItem.FieldModel.FieldName & "") = "" Then
            GoTo Continue
        End If
        
        Select Case objViewItem.FieldModel.DataType
        Case fldDataType.fldTypeText
            strSQL = strSQL & objViewItem.FieldModel.FieldName & "='" & _
                ConvertStringToSQL(objTrackerSheet.Range(objViewItem.FieldColumeA & lngRow).Value) & "'," & vbCrLf
        Case fldDataType.fldTypeLong
            If objTrackerSheet.Range(objViewItem.FieldColumeA & lngRow).Value & "" <> "" Then
                strSQL = strSQL & objViewItem.FieldModel.FieldName & "=" & _
                    ConvertStringToSQL(objTrackerSheet.Range(objViewItem.FieldColumeA & lngRow).Value) & "," & vbCrLf
            Else
                strSQL = strSQL & objViewItem.FieldModel.FieldName & "=NULL," & vbCrLf
            End If
        Case fldDataType.fldTypeDateTime
            If objTrackerSheet.Range(objViewItem.FieldColumeA & lngRow).Value & "" <> "" Then
                strSQL = strSQL & objViewItem.FieldModel.FieldName & "=#" & _
                    ConvertStringToSQL(objTrackerSheet.Range(objViewItem.FieldColumeA & lngRow).Value) & "#," & vbCrLf
            Else
                strSQL = strSQL & objViewItem.FieldModel.FieldName & "=NULL," & vbCrLf
            End If
        Case fldDataType.fldTypeCurrency
            If objTrackerSheet.Range(objViewItem.FieldColumeA & lngRow).Value & "" <> "" Then
                strSQL = strSQL & objViewItem.FieldModel.FieldName & "=" & _
                    ConvertStringToSQL(objTrackerSheet.Range(objViewItem.FieldColumeA & lngRow).Value) & "," & vbCrLf
            Else
                strSQL = strSQL & objViewItem.FieldModel.FieldName & "=NULL," & vbCrLf
            End If
        End Select
Continue:
    Next i
    
    If strSQL <> "" Then If Right(strSQL, 2) = vbCrLf Then strSQL = Left(strSQL, Len(strSQL) - 2)
    If strSQL <> "" Then If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1)
    
    GetConditionSQLStringOfUpdate = strSQL
End Function

'add all data
Public Function InsertALL(strTableName As String, strTrackerSheetName As String, _
    lngTrackFirstRow As Long, lngTrackLastRow As Long, ByRef strOutput As String) As Boolean
    Dim objRs As ADODB.Recordset
    Dim strSQL As String, strFields As String, strValues As String
    Dim dicViews As Dictionary
    Dim objViewItem As clsTrackerFieldModel
    Dim objTrackerSheet As Worksheet
    
On Error GoTo InsertALL_Error
    
    If m_objConn.State <> adStateOpen Then
        strOutput = "Connection is closed"
        InsertALL = False
        Exit Function
    End If
    
    Set objTrackerSheet = ThisWorkbook.Worksheets(strTrackerSheetName)
    
    If Not GetViewModels(strTableName, strTrackerSheetName, dicViews, strOutput) Then
        InsertALL = False
        Exit Function
    End If
    
    If Not CheckRequired(objTrackerSheet, dicViews, lngTrackFirstRow, lngTrackLastRow, strOutput) Then
        InsertALL = False
        Exit Function
    End If
    
    If Not CheckTextLength(objTrackerSheet, dicViews, lngTrackFirstRow, lngTrackLastRow, strOutput) Then
        InsertALL = False
        Exit Function
    End If
    
    
    Dim i As Long, j As Long
    
    m_objConn.BeginTrans
    For i = lngTrackFirstRow To lngTrackLastRow
        strSQL = "INSERT INTO " & strTableName & vbCrLf
        strSQL = strSQL & GetInsertContentSQLString(objTrackerSheet, i, dicViews)
        m_objConn.Execute strSQL
    Next i
    
    m_objConn.CommitTrans
    SetContent m_objConfigSheet.Name, strTableName, strTrackerSheetName, lngTrackFirstRow - 1, lngTrackLastRow - lngTrackLastRow + 1
    
    InsertALL = True
    
    On Error GoTo 0
    Exit Function

InsertALL_Error:
    m_objConn.RollbackTrans
    strOutput = "Error " & Err.Number & " (" & Err.Description & ") in procedure InsertALL of Class Module clsTrackerHelper"
    InsertALL = False
End Function

'get insert SQL string
Public Function GetInsertContentSQLString(ByRef objTrackerSheet As Worksheet, ByVal lngRow As Long, ByRef dicViews As Dictionary) As String
    Dim j As Long
    Dim strFields As String, strValues As String
    Dim objViewItem As clsTrackerFieldModel
    
    For j = 1 To dicViews.Count
        Set objViewItem = dicViews.Items(j - 1)
        
        If Not objViewItem.EnableUpdate Or Not objViewItem.Visible Then
            GoTo Continue
        End If
        
        If objViewItem.FieldModel.FieldName & "" = "" Then
            GoTo Continue
        End If
        
        strFields = strFields & Split(objViewItem.FieldModel.FieldName, ".")(UBound(Split(objViewItem.FieldModel.FieldName, "."))) & ","
        
        Select Case objViewItem.FieldModel.DataType
        Case fldDataType.fldTypeText
            strValues = strValues & "'" & ConvertStringToSQL(objTrackerSheet.Range(objViewItem.FieldColumeA & lngRow).Value) & "',"
        Case fldDataType.fldTypeLong
            If objTrackerSheet.Range(objViewItem.FieldColumeA & lngRow).Value & "" <> "" Then
                strValues = strValues & ConvertStringToSQL(objTrackerSheet.Range(objViewItem.FieldColumeA & lngRow).Value) & ","
            Else
                strValues = strValues & "NULL,"
            End If
        Case fldDataType.fldTypeDateTime
            If objTrackerSheet.Range(objViewItem.FieldColumeA & lngRow).Value & "" <> "" Then
                strValues = strValues & "#" & ConvertStringToSQL(objTrackerSheet.Range(objViewItem.FieldColumeA & lngRow).Value) & "#,"
            Else
                strValues = strValues & "NULL,"
            End If
        End Select
Continue:
    Next j
    
    If Right(strFields, 1) = "," Then strFields = Left(strFields, Len(strFields) - 1)
    If Right(strValues, 1) = "," Then strValues = Left(strValues, Len(strValues) - 1)
    
    GetInsertContentSQLString = "(" & strFields & ") Values" & vbCrLf & "(" & strValues & ")"
End Function

'delete data by primarykey
Public Function DeleteByPrimaryKey(strTableName As String, strTrackerSheetName As String, strKey As String, ByRef strOutput As String) As Boolean
    Dim con As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim strDBPath As String
    Dim Tablename As String
    Dim strSQL As String
    
On Error GoTo DeleteByPrimaryKey_Error
    
    Dim dicViews As Scripting.Dictionary
    
    If Not GetViewModels(strTableName, strTrackerSheetName, dicViews, strOutput) Then
        DeleteByPrimaryKey = False
        Exit Function
    End If
    
    If MsgBox("Are you sure delete this record: " & PrimaryKeyField(dicViews).FieldModel.Content & "=" & strKey, vbYesNo + vbExclamation, "Waring") <> vbYes Then
        DeleteByPrimaryKey = False
        Exit Function
    End If
    
    m_objConn.BeginTrans
    If m_objConn.State <> adStateOpen Then
        strOutput = "Connection is closed"
        DeleteByPrimaryKey = False
        Exit Function
    End If

    strSQL = "DELETE FROM " & strTableName & vbCrLf
    If PrimaryKeyField(dicViews).FieldModel.DataType = fldTypeLong Then
        strSQL = strSQL & "WHERE " & PrimaryKeyField(dicViews).FieldModel.FieldName & "=" & strKey
    Else
        strSQL = strSQL & "WHERE " & PrimaryKeyField(dicViews).FieldModel.FieldName & "='" & strKey & "'"
    End If
    
    Set rs = m_objConn.Execute(strSQL)
    m_objConn.CommitTrans
    
    DeleteByPrimaryKey = True
    On Error GoTo 0
    Exit Function

DeleteByPrimaryKey_Error:
    strOutput = "Error " & Err.Number & " (" & Err.Description & ") in procedure DeleteByPrimaryKey of Class Module clsTrackerHelper"
    m_objConn.RollbackTrans
    DeleteByPrimaryKey = False
End Function

'get SQL string of condition when query data
Private Function GetConditionSQLStringOfQuery(ByRef objTrackerSheet As Worksheet, ByRef dicViews As Dictionary) As String
    Dim strTmpContents As String
    Dim objViewItem As clsTrackerFieldModel
    Dim strSQL As String
    Dim aryQuery() As String
    Dim aryContain() As String, strContain As String
    Dim i As Long
    Dim strConditionAddress As String
    Dim dicConfitions As Dictionary
On Error GoTo GetConditionSQLStringOfQuery_Error
    
    For i = 1 To dicViews.Count
        Set objViewItem = dicViews.Items(i - 1)
        
        'Search condition is empty
        If objViewItem.QueryItems = "" Then
            GoTo Continue
        End If
        
        aryQuery = Split(objViewItem.QueryItems, ",")
        
        Select Case objViewItem.FieldModel.DataType
        Case fldDataType.fldTypeText
            If Left(objViewItem.QueryItems, 1) = SEARCH_TYPE_CHAR_TEXT_EQUEL Then  '固定值
                strConditionAddress = Mid(objViewItem.QueryItems, 2)
                If UCase(objTrackerSheet.Range(strConditionAddress).Value & "") = "NULL" Then
                    strSQL = strSQL & " AND " & objViewItem.FieldModel.FieldName & "&''=''" & vbCrLf
                ElseIf UCase(objTrackerSheet.Range(objViewItem.QueryItems).Value & "") = "NOT NULL" Then
                    strSQL = strSQL & " AND " & objViewItem.FieldModel.FieldName & _
                        "&''<>''" & vbCrLf
                Else
                    If objTrackerSheet.Range(objViewItem.QueryItems).Value & "" <> "" Then
                        strSQL = strSQL & " AND " & objViewItem.FieldModel.FieldName & "&''" & _
                            "='" & ConvertStringToSQL(objTrackerSheet.Range(objViewItem.QueryItems).Value) & "'" & vbCrLf
                    End If
                End If
            ElseIf Left(objViewItem.QueryItems, 1) = SEARCH_TYPE_CHAR_TEXT_CONTAIN Then  'In(*,*,*....)
                strConditionAddress = Mid(objViewItem.QueryItems, 2)
                If UCase(objTrackerSheet.Range(strConditionAddress).Value & "") = "NULL" Then
                    strSQL = strSQL & " AND " & objViewItem.FieldModel.FieldName & "&''=''" & vbCrLf
                ElseIf UCase(objTrackerSheet.Range(strConditionAddress).Value & "") = "NOT NULL" Then
                    strSQL = strSQL & " AND " & objViewItem.FieldModel.FieldName & _
                        "&''<>''" & vbCrLf
                Else
                    If objTrackerSheet.Range(strConditionAddress).Value & "" <> "" Then
                        aryContain = Split(ConvertStringToSQL(objTrackerSheet.Range(strConditionAddress).Value), SPLIT_CHAR_TEXT_CONTAIN)
                        strContain = Join(aryContain, "','")
                        strSQL = strSQL & " AND " & objViewItem.FieldModel.FieldName & _
                            " IN ('" & strContain & "')" & vbCrLf
                    End If
                End If
            Else
                If UCase(objTrackerSheet.Range(objViewItem.QueryItems).Value & "") = "NULL" Then
                    strSQL = strSQL & " AND " & objViewItem.FieldModel.FieldName & "=''" & vbCrLf
                ElseIf UCase(objTrackerSheet.Range(objViewItem.QueryItems).Value & "") = "NOT NULL" Then
                    strSQL = strSQL & " AND " & objViewItem.FieldModel.FieldName & _
                        "&''<>''" & vbCrLf
                Else
                    If objTrackerSheet.Range(objViewItem.QueryItems).Value & "" <> "" Then
                        strSQL = strSQL & " AND " & objViewItem.FieldModel.FieldName & _
                            " LIKE '%" & ConvertStringToSQL(objTrackerSheet.Range(objViewItem.QueryItems).Value) & "%'" & vbCrLf
                    End If
                End If
            End If
        Case fldDataType.fldTypeLong
            If Left(objViewItem.QueryItems, 1) = SEARCH_TYPE_CHAR_TEXT_CONTAIN Then
                strConditionAddress = Mid(objViewItem.QueryItems, 2)
                If objTrackerSheet.Range(strConditionAddress).Value & "" <> "" Then
                    aryContain = Split(objTrackerSheet.Range(strConditionAddress).Value, SPLIT_CHAR_NUM_CONTAIN)
                    strContain = Join(aryContain, ",")
                    strSQL = strSQL & " AND " & objViewItem.FieldModel.FieldName & _
                        "IN (" & objTrackerSheet.Range(strConditionAddress).Value & ")" & vbCrLf
                End If
            Else
                If objTrackerSheet.Range(objViewItem.QueryItems).Value & "" <> "" Then
                    strSQL = strSQL & " AND " & objViewItem.FieldModel.FieldName & _
                        "=" & objTrackerSheet.Range(objViewItem.QueryItems).Value & vbCrLf
                End If
            End If
        Case fldDataType.fldTypeDateTime
            If UBound(aryQuery) = 1 Then
                If objTrackerSheet.Range(aryQuery(0)).Value & "" <> "" And objTrackerSheet.Range(aryQuery(1)).Value & "" <> "" Then
                    strSQL = strSQL & " AND (" & objViewItem.FieldModel.FieldName & _
                        " BETWEEN #" & Format(objTrackerSheet.Range(aryQuery(0)).Value, "yyyy-MM-dd 00:00:00") & "#" & _
                        " AND #" & Format(objTrackerSheet.Range(aryQuery(1)).Value, "yyyy-MM-dd 23:23:59") & "#)" & vbCrLf
                ElseIf objTrackerSheet.Range(aryQuery(0)).Value & "" <> "" And objTrackerSheet.Range(aryQuery(1)).Value & "" = "" Then
                    strSQL = strSQL & " AND " & objViewItem.FieldModel.FieldName & " >=#" & _
                        Format(objTrackerSheet.Range(aryQuery(0)).Value, "yyyy-MM-dd 00:00:00") & "#" & vbCrLf
                ElseIf objTrackerSheet.Range(aryQuery(0)).Value & "" = "" And objTrackerSheet.Range(aryQuery(1)).Value & "" <> "" Then
                    strSQL = strSQL & " AND " & objViewItem.FieldModel.FieldName & " <=#" & _
                        Format(objTrackerSheet.Range(aryQuery(1)).Value, "yyyy-MM-dd 23:23:59") & "#" & vbCrLf
                End If
            Else
                If objTrackerSheet.Range(aryQuery(0)).Value & "" <> "" Then
                    strSQL = strSQL & " AND (" & objViewItem.FieldModel.FieldName & _
                        " BETWEEN #" & Format(objTrackerSheet.Range(objViewItem.QueryItems).Value, "yyyy-MM-dd 00:00:00") & "#" & _
                        " AND #" & Format(objTrackerSheet.Range(objViewItem.QueryItems).Value, "yyyy-MM-dd 23:59:59") & "#)" & vbCrLf
                End If
            End If
        Case fldDataType.fldTypeCurrency
            If objTrackerSheet.Range(objViewItem.QueryItems).Value & "" <> "" Then
                strSQL = strSQL & " AND " & objViewItem.FieldModel.FieldName & _
                    "=" & objTrackerSheet.Range(objViewItem.QueryItems).Value & vbCrLf
            End If
        End Select
Continue:
    Next i
    
    If strSQL <> "" Then If Right(strSQL, 2) = vbCrLf Then strSQL = Left(strSQL, Len(strSQL) - 2)
    
    GetConditionSQLStringOfQuery = strSQL
    On Error GoTo 0
    Exit Function

GetConditionSQLStringOfQuery_Error:
    GetConditionSQLStringOfQuery = ""
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetConditionSQLStringOfQuery of Class Module clsTrackerHelper"
End Function

'Convert SQL String of content to normal
Private Function ConvertStringToSQL(strMsg As String) As String
    strMsg = Replace(strMsg, "'", "''")
    strMsg = Replace(strMsg, "*", "**")
    strMsg = Replace(strMsg, "#", "##")
    strMsg = Replace(strMsg, "%", "%%")
'    If InStr(1, strMsg, "[") > 0 And InStrRev(strMsg, "]") > 0 Then
'        strMsg = Replace(strMsg, "[", "[[")
'        strMsg = Replace(strMsg, "]", "]]")
'    End If
'    strMsg = Replace(strMsg, "_", "__")
'    strMsg = Replace(strMsg, "?", "??")
    ConvertStringToSQL = strMsg
End Function

'export data to excel
Public Function ExportDataToExcel(ByVal strTableName As String, ByVal strTrackerSheetName As String, _
    ByVal strCopyAddress As String, ByVal strTargetFilePath As String, ByRef strOutput As String) As Boolean
    
    Dim objRs As ADODB.Recordset
    Dim strSQL As String
    Dim objTrackerSheet As Worksheet
    Dim dicViews As Dictionary
    Dim strMsg As String
           
    If Not GetViewModels(strTableName, strTrackerSheetName, dicViews, strMsg) Then
        strOutput = strMsg
        ExportDataToExcel = False
        Exit Function
    End If
           
    Set objTrackerSheet = ThisWorkbook.Worksheets(strTrackerSheetName)
        
    strSQL = "SELECT " & GetSelectSQLStringOfQuery(dicViews) & vbCrLf
    strSQL = strSQL & "FROM " & GetRealTableName(strTableName) & vbCrLf
    strSQL = strSQL & "WHERE 1=1 " & GetConditionSQLStringOfQuery(objTrackerSheet, dicViews) & vbCrLf
    strSQL = strSQL & "Order by " & PrimaryKeyField(dicViews).FieldModel.FieldName
    
    Set objRs = New ADODB.Recordset
    objRs.Open strSQL, m_objConn, adOpenKeyset, adLockReadOnly
    
    Application.ScreenUpdating = False
    
    If objRs.RecordCount = 0 Then
        strOutput = "Records count is 0"
        ExportDataToExcel = False
    Else
        Dim objSourcesht As Worksheet
        Dim i As Long
        Set objSourcesht = ActiveSheet
        
        Dim objWb As Workbook
        Dim objViewItem As clsTrackerFieldModel
        Set objWb = Application.Workbooks.Add
        
        
        For i = 1 To dicViews.Count
            Set objViewItem = dicViews.Items(i - 1)
            objWb.Worksheets(1).Cells(1, i).Value = objViewItem.FieldModel.Content
        Next i
        
        objWb.Worksheets(1).Range("A2").CopyFromRecordset objRs
        objWb.SaveAs strTargetFilePath
        objWb.Close False
        ExportDataToExcel = True
    End If
    
    Application.ScreenUpdating = True
End Function

'set table title and format
Public Sub SetTitle(strConfigSheetName As String, strTableName As String, _
    strTrackerSheetName As String, lngTitleRow As Long, _
    Optional ByVal lngMaxCount As Long = 200)
    Dim objTrackerSheet As Worksheet
    
    Dim objTmpHeadRange As Range
    Dim dicViews As Dictionary
    Dim strOutput As String
    
    Set m_objConfigSheet = ThisWorkbook.Worksheets(strConfigSheetName)
    
'    If UCase(m_objConfigSheet.Range("B2").Value) <> "Y" Then
'        Exit Sub
'    End IfRead
    
    Set objTrackerSheet = ThisWorkbook.Worksheets(strTrackerSheetName)
    
    If Not GetViewModels(strTableName, strTrackerSheetName, dicViews, strOutput) Then
        Exit Sub
    End If
     
    Dim i As Long
    Dim objViewItem As clsTrackerFieldModel
    
    objTrackerSheet.Range(lngTitleRow & ":" & objTrackerSheet.Rows.Count).Clear
    
    For i = 1 To dicViews.Count
        Set objViewItem = dicViews.Items(i - 1)
        If Not objViewItem.Visible Then
            GoTo Continue
        End If
        
        objTrackerSheet.Range(objViewItem.FieldColumeA & lngTitleRow).Value = objViewItem.FieldModel.Content
        'objTrackerSheet.Columns(objViewItem.FieldColume).Hidden = False
        objTrackerSheet.Columns(objViewItem.FieldColume).ColumnWidth = objViewItem.FieldModel.Width
        If objViewItem.EnableUpdate And objViewItem.Required Then
            m_objConfigSheet.Range("C3").Copy
            objTrackerSheet.Range(objViewItem.FieldColumeA & lngTitleRow).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
        ElseIf objViewItem.EnableUpdate Then
            m_objConfigSheet.Range("C2").Copy
            objTrackerSheet.Range(objViewItem.FieldColumeA & lngTitleRow).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
        Else
            m_objConfigSheet.Range("C1").Copy
            objTrackerSheet.Range(objViewItem.FieldColumeA & lngTitleRow).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
        End If
        
        objViewItem.FieldModel.FormatRange.Copy

        If objViewItem.FieldModel.FormatRange.Formula <> "" Then
            objTrackerSheet.Range(objViewItem.FieldColumeA & lngTitleRow + 1 & ":" & _
                objViewItem.FieldColumeA & lngTitleRow + lngMaxCount).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
                    SkipBlanks:=False, Transpose:=False
        Else
            objTrackerSheet.Range(objViewItem.FieldColumeA & lngTitleRow + 1 & ":" & _
                objViewItem.FieldColumeA & lngTitleRow + lngMaxCount).Formula = ""
        End If
        
        objTrackerSheet.Range(objViewItem.FieldColumeA & lngTitleRow + 1 & ":" & _
            objViewItem.FieldColumeA & lngTitleRow + lngMaxCount).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
                
        If objViewItem.EnableUpdate Then
            objTrackerSheet.Range(objViewItem.FieldColumeA & lngTitleRow + 1 & ":" & _
                objViewItem.FieldColumeA & lngTitleRow + lngMaxCount).PasteSpecial Paste:=xlPasteValidation, Operation:=xlNone, _
                    SkipBlanks:=False, Transpose:=False
        Else
             objTrackerSheet.Range(objViewItem.FieldColumeA & lngTitleRow + 1 & ":" & _
                objViewItem.FieldColumeA & lngTitleRow + lngMaxCount).Validation.Delete
        End If
        
        If Not objTrackerSheet.Range(objViewItem.FieldColumeA & lngTitleRow).Comment Is Nothing Then objTrackerSheet.Range(objViewItem.FieldColumeA & lngTitleRow).Comment.Delete
        If objViewItem.FieldModel.Description <> "" Then
            objTrackerSheet.Range(objViewItem.FieldColumeA & lngTitleRow).AddComment objViewItem.FieldModel.Description
        End If
        
Continue:
    Next i
    
    objTrackerSheet.Range("A1").Select
    
    Application.CutCopyMode = False
    Call ClearSearchCondition(strConfigSheetName, strTableName, strTrackerSheetName)
End Sub

'set table content area format
Public Sub SetContent(strConfigSheetName As String, strTableName As String, _
    strTrackerSheetName As String, lngTitleRow As Long, _
    Optional ByVal lngMaxCount As Long = 200)
    
    Dim objTrackerSheet As Worksheet
    Dim objTmpHeadRange As Range
    Dim dicViews As Dictionary
    Dim strOutput As String
    
    Set m_objConfigSheet = ThisWorkbook.Worksheets(strConfigSheetName)
    
'    If UCase(m_objConfigSheet.Range("B2").Value) <> "Y" Then
'        Exit Sub
'    End If
    
    Set objTrackerSheet = ThisWorkbook.Worksheets(strTrackerSheetName)
    
    If Not GetViewModels(strTableName, strTrackerSheetName, dicViews, strOutput) Then
        Exit Sub
    End If
     
    Dim i As Long, j As Long
    Dim objViewItem As clsTrackerFieldModel
    
    For i = 1 To dicViews.Count
        Set objViewItem = dicViews.Items(i - 1)
        If objViewItem.Visible Then

            objViewItem.FieldModel.FormatRange.Copy
            If objViewItem.FieldModel.FormatRange.Formula <> "" Then
                objTrackerSheet.Range(objViewItem.FieldColumeA & lngTitleRow + 1 & ":" & _
                    objViewItem.FieldColumeA & lngTitleRow + lngMaxCount).PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
            End If

            objTrackerSheet.Range(objViewItem.FieldColumeA & lngTitleRow + 1 & ":" & _
                objViewItem.FieldColumeA & lngTitleRow + lngMaxCount).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                    SkipBlanks:=False, Transpose:=False
            
            If objViewItem.FieldModel.DataType = fldTypeFilePath Then
                For j = lngTitleRow + 1 To lngMaxCount
                    If objTrackerSheet.Range(objViewItem.FieldColumeA & j).Value <> "" Then
                        objTrackerSheet.Range(objViewItem.FieldColumeA & j).Hyperlinks.Add _
                            Anchor:=objTrackerSheet.Range(objViewItem.FieldColumeA & j), _
                            Address:=objTrackerSheet.Range(objViewItem.FieldColumeA & j).Value, _
                            TextToDisplay:="Click View"
                    End If
                Next j
            End If
        End If
    Next i
    
    Application.CutCopyMode = False
End Sub

'clear sheet condition content
Public Sub ClearSearchCondition(strConfigSheetName As String, strTableName As String, strTrackerSheetName As String)
    Dim objTrackerSheet As Worksheet
    Dim objTmpHeadRange As Range
    Dim dicViews As Dictionary
    Dim strOutput As String
     
    Set m_objConfigSheet = ThisWorkbook.Worksheets(strConfigSheetName)
    Set objTrackerSheet = ThisWorkbook.Worksheets(strTrackerSheetName)
    
    If Not GetViewModels(strTableName, strTrackerSheetName, dicViews, strOutput) Then
        Exit Sub
    End If
    
    Dim i As Long, j As Long
    Dim objViewItem As clsTrackerFieldModel
    Dim aryQuery() As String

    For i = 1 To dicViews.Count
        Set objViewItem = dicViews.Items(i - 1)
        If objViewItem.QueryItems <> "" Then
            aryQuery = Split(objViewItem.QueryItems, ",")
            For j = LBound(aryQuery) To UBound(aryQuery)
                If Asc(UCase(Left(aryQuery(j), 1))) < 65 Or Asc(UCase(Left(aryQuery(j), 1))) > 90 Then
                    aryQuery(j) = Mid(aryQuery(j), 2)
                End If
                If UBound(aryQuery) < 1 Then objTrackerSheet.Range(aryQuery(j)).Offset(0, -1).Value = objViewItem.FieldModel.Content
                objTrackerSheet.Range(aryQuery(j)).ClearContents
            Next j
        End If
    Next i
End Sub
