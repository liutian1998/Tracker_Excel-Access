Attribute VB_Name = "modCommon"
'---------------------------------------------------------------------------------------
' Module    : modCommon
' Author    : tian.a.liu
' Date      : 3/31/2017 16:54
' Version   : 0.0.1
' Purpose   : Common module
'---------------------------------------------------------------------------------------

Option Explicit

Public Const GROUP_TYPE_PA_GROUPA As String = "PA_GROUPA"
Public Const GROUP_TYPE_PA_GROUPB As String = "PA_GROUPB"
Public Const GROUP_TYPE_PA_GROUPC As String = "PA_GROUPC"
Public Const GROUP_TYPE_PY As String = "PY"
Public Const GROUP_TYPE_ESC As String = "ESC"

Public Const ALERT_COLOR_LOW As Long = vbYellow
Public Const ALERT_COLOR_HIGHT As Long = vbRed

Public Const MESSAGE_NOT_AUTHORITY As String = "You don not have authority"

Public Const RANGE_ADDRESS_PROCESS_REQUESTNO = "G5"
Public Const RANGE_ADDRESS_PROCESS_CASEOWNER = "D5"

Public Const TASK_STATUS_OPEN = "Open"
Public Const TASK_STATUS_CLOSED = "Closed"
Public Const TASK_STATUS_ONHOLD = "On hold"
Public Const TASK_STATUS_REJECTED = "Rejected"
Public Const TASK_STATUS_REWORK = "Rework"

Public Const CASE_STATUS_OPEN = "Open"
Public Const CASE_STATUS_CLOSED = "Closed"
Public Const CASE_STATUS_ONHOLD = "On hold"
Public Const CASE_STATUS_REJECTED = "Rejected"
Public Const CASE_STATUS_REWORK = "Rework"

Public Const QC_RESULT_PASS As String = "Pass"
Public Const QC_RESULT_FAILED As String = "Failed"

Public Const NAME_LIST_ROLE As String = "ROLE_LIST"
Public Const NAME_LIST_USER_QC As String = "QC_USER_LIST"
Public Const NAME_LIST_USER_TL As String = "TL_USER_LIST"
Public Const NAME_LIST_USER_MEMBER As String = "MEMBER_USER_LIST"
Public Const NAME_LIST_REQUEST As String = "REQUEST_TYPE"
Public Const NAME_LIST_PROCESS As String = "PROCESS_LIST"
Public Const NAME_LIST_PROCESS_TABLE As String = "PROCESS_LIST_TABLE"
Public Const NAME_LIST_SUBPROCESS As String = "SUBPROCESS_LIST"
Public Const NAME_LIST_SUBPROCESS_PROCESS As String = "SUBPROCESS_LIST_PROCESS"
Public Const NAME_LIST_PSA As String = "PSA_LIST"
Public Const NAME_LIST_COMPANYCODE As String = "CompanyCode_LIST"
Public Const NAME_LIST_TASK_STATUS As String = "Task_Status"
Public Const NAME_LIST_CASE_STATUS As String = "Case_Status"

Public Function PickFilePath() As String
    Dim objDlg As FileDialog
    
    Set objDlg = Application.FileDialog(msoFileDialogFilePicker)
    objDlg.Show
    If objDlg.SelectedItems.Count > 0 Then
        PickFilePath = objDlg.SelectedItems(1)
    End If
End Function

Public Function PickSaveAs() As String
    Dim objDlg As FileDialog
    
    Set objDlg = Application.FileDialog(msoFileDialogSaveAs)
    objDlg.Show
    If objDlg.SelectedItems.Count > 0 Then
        PickSaveAs = objDlg.SelectedItems(1)
    End If
    
End Function

Public Sub SetAllBorders(Target As Range)
    Target.Borders(xlDiagonalDown).LineStyle = xlNone
    Target.Borders(xlDiagonalUp).LineStyle = xlNone
    With Target.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Target.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Target.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Target.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Target.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Target.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Public Sub CheckToolVersion()
    Dim strFileName As String
    Dim strVersion As String
    Dim strMyVersion As String
    Dim n As Long
    
On Error GoTo Run_Error
    If shtBasicConfig.ToolFolderPath = "" Then
        Exit Sub
    End If
    
    If Dir(shtBasicConfig.ToolFolderPath, vbDirectory) = "" Then
        Exit Sub
    End If
    
    strMyVersion = Mid(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "V") + 1, InStrRev(ThisWorkbook.Name, ".") - InStrRev(ThisWorkbook.Name, "V") - 1)
    strFileName = Dir(shtBasicConfig.ToolFolderPath & "\*.xls*", vbDirectory)
    Do While strFileName <> ""
        If strFileName <> "." And strFileName <> ".." Then
            If Left(strFileName, InStrRev(strFileName, "V")) = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, "V")) Then
                strVersion = Mid(strFileName, InStrRev(strFileName, "V") + 1, InStrRev(strFileName, ".") - InStrRev(strFileName, "V") - 1)
                If Val(Split(strVersion, ".")(0)) > Val(Split(strMyVersion, ".")(0)) Then
                    If MsgBox("Please update tools!" & vbCrLf & "There is new version:" & strFileName & vbCrLf & "Do you want to open the folder?", vbYesNo + vbExclamation, "Warning") = vbYes Then
                        Shell "Explorer.exe " & shtBasicConfig.ToolFolderPath, vbNormalFocus
                        Exit Do
                    End If
                ElseIf Val(Split(strVersion, ".")(0)) = Val(Split(strMyVersion, ".")(0)) Then
                    If Val(Split(strVersion, ".")(1)) > Val(Split(strMyVersion, ".")(1)) Then
                        If MsgBox("Please update tools!" & vbCrLf & "There is new version:" & strFileName & vbCrLf & "Do you want to open the folder?", vbYesNo + vbExclamation, "Warning") = vbYes Then
                            Shell "Explorer.exe " & shtBasicConfig.ToolFolderPath, vbNormalFocus
                            Exit Do
                        End If
                    ElseIf Val(Split(strVersion, ".")(1)) = Val(Split(strMyVersion, ".")(1)) Then
                        If Val(Split(strVersion, ".")(2)) > Val(Split(strMyVersion, ".")(2)) Then
                            If MsgBox("Please update tools!" & vbCrLf & "There is new version:" & strFileName & vbCrLf & "Do you want to open the folder?", vbYesNo + vbExclamation, "Warning") = vbYes Then
                                Shell "Explorer.exe " & shtBasicConfig.ToolFolderPath, vbNormalFocus
                                Exit Do
                            End If
                        End If
                    End If
                End If
            End If
        End If
        strFileName = Dir
    Loop
    
    Exit Sub
Run_Error:
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CheckToolVersion of Module modCommon", vbCritical, "Error"
End Sub

Public Function ConvertRangeToMultiSelectString(objTarget As Range) As String
    Dim objItem As Range
    Dim strTmp As String
    
    
    For Each objItem In objTarget
        strTmp = strTmp & objItem.Value & "|"
    Next objItem
    
    If strTmp <> "" Then If Right(strTmp, 1) = "|" Then strTmp = Left(strTmp, Len(strTmp) - 1)
    ConvertRangeToMultiSelectString = strTmp
End Function

Public Function FindWorkbookName(strName As String) As Name
    Dim i As Long
    
    For i = 1 To ThisWorkbook.Names.Count
        If UCase(ThisWorkbook.Names.Item(i).Name) = UCase(strName) Then
            Set FindWorkbookName = ThisWorkbook.Names.Item(i)
            Exit For
        End If
    Next i
End Function

'---------------------------------------------------------------------------------------
' Procedure : ViewSheet
' Author    : tian.a.liu
' Date      : 3/31/2017 16:53
' Return    :
' Purpose   : ViewSheet
'---------------------------------------------------------------------------------------
Public Sub ViewSheet(ByVal strSheetName As String)
    Dim i As Long

    'Application.EnableEvents = False
    Application.ScreenUpdating = False

    If strSheetName = "" Then strSheetName = "Main"
    
    ThisWorkbook.Worksheets(strSheetName).Visible = xlSheetVisible
    For i = 1 To Worksheets.Count
        If ThisWorkbook.Worksheets(i).Name <> strSheetName Then
            ThisWorkbook.Worksheets(i).Visible = xlSheetVeryHidden
        End If
    Next i
    
    ThisWorkbook.Activate
    ThisWorkbook.Worksheets(strSheetName).Select
    ThisWorkbook.Worksheets(strSheetName).Activate
    
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    
    Application.ScreenUpdating = True
    'Application.EnableEvents = True
End Sub

Public Function GetFilePathInFolderByFileName(strFolder As String, strName As String, Optional strType As String = "") As String
    Dim objItemSys As New Scripting.FileSystemObject
    Dim objItem As Scripting.File, objFolder As Scripting.Folder
    Dim i As Long
    
    Set objFolder = objItemSys.GetFolder(strFolder)
    
    For Each objItem In objFolder.Files
        If InStrRev(objItem.Name, strName) > 0 Then
            If strType <> "" Then
                If UCase(Right(objItem.Name, 3)) = UCase(strType) Then
                    GetFilePathInFolderByFileName = objItem.Path
                End If
            Else
                GetFilePathInFolderByFileName = objItem.Path
                Exit Function
            End If
        End If
    Next objItem
    
End Function

Public Function GetFolderPathInFolderByFileName(strFolder As String, strName As String) As String
    Dim objItemSys As New Scripting.FileSystemObject
    Dim objItem As Scripting.Folder, objFolder As Scripting.Folder
    Dim i As Long
    
    Set objFolder = objItemSys.GetFolder(strFolder)
    
    For Each objItem In objFolder.SubFolders
        If InStrRev(objItem.Name, strName) > 0 Then
            GetFolderPathInFolderByFileName = objItem.Path
            Exit Function
        End If
    Next objItem
    
End Function
