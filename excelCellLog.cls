VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "excelCellLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ログの書き出し位置
Private logStartRange As Range


' 初期化処理
Private Sub Class_Initialize()
'    Debug.Print "initalized"
End Sub

' 終了時処理
Private Sub Class_Terminate()
'    Debug.Print "terminated"
End Sub


'初期化、ログの開始Rangeアドレスを設定する
'引数RangeはWorkbook名とsheet名も使うので注意してね
Public Sub init(ByVal argLogStartRange As Range)
    
    On Error GoTo errCatch
    
    
    If argLogStartRange Is Nothing Then
    
        Set argLogStartRange = ThisWorkbook.Sheets(1).Range("A1")
    
    End If
    
    Set logStartRange = argLogStartRange
    
    Range(logStartRange.Row & ":" & logStartRange.Row).Interior.Color = RGB(192, 192, 192)
    
    logStartRange.Offset(0, 0) = "logNo"
    logStartRange.Offset(0, 1) = "Time"
    logStartRange.Offset(0, 2) = "ステータス"
    logStartRange.Offset(0, 3) = "本文"
'    logStartRange.Offset(0, 4) = "なんかその他"

    Exit Sub
      
errCatch:

    Debug.Print (Now & " : " & "なんかエラー" & "  ==========")
    Debug.Print ("Err.Source  : " & Err.Source)
    Debug.Print ("Err.Number : " & Err.Number)
    Debug.Print ("Err.Description : " & Err.Description)
   
    
End Sub



'ログ出力処理
Public Function writeLog(ByVal statasStr As String, ByVal logStr As String)
    
    On Error GoTo errCatch
    
    
    If logStartRange Is Nothing Then
        
        Call MsgBox("処理中止" & vbCrLf & "ログ出力クラス初期化処理してないのでエラーになるよ")
        Exit Function
    End If
    
    Dim logNo As Long
    
    logNo = getNewLogEndRow
    
    logStartRange.Offset(logNo, 0) = logNo
    logStartRange.Offset(logNo, 1) = Now
    logStartRange.Offset(logNo, 2) = statasStr
    logStartRange.Offset(logNo, 3) = logStr
'    logStartRange.Offset(logNo, 4) = "なんかその他"
    
    Cells.Columns.AutoFit
    
    Exit Function
    
errCatch:
    
    Debug.Print (Now & " : " & "なんかエラー" & "  ==========")
    Debug.Print ("Err.Source  : " & Err.Source)
    Debug.Print ("Err.Number : " & Err.Number)
    Debug.Print ("Err.Description : " & Err.Description)
    
End Function



'ログのクリア処理
Public Sub clearLog()

    On Error GoTo errCatch
    
    Call Workbooks(logStartRange.Parent.Parent.Name).Sheets(logStartRange.Parent.Name).Range(logStartRange.Offset(1, 0).Row & ":" & Rows.Count).Delete
    
    Cells.Columns.AutoFit

errCatch:
    
    Debug.Print (Now & " : " & "なんかエラー" & "  ==========")
    Debug.Print ("Err.Source  : " & Err.Source)
    Debug.Print ("Err.Number : " & Err.Number)
    Debug.Print ("Err.Description : " & Err.Description)
    
    
    
End Sub






'ログの最終行の一個下を取得
Private Function getNewLogEndRow() As Long


    On Error GoTo errCatch
    
    
    Dim reIndexNo As Long
    
    Dim tmpEndRange As Range
    Set tmpEndRange = Workbooks(logStartRange.Parent.Parent.Name).Worksheets(logStartRange.Parent.Name).Cells(Rows.Count, logStartRange.Column).End(xlUp)
    
    
    If tmpEndRange.Row < logStartRange.Row Then
        reIndexNo = 0
    Else
        reIndexNo = tmpEndRange.Row - logStartRange.Row + 1
    End If
    
    getNewLogEndRow = reIndexNo



errCatch:
    
    Debug.Print (Now & " : " & "なんかエラー" & "  ==========")
    Debug.Print ("Err.Source  : " & Err.Source)
    Debug.Print ("Err.Number : " & Err.Number)
    Debug.Print ("Err.Description : " & Err.Description)
    
End Function



