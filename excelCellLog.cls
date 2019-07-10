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



'初期化、ログの開始位置となるシートとRangeアドレスを設定する
Public Sub init(ByVal logSheetName As String, ByVal startRangeStr As String)
    
    If logSheetName = "" Then
        logSheetName = ThisWorkbook.Sheets(0).Name
    End If
    
    If startRangeStr = "" Then
        startRangeStr = "A1"
    End If
    
    Set logStartRange = ThisWorkbook.Sheets(logSheetName).Range(startRangeStr)
    
    Range(logStartRange.Row & ":" & logStartRange.Row).Interior.Color = RGB(192, 192, 192)
    
    logStartRange.Offset(0, 0) = "logNo"
    logStartRange.Offset(0, 1) = "Time"
    logStartRange.Offset(0, 2) = "ステータス"
    logStartRange.Offset(0, 3) = "本文"
'    logStartRange.Offset(0, 4) = "なんかその他"
    
End Sub



'ログ出力処理
Public Function writeLog(ByVal statasStr As String, ByVal logStr As String)
    
    Dim logNo As Long
    
    logNo = getNewLogEndRow
    
    logStartRange.Offset(logNo, 0) = logNo
    logStartRange.Offset(logNo, 1) = Now
    logStartRange.Offset(logNo, 2) = "ステータス"
    logStartRange.Offset(logNo, 3) = logStr
'    logStartRange.Offset(logNo, 4) = "なんかその他"


    Cells.Columns.AutoFit
    
End Function


'ログの最終行の一個下を取得
Private Function getNewLogEndRow() As Long
    
    Dim reIndexNo As Long
    
    Dim tmpEndRange As Range
    Set tmpEndRange = Sheets(logStartRange.Parent.Name).Cells(Rows.Count, logStartRange.Column).End(xlUp)
    
    
    If tmpEndRange.Row < logStartRange.Row Then
        reIndexNo = 0
    Else
        reIndexNo = tmpEndRange.Row - logStartRange.Row + 1
    End If
    
    getNewLogEndRow = reIndexNo

End Function



Public Sub clearLog()
    
    Call Sheets(logStartRange.Parent.Name).Range(logStartRange.Offset(1, 0).Row & ":" & Rows.Count).Delete
    
    Cells.Columns.AutoFit
    
End Sub
