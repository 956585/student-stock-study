Attribute VB_Name = "AutoStockReport"
Option Explicit

' ==============================================================================
' ツール名：企業データ自動整形＆高配当ハイライトマクロ
' 概要　　：Webツールから出力した企業データ表を1秒で整形し、
' 　　　　　配当利回り3.0%以上の優良銘柄を自動で色分け・集計するVBAマクロ
' ==============================================================================

Sub FormatAndAnalyzeCorporateData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dividendYield As Double
    Dim highYieldCount As Long
    
    Set ws = ActiveSheet
    
    ' 画面更新を一時停止して高速化（1秒で処理を完了させる定番テクニック）
    Application.ScreenUpdating = False
    
    ' 最終行を自動取得
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "データが見つかりません。企業データCSVを開いてから実行してください。", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' 1. 見出し行（1行目）のスタイルを一括設定
    With ws.Range("A1:I1")
        .Interior.Color = RGB(30, 41, 59)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    
    ' 2. データ行をループし、高配当（利回り3.0%以上）かつ予算内で買える銘柄をハイライト
    highYieldCount = 0
    For i = 2 To lastRow
        dividendYield = Val(ws.Cells(i, 5).Value) ' E列：配当利回り(%)
        
        ' 数値フォーマット（通貨・カンマ区切り）の設定
        ws.Cells(i, 4).NumberFormatLocal = "#,##0"
        ws.Cells(i, 8).NumberFormatLocal = "#,##0"
        ws.Cells(i, 9).NumberFormatLocal = "#,##0"
        
        If dividendYield >= 3# And Val(ws.Cells(i, 7).Value) > 0 Then
            ' 利回り3%以上かつ1株以上購入可能な行を薄い緑色でハイライト
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 9)).Interior.Color = RGB(220, 252, 231)
            highYieldCount = highYieldCount + 1
        Else
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 9)).Interior.ColorIndex = xlNone
        End If
    Next i
    
    ' 3. 罫線と列幅の自動調整
    ws.Range("A1:I" & lastRow).Borders.LineStyle = xlContinuous
    ws.Columns("A:I").AutoFit
    
    Application.ScreenUpdating = True
    
    MsgBox "自動整形が完了しました！" & vbCrLf & _
           "・対象企業数：" & (lastRow - 1) & " 社" & vbCrLf & _
           "・高配当（3.0%以上）で予算内購入できる銘柄：" & highYieldCount & " 社", vbInformation, "処理完了"
End Sub
