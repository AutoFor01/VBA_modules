Attribute VB_Name = "table"
Dim tmp_DT As ListObject
Dim last_r  As Long
Dim last_c As Long

Sub テーブルの基本操作一覧()

    '=========================
    'テーブルを作成する
    '=========================
    
        Dim tmp_DT As ListObject
        Set tmp_DT = ActiveSheet.ListObjects.Add(Source:=Cells.CurrentRegion)
            tmp_DT.TableStyle = "" 'テーブルの色をなくす
            tmp_DT.Name = "tmp_DT"  'テーブル名をつける

    '=========================
    'テーブルを範囲に戻す
    '=========================
        ActiveSheet.ListObjects(1).Unlist
        
    '=========================
    '最終行、最終列を求める
    '=========================
        last_c = tmp_DT.Range.Rows.Count
        last_r = tmp_DT.Range.Columns.Count
        
    '=========================
    '指定した文字列を含む列数を取得する
    '=========================
        Dim target_str As String
            target_str = "指定した文字列"
    
           c = tmp_DT.ListColumns(target_str).Index

End Sub



Sub 別ブックのテーブルを貼り付ける()

        '変数を設定
            Dim copy_wb As Workbook
            Set copy_wb = Workbooks("Book2")
            
            Dim past_wb As Workbook
            Set past_wb = Workbooks("Book1")
            
            last_r = past_wb.Sheets(1).ListObjects(1).Range.Rows.Count

        '最終行に貼り付け
            past_wb.Sheets(1).ListObjects(1).ListRows.Add
            copy_wb.Sheets(1).ListObjects(1).DataBodyRange.Copy
            past_wb.Sheets(1).ListObjects(1).ListColumns(1).Range(last_r + 1).PasteSpecial

End Sub
    
