Attribute VB_Name = "table_paste_another_table"
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
