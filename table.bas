Attribute VB_Name = "table"
Dim tmp_DT As ListObject
Dim last_r  As Long
Dim last_c As Long

Sub �e�[�u���̊�{����ꗗ()

    '=========================
    '�e�[�u�����쐬����
    '=========================
    
        Dim tmp_DT As ListObject
        Set tmp_DT = ActiveSheet.ListObjects.Add(Source:=Cells.CurrentRegion)
            tmp_DT.TableStyle = "" '�e�[�u���̐F���Ȃ���
            tmp_DT.Name = "tmp_DT"  '�e�[�u����������

    '=========================
    '�e�[�u����͈͂ɖ߂�
    '=========================
        ActiveSheet.ListObjects(1).Unlist
        
    '=========================
    '�ŏI�s�A�ŏI������߂�
    '=========================
        last_c = tmp_DT.Range.Rows.Count
        last_r = tmp_DT.Range.Columns.Count
        
    '=========================
    '�w�肵����������܂ޗ񐔂��擾����
    '=========================
        Dim target_str As String
            target_str = "�w�肵��������"
    
           c = tmp_DT.ListColumns(target_str).Index

End Sub



Sub �ʃu�b�N�̃e�[�u����\��t����()

        '�ϐ���ݒ�
            Dim copy_wb As Workbook
            Set copy_wb = Workbooks("Book2")
            
            Dim past_wb As Workbook
            Set past_wb = Workbooks("Book1")
            
            last_r = past_wb.Sheets(1).ListObjects(1).Range.Rows.Count

        '�ŏI�s�ɓ\��t��
            past_wb.Sheets(1).ListObjects(1).ListRows.Add
            copy_wb.Sheets(1).ListObjects(1).DataBodyRange.Copy
            past_wb.Sheets(1).ListObjects(1).ListColumns(1).Range(last_r + 1).PasteSpecial

End Sub
    
