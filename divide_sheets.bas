Attribute VB_Name = "divide_sheets"
Sub divide_sheets()


    Dim header_r As Long '�w�b�_�[�s
    Dim divided_col As Long
    Dim original_ws As String
    Dim divided_str As String
    
    '��������������������������
    header_r = 3
    divided_col = 3
    original_ws = "Sheet1"
    '��������������������������

    Do While Cells(header_r + 1, divided_col) <> ""
        
        Sheets(original_ws).Activate
        divided_str = Cells(header_r + 1, divided_col)
        
        '���V�[�g�𕡐����A�t�B���^�[��������
            
            ActiveSheet.Copy after:=Sheets(Sheets.Count)
            With Rows(header_r)
                .AutoFilter field:=divided_col, Criteria1:="<>" & divided_str
                .CurrentRegion.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
                .AutoFilter
                Cells(1, 1).Select
                ActiveSheet.Name = divided_str
            End With
            
        '���V�[�g���番�������s���폜����
            Sheets(original_ws).Activate
            With Rows(header_r)
                .AutoFilter field:=divided_col, Criteria1:=divided_str
                .CurrentRegion.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
                .AutoFilter
            End With
    
    Loop

    '���̃V�[�g���폜
        Application.DisplayAlerts = False
        Sheets(original_ws).Delete
        Application.DisplayAlerts = True


End Sub
  
