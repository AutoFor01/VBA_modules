    '====================
    '現在のパスを取得する
    '====================

        '---
        Dim fso as New FileSystemObject

        Dim ParentFolderPath as string  
            ParentFolderPath = fso.GetFolder(ThisWorkbook.Path).ParentFolder.Pat