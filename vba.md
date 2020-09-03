Sub 手写字体()
'
' 宏1 宏
'
'

Dim R_CHARACTER As Range


Dim FONTSIZE(5)
FONTSIZE(1) = "16"
FONTSIZE(2) = "18"
FONTSIZE(3) = "17"
FONTSIZE(4) = "15"
FONTSIZE(5) = "15"

Dim FONTNAME(1)
FONTNAME(1) = "祝宣淋手写"

Dim PARAGRAPHSPACE(5)
PARAGRAPHSPACE(1) = "22"
PARAGRAPHSPACE(2) = "23"
PARAGRAPHSPACE(3) = "20"
PARAGRAPHSPACE(4) = "18"
PARAGRAPHSPACE(5) = "21"


For Each R_CHARACTER In ActiveDocument.Characters
    VBA.Randomize
    R_CHARACTER.Font.Name = FONTNAME(1)
    
    R_CHARACTER.Font.Size = FONTSIZE(Int(VBA.Rnd * 3) + 1)
    
    R_CHARACTER.Font.Position = Int(VBA.Rnd * 3) + 1
     
    R_CHARACTER.Font.Spacing = 0

Next

Application.ScreenUpdating = True

For Each cur_paragraph In ActiveDocument.Paragraphs
    cur_paragraph.LineSpacing = PARAGRAPHSPACE(Int(VBA.Rnd * 5) + 1)
    
Next
Application.ScreenUpdating = True

End Sub



