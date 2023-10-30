Attribute VB_Name = "NewMacros"
Sub Random_fontsize_float()
'
' Random_fontsize_float 宏
'
'
' 字体修改 宏
'
    Dim R_Character As Range


    Dim FontSize(5)
    ' 字体大小在5个值之间进行波动，可以改写
    FontSize(1) = "15.5"
    FontSize(2) = "17"
    FontSize(3) = "16.8"
    FontSize(4) = "16.2"
    FontSize(5) = "15.7"



    Dim FontName(1)
    '字体名称在三种字体之间进行波动，可改写，但需要保证系统拥有下列字体
    FontName(1) = "萌妹子体"


    Dim ParagraphSpace(5)
    '行间距 在一定以下值中均等分布，可改写
    ParagraphSpace(1) = "28.1"
    ParagraphSpace(2) = "28.3"
    ParagraphSpace(3) = "27.5"
    ParagraphSpace(4) = "28"
    ParagraphSpace(5) = "28.9"

    '不懂原理的话，不建议修改下列代码

    For Each R_Character In ActiveDocument.Characters

        VBA.Randomize

        R_Character.Font.Name = FontName(1)

        R_Character.Font.Size = FontSize(Int(VBA.Rnd * 5) + 1)

        R_Character.Font.Position = Int(VBA.Rnd * 3) + 1

        R_Character.Font.Spacing = 0


    Next

    Application.ScreenUpdating = True



    For Each Cur_Paragraph In ActiveDocument.Paragraphs

        Cur_Paragraph.LineSpacing = ParagraphSpace(Int(VBA.Rnd * 5) + 1)


    Next
        Application.ScreenUpdating = True


End Sub
