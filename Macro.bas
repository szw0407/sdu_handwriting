Attribute VB_Name = "NewMacros"
Sub Random_fontsize_float()
'
' Random_fontsize_float ��
'
'
' �����޸� ��
'
    Dim R_Character As Range


    Dim FontSize(5)
    ' �����С��5��ֵ֮����в��������Ը�д
    FontSize(1) = "15.5"
    FontSize(2) = "17"
    FontSize(3) = "16.8"
    FontSize(4) = "16.2"
    FontSize(5) = "15.7"



    Dim FontName(1)
    '������������������֮����в������ɸ�д������Ҫ��֤ϵͳӵ����������
    FontName(1) = "��������"


    Dim ParagraphSpace(5)
    '�м�� ��һ������ֵ�о��ȷֲ����ɸ�д
    ParagraphSpace(1) = "28.1"
    ParagraphSpace(2) = "28.3"
    ParagraphSpace(3) = "27.5"
    ParagraphSpace(4) = "28"
    ParagraphSpace(5) = "28.9"

    '����ԭ��Ļ����������޸����д���

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
