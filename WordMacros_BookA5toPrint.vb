Sub ЗмінитиРозмірМаржинів()
    ' Встановлення розміру на A5 (148 x 210 мм)
    ActiveDocument.PageSetup.PaperSize = wdPaperA5
    
    ' Вирівнювання сторінок
    ActiveDocument.Sections(1).PageSetup.SectionStart = wdSectionContinuous
    ActiveDocument.Sections(1).PageSetup.OddAndEvenPagesHeaderFooter = True

    With ActiveDocument.PageSetup
        .MirrorMargins = True
        .LeftMargin = (22 / 25.4) * 72
        .RightMargin = (17 / 25.4) * 72
    End With

    ' Встановлення маржинів зверху (17 мм) та знизу (8 мм)
    ActiveDocument.PageSetup.TopMargin = CentimetersToPoints(1.7)
    ActiveDocument.PageSetup.BottomMargin = CentimetersToPoints(0.8)
             
    With ActiveDocument.PageSetup
        .HeaderDistance = InchesToPoints(0.4)
        .FooterDistance = InchesToPoints(0.4)
    End With

    ActiveDocument.Sections(1).Headers(1).Range.Font.Name = "Georgia"
    ActiveDocument.Sections(1).Headers(1).Range.Font.Size = 10
    ActiveDocument.Sections(1).Headers(1).Range.Text = ActiveDocument.BuiltInDocumentProperties("Author")
    
    ActiveDocument.Sections(1).Headers(3).Range.Font.Name = "Georgia"
    ActiveDocument.Sections(1).Headers(3).Range.Font.Size = 10
    ActiveDocument.Sections(1).Headers(3).Range.Text = ActiveDocument.BuiltInDocumentProperties("Title")
        
    ActiveDocument.Sections(1).Headers(1).Range.ParagraphFormat.Alignment = wdRight

    ' Налаштування для першого нижнього колонтитулу
    With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range
        .Font.Name = "Georgia"
        .Font.Size = 11
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Text = "- " & vbTab & " -"
        .Fields.Add Range:=.Paragraphs(1).Range.Characters(3), Type:=wdFieldPage
    End With
    
    ' Налаштування для непарного нижнього колонтитулу
    With ActiveDocument.Sections(1).Footers(wdHeaderFooterEvenPages).Range
        .Font.Name = "Georgia"
        .Font.Size = 11
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Text = "- " & vbTab & " -"
        .Fields.Add Range:=.Paragraphs(1).Range.Characters(3), Type:=wdFieldPage
    End With
End Sub

