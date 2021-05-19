Sub converterpdf()

ActiveSheet.Shapes.Item("btn").Visible = False

Dim NomPastTrab As String

NomPastTrab = VBA.Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))

 ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
        "c:\emissaonf.pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=True
        
ActiveSheet.Shapes.Item("btn").Visible = True

End Sub
