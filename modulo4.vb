Sub Criar()


ActiveSheet.Shapes.Item("btn").Visible = False
    
 Dim NovoArquivoXLS As Workbook
 Dim sPlanAEnviar As String
 
'Define a planilha que será criada. Ex.: Plan1, Balancete, Lista De Nomes, etc
  sPlanAEnviar = "Emissão de Nota Fiscal"
 
 'Cria um novo arquivo excel
  Set NovoArquivoXLS = Application.Workbooks.Add
 
 'Copia a planilha para o novo arquivo criado
  ThisWorkbook.Sheets(sPlanAEnviar).Copy Before:=NovoArquivoXLS.Sheets(1)
 
 'Salva o arquivo
  NovoArquivoXLS.SaveAs "c:" & "\" & sPlanAEnviar & ".xlsx"
 sExcluirAnexoTemporario = NovoArquivoXLS.FullName
 
 'Fecha o arquivo novo
  NovoArquivoXLS.Close
 
ActiveSheet.Shapes.Item("btn").Visible = True
 
End Sub
