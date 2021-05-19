Function EnviaEmail()

Criar
converterpdf

Dim data
Dim nome As String
Dim iMsg, iConf, Flds

Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields


data = Now

schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
'Configura o smtp
Flds.Item(schema & "smtpserver") = "zimbramail.penso.com.br"
'Configura a porta de envio de email
Flds.Item(schema & "smtpserverport") = 25
Flds.Item(schema & "smtpauthenticate") = 1
'Configura o email do remetente
Flds.Item(schema & "sendusername") = "teste@gmail.com.br"
'Configura a senha do email remetente
Flds.Item(schema & "sendpassword") = "teste2212"
Flds.Item(schema & "smtpusessl") = 1
Flds.Update

With iMsg
   'Email do destinatário
   .To = "destino@gmail.com.br"
   'Seu email
   .From = "Sistemas@gmail.com.br"
   'Título do email
   .Subject = "Emissão de NF interna"
   'Mensagem do e-mail, você pode enviar formatado em HTML
   .HTMLBody = "Segue como anexo o formulário para a emissão da NF.</br></br>Formulário enviado pelo úsuario (Terminal): " & Application.UserName & "</br>Formulário enviado ás:   " & data
   'Seu nome ou apelido
   '.Sender = "Nathalia"
   'Nome da sua organização
   '.Organization = "TESTE ORG"
   'e-mail de responder para
   '.ReplyTo = "teste@gmail.com.br"
   'Anexo a ser enviado na mensagem. Retire a aspa da linha abaixo e coloque o endereço do arquivo
   .AddAttachment ("c:\Emissão de Nota Fiscal.xlsx")
   .AddAttachment ("c:\emissaonf.pdf")
   Set .Configuration = iConf
   .Send
End With

Set iMsg = Nothing
Set iConf = Nothing
Set Flds = Nothing

Deletar

disparar

End Function
