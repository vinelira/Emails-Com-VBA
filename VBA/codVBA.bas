Attribute VB_Name = "Módulo1"
Sub script()
    Planilha1.Select
    
    Dim assunto, CC, texto, nome, email, anexo, Ref As String
    Dim linhaInicial, linhaFinal As Integer
    

    assunto = Range("F1").Value
    CC = Range("F2").Value
    linhaFinal = Range("F3").Value
    texto = Range("F4").Value
    linhaInical = 2

    
    
        
     For linha = linhaInical To linhaFinal
        nome = Range("A" & linha).Value
        email = Range("B" & linha).Value
        anexo = Range("C" & linha).Value
       
        
        
    
        Set obj_outlook = CreateObject("Outlook.application")
        Set novoemail = obj_outlook.createitem(0)
        
        With novoemail
            .display
            .To = email
            .CC = CC
            .Subject = assunto
            .body = texto
            .attachments.Add (ThisWorkbook.Path) & "\" & Cells(linha, 1).Value & ".pdf"           
            .send
                    
        End With
        
        Application.Wait (Now + TimeValue("00:00:04"))
        Set novoemail = Nothing
        
        
    Next linha
    
    MsgBox ("Foram enviados todos os emails")
         
End Sub
