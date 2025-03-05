Attribute VB_Name = "EnviarEmail"
Sub EnviarEmailPorGestor()
    Dim wsTratamento As Worksheet
    Dim lastRow As Long
    Dim uniqueManagers As Collection
    Dim cell As Range, visCell As Range
    Dim managerName As String
    Dim managerEmail As String
    Dim folderPath As String
    Dim newWb As Workbook
    Dim fileName As String
    Dim OutApp As Object, OutMail As Object
    Dim copyRange As Range
    Dim i As Long
    Dim emailDraftCount As Long
    Dim visRange As Range

    ' Define a planilha "Tratamento"
    On Error Resume Next
    Set wsTratamento = Sheets("Tratamento")
    On Error GoTo 0
    If wsTratamento Is Nothing Then
        MsgBox "A planilha 'Tratamento' não foi encontrada.", vbExclamation
        Exit Sub
    End If

    ' Determina a última linha com dados
    lastRow = wsTratamento.Cells(wsTratamento.Rows.Count, "L").End(xlUp).Row

    ' Seleciona somente os casos filtrados (registros visíveis)
    On Error Resume Next
    Set visRange = wsTratamento.Range("L3:L" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If visRange Is Nothing Then
        MsgBox "Não há registros filtrados na planilha 'Tratamento'.", vbExclamation
        Exit Sub
    End If

    ' Criar coleção de gestores únicos a partir dos casos filtrados
    Set uniqueManagers = New Collection
    On Error Resume Next
    For Each visCell In visRange
        If visCell.Value <> "" Then
            uniqueManagers.Add visCell.Value, CStr(visCell.Value)
        End If
    Next visCell
    On Error GoTo 0

    ' Define a pasta onde os arquivos serão salvos: mesma pasta do arquivo "Tratamento"
    folderPath = ThisWorkbook.Path & "\"

    ' Configura o Outlook para criação dos e-mails
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    If OutApp Is Nothing Then
        MsgBox "O Outlook não está disponível.", vbExclamation
        Exit Sub
    End If

    emailDraftCount = 0

    ' Para cada gestor único, aplica filtro adicional (na coluna L) sobre os dados já filtrados,
    ' copia os registros visíveis, salva o arquivo e gera o rascunho do e-mail.
    For i = 1 To uniqueManagers.Count
        managerName = uniqueManagers(i)
        
        ' Aplica filtro na coluna L para o gestor atual (mantendo os demais filtros já aplicados)
        wsTratamento.Range("A2:M" & lastRow).AutoFilter Field:=12, Criteria1:=managerName
        
        ' Seleciona os registros visíveis resultantes do filtro adicional
        On Error Resume Next
        Set copyRange = wsTratamento.Range("A2:M" & lastRow).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If Not copyRange Is Nothing Then
            ' Cria um novo workbook e copia os dados filtrados
            Set newWb = Workbooks.Add
            copyRange.Copy newWb.Sheets(1).Range("A1")
            
            ' Define o nome do arquivo e salva na mesma pasta do arquivo "Tratamento"
            fileName = folderPath & "Dados_" & managerName & ".xlsx"
            Application.DisplayAlerts = False
            newWb.SaveAs fileName:=fileName, FileFormat:=51  ' xlOpenXMLWorkbook (xlsx)
            Application.DisplayAlerts = True
            newWb.Close False
            
            ' Obter o e-mail do gestor: pega o valor da coluna M na primeira linha visível
            managerEmail = ""
            On Error Resume Next
            For Each cell In wsTratamento.Range("M3:M" & lastRow).SpecialCells(xlCellTypeVisible)
                If cell.Value <> "" Then
                    managerEmail = cell.Value
                    Exit For
                End If
            Next cell
            On Error GoTo 0
            
            ' Cria o rascunho do e-mail com o anexo
            Set OutMail = OutApp.CreateItem(0)  ' 0 = Novo e-mail
            With OutMail
                .To = managerEmail
                .CC = "Ponto Eletronico <pontoeletronico@haoc.com.br>"
                .BCC = ""
                .Subject = "Correção de Ponto: " & managerName
                .Body = "Prezado(a) " & managerName & "," & vbNewLine & vbNewLine & _
                        "Segue em anexo os dados com inconsistência de Ponto nas datas informadas." & vbNewLine & vbNewLine & _
                        "Atenciosamente,"
                .Attachments.Add fileName
                .Display  ' Exibe o e-mail como rascunho para revisão
            End With
            emailDraftCount = emailDraftCount + 1
        End If
        
        ' Restaura o filtro original dos casos já filtrados, removendo o critério de Gestor (coluna L)
        wsTratamento.Range("A2:M" & lastRow).AutoFilter Field:=12
    Next i
    
    MsgBox "Processamento concluído. " & emailDraftCount & " rascunhos de e-mail foram criados.", vbInformation
    
    ' Limpa as variáveis de objeto
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub




