Attribute VB_Name = "BasePontoTratamento"
Sub TratamentoPontoSenior()
    Dim ws As Worksheet, wsTratamento As Worksheet
    Dim lastRow As Long, lastRowTratamento As Long
    Dim copyRange As Range, fillRange As Range, blankCells As Range
    Dim cel As Range

    ' Define a planilha ativa (fonte dos dados)
    Set ws = ActiveSheet

    ' ===== Macro1: Inserir colunas, formatar, aplicar filtros e preencher f�rmulas =====
    ' Insere duas novas colunas � esquerda (coluna A)
    ws.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Aplica formata��o de fonte nas colunas A at� R
    With ws.Range("A:R").Font
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = -16777216
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With

    ' Determina a �ltima linha dinamicamente (usando a coluna C, por exemplo) e adiciona 10 linhas de seguran�a
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row + 10

    ' Aplica AutoFilter no intervalo A1 at� R(lastRow) com os crit�rios desejados
'    With ws.Range("A1:R" & lastRow)
'        .AutoFilter
'        .AutoFilter Field:=13, Criteria1:="="
'        .AutoFilter Field:=3, Criteria1:=Array("", "0001"), Operator:=xlFilterValues
'        .AutoFilter Field:=8, Criteria1:="<>"
'    End With

    With ws.Range("A1:R" & lastRow)
        .AutoFilter
        .AutoFilter Field:=13, Criteria1:="="
        .AutoFilter Field:=3, Criteria1:="<>"
        .AutoFilter Field:=8, Criteria1:="<>"
    End With

    ' Insere f�rmulas nas c�lulas A8 e B8
    ws.Range("A8").FormulaR1C1 = "=RC[2]"
    ws.Range("B8").FormulaR1C1 = "=RC[2]"

    ' Preenche para baixo as f�rmulas nas colunas A e B at� a �ltima linha
    ws.Range("A8:B" & lastRow).FillDown

    ' Remove os filtros aplicados
    ws.ShowAllData

    ' Converte as f�rmulas em valores nas colunas A e B
    With ws.Range("A:B")
        .Value = .Value
    End With

    ' ===== Macro2: Filtragem adicional, c�pia dos dados e tratamento na nova planilha =====
    ' Atualiza a �ltima linha (usando a coluna A) caso os dados tenham sido alterados
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Garante que apenas a coluna D ter� filtro ativo
    ws.AutoFilterMode = False ' Remove qualquer filtro anterior antes de aplicar um novo

    ' Aplica AutoFilter apenas na coluna 4 (D) com crit�rio "<>"
    With ws.Range("A1:R" & lastRow)
        .AutoFilter Field:=4, Criteria1:="<>"
    End With

    ' Copia os dados filtrados das colunas A at� M (apenas c�lulas vis�veis)
    On Error Resume Next
    Set copyRange = ws.Range("A1:M" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If copyRange Is Nothing Then
        MsgBox "Nenhuma c�lula vis�vel para copiar."
        Exit Sub
    End If
    copyRange.Copy

    ' Adiciona uma nova planilha e cola os dados copiados
    Set wsTratamento = Sheets.Add(After:=ws)
    wsTratamento.Paste
    Application.CutCopyMode = False

    ' Renomeia a nova planilha para "Tratamento"
    On Error Resume Next
    Application.DisplayAlerts = False
    wsTratamento.Name = "Tratamento"
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Determina a �ltima linha na planilha "Tratamento"
    lastRowTratamento = wsTratamento.Cells(wsTratamento.Rows.Count, "A").End(xlUp).Row

    ' Preenche c�lulas em branco no intervalo A3:B(lastRowTratamento) com o valor da c�lula imediatamente acima
    Set fillRange = wsTratamento.Range("A3:B" & lastRowTratamento)
    On Error Resume Next
    Set blankCells = fillRange.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0
    If Not blankCells Is Nothing Then
        For Each cel In blankCells
            cel.Value = cel.Offset(-1, 0).Value
        Next cel
    End If

    ' Converte eventuais f�rmulas em valores nas colunas A e B da planilha "Tratamento"
    With wsTratamento.Range("A:B")
        .Value = .Value
    End With

    ' Exclui as colunas F e I na planilha "Tratamento"
    ' Deleta duas vezes a coluna F conforme o comportamento original
    wsTratamento.Columns("G:G").Delete Shift:=xlToLeft
    wsTratamento.Columns("J:J").Delete Shift:=xlToLeft

    ' ===== Configura os nomes das colunas na nova planilha =====
    With wsTratamento
        ' Define os cabe�alhos na linha 2
        .Range("A2").Value = "Mat"
        .Range("B2").Value = "Nome"
        .Range("C2").Value = "Data"
        .Range("D2").Value = "Semana"
        .Range("E2").Value = "Hora1"
        .Range("F2").Value = "Hora2"
        .Range("G2").Value = "C.Custo"
        .Range("H2").Value = "Tipo"
        .Range("I2").Value = "Descri��o"
        .Range("J2").Value = "Qtd"
        .Range("K2").Value = "A��o"
        .Range("L2").Value = "Gestor"
        .Range("M2").Value = "e-mail"
    End With

    ' ===== PROCX para localizar Gestor e E-mail =====
    ' Considera que a planilha "DadosGestores" possui:
    ' Coluna A: Matr�cula; Coluna B: Gestor; Coluna C: E-mail.
    ' Na planilha "Tratamento", a coluna A cont�m a Matr�cula.
    
    With wsTratamento
        ' Insere a f�rmula em L3 usando FormulaLocal (Excel em portugu�s)
        .Range("L3").FormulaLocal = "=PROCX(A3;INDIRETO(""DadosGestores!F:F"");INDIRETO(""DadosGestores!H:I"");""N�o encontrado"")"
        
        ' Preenche a f�rmula da linha 3 at� a �ltima linha (lastRowTratamento)
        .Range("L3:L" & lastRowTratamento).FillDown
        
        ' Se quiser converter as f�rmulas em valores:
        '.Range("L3:L" & lastRowTratamento).Value = .Range("L3:L" & lastRowTratamento).Value
    End With

    ' Remove os filtros da planilha original, se ainda estiverem ativos
    ws.ShowAllData

    ' Exibe mensagem de finaliza��o
    MsgBox "Processamento conclu�do com sucesso!", vbInformation
End Sub

