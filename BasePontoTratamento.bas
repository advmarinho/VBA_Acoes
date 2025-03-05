Attribute VB_Name = "BasePontoTratamento"
Sub TratamentoPontoSenior()
    Dim ws As Worksheet, wsTratamento As Worksheet
    Dim lastRow As Long, lastRowTratamento As Long
    Dim copyRange As Range, fillRange As Range, blankCells As Range
    Dim cel As Range

    ' Define a planilha ativa (fonte dos dados)
    Set ws = ActiveSheet

    ' ===== Macro1: Inserir colunas, formatar, aplicar filtros e preencher fórmulas =====
    ' Insere duas novas colunas à esquerda (coluna A)
    ws.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Aplica formatação de fonte nas colunas A até R
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

    ' Determina a última linha dinamicamente (usando a coluna C, por exemplo) e adiciona 10 linhas de segurança
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row + 10

    ' Aplica AutoFilter no intervalo A1 até R(lastRow) com os critérios desejados
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

    ' Insere fórmulas nas células A8 e B8
    ws.Range("A8").FormulaR1C1 = "=RC[2]"
    ws.Range("B8").FormulaR1C1 = "=RC[2]"

    ' Preenche para baixo as fórmulas nas colunas A e B até a última linha
    ws.Range("A8:B" & lastRow).FillDown

    ' Remove os filtros aplicados
    ws.ShowAllData

    ' Converte as fórmulas em valores nas colunas A e B
    With ws.Range("A:B")
        .Value = .Value
    End With

    ' ===== Macro2: Filtragem adicional, cópia dos dados e tratamento na nova planilha =====
    ' Atualiza a última linha (usando a coluna A) caso os dados tenham sido alterados
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Garante que apenas a coluna D terá filtro ativo
    ws.AutoFilterMode = False ' Remove qualquer filtro anterior antes de aplicar um novo

    ' Aplica AutoFilter apenas na coluna 4 (D) com critério "<>"
    With ws.Range("A1:R" & lastRow)
        .AutoFilter Field:=4, Criteria1:="<>"
    End With

    ' Copia os dados filtrados das colunas A até M (apenas células visíveis)
    On Error Resume Next
    Set copyRange = ws.Range("A1:M" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If copyRange Is Nothing Then
        MsgBox "Nenhuma célula visível para copiar."
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

    ' Determina a última linha na planilha "Tratamento"
    lastRowTratamento = wsTratamento.Cells(wsTratamento.Rows.Count, "A").End(xlUp).Row

    ' Preenche células em branco no intervalo A3:B(lastRowTratamento) com o valor da célula imediatamente acima
    Set fillRange = wsTratamento.Range("A3:B" & lastRowTratamento)
    On Error Resume Next
    Set blankCells = fillRange.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0
    If Not blankCells Is Nothing Then
        For Each cel In blankCells
            cel.Value = cel.Offset(-1, 0).Value
        Next cel
    End If

    ' Converte eventuais fórmulas em valores nas colunas A e B da planilha "Tratamento"
    With wsTratamento.Range("A:B")
        .Value = .Value
    End With

    ' Exclui as colunas F e I na planilha "Tratamento"
    ' Deleta duas vezes a coluna F conforme o comportamento original
    wsTratamento.Columns("G:G").Delete Shift:=xlToLeft
    wsTratamento.Columns("J:J").Delete Shift:=xlToLeft

    ' ===== Configura os nomes das colunas na nova planilha =====
    With wsTratamento
        ' Define os cabeçalhos na linha 2
        .Range("A2").Value = "Mat"
        .Range("B2").Value = "Nome"
        .Range("C2").Value = "Data"
        .Range("D2").Value = "Semana"
        .Range("E2").Value = "Hora1"
        .Range("F2").Value = "Hora2"
        .Range("G2").Value = "C.Custo"
        .Range("H2").Value = "Tipo"
        .Range("I2").Value = "Descrição"
        .Range("J2").Value = "Qtd"
        .Range("K2").Value = "Ação"
        .Range("L2").Value = "Gestor"
        .Range("M2").Value = "e-mail"
    End With

    ' ===== PROCX para localizar Gestor e E-mail =====
    ' Considera que a planilha "DadosGestores" possui:
    ' Coluna A: Matrícula; Coluna B: Gestor; Coluna C: E-mail.
    ' Na planilha "Tratamento", a coluna A contém a Matrícula.
    
    With wsTratamento
        ' Insere a fórmula em L3 usando FormulaLocal (Excel em português)
        .Range("L3").FormulaLocal = "=PROCX(A3;INDIRETO(""DadosGestores!F:F"");INDIRETO(""DadosGestores!H:I"");""Não encontrado"")"
        
        ' Preenche a fórmula da linha 3 até a última linha (lastRowTratamento)
        .Range("L3:L" & lastRowTratamento).FillDown
        
        ' Se quiser converter as fórmulas em valores:
        '.Range("L3:L" & lastRowTratamento).Value = .Range("L3:L" & lastRowTratamento).Value
    End With

    ' Remove os filtros da planilha original, se ainda estiverem ativos
    ws.ShowAllData

    ' Exibe mensagem de finalização
    MsgBox "Processamento concluído com sucesso!", vbInformation
End Sub

