Sub TestExportToPowerPoint()
    Dim pptApp As Object
    Dim pptPresentation As Object
    Dim pptSlide As Object
    Dim produto As String
    Dim slideNumber As Integer
    Dim exportData As Range
    Dim pptShape As Object
    Dim fundo As Object
    Dim ws As Worksheet
    Dim col As Integer
    Dim ultCol As Integer
    Dim limiteCol As Integer
    
    ' Definir a planilha ACTUAL
    Set ws = ThisWorkbook.Sheets("ACTUAL")
    
    ' Capturar o valor do produto a partir da célula Q17 antes de qualquer operação
    produto = Trim(ws.Range("Q17").Value) ' Produto selecionado manualmente na célula Q17
    
    ' Mapear os slides para cada produto
    Select Case produto
        Case "HJF XJF"
            slideNumber = 8
        Case "HJD XJD"
            slideNumber = 17
        Case "B52 X52"
            slideNumber = 10
        Case "BBB XBB"
            slideNumber = 31
        Case Else
            MsgBox "Produto não encontrado. Verifique o valor na célula Q17.", vbExclamation
            Exit Sub
    End Select
    
    ' Inicializar o PowerPoint
    On Error Resume Next
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    On Error GoTo 0
    
    ' Caminho do template do PowerPoint (agora puxando de R24)
    Dim caminhoArquivo As String
    caminhoArquivo = Trim(ws.Range("R24").Value)
    
    ' Verificar se o caminho foi preenchido
    If caminhoArquivo = "" Then
        MsgBox "O caminho do arquivo PowerPoint não foi especificado em R24.", vbExclamation
        Exit Sub
    End If
    
    ' Abrir a apresentação PowerPoint
    On Error Resume Next
    Set pptPresentation = pptApp.Presentations.Open(caminhoArquivo)
    If pptPresentation Is Nothing Then
        MsgBox "Não foi possível abrir o arquivo PowerPoint. Verifique o caminho em R24.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Determinar até onde devemos pegar os dados (antes da palavra "Limite" na linha 1)
    ultCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column ' Última coluna preenchida na linha 1
    
    For col = 1 To ultCol
        If LCase(Trim(ws.Cells(1, col).Value)) = "limite" Then
            limiteCol = col - 1 ' A primeira coluna antes da palavra "Limite"
            Exit For
        End If
    Next col
    
    ' Se "Limite" não for encontrado, usar todas as colunas
    If limiteCol = 0 Then limiteCol = ultCol
    
    ' Selecionar a tabela até a coluna antes de "Limite"
    Set exportData = ws.Range(ws.Cells(2, 2), ws.Cells(31, limiteCol)) ' Dados de B2 até última coluna antes de "Limite"
    
    ' Selecionar o slide específico
    Set pptSlide = pptPresentation.Slides(slideNumber)
    
    ' Criar um retângulo branco como fundo antes de colar a imagem
    Set fundo = pptSlide.Shapes.AddShape(msoShapeRectangle, 45, 145, 810, 410)
    fundo.Fill.ForeColor.RGB = RGB(255, 255, 255) ' Branco
    fundo.Line.Visible = msoFalse ' Remover borda do retângulo
    
    ' Copiar apenas os dados até a coluna antes de "Limite" (ou toda a tabela se não houver limite)
    exportData.Copy
    pptSlide.Shapes.PasteSpecial DataType:=2 ' Cola como imagem (DataType:=2)
    
    ' Ajustar a posição e o tamanho da imagem no slide
    Set pptShape = pptSlide.Shapes(pptSlide.Shapes.Count) ' Última forma colada
    pptShape.Left = 30 ' Ajuste horizontal (esquerda)
    pptShape.Top = 60 ' Ajuste vertical (superior)
    pptShape.Width = 800 ' Largura da área
    pptShape.Height = 400 ' Altura da área
    
    ' Enviar a imagem para frente, garantindo que o fundo fique atrás
    fundo.ZOrder msoSendToBack
    pptShape.ZOrder msoBringToFront
    
    MsgBox "Dados exportados para o slide " & slideNumber
End Sub



    ' Script escrito por Giovanni Müller | Economics Studies Intern |
    ' Email: giovanni.muller-extern@renault.com
