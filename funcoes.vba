Sub cabecalhoEstimativa(wsDados As Worksheet)
    
    ' Cabeçalho definido para preenchimento
    wsDados.Cells(2, 2).Value = "Parcela"
    wsDados.Cells(2, 3).Value = "Saldo Devedor"
    wsDados.Cells(2, 4).Value = "Prestação"
    wsDados.Cells(2, 5).Value = "Amortização"
    wsDados.Cells(2, 6).Value = "Juros"
    wsDados.Cells(2, 7).Value = "Seguro"
    wsDados.Cells(2, 8).Value = "Administração"
        
End Sub

Sub parcelasSimulador(wsBase As Worksheet, wsDados As Worksheet, ultimaParcela As Integer)

    Dim valorFinanciado As Double
    Dim jurosDefinido As Double
    Dim percentualSeguro As Double
    Dim percentualAdministracao As Double
    
    valorFinanciado = wsBase.Range("C4").Value
    jurosDefinido = ((1 + wsBase.Range("C5").Value) ^ (1 / 12) - 1)
    percentualSeguro = valorFinanciado * wsBase.Range("C7").Value
    percentualAdministracao = valorFinanciado * wsBase.Range("C8").Value
    baixaFinanciamento = Pmt(jurosDefinido, ultimaParcela, -valorFinanciado)
    jurosCalculado = valorFinanciado * jurosDefinido

    primeiraLinha = 3
    parcela = 0
    
    Do While parcela <= ultimaParcela
        If parcela = 0 Then
            wsDados.Cells(primeiraLinha, 2).Value = parcela
            wsDados.Cells(primeiraLinha, 3).Value = valorFinanciado
            
            parcela = parcela + 1
            valorFinanciado = valorFinanciado + jurosCalculado - baixaFinanciamento
        Else
            wsDados.Cells(primeiraLinha + parcela, 2).Value = parcela
            wsDados.Cells(primeiraLinha + parcela, 3).Value = valorFinanciado
            wsDados.Cells(primeiraLinha + parcela, 4).Value = baixaFinanciamento + percentualSeguro + percentualAdministracao
            wsDados.Cells(primeiraLinha + parcela, 5).Value = baixaFinanciamento - jurosCalculado
            wsDados.Cells(primeiraLinha + parcela, 6).Value = jurosCalculado
            wsDados.Cells(primeiraLinha + parcela, 7).Value = percentualSeguro
            wsDados.Cells(primeiraLinha + parcela, 8).Value = percentualAdministracao
            
            parcela = parcela + 1
            jurosCalculado = valorFinanciado * jurosDefinido
            valorFinanciado = valorFinanciado + jurosCalculado - baixaFinanciamento

        End If
    Loop
   
End Sub

Sub formataTabela(wsBase As Worksheet, wsDados As Worksheet, ultimaParcela As Integer)

    Dim corBorda As Long
    Dim corFundoCabecalho As Long
    Dim corTextoCabecalho As Long
    Dim tabela As Range
    Dim cabecalho As Range
    Dim dadosNumero As Range
    
    ultimaLinha = ultimaParcela + 3
    ultimaColuna = wsDados.Cells(2, 2).End(xlToRight).Column
    
    Set tabela = Range(wsDados.Cells(2, 2), wsDados.Cells(ultimaLinha, ultimaColuna))
    Set cabecalho = Range(wsDados.Cells(2, 2), wsDados.Cells(2, ultimaColuna))
    Set dadosNumero = Range(wsDados.Cells(3, 3), wsDados.Cells(ultimaLinha, ultimaColuna))
    
    corBorda = RGB(166, 166, 166)
    corFundoCabecalho = RGB(0, 112, 192)
    corTextoCabecalho = RGB(255, 255, 255)
    
    With tabela
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 10
        .Font.Name = "Calibri"
        .Borders.LineStyle = xlContinuous
        .Borders.Color = corBorda
    End With
    
    With cabecalho
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = corTextoCabecalho
        .Interior.Color = corFundoCabecalho
    End With
        
    With dadosNumero
        .NumberFormat = "#,##0.00"
    End With

End Sub

Sub parcelasEfetivo(wsBase As Worksheet, wsDados As Worksheet, ultimaParcela As Integer)

    Dim valorFinanciado As Double
    Dim jurosDefinido As Double
    Dim percentualSeguro As Double
    Dim percentualAdministracao As Double
    Dim baixaFinanciamento As Double
    Dim amortizacao As Double
    
    valorFinanciado = wsBase.Range("H4").Value
    jurosDefinido = ((1 + wsBase.Range("H5").Value) ^ (1 / 12) - 1)
    percentualSeguro = valorFinanciado * wsBase.Range("C7").Value
    percentualAdministracao = valorFinanciado * wsBase.Range("H8").Value
    jurosCalculado = valorFinanciado * jurosDefinido
    
    If Pmt(jurosDefinido, ultimaParcela, -valorFinanciado) > wsBase.Range("H12").Value Then
        wsBase.Range("I12").Value = Pmt(jurosDefinido, ultimaParcela, -valorFinanciado)
        wsBase.Range("I12").NumberFormat = "#,##0.00"
        baixaFinanciamento = Pmt(jurosDefinido, ultimaParcela, -valorFinanciado)
    Else
        wsBase.Range("I12").ClearContents
        baixaFinanciamento = wsBase.Range("H12")
    End If
    
    amortizacao = baixaFinanciamento - jurosCalculado

    primeiraLinha = 3
    parcela = 0
    
    Do While parcela <= ultimaParcela
        If parcela = 0 Then
            wsDados.Cells(primeiraLinha, 2).Value = parcela
            wsDados.Cells(primeiraLinha, 3).Value = valorFinanciado
            
            parcela = parcela + 1
            valorFinanciado = valorFinanciado + jurosCalculado - baixaFinanciamento
        Else
            wsDados.Cells(primeiraLinha + parcela, 2).Value = parcela
            wsDados.Cells(primeiraLinha + parcela, 4).Formula = "=SUM(E" & primeiraLinha + parcela & ":H" & primeiraLinha + parcela & ")"
            wsDados.Cells(primeiraLinha + parcela, 5).Value = amortizacao
            wsDados.Cells(primeiraLinha + parcela, 6).Value = jurosCalculado
            wsDados.Cells(primeiraLinha + parcela, 7).Value = percentualSeguro
            wsDados.Cells(primeiraLinha + parcela, 8).Value = percentualAdministracao
            wsDados.Cells(primeiraLinha + parcela, 3).Value = valorFinanciado
            
            parcela = parcela + 1
            jurosCalculado = valorFinanciado * jurosDefinido
            amortizacao = WorksheetFunction.Min(valorFinanciado, baixaFinanciamento - jurosCalculado)
            valorFinanciado = WorksheetFunction.Max(0, valorFinanciado + jurosCalculado - baixaFinanciamento)

        End If
    Loop
   
End Sub

Sub parcelasRealizado(wsBase As Worksheet, wsDados As Worksheet, ultimaParcela As Integer, wsPagto As Worksheet)

    Dim valorFinanciado As Double
    Dim jurosDefinido As Double
    Dim baixaFinanciamento As Double
    Dim amortizacao As Double
    Dim wsTaxa As Worksheet
    Dim correcaoMonetaria As Double
    
    Set wsTaxa = ThisWorkbook.Sheets("Efetivo")
    
    valorFinanciado = wsBase.Range("M4").Value
    jurosDefinido = ((1 + wsBase.Range("M5").Value) ^ (1 / 12) - 1)
    jurosCalculado = valorFinanciado * jurosDefinido
    correcaoMonetaria = wsBase.Range("M12").Value
    
    primeiraLinha = 3
    parcela = 0
    
    ultimopagto = wsPagto.Cells(wsPagto.Rows.Count, 2).End(xlUp).Value
    
    Do While parcela <= ultimopagto
        If parcela = 0 Then
            wsDados.Cells(primeiraLinha, 2).Value = parcela
            wsDados.Cells(primeiraLinha, 3).Value = valorFinanciado
    
            parcela = parcela + 1
    
        Else
            linhaAtual = primeiraLinha + parcela
            
            txAdm = wsTaxa.Cells(linhaAtual, 8).Value
            txSeguro = wsTaxa.Cells(linhaAtual, 7).Value
            pagamento = wsPagto.Cells(linhaAtual, 7).Value
            amortizaParcela = pagamento - txAdm - txSeguro - jurosCalculado
            valorCorrecaoMonetaria = valorFinanciado * correcaoMonetaria
            valorFinanciado = valorFinanciado + valorCorrecaoMonetaria - amortizaParcela
            
            wsDados.Cells(linhaAtual, 2).Value = parcela
            wsDados.Cells(linhaAtual, 8).Value = txAdm
            If parcela = ultimaParcela Then
                wsDados.Cells(linhaAtual, 7).Value = 0
            Else
                wsDados.Cells(linhaAtual, 7).Value = txSeguro
            End If
            wsDados.Cells(linhaAtual, 6).Value = jurosCalculado
            wsDados.Cells(linhaAtual, 4).Value = pagamento
            wsDados.Cells(linhaAtual, 5).Value = amortizaParcela
            wsDados.Cells(linhaAtual, 3).Value = valorFinanciado
            wsDados.Cells(linhaAtual, 9).Value = valorCorrecaoMonetaria
            
            parcela = parcela + 1
            jurosCalculado = valorFinanciado * jurosDefinido
        End If
    Loop
    
    ultimaLinhaPagto = wsPagto.Cells(wsPagto.Rows.Count, 2).End(xlUp).Row
    linhaPrazoRestante = wsPagto.Cells(ultimaLinhaPagto, 8).End(xlUp).Row
    prazoRestante = wsPagto.Cells(ultimaLinhaPagto, 8).End(xlUp).Value
    
    If prazoRestante = "Prazo Restante" Then
        prazo = wsBase.Range("M10").Value
    Else
        prazo = wsPagto.Cells(linhaPrazoRestante, 2).Value + wsPagto.Cells(linhaPrazoRestante, 8).Value
        wsBase.Range("M10").Value = prazo
    End If
    
    pagamento = wsPagto.Cells(linhaPrazoRestante, 3).Value
    
    Do While parcela <= prazo

        linhaAtual = primeiraLinha + parcela

        txAdm = wsTaxa.Cells(linhaAtual, 8).Value
        txSeguro = wsTaxa.Cells(linhaAtual, 7).Value
        amortizaParcela = pagamento - txAdm - txSeguro - jurosCalculado
        valorCorrecaoMonetaria = 0
        
        wsDados.Cells(linhaAtual, 2).Value = parcela
        wsDados.Cells(linhaAtual, 8).Value = txAdm
        If parcela = prazo Then
            wsDados.Cells(linhaAtual, 7).Value = 0
        Else
            wsDados.Cells(linhaAtual, 7).Value = txSeguro
        End If
        wsDados.Cells(linhaAtual, 6).Value = jurosCalculado
        wsDados.Cells(linhaAtual, 5).Value = WorksheetFunction.Min(valorFinanciado, amortizaParcela)
        valorFinanciado = valorFinanciado + valorCorrecaoMonetaria - amortizaParcela
        wsDados.Cells(linhaAtual, 3).Value = WorksheetFunction.Max(0, valorFinanciado)
        
        parcela = parcela + 1
        jurosCalculado = valorFinanciado * jurosDefinido

    Loop
   
End Sub
