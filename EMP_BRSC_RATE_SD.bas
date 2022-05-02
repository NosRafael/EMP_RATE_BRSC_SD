Option Compare Database
Option Explicit

'==============================================================================
'VARIAVEIS PUBLICAS DO MODULO
Dim FOLDERORNAME As String
Dim FOLDERDESTNAME As String
Dim dataref As String
Dim sOperadora As String
Dim sRamo As String

'==============================================================================
Sub EMP_BRSC_RATE_SA()
'==============================================================================
'ROTINA PARA EXECUCAO DO RATEIO DA EMPRESA
'
'ARQUIVOS:
'   - FM
'   - PC
'   - BaseEmp_Ativos.xlsx
'   - BaseEmp_Desligados.xlsx
'   - Taxas.xlsx
'
'VERSAO 001 - 02/05/2022 - RAFAEL BICALHO PAIVA - IMPLANTAÇÃO

'==============================================================================

    FOLDERORNAME = ""
    FOLDERDESTNAME = FOLDERORNAME & "\Finalizado"
    
    dataref = InputBox("Informe a referência de processamento: (YYYY_MM)", "Referência de processamento", Year(Date) & "_" & Format(Month(Date), "00"))
    If Trim(dataref) = Empty Then dataref = Year(Date) & "_" & Format(Month(Date), "00")
    
    sOperadora = InputBox("Informe a operadora", "Operadora", "Bradesco")
    sRamo = InputBox("Informe o ramo", "Saude", "Saude")
    
    Excl_Tabs
        
    ImportaArquivo
    Processamento
    ProcessamentoCusto
    ExportaArquivo
    
    MsgBox "Processo finalizado com sucesso!", vbInformation

End Sub


Private Sub ImportaArquivo()

    Dim filebase1 As String
    Dim fs As Object
    Dim choice As String
    Dim i As Integer
    
    DoCmd.SetWarnings False
    
    
    '==========================================================================
    'IMPORTA ARQUIVO TAXAS
    '==========================================================================
    Set fs = New clFileSearch
    With fs
        
        .NewSearch
        .LookIn = FOLDERORNAME
        .filename = "TAXAS*.xlsx"
        .Execute
        
        If .FoundFiles.Count = 0 Then
            MsgBox "Arquivo " & .filename & " não localizado!", vbCritical
            End
        End If
        
        For i = 1 To .FoundFiles.Count
            choice = .FoundFiles(i)
            DoCmd.TransferSpreadsheet acImport, , "TAXAS", choice, True
        Next i
        
    End With
    
    '==========================================================================
    'IMPORTA ARQUIVO DEPARA
    '==========================================================================
    Set fs = New clFileSearch
    With fs
        
        .NewSearch
        .LookIn = FOLDERORNAME
        .filename = "DEPARA*.xlsx"
        .Execute
        
        If .FoundFiles.Count = 0 Then
            MsgBox "Arquivo " & .filename & " não localizado!", vbCritical
            End
        End If
        
        For i = 1 To .FoundFiles.Count
            choice = .FoundFiles(i)
            DoCmd.TransferSpreadsheet acImport, , "DEPARA", choice, True
        Next i
        
    End With
    
    '==========================================================================
    'IMPORTA ARQUIVO LOCALIZADOS
    '==========================================================================
        
    Set fs = New clFileSearch
    With fs
        
        .NewSearch
        .LookIn = FOLDERORNAME
        .filename = "*NaoLocalizados*.xls*"
        .Execute
        
        For i = 1 To .FoundFiles.Count
            choice = .FoundFiles(i)
            DoCmd.TransferSpreadsheet acImport, , "LOCALIZADOS", choice, True
        Next i
    End With
    
'    '==========================================================================
'    'IMPORTA RELATORIO DE COPARTICIPACAO
'    '==========================================================================
'
'    Set fs = New clFileSearch
'    With fs
'
'        .NewSearch
'        .LookIn = FOLDERORNAME
'        .filename = "HDI-Bradesco(Saude)_Coparticipacao(Conferencia)*.xls*"
'        .Execute
'
'        For i = 1 To .FoundFiles.Count
'            CHOICE = .FoundFiles(i)
'            'ts(i) = Right(Replace(GetFileName(CHOICE), " ", ""), 6)
'
'            DoCmd.TransferSpreadsheet acImport, , "COPA", CHOICE, True, "RESUMO!"
'
'        Next i
'    End With

    '==========================================================================
    'IMPORTA ARQUIVO DE CONFERENCIA DA EMPRESA HDI
    '==========================================================================
    Set fs = New clFileSearch
    With fs

        .NewSearch
        .LookIn = FOLDERORNAME & "\Arquivos\Empresa"
        .filename = "Conferencia_Beneficios_HDI*.xls*"
        .Execute
        
        If .FoundFiles.Count = 0 Then
            MsgBox "Arquivo " & .filename & " não localizado!", vbCritical
            End
        End If

        For i = 1 To .FoundFiles.Count
            choice = .FoundFiles(i)
            DoCmd.TransferSpreadsheet acImport, , "CONF_HDI", choice, True
        Next i
        
    End With

    '==========================================================================
    'IMPORTA ARQUIVO DE CONFERENCIA DA EMPRESA GLOBAL
    '==========================================================================

    Set fs = New clFileSearch
    With fs

        .NewSearch
        .LookIn = FOLDERORNAME & "\Arquivos\Empresa"
        .filename = "Conferencia_Beneficios_Global*.xls*"
        .Execute
        
        If .FoundFiles.Count = 0 Then
            MsgBox "Arquivo " & .filename & " não localizado!", vbCritical
            End
        End If

        For i = 1 To .FoundFiles.Count
            choice = .FoundFiles(i)
            DoCmd.TransferSpreadsheet acImport, , "CONF_GLOBAL", choice, True
        Next i
        
    End With

    '==========================================================================
    'IMPORTA ARQUIVO DE CONFERENCIA DA EMPRESA SANTANDER
    '==========================================================================

    Set fs = New clFileSearch
    With fs

        .NewSearch
        .LookIn = FOLDERORNAME & "\Arquivos\Empresa"
        .filename = "Conferencia_Beneficios_Santander*.xls*"
        .Execute
        
        If .FoundFiles.Count = 0 Then
            MsgBox "Arquivo " & .filename & " não localizado!", vbCritical
            End
        End If

        For i = 1 To .FoundFiles.Count
            choice = .FoundFiles(i)
            DoCmd.TransferSpreadsheet acImport, , "CONF_SANTANDER", choice, True
        Next i
        
    End With
    
    '==========================================================================
    'IMPORTA BASE EMPRESA - MALA DIRETA
    '==========================================================================
    Set fs = New clFileSearch
    With fs
        
        .NewSearch
        .LookIn = FOLDERORNAME & "\Arquivos\Empresa"
        .filename = "*MALA DIRETA*.xls*"
        .Execute
        
        If .FoundFiles.Count = 0 Then
            MsgBox "Arquivo " & .filename & " não localizado!", vbCritical
            End
        End If
        
        For i = 1 To .FoundFiles.Count
            choice = .FoundFiles(i)
            DoCmd.TransferSpreadsheet acImport, , "EMPRESA", choice, True, "MALA DIRETA!A1:AP65000"
            DoCmd.TransferSpreadsheet acImport, , "EMPRESA", choice, True, "Global!A1:AP65000"
            DoCmd.TransferSpreadsheet acImport, , "EMPRESA", choice, True, "Santander Auto!A1:AP65000"
        Next i
        
    End With
     
    '==========================================================================
    'IMPORTA ARQUIVO FM
    '==========================================================================
    Set fs = New clFileSearch
    With fs
        
        .NewSearch
        .LookIn = FOLDERORNAME & "\Arquivos\Operadora"
        .filename = "FM*.TXT"
        .Execute
        
        If .FoundFiles.Count = 0 Then
            MsgBox "Arquivo " & .filename & " não localizado!", vbCritical
            End
        End If
        
        For i = 1 To .FoundFiles.Count
            choice = .FoundFiles(i)
            DoCmd.TransferText acImportFixed, "IM_BRSC_FATU_REG3", "FM", choice
        Next i
        
    End With
    
    '==========================================================================
    'IMPORTA ARQUIVO PC PARA BUSCAR CPF E NOME TIT
    '==========================================================================
    Set fs = New clFileSearch
    With fs
        
        .NewSearch
        .LookIn = FOLDERORNAME & "\Arquivos\Operadora\PC"
        .filename = "PC*.TXT"
        .Execute
        
        If .FoundFiles.Count = 0 Then
            MsgBox "Arquivo " & .filename & " não localizado!", vbCritical
            End
        End If
        
        For i = 1 To .FoundFiles.Count
            choice = .FoundFiles(i)
            DoCmd.TransferText acImportFixed, "IM_BRSC_ATVO_TIT", "PC_TIT", choice
        Next i
        
    End With
    
    DoCmd.SetWarnings True
        
End Sub

Private Sub Processamento()
    
    Const cIOF As Double = 0.0238      'CONSTANTE VALOR DO IOF
    Dim sSQL As String
    Dim ultimoDia As Date
    
    DoCmd.SetWarnings False
    
    '==============================================================================
    'TRATA TABELA EMPRESA
    '==============================================================================
    
    DoCmd.RunSQL "ALTER TABLE EMPRESA ADD COLUMN EMPRESA TEXT"
    DoCmd.RunSQL "ALTER TABLE EMPRESA ADD COLUMN STATUS TEXT"
    
    DoCmd.RunSQL "ALTER TABLE EMPRESA ALTER COLUMN Cpf TEXT"
    
    DoCmd.RunSQL "DELETE * FROM EMPRESA WHERE Nome IS NULL OR Nome = ''"
    
    DoCmd.RunSQL "UPDATE EMPRESA SET EMPRESA = 'GLOBAL' WHERE Unidade = 'HDI GLOBAL'"
    DoCmd.RunSQL "UPDATE EMPRESA SET EMPRESA = 'SANTANDER' WHERE Unidade = '01-SANTANDER AUTO SA'"
    DoCmd.RunSQL "UPDATE EMPRESA SET EMPRESA = 'HDI' WHERE EMPRESA IS NULL"
    DoCmd.RunSQL "UPDATE EMPRESA SET STATUS = 'DESLIGADO' WHERE LEFT(Situacao, 4) = 'DEMI'"
    DoCmd.RunSQL "UPDATE EMPRESA SET STATUS = 'ATIVO' WHERE STATUS IS NULL"
    DoCmd.RunSQL "UPDATE EMPRESA SET CPF = FORMAT(REPLACE(REPLACE(CPF, '.', ''), '-', ''), '00000000000')"
        
    
    '==============================================================================
    'TRATA TABELAS
    '==============================================================================

    'Deleta registros
    DoCmd.RunSQL "DELETE FROM FM WHERE [TIPO DE REGISTRO] <> '3'"
    DoCmd.RunSQL "DELETE FROM FM WHERE [NÚMERO DA SUBFATURA] IN ('0005', '0500', '0851', '0852')"
    'DoCmd.RunSQL "DELETE FROM COPA WHERE [SUB] IN ('005', '500', '851', '852')"
    DoCmd.RunSQL "DELETE FROM PC_TIT WHERE  [TIPO DE REGISTRO] <> '2'"
    
    'Aumenta tamanho campos
    DoCmd.RunSQL "ALTER TABLE FM ALTER COLUMN [DATA DE NASCIMENTO] TEXT(10)"
    DoCmd.RunSQL "ALTER TABLE FM ALTER COLUMN [DATA INÍCIO VIGÊNCIA] TEXT(10)"
    DoCmd.RunSQL "ALTER TABLE FM ALTER COLUMN [DATA DE LANÇAMENTO] TEXT(10)"
    DoCmd.RunSQL "ALTER TABLE FM ALTER COLUMN [CÓD GRAU PARENTESCO] TEXT"
    
    'Cria campos
    DoCmd.RunSQL "ALTER TABLE FM ADD COLUMN" & _
                " TIPO TEXT" & _
                ", CPF TEXT" & _
                ", MATRICULA TEXT" & _
                ", IOF DOUBLE" & _
                ", VALOR_FINAL DOUBLE" & _
                ", VALOR2 DOUBLE" & _
                ", UNID TEXT" & _
                ", FUNCAO TEXT" & _
                ", CR TEXT" & _
                ", NOME_TIT TEXT" & _
                ", SITUACAO TEXT" & _
                ", EMPRESA TEXT" & _
                ", VALOR_TAXA DOUBLE"

                

    '==============================================================================
    'PROCESSA
    '==============================================================================
    
    'SelfJoin para atualizar o plano
    sSQL = "UPDATE FM F "
    sSQL = sSQL & " INNER JOIN FM T ON F.[NÚMERO DO CERTIFICADO] = T.[NÚMERO DO CERTIFICADO] AND CDBL(F.[NÚMERO DA SUBFATURA]) = CDBL(T.[NÚMERO DA SUBFATURA])"
    sSQL = sSQL & " SET"
    sSQL = sSQL & " F.[CÓDIGO DO PLANO] = T.[CÓDIGO DO PLANO]"
    sSQL = sSQL & " WHERE F.[NÚMERO DO CERTIFICADO] <> '0000000'"
    DoCmd.RunSQL sSQL


    'Join FM x PC_TIT para buscar nome dos titulares e cpf
    sSQL = "UPDATE FM F "
    sSQL = sSQL & " INNER JOIN PC_TIT T ON F.[NÚMERO DO CERTIFICADO] = T.[NÚMERO DO CERTIFICADO] AND CDBL(F.[NÚMERO DA SUBFATURA]) = CDBL(T.[NÚMERO DA SUBFATURA])"
    sSQL = sSQL & " SET"
    sSQL = sSQL & "      F.NOME_TIT  =      T.[NOME DO SEGURADO]"
    sSQL = sSQL & ",     F.[CPF]     =      T.[CPF]"
    DoCmd.RunSQL sSQL
    
    'Join FM x DEPARA para buscar nome da empresa
    sSQL = "UPDATE FM F "
    sSQL = sSQL & " INNER JOIN DEPARA D ON F.[NÚMERO DA SUBFATURA] = D.[SUBFATURA]"
    sSQL = sSQL & " SET"
    sSQL = sSQL & "      F.[EMPRESA]  =      D.[EMPRESA]"
    DoCmd.RunSQL sSQL
    
       
    'Seta valor negativo se COD_LANCAMENTO > 49
    DoCmd.RunSQL "UPDATE FM SET" & _
                " VALOR2 = IIF(VAL([CÓDIGO DO LANÇAMENTO]) > 49, (CDBL([VALOR DO LANÇAMENTO]) / 100) * -1, CDBL([VALOR DO LANÇAMENTO]) / 100)"
                    
  

'    'Join FM x BASE EMPRESA ATIVOS para puxar a matrícula'

    DoCmd.RunSQL "ALTER TABLE EMPRESA ALTER COLUMN CPF DOUBLE"
    DoCmd.RunSQL "ALTER TABLE FM ALTER COLUMN CPF DOUBLE"

    sSQL = "UPDATE FM F " & Chr(13)
    sSQL = sSQL & " INNER JOIN EMPRESA E ON F.CPF = E.Cpf AND F.EMPRESA = E.Empresa" & Chr(13)
    sSQL = sSQL & " SET" & Chr(13)
    sSQL = sSQL & "       F.[MATRICULA]          = E.[Registro]" & Chr(13)
    sSQL = sSQL & ",      F.[UNID]               = E.[Unidade]" & Chr(13)
    sSQL = sSQL & ",      F.[FUNCAO]             = E.[Nome_Funcao]" & Chr(13)
    sSQL = sSQL & ",      F.[CR]                 = E.[Localização]" & Chr(13)
    sSQL = sSQL & ",      F.[SITUACAO]           = E.[Situacao]" & Chr(13)
    sSQL = sSQL & " WHERE E.[Status] = 'Ativo'"
    DoCmd.RunSQL sSQL
    
    '    'Join FM x BASE EMPRESA ATIVOS para puxar a matrícula'

    sSQL = "UPDATE FM F " & Chr(13)
    sSQL = sSQL & " INNER JOIN EMPRESA E ON F.CPF = E.Cpf AND F.EMPRESA = E.Empresa" & Chr(13)
    sSQL = sSQL & " SET" & Chr(13)
    sSQL = sSQL & "       F.[MATRICULA]          = E.[Registro]" & Chr(13)
    sSQL = sSQL & ",      F.[UNID]               = E.[Unidade]" & Chr(13)
    sSQL = sSQL & ",      F.[FUNCAO]             = E.[Nome_Funcao]" & Chr(13)
    sSQL = sSQL & ",      F.[CR]                 = E.[Localização]" & Chr(13)
    sSQL = sSQL & ",      F.[SITUACAO]           = E.[Situacao]" & Chr(13)
    sSQL = sSQL & " WHERE F.[CR] = '' OR F.[CR] IS NULL"
    DoCmd.RunSQL sSQL

    
                    
    'Gera tabela final somente de beneficiarios (CERTIFICADOS <> 0000000)
    '(Rodrigo-19/01/2015) Acrescentei [NOME DO SEGURADO] no GROUP BY
    '(Rodrigo-20/03/2015) Acrescentei NOME_TIT e TIPO_DEP no GROUP BY
    sSQL = "SELECT"
        sSQL = sSQL & " [NÚMERO DA SUBFATURA] AS SUBFATURA"
        sSQL = sSQL & ", [EMPRESA]"
        sSQL = sSQL & ", [NÚMERO DO CERTIFICADO] AS CERTIFICADO"
        sSQL = sSQL & ", [COMPLEMENTO DO CERTIFICADO] AS [COMPLEMENTO CERTIFICADO]"
        sSQL = sSQL & ", MATRICULA"
        sSQL = sSQL & ", CPF"
        sSQL = sSQL & ", [NOME_TIT]"
        sSQL = sSQL & ", [NOME DO SEGURADO] AS [NOME SEGURADO/DEPENDENTE]"
        sSQL = sSQL & ", [CÓD GRAU PARENTESCO]"
        sSQL = sSQL & ", [UNID]"
        sSQL = sSQL & ", [FUNCAO]"
        sSQL = sSQL & ", [CR]"
        sSQL = sSQL & ", [CÓDIGO DO PLANO]"
        sSQL = sSQL & ", [SITUACAO]"
        sSQL = sSQL & ", '' AS [CO-PART]"
        sSQL = sSQL & ", '' AS [TX IMP / 2º VIA CART]"
        sSQL = sSQL & ", SUM([VALOR2]) AS [CUSTO R$]"
        sSQL = sSQL & ", SUM(FM.IOF) AS IOF"
        sSQL = sSQL & ", SUM([VALOR_FINAL]) AS [CUSTO COM IOF]"
    sSQL = sSQL & " INTO FINAL"
    sSQL = sSQL & " FROM FM"
    sSQL = sSQL & " GROUP BY [NÚMERO DA SUBFATURA], EMPRESA, [NÚMERO DO CERTIFICADO], [COMPLEMENTO DO CERTIFICADO], MATRICULA, CPF, [NOME_TIT], [CÓD GRAU PARENTESCO], [NOME DO SEGURADO], [UNID], [FUNCAO], [CR], [CÓDIGO DO PLANO], [SITUACAO]"
    DoCmd.RunSQL sSQL
    
    'Calculo IOF
    
    sSQL = "UPDATE FINAL"
    sSQL = sSQL & " SET IOF = [CUSTO R$] * 0.0238"
    
    DoCmd.RunSQL sSQL
    
    sSQL = "UPDATE FINAL"
    sSQL = sSQL & " SET [CUSTO COM IOF] = [CUSTO R$] + IOF"
    
    DoCmd.RunSQL sSQL
    
    
    '(Rodrigo-19/06/2015) Join FM X TAXAS para buscar taxas do segurado
    DoCmd.RunSQL "ALTER TABLE FINAL ALTER COLUMN [TX IMP / 2º VIA CART] DOUBLE"
    DoCmd.RunSQL "UPDATE FINAL A" & _
                " INNER JOIN TAXAS B ON B.CERTIFICADO = A.CERTIFICADO AND B.COMPLEMENTO = A.[COMPLEMENTO CERTIFICADO]" & _
                " SET A.[TX IMP / 2º VIA CART] = B.VALOR"
                
    '(Rodrigo-19/06/2015) Calcula IOF novamente para incluir valor das taxas
    DoCmd.RunSQL "UPDATE FINAL SET" & _
                " IOF = IOF + ([TX IMP / 2º VIA CART] * " & Replace(cIOF, ",", ".") & ")" & _
                ", [CUSTO COM IOF] = [CUSTO COM IOF] + [TX IMP / 2º VIA CART] + ([TX IMP / 2º VIA CART] * " & Replace(cIOF, ",", ".") & ")" & _
                " WHERE [TX IMP / 2º VIA CART] IS NOT NULL"
                
    'GERA TABELA COM OS VALORES DO CERTIFICADO 0000000
    
    sSQL = "SELECT"
    sSQL = sSQL & "     [SUBFATURA]"
    sSQL = sSQL & ",    [NOME SEGURADO/DEPENDENTE] AS DESCRICAO "
    sSQL = sSQL & ",    [CUSTO R$] AS [TOTAL SEM IOF]"
    sSQL = sSQL & "         INTO ACERTOS"
    sSQL = sSQL & "         FROM FINAL"
    sSQL = sSQL & "         WHERE CERTIFICADO = '0000000'"
    DoCmd.RunSQL sSQL
    
                
    '(Rodrigo-19/06/2015) Exclui lançamento original de cobranca de taxas e copay
        DoCmd.RunSQL "DELETE FROM FINAL" & _
                " WHERE CERTIFICADO = '0000000' "
                
    'Join FINAL x LOCALIZADOS
    sSQL = "UPDATE FINAL F"
    sSQL = sSQL & " INNER JOIN LOCALIZADOS L ON F.[CERTIFICADO] = L.[CERTIFICADO] AND F.[SUBFATURA] = L.[SUBFATURA]"
    sSQL = sSQL & " SET"
    sSQL = sSQL & "  F.[MATRICULA]      = L.[RE]"
    sSQL = sSQL & ", F.[UNID]           = L.[UNID]"
    sSQL = sSQL & ", F.[FUNCAO]         = L.[FUNCAO]"
    sSQL = sSQL & ", F.[CR]             = L.[CR]"
    sSQL = sSQL & ", F.[SITUACAO]       = L.[SITUACAO]"
    sSQL = sSQL & " WHERE F.MATRICULA = '' OR F.MATRICULA IS NULL"
    DoCmd.RunSQL sSQL
    
    sSQL = "UPDATE FINAL F"
    sSQL = sSQL & " INNER JOIN LOCALIZADOS L ON F.[CERTIFICADO] = L.[CERTIFICADO] AND F.[SUBFATURA] = L.[SUBFATURA]"
    sSQL = sSQL & " SET"
    sSQL = sSQL & "  F.[NOME_TIT]       = L.[NOME DO TITULAR]"
    sSQL = sSQL & ", F.[CPF]            = CDBL(L.[CPF])"
    sSQL = sSQL & " WHERE F.[NOME_TIT] = '' OR F.[NOME_TIT] IS NULL"
    DoCmd.RunSQL sSQL
                
        'Gera não localizados
    sSQL = "SELECT DISTINCT"
        sSQL = sSQL & "  SUBFATURA"
        sSQL = sSQL & ", CERTIFICADO"
        sSQL = sSQL & ", CPF"
        sSQL = sSQL & ", [NOME_TIT] AS [NOME DO TITULAR]"
        sSQL = sSQL & ", [NOME SEGURADO/DEPENDENTE] AS [NOME DO SEGURADO]"
        sSQL = sSQL & ", MATRICULA AS RE"
        sSQL = sSQL & ", UNID"
        sSQL = sSQL & ", FUNCAO"
        sSQL = sSQL & ", CR"
        sSQL = sSQL & ", SITUACAO"
    sSQL = sSQL & " INTO ND"
    sSQL = sSQL & " FROM FINAL"
    sSQL = sSQL & " WHERE MATRICULA = '' OR MATRICULA IS NULL"
    DoCmd.RunSQL sSQL
           
   
    'Gera relatório final
    sSQL = "SELECT"
        sSQL = sSQL & "     SUBFATURA"
        sSQL = sSQL & ",    EMPRESA"
        sSQL = sSQL & ",    CPF"
        sSQL = sSQL & ",    MATRICULA AS RE"
        sSQL = sSQL & ",    UNID"
        sSQL = sSQL & ",    NOME_TIT AS NOME"
        sSQL = sSQL & ",    FUNCAO"
        sSQL = sSQL & ",    CR"
        sSQL = sSQL & ",    COUNT([NOME SEGURADO/DEPENDENTE]) AS [QT]"
        sSQL = sSQL & ",    [CÓDIGO DO PLANO] AS PLANO"
        sSQL = sSQL & ",    [CÓDIGO DO PLANO] AS [BDESC]"
        sSQL = sSQL & ",    SUM([CUSTO R$]) AS [C_CIA]"
        sSQL = sSQL & ",    SITUACAO"
    sSQL = sSQL & "     INTO FINAL2"
    sSQL = sSQL & "     FROM FINAL"
    sSQL = sSQL & "     GROUP BY SUBFATURA, EMPRESA, CPF, MATRICULA, UNID, NOME_TIT, FUNCAO, CR, [CÓDIGO DO PLANO], [CÓDIGO DO PLANO], SITUACAO"
    DoCmd.RunSQL sSQL
    
'    'Gera resumo por centro de custo
'    sSQL = "SELECT"
'        sSQL = sSQL & "     SUBFATURA"
'        sSQL = sSQL & ",    CR"
'        sSQL = sSQL & ",    SUM([CUSTO COM IOF]) AS [TOTAL]"
'        sSQL = sSQL & ",    'FATURAMENTO' AS [TIPO DE LANCAMENTO]"
'    sSQL = sSQL & "     INTO RESUMO"
'    sSQL = sSQL & "     FROM FINAL"
'    sSQL = sSQL & "     GROUP BY"
'        sSQL = sSQL & "     SUBFATURA"
'        sSQL = sSQL & ",    CR"
'    DoCmd.RunSQL sSQL
    
'    'Insere lançamentos de coparticipação na tabela resumo
'        sSQL = "INSERT INTO RESUMO"
'        sSQL = sSQL & "     (SUBFATURA"
'        sSQL = sSQL & ",    CR"
'        sSQL = sSQL & ",    [TOTAL]"
'        sSQL = sSQL & ",    [TIPO DE LANCAMENTO])"
'    sSQL = sSQL & "     SELECT"
'        sSQL = sSQL & "     [SUB]"
'        sSQL = sSQL & ",    [C_CUSTO]"
'        sSQL = sSQL & ",    [VALOR TOTAL]"
'        sSQL = sSQL & ",    'COPARTICIPACAO'"
'    sSQL = sSQL & "     FROM COPA"
'    DoCmd.RunSQL sSQL
    
'    'atualiza campos subfatura da tabela resumo
'    DoCmd.RunSQL "UPDATE RESUMO SET SUBFATURA = FORMAT([SUBFATURA], '0000')"
    
    DoCmd.SetWarnings True
        
End Sub

Private Sub ProcessamentoCusto()

Dim sSQL As String
    
    DoCmd.SetWarnings False
    
        'Cria campos
        DoCmd.RunSQL "ALTER TABLE FINAL2 ADD COLUMN CUSTO DOUBLE"
        DoCmd.RunSQL "ALTER TABLE FINAL2 ADD COLUMN DESCONTO DOUBLE"
        
        'Atualiza campos
        DoCmd.RunSQL "UPDATE FINAL2 SET CUSTO = 0"
        DoCmd.RunSQL "UPDATE FINAL2 SET DESCONTO = 0"
    
        'Altera campos
        DoCmd.RunSQL "ALTER TABLE FINAL2            ALTER COLUMN [RE] TEXT"
        DoCmd.RunSQL "ALTER TABLE CONF_HDI          ALTER COLUMN [Matrícula ] TEXT"
        DoCmd.RunSQL "ALTER TABLE CONF_GLOBAL       ALTER COLUMN [Matrícula ] TEXT"
        DoCmd.RunSQL "ALTER TABLE CONF_SANTANDER    ALTER COLUMN [Matrícula ] TEXT"
        DoCmd.RunSQL "ALTER TABLE CONF_HDI          ALTER COLUMN [Bradesco - Custo ] DOUBLE"
        DoCmd.RunSQL "ALTER TABLE CONF_GLOBAL       ALTER COLUMN [Bradesco - Custo] DOUBLE"
        DoCmd.RunSQL "ALTER TABLE CONF_SANTANDER    ALTER COLUMN [Bradesco - Custo] DOUBLE"
        DoCmd.RunSQL "ALTER TABLE CONF_HDI          ALTER COLUMN [Bradesco - Desconto ] DOUBLE"
        DoCmd.RunSQL "ALTER TABLE CONF_GLOBAL       ALTER COLUMN [Bradesco - Desconto] DOUBLE"
        DoCmd.RunSQL "ALTER TABLE CONF_SANTANDER    ALTER COLUMN [Bradesco - Desconto] DOUBLE"
        
        'Padroniza matrícula em todas as tabelas
        DoCmd.RunSQL "UPDATE FINAL2         SET [RE]            = FORMAT([RE], '0')"
        DoCmd.RunSQL "UPDATE CONF_HDI       SET [Matrícula ]    = FORMAT([Matrícula ], '0')"
        DoCmd.RunSQL "UPDATE CONF_GLOBAL    SET [Matrícula ]    = FORMAT([Matrícula ], '0')"
        DoCmd.RunSQL "UPDATE CONF_SANTANDER SET [Matrícula ]    = FORMAT([Matrícula ], '0')"
    
        'Join FINAL2 x CONF_HDI (CUSTO, DESCONTO)
        sSQL = "UPDATE FINAL2 F"
        sSQL = sSQL & " INNER JOIN CONF_HDI C ON"
        sSQL = sSQL & "  F.[RE] = C.[Matrícula ]"
        sSQL = sSQL & " SET"
        sSQL = sSQL & "  F.CUSTO = C.[Bradesco - Custo ]"
        sSQL = sSQL & ", F.DESCONTO = C.[Bradesco - Desconto ]"
        sSQL = sSQL & " WHERE F.EMPRESA = 'HDI'"
        DoCmd.RunSQL sSQL
    
        'Join FINAL2 x CONF para buscar o custo e desconto previsto pela GLOBAL
        sSQL = "UPDATE FINAL2 F"
        sSQL = sSQL & " INNER JOIN CONF_GLOBAL C ON"
        sSQL = sSQL & " F.[RE] = C.[Matrícula ]"
        sSQL = sSQL & " SET"
        sSQL = sSQL & "  F.CUSTO = C.[Bradesco - Custo]"
        sSQL = sSQL & ", F.DESCONTO = C.[Bradesco - Desconto]"
        sSQL = sSQL & " WHERE F.EMPRESA = 'GLOBAL'"
        DoCmd.RunSQL sSQL
    
        'Join FINAL2 x CONF para buscar o custo e desconto previsto pela SANTANDER
        sSQL = "UPDATE FINAL2 F"
        sSQL = sSQL & " INNER JOIN CONF_SANTANDER C ON"
        sSQL = sSQL & " F.[RE] = C.[Matrícula ]"
        sSQL = sSQL & " SET"
        sSQL = sSQL & "  F.CUSTO = C.[Bradesco - Custo]"
        sSQL = sSQL & ", F.DESCONTO = C.[Bradesco - Desconto]"
        sSQL = sSQL & " WHERE F.EMPRESA = 'SANTANDER'"
        DoCmd.RunSQL sSQL
        
        'Gera tabela de custos não localizados na folha
        sSQL = " SELECT *"
        sSQL = sSQL & " INTO CUSTO_ND"
        sSQL = sSQL & " FROM FINAL2"
        sSQL = sSQL & " WHERE [CUSTO] = 0"
        DoCmd.RunSQL sSQL
    

    DoCmd.SetWarnings True
        
End Sub

Private Sub ExportaArquivo()

Dim sFileName As String         'nome do arquivo final
Dim sFileRateio As String
Dim sFileTemplate As String
Dim sSubfatura As String
Dim db As Database
Dim rs As Recordset
Dim i As Integer
    
    DoCmd.SetWarnings False
    
    '==============================================================================
    'EXPORTA RATEIO
    '==============================================================================
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT DISTINCT SUBFATURA FROM FINAL2")
    
    Do Until rs.EOF
    
    sSubfatura = rs!Subfatura
    
    sFileName = FOLDERDESTNAME & "\HDI(" & rs!Subfatura & ")-Bradesco(Saude)_Rateio(Parcial)_(" & dataref & ").xlsx"
    DoCmd.RunSQL "SELECT * INTO RATEIO_" & rs!Subfatura & " FROM FINAL2 WHERE SUBFATURA = '" & rs!Subfatura & "'"
    DoCmd.RunSQL "SELECT * INTO ACERTOS_" & rs!Subfatura & " FROM ACERTOS WHERE SUBFATURA = '" & rs!Subfatura & "'"
    DoCmd.RunSQL "SELECT * INTO CUSTO_ND_" & rs!Subfatura & " FROM CUSTO_ND WHERE SUBFATURA = '" & rs!Subfatura & "'"
    DoCmd.TransferSpreadsheet acExport, , "RATEIO_" & rs!Subfatura, sFileName, , "RATEIO"
    DoCmd.TransferSpreadsheet acExport, , "ACERTOS_" & rs!Subfatura, sFileName, , "ACERTOS"
    DoCmd.TransferSpreadsheet acExport, , "CUSTO_ND_" & rs!Subfatura, sFileName, , "CUSTO_ND"
    Call Format_Excel(sFileName)
    
    'COPIA O TEMPLATE
    sFileRateio = FOLDERDESTNAME & "\HDI(" & rs!Subfatura & ")-Bradesco(Saude)_Rateio_(" & dataref & ").xlsm"
    sFileTemplate = FOLDERORNAME & "\Arquivos\Template\Template-Bradesco.xlsm"
    FileCopy sFileTemplate, sFileRateio
    
    Call AddFormulas(sFileRateio, sFileName, sSubfatura)
    
    rs.MoveNext
    Loop
    
   '==============================================================================
   'EXPORTA NÃO LOCALIZADOS
   '==============================================================================
    If DCount("CERTIFICADO", "ND") > 0 Then
        sFileName = FOLDERDESTNAME & "\HDI-Bradesco(Saude)_Rateio_NaoLocalizados_(" & dataref & ").xlsx"
        DoCmd.TransferSpreadsheet acExport, , "ND", sFileName, , "NaoLocalizados"
        Call Format_Excel(sFileName)
    End If
    
    DoCmd.SetWarnings True
        
End Sub


Private Sub AddFormulas(sFileRateio As String, sFileSource As String, sSubfatura As String)

Dim appXls As Excel.Application
Dim i As Integer
Dim linha As Integer

If FileExists(sFileRateio) = False Then
    MsgBox "Arquivo " & sFileRateio & " não encontrado!!!", vbCritical
    Exit Sub
End If

If FileExists(sFileSource) = False Then
    MsgBox "Arquivo " & sFileSource & " não encontrado!!!", vbCritical
    Exit Sub
End If

Set appXls = New Excel.Application
With appXls
    .Workbooks.Open filename:=sFileRateio
    .Workbooks.Open filename:=sFileSource
    
    linha = .Workbooks(1).Sheets(2).Cells(.Rows.Count, 1).End(xlUp).Row + 1
    
    'atualiza formulas do detalhado
    For i = 2 To .Workbooks(2).Sheets("RATEIO").Cells(.Rows.Count, 1).End(xlUp).Row
        .Workbooks(1).Sheets(2).Rows("" & linha & ":" & linha & "").Insert Shift:=xlDown
        .Workbooks(1).Sheets(2).Cells(linha, 1) = .Workbooks(2).Sheets("RATEIO").Cells(i, 4)
        .Workbooks(1).Sheets(2).Cells(linha, 2).FormulaR1C1 = "=VLOOKUP(RC[-1],'[HDI(" & sSubfatura & ")-Bradesco(Saude)_Rateio(Parcial)_(" & dataref & ").xlsx]RATEIO'!R1C4:R65533C13,2,0)"
        .Workbooks(1).Sheets(2).Cells(linha, 3).FormulaR1C1 = "=VLOOKUP(RC[-2],'[HDI(" & sSubfatura & ")-Bradesco(Saude)_Rateio(Parcial)_(" & dataref & ").xlsx]RATEIO'!R1C4:R65533C13,3,0)"
        .Workbooks(1).Sheets(2).Cells(linha, 4).FormulaR1C1 = "=VLOOKUP(RC[-3],'[HDI(" & sSubfatura & ")-Bradesco(Saude)_Rateio(Parcial)_(" & dataref & ").xlsx]RATEIO'!R1C4:R65533C13,4,0)"
        .Workbooks(1).Sheets(2).Cells(linha, 5).FormulaR1C1 = "=VLOOKUP(RC[-4],'[HDI(" & sSubfatura & ")-Bradesco(Saude)_Rateio(Parcial)_(" & dataref & ").xlsx]RATEIO'!R1C4:R65533C13,5,0)"
        .Workbooks(1).Sheets(2).Cells(linha, 6).FormulaR1C1 = "=VLOOKUP(RC[-5],'[HDI(" & sSubfatura & ")-Bradesco(Saude)_Rateio(Parcial)_(" & dataref & ").xlsx]RATEIO'!R1C4:R65533C13,6,0)"
        .Workbooks(1).Sheets(2).Cells(linha, 7).FormulaR1C1 = "=VLOOKUP(RC[-6],'[HDI(" & sSubfatura & ")-Bradesco(Saude)_Rateio(Parcial)_(" & dataref & ").xlsx]RATEIO'!R1C4:R65533C13,7,0)"
        .Workbooks(1).Sheets(2).Cells(linha, 8).FormulaR1C1 = "=VLOOKUP(RC[-7],'[HDI(" & sSubfatura & ")-Bradesco(Saude)_Rateio(Parcial)_(" & dataref & ").xlsx]RATEIO'!R1C4:R65533C13,8,0)"
        .Workbooks(1).Sheets(2).Cells(linha, 9).FormulaR1C1 = "=VLOOKUP(RC[-8],'[HDI(" & sSubfatura & ")-Bradesco(Saude)_Rateio(Parcial)_(" & dataref & ").xlsx]RATEIO'!R1C4:R65533C13,9,0)"
        .Workbooks(1).Sheets(2).Cells(linha, 10).FormulaR1C1 = "=RC[-1]*0.0238"
        .Workbooks(1).Sheets(2).Cells(linha, 11).FormulaR1C1 = "=RC[-2]+RC[-1]"
        .Workbooks(1).Sheets(2).Cells(linha, 12).FormulaR1C1 = ""
        .Workbooks(1).Sheets(2).Cells(linha, 13).FormulaR1C1 = "=VLOOKUP(RC[-6],R1C7:R6C10,4,0)*RC[-7]"
        .Workbooks(1).Sheets(2).Cells(linha, 14).FormulaR1C1 = "=RC[-3]+RC[-1]"
        .Workbooks(1).Sheets(2).Cells(linha, 15).FormulaR1C1 = "=(VLOOKUP(RC8,R1C7:R6C9,2,FALSE)*RC[-9])-RC[-6]"
        .Workbooks(1).Sheets(2).Cells(linha, 16).FormulaR1C1 = "=RC[-5]"
        .Workbooks(1).Sheets(2).Cells(linha, 17).FormulaR1C1 = "=VLOOKUP(RC[-10],R1C7:R6C10,3,0)*RC[-11]"
        .Workbooks(1).Sheets(2).Cells(linha, 18).FormulaR1C1 = "=IF(AND(RC[-2]-RC[-1]>-0.05,RC[-2]-RC[-1]<0.05),0,RC[-2]-RC[-1])"
        .Workbooks(1).Sheets(2).Cells(linha, 19).FormulaR1C1 = "=VLOOKUP(RC[-18],'[HDI(" & sSubfatura & ")-Bradesco(Saude)_Rateio(Parcial)_(" & dataref & ").xlsx]RATEIO'!C4:C15,11,0)"
        .Workbooks(1).Sheets(2).Cells(linha, 20).FormulaR1C1 = "=VLOOKUP(RC[-19],'[HDI(" & sSubfatura & ")-Bradesco(Saude)_Rateio(Parcial)_(" & dataref & ").xlsx]RATEIO'!C4:C15,12,0)"
        .Workbooks(1).Sheets(2).Cells(linha, 21).FormulaR1C1 = "=IF(AND(RC[-10]-RC[-2]>-0.05,RC[-10]-RC[-2]<0.05),0,RC[-10]-RC[-2])"
        .Workbooks(1).Sheets(2).Cells(linha, 22).FormulaR1C1 = "=RC[-7]-RC[-2]"
        .Workbooks(1).Sheets(2).Cells(linha, 23).FormulaR1C1 = "=VLOOKUP(RC[-22],'[HDI(" & sSubfatura & ")-Bradesco(Saude)_Rateio(Parcial)_(" & dataref & ").xlsx]RATEIO'!R1C4:R65533C13,10,0)"
        linha = linha + 1
    Next i
    
    'atualiza base rateio
    For i = 2 To .Workbooks(2).Sheets("RATEIO").Cells(.Rows.Count, 1).End(xlUp).Row
        .Workbooks(1).Sheets("BaseRateio").Cells(i, 1) = .Workbooks(2).Sheets("RATEIO").Cells(i, 8)
        .Workbooks(1).Sheets("BaseRateio").Cells(i, 2) = .Workbooks(2).Sheets("RATEIO").Cells(i, 12) * 1.0238
    Next i
    
    'atualiza totais do detalhado
    .Workbooks(1).Sheets(2).Cells(linha, 6).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    .Workbooks(1).Sheets(2).Cells(linha, 9).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    .Workbooks(1).Sheets(2).Cells(linha, 10).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    .Workbooks(1).Sheets(2).Cells(linha, 11).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    .Workbooks(1).Sheets(2).Cells(linha, 12).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    .Workbooks(1).Sheets(2).Cells(linha, 13).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    .Workbooks(1).Sheets(2).Cells(linha, 14).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    .Workbooks(1).Sheets(2).Cells(linha, 15).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    .Workbooks(1).Sheets(2).Cells(linha, 16).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    .Workbooks(1).Sheets(2).Cells(linha, 17).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    .Workbooks(1).Sheets(2).Cells(linha, 18).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    .Workbooks(1).Sheets(2).Cells(linha, 19).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    .Workbooks(1).Sheets(2).Cells(linha, 20).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    .Workbooks(1).Sheets(2).Cells(linha, 21).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    .Workbooks(1).Sheets(2).Cells(linha, 22).FormulaR1C1 = "=SUM(R[-1]C:R[-" & linha - 10 & "]C)"
    
    'Atualiza nome da aba detalhado com base na data ref
    Select Case Month(DateAdd("m", 1, "01/" & Right(dataref, 2) & "/" & Left(dataref, 4)))
        Case 1
        .Workbooks(1).Sheets(2).Name = "Janeiro"
        Case 2
        .Workbooks(1).Sheets(2).Name = "Fevereiro"
        Case 3
        .Workbooks(1).Sheets(2).Name = "Março"
        Case 4
        .Workbooks(1).Sheets(2).Name = "Abril"
        Case 5
        .Workbooks(1).Sheets(2).Name = "Maio"
        Case 6
        .Workbooks(1).Sheets(2).Name = "Junho"
        Case 7
        .Workbooks(1).Sheets(2).Name = "Julho"
        Case 8
        .Workbooks(1).Sheets(2).Name = "Agosto"
        Case 9
        .Workbooks(1).Sheets(2).Name = "Setembro"
        Case 10
        .Workbooks(1).Sheets(2).Name = "Outubro"
        Case 11
        .Workbooks(1).Sheets(2).Name = "Novembro"
        Case 12
        .Workbooks(1).Sheets(2).Name = "Dezembro"
    End Select
    
    'define range de formatação
    With .Workbooks(1).Sheets(2)
        With .Range(.Cells(10, 1), .Cells(linha - 1, 24))
            
            'formata alinhamento e bold
            .Font.Bold = False
            .Font.Size = 8
            .Interior.ColorIndex = 0
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlBottom
            
            'formata grade das celulas
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
    End With
    
    'formata largura das colunas como auto fit
    'appXls.Columns.AutoFit
    
    'posiciona cursor para celula A1 de todas as sheets
    i = .Workbooks(1).Sheets.Count
    Do Until i = 0
        .Workbooks(1).Sheets(i).Activate
        .Cells(1, 1).Activate
        i = i - 1
    Loop
    
    For i = 1 To .Workbooks.Count
    .Workbooks(1).Save
    .Workbooks(1).Close
    Next i
    .Quit
End With

Set appXls = Nothing

End Sub



