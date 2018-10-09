


Imports SeleniumWrapper.WebDriver
Imports Microsoft.Office.Interop.Excel.XlFileFormat
Imports Microsoft.Office.Interop.Excel.XlDirection
Imports Microsoft.Office.Interop.Excel.XlCellType
Imports Microsoft.Office.Interop.Excel.OLEObjectClass
Imports Microsoft.VisualBasic









Module Module1

    Sub Main()

        Dim driver As New SeleniumWrapper.WebDriver
        Dim Elemento As SeleniumWrapper.WebElement
        Dim objPlan As Object



        Dim Descricao As String
        Dim DtVencimento As String
        Dim TaxaRend As String
        Dim ValorMin As String
        Dim PrecUni As String


        '' inicia aqui a planilha 

        objPlan = CreateObject("Excel.Application")


        If (ExisteArquivo("TJ" + Today.Year.ToString + ".xlsx") = 0) Then

            IniciarDados(CreateObject("Excel.Application"))
            ''o Arquivo nao existe ele cri a planilha e fecha

        End If









        driver.start("chrome")
        driver.get("http://www.tesouro.fazenda.gov.br/tesouro-direto-precos-e-taxas-dos-titulos")
        driver.setImplicitWait(2000)

        'Descricao'
        Elemento = driver.findElementByCssSelector("#p_p_id_precosetaxas_WAR_tesourodiretoportlet_ > div > div > div > table.tabelaPrecoseTaxas > tbody > tr:nth-child(11) > td.listing0")
        Descricao = Elemento.Text

        'Data Vencimento'
        Elemento = driver.findElementByCssSelector("#p_p_id_precosetaxas_WAR_tesourodiretoportlet_ > div > div > div > table.tabelaPrecoseTaxas > tbody > tr:nth-child(11) > td:nth-child(2)")
        DtVencimento = Elemento.Text

        'Taxa Rendimento'
        Elemento = driver.findElementByCssSelector("#p_p_id_precosetaxas_WAR_tesourodiretoportlet_ > div > div > div > table.tabelaPrecoseTaxas > tbody > tr:nth-child(11) > td:nth-child(3)")
        TaxaRend = Elemento.Text

        'valor Minimo'
        Elemento = driver.findElementByCssSelector("#p_p_id_precosetaxas_WAR_tesourodiretoportlet_ > div > div > div > table.tabelaPrecoseTaxas > tbody > tr:nth-child(11) > td:nth-child(4)")
        ValorMin = Elemento.Text

        'Preço Unitario'
        Elemento = driver.findElementByCssSelector("#p_p_id_precosetaxas_WAR_tesourodiretoportlet_ > div > div > div > table.tabelaPrecoseTaxas > tbody > tr:nth-child(11) > td:nth-child(5)")
        PrecUni = Elemento.Text


        Dim Params() As String = {Descricao, DtVencimento, TaxaRend, ValorMin, PrecUni}


        RegDados(Params, objPlan)

        EncerraPlan(objPlan)


        EncerraNav(driver)



    End Sub


    Sub RegDados(Params() As String, xlApp As Object)

        Dim NomePlanilha As String
        Dim Indice As Integer


        NomePlanilha = "TJ" + Today.Year.ToString + ".xlsx"


        xlApp.Workbooks.Open(Filename:="D:\Multimidia\documentos\Orçamento\Investimento\" + NomePlanilha)
        xlApp.Visible = True

        xlApp.ActiveWorkbook.Worksheets(1).Range("A1").Select


        Indice = xlApp.ActiveWorkbook.Worksheets(1).Range("A1", xlApp.ActiveWorkbook.Worksheets(1).Range("A" & xlApp.ActiveWorkbook.Worksheets(1).Rows.Count).End(xlUp)).Rows.Count

        Indice = Indice + 1

        xlApp.ActiveWorkbook.Worksheets(1).Columns("A:F").Rows(Indice).Font.Bold = False

        xlApp.ActiveWorkbook.Worksheets(1).Cells(Indice, 1).Value = Params(0)
        xlApp.ActiveWorkbook.Worksheets(1).Cells(Indice, 2).Value = Params(1)
        xlApp.ActiveWorkbook.Worksheets(1).Cells(Indice, 3).Value = Params(2)
        xlApp.ActiveWorkbook.Worksheets(1).Cells(Indice, 4).Value = Params(3)
        xlApp.ActiveWorkbook.Worksheets(1).Cells(Indice, 5).Value = Params(4)
        xlApp.ActiveWorkbook.Worksheets(1).Cells(Indice, 6).Value = Today


    End Sub



    Public Function ExisteArquivo(NomeArquivo As String) As Integer

        Dim caminhoArquivo As String

        caminhoArquivo = "D:\Multimidia\documentos\Orçamento\Investimento\" + NomeArquivo

        If (Dir(caminhoArquivo) = vbNullString) Then
            Return 0
        Else
            Return 1
        End If

    End Function



    Sub IniciarDados(xlApp As Object)


        Dim NomePlanilha As String

        NomePlanilha = "TJ" + Today.Year.ToString + ".xlsx"
        xlApp.Visible = False
        xlApp.Workbooks.add


        xlApp.ActiveWorkbook.Worksheets(1).Columns("A:F").ColumnWidth = 23.71
        xlApp.ActiveWorkbook.Worksheets(1).Columns("A:F").Font.Bold = True

        xlApp.ActiveWorkbook.Worksheets(1).Cells(1, 1).Value = "Descricao"
        xlApp.ActiveWorkbook.Worksheets(1).Cells(1, 2).Value = "Data Vencimento"
        xlApp.ActiveWorkbook.Worksheets(1).Cells(1, 3).Value = "Taxa Rendimento"
        xlApp.ActiveWorkbook.Worksheets(1).Cells(1, 4).Value = "valor Minimo"
        xlApp.ActiveWorkbook.Worksheets(1).Cells(1, 5).Value = "Taxa Rendimento"
        xlApp.ActiveWorkbook.Worksheets(1).Cells(1, 6).Value = "Data Execução"

        xlApp.ActiveWorkbook.Worksheets(1).Range("A1").Select


        xlApp.ActiveWorkbook.saveAs(Filename:="D:\Multimidia\documentos\Orçamento\Investimento\" + NomePlanilha, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False)


        xlApp.ActiveWorkbook.Close(SaveChanges:=False)
        xlApp.Quit()

    End Sub


    Sub EncerraPlan(xlApp As Object)

        xlApp.ActiveWorkbook.Close(SaveChanges:=True)
    End Sub

    Sub EncerraNav(browser As Object)
        browser.close()
        Shell("cmd.exe /c " & "Taskkill /IM chromedriver.exe /F ")
    End Sub



End Module
