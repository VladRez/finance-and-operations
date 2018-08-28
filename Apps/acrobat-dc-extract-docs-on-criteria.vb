Private Sub main()
        Dim gApp As Acrobat.CAcroApp
        Dim gPDDoc As Acrobat.CAcroPDDoc
        Dim gTarget As Acrobat.CAcroPDDoc
        Dim jso As Object
        gApp = CreateObject("AcroExch.App")
        gPDDoc = CreateObject("AcroExch.PDDoc")
        gPDDoc.Open(TextBox1.Text.ToString())
        jso = gPDDoc.GetJSObject



        ''''VARIALBES''''''
        Dim InvoiceNum As Integer
        Dim TotalPageNum As Integer

        Dim WordCount As Integer

        Dim currentWord As String
        Dim tempPageCount As Integer

        TotalPageNum = gPDDoc.GetNumPages
        ''Console.WriteLine("TOTAL # OF PAGES " + TotalPageNum.ToString)
        Label5.Text = TotalPageNum.ToString()

        For curPage As Integer = 0 To TotalPageNum - 1
            gTarget = CreateObject("AcroExch.PDDoc")

            For WordPos As Integer = 0 To 50
                currentWord = jso.getPageNthWord(curPage, WordPos, True)
                If currentWord = "CRITERIA" Or currentWord = "CRITERIA" Then
                    InvoiceNum = jso.getPageNthWord(curPage, WordPos + 2, True)

                    ' Console.WriteLine("Invoice: " + InvoiceNum.ToString)

                    tempPageCount = jso.getPageNthWord(curPage, WordPos + 10, True)

                    ' Console.WriteLine("Current Page: " + curPage.ToString)
                    ' Console.WriteLine("Page Range: " + tempPageCount.ToString)
                    gTarget.Create()

                    ' Console.WriteLine(".........Creating doc..............")

                    ''Console.WriteLine("PAGE: " + curPage.ToString + " out of " + TotalPageNum.ToString + " - Invoice: " + InvoiceNum.ToString)
                    Label5.Text = "PAGE: " + (curPage + 1).ToString + " out of " + TotalPageNum.ToString + " - Invoice: " + InvoiceNum.ToString
                    gTarget.InsertPages(-1, gPDDoc, curPage, tempPageCount, False)
                    ' Console.WriteLine("Begin index: " + curPage.ToString)
                    'Console.WriteLine("End index: " + tempPageCount.ToString)
                    gTarget.Save(&H1, TextBox2.Text.ToString() & "\" & InvoiceNum.ToString & ".pdf")
                    curPage = curPage - 1 + tempPageCount
                    'Console.WriteLine("Moving to page: " + curPage.ToString)
                    Exit For
                End If


            Next WordPos
            gTarget.Close()
        Next curPage
        Label5.Text = "DONE!"
    End Sub