Imports System.Net, System.IO, System.Text
Imports firefox = OpenQA.Selenium.Firefox
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Support
Imports OpenQA.Selenium.Support.UI
Imports OpenQA.Selenium.Firefox
Imports EC = SeleniumExtras.WaitHelpers.ExpectedConditions
Imports EX = Microsoft.Office.Interop.Excel


Public Class Form3
    Private Sub B_Label_Loader_Click(sender As Object, e As EventArgs) Handles B_Label_Loader.Click
        If TB_Label_Folder.Text = "Click to select folder" Then
            MsgBox("Please choose output folder!")
            Exit Sub
        End If

        'Dim xlApp As New EX.Application

        'Try
        '    Dim ffprofile As New FirefoxOptions
        '    ffprofile.SetPreference("browser.download.folderList", 2)
        '    ffprofile.SetPreference("browser.download.manager.showWhenStarting", False)
        '    ffprofile.SetPreference("pdfjs.disabled", True)
        '    ffprofile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/pdf, application/octet-stream")
        '    ffprofile.SetPreference("browser.download.dir", TB_Label_Folder.Text)
        '    Dim ff_Driver As New FirefoxDriver(ffprofile)

        '    Dim input As New OpenFileDialog
        '    input.Filter = "Excel Files (*.xlsx)|*.xlsx"
        '    input.ShowDialog()



        '    Dim xlBook = xlApp.Workbooks.Open(input.FileName)
        '    Dim xltab1 = xlBook.Worksheets("DOWNLOAD")
        '    Dim xlUp = EX.XlDirection.xlUp
        '    Dim lastentry As Integer
        '    Dim add As String
        '    lastentry = xltab1.Range("A" & xltab1.Rows.Count).End(xlUp).Row
        '    'lastentry = xltab1.Range("A1:A" & lastentry).Value

        '    For row As Integer = 2 To lastentry
        '        add = xltab1.Range("A" & row).Value
        '        ff_Driver.Navigate.GoToUrl("https://eprel.ec.europa.eu/api/products/tyres/" + add + "/labels")


        '    Next

        '    ff_Driver.Close()


        '    'ff_Driver.Navigate.GoToUrl("https://eprel.ec.europa.eu/api/products/tyres/381324/labels?format=PDF")

        'Catch ex As Exception
        '    xlApp.Quit()
        '    Exit Sub
        'End Try


        'xlApp.Quit()


        Dim request As WebRequest = WebRequest.Create("https://eprel.ec.europa.eu/api/products/tyres/381324/labels?format=PDF")
        request.Credentials = CredentialCache.DefaultCredentials

        Dim response As WebResponse = request.GetResponse()
        Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
        Console.WriteLine(response)

        Using dataStream As Stream = response.GetResponseStream()
            Dim reader As New StreamReader(dataStream)
            Dim responsefromServer As String = reader.ReadToEnd()
            Console.WriteLine(responsefromServer)
        End Using



        'Dim webClient As New System.Net.WebClient
        'Dim result As String = webClient.DownloadString("https://eprel.ec.europa.eu/api/products/tyres/381324/labels?noRedirect=true&format=PDF")
        'Console.WriteLine(result)

        response.Close()


    End Sub



    Private Sub TB_Label_Folder_Click(sender As Object, e As EventArgs) Handles TB_Label_Folder.Click
        Dim folder As New FolderBrowserDialog()
        'folder.ShowDialog()
        'TB_Label_Folder.Text = folder.SelectedPath + "\"

        If folder.ShowDialog = DialogResult.Cancel Then
            MsgBox("Please select folder!")
        Else
            TB_Label_Folder.Text = folder.SelectedPath + "\"
        End If

    End Sub


End Class