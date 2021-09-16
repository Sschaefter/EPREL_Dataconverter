Imports System.Net
Imports System.IO
Imports System.Text
Imports firefox = OpenQA.Selenium.Firefox
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Support
Imports OpenQA.Selenium.Support.UI
Imports OpenQA.Selenium.Firefox
Imports EC = SeleniumExtras.WaitHelpers.ExpectedConditions
Imports EX = Microsoft.Office.Interop.Excel


Public Class Form3
    Private Async Sub B_Label_Loader_Click(sender As Object, e As EventArgs) Handles B_Label_Loader.Click
        If TB_Label_Folder.Text = "Click to select folder" Then
            MsgBox("Please choose output folder!")
            Exit Sub
        End If


        Dim pdfBytes As Byte()
        Dim pdfFilePath As String
        Dim xlApp As New EX.Application

        Dim input As New OpenFileDialog
        input.Filter = "Excel Files (*.xlsx)|*.xlsx"
        input.ShowDialog()



        Dim xlBook = xlApp.Workbooks.Open(input.FileName)
        Dim xltab1 = xlBook.Worksheets("DOWNLOAD")
        Dim xlUp = EX.XlDirection.xlUp
        Dim lastentry As Integer
        Dim add As String
        Dim LB_COL As String = "COLOR"
        Dim LB_SZ As String = "BIG"
        Dim LB_FF As String = "PDF"

        If RB_BW.Checked = True Then
            LB_COL = "BW"
        Else
            LB_COL = "COLOR"
        End If

        If RB_BIG.Checked = True Then
            LB_SZ = "BIG"
        Else
            LB_SZ = "SMALL"
        End If

        If RB_PDF.Checked = True Then
            LB_FF = "PDF"
        ElseIf RB_PNG.Checked = True Then
            LB_FF = "PNG"
        Else
            LB_FF = "SVG"
        End If


        lastentry = xltab1.Range("B" & xltab1.Rows.Count).End(xlUp).Row
        'lastentry = xltab1.Range("A" & xltab1.Rows.Count).End(xlUp).Row
        'lastentry = xltab1.Range("A1:A" & lastentry).Value




        Try
            For row As Integer = 2 To lastentry
                add = xltab1.Range("A" & row).Value
                If CBox_Zipall_Label.Checked = False Then
                    pdfBytes = Await GetPDFResourceAsync(New Uri("https://eprel.ec.europa.eu/api/products/lightsources/" & add & "/labels?format=" & LB_FF & "&type=" & LB_SZ & "_" & LB_COL))
                    pdfFilePath = Path.Combine(TB_Label_Folder.Text, "label_" & xltab1.Range("B" & row).Value & "_" & add & "_" & LB_SZ & "_" & LB_COL & "." & LB_FF)
                    File.WriteAllBytes(pdfFilePath, pdfBytes)
                ElseIf CBox_Zipall_Label.Checked = True Then
                    pdfBytes = Await GetPDFResourceAsync(New Uri("https://eprel.ec.europa.eu/api/products/lightsources/" & add & "/labels"))
                    pdfFilePath = Path.Combine(TB_Label_Folder.Text, xltab1.Range("B" & row).Value & "_" & add & ".zip")
                    File.WriteAllBytes(pdfFilePath, pdfBytes)
                End If
            Next

        Catch ex As Exception
            xlBook.Close(SaveChanges:=False)
            xlApp.Quit()
            Exit Sub
        End Try

        xlBook.Close(SaveChanges:=False)
        xlApp.Quit()

        MsgBox("Download finished!", MsgBoxStyle.OkOnly)



    End Sub

    Public Async Function GetPDFResourceAsync(resourceUri As Uri) As Task(Of Byte())
        Dim request = WebRequest.CreateHttp(resourceUri)
        InitializeWebRequest(request)
        Using locResponse As HttpWebResponse = DirectCast(Await request.GetResponseAsync(), HttpWebResponse)
            If locResponse.StatusCode = HttpStatusCode.OK Then
                Return Await GetPDFResourceDirectAsync(locResponse.ResponseUri)
            Else
                Return Nothing
            End If
        End Using
    End Function

    Public Async Function GetPDFResourceDirectAsync(resourceUri As Uri) As Task(Of Byte())
        Dim request = WebRequest.CreateHttp(resourceUri)
        InitializeWebRequest(request)

        Dim buffersize As Integer = 132072
        Dim buffer As Byte() = New Byte(buffersize - 1) {}

        Dim dataResponse = DirectCast(Await request.GetResponseAsync(), HttpWebResponse)
        If dataResponse.StatusCode = HttpStatusCode.OK Then

            Using responseStream As Stream = dataResponse.GetResponseStream(),
            mStream As MemoryStream = New MemoryStream()
                Dim read As Integer = 0
                Do
                    read = Await responseStream.ReadAsync(buffer, 0, buffer.Length)
                    Await mStream.WriteAsync(buffer, 0, read)
                Loop While read > 0
                Return mStream.ToArray()
            End Using
        End If
    End Function

    Private Sub InitializeWebRequest(request As HttpWebRequest)
        request.UserAgent = "Mozilla/5.0 (Windows NT 10; WOW64; Trident/7.0; rv:11.0) like Gecko"
        request.AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate
        request.Headers.Add(HttpRequestHeader.AcceptEncoding, "gzip, deflate;q=0.8")
        request.Headers.Add(HttpRequestHeader.CacheControl, "no-cache")
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

    Private Async Sub B_Fiches_Loader_Click(sender As Object, e As EventArgs) Handles B_Fiches_Loader.Click
        If TB_Fiches_Folder.Text = "Click to select folder" Then
            MsgBox("Please choose output folder!")
            Exit Sub
        End If


        Dim pdfBytes As Byte()
        Dim pdfFilePath As String
        Dim xlApp As New EX.Application

        Dim input As New OpenFileDialog
        input.Filter = "Excel Files (*.xlsx)|*.xlsx"
        input.ShowDialog()



        Dim xlBook = xlApp.Workbooks.Open(input.FileName)
        Dim xltab1 = xlBook.Worksheets("DOWNLOAD")
        Dim xlUp = EX.XlDirection.xlUp
        Dim lastentry As Integer
        Dim add As String

        lastentry = xltab1.Range("B" & xltab1.Rows.Count).End(xlUp).Row
        'lastentry = xltab1.Range("A" & xltab1.Rows.Count).End(xlUp).Row
        'lastentry = xltab1.Range("A1:A" & lastentry).Value



        Try
            For row As Integer = 2 To lastentry
                add = xltab1.Range("A" & row).Value
                If CBox_Zipall_Fiche.Checked = False Then
                    pdfBytes = Await GetPDFResourceAsync(New Uri("https://eprel.ec.europa.eu/api/products/lightsources/" & add & "/fiches?language=" & CB_LB_LANG.Text))
                    pdfFilePath = Path.Combine(TB_Fiches_Folder.Text, "fiche_" & xltab1.Range("B" & row).Value & "_" & add & "_" & CB_LB_LANG.Text & ".pdf")
                    File.WriteAllBytes(pdfFilePath, pdfBytes)
                ElseIf CBox_Zipall_Fiche.Checked = True Then
                    pdfBytes = Await GetPDFResourceAsync(New Uri("https://eprel.ec.europa.eu/api/products/lightsources/" & add & "/fiches"))
                    pdfFilePath = Path.Combine(TB_Fiches_Folder.Text, xltab1.Range("B" & row).Value & "_" & add & ".zip")
                    File.WriteAllBytes(pdfFilePath, pdfBytes)
                End If
            Next

        Catch ex As Exception
            xlBook.Close(SaveChanges:=False)
            xlApp.Quit()
            Exit Sub
        End Try

        xlBook.Close(SaveChanges:=False)
        xlApp.Quit()

        MsgBox("Download finished!", MsgBoxStyle.OkOnly)

    End Sub

    Private Sub CBox_Zipall_CheckedChanged(sender As Object, e As EventArgs) Handles CBox_Zipall_Label.CheckedChanged
        If CBox_Zipall_Label.Checked = True Then
            Panel1.Enabled = False
            Panel2.Enabled = False
            Panel3.Enabled = False
        Else
            Panel1.Enabled = True
            Panel2.Enabled = True
            Panel3.Enabled = True
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CBox_Zipall_Fiche.CheckedChanged
        If CBox_Zipall_Fiche.Checked = True Then
            CB_LB_LANG.Enabled = False
        Else
            CB_LB_LANG.Enabled = True

        End If
    End Sub

    Private Sub TB_Fiches_Folder_Click(sender As Object, e As EventArgs) Handles TB_Fiches_Folder.Click
        Dim folder As New FolderBrowserDialog()

        If folder.ShowDialog = DialogResult.Cancel Then
            MsgBox("Please select folder!")
        Else
            TB_Fiches_Folder.Text = folder.SelectedPath + "\"
        End If

    End Sub
End Class