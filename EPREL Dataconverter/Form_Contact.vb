Imports System.Globalization
Imports System.Text.RegularExpressions


Public Class Form_Contact
    Public mailcheck As Boolean = False
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CB_ContactDetails.CheckedChanged
        If CB_ContactDetails.Checked = True Then
            P_ContactDetails.Enabled = True
            Form1.Txt_ContactRef.Enabled = False
        Else
            P_ContactDetails.Enabled = False
            Form1.Txt_ContactRef.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If CB_ContactDetails.Checked = True And (TB_ContactName.Text = "" Or TB_FirstName.Text = "" Or TB_LastName.Text = "" Or TB_PhoneNumber.Text = "") Then
            MsgBox("Please fill mandatory data!")
        ElseIf CB_ContactDetails.Checked = True And mailcheck = False Then
            MsgBox("Email not valid!")
        Else
            Hide()
        End If
    End Sub


    Private Sub TB_Email_Leave(sender As Object, e As EventArgs) Handles TB_Email.Leave
        Dim email As System.Net.Mail.MailAddress
        Try
            email = New Net.Mail.MailAddress(TB_Email.Text)
            mailcheck = True
        Catch ex As Exception
            If TB_Email.Text <> "" Then
                MsgBox("Email not valid!")
                mailcheck = False
            Else
                mailcheck = True
                Exit Try
            End If
        End Try
    End Sub

End Class