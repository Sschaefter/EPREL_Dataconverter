<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Contact
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.CB_ContactDetails = New System.Windows.Forms.CheckBox()
        Me.P_ContactDetails = New System.Windows.Forms.Panel()
        Me.TB_Email = New System.Windows.Forms.TextBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.CBox_Country = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.TB_Postcode = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.TB_Province = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TB_Municipality = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TB_City = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.TB_Number = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TB_StreetName = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TB_URL = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TB_PhoneNumber = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TB_LastName = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TB_FirstName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TB_ContactName = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.P_ContactDetails.SuspendLayout()
        Me.SuspendLayout()
        '
        'CB_ContactDetails
        '
        Me.CB_ContactDetails.AutoSize = True
        Me.CB_ContactDetails.Location = New System.Drawing.Point(12, 12)
        Me.CB_ContactDetails.Name = "CB_ContactDetails"
        Me.CB_ContactDetails.Size = New System.Drawing.Size(118, 17)
        Me.CB_ContactDetails.TabIndex = 0
        Me.CB_ContactDetails.Text = "Use Contact details"
        Me.CB_ContactDetails.UseVisualStyleBackColor = True
        '
        'P_ContactDetails
        '
        Me.P_ContactDetails.Controls.Add(Me.TB_Email)
        Me.P_ContactDetails.Controls.Add(Me.Label14)
        Me.P_ContactDetails.Controls.Add(Me.CBox_Country)
        Me.P_ContactDetails.Controls.Add(Me.Label13)
        Me.P_ContactDetails.Controls.Add(Me.Label12)
        Me.P_ContactDetails.Controls.Add(Me.TB_Postcode)
        Me.P_ContactDetails.Controls.Add(Me.Label11)
        Me.P_ContactDetails.Controls.Add(Me.TB_Province)
        Me.P_ContactDetails.Controls.Add(Me.Label10)
        Me.P_ContactDetails.Controls.Add(Me.TB_Municipality)
        Me.P_ContactDetails.Controls.Add(Me.Label9)
        Me.P_ContactDetails.Controls.Add(Me.TB_City)
        Me.P_ContactDetails.Controls.Add(Me.Label8)
        Me.P_ContactDetails.Controls.Add(Me.TB_Number)
        Me.P_ContactDetails.Controls.Add(Me.Label7)
        Me.P_ContactDetails.Controls.Add(Me.TB_StreetName)
        Me.P_ContactDetails.Controls.Add(Me.Label6)
        Me.P_ContactDetails.Controls.Add(Me.TB_URL)
        Me.P_ContactDetails.Controls.Add(Me.Label5)
        Me.P_ContactDetails.Controls.Add(Me.Label4)
        Me.P_ContactDetails.Controls.Add(Me.TB_PhoneNumber)
        Me.P_ContactDetails.Controls.Add(Me.Label3)
        Me.P_ContactDetails.Controls.Add(Me.TB_LastName)
        Me.P_ContactDetails.Controls.Add(Me.Label2)
        Me.P_ContactDetails.Controls.Add(Me.TB_FirstName)
        Me.P_ContactDetails.Controls.Add(Me.Label1)
        Me.P_ContactDetails.Controls.Add(Me.TB_ContactName)
        Me.P_ContactDetails.Enabled = False
        Me.P_ContactDetails.Location = New System.Drawing.Point(12, 35)
        Me.P_ContactDetails.Name = "P_ContactDetails"
        Me.P_ContactDetails.Size = New System.Drawing.Size(238, 433)
        Me.P_ContactDetails.TabIndex = 1
        '
        'TB_Email
        '
        Me.TB_Email.Location = New System.Drawing.Point(6, 133)
        Me.TB_Email.Name = "TB_Email"
        Me.TB_Email.Size = New System.Drawing.Size(206, 20)
        Me.TB_Email.TabIndex = 28
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(171, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(64, 13)
        Me.Label14.TabIndex = 27
        Me.Label14.Text = "* Mandatory"
        '
        'CBox_Country
        '
        Me.CBox_Country.FormattingEnabled = True
        Me.CBox_Country.Items.AddRange(New Object() {"AT", "BE", "BG", "CY", "CZ", "DE", "DK", "EE", "EL", "ES", "FI", "FR", "HR", "HU", "IE", "IT", "LT", "LU", "LV", "MT", "NL", "PL", "PT", "RO", "SE", "SI", "SK", "UK", "LI", "NO", "IS", "XI"})
        Me.CBox_Country.Location = New System.Drawing.Point(6, 406)
        Me.CBox_Country.Name = "CBox_Country"
        Me.CBox_Country.Size = New System.Drawing.Size(100, 21)
        Me.CBox_Country.TabIndex = 26
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(3, 390)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(43, 13)
        Me.Label13.TabIndex = 25
        Me.Label13.Text = "Country"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(3, 351)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(56, 13)
        Me.Label12.TabIndex = 24
        Me.Label12.Text = "Post Code"
        '
        'TB_Postcode
        '
        Me.TB_Postcode.Location = New System.Drawing.Point(6, 367)
        Me.TB_Postcode.Name = "TB_Postcode"
        Me.TB_Postcode.Size = New System.Drawing.Size(207, 20)
        Me.TB_Postcode.TabIndex = 23
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 312)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(49, 13)
        Me.Label11.TabIndex = 22
        Me.Label11.Text = "Province"
        '
        'TB_Province
        '
        Me.TB_Province.Location = New System.Drawing.Point(6, 328)
        Me.TB_Province.Name = "TB_Province"
        Me.TB_Province.Size = New System.Drawing.Size(207, 20)
        Me.TB_Province.TabIndex = 21
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(3, 273)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(62, 13)
        Me.Label10.TabIndex = 20
        Me.Label10.Text = "Municipality"
        '
        'TB_Municipality
        '
        Me.TB_Municipality.Location = New System.Drawing.Point(6, 289)
        Me.TB_Municipality.Name = "TB_Municipality"
        Me.TB_Municipality.Size = New System.Drawing.Size(207, 20)
        Me.TB_Municipality.TabIndex = 19
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(3, 234)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(24, 13)
        Me.Label9.TabIndex = 18
        Me.Label9.Text = "City"
        '
        'TB_City
        '
        Me.TB_City.Location = New System.Drawing.Point(6, 250)
        Me.TB_City.Name = "TB_City"
        Me.TB_City.Size = New System.Drawing.Size(207, 20)
        Me.TB_City.TabIndex = 17
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(171, 195)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(44, 13)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "Number"
        '
        'TB_Number
        '
        Me.TB_Number.Location = New System.Drawing.Point(174, 211)
        Me.TB_Number.Name = "TB_Number"
        Me.TB_Number.Size = New System.Drawing.Size(41, 20)
        Me.TB_Number.TabIndex = 15
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 195)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(66, 13)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Street Name"
        '
        'TB_StreetName
        '
        Me.TB_StreetName.Location = New System.Drawing.Point(6, 211)
        Me.TB_StreetName.Name = "TB_StreetName"
        Me.TB_StreetName.Size = New System.Drawing.Size(161, 20)
        Me.TB_StreetName.TabIndex = 13
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 156)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(29, 13)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "URL"
        '
        'TB_URL
        '
        Me.TB_URL.Location = New System.Drawing.Point(6, 172)
        Me.TB_URL.Name = "TB_URL"
        Me.TB_URL.Size = New System.Drawing.Size(207, 20)
        Me.TB_URL.TabIndex = 11
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 117)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(77, 13)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Email Address*"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 78)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(82, 13)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Phone Number*"
        '
        'TB_PhoneNumber
        '
        Me.TB_PhoneNumber.Location = New System.Drawing.Point(6, 94)
        Me.TB_PhoneNumber.Name = "TB_PhoneNumber"
        Me.TB_PhoneNumber.Size = New System.Drawing.Size(207, 20)
        Me.TB_PhoneNumber.TabIndex = 6
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(109, 39)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(58, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Last Name"
        '
        'TB_LastName
        '
        Me.TB_LastName.Location = New System.Drawing.Point(112, 55)
        Me.TB_LastName.Name = "TB_LastName"
        Me.TB_LastName.Size = New System.Drawing.Size(100, 20)
        Me.TB_LastName.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "First Name"
        '
        'TB_FirstName
        '
        Me.TB_FirstName.Location = New System.Drawing.Point(6, 55)
        Me.TB_FirstName.Name = "TB_FirstName"
        Me.TB_FirstName.Size = New System.Drawing.Size(100, 20)
        Me.TB_FirstName.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Contact Name*"
        '
        'TB_ContactName
        '
        Me.TB_ContactName.Location = New System.Drawing.Point(6, 16)
        Me.TB_ContactName.Name = "TB_ContactName"
        Me.TB_ContactName.Size = New System.Drawing.Size(100, 20)
        Me.TB_ContactName.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(97, 474)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Close"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form_Contact
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(268, 505)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.P_ContactDetails)
        Me.Controls.Add(Me.CB_ContactDetails)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Form_Contact"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Contact Details"
        Me.P_ContactDetails.ResumeLayout(False)
        Me.P_ContactDetails.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CB_ContactDetails As CheckBox
    Friend WithEvents P_ContactDetails As Panel
    Friend WithEvents Label9 As Label
    Friend WithEvents TB_City As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents TB_Number As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents TB_StreetName As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents TB_URL As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents TB_PhoneNumber As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents TB_LastName As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents TB_FirstName As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TB_ContactName As TextBox
    Friend WithEvents CBox_Country As ComboBox
    Friend WithEvents Label13 As Label
    Friend WithEvents Label12 As Label
    Friend WithEvents TB_Postcode As TextBox
    Friend WithEvents Label11 As Label
    Friend WithEvents TB_Province As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents TB_Municipality As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Label14 As Label
    Friend WithEvents TB_Email As TextBox
End Class
