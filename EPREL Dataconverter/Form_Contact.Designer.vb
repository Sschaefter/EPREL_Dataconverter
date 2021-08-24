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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form_Contact))
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
        resources.ApplyResources(Me.CB_ContactDetails, "CB_ContactDetails")
        Me.CB_ContactDetails.Name = "CB_ContactDetails"
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
        resources.ApplyResources(Me.P_ContactDetails, "P_ContactDetails")
        Me.P_ContactDetails.Name = "P_ContactDetails"
        '
        'TB_Email
        '
        resources.ApplyResources(Me.TB_Email, "TB_Email")
        Me.TB_Email.Name = "TB_Email"
        '
        'Label14
        '
        resources.ApplyResources(Me.Label14, "Label14")
        Me.Label14.Name = "Label14"
        '
        'CBox_Country
        '
        Me.CBox_Country.FormattingEnabled = True
        Me.CBox_Country.Items.AddRange(New Object() {resources.GetString("CBox_Country.Items"), resources.GetString("CBox_Country.Items1"), resources.GetString("CBox_Country.Items2"), resources.GetString("CBox_Country.Items3"), resources.GetString("CBox_Country.Items4"), resources.GetString("CBox_Country.Items5"), resources.GetString("CBox_Country.Items6"), resources.GetString("CBox_Country.Items7"), resources.GetString("CBox_Country.Items8"), resources.GetString("CBox_Country.Items9"), resources.GetString("CBox_Country.Items10"), resources.GetString("CBox_Country.Items11"), resources.GetString("CBox_Country.Items12"), resources.GetString("CBox_Country.Items13"), resources.GetString("CBox_Country.Items14"), resources.GetString("CBox_Country.Items15"), resources.GetString("CBox_Country.Items16"), resources.GetString("CBox_Country.Items17"), resources.GetString("CBox_Country.Items18"), resources.GetString("CBox_Country.Items19"), resources.GetString("CBox_Country.Items20"), resources.GetString("CBox_Country.Items21"), resources.GetString("CBox_Country.Items22"), resources.GetString("CBox_Country.Items23"), resources.GetString("CBox_Country.Items24"), resources.GetString("CBox_Country.Items25"), resources.GetString("CBox_Country.Items26"), resources.GetString("CBox_Country.Items27"), resources.GetString("CBox_Country.Items28"), resources.GetString("CBox_Country.Items29"), resources.GetString("CBox_Country.Items30"), resources.GetString("CBox_Country.Items31")})
        resources.ApplyResources(Me.CBox_Country, "CBox_Country")
        Me.CBox_Country.Name = "CBox_Country"
        '
        'Label13
        '
        resources.ApplyResources(Me.Label13, "Label13")
        Me.Label13.Name = "Label13"
        '
        'Label12
        '
        resources.ApplyResources(Me.Label12, "Label12")
        Me.Label12.Name = "Label12"
        '
        'TB_Postcode
        '
        resources.ApplyResources(Me.TB_Postcode, "TB_Postcode")
        Me.TB_Postcode.Name = "TB_Postcode"
        '
        'Label11
        '
        resources.ApplyResources(Me.Label11, "Label11")
        Me.Label11.Name = "Label11"
        '
        'TB_Province
        '
        resources.ApplyResources(Me.TB_Province, "TB_Province")
        Me.TB_Province.Name = "TB_Province"
        '
        'Label10
        '
        resources.ApplyResources(Me.Label10, "Label10")
        Me.Label10.Name = "Label10"
        '
        'TB_Municipality
        '
        resources.ApplyResources(Me.TB_Municipality, "TB_Municipality")
        Me.TB_Municipality.Name = "TB_Municipality"
        '
        'Label9
        '
        resources.ApplyResources(Me.Label9, "Label9")
        Me.Label9.Name = "Label9"
        '
        'TB_City
        '
        resources.ApplyResources(Me.TB_City, "TB_City")
        Me.TB_City.Name = "TB_City"
        '
        'Label8
        '
        resources.ApplyResources(Me.Label8, "Label8")
        Me.Label8.Name = "Label8"
        '
        'TB_Number
        '
        resources.ApplyResources(Me.TB_Number, "TB_Number")
        Me.TB_Number.Name = "TB_Number"
        '
        'Label7
        '
        resources.ApplyResources(Me.Label7, "Label7")
        Me.Label7.Name = "Label7"
        '
        'TB_StreetName
        '
        resources.ApplyResources(Me.TB_StreetName, "TB_StreetName")
        Me.TB_StreetName.Name = "TB_StreetName"
        '
        'Label6
        '
        resources.ApplyResources(Me.Label6, "Label6")
        Me.Label6.Name = "Label6"
        '
        'TB_URL
        '
        resources.ApplyResources(Me.TB_URL, "TB_URL")
        Me.TB_URL.Name = "TB_URL"
        '
        'Label5
        '
        resources.ApplyResources(Me.Label5, "Label5")
        Me.Label5.Name = "Label5"
        '
        'Label4
        '
        resources.ApplyResources(Me.Label4, "Label4")
        Me.Label4.Name = "Label4"
        '
        'TB_PhoneNumber
        '
        resources.ApplyResources(Me.TB_PhoneNumber, "TB_PhoneNumber")
        Me.TB_PhoneNumber.Name = "TB_PhoneNumber"
        '
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
        '
        'TB_LastName
        '
        resources.ApplyResources(Me.TB_LastName, "TB_LastName")
        Me.TB_LastName.Name = "TB_LastName"
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'TB_FirstName
        '
        resources.ApplyResources(Me.TB_FirstName, "TB_FirstName")
        Me.TB_FirstName.Name = "TB_FirstName"
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'TB_ContactName
        '
        resources.ApplyResources(Me.TB_ContactName, "TB_ContactName")
        Me.TB_ContactName.Name = "TB_ContactName"
        '
        'Button1
        '
        resources.ApplyResources(Me.Button1, "Button1")
        Me.Button1.Name = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form_Contact
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ControlBox = False
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.P_ContactDetails)
        Me.Controls.Add(Me.CB_ContactDetails)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Form_Contact"
        Me.ShowIcon = False
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
