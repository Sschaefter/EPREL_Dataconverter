<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Txt_TrademarkRef = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Txt_Request = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.CB_Trademark = New System.Windows.Forms.CheckBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.CB_RegistrantNature = New System.Windows.Forms.ComboBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Txt_ContactRef = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.CB_ReasonChange = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.CheckB_Log = New System.Windows.Forms.CheckBox()
        Me.CB_OperationType = New System.Windows.Forms.ComboBox()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.BT_Tools = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(73, 356)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Start"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(23, 46)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Operation Type"
        '
        'Txt_TrademarkRef
        '
        Me.Txt_TrademarkRef.Location = New System.Drawing.Point(23, 142)
        Me.Txt_TrademarkRef.Name = "Txt_TrademarkRef"
        Me.Txt_TrademarkRef.Size = New System.Drawing.Size(221, 20)
        Me.Txt_TrademarkRef.TabIndex = 4
        Me.Txt_TrademarkRef.Text = "REF001"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(23, 126)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(111, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Trademark Reference"
        '
        'Txt_Request
        '
        Me.Txt_Request.Location = New System.Drawing.Point(23, 23)
        Me.Txt_Request.Name = "Txt_Request"
        Me.Txt_Request.Size = New System.Drawing.Size(221, 20)
        Me.Txt_Request.TabIndex = 6
        Me.Txt_Request.Text = "Request"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(23, 7)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Request ID"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.CB_Trademark)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.CB_RegistrantNature)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Txt_ContactRef)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.CB_ReasonChange)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.CheckB_Log)
        Me.Panel1.Controls.Add(Me.CB_OperationType)
        Me.Panel1.Controls.Add(Me.Txt_Request)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Txt_TrademarkRef)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(45, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(281, 318)
        Me.Panel1.TabIndex = 7
        '
        'CB_Trademark
        '
        Me.CB_Trademark.AutoSize = True
        Me.CB_Trademark.Location = New System.Drawing.Point(167, 125)
        Me.CB_Trademark.Name = "CB_Trademark"
        Me.CB_Trademark.Size = New System.Drawing.Size(77, 17)
        Me.CB_Trademark.TabIndex = 16
        Me.CB_Trademark.Text = "Trademark"
        Me.CB_Trademark.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(23, 203)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(102, 13)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Nature of Registrant"
        '
        'CB_RegistrantNature
        '
        Me.CB_RegistrantNature.FormattingEnabled = True
        Me.CB_RegistrantNature.IntegralHeight = False
        Me.CB_RegistrantNature.Items.AddRange(New Object() {"AUTHORISED_REPRESENTATIVE", "IMPORTER", "MANUFACTURER"})
        Me.CB_RegistrantNature.Location = New System.Drawing.Point(23, 219)
        Me.CB_RegistrantNature.Name = "CB_RegistrantNature"
        Me.CB_RegistrantNature.Size = New System.Drawing.Size(221, 21)
        Me.CB_RegistrantNature.TabIndex = 14
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(169, 258)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 13
        Me.Button2.Text = "Contact"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Txt_ContactRef
        '
        Me.Txt_ContactRef.Location = New System.Drawing.Point(23, 180)
        Me.Txt_ContactRef.Name = "Txt_ContactRef"
        Me.Txt_ContactRef.Size = New System.Drawing.Size(221, 20)
        Me.Txt_ContactRef.TabIndex = 12
        Me.Txt_ContactRef.Text = "REF001"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(23, 164)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(97, 13)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Contact Reference"
        '
        'CB_ReasonChange
        '
        Me.CB_ReasonChange.Enabled = False
        Me.CB_ReasonChange.FormattingEnabled = True
        Me.CB_ReasonChange.Items.AddRange(New Object() {"CORRECT_TYPO", "CHANGE_IN_STANDARDS", "LABEL_SCALE_RANGE_CHANGE", "CHANGE_REQUESTED_BY_MSA", "ADDED_INFORMATION_NO_EFFECT_ON_DECLARATION", "REQUEST_CHANGE_BY_EXTERNAL_BODY"})
        Me.CB_ReasonChange.Location = New System.Drawing.Point(23, 102)
        Me.CB_ReasonChange.Name = "CB_ReasonChange"
        Me.CB_ReasonChange.Size = New System.Drawing.Size(221, 21)
        Me.CB_ReasonChange.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(23, 86)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(99, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "Reason for Change"
        '
        'CheckB_Log
        '
        Me.CheckB_Log.AutoSize = True
        Me.CheckB_Log.Location = New System.Drawing.Point(23, 258)
        Me.CheckB_Log.Name = "CheckB_Log"
        Me.CheckB_Log.Size = New System.Drawing.Size(44, 17)
        Me.CheckB_Log.TabIndex = 8
        Me.CheckB_Log.Text = "Log"
        Me.CheckB_Log.UseVisualStyleBackColor = True
        '
        'CB_OperationType
        '
        Me.CB_OperationType.FormattingEnabled = True
        Me.CB_OperationType.IntegralHeight = False
        Me.CB_OperationType.Items.AddRange(New Object() {"DECLARE_END_DATE_OF_PLACEMENT_ON_MARKET", "PREREGISTER_PRODUCT_MODEL", "REGISTER_PRODUCT_MODEL", "UPDATE_PRODUCT_MODEL"})
        Me.CB_OperationType.Location = New System.Drawing.Point(23, 62)
        Me.CB_OperationType.Name = "CB_OperationType"
        Me.CB_OperationType.Size = New System.Drawing.Size(221, 21)
        Me.CB_OperationType.TabIndex = 7
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(115, 409)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(144, 13)
        Me.LinkLabel1.TabIndex = 8
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "©Mario Planeck, 07.07.2021"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(154, 356)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 9
        Me.Button3.Text = "Validate Zip"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'BT_Tools
        '
        Me.BT_Tools.Location = New System.Drawing.Point(235, 356)
        Me.BT_Tools.Name = "BT_Tools"
        Me.BT_Tools.Size = New System.Drawing.Size(75, 23)
        Me.BT_Tools.TabIndex = 10
        Me.BT_Tools.Text = "Tools"
        Me.BT_Tools.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(374, 431)
        Me.Controls.Add(Me.BT_Tools)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Button1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "EPREL Dataconverter V1.0.1"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Txt_TrademarkRef As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Txt_Request As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents CB_OperationType As ComboBox

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()
        CB_OperationType.SelectedIndex = 1
        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.

    End Sub

    Friend WithEvents LinkLabel1 As LinkLabel
    Friend WithEvents CheckB_Log As CheckBox
    Friend WithEvents CB_ReasonChange As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Txt_ContactRef As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Label6 As Label
    Friend WithEvents CB_RegistrantNature As ComboBox
    Friend WithEvents CB_Trademark As CheckBox
    Friend WithEvents BT_Tools As Button
End Class
