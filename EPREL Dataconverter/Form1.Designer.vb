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
        resources.ApplyResources(Me.Button1, "Button1")
        Me.Button1.Name = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'Txt_TrademarkRef
        '
        resources.ApplyResources(Me.Txt_TrademarkRef, "Txt_TrademarkRef")
        Me.Txt_TrademarkRef.Name = "Txt_TrademarkRef"
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'Txt_Request
        '
        resources.ApplyResources(Me.Txt_Request, "Txt_Request")
        Me.Txt_Request.Name = "Txt_Request"
        '
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
        '
        'Panel1
        '
        resources.ApplyResources(Me.Panel1, "Panel1")
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
        Me.Panel1.Name = "Panel1"
        '
        'Label6
        '
        resources.ApplyResources(Me.Label6, "Label6")
        Me.Label6.Name = "Label6"
        '
        'CB_RegistrantNature
        '
        resources.ApplyResources(Me.CB_RegistrantNature, "CB_RegistrantNature")
        Me.CB_RegistrantNature.FormattingEnabled = True
        Me.CB_RegistrantNature.Items.AddRange(New Object() {resources.GetString("CB_RegistrantNature.Items"), resources.GetString("CB_RegistrantNature.Items1"), resources.GetString("CB_RegistrantNature.Items2")})
        Me.CB_RegistrantNature.Name = "CB_RegistrantNature"
        '
        'Button2
        '
        resources.ApplyResources(Me.Button2, "Button2")
        Me.Button2.Name = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Txt_ContactRef
        '
        resources.ApplyResources(Me.Txt_ContactRef, "Txt_ContactRef")
        Me.Txt_ContactRef.Name = "Txt_ContactRef"
        '
        'Label5
        '
        resources.ApplyResources(Me.Label5, "Label5")
        Me.Label5.Name = "Label5"
        '
        'CB_ReasonChange
        '
        resources.ApplyResources(Me.CB_ReasonChange, "CB_ReasonChange")
        Me.CB_ReasonChange.FormattingEnabled = True
        Me.CB_ReasonChange.Items.AddRange(New Object() {resources.GetString("CB_ReasonChange.Items"), resources.GetString("CB_ReasonChange.Items1"), resources.GetString("CB_ReasonChange.Items2"), resources.GetString("CB_ReasonChange.Items3"), resources.GetString("CB_ReasonChange.Items4"), resources.GetString("CB_ReasonChange.Items5")})
        Me.CB_ReasonChange.Name = "CB_ReasonChange"
        '
        'Label4
        '
        resources.ApplyResources(Me.Label4, "Label4")
        Me.Label4.Name = "Label4"
        '
        'CheckB_Log
        '
        resources.ApplyResources(Me.CheckB_Log, "CheckB_Log")
        Me.CheckB_Log.Name = "CheckB_Log"
        Me.CheckB_Log.UseVisualStyleBackColor = True
        '
        'CB_OperationType
        '
        resources.ApplyResources(Me.CB_OperationType, "CB_OperationType")
        Me.CB_OperationType.FormattingEnabled = True
        Me.CB_OperationType.Items.AddRange(New Object() {resources.GetString("CB_OperationType.Items"), resources.GetString("CB_OperationType.Items1"), resources.GetString("CB_OperationType.Items2"), resources.GetString("CB_OperationType.Items3")})
        Me.CB_OperationType.Name = "CB_OperationType"
        '
        'LinkLabel1
        '
        resources.ApplyResources(Me.LinkLabel1, "LinkLabel1")
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.TabStop = True
        '
        'Button3
        '
        resources.ApplyResources(Me.Button3, "Button3")
        Me.Button3.Name = "Button3"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'BT_Tools
        '
        resources.ApplyResources(Me.BT_Tools, "BT_Tools")
        Me.BT_Tools.Name = "BT_Tools"
        Me.BT_Tools.UseVisualStyleBackColor = True
        '
        'Form1
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.BT_Tools)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
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
    Friend WithEvents BT_Tools As Button
End Class
