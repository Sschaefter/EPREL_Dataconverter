<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form3
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form3))
        Me.B_Label_Loader = New System.Windows.Forms.Button()
        Me.TB_Label_Folder = New System.Windows.Forms.TextBox()
        Me.B_Fiches_Loader = New System.Windows.Forms.Button()
        Me.TB_Fiches_Folder = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.RB_BW = New System.Windows.Forms.RadioButton()
        Me.RB_COL = New System.Windows.Forms.RadioButton()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.RB_SVG = New System.Windows.Forms.RadioButton()
        Me.RB_PNG = New System.Windows.Forms.RadioButton()
        Me.RB_PDF = New System.Windows.Forms.RadioButton()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.RB_SMALL = New System.Windows.Forms.RadioButton()
        Me.RB_BIG = New System.Windows.Forms.RadioButton()
        Me.CBox_Zipall_Label = New System.Windows.Forms.CheckBox()
        Me.CB_LB_LANG = New System.Windows.Forms.ComboBox()
        Me.CBox_Zipall_Fiche = New System.Windows.Forms.CheckBox()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'B_Label_Loader
        '
        Me.B_Label_Loader.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.B_Label_Loader, "B_Label_Loader")
        Me.B_Label_Loader.Name = "B_Label_Loader"
        Me.B_Label_Loader.UseVisualStyleBackColor = True
        '
        'TB_Label_Folder
        '
        Me.TB_Label_Folder.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.TB_Label_Folder, "TB_Label_Folder")
        Me.TB_Label_Folder.Name = "TB_Label_Folder"
        '
        'B_Fiches_Loader
        '
        Me.B_Fiches_Loader.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.B_Fiches_Loader, "B_Fiches_Loader")
        Me.B_Fiches_Loader.Name = "B_Fiches_Loader"
        Me.B_Fiches_Loader.UseVisualStyleBackColor = True
        '
        'TB_Fiches_Folder
        '
        Me.TB_Fiches_Folder.Cursor = System.Windows.Forms.Cursors.Default
        resources.ApplyResources(Me.TB_Fiches_Folder, "TB_Fiches_Folder")
        Me.TB_Fiches_Folder.Name = "TB_Fiches_Folder"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.RB_BW)
        Me.Panel1.Controls.Add(Me.RB_COL)
        resources.ApplyResources(Me.Panel1, "Panel1")
        Me.Panel1.Name = "Panel1"
        '
        'RB_BW
        '
        resources.ApplyResources(Me.RB_BW, "RB_BW")
        Me.RB_BW.Name = "RB_BW"
        Me.RB_BW.UseVisualStyleBackColor = True
        '
        'RB_COL
        '
        resources.ApplyResources(Me.RB_COL, "RB_COL")
        Me.RB_COL.Checked = True
        Me.RB_COL.Name = "RB_COL"
        Me.RB_COL.TabStop = True
        Me.RB_COL.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.RB_SVG)
        Me.Panel2.Controls.Add(Me.RB_PNG)
        Me.Panel2.Controls.Add(Me.RB_PDF)
        resources.ApplyResources(Me.Panel2, "Panel2")
        Me.Panel2.Name = "Panel2"
        '
        'RB_SVG
        '
        resources.ApplyResources(Me.RB_SVG, "RB_SVG")
        Me.RB_SVG.Name = "RB_SVG"
        Me.RB_SVG.UseVisualStyleBackColor = True
        '
        'RB_PNG
        '
        resources.ApplyResources(Me.RB_PNG, "RB_PNG")
        Me.RB_PNG.Name = "RB_PNG"
        Me.RB_PNG.UseVisualStyleBackColor = True
        '
        'RB_PDF
        '
        resources.ApplyResources(Me.RB_PDF, "RB_PDF")
        Me.RB_PDF.Checked = True
        Me.RB_PDF.Name = "RB_PDF"
        Me.RB_PDF.TabStop = True
        Me.RB_PDF.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.RB_SMALL)
        Me.Panel3.Controls.Add(Me.RB_BIG)
        resources.ApplyResources(Me.Panel3, "Panel3")
        Me.Panel3.Name = "Panel3"
        '
        'RB_SMALL
        '
        resources.ApplyResources(Me.RB_SMALL, "RB_SMALL")
        Me.RB_SMALL.Name = "RB_SMALL"
        Me.RB_SMALL.UseVisualStyleBackColor = True
        '
        'RB_BIG
        '
        resources.ApplyResources(Me.RB_BIG, "RB_BIG")
        Me.RB_BIG.Checked = True
        Me.RB_BIG.Name = "RB_BIG"
        Me.RB_BIG.TabStop = True
        Me.RB_BIG.UseVisualStyleBackColor = True
        '
        'CBox_Zipall_Label
        '
        resources.ApplyResources(Me.CBox_Zipall_Label, "CBox_Zipall_Label")
        Me.CBox_Zipall_Label.Name = "CBox_Zipall_Label"
        Me.CBox_Zipall_Label.UseVisualStyleBackColor = True
        '
        'CB_LB_LANG
        '
        Me.CB_LB_LANG.FormattingEnabled = True
        Me.CB_LB_LANG.Items.AddRange(New Object() {resources.GetString("CB_LB_LANG.Items"), resources.GetString("CB_LB_LANG.Items1"), resources.GetString("CB_LB_LANG.Items2"), resources.GetString("CB_LB_LANG.Items3"), resources.GetString("CB_LB_LANG.Items4"), resources.GetString("CB_LB_LANG.Items5"), resources.GetString("CB_LB_LANG.Items6"), resources.GetString("CB_LB_LANG.Items7"), resources.GetString("CB_LB_LANG.Items8"), resources.GetString("CB_LB_LANG.Items9"), resources.GetString("CB_LB_LANG.Items10"), resources.GetString("CB_LB_LANG.Items11"), resources.GetString("CB_LB_LANG.Items12"), resources.GetString("CB_LB_LANG.Items13"), resources.GetString("CB_LB_LANG.Items14"), resources.GetString("CB_LB_LANG.Items15"), resources.GetString("CB_LB_LANG.Items16"), resources.GetString("CB_LB_LANG.Items17"), resources.GetString("CB_LB_LANG.Items18"), resources.GetString("CB_LB_LANG.Items19"), resources.GetString("CB_LB_LANG.Items20"), resources.GetString("CB_LB_LANG.Items21"), resources.GetString("CB_LB_LANG.Items22"), resources.GetString("CB_LB_LANG.Items23")})
        resources.ApplyResources(Me.CB_LB_LANG, "CB_LB_LANG")
        Me.CB_LB_LANG.Name = "CB_LB_LANG"
        '
        'CBox_Zipall_Fiche
        '
        resources.ApplyResources(Me.CBox_Zipall_Fiche, "CBox_Zipall_Fiche")
        Me.CBox_Zipall_Fiche.Name = "CBox_Zipall_Fiche"
        Me.CBox_Zipall_Fiche.UseVisualStyleBackColor = True
        '
        'Form3
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.CBox_Zipall_Fiche)
        Me.Controls.Add(Me.CB_LB_LANG)
        Me.Controls.Add(Me.CBox_Zipall_Label)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TB_Fiches_Folder)
        Me.Controls.Add(Me.B_Fiches_Loader)
        Me.Controls.Add(Me.TB_Label_Folder)
        Me.Controls.Add(Me.B_Label_Loader)
        Me.Name = "Form3"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents B_Label_Loader As Button
    Friend WithEvents TB_Label_Folder As TextBox
    Friend WithEvents B_Fiches_Loader As Button
    Friend WithEvents TB_Fiches_Folder As TextBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents RB_BW As RadioButton
    Friend WithEvents RB_COL As RadioButton
    Friend WithEvents Panel2 As Panel
    Friend WithEvents RB_SVG As RadioButton
    Friend WithEvents RB_PNG As RadioButton
    Friend WithEvents RB_PDF As RadioButton
    Friend WithEvents Panel3 As Panel
    Friend WithEvents RB_SMALL As RadioButton
    Friend WithEvents RB_BIG As RadioButton
    Friend WithEvents CBox_Zipall_Label As CheckBox
    Friend WithEvents CB_LB_LANG As ComboBox
    Friend WithEvents CBox_Zipall_Fiche As CheckBox
End Class
