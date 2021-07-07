<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3
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
        Me.B_Label_Loader = New System.Windows.Forms.Button()
        Me.TB_Label_Folder = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.DirectorySearcher1 = New System.DirectoryServices.DirectorySearcher()
        Me.SuspendLayout()
        '
        'B_Label_Loader
        '
        Me.B_Label_Loader.Location = New System.Drawing.Point(323, 12)
        Me.B_Label_Loader.Name = "B_Label_Loader"
        Me.B_Label_Loader.Size = New System.Drawing.Size(115, 23)
        Me.B_Label_Loader.TabIndex = 0
        Me.B_Label_Loader.Text = "Download Labels"
        Me.B_Label_Loader.UseVisualStyleBackColor = True
        '
        'TB_Label_Folder
        '
        Me.TB_Label_Folder.Location = New System.Drawing.Point(12, 14)
        Me.TB_Label_Folder.Name = "TB_Label_Folder"
        Me.TB_Label_Folder.Size = New System.Drawing.Size(305, 20)
        Me.TB_Label_Folder.TabIndex = 1
        Me.TB_Label_Folder.Text = "Click to select folder"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(323, 41)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(115, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Download Fiches"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(12, 43)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(305, 20)
        Me.TextBox1.TabIndex = 3
        '
        'DirectorySearcher1
        '
        Me.DirectorySearcher1.ClientTimeout = System.TimeSpan.Parse("-00:00:01")
        Me.DirectorySearcher1.ServerPageTimeLimit = System.TimeSpan.Parse("-00:00:01")
        Me.DirectorySearcher1.ServerTimeLimit = System.TimeSpan.Parse("-00:00:01")
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(452, 103)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TB_Label_Folder)
        Me.Controls.Add(Me.B_Label_Loader)
        Me.Name = "Form3"
        Me.Text = "Tools"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents B_Label_Loader As Button
    Friend WithEvents TB_Label_Folder As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents DirectorySearcher1 As DirectoryServices.DirectorySearcher
End Class
