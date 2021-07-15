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
        Me.B_Fiches_Loader = New System.Windows.Forms.Button()
        Me.TB_Fiches_Folder = New System.Windows.Forms.TextBox()
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
        Me.B_Label_Loader.UseWaitCursor = True
        '
        'TB_Label_Folder
        '
        Me.TB_Label_Folder.Location = New System.Drawing.Point(12, 14)
        Me.TB_Label_Folder.Name = "TB_Label_Folder"
        Me.TB_Label_Folder.Size = New System.Drawing.Size(305, 20)
        Me.TB_Label_Folder.TabIndex = 1
        Me.TB_Label_Folder.Text = "Click to select folder"
        Me.TB_Label_Folder.UseWaitCursor = True
        '
        'B_Fiches_Loader
        '
        Me.B_Fiches_Loader.Location = New System.Drawing.Point(323, 41)
        Me.B_Fiches_Loader.Name = "B_Fiches_Loader"
        Me.B_Fiches_Loader.Size = New System.Drawing.Size(115, 23)
        Me.B_Fiches_Loader.TabIndex = 2
        Me.B_Fiches_Loader.Text = "Download Fiches"
        Me.B_Fiches_Loader.UseVisualStyleBackColor = True
        Me.B_Fiches_Loader.UseWaitCursor = True
        '
        'TB_Fiches_Folder
        '
        Me.TB_Fiches_Folder.Location = New System.Drawing.Point(12, 43)
        Me.TB_Fiches_Folder.Name = "TB_Fiches_Folder"
        Me.TB_Fiches_Folder.Size = New System.Drawing.Size(305, 20)
        Me.TB_Fiches_Folder.TabIndex = 3
        Me.TB_Fiches_Folder.Text = "Click to select folder"
        Me.TB_Fiches_Folder.UseWaitCursor = True
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
        Me.Controls.Add(Me.TB_Fiches_Folder)
        Me.Controls.Add(Me.B_Fiches_Loader)
        Me.Controls.Add(Me.TB_Label_Folder)
        Me.Controls.Add(Me.B_Label_Loader)
        Me.Name = "Form3"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Tools"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents B_Label_Loader As Button
    Friend WithEvents TB_Label_Folder As TextBox
    Friend WithEvents B_Fiches_Loader As Button
    Friend WithEvents TB_Fiches_Folder As TextBox
    Friend WithEvents DirectorySearcher1 As DirectoryServices.DirectorySearcher
End Class
