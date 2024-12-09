<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class sBNF
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.sBNF_Grammar = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'sBNF_Grammar
        '
        Me.sBNF_Grammar.Dock = System.Windows.Forms.DockStyle.Fill
        Me.sBNF_Grammar.FormattingEnabled = True
        Me.sBNF_Grammar.Location = New System.Drawing.Point(0, 0)
        Me.sBNF_Grammar.Name = "sBNF_Grammar"
        Me.sBNF_Grammar.Size = New System.Drawing.Size(800, 450)
        Me.sBNF_Grammar.Sorted = True
        Me.sBNF_Grammar.TabIndex = 0
        '
        'sBNF
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.sBNF_Grammar)
        Me.Name = "sBNF"
        Me.Text = "sBNF"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents sBNF_Grammar As ListBox
End Class
