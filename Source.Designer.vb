<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Source
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.SourceCode = New System.Windows.Forms.RichTextBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ViewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SourceCodeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FlowChartToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FlowChartPictureBox = New System.Windows.Forms.PictureBox()
        Me.ListBoxFlowChart = New System.Windows.Forms.ListBox()
        Me.ListBoxVariables = New System.Windows.Forms.ListBox()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.FlowChartPictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SourceCode
        '
        Me.SourceCode.Location = New System.Drawing.Point(12, 27)
        Me.SourceCode.Name = "SourceCode"
        Me.SourceCode.Size = New System.Drawing.Size(1281, 73)
        Me.SourceCode.TabIndex = 0
        Me.SourceCode.Text = ""
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ViewToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1424, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ViewToolStripMenuItem
        '
        Me.ViewToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SourceCodeToolStripMenuItem, Me.FlowChartToolStripMenuItem})
        Me.ViewToolStripMenuItem.Name = "ViewToolStripMenuItem"
        Me.ViewToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.ViewToolStripMenuItem.Text = "View"
        '
        'SourceCodeToolStripMenuItem
        '
        Me.SourceCodeToolStripMenuItem.Name = "SourceCodeToolStripMenuItem"
        Me.SourceCodeToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.SourceCodeToolStripMenuItem.Text = "Source Code"
        '
        'FlowChartToolStripMenuItem
        '
        Me.FlowChartToolStripMenuItem.Name = "FlowChartToolStripMenuItem"
        Me.FlowChartToolStripMenuItem.Size = New System.Drawing.Size(141, 22)
        Me.FlowChartToolStripMenuItem.Text = "Flow Chart"
        '
        'FlowChartPictureBox
        '
        Me.FlowChartPictureBox.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.FlowChartPictureBox.Location = New System.Drawing.Point(0, 96)
        Me.FlowChartPictureBox.Name = "FlowChartPictureBox"
        Me.FlowChartPictureBox.Size = New System.Drawing.Size(1293, 631)
        Me.FlowChartPictureBox.TabIndex = 2
        Me.FlowChartPictureBox.TabStop = False
        '
        'ListBoxFlowChart
        '
        Me.ListBoxFlowChart.FormattingEnabled = True
        Me.ListBoxFlowChart.HorizontalScrollbar = True
        Me.ListBoxFlowChart.Location = New System.Drawing.Point(1289, 27)
        Me.ListBoxFlowChart.Name = "ListBoxFlowChart"
        Me.ListBoxFlowChart.Size = New System.Drawing.Size(135, 407)
        Me.ListBoxFlowChart.Sorted = True
        Me.ListBoxFlowChart.TabIndex = 3
        '
        'ListBoxVariables
        '
        Me.ListBoxVariables.FormattingEnabled = True
        Me.ListBoxVariables.HorizontalScrollbar = True
        Me.ListBoxVariables.Location = New System.Drawing.Point(1315, 554)
        Me.ListBoxVariables.Name = "ListBoxVariables"
        Me.ListBoxVariables.Size = New System.Drawing.Size(109, 173)
        Me.ListBoxVariables.Sorted = True
        Me.ListBoxVariables.TabIndex = 4
        '
        'Source
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1424, 726)
        Me.Controls.Add(Me.FlowChartPictureBox)
        Me.Controls.Add(Me.ListBoxVariables)
        Me.Controls.Add(Me.ListBoxFlowChart)
        Me.Controls.Add(Me.SourceCode)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "Source"
        Me.Text = "Source"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.FlowChartPictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents SourceCode As RichTextBox
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents ViewToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SourceCodeToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FlowChartToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FlowChartPictureBox As PictureBox
    Friend WithEvents ListBoxFlowChart As ListBox
    Friend WithEvents ListBoxVariables As ListBox
End Class
