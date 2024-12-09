Public Class FlowChart2025
    Private Sub NewSourceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewSourceToolStripMenuItem.Click
        Dim FileData, FileName As String
        Dim newMDIChild3 As New Source
        '''''Dim I As Integer
        '''''For I = Options.CheckedListBoxFlowChartOptions.Items.Count To 78
        '''''Options.CheckedListBoxFlowChartOptions.Items.Add(I.ToString)
        '''''Next

        newMDIChild3.MdiParent = Me
        FileName = XOpenFile("read", "Source file ", "Language")
        If FileName <> "" Then
            FileData = IO.File.ReadAllText(FileName)
            'FileData = FileData.Replace(vbLf, Environment.NewLine)
            newMDIChild3.SourceCode.Text = FileData
            FileData = newMDIChild3.SourceCode.Text
        End If
        newMDIChild3.Text = "Source : " & FileName
        newMDIChild3.Show()
        Application.DoEvents()
    End Sub

    Private Sub FlowChart2025_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Init(Source)
    End Sub

    Private Sub AddBlankLineToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddBlankLineToolStripMenuItem.Click
        'todo need to save any changes.
        Application.Exit()
    End Sub
End Class
