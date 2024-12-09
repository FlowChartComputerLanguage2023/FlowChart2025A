


Public Class Source
    'Source to Flow Chart Menu item
    Private Sub SourceCodeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FlowChartToolStripMenuItem.Click
        ListBoxFlowChart.Items.Clear()
        Log("?", 1063, Err3, "Source to flow chart ")
        SourceCode.Visible = True 'False
        FlowChartPictureBox.Visible = True
        ListBoxFlowChart.Visible = True

        'convert the source code on this window to a flowchart on this window
        Application.DoEvents()
        ConvertSource2FlowChart(Me) ', Me.ListBoxFlowChart, Me. DrawingArea tureBox)
        Application.DoEvents()

    End Sub



    Private Sub SourceCodeToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles SourceCodeToolStripMenuItem.Click

        DebugLog("?", 1029, Err7, "flowchart to source " & CT & sender.ToString & CT & e.ToString)
        SourceCode.Visible = True
        '        Me.SourceCode.Dock = DockStyle.Fill
        FlowChartPictureBox.Visible = True 'False
        ListBoxFlowChart.Visible = True 'False
        ConvertFlowChart2Source(Me) ', Me. DrawingArea tureBox)
        Application.DoEvents()

    End Sub

    Private Sub FlowChartPictureBox_Click(sender As Object, e As EventArgs) Handles FlowChartPictureBox.Click
        Dim A As Point
        Dim B As Point
        A.X = 10
        A.Y = 10
        B.X = 100
        B.Y = 100
        FlowChartPictureBox.CreateGraphics.DrawLine(Pens.Black, AutoScrollPosition, B)
    End Sub

    Private Sub Source_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Dim X1, Y1, X2, Y2, X3, Y3 As Integer
        X1 = 0
        Y1 = 0
        'Debug.Print("Resize window " & Me.Size.ToString)
        X3 = Me.Size.Width
        Y3 = Me.Size.Height
        X2 = CInt(X3 * 0.8)
        Y2 = CInt(Y3 * 0.8)
        MySetSize(Me.FlowChartPictureBox, X1, Y1, X2, Y2)
        MySetSize(Me.SourceCode, X2, Y1, X3 - X2 - 20, Y2)
        MySetSize(Me.ListBoxFlowChart, X1, Y2, X2, Y3 - Y2 - 20)
        MySetSize(Me.ListBoxVariables, X2, Y2, X3 - X2 - 20, Y3 - Y2 - 20)
        Application.DoEvents()
        'Debug.Print("(" & X1.ToString & "," & Y1.ToString & ")-(" & X2.ToString & "," & Y2.ToString & ")")
    End Sub

    Friend Sub MySetSize(LB As PictureBox, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
        LB.Top = Y1
        LB.Left = X1
        LB.Height = Y2 - 10
        LB.Width = X2 - 10
        Application.DoEvents()
        'Debug.Print("(" & X1.ToString & "," & Y1.ToString & ")-(" & X2.ToString & "," & Y2.ToString & ")")
    End Sub

    Friend Sub MySetSize(LB As ListBox, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
        LB.Top = Y1
        LB.Left = X1
        LB.Height = Y2 - 10
        LB.Width = X2 - 10
        Application.DoEvents()
        'Debug.Print("(" & X1.ToString & "," & Y1.ToString & ")-(" & X2.ToString & "," & Y2.ToString & ")")
    End Sub

    Friend Sub MySetSize(LB As RichTextBox, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
        LB.Top = Y1
        LB.Left = X1
        LB.Height = Y2 - 10
        LB.Width = X2 - 10
        Application.DoEvents()
        'Debug.Print("(" & X1.ToString & "," & Y1.ToString & ")-(" & X2.ToString & "," & Y2.ToString & ")")
    End Sub

End Class


