Public Class Options

    'todo make sure that this input is in the valid format 
    Private Sub AddBlankLine(NewLine As String)
        Dim I As Integer
        Dim X As String
        I = Me.OptionsTabCTL.SelectedIndex
        X = Me.OptionsTabCTL.Controls(I).Name
        Debug.Print(X)
        Debug.Print(Me.OptionsTabCTL.Controls(I).ToString)
        Select Case X
            Case "TabKeyWords"
                MatchFormat(38, NewLine)
                If InStr(NewLine, " ") <> 0 Then Log("", 1060, Err1, "A white space is invalid in a keyword " & ToHex(NewLine))
                Me.ListBox_KeyWords.Items.Add(NewLine)
                Application.DoEvents()
            Case "TabDataTypes"
                MatchFormat(32, NewLine)
                Me.ListBoxDataTypes.Items.Add(NewLine)
            Case "TabSymbols"
                Me.ListBoxSymbols.Items.Add(NewLine)
            Case "TabColors"
                Me.ListBoxColors.Items.Add(NewLine)
            Case "TabsBNF"
                Me.ListBoxGrammar.Items.Add(NewLine)
            Case "TabsBNFroots"
                Me.ListBoxGrammarStarts.Items.Add(NewLine)
            Case "TabPathIO"
                Me.ListBoxInputOutput.Items.Add(NewLine)
            Case "TabPathStyle"
                Me.ListboxPathLineStyle.Items.Add(NewLine)
            Case "TabPathStart"
                Me.ListboxPathStart.Items.Add(NewLine)
            Case "TabPathEnd"
                Me.ListboxPathEnd.Items.Add(NewLine)
            Case "TabRotation"
                Me.ListBoxRotation.Items.Add(NewLine)
            Case "TabBlocks"
                Me.ListBoxBlocks.Items.Add(NewLine)
            Case "TabKeyWordGraphics"
                Me.ListBoxSymbolKeyWordGraphics.Items.Add(NewLine)
            Case "TabGrammarGraphics"
                Me.ListBoxGrammarGraphics.Items.Add(NewLine)
            Case "TabPage17"
                Me.ListBoxLanguageClass.Items.Add(NewLine)
            Case "TabPage18"
                Me.ListBoxPage18.Items.Add(NewLine)
            Case Else
                Log("", 1061, Err1, "Invalid to add a line to this list " & X & NewLine)
        End Select
        Application.DoEvents()
    End Sub

    Private Sub AddToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddToolStripMenuItem.Click
        AddBlankLine(InputBox("Enter a new Line"))
    End Sub



    Friend Sub ReplaceLine(ByRef LB As ListBox, SpaceAllowed As Boolean, DeleteThenAdd As Boolean)
        Dim newline As String
        Dim I As Integer
        I = LB.SelectedIndex
        If I < 0 Then Return
        If I > LB.Items.Count - 1 Then Return
        newline = InputBox("Enter new value", "Keyword ", LB.Items.Item(I).ToString)
        If SpaceAllowed = False And InStr(newline, " ") <> 0 Then
            Log("", 1062, Err1, "A white space is invalid in a keyword " & ToHex(newline))
        Else
            LB.Items.RemoveAt(I)
            LB.Items.Add(newline)
            LB.SelectedItem = newline
        End If
        Application.DoEvents()
    End Sub


    Private Sub ListBox_KeyWords_Click(sender As Object, e As EventArgs) Handles ListBox_KeyWords.Click
        ReplaceLine(Me.ListBox_KeyWords, False, True)
    End Sub

    Private Sub ListBoxSymbolData_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxSymbolData.DoubleClick
        ReplaceLine(Me.ListBoxSymbolData, True, True)
    End Sub

    Private Sub ListBoxBlocks_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxBlocks.DoubleClick
        ReplaceLine(Me.ListBoxBlocks, False, True)
    End Sub

    Private Sub ListBoxColors_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxColors.DoubleClick
        ReplaceLine(Me.ListBoxColors, True, True)
    End Sub

    Private Sub ListBoxDataTypes_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxDataTypes.DoubleClick
        ReplaceLine(Me.ListBoxDataTypes, True, True)
    End Sub

    Private Sub ListBoxGrammar_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxGrammar.DoubleClick
        ReplaceLine(Me.ListBoxGrammar, True, False)
    End Sub

    Private Sub ListBoxGrammarGraphics_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxGrammarGraphics.DoubleClick
        ReplaceLine(Me.ListBoxGrammarGraphics, True, True)
    End Sub

    Private Sub ListBoxGrammarStarts_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxGrammarStarts.DoubleClick
        ReplaceLine(Me.ListBoxGrammarStarts, True, True)
    End Sub

    Private Sub ListBoxInputOutput_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxInputOutput.DoubleClick
        ReplaceLine(Me.ListBoxInputOutput, True, True)
    End Sub

    Private Sub ListBoxLanguageClass_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxLanguageClass.DoubleClick
        ReplaceLine(Me.ListBoxLanguageClass, True, True)
    End Sub

    Private Sub ListBoxPage18_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxPage18.DoubleClick
        ReplaceLine(Me.ListBoxPage18, True, False)
    End Sub

    Private Sub ListboxPathEnd_DoubleClick(sender As Object, e As EventArgs) Handles ListboxPathEnd.DoubleClick
        ReplaceLine(Me.ListboxPathEnd, True, True)
    End Sub

    Private Sub ListboxPathLineStyle_DoubleClick(sender As Object, e As EventArgs) Handles ListboxPathLineStyle.DoubleClick
        ReplaceLine(Me.ListboxPathLineStyle, True, True)
    End Sub

    Private Sub ListboxPathStart_DoubleClick(sender As Object, e As EventArgs) Handles ListboxPathStart.DoubleClick
        ReplaceLine(Me.ListboxPathStart, True, True)
    End Sub

    Private Sub ListBoxRotation_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxRotation.DoubleClick
        ReplaceLine(Me.ListBoxRotation, True, True)
    End Sub

    Private Sub ListBoxSymbolKeyWordGraphics_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxSymbolKeyWordGraphics.DoubleClick
        ReplaceLine(Me.ListBoxSymbolKeyWordGraphics, True, True)
    End Sub

    Private Sub ListBoxSymbols_DoubleClick(sender As Object, e As EventArgs) Handles ListBoxSymbols.DoubleClick
        ReplaceLine(Me.ListBoxSymbols, True, True)
    End Sub

    Private Sub ListboxPathEnd_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListboxPathEnd.SelectedIndexChanged

    End Sub
End Class