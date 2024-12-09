


'(1118

'todo Stevens rule #1 in a Grammarrule, anytime there is a "|" then that Grammar
'todo       rule is an option with the point in the symbol being the list of available
'todo       options (with the path constant being the selected option), and the point
'todo       name being the Grammar Rule Name
'todo Stevens rule #2 Symbols are made up of points (where paths connect to)
'todo       and lines which the user can +-modify
'todo 



'there is 1 copy of library and options for all source and flowchart
'each source code routine should have its own flowchart 

'done need to make the Options Form have tabs for each of the different list
'??? There should only be 3 types of syntax
'       A keyword
'       an options from a Grammar defined list (All literals)
'       everything else should be a point in the symbol, and a path to other symbols (ie: variables, Addresses (including sub names, labels (for goto,branch,select,switch...)))

Imports System.IO

'Imports Microsoft.VisualBasic.CallType


Module CommonRoutines

    Structure Universestructure
        Dim X As Integer 'This is where the last symbol was drawn (mydraw())
        Dim Y As Integer
        Dim MaxRows As Integer ' this is the number of lines down before coming back up to the top moving over and starting down again
        Dim Spacing1 As Integer ' This is the spacing between each symbol 
        Dim LogNumber As Integer
    End Structure

    Public MyUniverse As UniverseStructure
    Friend MyDebugLevel As Integer = 11


    Public Const Err0 As Integer = 0 'This is only for not doing anything
    Public Const Err1 As Integer = 1 'errors
    Public Const Err2 As Integer = 2 'Warnings
    Public Const Err3 As Integer = 4 'Information
    Public Const Err4 As Integer = 8 'These are no longer important (But good for debug)
    Public Const Err5 As Integer = 16 'Debugging
    Public Const Err6 As Integer = 32 'possible current bug problem 
    Public Const Err7 As Integer = 64 'untested
    Public Const Err8 As Integer = 128

    Public Const DebugShow As Integer = 4
    Public Const FC_Char = "/" ' Flow Chart Start Character

    Structure FlowChartRecordStructure
        Dim Code As Integer
        Dim X1 As Integer
        Dim y1 As Integer
        Dim X2 As Integer
        Dim IO As String 'for library symbols encoded into x2
        Dim Rotation As String ' for flowchart
        Dim Y2 As Integer
        Dim DT As String ' for symbols encoded into y2
        Dim Options As String ' defines how this symbol is used (reserved)
        Dim Ref As String
    End Structure

    Dim CurrentRecords(24) As FlowChartRecordStructure
    Dim CurrentSymbol(24) As FlowChartRecordStructure

    Structure Ustruct
        Dim CurrentSymbol As Integer
        Dim White_Space As String
        Dim PARSE1 As String
        Dim FlowChartComputerLanguage As Size
    End Structure

    Friend U As Ustruct

    Structure SourceSyntaxStructure

        Dim GrammarRuleName As String
        Dim GrammarLine As String
        Dim Grammars() As String        'holds the parsed grammarline
        Dim GrammarWhat() As String     'hold what it is (grammarrulename & -+*? , literal, '|')
        Dim FlowChart2Source() As String 'holds the translation from a flowchart symbol to source code.
        Dim AtParse As Integer

        Dim Source_Line_Number As Integer
        Dim SourceCodeLine As String
        Dim Source() As String          'holds the parsed source code line
        Dim SourceWhat() As String      'holds what it is ...

        Dim Source2FlowChart() As String       'This should hold the 'syntax' for the translation from source code to flow chart code

    End Structure

    Structure ConvertsStructure
        Dim SourceCode As String
        Dim GrammarCode As String
        Dim FCCode As String
    End Structure




    'Friend SSS As SourceSyntaxStructure

    Friend CT As String = vbCrLf & vbTab
    Friend CR As String = vbCrLf
    Friend MyDebugBits As Integer = &HFFFF ' zero is off


    Friend Function Src2FC(ByRef SourceForm As Source, SSS As SourceSyntaxStructure, depth As Integer) As String
        Dim GrammarRuleName As String
        Dim SaveGrammarRule As String
        Dim X, Y, Grammarline As String
        '1 text if starting with a grammar rule
        ParseAll(SourceForm, SSS)
        If SSS.AtParse >= SSS.Source.Length Then Return MyTrace(SSS, depth) & FC_Char & "Ignore=Finished Testing "
        Grammarline = MakeGrammar(SSS)
        Log("?", 9999, Err8, "Restart Src2FC " & Space(depth * 4) & depth.ToString & CT & SSS.SourceCodeLine & CT & SSS.GrammarLine)
        If GrammarLine = "" Then Return ""
        If InStr(SSS.GrammarLine, "::=") <> 0 Then
            GrammarRuleName = Parse(SSS.GrammarLine, "::=", "")
            GrammarLine = DropParse(SSS.GrammarLine, "::=", "")
            Log("", 1000, Err6, MyGraphicTrace(depth, SSS) & "New Grammar Rule " & HL(GrammarRuleName))
        Else
            GrammarRuleName = "" 'This is just part of a grammar rule
        End If
        '2 are we testing a grammar rule | option
        If InStr(GrammarLine, " | ") <> 0 Then
            SaveGrammarRule = SSS.GrammarLine
            X = GrammarLine
            While Len(X) <> 0
                SSS.GrammarLine = Parse(X, " | ", "")
                Log("", 1001, Err6, MyGraphicTrace(depth, SSS) & "New option " & SSS.GrammarLine & " for rule " & GrammarRuleName)
                Y = Src2FC(SourceForm, SSS, depth + 1)
                If Y <> "" Then
                    Log("", 1002, Err6, MyGraphicTrace(depth, SSS) & SSS.Source(SSS.AtParse) & " success for " & SSS.GrammarLine & " for rule " & GrammarRuleName)
                    Grammarline = Parse(Y, ",", "")
                    GrammarLine = Parse(GrammarLine, "", U.PARSE1)
                    AddGraphic(SourceForm, SSS, MyTrace(SSS, depth) & FC_Char & "Ignore=(J)" & GrammarRuleName & "," & SaveGrammarRule, depth)
                    If Len(Grammarline) > 1 Then AddGraphic(SourceForm, SSS, MyTrace(SSS, depth) & FC_Char & "symbol=" & GrammarRuleName, depth)
                    Return MyTrace(SSS, depth) & FC_Char & "Ignore=(K)" & GrammarRuleName & "," & SaveGrammarRule &
                        vbCrLf & MyTrace(SSS, depth) & FC_Char & "symbol=" & Parse(Y, ",", "") &
                        vbCrLf & MyTrace(SSS, depth) & FC_Char & "symbol=" & Grammarline
                End If
                X = DropParse(X, " | ", "")
            End While
        Else
            '3 test if this is a literal
            '3A test if it is a literal that matches
            If IsThisALiteral(GrammarLine) <> "" Then
                'todo fix this hack
                While SSS.AtParse >= SSS.Source.Length
                    ReDim Preserve SSS.Source(SSS.Source.Length + 1)
                    ReDim Preserve SSS.SourceWhat(SSS.SourceWhat.Length + 1)
                End While
                'Debug.Print(MyShowSSS(SSS))
                'Debug.Print(SSS.Source(SSS.AtParse) & ", " & SSS.GrammarLine)

                If IsThisATerminalMatch(SourceForm, SSS.Source(SSS.AtParse), IsThisALiteral(GrammarLine), depth) <> "" Then
                    Log("", 1003, Err6, MyGraphicTrace(depth, SSS) & SSS.Source(SSS.AtParse) & " literal success for " & SSS.GrammarLine & " for rule " & HL(GrammarRuleName) & CT & MyShowSSS(SSS))
                    'Debug.Print(MyShowSSS(SSS))
                    SSS.AtParse += 1
                    SSS.GrammarLine = DropParse(GrammarLine, "", U.PARSE1)
                    ParseAll(SourceForm, SSS)
                    'Debug.Print(MyShowSSS(SSS))
                    Return MyTrace(SSS, depth) & FC_Char & "Ignore=" & GrammarRuleName & "," & Grammarline & ", " & Src2FC(SourceForm, SSS, depth + 1)
                Else
                    Log("", 1004, Err6, MyGraphicTrace(depth, SSS) & SSS.Source(SSS.AtParse) & " literal failed for " & SSS.GrammarLine & " for rule " & GrammarRuleName)
                    Return "" 'failed to match to keyword
                End If
            End If

            If IsThisAGrammarRule(Parse(GrammarLine, "", U.PARSE1)) <> "" Then
                Log("", 1005, Err6, "Grammar rule modifier : " & HL(IsThisAGrammarRuleModifier(Parse(Grammarline, " ", ""))) & CT & HL(Grammarline))
                Select Case IsThisAGrammarRuleModifier(Parse(GrammarLine, " ", ""))
                    Case "*"
                        Log("", 1006, Err6, "* rule modifier " & Grammarline)
                    Case "-"
                        Log("", 1007, Err6, "- rule modifier " & Grammarline)
                    Case "?"
                        Log("", 1008, Err6, "? rule modifier " & Grammarline)
                    Case "+"
                        Log("", 1009, Err6, "+ rule modifier " & Grammarline)
                    Case Else
                        Log("", 1010, Err6, "no rule modifier " & Grammarline)
                End Select

                SSS.GrammarLine = IsThisAGrammarRule(Parse(Grammarline, "", U.PARSE1))

                Debug.Print(depth.ToString & ", " & SSS.AtParse.ToString & "] " & SSS.Source.Length.ToString)
                Debug.Print(SSS.AtParse.ToString)
                Debug.Print(HL(SSS.Source(SSS.AtParse)))
                Debug.Print(GrammarRuleName)



                Log("", 1011, Err6, MyGraphicTrace(depth, SSS) &
                    HL(SSS.Source(SSS.AtParse)) & " trying grammar rule " & GrammarRuleName)
                Y = Src2FC(SourceForm, SSS, depth + 1)
                If Y <> "" Then
                    Log("", 1012, Err6, MyGraphicTrace(depth, SSS) & SSS.Source(SSS.AtParse) & " success for " & SSS.GrammarLine & " for rule " & GrammarRuleName & CT & Y)
                    SSS.AtParse += 1
                    Y = Y & Src2FC(SourceForm, SSS, depth + 1)
                    Return MyTrace(SSS, depth) & Y
                Else
                    Log("", 1013, Err6, " Why?")
                End If
                'todo  this is a hack for now
                '3B test if a predefined grammar rule (for various types of variable names (Paths)
                'sBNF_Variable  (makes a path between)
                'sBNF_Function  (Is a constant option name )
                'sBNF_operator  (Is an operator from a list of |)
            End If
        End If
        Return ""
    End Function

    Friend Function MakeGrammar(sss As SourceSyntaxStructure) As String
        Dim I As Integer
        MakeGrammar = ""
        'Debug.Print(MyShowSSS(sss))
        For I = 0 To sss.Grammars.Length - 1
            If Len(sss.Grammars(I)) > 0 Then MakeGrammar &= " " & sss.Grammars(I)
        Next I
        MakeGrammar = Trim(MakeGrammar.Replace(" : : = ", " ::= "))
        'Debug.Print(MakeGrammar)
    End Function

    Friend Sub SetOption(Number As Integer, Value As String)
        Options.ListBoxSymbolData.Items.Item(Number) = Bin2Str(Number, 2) & "," & Value
    End Sub
    Friend Function GetOption(Number As Integer) As String
        DebugLog("", 1200, Err8, "Get Options " & Number.ToString & ") " & Options.ListBoxSymbolData.Items.Item(Number).ToString)
        Return Options.ListBoxSymbolData.Items.Item(Number).ToString
    End Function


    Friend Sub SetFCOption(Number As Integer, Value As String)
        Options.CheckedListBoxFlowChartOptions.Items.Item(Number) = Bin2Str(Number, 2) & "," & Value
    End Sub
    Friend Function GetFCOption(Number As Integer) As String
        Return Options.CheckedListBoxFlowChartOptions.Items.Item(Number).ToString
    End Function




    Friend Function HL(S As String) As String
        Return "-->" & S & "<--"
    End Function
    Friend Function NoCR(S As String) As String
        Return (S.Replace(vbCr, " ")).Replace(vbLf, " ")
    End Function
    Friend Function Max(a As Integer, b As Integer) As Integer
        If a > b Then Return a
        Return b
    End Function




    Public Sub NewSourceFlowChartForm()
        Dim FileData, FileName As String
        Dim newMDIChild3 As New Source

        newMDIChild3.MdiParent = FlowChart2025
        FileName = XOpenFile("read", "Source file ", "Language")
        If FileName <> "" Then
            FileData = IO.File.ReadAllText(FileName)
            newMDIChild3.SourceCode.Text = FileData.Replace(vbLf, Environment.NewLine)
        End If
        newMDIChild3.Text = "Source : " & FileName
        newMDIChild3.Show()
        Application.DoEvents()
    End Sub


    Friend Function ConverLineNumber2FlowChartXY(ByRef SourceForm As Source, LineNumber As Integer) As FlowChartRecordStructure
        Dim RTN As FlowChartRecordStructure = Nothing
        RTN.Code = 1
        RTN.X1 = 10
        RTN.y1 = LineNumber * 20
        RTN.X2 = 20
        RTN.Y2 = LineNumber * 20 - 20
        RTN.Ref = "??????????"
        Return RTN
    End Function


    'convert the flowchart back to the source code
    Friend Sub ConvertFlowChart2Source(F As Source) ', P As PictureBox)
        Dim FC As String
        FC = F.ListBoxFlowChart.ToString
        Log("", 1014, Err1, "FlowChart to text not written yet" & CT & F.Text & CT & HL(FC.ToString))
        Select Case MsgBox("Not re-written yet " & CT & F.ListBoxFlowChart.ToString & CT & F.SourceCode.ToString, MsgBoxStyle.AbortRetryIgnore)
            Case MsgBoxResult.Abort
                Application.Exit()
            Case MsgBoxResult.Cancel
            Case MsgBoxResult.Ignore
            Case MsgBoxResult.No
            Case MsgBoxResult.Ok
            Case MsgBoxResult.Retry
            Case MsgBoxResult.Yes
        End Select
    End Sub


    'conver the source code to a flowchart on this source window (sourceform). 
    Friend Function ConvertSource2FlowChart(ByRef SourceForm As Source) As String 'F As Form, S As TextBox, FC As RichTextBox, P As PictureBox)
        'This should make the flowchart from the source code
        Dim SSS As SourceSyntaxStructure = Nothing 'Hols the sourcedoe (to be converted) and the grammar being used to convert it
        Dim X As String = ""
        Dim Depth As Integer = 0
        SSS.Source_Line_Number = 0
        SSS.AtParse = 1

        'debug checking that I am working with what i think i am working with
        'make sure that this is using vbcrlf (sometimes vblf only is used (and vbcr, but im ignoring that for now
        'SourceForm.SourceCode.ToString()
        SSS.SourceCodeLine = ""
        For i = 0 To SourceForm.SourceCode.Lines.Count - 1
            SSS.SourceCodeLine &= SourceForm.SourceCode.Lines(i) & vbCrLf
        Next i

        'richtext box returns with heading
        'If it is a rich text passed then remove that from the front of the file
        If InStr(SSS.SourceCodeLine, "System.Windows.Forms.RichTextBox, Text:") <> 0 Then
            Log("", 1015, Err2, MyGraphicTrace(Depth, SSS) & " Removing Rich Text heading ")
            SSS.SourceCodeLine = DropParse(X, "System.Windows.Forms.RichTextBox, Text:")
        End If
        'SSS.SourceCodeLine = SSS.SourceCodeLine.Replace(vbLf, Environment.NewLine)

        'save off a copy (cause I like to work with small variable names that are meaningless
        X = SSS.SourceCodeLine
        ConvertSource2FlowChart = "" 'This should become the flowchart code to draw this line of code.

        'going through each line of code
        While Len(X) > 0
            SSS.Source_Line_Number += 1 'counting the line numbers (used to place the flow chart symbol
            'each line of code should be one symbol on the flow chart (so there should be a one to one associaation
            Log("", 1016, Err8, "^^^^^^^^^ Starting new line of source code " & SSS.Source_Line_Number.ToString & vbTab & Parse(X, vbCrLf, ""))
            SSS.SourceCodeLine = Parse(X, vbCrLf, "")
            ReDim CurrentRecords(24) 'erase everthing
            ReDim CurrentSymbol(24)
            AddGraphic(SourceForm, SSS, MyTrace(SSS, Depth) & FC_Char & "Ignore=(L) New Source Code Line  --> " & SSS.SourceCodeLine, Depth)
            Application.DoEvents()
            ConvertSource2FlowChart &= GetSymbolFromSource(SourceForm, SSS, Depth + 1)
            SourceLog("", 1300, SourceForm, Err5, "convert source 2 FC " & HL(Parse(X, vbCrLf, "")) & CT & ConvertSource2FlowChart & CT)
            If Len(ConvertSource2FlowChart) = 0 Then
                ConvertSource2FlowChart = FC_Char & "Message= This line of code did not pass the sBNF test & '0X000D'" & SSS.SourceCodeLine
                ConvertSource2FlowChart &= FC_Char & "Use=" & (SSS.Source_Line_Number * 10).ToString & ",0,Dafault,BLANKSYMBOL"
            End If
            MyDraw(SourceForm, SSS, ToHex(ConvertSource2FlowChart))
            X = DropParse(X, vbCrLf)
        End While
        Log("", 1017, Err1, "test Source Code to FlowChart not completed yet ")
        Return ConvertSource2FlowChart
    End Function


    Friend Function AddGraphic(ByRef Sourceform As Source, SSS As SourceSyntaxStructure, S As String, depth As Integer) As String
        Dim X As String
        While Len(S) > 0
            X = MyTrace(SSS, depth) & Parse(S, vbCrLf, "")
            Sourceform.ListBoxFlowChart.Items.Add(X)
            S = DropParse(S, vbCrLf, "")
            Debug.Print(X & CT & S)
        End While
        Application.DoEvents()
        Return ""
    End Function

    Friend Function SyntaxMatchesSource_X(Syntax As String, CodeLine As String) As String
        Dim Temp As String
        Temp = Syntax
        Log("", 1018, 256, "Watch " & Temp & vbTab & vbTab & ToHex(CodeLine))
        'todo this needs to match to get the next one
        Temp = Parse(Temp, "::=", "")
        Temp = DropParse(Temp, "::=")
        If Temp = "" Then Return ""
        Temp = Parse(Temp, "", U.White_Space)
        If IsThisAGrammarRule(Temp) <> "" Then
            Temp = Parse(IsThisAGrammarRule(Temp), "", U.PARSE1)
            Temp = SyntaxMatchesSource_X(Temp, CodeLine)
            Log("", 1019, 256, "Matched " & Temp & CT & CodeLine & CT & Syntax)
            Return Temp
        Else
            Log("", 1020, Err8, "Does " & Temp & " Match " & HL(ToHex(CodeLine)))
            Return "" ' failed to match
        End If

    End Function

    Friend Function GetSymbolFromSource(ByRef SourceForm As Source, SSS As SourceSyntaxStructure, Depth As Integer) As String
        Dim I As Integer
        Dim Temp As String
        Dim SaveGrammarLine, SaveSourceLine As String
        ReDim SSS.Source(2)
        ReDim SSS.Grammars(2)
        ReDim SSS.SourceWhat(2)
        ReDim SSS.GrammarWhat(2)
        ReDim SSS.Source2FlowChart(2)
        ReDim SSS.FlowChart2Source(2)
        GetSymbolFromSource = MyTrace(SSS, Depth) & FC_Char & "ignore=(M)" & Bin2Str(SSS.Source_Line_Number, 5) & ":" & Bin2Str(SSS.AtParse, 3) & ") " & SSS.SourceCodeLine & vbCrLf
        'SSS.SourceLine = Sline
        For I = 0 To Options.ListBoxGrammarStarts.Items.Count - 1
            Temp = Options.ListBoxGrammarStarts.Items(I).ToString
            SaveSourceLine = SSS.SourceCodeLine
            SaveGrammarLine = SSS.GrammarLine
            'get this rule by name, or else return nothing.
            SSS.GrammarLine = IsThisAGrammarRule(Temp)
            ParseAll(SourceForm, SSS)
            SourceLog("", 1301, SourceForm, Err8, MyGraphicTrace(Depth, SSS) & " +++++++++++ Starting new grammar check --> " & Options.ListBoxGrammarStarts.Items(I).ToString)
            Temp = Src2FC(SourceForm, SSS, Depth + 1)
            'Temp = MatchSyntax2Source(SourceForm, SSS, Depth + 1)
            SourceLog("", 1302, SourceForm, Err8, MyGraphicTrace(Depth, SSS) & " ^^^^^^^^^^ Returning Grammar check--> " & HL(Temp))
            If Temp <> "" Then
                'todo need to make FlowChartCode
                Log("", 1021, Err7, "*********************   Match ====>   " &
                    CT & HL(Temp) &
                    CT & SSS.SourceCodeLine &
                    CT & SSS.GrammarLine)
                'Log(2101, Err1, "Need to see if the rest of the line matches ")
                Log("", 1022, Err8, "Need to see if the rest of the line matches " & MyShowSSS(SSS))
                GetSymbolFromSource &= Temp & vbCrLf
            Else
                SourceLog("", 1304, SourceForm, Err8, MyGraphicTrace(Depth, SSS) & " -------- does not match " & vbTab & HL(SSS.SourceCodeLine) & "   " & HL(SSS.GrammarLine))
            End If
        Next
        Log("", 1024, Err2, "This is the code to draw a flowchart " & HL(GetSymbolFromSource) & HL(SSS.SourceCodeLine))
    End Function

    Friend Function MakeSymbolFromSyntax(ByRef SourceForm As Source, NameOfGrammarRule As String) As String ' Returns a import/export for this syntax()
        Dim GrammarRuleName, Syntax As String
        Dim RTN As String = ""
        GrammarRuleName = NameOfGrammarRule
        Syntax = IsThisAGrammarRule(GrammarRuleName)
        If Syntax = "" Then Return "" ' not a valid Grammar rule name 
        Syntax = DropParse(Syntax, "::=")


        While Len(Syntax) > 1
            ' We now have all of the rules
            ' We only need 3 keywords,terminals,option list Grammar name (where the Grammar is only a list of literals (no Grammar names)
            Select Case WhatIsThis(SourceForm, Syntax)' This is a list of whats what
                'This defines the name of the symbol (Which has this syntax
                Case "KeyWords",    ' Defines the symbol name
                     "Unicode",      ' Defines the symbol name
                     "Unknown" '
                    Log("", 1025, Err1, "Error 1")

                     'If the Grammar rule has a | in it then these are from a list of options
                     'If it does not have | then these are the constants on a path
                Case "Literal", 'Can be a keyword, or point or a constant on a path, but never a path name, 
                     "Character",
                     "Special" 'Option list, 
                    Log("", 1026, Err1, "Error 2")

                Case "BNF"         ' Defines the symbol name (The GrammarRuleName only)
                    RTN &= ", " & MakeSymbolFromSyntax(SourceForm, Parse(Syntax, "", U.PARSE1))

                'This defines information about a path and points in the symbol
                Case "Variable" 'this defines the name of a path
                    Log("", 1027, Err1, "Error 3")
                Case "color" 'todo Ignore these??
                    Log("", 1028, Err1, "Error 4")
                'todo color and datatype defines all unnamed paths (...as integer) (or optional all path names in the future (int variable;)
                Case "DataType" 'This is for a named path. 
                    Log("", 1029, Err1, "Error 5")
                Case "WhiteSpace" 'ignored (Assumed a space between all symbol points)
                    Log("", 1030, Err1, "Error 6")
                Case Else
                    Log("", 1031, Err1, "Error 7")
            End Select
            Syntax = DropParse(Syntax, "", U.PARSE1)
        End While
        Return ""
    End Function


    '   "50  fc_char  & Author=name"
    '    fc_char & "Block = begging String ending String”
    '    fc_char & "Blocks= begging string ending string comma begging string ending string …'Each string is divided in half. "
    '    fc_char & "Blocks= begging string ending string comma begging string ending string …'Each string is divided in half. "
    '    fc_char & "Color=Color Name, Alpha, Red, Green, Blue, Style, StartCap, EndCap"
    '    fc_char & "Constant=name, X, Y, Value"
    '    fc_char & "ConstantBranchToNextLine = branchto (Internal marker default ⬂BranchTo⬃)"
    '    fc_char & "ConstantDelimiter=CharacterCharacterCharacter…’ Should be in form ‘0X000D’ …"
    '    fc_char & "ConstantDelimiters=CharacterCharacterCharacter…’ Should be in form ‘0X000D’ …"
    '    fc_char & "ConstantDistanceBetweenControls=number (Pixel distance between FCCL controls)"
    '    fc_char & "ConstantDistanceToMovePaths=number (min distance between paths to allow movement Default 101)"
    '    fc_char & "DataType=DataTypeName, Number Of Bytes, Color Name, Color Width, Description"
    '    fc_char & "debug=number, number (First number Is the level of debug messages (MsgBox & Log [1=Errors, 2=warnings, 3=notices], [log 4=Info, 5=display, 6=status], [special 7-9]) "
    '    fc_char & "Delete= (Never used for input or output)"
    '    fc_char & "drilldown=path/file.extension (This will open and read the file)"
    '    fc_char & "dump=Dump File 1, File 2, File 3"
    '    fc_char & "Error   ' Code, name, x1, y1, Name, {other things}"
    '    fc_char & "exit=”  (Will exit the program!  Forced to quit, and nothing is saved.)"
    '    fc_char & "export=  (This will export the two log files 1 And 2, reserved for future usage )"
    '   ”(20)   fc_char  & FCCL_case= {Yes/No, 0/1, True/False, Yes/Ne}  (This is if the computer language is case sensitive)"
    '   ”(21)   fc_char  & FCCL_comment=string (This Is the marker for syntax default ⬂comment⬃)"
    '   ”(22)   fc_char  & FCCL_Default_Root=Grammar rule name, Symbol name"
    '   ”(22)   fc_char  & FCCL_Default_Roots=Grammar rule name 1, Symbol name 1 ,Grammar rule name 2, Symbol name 2 …"
    '   ”(23)   fc_char  & FCCL_DialectName = (replaced with /Language)"
    '    fc_char & "FCCL_Dimension= strings  (This must have 'Variable’ and ‘DataType’ in the strings it defines the syntax code text replacement for all paths)"
    '   ”(25)   fc_char  & FCCL_Extensions=extentions, … (These are the extensions that the computer language uses.  The First one will be the default of any output)"
    '    fc_char & "FCCL_multiLine= character(s)  (This is used as a divider between statements on the same line (as allowed by the computer language) default is : )"
    '    fc_char & "FCCL_Root=grammar rule name, grammar rule name, … (This Is a list of grammar rule names that care Not called by another grammar rule.)"
    '    fc_char & "FCCL_Roots=Grammar rule name, Symbol name , rule, symbol, .. (This will make this symbol be displayed when the grammar rule us used for this symbol  - Note that only lines are allowed in these special symbols Under construction))"
    '    fc_char & "FCCL_varchars= string  (These are characters that are allowed in variable name (in addition to A-Z, a-z, 0-9)"
    '    fc_char & "Fcfinish=(No longer used)"
    '    fc_char & "FileName=Device:/Path/FileName.Extension"
    '    fc_char & "function=FunctionName  or /Functions = FunctionName1, FunctionName2, . . . "
    '    fc_char & "functions=list of  Function names (This Is a list of functions (to be treated as keywords, non changeable)"
    '    fc_char & "Grammar=GrammarName '::=' Simple BNF "
    '    fc_char & "ignore=Ignores everything on this line"
    '    fc_char & "import  = (under construction see drilldown)"
    '    fc_char & "keyword=keyword  or    fc_char  & Keywords=keyword, keyword2, ... "
    '    fc_char & "KeywordS="
    '   "(1)    fc_char  & Language=Language class, Language dialect"
    '    fc_char & "Line=x1, y1, x2, y2, Color"
    '    fc_char & "Lines=Color, x1, y1, x2, y2, x3, y3, x4, y4 ... "
    '    fc_char & "login= (under construction)"
    '    fc_char & "MarkerAlpha=MarkerString (This is used internally only default Alpha)"
    '    fc_char & "MarkerBranchToNextLine=Markerstring (This Is used internally only default BranchTo)"
    '    fc_char & "MarkerCameFromLine=MarkerString (This is used internally only default CameFrom)"
    '    fc_char & "MarkerComment=MarkerString (This Is used internally only default Comment)"
    '    fc_char & "MarkerNumber=MarkerString (This is used internally only default Numbers)"
    '    fc_char & "MarkerQuotes=MarkerString (This Is used internally only quote)"
    '    fc_char & "MarkerSpecialCharacters=MarkerString (This is used internally only default Special)"
    '    fc_char & "Message = text to be displayed"
    '    fc_char & "Message= text  (This will display a message box with what ever text Is on the line)"
    '    fc_char & "MicroCodeText=Order section name, Text [replacements] text ... "
    '    fc_char & "Name=Symbol Name, options"
    '    fc_char & "Notes={}"
    '    fc_char & "OpCode={}"
    '    fc_char & "Operator=operator  or     fc_char  & operators = operator1 , operator2 , operator3 ... "
    '    fc_char & "operators="
    '    fc_char & "Option=number, {on/off, Yes/No, 1/0, True,False}  or /Option = ComputerLanguage???? "
    '    fc_char & "option=Option Number, {On/Off, 1/0, True/false}, Description" &  ct  &
    '    fc_char & "option=Option Number, {On/Off, 1/0, True/false}, New Description" &  ct  &
    '    fc_char & "Path=Name, x1, y1, x2, y2, Data type"
    '    fc_char & "Point=X, Y, {Input-Output ... }, Data Type, Name"
    '    fc_char & "route= [start] [, end] (This will try to connect all paths    fc_char  & route=, one path    fc_char  & route=10, Or a range of paths    fc_char  & route=1,1000"
    '    fc_char & "set={Points, Text, Delimiters, Language, Options, Option, Scale, Grids, Dump}" &  ct  &
    '    fc_char & "set={Points,Text,Delimiters,Language, Options, Option, Scale,Grids,Dump}" &  ct  &
    '    fc_char & "set=Delimiters, Start of marker, end of marker (no comma between)" &  ct  &
    '    fc_char & "set=Delimiters,Start of marker, end of marker" &  ct  &
    '    fc_char & "Set=Dump, DumpFile1.txt, DumpFile2.txt, DumpFile3.txt" &  ct  &
    '    fc_char & "Set=Dump,DumpFile1.txt, DumpFile2.txt, DumpFile3.txt" &  ct  &
    '    fc_char & "Set=Grids, Lines(1), PathsPoints(250), Symbols pathsPoints, 10000) " &  ct  &
    '    fc_char & "Set=Grids,Lines(1-10),PathsPoints(Lines,250),Symbols pathsPoints,10000) " &  ct  &
    '    fc_char & "set=language, class,  dialect"
    '    fc_char & "set=language,class, dialect"
    '    fc_char & "Set=Option, 1-64,{0/1,on/off,true/false}" &  ct  &
    '    fc_char & "Set=points, Index, X, Y" &  ct  &
    '    fc_char & "Set=points,Index,X,Y" &  ct  &
    '    fc_char & "Set=Scale, 625-10000" &  ct  &
    '    fc_char & "Set=Scale,625-10000" &  ct  &
    '    fc_char & "Set=text, Index, X, Y" &  ct  &
    '    fc_char & "Set=text,Index,X,Y" &  ct  &
    '    fc_char & "Stroke={}"
    '    fc_char & "Syntax={keyword, special characters, variables"
    '    fc_char & "Syntax={keyword, special characters" & COMMA & RMS & "variable" & FCCL_WhiteSpace & COMMA & RMS & "quote" & COMMA & RMS & "number, Alphabetics, and so on}"
    '    fc_char & "translate=English, spoken language  (This will replace all text from the english (word or phrase) to a spoken language (word or phrase) Note that this is a word for word translation, so the output would be in the order of the english input)"
    '    fc_char & "Unknown=used internally"
    '    fc_char & "Use=Name, X, Y, rotation, future dynamic options"
    '    fc_char & "Version= "
    '    fc_char & "x1=not used"
    '    fc_char & "x2=Not used"
    '    fc_char & "y1=not used"
    '    fc_char & "y2=Not used"


    'todo need to make sure that there is not two cases of the same Flow Chart command in FCcmd (I should but them in alphabetic order also)
    Friend Function FCcmd(ByRef SourceForm As Source, S As String, LineNumber As Integer) As String ' This should set parameters
        Dim X, XX, y As String 'Y1
        Dim C, Cx1, Cy1, Cx2, Cy2 As String
        Dim x1, y1 As Integer
        FCcmd = ""
        Debug.Print(S)
        If S = "" Then Return ""
        If S = FC_Char Then Return ""
        If InStr(S, ")") + 1 < InStr(S, "/") Then
            S = Mid(S, InStr(S, "/"))
        End If
        If Left(S, Len(FC_Char)) = FC_Char Then  'This is a flowchart command
            X = Mid(S, 2, Len(S))
            y = Parse(X, "=", "")
            X = DropParse(X, "=")
            Debug.Print(y & vbTab & vbTab & X & vbTab & vbTab & S)
            Select Case LCase(Trim(y))
                'goes on symbol
                'todo need to check the format of everything input 
                Case "author" '1
                    MatchFormat(50, X)
                    SetOption(50, X)
                    Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol) = CStr(Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol)) & CT & FC_Char & y & "=" & X
                Case "block", "blocks" 'blocks are implied subroutines (between {} or sub/end sub ...) 2
                    While Len(X) > 0
                        XX = Parse(X, ",", "")
                        Options.ListBoxBlocks.Items.Add(XX)
                        X = Trim(DropParse(X, ","))
                    End While
                    'goes on options.listbox colors
                Case "c_l_case" '3
                    SetOption(3, X)
                Case "c_l_comment" '4
                    SetOption(4, X)
                Case "c_l_dialectname" '5
                    SetOption(5, X)
                Case "c_l_extention"
                    SetOption(6, X)
                Case "c_l_multiline"
                    SetOption(7, X)
                Case "c_l_process"
                    SetOption(8, X)
                Case "c_l_varch"
                    SetOption(9, X)
                Case "color"
                    MatchFormat(31, X)
                    Options.ListBoxColors.Items.Add(X)
                Case "constant"
                    Options.ListBoxSymbols.Items.Add(FC_Char & y & " = " & X)
                Case "constantbranchtonextline"
                    SetOption(11, X)
                Case "constantdelimiters"
                    SetOption(12, X)
                Case "constantquote"
                    SetOption(13, X)

                Case "ConstantDistanceBetweenControls"
                    SourceLog("", 1305, SourceForm, Err1, "Not written yet ")
                Case "ConstantDistanceToMovePaths"
                    SourceLog("", 1306, SourceForm, Err1, "Not written yet ")

                Case "DataType"
                    MatchFormat(32, X)
                    Options.ListBoxDataTypes.Items.Add(X)
                Case "debug"
                    SourceLog("", 1307, SourceForm, Err1, FC_Char & "debug Not written yet " & X)
                Case "delete" 'todo ignore this
                    SetOption(15, X)
                Case "drilldown" 'open(new input file)
                    Log("", 1032, Err1, "Drilldown not done again yet)")
                    SetOption(16, X)
                    'goes on FC
                Case "dump"
                    SourceLog("", 1308, SourceForm, Err1, FC_Char & "dump Not written yet " & X)
                Case "error"
                    SetOption(17, X)
                Case "exit"
                    Application.Exit()
                        'Options.ListOptions.Items.Item(18,x)
                Case "export"
                    SourceLog("", 1309, SourceForm, Err1, FC_Char & "export Not written yet " & X)
                Case "extension"
                    SetOption(19, X)
                Case "fccl_case"
                    SetOption(20, X)
                Case "fccl_comment" 'This is the character(s) used for the start of a comment at the end of a line
                    SetOption(21, X)
                Case "FCCL_Default_Root"
                    SourceLog("", 1310, SourceForm, Err1, "Not written yet ")
                Case "fccl_default_roots"
                    SetOption(22, X)
                Case "fccl_dialectname"
                    SetOption(23, X)
                Case "FCCL_Dimension"
                    SourceLog("", 1311, SourceForm, Err1, "Not written yet ")
                Case "fccl_errormessage"
                    SetOption(24, X)
                Case "fccl_extension", "FCCL_Extensions"
                    SetOption(25, CStr(GetOption(25)) & X & ",")
                    SourceLog("", 1312, SourceForm, Err1, "Not written yet ")
                Case "fccl_process"
                    SetOption(26, X)
                Case "FCCL_Roots"
                    SourceLog("", 1313, SourceForm, Err1, "Not written yet ")
                Case "fccl_root" 'list of the starting places for the top of the bnf
                    SetOption(27, CStr(GetOption(27)) & X & ",")
                    SetOption(27, TrimOff(CStr(GetOption(27)), ","))
                    SetOption(27, TrimOff(CStr(GetOption(27)), " ") & ",")
                Case "fccl_varchars" 'These are characters allowed in a variable name (implied 0-9, a-z, A-Z)
                    SetOption(28, CStr(GetOption(28)) & X)
                    'Case "finish"
                    'Application.Exit() 'this was to be the last line in the clipboard
                     'Options.ListOptions.Items.Item(30,x)
                Case "Fcfinish"
                    SourceLog("", 1314, SourceForm, Err1, "Not written yet ")
                Case "FlowChart"
                    Options.ListBoxGrammarGraphics.Items.Add(X)
                    X = Trim(DropParse(X, ",", ""))
                    X = Trim(DropParse(X, ",", ""))
                    MatchFormat(0, X) ' Format 0 is special line... format
                Case "filename"
                    Library.LIB_parameters.Text = CStr(Library.LIB_parameters.Text) & FC_Char & y & " = " & X & vbCrLf
                        'Options.ListOptions.Items.Item(29,x)
                    'goes on Symbol
                Case "formatcolor"
                    SetOption(31, X)
                Case "formatdatatype"
                    SetOption(32, X)
                Case "formatfilename"
                    SetOption(33, X)
                Case "formatlanguage"
                    SetOption(34, X)
                Case "formatline"
                    SetOption(35, X)
                Case "formatmicrocodetext"
                    SetOption(36, X)
                Case "formatnotes"
                    SetOption(37, X)
                Case "FormatString"
                    SetOption(38, X)
                Case "formatpath"
                    SetOption(39, X)
                Case "formatpoint"
                    SetOption(40, X)
                Case "formatset"
                    SetOption(41, X)
                Case "formatstroke"
                    SetOption(42, X)
                Case "formatsyntax"
                    SetOption(43, X)
                Case "formatuse"
                    SetOption(44, X)
                Case "function", "functions" 'not used
                    SetOption(45, X)
                    SourceLog("", 1315, SourceForm, Err1, "Not written yet ")
                Case "grammar" 'sBNF rules
                    'make sure this has a ::= in it.
                    If InStr(X, "::=") = 0 Then
                        LogError(1039, "Grammar rule is missing a ::= " & CT & X, "Invalid Grammar rule")
                    End If
                    sBNF.sBNF_Grammar.Items.Add(Trim(X))
                Case "ignore"
                    SetOption(48, X)
                    'goes on options.listbox Keywords
                    'goes on FC
                Case "import"
                    SourceLog("", 1316, SourceForm, Err1, FC_Char & "import Not written yet " & X)
                Case "keywords", "keyword"
                    'todo need to make up the keywords when reading the sBNF file (everthing that is not a single character (or range)
                    'Options.ListOptions.Items.Item(49) &= x & ","
                    While Len(X) > 0
                        XX = Parse(X, ",", "")
                        Options.ListBox_KeyWords.Items.Add(XX)
                        X = Trim(DropParse(X, ","))
                    End While
                Case "language" 'This is the language class comma and the dialect of that 
                    SetOption(1, X)
                    sBNF.Text = sBNF.Name & " : " & X
                    Application.DoEvents()
                Case "line"
                    Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol) = CStr(Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol)) & CT & FC_Char & y & "=" & X
                    SetOption(35, X)
                    'goes on Symbol
                Case "lines"
                    Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol) = CStr(Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol)) & CT & FC_Char & y & "=" & X
                    SetOption(52, X)
                    'goes on Symbol
                Case "login"
                    SourceLog("", 1317, SourceForm, Err1, "Not written yet ")
                Case "markerbranchtonextline"
                    SetOption(53, X)
                Case "markercamefromline"
                    SetOption(54, X)
                Case "MarkerAlpha"
                    SourceLog("", 1318, SourceForm, Err1, "Not written yet ")
                Case "MarkerBranchToNextLine"
                    SourceLog("", 1319, SourceForm, Err1, "Not written yet ")
                Case "MarkerCameFromLine"
                    SourceLog("", 1320, SourceForm, Err1, "Not written yet ")
                Case "MarkerComment"
                    SourceLog("", 1321, SourceForm, Err1, "Not written yet ")
                Case "MarkerNumber"
                    SourceLog("", 1322, SourceForm, Err1, "Not written yet ")
                Case "MarkerQuotes"
                    SourceLog("", 1323, SourceForm, Err1, "Not written yet ")
                Case "MarkerSpecialCharacters"
                    SourceLog("", 1324, SourceForm, Err1, "Not written yet ")
                Case "message" 'This is a message to the user from the language definition file (ignored )
                    Select Case MsgBox(ToHex(X), MsgBoxStyle.Information, "Imbedded Message")
                        Case MsgBoxResult.Abort
                            Application.Exit()
                        Case MsgBoxResult.Cancel
                        Case MsgBoxResult.Ignore
                        Case MsgBoxResult.No
                        Case MsgBoxResult.Ok
                        Case MsgBoxResult.Retry
                        Case MsgBoxResult.Yes
                    End Select
                    SetOption(55, FC_Char & "ignore=" & NoCR(X))
                Case "microcodetext"
                    Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol) = CStr(Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol)) & CT & FC_Char & y & "=" & X
                    SetOption(56, X)
                    SourceLog("", 1326, SourceForm, Err1, "Not written yet ")
                Case "name"
                    U.CurrentSymbol += 1
                    Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol) = CStr(Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol)) & CT & FC_Char & y & "=" & X
                    SetOption(57, X)
                    'goes on Symbol
                Case "notes"
                    MyDrawText(SourceForm, X, MyUniverse.X, MyUniverse.Y, "Red")
                    Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol) = CStr(Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol)) & CT & FC_Char & y & "=" & X
                    SetOption(58, X)
                    'goes on Symbol
                Case "operator", "operators" 'List of operators (should be replaced with the name of the operators that are allowed in the BNF
                    SourceLog("", 1303, SourceForm, Err1, "Not written yet ")
                    SetOption(60, X)
                Case "option"
                    SetOption(61, X)
                    'not used any move replace with sBNF
                Case "order"
                    SetOption(62, X)
                Case "opcode"
                    Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol) = CStr(Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol)) & CT & FC_Char & y & "=" & X
                    SetOption(59, X)
                    'goes on Symbol
                Case "path"
                    Options.ListBoxSymbols.Items.Add(FC_Char & y & " = " & X)
                    SetOption(63, X)
                    'goes on FC
                Case "point"
                    Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol) = CStr(Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol)) & CT & FC_Char & y & "=" & X
                    SetOption(64, X)
                    'goes on Symbol
                Case "route" 'starts to rout all of the unconnected paths
                    SetOption(65, X)
                    RoutePaths()
                Case "set"
                    Select Case Parse(S, ",", "")
                        Case "delimiters"
                            SetOption(66, X)
                        Case "dump"
                            SetOption(67, X)
                        Case "grids"
                            SetOption(68, X)
                            SourceLog("", 1327, SourceForm, Err1, "Not written yet ")
                        Case "Option" '   fc_char  & set=option,... is not the same as    fc_char  & option=...
                            SetOption(69, X)
                            'goes on options.listbox ??? same as above
                        Case "options", "Option"
                            SetOption(70, X)
                            SourceLog("", 1328, SourceForm, Err1, "Not written yet ")
                            'goes on options list box 
                        Case "points"
                            SourceLog("", 1329, SourceForm, Err1, "Not written yet ")
                            SetOption(71, X)
                        Case "scale"
                            SetOption(72, X)
                            SourceLog("", 1330, SourceForm, Err1, "Not written yet ")
                        Case "text"
                            SetOption(73, X)
                            SourceLog("", 1331, SourceForm, Err1, "Not written yet ")
                        Case "set=language"
                            SourceLog("", 1332, SourceForm, Err1, "Not written yet ")
                        Case Else
                            Log("", 1033, Err1, "This /set code is not valid. ")
                    End Select
                Case "stroke"
                    SetOption(74, X)
                    SourceLog("", 1333, SourceForm, Err1, "Not written yet ")
                Case "Symbol"
                    Log("", 1034, Err2, X)
                    Debug.Print(FindKeyWordGraphic(X))
                    X = FindKeyWordGraphic(X)
                    X = DropParse(X, ",", "") 'name of of symbol
                    X = DropParse(X, ",", "") 'type of this symbol (line, point, ...)
                    Cx1 = "0"
                    Cy1 = "0"
                    C = "Red"
                    While Len(X) > 0
                        If IsThisAColor(X) <> "" Then
                            C = IsThisAColor(X)
                            X = DropParse(X, ",", "")
                            Cx1 = Parse(X, ",", "") : X = DropParse(X, ",", "")
                            Cy1 = Parse(X, ",", "") : X = DropParse(X, ",", "")
                            Cx2 = Cx1
                            Cy2 = Cy1
                        Else
                            Cx2 = Parse(X, ",", "") : X = DropParse(X, ",", "")
                            Cy2 = Parse(X, ",", "") : X = DropParse(X, ",", "")
                            MyDrawLine(SourceForm,
                                       Str2Bin(Cx1), Str2Bin(Cy1),
                                       Str2Bin(Cx2), Str2Bin(Cy2),
                                       C)
                            Cx1 = Cx2 'to draw from last point
                            Cy1 = Cy2
                        End If
                    End While
                Case "symbollines"
                    'each line starts with the name of a grammar rule, then a color and then the start of lines(x,y,x2,y2,X3,Y3...) until the X is a color and not a value  which starts a new line (could be the same color)
                    Options.ListBoxSymbolKeyWordGraphics.Items.Add(X)
                    DebugLog("", 1201, Err8, X)
                Case "syntax"
                    Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol) = CStr(Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol)) & CT & FC_Char & y & "=" & X
                    SetOption(75, X)
                Case "Text"
                    x1 = Str2Bin(Parse(X, ",", "")) : X = DropParse(X, ",", "")
                    y1 = Str2Bin(Parse(X, ",", "")) : X = DropParse(X, ",", "")
                    MyDrawText(SourceForm, X, x1, y1 * 10, "Red")
                    'goes on Symbol
                Case "translate"
                    SetOption(76, X)
                Case "Unknown"
                    SourceLog("", 1334, SourceForm, Err1, "Not written yet ")
                Case "use"
                    Options.ListBoxSymbols.Items.Add(FC_Char & y & " = " & X)
                    SetOption(77, X)
                Case "version" 'version of the symbol in the language
                    Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol) = CStr(Library.LIB_ListBoxProgramOptions.Items.Item(U.CurrentSymbol)) & CT & FC_Char & y & "=" & X
                    SetOption(78, X)
                    'this is a comment in the definition file and is ignored
                Case "x1"
                    SourceLog("", 1335, SourceForm, Err1, "Not written yet ")
                Case "x2"
                    SourceLog("", 1336, SourceForm, Err1, "Not written yet ")
                Case "y1"
                    SourceLog("", 1337, SourceForm, Err1, "Not written yet ")
                Case "y2"
                    SourceLog("", 1338, SourceForm, Err1, "Not written yet ")

                Case Else
                    Select Case MsgBox("Unable to complete this flow chart / command " & CT & y & CT & X, MsgBoxStyle.AbortRetryIgnore)
                        Case MsgBoxResult.Abort
                            Application.Exit()
                        Case MsgBoxResult.Cancel
                        Case MsgBoxResult.Ignore
                        Case MsgBoxResult.No
                        Case MsgBoxResult.Ok
                        Case MsgBoxResult.Retry
                        Case MsgBoxResult.Yes
                    End Select
            End Select
        Else 'this must be a source code
            DebugLog("", 1202, Err1, "Halt trying to do a flowchart command (most likely from a source code file)" & S)
        End If
        Return FCcmd
    End Function


    Friend Sub LoadLanguage(ByRef SourceForm As Source, FileText As String)
        Dim TextLine As String

        While Len(FileText) > 0
            TextLine = Parse(FileText, vbCrLf, "")
            FileText = DropParse(FileText, vbCrLf)
            FCcmd(SourceForm, TextLine, 0)
        End While
    End Sub


    Friend Sub LogError(RefNumber As Integer, Titled As String, Message As String)
        Select Case MsgBox(Titled, MsgBoxStyle.AbortRetryIgnore, RefNumber.ToString & ")  " & Message)
            Case MsgBoxResult.Abort
                Application.Exit()
            Case MsgBoxResult.Cancel
            Case MsgBoxResult.Ignore
            Case MsgBoxResult.No
            Case MsgBoxResult.Ok
            Case MsgBoxResult.Retry
            Case MsgBoxResult.Yes
        End Select
    End Sub


    Friend Function FailedLog(ByRef SourceForm As Source, ErrorNumber As Integer, FlowChartCode As String, Rule As String, MSG As String, Depth As Integer) As String
        If Depth < 1 Then Depth = 1
        SourceLog("", ErrorNumber, SourceForm, Err8, "[" & Depth.ToString & "] *****  Failed ***** " & MSG & vbTab & HL(ToHex(FlowChartCode)) & vbTab & HL(ToHex(Rule)))
        Return ""
    End Function




    Friend Sub SourceLog(junk As String, Errornumber As Integer, ByRef SourceForm As Source, Level As Integer, S As String)
        If Len(My.Application.Info.StackTrace) > 10000 Then
            Log("", Errornumber, Err1, "Stack overflow problem " & My.Application.Info.StackTrace)
        End If
        Log("", Errornumber, Err8, S)
    End Sub


    Friend Sub DebugLog(junk As String, ErrorNumber As Integer, Level As Integer, s As String)
        Log("", ErrorNumber, Level, s)
        If Len(My.Application.Info.StackTrace) > 10000 Then
            Log("", ErrorNumber, Err8, "Stack overflow problem " & My.Application.Info.StackTrace)
        End If
    End Sub


    Friend Function MyGraphicTrace(Depth As Integer, SSS As SourceSyntaxStructure) As String
        Return "[" & Depth.ToString & ", " & SSS.Source_Line_Number.ToString & ", " & SSS.AtParse.ToString & "] " & Space(Depth * 4)
    End Function

    Friend Function MyTrace(sss As SourceSyntaxStructure, Depth As Integer) As String
        Return Bin2Str(sss.Source_Line_Number, 5) & ":" & Bin2Str(sss.AtParse, 3) & ":" & Depth.ToString & ") " & FC_Char & "text=" & sss.AtParse.ToString & "," & Depth.ToString & "," & sss.GrammarLine & vbCrLf
    End Function


    Friend Sub Log(junk As String, ErrorNumber As Integer, Level As Integer, s As String, Optional A As String = "")

        If ErrorNumber = 1037 Then Return
        If ErrorNumber = 1215 Then Return
        If ErrorNumber = 1200 Then Return
        If ErrorNumber = 1000 Then Return
        If ErrorNumber = 1001 Then Return
        If ErrorNumber = 1004 Then Return
        If ErrorNumber = 1302 Then Return
        If ErrorNumber = 1010 Then Return
        If ErrorNumber = 99999 Then Return

        MyUniverse.LogNumber += 1
        IO.File.AppendAllText("FlowChart2025.txt", MyUniverse.LogNumber.ToString & "," & ErrorNumber.ToString & "[" & Level.ToString & "]) " & s & vbCrLf)
        Select Case MyDebugBits And Level
            Case Err0 ' Do Nothing
                Return
            Case Err1 'error
                Debug.Print(ErrorNumber.ToString & ")  Error> " & vbTab & vbTab & s)
                'todo make this optional MsgBox("Error : " & s, -1, "Error")
                'todo need to have a error list 
                IO.File.AppendAllText("FlowChart2025.txt", " * * * * * *  * * * * * * * Error  * * * * * *  * * * * * * * * " & CT & ErrorNumber.ToString & "[" & Level.ToString & "]) " & s & vbCrLf)
                Select Case MsgBox(s, MsgBoxStyle.AbortRetryIgnore, "Error #" & ErrorNumber.ToString & ":" & Level.ToString)
                    Case MsgBoxResult.Abort
                        Application.Exit()
                    Case MsgBoxResult.Cancel
                    Case MsgBoxResult.Ignore
                    Case MsgBoxResult.No
                    Case MsgBoxResult.Ok
                    Case MsgBoxResult.Retry
                    Case MsgBoxResult.Yes
                End Select
                Return
            Case Err2 'warning
                Debug.Print(ErrorNumber.ToString & ")  Warning> " & vbTab & vbTab & s)
                Return
            Case Err3 'information
                Debug.Print(ErrorNumber.ToString & ")  Information> " & vbTab & vbTab & s)
                Return
            Case Err4 'notice
                Debug.Print("^" & ErrorNumber.ToString & ") " & Level.ToString & " Other >" & vbTab & vbTab & ToHex(s))
                Return
            Case Err5 'debug
                Debug.Print("^" & ErrorNumber.ToString & ") " & Level.ToString & " Debug >" & vbTab & vbTab & s)
                Return
            Case Err6 'current debug
                Debug.Print("(" & MyUniverse.LogNumber.ToString & "," & ErrorNumber.ToString & ", " & Level.ToString & ") " & " Watch6 >" & vbTab & vbTab & s)
                Return
            Case Err7 'not looked at yet
                Debug.Print("(" & MyUniverse.LogNumber.ToString & "," & ErrorNumber.ToString & ", " & Level.ToString & ") " & " Watch7 >" & vbTab & vbTab & s)
            Case Err8
                Debug.Print("(" & MyUniverse.LogNumber.ToString & "," & ErrorNumber.ToString & ", " & Level.ToString & ") " & " Watch8 >" & vbTab & vbTab & s)
                Return
            Case Else 'unknown error level
                Debug.Print("@@@@@@@@@@@@@@@@@@ Unknown Error Level  @@@@@@@@@@@@@@@@@ " & ErrorNumber)
                Debug.Print(ErrorNumber.ToString & ")  Unknown Error " & Level.ToString & " >" & vbTab & vbTab & s)
                'MsgBox("Unknown Error Level : " & Level.ToString & "  " & s)
        End Select
    End Sub


    Friend Sub SetFormat(F As Form, L As Integer, T As Integer, W As Integer, H As Integer)
        F.Left = L
        F.Top = T
        F.Width = W
        F.Height = H
        F.Show()
    End Sub

    Friend Sub ShowAllForms()
        Dim x, y As Integer
        Dim X1, X2, X3 As Integer
        Dim Y1, Y2 As Integer


        x = FlowChart2025.Width - 25
        y = FlowChart2025.Height - 100
        X1 = 1
        X2 = CInt(x / 5 * 2)
        X3 = CInt(x / 5 * 4)
        Y1 = 1
        Y2 = CInt(y / 3 * 2)
        SetFormat(Library, X3, Y1, CInt(x / 5), Y2)
        SetFormat(sBNF, X1, Y2, x, CInt(y / 3))
        Options.WindowState = FormWindowState.Minimized
        Application.DoEvents()

    End Sub




    Friend Sub Init(ByRef SourceForm As Source)
        Dim FileName, FileData As String
        Dim I As Integer

        Debug.Print(CPUID())


        MyUniverse.X = 10
        MyUniverse.Y = 10
        MyUniverse.MaxRows = 10
        MyUniverse.Spacing1 = 100

        If Dir("FlowChart2025.txt") <> "" Then Kill("FlowChart2025.txt")

        For I = Options.ListBoxSymbolData.Items.Count To 78
            Options.ListBoxSymbolData.Items.Add(Format(I, Bin2Str(I, 2) & ","))
        Next
        SetOption(1, "Source Code Language")
        SetOption(3, "CL case")
        SetOption(4, "CL comment")
        SetOption(5, "CL DialectName")
        SetOption(6, "CL Extention")
        SetOption(7, "CL Multiline")
        SetOption(8, "CL Process")
        SetOption(9, "CL varch")
        SetOption(11, "CL constantBranchToNextLine")
        SetOption(12, "ConstantDelimiters")
        SetOption(13, "ConstantQuote")
        SetOption(15, "Delete")
        SetOption(16, "DrillDown")
        SetOption(17, "Error")
        SetOption(18, "(Was Exit)")
        SetOption(19, "Extension (see 6, 25)")
        SetOption(20, "FCCL Case ( See 3)")
        SetOption(21, "FCCL comment (See 4)")
        SetOption(22, "FCCL Default_roots")
        SetOption(23, "FCCL DialectName (see 5)")
        SetOption(24, "FCCL ErrorMessage")
        SetOption(25, "FCCL Extension (See 19, 6)")
        SetOption(26, "FCCL process (See 8)")
        SetOption(27, "FCCL root (See 22)")
        SetOption(28, "FCCL varchars (See 9)")
        SetOption(30, "finish (See 18)")
        SetOption(31, FC_Char & "color,$,#,#,#,#,$,$,$") 'name, Red, Green, Blue, Alpha, line style, start cap, end cap
        SetOption(32, FC_Char & "datatype,$,#,$,#,$") 'name, bytes, color, width, description (describtion should be the names of datatypes for structuress, firlds ararrys etc)
        SetOption(33, FC_Char & "FileName,$") ' device:/path(s)/filename.extention
        SetOption(34, FC_Char & "language,$,$") 'language, dialect
        '   fc_char  & lines=color,x1,y1,x2,y2,x3,y3... color,x1,y1,x2,y2,...  (new color name starts new line at x1,y1 to x2,y2 then x2,y2, to x3,y3 ...
        SetOption(52, FC_Char & "Lines,$,#,#,#,#") 'todo This should be recursive because any colorname in the X1 position, restarets xy's
        SetOption(35, FC_Char & "line,#,#,#,#,$") '   fc_char  & line=x,y,x,y,color 
        'SetOption(51,  fc_char & "Line,#,#,#,#,$")

        SetOption(36, "format micro code text")
        SetOption(37, FC_Char & "notes")
        SetOption(38, FC_Char & "String,$")
        SetOption(39, FC_Char & "path")
        SetOption(40, FC_Char & "point")
        SetOption(41, FC_Char & "set")
        SetOption(42, FC_Char & "stroke")
        SetOption(43, FC_Char & "syntax")
        SetOption(44, FC_Char & "use")
        SetOption(45, FC_Char & "function")
        SetOption(48, FC_Char & "Ignore")
        SetOption(49, FC_Char & "keyword/Keywords")
        SetOption(50, FC_Char & "author")
        SetOption(53, "Marker Branch To Next Line")
        SetOption(54, "Marker Came From Line")
        SetOption(56, "MicroCode")
        SetOption(57, "Name")
        SetOption(58, "Notes")
        SetOption(60, "Operator, Operators")
        SetOption(61, "Option")
        SetOption(62, "Order")
        SetOption(63, "Path")
        SetOption(64, "Point")
        SetOption(65, "Route")
        SetOption(66, "delimiters")
        SetOption(67, "dump")
        SetOption(68, "grids")
        SetOption(69, "option")
        SetOption(70, "options")
        SetOption(71, "points")
        SetOption(72, "scale")
        SetOption(73, "text")
        SetOption(74, "Stroke")
        SetOption(75, "Syntax")
        SetOption(76, "translate")
        SetOption(77, "Use")
        SetOption(78, "Version")


        'Set Defaults
        For I = Options.CheckedListBoxFlowChartOptions.Items.Count To 78
            Options.CheckedListBoxFlowChartOptions.Items.Add(Bin2Str(I, 2))
        Next

        'These Options are fixed in code (use to be in a structure)
        SetFCOption(0, "Check List")
        SetFCOption(1, "  Display Path Name")
        SetFCOption(2, "  Display Symbol Name Top")
        SetFCOption(3, "  Display ID Stroke")
        SetFCOption(4, "  Display File Name")
        SetFCOption(5, "  Display Notes")
        SetFCOption(6, "  Display OpCode")
        SetFCOption(7, "  Display Code")
        SetFCOption(8, "  Display Index Short Cut Pointer")
        SetFCOption(9, "  Display ErrorText")
        SetFCOption(10, "  Display InputOutPut")
        SetFCOption(11, "  Display Errors")
        SetFCOption(12, "  Display PathNames")
        SetFCOption(13, "  Display Constants")
        SetFCOption(14, "  Make Paths Orthogonal")
        SetFCOption(15, "  Move Symbols from On top Of Each other")
        SetFCOption(16, "  Output Line Numbers")
        SetFCOption(17, "  Display Data Value On Paths")
        SetFCOption(18, "  Disable ClipBoard Processor")
        SetFCOption(19, "  Auto Route")
        SetFCOption(20, "  Display Point data type")
        SetFCOption(21, "  Display Symbol Name Top")
        SetFCOption(23, "  Use hand To scroll")
        SetFCOption(24, "  Number Points")
        SetFCOption(25, "  Auto Correct Problems")
        SetFCOption(26, "  Stop Display debug ")
        SetFCOption(27, "  Dump Status")
        SetFCOption(28, "  Dump Messages")
        SetFCOption(29, "  Dump errors")
        SetFCOption(30, "  Dump Bugs")
        SetFCOption(31, "  Internal Checks")
        SetFCOption(33, "  Display Order Index ")
        SetFCOption(34, "  Display Text Ref #")
        SetFCOption(35, "  Animate")
        SetFCOption(36, "  Show Mouse Movement ")
        SetFCOption(37, "  Quick Load")
        SetFCOption(58, "  FIx <> In /Grammar")
        SetFCOption(60, "  DataType Variable")
        SetFCOption(61, "  Debug")
        SetFCOption(62, "  Ignore Grammar Checking ")

        U.White_Space = " " & vbCr & vbTab & vbLf & vbBack
        U.PARSE1 = U.White_Space & "' `!@#$%^&*()-+={[}]|\:;<,>?/"



        Source.MdiParent = FlowChart2025
        sBNF.MdiParent = FlowChart2025
        Library.MdiParent = FlowChart2025
        Source.MdiParent = FlowChart2025
        Options.MdiParent = FlowChart2025
        Source.WindowState = FormWindowState.Normal
        sBNF.WindowState = FormWindowState.Minimized
        Library.WindowState = FormWindowState.Minimized
        Source.WindowState = FormWindowState.Normal
        Options.WindowState = FormWindowState.Minimized

        FlowChart2025.WindowState = FormWindowState.Maximized
        FlowChart2025.Show()
        Source.WindowState = FormWindowState.Normal
        'Source.Show()
        sBNF.WindowState = FormWindowState.Normal
        sBNF.Show()
        Options.WindowState = FormWindowState.Normal
        Options.Show()
        Library.WindowState = FormWindowState.Normal
        Library.Show()
        Application.DoEvents()

        ShowAllForms()
        Application.DoEvents()

        'Dim newMDIChild2 As New Library
        'newMDIChild2.MdiParent = FlowChart2025A
        Library.MdiParent = FlowChart2025
        FileName = XOpenFile("Language", "Computer Language Definition File", "Language")
        If FileName <> "" Then
            FileData = IO.File.ReadAllText(FileName)
            LoadLanguage(SourceForm, FileData)
            'newMDIChild2.TextBox1.Text = FileData
            Library.Text = "Library : " & FileName
        End If
        Log("", 1035, Err3, "Definition file : " & Library.Text) 'information
        Library.Show()
        Application.DoEvents()

        'Dim newMDIChild4 As New FC
        Source.MdiParent = FlowChart2025
        'Source.Show()
        'Application.DoEvents()

        'Dim newMDIChild5 As New Options
        Options.MdiParent = FlowChart2025
        Options.Show()
        Application.DoEvents()

        MakeLibrary() 'Convert library into the sbnf file

        'FileName = XOpenFile("read", "Source Code File", "Source Code File")
        'If FileName <> "" Then
        ' FileData = IO.File.ReadAllText(FileName)
        ' MakeNewRoutine(FileName, FileData)
        ' End If

        '''''''MyTraceSystem()
    End Sub


    'Routine Returns the file name opened.(Does not change the input file name
    Friend Function XOpenFile(ReadOrWrite As String, MyTitle As String, Optional My_FileName As String = "*.*") As String
        My_FileName = Nothing
        Select Case LCase(ReadOrWrite)
            Case "read"
                Dim openFileDialog1 As New OpenFileDialog
                openFileDialog1.Title = MyTitle
                openFileDialog1.FileName = My_FileName
                openFileDialog1.InitialDirectory = Directory.GetCurrentDirectory
                openFileDialog1.Filter = "All Files (*.*)|*.*|Flow Chart Computer Language files (*.FlowChart)|*.FlowChart|Symbol files (*.Symbol)|*.Symbol|TextFile (*.txt)|*.txt"
                openFileDialog1.RestoreDirectory = True
                openFileDialog1.AddExtension = True
                openFileDialog1.DefaultExt = ".*"
                openFileDialog1.Multiselect = False
                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                    My_FileName = openFileDialog1.FileName
                Else
                    My_FileName = Nothing
                End If
                XOpenFile = My_FileName
                openFileDialog1.Dispose()

            Case "write"
                Dim SaveFileDialog1 As New SaveFileDialog()
                If My_FileName = Nothing Then My_FileName = "Start"
                SaveFileDialog1.Title = MyTitle
                SaveFileDialog1.FileName = My_FileName
                SaveFileDialog1.InitialDirectory = My.Application.Info.DirectoryPath
                SaveFileDialog1.Filter = "All Files (*.*)|*.*|Software Schematic files (*.FlowChart)|*.FlowChart|Symbol files (*.Symbol)|*.Symbol|TextFile (*.txt)|*.txt"
                SaveFileDialog1.RestoreDirectory = True
                SaveFileDialog1.AddExtension = True
                SaveFileDialog1.DefaultExt = ".FlowChart"
                SaveFileDialog1.CheckFileExists = False
                SaveFileDialog1.CheckPathExists = True
                If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
                    My_FileName = SaveFileDialog1.FileName
                Else
                    My_FileName = Nothing
                End If
                XOpenFile = My_FileName
                SaveFileDialog1.Dispose()

            Case "language"
                Dim openFileDialog1 As New OpenFileDialog()
                openFileDialog1.Title = MyTitle
                openFileDialog1.FileName = My_FileName
                openFileDialog1.InitialDirectory = My.Application.Info.DirectoryPath & "\Languages\"
                Log("", 1036, Err3, "Program Dir : " & openFileDialog1.InitialDirectory.ToString)
                'todo need to only show the standard file extentions for this language
                openFileDialog1.Filter = "All Files (*.*)|*.*|Flow Chart Computer Language files (*.FlowChart)|*.FlowChart|Symbol files (*.Symbol)|*.Symbol|TextFile (*.txt)|*.txt"
                openFileDialog1.RestoreDirectory = True
                openFileDialog1.AddExtension = True
                openFileDialog1.DefaultExt = ".*"
                openFileDialog1.Multiselect = False
                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                    My_FileName = openFileDialog1.FileName
                Else
                    My_FileName = Nothing
                End If
                XOpenFile = My_FileName
                openFileDialog1.Dispose()


            Case Else
                XOpenFile = Nothing
        End Select
    End Function

    Friend Function TrimOff(S As String, Remove As String, Optional RemoveAllThese As String = "") As String
        TrimOff = S
        If Len(RemoveAllThese) > 0 Then
            While Len(RemoveAllThese) > 0
                Remove = Left(RemoveAllThese, 1)
                RemoveAllThese = Mid(RemoveAllThese, 2, Len(RemoveAllThese))
                While LCase(Left(TrimOff, Len(Remove))) = LCase(Remove)
                    TrimOff = Mid(TrimOff, Len(Remove), Len(TrimOff) - 1)
                End While
                While LCase(Right(TrimOff, Len(Remove))) = LCase(Remove)
                    TrimOff = Mid(TrimOff, 1, Len(TrimOff) - Len(Remove))
                End While
            End While
        Else
            While LCase(Left(TrimOff, Len(Remove))) = LCase(Remove)
                TrimOff = Mid(TrimOff, Len(Remove) + 1, Len(TrimOff))
            End While
            While LCase(Right(TrimOff, Len(Remove))) = LCase(Remove)
                TrimOff = Mid(TrimOff, 1, Len(TrimOff) - Len(Remove))
            End While
        End If
    End Function


    Friend Sub ParseAll(ByRef SourceForm As Source, ByRef SSS As SourceSyntaxStructure)
        Dim I As Integer
        Dim X As String
        Dim SaveGrammarCodeLine As String = SSS.GrammarLine

        'get rid of any old stuff 
        ReDim SSS.Grammars(2)
        ReDim SSS.GrammarWhat(2)
        ReDim SSS.Source(2)
        ReDim SSS.SourceWhat(2)



        X = SSS.GrammarLine
        I = 0
        While Len(X) > 0
            I += 1
            ReDim Preserve SSS.Grammars(I + 1)
            ReDim Preserve SSS.GrammarWhat(I + 1)
            SSS.Grammars(I) = Parse(X, "", U.PARSE1)
            SSS.GrammarWhat(I) = WhatIsThis(SourceForm, SSS.Grammars(I))
            X = DropParse(X, "", U.PARSE1)
        End While

        X = SSS.SourceCodeLine
        I = 0
        While Len(X) > 0
            I += 1
            ReDim Preserve SSS.Source(I + 1)
            ReDim Preserve SSS.SourceWhat(I + 1)
            '?debug.print(X)
            SSS.Source(I) = Parse(X & " ", "", U.PARSE1)
            SSS.SourceWhat(I) = WhatIsThis(SourceForm, SSS.Source(I))
            X = Trim(DropParse(X & " ", "", U.PARSE1))
        End While
        FixParse(SSS)
        If MyDebugLevel > 10 Then Log("", 1037, Err6, "Parse all " & MyShowSSS(SSS))
    End Sub

    Friend Function MyShowSSS(sss As SourceSyntaxStructure) As String
        Dim I, J As Integer
        Dim X As String
        Dim S As Integer = 32
        MyShowSSS = ""
        MyShowSSS &= CT & "source code = " & sss.SourceCodeLine
        MyShowSSS &= CT & "grammar =" & sss.GrammarLine
        MyShowSSS &= CT & "line # = " & sss.Source_Line_Number.ToString & " Parse = " & sss.AtParse.ToString
        I = 0
        J = 0

        I = 0
        I = Max(I, sss.FlowChart2Source.Length)
        I = Max(I, sss.Grammars.Length)
        I = Max(I, sss.GrammarWhat.Length)
        I = Max(I, sss.Source.Length)
        I = Max(I, sss.Source2FlowChart.Length)
        I = Max(I, sss.SourceWhat.Length)
        'todo make sure that you dont go past the end 
        For J = 1 To I
            If J = sss.AtParse Then
                X = CT & J.ToString & "* A) "
            Else
                X = CT & J.ToString & ") A) "
            End If
            If J < sss.Source.Length Then
                X &= "Source = " & sss.Source(J)
            Else
                X &= " nil "
            End If

            X = Left(X & Space(S), S) & " b) "
            If J < sss.Grammars.Length Then
                X &= "Grammar = " & sss.Grammars(J).ToString
            Else
                X &= " nil "
            End If

            X = Left(X & Space(S * 2), S * 2) & " c) "
            If J < sss.GrammarWhat.Length Then
                X &= "Gr What = " & sss.GrammarWhat(J)
            Else
                X &= " nil "
            End If

            X = Left(X & Space(S * 3), S * 3) & " d) "
            If J < sss.SourceWhat.Length Then
                X &= "Src What = " & sss.SourceWhat(J)
            Else
                X &= " nil "
            End If

            X = Left(X & Space(S * 4), S * 4) & " e) "
            If J < sss.Source2FlowChart.Length Then
                X &= "Src2FC = " & sss.Source2FlowChart(J)
            Else
                X &= " nil "
            End If

            X = Left(X & Space(S * 5), S * 5) & " f) "
            If J < sss.FlowChart2Source.Length Then
                X &= "Fc2Src = " & sss.FlowChart2Source(J)
            Else
                X &= " nil "
            End If
            MyShowSSS &= X
        Next J
    End Function


    Friend Function UnParse(parsedArray() As String) As String
        Dim I As Integer
        UnParse = ""
        For I = 0 To parsedArray.Length - 1
            UnParse &= Left(U.White_Space, 1) & parsedArray(I)
        Next
        UnParse = UnParse.Replace(vbCr & " " & vbLf, vbCrLf)
    End Function


    'at is a string to find up to.
    'ats is a string to find the first character in that string ie: search linein for any character in ats
    'Returns each atom of string textline, (Divided by u.parse1)
    Friend Function MyParse(Parsed() As String, TextLine As String) As String
        Dim I As Integer = 0
        MyParse = ""
        While Len(TextLine) > 0
            ReDim Parsed(I + 1)
            Parsed(I) = Parse(TextLine, "", U.PARSE1)
            If IsThisAKeyword(Parsed(I)) <> "" Then
                MyParse &= "," & Parsed(I)
            End If
            I += 1
        End While
        MyParse = TrimOff(MyParse, ",")
    End Function
    Friend Function Parse(LineIn As String, At As String, Ats As String) As String
        Dim S As String
        Dim I, Lowest As Integer
        Dim X As String
        S = LineIn

        If S = "" Then Return ""
        If IsThisALiteral(LineIn) <> "" Then Return IsThisALiteral(LineIn)
        If Len(Ats) > 0 Then 'for list of possible delimiters
            X = Ats
            Lowest = Len(S) + 1
            While Len(X) > 0
                I = InStr(S, Left(X, 1))
                If I > 0 And I < Lowest Then Lowest = I
                X = Mid(X, 2, Len(X) - 1)
            End While
            If Lowest = 0 Then Lowest = Len(S)
            Parse = Mid(S, 1, Lowest - 1)
            If Len(Parse) <= 1 Then Return Left(LineIn, 1)
        Else 'for a string as a delimiter
            I = InStr(S, At)
            If I = 0 Then I = Len(S) + 1
            Parse = Mid(S, 1, I - 1)
        End If
        '    DebugLog( 1026  ,9,HL(S &  ct  &  "+" & Parse)
    End Function

    Friend Function DropParse(Linein As String, at As String, Optional Ats As String = "") As String
        Dim S As String
        Dim I, Lowest As Integer
        Dim X As String
        If Linein = "" Then Return ""
        S = Linein

        '+2 for the beginning and ending quotes (squotes)
        X = IsThisALiteral(Linein)
        If X <> "" Then
            Return Mid(Linein, Len(X) + 1, Len(Linein))
        End If
        If Len(Ats) > 0 Then 'for list of possible delimiters
            X = Ats
            Lowest = Len(S) + 1
            While Len(X) > 0
                I = InStr(S, Left(X, 1))
                If I > 0 And I <= Lowest Then Lowest = I
                X = Mid(X, 2, Len(X))
            End While
            If Lowest = 0 Then Lowest = Len(S) + 1
            DropParse = Mid(S, Lowest, Len(S))
            If Len(DropParse) = Len(Linein) Then Return Mid(S, 2, Len(S))
        Else 'for a string as a delimiter
            I = InStr(S, at)
            If I = 0 Then I = Len(S) + 1
            DropParse = Trim(Mid(S, I + Len(at), Len(S)))
        End If
    End Function

    Friend Sub MakeLibrary()
        Dim X, Y, Z As String
        Dim I As Integer
        For I = 0 To sBNF.sBNF_Grammar.Items.Count - 1
            'todo make up the keywords of the computer language
            Y = sBNF.sBNF_Grammar.Items.Item(I).ToString
            Z = Trim(Parse(Y, "::=", ""))
            Y = TrimOff(Y, " ;")
            Y = DropParse(Y, "::=")
            While Len(Y) > 0
                Z = IsThisALiteral(Y)
                If Z <> "" Then
                    'todo need to make sure this is not a character test (and assume that keywords can not be just one character!!!!!
                    If IsThisACharacterRange(TrimOff(Z, "'")) <> "" Then
                        'todo assume that this is not a keyword (with a dash in the middle of the word it canot be!!!!!
                    Else
                        If Len(Z) = 3 And Left(Z, 1) = "'" And Right(Z, 1) = "'" Then
                            'todo also make sure that it is spacial character and not a letter or number??????
                            'Debug.Print("Single character is required here " & Z)
                        Else
                            If IsThisAKeyword(TrimOff(Z, "'")) = "" Then ' check if it is already in the list
                                If MyDebugLevel > 11 Then Log("", 1038, Err8, "Added keyword " & Z)
                                Options.ListBox_KeyWords.Items.Add(TrimOff(Z, "'"))
                            Else
                                Log("", 1039, Err2, " Keyword " & Z & " used twice in the /grammar definition file")
                            End If
                        End If
                    End If
                    Y = Mid(Y, Len(Z) + 1, Len(Y))
                Else
                    Y = DropParse(Y, "", U.PARSE1)
                End If
            End While


            X = FindRootswAndTerminals(sBNF.sBNF_Grammar.Items(I).ToString)
            If Len(X) = 0 Then
                'todo need to now make a symbol object for this 
                If MyDebugLevel > 12 Then Log("", 1040, Err3, "Root or Terminal => " & sBNF.sBNF_Grammar.Items(I).ToString)
                Options.ListBoxGrammarStarts.Items.Add(FindRuleNameOnly(sBNF.sBNF_Grammar.Items(I).ToString))
            Else
                If MyDebugLevel > 12 Then DebugLog("", 1202, Err4, " This Rule is not a root/terminal => " & Parse(sBNF.sBNF_Grammar.Items(I).ToString, "::=", "") & " Is used by " & X)
            End If
        Next I

    End Sub


    Friend Sub RoutePaths()
        Log("", 1041, Err1, "Not Making Paths Right now ")
    End Sub



    'finds terminals and also if it is a root (top most that is not called by anyother)
    'todo this is not getting roots correctly (Gets some it shouldn't?)
    'todo This should also find out undefined grammar rules
    Friend Function FindRootswAndTerminals(Syntax As String) As String
        Dim SearchingFor, Found, Y, Z As String
        Dim SyntaxFormat As String = ""
        Dim E As String
        Dim J As Integer
        Found = ""
        Z = ""
        E = ""
        SearchingFor = " " & Trim(Parse(Syntax, "::=", "")) & " "

        For J = 0 To sBNF.sBNF_Grammar.Items.Count - 1
            Y = sBNF.sBNF_Grammar.Items.Item(J).ToString
            If Syntax <> Y Then
                Z = Trim(Parse(Y, "::=", ""))
                Y = TrimOff(Y, " ;")
                Y = DropParse(Y, "::=")
                If InStr(" " & Trim(Y) & " ", SearchingFor) <> 0 Then
                    Found &= Z & ","
                    If MyDebugLevel > 12 Then DebugLog("", 1203, Err8, "Found rule " & SearchingFor & CT & " in grammar rule " & sBNF.sBNF_Grammar.Items.Item(J).ToString)
                Else
                    If InStr(Y, Trim(SearchingFor)) <> 0 Then
                        If MyDebugLevel > 12 Then DebugLog("", 1204, Err8, "Found (maybe followed by +-*?)  " & SearchingFor & vbTab & " in " & vbTab & Z & " ::= " & Y)
                        E = Mid(Y, InStr(Y, Trim(SearchingFor)) + Len(SearchingFor) - 2, 1)
                        If InStr("+-*?", E) <> 0 Then
                            If InStr(" " & Trim(Y) & " ", " " & Trim(SearchingFor) & E & " ") <> 0 Then
                                If MyDebugLevel > 12 Then DebugLog("", 1205, Err4, "Found " & SearchingFor & " with an : " & E & " after with no spaces around it " & CT & sBNF.sBNF_Grammar.Items.Item(J).ToString)
                                Found &= Z & E & ", "
                            Else
                                'This is an error because most times you are using a Grammar rule with a prefix
                                Log("", 1042, Err1, " Grammar Name Rule error -->" & SearchingFor & "<-- with an " & E & " after with no spaces around it " & CT & sBNF.sBNF_Grammar.Items.Item(J).ToString)
                            End If
                        Else
                            Log("", 1043, Err1, "Found >" & SearchingFor & "< no spaces around it " & CT & sBNF.sBNF_Grammar.Items.Item(J).ToString)
                        End If
                        'MsgBox("Found it with no spaces around it " & SearchingFor & CT & Z & "::=" & Y)
                    Else
                        'Debug.Print("NOT FOUND " & SearchingFor & vbTab & " in " & vbTab & Z & " ::= " & Y)
                    End If
                End If
            End If
        Next


        If MyDebugLevel > 12 Then Log("", 1044, Err8, "Now check each for being a terminal or rulename " & Syntax)
        Y = DropParse(Syntax, "::=", "")
        SyntaxFormat = ""
        While Len(Y) <> 0
            Z = Parse(Y, "", U.PARSE1)

            'Debug.Print(HL(ToHex(Z)))
            If IsThisALiteral(Z) <> "" Then
                SyntaxFormat &= "L" 'is a terminal
            ElseIf IsThisAGrammarRule(Z) <> "" Then
                SyntaxFormat &= "R" 'is a RuleName
            ElseIf IsThisAWhiteSpace(Z) <> "" Then
                SyntaxFormat &= "_" 'is white space
            ElseIf Z = "|" Or Z = "*" Or Z = "-" Or Z = "?" Or Z = "+" Then
                SyntaxFormat &= Z '| is an OR, *-+? is parsed for after a grammar rule
            Else
                Log("", 1045, Err1, "This is not a valid grammar rule " & HL(ToHex(Z)) & ToHex(Syntax) & CT & HL(Syntax))
            End If
            Y = DropParse(Y, "", U.PARSE1)
        End While

        If Len(Found) > 0 Then
            If MyDebugLevel > 12 Then DebugLog("", 1206, Err8, "Found rule " & SearchingFor & "  <==== int grammar rule >  " & Found)
            Return Found
        End If
        'todo make sure that a message is given for this invalid grammar rule name not found 
        Return "" 'Trim(SearchingFor)
    End Function





    Friend Sub RemoveParseArray(ByRef ParsedArray() As String, At As Integer)
        Dim I As Integer
        For I = At To ParsedArray.Length - 2
            'Debug.Print(I.ToString & ") " & ParsedArray(I + 1))
            ParsedArray(I) = ParsedArray(I + 1)
        Next
        ParsedArray(ParsedArray.Length - 1) = ""
        'Log(1060, Err2, "last parse " & HL(ParsedArray(ParsedArray.Length - 1)) & ", " & ParsedArray.Length.ToString)
        If ParsedArray.Length > 2 Then
            While ParsedArray.Length > 2 And Len(ParsedArray(ParsedArray.Length - 1)) = 0
                'Log(1061, Err2, "remove last parse " & HL(ParsedArray(ParsedArray.Length - 1)) & ", " & ParsedArray.Length.ToString)
                ReDim Preserve ParsedArray(ParsedArray.Length - 2)
            End While
        End If
        'Log(1062, Err2, "last parse " & HL(ParsedArray(ParsedArray.Length - 1)) & ", " & ParsedArray.Length.ToString)
    End Sub



    Friend Function Success(ByRef SourceForm As Source, SSS As SourceSyntaxStructure, Depth As Integer) As String
        Dim Zero As Integer = SSS.Source.Count - 1
        Log("", 1047, Err8, "Success FOUND : " &
            FC_Char & "Message= characters match " & ToHex(SSS.GrammarLine) & UnParse(SSS.Source) & SSS.GrammarLine)
        AddGraphic(SourceForm, SSS, FC_Char & "Ignore=(A) Try  --> " & SSS.GrammarLine, Depth)
        DebugLog("", 1207, Err2, MyShowSSS(SSS))
        Application.DoEvents()
        Success = ""

        Success &= Bin2Str(SSS.Source_Line_Number, 5) & ":" & Bin2Str(SSS.AtParse, 3) & " B =" & FC_Char & "Message= characters match " & ToHex(SSS.GrammarLine) & UnParse(SSS.Source) & SSS.GrammarLine
    End Function



    'bug, it skips over the next one, when it deletes one
    'removes all of the blank parsing 
    'todo need to fix parse so that it does not have null strings
    'it also removes it if it is a white space in the /grammar 
    Friend Sub FixParse(ByRef SSS As SourceSyntaxStructure)
        Dim I As Integer
        Dim Flag As Boolean = False
        I = 0
        While SSS.Grammars.Length - 1 > 1 And SSS.Grammars(SSS.Grammars.Length - 1) = ""
            RemoveParseArray(SSS.Grammars, SSS.Grammars.Length - 1)
        End While
        While I < SSS.Grammars.Length - 1
            I += 1
            If I <= SSS.Grammars.Length - 2 Then
                If IsThisAGrammarRule(SSS.Grammars(I).ToString) <> "" Then
                    If SSS.Grammars(I + 1) = IsThisAGrammarRuleModifier(SSS.Grammars(I + 1)) Then
                        Log("", 1048, Err8, CT & "C)" & I.ToString & ">" & SSS.Grammars.Length.ToString & vbTab & SSS.Grammars(I).ToString & CT &
                            "D)" & I.ToString & ">" & SSS.Grammars(I).ToString & vbTab & SSS.Grammars(I).ToString & CT &
                            "E)" & I.ToString & "+1 >" & SSS.Grammars(I + 1).ToString & vbTab & SSS.Grammars(I + 1).ToString)
                        SSS.Grammars(I) = SSS.Grammars(I).ToString & SSS.Grammars(I + 1).ToString
                        SSS.Grammars(I + 1) = ""
                        RemoveParseArray(SSS.Grammars, I + 1)
                        RemoveParseArray(SSS.GrammarWhat, I + 1)
                        I -= 1
                        '?log(ErrorNumber, err8, SSS.Grammars.Length.ToString & vbTab & SSS.SourceForm.Length.ToString)
                    End If
                End If

                If Len(SSS.Grammars(I)) = 0 Or SSS.Grammars(I) = " " Then
                    Flag = True
                    RemoveParseArray(SSS.Grammars, I)
                    RemoveParseArray(SSS.GrammarWhat, I)
                End If
            Else
                If SSS.Grammars(SSS.Grammars.Length - 1) = "" Then
                    RemoveParseArray(SSS.Grammars, SSS.Grammars.Length - 1)
                    RemoveParseArray(SSS.GrammarWhat, SSS.Grammars.Length - 1)
                End If
            End If
        End While


        I = 0
        While I < SSS.Source.Length - 2
            I += 1
            If I < SSS.Source.Length - 1 Then
                If Len(SSS.Source(I)) = 0 Then
                    DebugLog("", 1208, Err8, I.ToString & "<" & SSS.Source.Length.ToString)
                    Flag = True
                    RemoveParseArray(SSS.Source, I)
                    RemoveParseArray(SSS.SourceWhat, I)
                    '?log(ErrorNumber, err8, SSS.Grammars.Length.ToString & vbTab & SSS.Source.Length.ToString)
                End If
            Else
                Log("", 1049, Err8, I.ToString & ">=" & SSS.Source.Length.ToString)
            End If
        End While
        If Flag = True Then
            FixParse(SSS)
        End If
    End Sub


    'assume that both are literals
    Friend Function IsThisATerminalMatch(ByRef SourceForm As Source, SourceAtom As String, GrammarRule As String, Depth As Integer) As String
        If GrammarRule = "" Then Return ""
        If SourceAtom = "" Then Return FC_Char & "ignore=" & SourceAtom & "," & GrammarRule
        If SourceAtom = GrammarRule Then Return SourceAtom
        If Len(SourceAtom) = Len(GrammarRule) - 2 Then
            If Left(GrammarRule, 1) = "'" And Right(GrammarRule, 1) = "'" Then
                If SourceAtom = Mid(GrammarRule, 2, Len(GrammarRule) - 2) Then
                    SourceLog("", 1334, SourceForm, Err5, Space(15) & "[?] ????????? Checking if terminals match   " & HL(ToHex(SourceAtom)) & " =? " & HL(ToHex(GrammarRule)))
                    Return SourceAtom
                End If
            End If
        End If
        If Len(SourceAtom) = 1 Then
            If IsThisACharacterRange(GrammarRule) <> "" Then
                If Hex2Bin(GrammarRule) = Hex2Bin(ToHex(SourceAtom)) Then
                    SourceLog("", 1334, SourceForm, Err5, Space(15) & "[?] ????????? Checking if terminals match   " & HL(ToHex(SourceAtom)) & " =? " & HL(ToHex(GrammarRule)))
                    Return SourceAtom
                End If
            End If
        End If
        Return "" 'Return Failed(SourceForm, 9200, SourceAtom, GrammarRule, "Literals do not match", Depth + 1)
    End Function




    Friend Function MatchSyntax2Source(ByRef SourceForm As Source,
                                       ByRef SSS As SourceSyntaxStructure,
                                       Depth As Integer) As String
        Dim SaveSource As String = SSS.SourceCodeLine
        Dim SaveGrammar As String = SSS.GrammarLine
        Dim LineNumberString As String = ""
        Dim X, Y, Z As String
        If SSS.AtParse <= SSS.Source.Length Then
            ReDim Preserve SSS.Source(SSS.AtParse + 1)
        End If
        If SSS.Source(SSS.AtParse) = "" Then Return Success(SourceForm, SSS, Depth)
        SourceLog("", 1335, SourceForm, Err8,
                  "[" & Depth.ToString & "," & SSS.AtParse.ToString & "] " &
                                    " Starting Syntax2Source " &
                                    HL(SSS.Source(SSS.AtParse)) & vbTab & vbTab & HL(SSS.Grammars(SSS.AtParse)))

        ''''''''''Depth += 1
        Z = ""
        MatchSyntax2Source = ""
        ParseAll(SourceForm, SSS)
        If SSS.Source.Length <= SSS.AtParse Then
            ReDim Preserve SSS.Source(SSS.AtParse + 1)
        Else
            Log("", 1050, Err6, "Parse at " & SSS.AtParse.ToString & ", # code atoms " & SSS.Source.Length.ToString & vbTab & HL(SSS.Source(SSS.AtParse)))
        End If
        Y = WhatIsThis(SourceForm, SSS.SourceCodeLine)
        X = IsThisAGrammarRule(Parse(SSS.GrammarLine, "", U.PARSE1))
        DebugLog("", 1209, Err8, "[" & Depth.ToString & "," & SSS.AtParse.ToString & "]" & " What is this " & HL(Y) &
                 CT & ", Grammar Rule Name? " & HL(SSS.GrammarRuleName) &
                 CT & ", ======== Starting New testing Of : " & HL(SSS.Source(SSS.AtParse)) & HL(SSS.Grammars(SSS.AtParse)))
        'first strip off the grammar rule name  and pass just the rules for this point in the source
        Log("", 1051, Err6, "Should I delete the grammar rule name first? " & SSS.GrammarRuleName)
        'SSS.GrammarRuleName = "" 'assume no grammar rule name
        If InStr(SSS.GrammarLine, "::=") <> 0 Then
            'SourceForm.ListBoxFlowChart.Items.addText(FC_Char & "Ignore=(B) Grammar rule --> " & Parse(SSS.GrammarLine, "::=", ""))
            SSS.GrammarRuleName = Trim(Parse(SSS.GrammarLine, "::=", ""))
            SourceLog("", 1336, SourceForm, Err8, "[" & Depth.ToString & ", " & SSS.AtParse.ToString & "] At Grammar Rule " & HL(SSS.GrammarRuleName))
            SSS.GrammarLine = DropParse(SSS.GrammarLine, "::=", "")
            X = SSS.GrammarLine
            While X <> ""
                Y = Parse(X, " | ", "")
                If Y <> "" Then
                    If SSS.AtParse < SSS.Source.Length Then
                        SourceLog("", 1337, SourceForm, Err8, "doing " & SSS.GrammarRuleName & "--" & HL(Y) & " with " & HL(SSS.Source(SSS.AtParse)) & " source " & HL(SSS.SourceCodeLine))
                    Else
                        Log("", 1052, Err1, "Program  error in logic ")
                    End If
                    If SSS.AtParse < SSS.Source.Length Then
                        Log("", 1053, Err8, SSS.GrammarLine & "<<<<>>>>>" & vbTab & Y & " with " & SSS.Source(SSS.AtParse) & CT & SaveSource & CT & SaveGrammar)
                    Else
                        Log("", 1054, Err1, "Program error in logic - blank line")
                    End If
                    SSS.GrammarLine = Y
                    Z = MatchSyntax2Source(SourceForm, SSS, Depth + 1)
                    If Z <> "" And Z <> vbCrLf & "\" Then
                        While Len(Z) <> 0
                            AddGraphic(SourceForm, SSS, FC_Char & "Ignore=(C) Found " & Parse(Z, vbCrLf, "") & "," & Parse(SaveGrammar, "::=", ""), Depth)
                            Z = DropParse(Z, vbCrLf, "")
                        End While
                        SourceLog("", 1338, SourceForm, Err8, Bin2Str(SSS.Source_Line_Number, 5) & ":" & Bin2Str(SSS.AtParse, 3) & "= [" & Depth.ToString & ", " & SSS.AtParse.ToString & "]" & CT & Y & CT & FC_Char & "Ignore=(D) Found " & Z & CT & SaveSource & CT & SaveGrammar & CT & MyShowSSS(SSS))
                        Application.DoEvents()
                        SourceLog("", 1339, SourceForm, Err8, "FOUND ?????  [" & Depth.ToString & "," & SSS.AtParse.ToString & "]" & Z & CT & SaveSource & CT & SaveGrammar & CT & MyShowSSS(SSS))
                        Return MatchSyntax2Source & Z & vbCrLf & FindKeyWordGraphic(Parse(SaveGrammar, ": :=", ""))
                    End If
                End If
                X = DropParse(X, " | ", "")
            End While
            Return MatchSyntax2Source & MatchSyntax2Source(SourceForm, SSS, Depth + 1)
        End If

        'Next see if we are testing a grammar rule named
        Z = X
        X = IsThisAGrammarRule(Parse(SSS.GrammarLine, "", U.PARSE1))
        If X <> Z Then
            SourceLog("", 1340, SourceForm, Err8, "[" & Depth.ToString & "," & SSS.AtParse.ToString & "] Problem???")
        End If
        If X <> "" Then
            SourceLog("", 1341, SourceForm, Err8, "[" & Depth.ToString & "," & SSS.AtParse.ToString &
                      "] ^^^^^^^^ This Is a Grammar Rule Name :  " & SSS.GrammarLine &
                      CT & "What=" & Y &
                      CT & "code=" & SSS.Source(SSS.AtParse))
            '''''Source.ListBoxFlowChart.Items.add(FC_Char & "Ignore=(E) Trying  --> " & SSS.GrammarLine)
            SSS.GrammarLine = X
            'todo need to restore the grammar line before turning
            X = MatchSyntax2Source(SourceForm, SSS, Depth + 1)
            Return MatchSyntax2Source & X
        End If

        DebugLog("", 1210, Err8, "[" & Depth.ToString & "," & SSS.AtParse.ToString & "] do literals match " & HL(SSS.Source(SSS.AtParse)) & HL(SSS.GrammarLine))
        Y = IsThisATerminalMatch(SourceForm, SSS.Source(SSS.AtParse), SSS.GrammarLine, Depth + 1)
        If Y <> "" Then
            SourceLog("", 1342, SourceForm, Err8, "[" & Depth.ToString & "," & SSS.AtParse.ToString & "] Match Keyword " & vbTab & SSS.AtParse.ToString & ")" & SSS.Source(SSS.AtParse) & vbTab & HL(ToHex(SSS.GrammarLine)) & vbTab & Y)
            AddGraphic(SourceForm, SSS, FC_Char & "Ignore=(F) " & Y & "    " &
                       vbCrLf & FC_Char & "ignore=(G)" & SaveSource & "," & SaveGrammar & "[" & Depth.ToString & "," & SSS.AtParse.ToString & "] Match Keyword " & vbTab & SSS.Source(SSS.AtParse) & vbTab & HL(ToHex(SSS.GrammarLine)) & vbTab & Y, Depth)
            Application.DoEvents()
            X = Success(SourceForm, SSS, Depth)
            ReDim Preserve SSS.Source2FlowChart(SSS.AtParse + 1)
            SSS.Source2FlowChart(SSS.AtParse) &= SSS.GrammarRuleName & ","
            Log("", 1055, Err6, MyShowSSS(SSS))
            Y = X
            While X <> "" And X <> vbCrLf & FC_Char
                Log("", 1056, Err8, Parse(X, vbCrLf, "") & FC_Char & "Temp Grammar=" & SaveGrammar & FC_Char & "Save Source=" & SaveSource)
                LineNumberString = ""
                If Bin2Str(SSS.Source_Line_Number, 5) & ":" & Bin2Str(SSS.AtParse, 3) & "=" <> Left(Parse(X, vbCrLf, ""), 6) Then LineNumberString = Bin2Str(SSS.Source_Line_Number, 5) & ":" & Bin2Str(SSS.AtParse, 3) & "="
                AddGraphic(SourceForm, SSS, Parse(X, vbCrLf, "") &
                    vbCrLf & FC_Char & "ignore=(H)temp Grammar=" & SaveGrammar &
                    vbCrLf & FC_Char & "ignore=(I)save source=" & SaveSource, Depth)
                Application.DoEvents()
                X = DropParse(X, vbCrLf, "")
            End While
            Application.DoEvents()
            SSS.Source2FlowChart(SSS.AtParse) &= "," & Y
            SSS.AtParse += 1
            'SSS.GrammarLine = SaveGrammar
            'ParseAll(SourceForm, SSS)
            Return SSS.Source2FlowChart(SSS.AtParse - 1)
            'MatchSyntax2Source & "," & Bin2Str(SSS.Source_Line_Number, 5) & ":" & Bin2Str(SSS.AtParse, 3) & "=" & FC_Char & "constant=" & SSS.GrammarRuleName & "," & SSS.Source(SSS.AtParse)
        End If
        Return "" '        Return MatchSyntax2Source & Failed(SourceForm, 9201, SSS.Source(sss.atparse), SSS.Grammars(sss.atparse), SSS.Grammars(sss.atparse) & "   ---   " & SSS.Source(sss.atparse), Depth + 1)
    End Function

    'todo assume that this is source code or fccmd and color/font it (change text box to Rich Text Box.)
    'If the syntax matches the source code then return the name of the rule it matched, else return ""
    Friend Function MatchSyntax2Source_X(ByRef SourceForm As Source, ByRef SSS As SourceSyntaxStructure, Depth As Integer) As String
        Dim ThisRuleName, Rules, SourceCodeAtom, FirstRuleName, RuleResults, GrammarNameOption, Temp As String
        Dim I, J, K As Integer
        Dim RTN As String = ""
        Dim Temp1 As FlowChartRecordStructure
        'Saving them incase I need to store them on failure (things with '|' )
        Dim SaveGrammarLine As String = SSS.GrammarLine
        Dim SaveSourceCodeLine As String = SSS.SourceCodeLine
        ParseAll(SourceForm, SSS) ' reparse everything, just to make sure that we are dealing with the current atom 
        SourceCodeAtom = ""


        'ThisRuleName is if the entire rule is passed, other wise this is only part (or all) of a rule
        If Parse(SSS.GrammarLine, "::=", "") <> "" Then
            ThisRuleName = Trim(Parse(SSS.GrammarLine, "::=", ""))
            Rules = DropParse(SSS.GrammarLine, "::=", "")
        Else 'This is just rules to be tested, not the orginal *(recursive)
            ThisRuleName = ""
            Rules = SSS.GrammarLine
        End If
        FirstRuleName = Parse(Rules, "", U.PARSE1)
        GrammarNameOption = ""
        'Y is the first parseable word (keyword, variablename, special characters ...
        If IsThisACharacterRange(Rules) <> "" Then
            If Hex2Bin(Rules) = Hex2Bin(ToHex(SSS.Source(SSS.AtParse))) Then
                SourceLog("", 1342, SourceForm, Err6, "***** matches  (" & ThisRuleName & "," & ToHex(SSS.Source(SSS.AtParse)) & ")")
                Temp1 = ConverLineNumber2FlowChartXY(SourceForm, SSS.AtParse)
                RTN &= Success(SourceForm, SSS, Depth)
            Else
                SourceLog("", 1343, SourceForm, Err7, FailedLog(SourceForm, 1343, SSS.Source(SSS.AtParse), Rules, "< Does Not Match >", Depth + 1))
            End If
        End If


        'Now The grammar only has the rules, 

        If ThisRuleName <> "" Then SSS.GrammarLine = DropParse(SSS.GrammarLine, "::=")
        ParseAll(SourceForm, SSS)

        If InStr(Rules, " | ") <> 0 Then
            While Len(Rules) > 0
                Do
                    Temp = Parse(Rules, "", U.PARSE1)
                    '''''Temp2 = SSS.SourceCodeLine
                    SSS.GrammarLine = Parse(Temp, "|", "")
                    RuleResults = MatchSyntax2Source(SourceForm, SSS, Depth + 1)
                    SourceLog("", 1344, SourceForm, Err8, "  ***** Looping through options : >" & RuleResults & "<" & vbTab & Temp)

                    'This should see if this is a conditional grammar rule name 
                    If RuleResults <> "" Then
                        If IsThisAGrammarRule(Parse(RuleResults, "", U.PARSE1)) <> "" Then
                            GrammarNameOption = SSS.Source(I + 1)
                        Else
                            GrammarNameOption = ""
                        End If

                        If IsThisAGrammarRule(RuleResults) <> "" Then
                            SourceLog("", 1345, SourceForm, Err8, ToHex(FirstRuleName) & " Try This Rule : " & HL(IsThisAGrammarRule(FirstRuleName)) & vbTab & vbTab & HL(Temp))
                        Else
                            SourceLog("", 1346, SourceForm, Err6, "This Is Not a rule " & HL(FirstRuleName) & vbTab & HL(IsThisAGrammarRule(FirstRuleName)) & vbTab & HL(Temp))
                        End If
                    End If ' end of test if RuleResults <> "" also 
                    If FirstRuleName <> "" Then
                        SourceLog("", 1347, SourceForm, Err6, "***** matches  (" & ThisRuleName & ") (" & SSS.Source(SSS.AtParse) & ")(" & Rules & ")")
                        RTN &= Success(SourceForm, SSS, Depth)
                    Else
                        SourceLog("", 1348, SourceForm, Err8, "*****  Rule fails " & CT &
                                          HL(ToHex(Parse(Rules, "|", ""))) & "," & CT &
                                          HL("'" & ToHex(Trim(UnParse(SSS.Source))) & "'") & "," & CT &
                                          HL("'" & ToHex(FirstRuleName)) & "'")
                    End If
                    Temp = DropParse(Rules, U.PARSE1)
                Loop While Len(Temp) > 0
                Rules = TrimOff(Rules, "|")
                Rules = DropParse(Rules, "|") 'end of this options
                DebugLog("", 1211, Err6, "***** Matched rulename 1) " & HL(FirstRuleName) & CT & "  rule 2) " & HL(Rules) & CT & " rule testing 3)" & Parse(Rules, "|", ""))
            End While
        End If

        If FirstRuleName = "" And IsThisAGrammarRule(ThisRuleName) <> "" Then
            Log("", 1057, Err8, "Recursive here " & FirstRuleName)
            RTN &= MatchSyntax2Source(SourceForm, SSS, Depth + 1)
            SSS.SourceCodeLine = SaveSourceCodeLine
            Return RTN
        Else
            Log("", 1058, Err8, FirstRuleName & vbTab & vbTab & ThisRuleName & vbTab & vbTab & IsThisAGrammarRule(ThisRuleName))
        End If

        J = 1 'pointer to the source array
        K = 1 'pointer to the syntax array
        'todo need to get rid of the white spaces??
        If Len(FirstRuleName) > 0 Then
            SourceLog("", 1350, SourceForm, Err7, FirstRuleName.ToString)
        End If
        For I = 0 To SSS.Source.Length
            'RuleName is the first parsed syntax word
            FirstRuleName = Trim(Parse(Rules, "", U.PARSE1)) ' first syntax word
            If FirstRuleName = "" Then
                DebugLog("", 1212, Err6, "***** fails The rule is blank " & HL(ToHex(FirstRuleName)) & vbTab & HL(ToHex(ThisRuleName)) & CT & HL(ToHex(Rules)) & vbTab & HL(ToHex(SourceCodeAtom)) & vbTab & "RTN = " & HL(RTN))
                RTN &= FailedLog(SourceForm, 1394, SSS.Source(I).ToString, I.ToString & ")" & Rules, "*****  The Rule is blank " & ThisRuleName, Depth + 1)
                Return RTN
            End If
            'A is the first parsed source code word
            If I < SSS.Source.Length Then
                SourceCodeAtom = SSS.Source(I)
                'todo why am I checking for a grammar name when dealing with a source code line?
            Else
                SourceLog("", 1351, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                RTN &= FailedLog(SourceForm, 1395, SSS.SourceCodeLine, Rules, "The number of atoms is less than the number of rules ", Depth + 1)
            End If

            'Does the source word match the syntax word
            'source code ':' syntax code
            'Each loop that gets return something should add another point to the symbol
            SourceLog("", 1352, SourceForm, Err8, HL(SourceCodeAtom) & vbTab & vbTab & WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName) & vbTab & vbTab & HL(FirstRuleName))
            Select Case WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName)
                Case "BLANK:Keyword"
                    SourceLog("", 1353, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                    WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "BLANK:Literal"
                    SourceLog("", 1354, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                    SourceLog("", 1355, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    If SourceCodeAtom = Mid(FirstRuleName, 2, Len(FirstRuleName) - 2) Then
                        SourceLog("", 1356, SourceForm, Err6, "***** matches  This is a keyword match ")
                        RTN &= Success(SourceForm, SSS, Depth)
                    End If
                Case "BLANK:BNF"
                    SourceLog("", 1357, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "BLANK:Variable"
                    SourceLog("", 1358, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                Case "BLANK:WhiteSpace"
                    SourceLog("", 1359, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName) & CT &
                              "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                    RTN &= "(WhiteSpace)"
                    SourceLog("", 1360, SourceForm, Err6, "White Space " & RTN)
                Case "BLANK:Color"
                    SourceLog("", 1361, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName) & CT &
                              "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                Case "BLANK:DataType"
                    SourceLog("", 1362, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName) & CT &
                              "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                Case "BLANK:Special"
                    SourceLog("", 1363, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    SourceLog("", 1427, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                Case "BNF:Keyword"
                    SourceLog("", 1428, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName) & CT &
                              "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                Case "BNF:Literal"
                    SourceLog("", 1429, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName) & CT &
                              "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                Case "BNF:BNF"
                    SourceLog("", 1430, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName) & CT &
                              "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                Case "BNF:Variable"
                    SourceLog("", 1431, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName) & CT &
                              "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                Case "BNF:WhiteSpace"
                    SourceLog("", 1432, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName) & CT &
                              "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                Case "BNF:Color"
                    SourceLog("", 1433, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "BNF:DataType"
                    SourceLog("", 1434, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "BNF:Special"
                    SourceLog("", 1435, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Character:Keyword"
                    SourceLog("", 1436, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Character:Literal"
                    SourceLog("", 1437, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    If SourceCodeAtom = Mid(FirstRuleName, 2, Len(FirstRuleName) - 2) Then
                        SourceLog("", 1364, SourceForm, Err6, "***** matches  This is a keyword match ")
                        RTN &= Success(SourceForm, SSS, Depth)
                    End If
                    SourceLog("", 1364, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Character:BNF"
                    SourceLog("", 1438, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Character:Variable"
                    SourceLog("", 1439, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Character:WhiteSpace"
                    SourceLog("", 1441, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Character:Color"
                    SourceLog("", 1442, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Character:DataType"
                    SourceLog("", 1443, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Character:Special"
                    SourceLog("", 1444, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Color:Keyword"
                    SourceLog("", 1445, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Color:Literal"
                    SourceLog("", 1446, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Color:BNF"
                    SourceLog("", 1447, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Color:Variable"
                    SourceLog("", 1448, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Color:WhiteSpace"
                    SourceLog("", 1449, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Color:Color"
                    SourceLog("", 1451, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Color:DataType"
                    SourceLog("", 1452, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Color:Special"
                    SourceLog("", 1453, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "DataType:Keyword"
                    SourceLog("", 1454, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "DataType:Literal"
                    SourceLog("", 1455, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "DataType:BNF"
                    SourceLog("", 1456, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "DataType:Variable"
                    SourceLog("", 1457, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "DataType:WhiteSpace"
                    SourceLog("", 1458, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "DataType:Color"
                    SourceLog("", 1459, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "DataType:DataType"
                    SourceLog("", 1461, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                Case "DataType:Special"
                    SourceLog("", 1462, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                Case "KeyWord:BLANK"
                    Select Case MsgBox("Need to get rid of blank grammer rules. ", MsgBoxStyle.AbortRetryIgnore, WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                        Case MsgBoxResult.Abort
                            Application.Exit()
                        Case MsgBoxResult.Cancel
                        Case MsgBoxResult.Ignore
                        Case MsgBoxResult.No
                        Case MsgBoxResult.Ok
                        Case MsgBoxResult.Retry
                        Case MsgBoxResult.Yes
                    End Select
                    SourceLog("", 1365, SourceForm, Err8, "MatchSyntax2Source" &
                     CT & WhatIsThis(SourceForm, SourceCodeAtom) &
                     vbTab & ToHex(SourceCodeAtom) &
                     CT & WhatIsThis(SourceForm, FirstRuleName) &
                     vbTab & ToHex(FirstRuleName) &
                     CT & HL(WhatIsThis(SourceForm, Rules)) &
                     vbTab & HL(SSS.Source(SSS.AtParse)) &
                     "-->" & HL(SSS.Grammars(SSS.AtParse)) &
                     "-->" & HL(SSS.Grammars(SSS.AtParse + 1)))
                    RTN &= FailedLog(SourceForm, 1395, FC_Char & "Message= There should never be a blank for a grammer rule ", FirstRuleName & ", " & ThisRuleName, "Blank rule for " & ThisRuleName, Depth + 1)
                Case "Keyword:Keyword"
                    SourceLog("", 1463, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT & WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    If SourceCodeAtom = FirstRuleName Then
                        ReDim Preserve SSS.Grammars(SSS.Grammars.Length + 1)
                        ReDim Preserve SSS.GrammarWhat(SSS.Grammars.Length + 1)
                        ReDim Preserve SSS.Source(SSS.Source.Length + 1)
                        ReDim Preserve SSS.SourceWhat(SSS.Source.Length + 1)
                        SSS.Grammars(SSS.Grammars.Length) = SourceCodeAtom
                        SSS.Grammars(SSS.GrammarWhat.Length) = WhatIsThis(SourceForm, SourceCodeAtom)
                        SSS.Source(SSS.Source.Length) = FirstRuleName
                        SSS.Source(SSS.SourceWhat.Length) = WhatIsThis(SourceForm, FirstRuleName)
                    End If
                Case "Keyword:Literal"
                    SourceLog("", 1464, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    If SourceCodeAtom = Mid(FirstRuleName, 2, Len(FirstRuleName) - 2) Then
                        SourceLog("", 1365, SourceForm, Err6, "***** matches  This is a keyword match ")
                        RTN &= Success(SourceForm, SSS, Depth)
                    End If
                Case "Keyword:BNF"
                    SourceLog("", 1465, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    If FirstRuleName = "" And IsThisAGrammarRule(FirstRuleName) <> "" Then
                        Log("", 1063, Err8, "Recursive here " & FirstRuleName)
                        RTN &= MatchSyntax2Source(SourceForm, SSS, Depth + 1)
                        SSS.SourceCodeLine = SaveSourceCodeLine
                        Return RTN
                    End If
                    RTN &= MatchSyntax2Source(SourceForm, SSS, Depth + 1)
                Case "Keyword:Variable" 'assume keywords are not variables.
                    SourceLog("", 1466, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Keyword:WhiteSpace"
                    SourceLog("", 1467, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Keyword:Color"
                    SourceLog("", 1468, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Keyword:DataType"
                    SourceLog("", 1469, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Keyword:Special"
                    SourceLog("", 1471, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Literal:Keyword"
                    SourceLog("", 1472, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Literal:Literal"
                    SourceLog("", 1473, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Literal:BNF"
                    SourceLog("", 1474, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Literal:Variable"
                    SourceLog("", 1475, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Literal:WhiteSpace"
                    SourceLog("", 1476, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Literal:Color"
                    SourceLog("", 1477, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Literal:DataType"
                    SourceLog("", 1478, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Literal:Special"
                    SourceLog("", 1479, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Special:BLANK"
                    SourceLog("", 1365, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Special:Keyword"
                    SourceLog("", 1481, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Special:Literal"
                    SourceLog("", 1482, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Special:BNF"
                    SourceLog("", 1483, SourceForm, Err6, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Special:Variable"
                    SourceLog("", 1484, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Special:WhiteSpace"
                    SourceLog("", 1485, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Special:Color"
                    SourceLog("", 1486, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Special:DataType"
                    SourceLog("", 1487, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Special:Special"
                    SourceLog("", 1488, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                    WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Variable:BLANK"
                    SourceLog("", 1366, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Variable:Keyword"
                    SourceLog("", 1489, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Variable:Literal"
                    SourceLog("", 1491, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Variable:BNF" 'todo error a variable can not be a rule name
                    SourceLog("", 1366, SourceForm, Err4, "***** This Source word " & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "--> can not match this grammar word " & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName) & CT &
                              "***** Failed MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    Log("", 1064, Err8, I.ToString & ">=" & SSS.Grammars.Length.ToString & vbTab & SSS.Source.Length.ToString)
                    SourceLog("", 1367, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    RTN &= FailedLog(SourceForm, 1397, SourceCodeAtom, Rules, "Variable name does not match rule requirement ", Depth + 1)
                    SourceLog("", 1368, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Variable:Variable"
                    SourceLog("", 1493, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Variable:WhiteSpace"
                    SourceLog("", 1494, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Variable:Color"
                    SourceLog("", 1495, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Variable:DataType"
                    SourceLog("", 1496, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "Variable:Special"
                    SourceLog("", 1497, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "WhiteSpace:BLANK"
                    SourceLog("", 1498, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "WhiteSpace:Character"
                    SourceLog("", 1370, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName) & CT &
                              "***** fails >" & HL(ToHex(FirstRuleName)) & vbTab & HL(ToHex(ThisRuleName)) & CT & HL(ToHex(Rules)) & vbTab & HL(ToHex(SourceCodeAtom)) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    RTN &= FailedLog(SourceForm, 1398, SourceCodeAtom, Rules, "White Space where a character was expected", Depth + 1)
                Case "WhiteSpace:Keyword"
                    SourceLog("", 1371, SourceForm, Err7, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    'todo need to see if the keyword is a whitespace
                    If FirstRuleName = "Unicode" Then
                        If MyUnicode.UnicodeClassCode(SourceCodeAtom) = Rules Then
                            RTN &= "(" & SourceCodeAtom & vbTab & Rules & ")"
                        Else
                            SourceLog("", 1372, SourceForm, Err6, "***** fails " & HL(ToHex(ThisRuleName)) & CT & HL(ToHex(Rules)) & CT & HL(ToHex(SourceCodeAtom)) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                            RTN &= FailedLog(SourceForm, 1399, SourceCodeAtom, Rules, "White Space where a keyword was expected ", Depth + 1)
                        End If
                    End If
                    SourceLog("", 1373, SourceForm, Err1, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    If IsThisACharacterRange(SourceCodeAtom) <> "" Then
                        SourceLog("", 1374, SourceForm, Err8, "HALT " & IsThisACharacterRange(FirstRuleName))
                    End If
                    SourceLog("", 1375, SourceForm, Err8, "***** fails " & HL(ToHex(FirstRuleName)) & vbTab & HL(ToHex(ThisRuleName)) & CT & HL(ToHex(Rules)) & vbTab & HL(ToHex(SourceCodeAtom)) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    RTN &= Success(SourceForm, SSS, Depth)
                Case "WhiteSpace:Literal"
                    'todo need to check that the literal is not a white space
                    SourceLog("", 1376, SourceForm, Err7, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    If Len(SourceCodeAtom) <> 0 And Len(FirstRuleName) <> 0 Then
                        If AscW(SourceCodeAtom) = Hex2Bin(FirstRuleName) Then
                            RTN &= "(" & SourceCodeAtom & ")"
                        Else
                            SourceLog("", 1377, SourceForm, Err6, "***** fails Rule=" & Rules & vbTab & " Source=" & ToHex(SourceCodeAtom))
                            RTN &= FailedLog(SourceForm, 1400, SourceCodeAtom, Rules, "White space where a literal was expected ", Depth + 1)
                        End If
                    Else
                        SourceLog("", 1377, SourceForm, Err6, "***** fails Rule=" & ThisRuleName & vbTab & " Rule=" & Rules & vbTab & " Source=" & ToHex(SourceCodeAtom) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                        RTN &= FailedLog(SourceForm, 1401, SourceCodeAtom, Rules, "White space where a literal was expected ", Depth + 1) 'failed to match 
                    End If
                Case "WhiteSpace:BNF"
                    SourceLog("", 1378, SourceForm, Err7, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    '                            HL(ToHex(RuleName)) & vbTab &
                    SSS.GrammarLine = IsThisAGrammarRule(Parse(Rules, "", U.PARSE1))
                    ParseAll(SourceForm, SSS)
                    Temp = MatchSyntax2Source(SourceForm, SSS, Depth + 1)
                    If Temp = "" Then
                        SourceLog("", 1380, SourceForm, Err6, "***** fails " &
                            HL(Temp) & vbTab &
                            HL(ToHex(FirstRuleName)) & vbTab &
                        HL(ToHex(Rules)) & vbTab &
                            HL(ToHex(SourceCodeAtom)) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                        RTN &= FailedLog(SourceForm, 1404, SourceCodeAtom, Rules, "White space where a rule was expected", Depth + 1)
                    Else
                        SourceLog("", 1381, SourceForm, Err6, "***** matches " & HL(Temp) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                        RTN &= Success(SourceForm, SSS, Depth)
                    End If
                    SourceLog("", 1382, SourceForm, Err8, WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "WhiteSpace:Variable"
                    SourceLog("", 1383, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "WhiteSpace:WhiteSpace"
                    SourceLog("", 1384, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1))
                    RTN &= "(WhiteSpace)"
                Case "WhiteSpace:Color"
                    SourceLog("", 1385, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "WhiteSpace:DataType"
                    SourceLog("", 1386, SourceForm, Err8, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case "WhiteSpace:Special"
                    SourceLog("", 1387, SourceForm, Err7, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                    If FirstRuleName = ";" Then
                        SourceLog("", 1388, SourceForm, Err6, "***** matches " & SourceCodeAtom & CT &
                                  WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                        RTN &= Success(SourceForm, SSS, Depth) ';This is the end of the syntax
                    End If
                    SourceLog("", 1389, SourceForm, Err3, "MatchSyntax2Source" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & vbTab & ToHex(SourceCodeAtom) & CT & WhatIsThis(SourceForm, FirstRuleName) & vbTab & ToHex(FirstRuleName) & CT & WhatIsThis(SourceForm, Rules) & vbTab & SSS.Source(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse) & "-->" & SSS.Grammars(SSS.AtParse + 1) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
                Case Else
                    DebugLog("", 1213, Err1, "[" & SourceCodeAtom & "," & ToHex(SourceCodeAtom) & "]" & CT & WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName) & CT & "[" & FirstRuleName & ToHex(FirstRuleName) & "]")
                    SourceLog("", 1390, SourceForm, Err8, "Program Problem unknown combinition ---> " & WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
            End Select
        Next
        SourceLog("", 1391, SourceForm, Err6, "***** Syntax 2 Source RESULTS " & HL(RTN) & CT &
                              WhatIsThis(SourceForm, SourceCodeAtom) & ":" & WhatIsThis(SourceForm, FirstRuleName))
        Return RTN
    End Function


    Friend Sub WhatAreThese(ByRef SourceForm As Source, ByRef SSS As SourceSyntaxStructure)
        Dim I As Integer
        If SSS.Grammars.Length <> SSS.GrammarWhat.Length Then
            ReDim SSS.GrammarWhat(SSS.Grammars.Length - 1)
        End If
        For I = 0 To SSS.Grammars.Length - 1
            SSS.GrammarWhat(I) = WhatIsThis(SourceForm, SSS.Grammars(I))
        Next


        If SSS.Source.Length <> SSS.SourceWhat.Length Then
            ReDim SSS.SourceWhat(SSS.Source.Length - 1)
        End If
        For I = 0 To SSS.Source.Length - 1
            SSS.SourceWhat(I) = WhatIsThis(SourceForm, SSS.Source(I))
        Next
    End Sub



    Friend Function WhatIsThis(SourceForm As Source, LineIn As String) As String
        Dim S As String
        If LineIn = "" Then Return "BLANK"
        S = Parse(LineIn, U.PARSE1, "")

        S = Trim(Parse(LineIn, "::=", ""))
        S = Trim(Parse(S, ",", ""))
        If S <> " " Then S = Trim(Parse(S, " ", ""))
        'S = Trim(Parse(S, vbTab))

        If IsThisAKeyword(S) <> "" Then Return "Keyword"
        If Left(LineIn, 8) = "Unicode." Then Return "Unicode"
        If IsThisALiteral(S) <> "" Then Return "Literal"
        If IsThisAGrammarRule(S) <> "" Then Return "BNF"
        If IsThisAWhiteSpace(S) <> "" Then Return "WhiteSpace"
        If IsThisACharacterRange(S) <> "" Then Return "Character"
        If LineIn = " " Then Return "WhiteSpace" 'todo this should have return white space from above
        If IsThisAColor(S) <> "" Then Return "Color"
        If IsthisADataType(S) <> "" Then Return "DataType"
        If IsThisANumber(S) <> "" Then Return "Number"
        If IsThisASpecial(S) <> "" Then Return "Special"
        If IsThisAVariable(S) <> "" Then
            If FindVariable(SourceForm, S) = "" Then
                AddVariable(SourceForm, S)
            End If
            Return "Variable"
        End If
        If LineIn = "'" Or LineIn = Chr(34) Then
            Return "Literal"
        End If
        Log("", 1065, Err1, "do you know what this is " & HL(S))
        Return "Unknown"
    End Function




    'todo This should be extended to say the field name (of the format, so that input can be in any format. instead of string $, or number # check only)
    'This should check that the input data is in the correct format 
    Friend Sub MatchFormat(FMTNumber As Integer, MyData As String)
        Dim MyFormat As String
        Dim A, B, X, Y As String
        Dim Flag As Boolean = True
        If FMTNumber <> 0 Then
            MyFormat = GetOption(FMTNumber)
            X = MyFormat
            X = Trim(DropParse(X, ",", "")) 'drop the number
            X = Trim(DropParse(X, ",", "")) ' drop the /command
            Y = MyData
            While X <> "" And Y <> ""
                A = Parse(X, ",", "")
                B = Parse(Y, ",", "")
                'Log(1058, Err6, A & vbTab & vbTab & B)
                Select Case A
                    Case "$" 'character string
                    'everycharacter string is OK
                    Case "#" 'must be an integer (for now)
                        If IsThisANumber(B) = "" Then Flag = False
                    Case Else
                        Log("", 1066, Err1, "Error in format " & CT & X & CT & Y)
                End Select
                X = Trim(DropParse(X, ",", ""))
                Y = DropParse(Y, ",", "")
            End While
        Else
            'special, Lines format #0, color, x1,y1,x2,xy [,x3,x4... | color,x1,y1,x2,y2,...]
            MyFormat = "Lines"
            'need to test for a color being first
            X = MyData
            While Len(X) > 0
                Y = Parse(X, ",", "") : X = Trim(DropParse(X, ",", ""))
                If IsThisAColor(Y) <> "" Then 'This is a color name 
                    Y = Parse(X, ",", "") : X = Trim(DropParse(X, ",", "")) 'X1
                    If IsThisANumber(Y) = "" Then Log("", 1066, Err1, "This format is incorrect for a X,Y " & HL(ToHex(Y)) & ", " & HL(ToHex(X)) & CT & ToHex(MyData))
                    Y = Parse(X, ",", "") : X = Trim(DropParse(X, ",", "")) 'Y1
                    If IsThisANumber(Y) = "" Then Log("", 1067, Err1, "This format is incorrect for a X,Y " & HL(ToHex(Y)) & ", " & HL(ToHex(X)) & CT & ToHex(MyData))
                    Y = Parse(X, ",", "") : X = Trim(DropParse(X, ",", "")) 'x2
                    If IsThisANumber(Y) = "" Then Log("", 1068, Err1, "This format is incorrect for a X,Y " & HL(ToHex(Y)) & ", " & HL(ToHex(X)) & CT & ToHex(MyData))
                    Y = Parse(X, ",", "") : X = Trim(DropParse(X, ",", "")) 'y2
                    If IsThisANumber(Y) = "" Then Log("", 1069, Err1, "This format is incorrect for a X,Y " & HL(ToHex(Y)) & ", " & HL(ToHex(X)) & CT & ToHex(MyData))
                ElseIf IsThisANumber(Y) <> "" Then ' This is x3...
                    'X = Trim(DropParse(X, ",", "")) 'X3....
                    Y = Parse(X, ",", "") : X = Trim(DropParse(X, ",", "")) 'Y3....
                    If IsThisANumber(Y) = "" Then Log("", 1070, Err1, "This format is incorrect for a X,Y " & HL(ToHex(Y)) & ", " & HL(ToHex(X)) & CT & ToHex(MyData))
                Else
                    Log("", 1071, Err1, "The definition file input is not in a valid format " & CT & MyFormat & CT & MyData & CT & HL(Y) & vbTab & vbTab & HL(X))
                End If
            End While
        End If
        If Flag <> True Then
            Log("", 1072, Err1, "The definition file input is not in a valid format " & CT & MyFormat & CT & MyData)
        End If
    End Sub

    Friend Function WhatDataTypeIsThis(ByRef SourceForm As Source, S As String) As String
        If Len(S) = 0 Then Return "Error"
        Select Case WhatIsThis(SourceForm, S)
            Case "Keyword", "Special"
                Return "Error"
            Case "variable"
                Return FindDataType(S)
            Case "literal"
                Return "String"
            Case "Number"
                Return "Number"
            Case Else
                Log("", 1073, Err1, S & " is an unknown flow chart catagory " & WhatIsThis(SourceForm, S))
        End Select
        Return "Error"
    End Function


    Friend Function CheckNotInList(ByRef LB As ListBox, WhatToFind As String) As Integer
        Return -1
    End Function


    Friend Function BinarySearch4Index(ListB As ListBox, Searching_For As String) As Integer ', Optional ClosestTo As String = "") As Integer
        Dim Low, High, At As Integer
        Dim D1, D2, D3, D4 As String
        D1 = "" : D2 = "" : D3 = ""
        If Searching_For = "" Then Return 0 ' can never find nothing
        'If ClosestTo <> "" Then Searching_For = ClosestTo
        D4 = Parse(Searching_For, "::=", "")
        D4 = LCase(Trim(Parse(D4, ",", "")))
        Low = 0
        High = ListB.Items.Count - 1
        At = CInt(High / 2)
        If High < 1 Or At < 1 Then
            Return 0
        End If
        If LCase(D4) < LCase(ListB.Items.Item(0).ToString) Then
            'If ClosestTo <> "" Then Return 0
            Return 0
        End If
        If LCase(D4) > LCase(ListB.Items.Item(High).ToString) Then
            DebugLog("", 1214, Err8, D4 & "   Higher than last one " & D4 & CT & ListB.Items.Item(High - 1).ToString & CT & ListB.Items.Item(High).ToString)
            'If ClosestTo <> "" Then Return High
            Return 0
        End If
        While D4 <> ListB.Items.Item(At).ToString And High - Low > 1
            D1 = LCase(ListB.Items.Item(Low).ToString)
            D2 = LCase(ListB.Items.Item(At).ToString)
            D3 = LCase(ListB.Items.Item(High).ToString)
            If D4 > D2 Then Low = At
            If D4 < D2 Then High = At
            If D4 = D2 Then Return At
            At = CInt((High + Low) / 2)
        End While


        If LCase(Left(ListB.Items.Item(Low).ToString, Len(D4))) = D4 Then Return Low
        If D4 = LCase(ListB.Items.Item(At).ToString.ToString) Then Return At
        If LCase(Left(ListB.Items.Item(At).ToString, Len(D4))) = LCase(D4) Then Return At
        If LCase(Left(ListB.Items.Item(High).ToString, Len(D4))) = LCase(D4) Then Return High

        If D4 <> CStr(ListB.Items.Item(At)) Then
            If D4 < CStr(ListB.Items.Item(High)) Then
                If D4 > CStr(ListB.Items.Item(Low)) Then
                    'If ClosestTo <> "" Then Return At
                    Return 0
                End If
            End If
        End If
        If D4 = D1 Then Return Low
        If D4 = D2 Then Return At
        If D4 = D3 Then Return High
        Return 0
    End Function


    Friend Function BinarySearchList(ListB As ListBox, SearchFor As String, Optional ClosestTo As String = "") As String
        Dim Low, High, At As Integer
        Dim D1, D2, D3, D4, Searching_For, MyDefault As String
        D1 = "" : D2 = "" : D3 = "" : MyDefault = ""
        Searching_For = Trim(SearchFor)
        If Searching_For = "" Then Searching_For = Trim(ClosestTo)
        If Searching_For = "" Then Return ""
        D4 = Parse(Searching_For, "::=", "")
        D4 = LCase(Trim(Parse(D4, ",", "")))
        Low = 0
        High = ListB.Items.Count - 1
        At = CInt(High / 2)
        If High < 1 Or At < 0 Then
            Return ""
        End If
        'Debug.Print(HL(LCase(D4)) & vbTab & vbTab & HL(ToHex(LCase(Trim(Parse(ListB.Items.Item(0).ToString, "", U.PARSE1))))))
        If LCase(D4) < LCase(Trim(Parse(ListB.Items.Item(0).ToString, "", U.PARSE1))) Then
            If ClosestTo <> "" Then MyDefault = ListB.Items.Item(0).ToString
            Return MyDefault
        End If
        If LCase(D4) > LCase(ListB.Items.Item(High).ToString) Then
            DebugLog("", 1215, Err8, "Higher than last one >" & LCase(D4) & "<" & CT & ">" & ListB.Items.Item(High).ToString & "<" & CT & ">" & ListB.Items.Item(High).ToString & "<")
            If ClosestTo <> "" Then MyDefault = ListB.Items.Item(High).ToString
            Return MyDefault
        End If
        While D4 <> ListB.Items.Item(At).ToString And High - Low > 1
            D1 = LCase(ListB.Items.Item(Low).ToString)
            D2 = LCase(ListB.Items.Item(At).ToString)
            D3 = LCase(ListB.Items.Item(High).ToString)
            If D4 > D2 Then Low = At
            If D4 < D2 Then High = At
            If D4 = D2 Then Return ListB.Items.Item(At).ToString
            At = CInt((High + Low) / 2)
        End While



        If LCase(Left(ListB.Items.Item(Low).ToString, Len(D4))) = D4 Then Return ListB.Items.Item(Low).ToString
        If D4 = LCase(ListB.Items.Item(At).ToString.ToString) Then Return ListB.Items.Item(At).ToString
        If LCase(Left(ListB.Items.Item(At).ToString, Len(D4))) = LCase(D4) Then Return ListB.Items.Item(At).ToString
        If LCase(Left(ListB.Items.Item(High).ToString, Len(D4))) = LCase(D4) Then Return ListB.Items.Item(High).ToString

        If D4 <> CStr(ListB.Items.Item(At)) Then
            If D4 < CStr(ListB.Items.Item(High)) Then
                If D4 > CStr(ListB.Items.Item(Low)) Then
                    If ClosestTo <> "" Then MyDefault = ListB.Items.Item(At).ToString
                    Return MyDefault
                End If
            End If
        End If
        If D4 = D1 Then Return ListB.Items.Item(Low).ToString
        If D4 = D2 Then Return ListB.Items.Item(At).ToString
        If D4 = D3 Then Return ListB.Items.Item(High).ToString
        Return ""
    End Function



    Friend Function IsThisAUnicode(s As String) As String
        Return MyUnicode.UnicodeClassCode(s)
    End Function
    'assume that everything else has been tried
    Friend Function IsThisAVariable(s As String) As String
        Dim I As Integer
        I = 1
        'rule one must start with a character
        'OR underline (for microsoft stuff)
        Select Case MyUnicode.UnicodeClassCode(s)
                'codes that can not be the start of a variable name
            Case "", "Unicode.Ws", "Unicode.Cc", "Unicode.Ps", "Unicode.Po", "Unicode.Sm", "Unicode.Nd"
                Log("", 1074, Err5, ToHex(Mid(s, I, 1)) & " Unicode class " & MyUnicode.UnicodeClassCode(Mid(s, I, 1)))
                Return ""
            Case "Unicode.Lu", "Unicode.Ll"
                DebugLog("", 1216, Err8, ToHex(Mid(s, I, 1)) & " Unicode class " & MyUnicode.UnicodeClassCode(Mid(s, I, 1)))
            Case Else
                Log("", 1075, Err1, ToHex(Mid(s, I, 1)) & " Unicode class " & MyUnicode.UnicodeClassCode(Mid(s, I, 1)))
        End Select
        For I = 2 To Len(s)
            Select Case MyUnicode.UnicodeClassCode(Mid(s, I, 1))
                    'codes that can not be in a variable name
                Case "Unicode.Ws", "Unicode.Cc"
                    Log("", 1076, Err5, ToHex(Mid(s, I, 1)) & " Unicode class " & MyUnicode.UnicodeClassCode(Mid(s, I, 1)))
                    Return ""
                Case "Unicode.Lu", "Unicode.Ll", "Unicode.Nd" 'These letters are ok
                    If MyDebugLevel > 12 Then DebugLog("", 1217, Err5, "Ok For A Variable Name >" & ToHex(Mid(s, I, 1)) & "< Unicode class " & MyUnicode.UnicodeClassCode(Mid(s, I, 1)))
                Case "unicode.Po", "Unicode.Pe", "Unicode.Pc"
                    Log("", 1077, Err5, ToHex(Mid(s, I, 1)) & " Unicode class " & MyUnicode.UnicodeClassCode(Mid(s, I, 1)) & vbTab & ToHex(s))
                Case Else
                    'todo need to make it so that it checks the possible exceptions
                    Log("", 1078, Err1, "This is not allowed in a non keyword name " & ToHex(Mid(s, I, 1)) & " Unicode class " & MyUnicode.UnicodeClassCode(Mid(s, I, 1)))
                    'todo need to see if it is one of the language dependant characters.
                    Return "" ' not a variable name 
            End Select
        Next
        Return s
    End Function

    Friend Function IsThisASpecial(s As String) As String
        If InStr(U.PARSE1, Left(s, 1)) <> 0 Then Return "Special"
        Return ""
    End Function


    Friend Function IsThisAKeyword(s As String) As String
        Return BinarySearchList(Options.ListBox_KeyWords, s)
    End Function

    Friend Function IsThisAGrammarRuleModifier(s As String) As String
        Select Case Right(s, 1)
            Case "?", "+", "-", "*"
                If Len(s) < 2 Then
                    Return ""
                End If
                If IsThisAGrammarRule(Mid(s, 1, Len(s) - 2)) <> "" Then
                    Return Right(s, 1)
                Else
                    Return Right(s, 1)
                End If
            Case Else
        End Select
        Return ""
    End Function

    Friend Function IsThisAGrammarRule(s As String) As String
        Dim X As String = s
        If IsThisAGrammarRuleModifier(s) <> "" Then
            X = Left(X, Len(X) - 1)
        End If
        Return BinarySearchList(sBNF.sBNF_Grammar, X)
    End Function

    Friend Function IsThisAWhiteSpace(s As String) As String
        If InStr(U.White_Space, Left(s, 1)) <> 0 Then Return Left(s, 1)
        Return ""
    End Function


    'always returns a character range in the form 0x1234-0x4321'
    Friend Function IsThisACharacterRange(s As String) As String
        '       This can have the following formats:
        ' '         a character between squotes
        '0x0011'    A two or four digit hex value
        'a-z'       A character in this range
        '0x1234-0x4321' A character in this range
        'Dim I As Integer
        If Matches("'0X????-0X????'", s) <> "" Then Return s
        If Matches("'0X??-0X??'", s) <> "" Then Return s
        If Matches("'0X????'", s) <> "" Then Return s
        If Matches("'0X??'", s) <> "" Then Return s
        'If Matches("'?'", s) <> "" Then Return s

        If Matches("0X????-0X????", s) <> "" Then Return s
        If Matches("0X??-0X??", s) <> "" Then Return s
        If Matches("0X????", s) <> "" Then Return s
        If Matches("0X??", s) <> "" Then Return s
        'If Matches("?", s) <> "" Then Return s
        If Matches("?-?", s) <> "" Then Return s
        Return ""
    End Function


    Friend Function Matches(MyFormat As String, input As String) As String
        Dim I As Integer
        If input = "" Then Return ""
        Matches = MyFormat
        For I = 1 To Len(MyFormat)
            Select Case Mid(MyFormat, I, 1)
                Case "?" ' This can be anything
                    Mid(Matches, I, 1) = Mid(input, I, 1)
                Case Else
                    If Mid(MyFormat, I, 1) <> Mid(input, I, 1) Then Return ""
            End Select
        Next I
        Return Matches
    End Function


    Friend Function IsthisADataType(S As String) As String
        Return BinarySearchList(Options.ListBoxDataTypes, S)
    End Function
    Friend Function IsThisAColor(s As String) As String
        Return BinarySearchList(Options.ListBoxColors, s, "")
    End Function

    Friend Function IsThisANumber(S As String) As String
        'This should return the string if it is any kind of number 
        'todo expand this test to include integer, readl(decimal), Scientific, ... (Not hex, oct, binary...)
        Dim X As String = ""
        Dim I As Integer = 1
        Dim ValidNumberCharacters As String = "0123456789-"
        If S = "" Then Return ""
        While InStr(ValidNumberCharacters, Mid(S, I, 1)) > 0
            X &= Mid(S, I, 1)
            I += 1
            If I >= Len(S) Then Return X
        End While
        Return X
    End Function
    Friend Function IsThisALiteral(s As String) As String
        Dim X As String
        Dim I As Integer
        X = Left(s, 1)
        If X <> "'" And X <> ChrW(34) Then Return ""
        I = InStr(Mid(s, 2, Len(s)), X)
        If I <> 0 Then
            Return Mid(s, 1, I + 1)
        End If
        Return ""
    End Function



    Friend Function FindKeyWordGraphic(KeyWord As String) As String
        FindKeyWordGraphic = BinarySearchList(Options.ListBoxGrammarGraphics, KeyWord, "")
        If FindKeyWordGraphic <> "" Then FindKeyWordGraphic = FC_Char & FindKeyWordGraphic
    End Function

    Friend Function FindDataType(S As String) As String
        'todo fix this
        'First find out if it is a defined variable name, 
        'second make sure that is is not a keyword.....
        'third make sure that you are actually looking for the datatype of what ever this is

        Return BinarySearchList(Options.ListBoxDataTypes, "", S)
    End Function

    Friend Function FindSymbol(S As String) As String
        Return BinarySearchList(Library.LIB_ListBoxProgramOptions, "", S)
    End Function

    Friend Function FindRuleNameOnly(S As String) As String
        FindRuleNameOnly = Trim(Parse(S, "::=", ""))
    End Function


    Friend Sub AddVariable(SourceFOrm As Source, S As String)
        SourceFOrm.ListBoxVariables.Items.Add(S) ' Add new variable name
    End Sub
    Friend Function FindVariable(SourceForm As Source, S As String) As String
        Return BinarySearchList(SourceForm.ListBoxVariables, S, "")
    End Function

    Friend Sub MakeNewRoutine_notUsed(FileName As String, SourceCode As String)
        Dim newMDIChild1 As New Source
        newMDIChild1.MdiParent = FlowChart2025
        'newMDIChild1.Source.Text = SourceCode
        newMDIChild1.SourceCode.Text = SourceCode.Replace(vbLf, Environment.NewLine)
        Log("", 1078, Err6, "Make new routine " & CT & HL(newMDIChild1.Text) & CT & HL(ToHex(Parse(SourceCode, vbCrLf, ""))))
        newMDIChild1.Text = "Source : " & FileName
        newMDIChild1.Show()
        Application.DoEvents()
    End Sub


    Friend Function Hex2Bin(HexStr As String) As Integer
        If HexStr = "" Then Return 0
        HexStr = TrimOff(HexStr, "'")
        If Left(HexStr, 2) = "0X" Or Left(HexStr, 2) = "0x" Then
            HexStr = Mid(HexStr, 3, Len(HexStr))
        End If
        For i = 1 To HexStr.Length
            If InStr("0123456789AaBbCcDdEeFf", Mid(HexStr, i, 1)) = 0 Then
                Return 0
            End If
        Next
        If Len(HexStr) <> 4 Then Return 0

        Return Convert.ToInt32(HexStr, 16)
    End Function

    Friend Function Bin2Str(N As Integer, Digits As Integer) As String
        Return Right("000000" & N.ToString, Digits)
    End Function


    Friend Function Str2Bin(S1 As String) As Integer
        Dim Sign As Integer
        Dim S As String
        Str2Bin = 0
        Sign = 1
        S = S1
        While Len(S) > 0
            Select Case Left(S, 1)
                Case "0"
                    Str2Bin = Str2Bin * 10 + 0
                Case "1"
                    Str2Bin = Str2Bin * 10 + 1
                Case "2"
                    Str2Bin = Str2Bin * 10 + 2
                Case "3"
                    Str2Bin = Str2Bin * 10 + 3
                Case "4"
                    Str2Bin = Str2Bin * 10 + 4
                Case "5"
                    Str2Bin = Str2Bin * 10 + 5
                Case "6"
                    Str2Bin = Str2Bin * 10 + 6
                Case "7"
                    Str2Bin = Str2Bin * 10 + 7
                Case "8"
                    Str2Bin = Str2Bin * 10 + 8
                Case "9"
                    Str2Bin = Str2Bin * 10 + 9
                Case "-"
                    Sign = -1
                Case Else
                    Log("", 1079, Err1, "This is not a number " & S1)
                    Return Nothing
            End Select
            S = Mid(S, 2)
        End While
        Return Str2Bin * Sign
    End Function



    Friend Function ToHex(s As String) As String
        Dim Digits() As String = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F"}
        Dim I As Integer
        'Dim T(1) as integer
        Dim T As Integer
        Dim Str As String
        If s = "" Then Return "0X????"
        ToHex = s
        I = 0
        While 1 = 1
            If I >= ToHex.Length Then Return ToHex
            I += 1
            Str = Trim(MyUnicode.UnicodeClassCode(Mid(ToHex, I, 1)))
            If Mid(ToHex, I, 1) = " " Then Str = ""
            If Str = "Unicode.Cc" Or Str = "Unicode.Ws" Or AscW(Left(ToHex, 1)) < 32 Then
                T = CByte(AscW(Mid(ToHex, I, 1)))
                'ToHex = Mid(ToHex, 1, I - 1) & "0x" & Convert.ToHexString(T) & Mid(ToHex, I + 1, Len(ToHex))

                ToHex = Mid(ToHex, 1, I - 1) &
                            "'0X" & Digits(CInt(Fix(T / 16))) &
                            Digits(CInt(Fix(T / 256))) &
                            Digits(CInt(Fix(T / 16) Mod 16)) &
                            Digits(CInt(Fix(T Mod 16))) &
                            "'" & Mid(ToHex, I + 1, Len(ToHex))
                DebugLog("", 1218, Err6, "tohex : " & I.ToString & ")>" & ToHex & "< & >" & s & "<")
            Else
            End If
        End While
        Select Case MsgBox("ERROR ToHex()" & CT & s & CT & ToHex & CT & I.ToString & CT & s.Length.ToString)
            Case MsgBoxResult.Abort
                Application.Exit()
            Case MsgBoxResult.Cancel
            Case MsgBoxResult.Ignore
            Case MsgBoxResult.No
            Case MsgBoxResult.Ok
            Case MsgBoxResult.Retry
            Case MsgBoxResult.Yes
        End Select
        Stop
    End Function



    'TODO NEED TO ADD THESE 'special' characters to the /Grammar (overridable)
    'todo need to check that these are in one of the /Grammars ?
    'WS    White Space (tested for first, user controled)
    ' Cc         Control             65
    ' Cf         Format              161
    ' Co         Private Use         0
    ' Cs         Surrogate           0
    ' Ll         Lowercase Letter    2,155
    ' Lm         Modifier Letter     260
    ' Lo         Other Letter        127,004
    ' Lt         TitleCase Letter    31
    ' Lu         Uppercase Letter    1,791
    ' Mc         Spacing Mark        443
    ' Me         Enclosing Mark      13
    ' Mn         NonSpacing Mark     1,839
    ' Nd         Decimal Number      650
    ' Nl         Letter Number       236
    ' No         Other Number        895
    ' Pc         Connector Punctuation 10
    ' Pd         Dash Punctuation    25
    ' Pe         Close Punctuation   73
    ' Pf         Final Punctuation   10
    ' Pi         Initial Punctuation 12
    ' Po         Other Punctuation   593
    ' Ps         Open Punctuation    75
    ' Sc         Currency Symbol     62
    ' Sk         Modifier Symbol     123
    ' Sm         Math Symbol         948
    ' So         Other Symbol        6,431
    ' Zl         Line Separator      1
    ' Zp         Paragraph Separator 1
    ' Zs         Space Separator     17




    'special predefined Grammar key words
    'FCCL_EndOfLine		0x000D	0x000A	0x2028	0x2029
    'FCCL_LessThan		<
    'FCCL_GreaterThan	>
    'FCCL_SQuote		0x0027	0x2018	0x2019
    'FCCL_WhiteSpace    Unicode.Zs  0x0009


    Friend Function GetFieldNumber(MyTableField As String) As Integer
        'This is a data dictionary Junk
        Select Case MyTableField
            Case "Color:Name"     'ColorName
                Return 1
            Case "Color:Alpha"
                Return 2
            Case "Color:Red"
                Return 3
            Case "Color:Green"
                Return 4
            Case "Color:Blue"
                Return 5
            Case "Color:Style"    'Solid, dotted, dashed...
                Return 6
            Case "Color:Start"    'Microsoft line caps
                Return 7
            Case "Color:End"      'Microsoft line caps
                Return 8
            Case Else
                Return -1
        End Select
        Return -1
    End Function


    Friend Function GetMyField(ByRef LB As ListBox, MyTableField As String, RecordNumber As Integer) As String
        Dim X As String
        Dim I As Integer
        X = GetMyRecord(LB, RecordNumber)
        I = GetFieldNumber(MyTableField)
        While I > 0
            I -= 1
            X = Parse(X, ",", "")
        End While
        Return Parse(X, ",", "")
    End Function


    'This returns the 'record' as a string (should return a shared common record union format with all of the possible answers)
    Friend Function GetMyRecord(LB As ListBox, Index As Integer) As String
        Return LB.Items(Index).ToString
    End Function

    Friend Function SetField(Recordin As String, FieldNumber As Integer, ToWhat As String) As String
        Dim I As Integer = 0
        SetField = ""
        While Len(Recordin) > 0
            If I = FieldNumber Then
                SetField &= "," & ToWhat
                Recordin = DropParse(Recordin, ",", "")
            Else
                SetField &= "," & Parse(Recordin, ",", "")
                Recordin = DropParse(Recordin, ",", "")
            End If
        End While
        Return SetField
    End Function

    Friend Sub SetMyField(LB As ListBox, MyField As String, RecordNumber As Integer, ToWhat As String)
        Dim RecordIn, RecordOut As String
        RecordOut = ""
        RecordIn = GetMyRecord(LB, RecordNumber)
        RecordOut = SetField(RecordIn, GetFieldNumber(MyField), ToWhat)
        SetMyRecord(LB, RecordNumber, RecordOut)
    End Sub
    Friend Sub SetMyRecord(LB As ListBox, RecordNumber As Integer, ToWhat As String)
        LB.Items(RecordNumber) = ToWhat
    End Sub


    Public Class MyUnicode
        Public Shared Function UnicodeWhiteSpace(Letter As String) As String
            If Letter = "" Then Return ""
            If InStr(CommonRoutines.U.White_Space, Left(Letter, 1)) <> 0 Then Return "WS"
            Return ""
        End Function


        Public Shared Function Pop_Unicode_Used(Codeline As String) As String
            Dim UCC, RTN As String
            RTN = ""
            UCC = UnicodeClassCode(Codeline) ' get the catagor class of the first letter, and all other letters must be that code
            While IsThisAUnicodeClass(Codeline) ' to end checking
                If UnicodeClassCode(Codeline) = UCC Then
                    RTN = Left(Codeline, 1) 'Get this letter (also)
                    Codeline = Mid(Codeline, 2) 'chip off the first letter, and move all of those other letters over
                Else
                    Return RTN
                End If
            End While
            Return RTN
        End Function


        Public Shared Function UnicodeClassCode(Letter As String) As String
            Dim Category() As String = {"Lu", "Ll", "Lt", "Lm", "Lo", "Mn", "Mc", "Me", "Nd", "Nl", "No", "Zs", "Zl", "Zp", "Cc", "Cf", "Cs", "Co", "Pc", "Pd", "Ps", "Pe", "Pi", "Pf", "Po", "Sm", "Sc", "Sk", "So", "Cn", "Ws"}
            Dim RTN As Integer
            If Letter = "" Then Return ""
            If UnicodeWhiteSpace(Letter) <> "" Then
                Return "Unicode. Ws" 'whitespace
            End If
            RTN = Char.GetUnicodeCategory(Letter, 0)
            Return "Unicode." & Category(RTN)
        End Function


        Public Shared Function IsThisAUnicodeClass(A As String) As Boolean
            'todo need to change it so that unicode. is required instead of optioonal (Or better yet, but it in an option
            If Left(UCase(A), 8) <> "UNICODE." Then Return False
            A = FindRuleNameOnly(A)
            Select Case A
                    'todo It is now an error to only have the classname
                Case "Cc", "Cf", "Co", "Cs", "Ll", "Lm", "Lo", "Lt", "Lu", "Mc", "Me", "Mn", "Nd", "Nl", "No", "Pc", "Pd", "Pe", "Pf", "Pi", "Po", "Ps", "Sc", "Sk", "Sm", "So", "Zl", "Zp", "Zs"
                    Return False'todo remove this case condition, because it must have unicode.
                Case "Unicode.Cc", "Unicode.Cf", "Unicode.Co", "Unicode.Cs", "Unicode.Ll", "Unicode.Lm", "Unicode.Lo", "Unicode.Lt", "Unicode.Lu", "Unicode.Mc"
                    Return True
                Case "Unicode.Me", "Unicode.Mn", "Unicode.Nd", "Unicode.Nl", "Unicode.No", "Unicode.Pc", "Unicode.Pd", "Unicode.Pe", "Unicode.Pf", "Unicode.Pi"
                    Return True
                Case "Unicode.Po", "Unicode.Ps", "Unicode.Sc", "Unicode.Sk", "Unicode.Sm", "Unicode.So", "Unicode.Zl", "Unicode.Zp", "Unicode.Zs", "Unicode.Ws"
                    Return True
                Case Else
                    Select Case MsgBox("This does not belong to any Unicode class " & CT & A, MsgBoxStyle.Critical, "Unknown Unicode Class : ")
                        Case MsgBoxResult.Abort
                            Application.Exit()
                        Case MsgBoxResult.Cancel
                        Case MsgBoxResult.Ignore
                        Case MsgBoxResult.No
                        Case MsgBoxResult.Ok
                        Case MsgBoxResult.Retry
                        Case MsgBoxResult.Yes
                    End Select
            End Select
            Return False
        End Function

        Public Shared Function IsThisAHexPrefix_NotUsed(A As String) As Boolean
            If InStr("0X", A) <> 0 Then Return True
            If InStr("&H", A) <> 0 Then Return True
            Return False
        End Function
        Public Shared Function TestUnicodeRange_Used(Hex_0X As String, Letter As String) As Boolean
            Dim I, J, K, L As Integer
            Dim A, B, C, D As String
            A = Hex_0X
            A = Trim(TrimOff(A, "'"))
            A = Trim(TrimOff(A, "<"))
            A = Trim(TrimOff(A, ">"))
            A = Trim(TrimOff(A, ","))
            A = Trim(A)

            If InStr("0X", A) = 0 Then Return False
            If InStr("-", A) = 0 Then
                If InStr("0X", A) <> 0 Then
                    I = Hex2Bin(A)
                Else
                    I = AscW(A)
                End If
                If InStr("0X", Letter) <> 0 Then
                    J = Hex2Bin(Letter)
                Else
                    J = AscW(Letter)
                End If
                If I = J Then Return True
                Return False
            End If
            'else hex a a range of characters
            I = InStr("0X", A)
            J = InStr(I + 1, "0X", A)
            K = InStr("0X", Letter)
            L = InStr("-", A)
            B = Mid(A, 1, L - 1)
            C = Mid(A, L + 1)


            If I <> 0 Then
                I = Hex2Bin(B) 'i+2
            Else
                I = AscW(A)
            End If

            If J <> 0 Then
                J = Hex2Bin(C) 'j+2
            Else
                J = AscW(Right(A, 1))
            End If

            If Letter = "" Then Return False
            D = Letter
            If K <> 0 Then
                If Right(D, 1) = ";" Then D = Mid(D, 1, Len(D) - 1)
                K = Hex2Bin(TrimOff(Letter, "'"))
                'K = Hex2Bin(Mid(Letter, 3))
            Else
                K = AscW(Letter)
            End If

            If K >= I And K <= J Then
                Return True
            End If
            Return False
        End Function
        ' Ws white space

        'Lu 00 Upper case Letter 	"Lu" (letter, upper case). 0.
        'Ll 01 Lower case Letter 	"Ll" (letter, lower case). 1.
        'Lt 02 Title case Letter 	"Lt" (letter, title case). 2.
        'Lm 03 ModifierLetter 	Modifier letter character, which Is free-standing spacing character that indicates modifications of a preceding letter. "Lm" (letter, modifier). 3.
        'Lo 04 OtherLetter 	Not an upper Case letter, a lower Case letter, a title Case letter, Or a modifier letter. "Lo" (letter, other). 4.
        'Mn 05 NonSpacingMark 	5 	NonSpacing character that indicates modifications of a base character. "Mn" (mark, non spacing). 5.
        'Mc 06 SpacingCombiningMark 	6 	Spacing character that indicates modifications of a base character And affects the width of the glyph for that base character. "Mc" (mark, spacing combining). 6.
        'Me 07 EnclosingMark 	7 	Enclosing mark character, which Is a non spacing combining character that surrounds all previous characters up to And including a base character. "Me" (mark, enclosing). 7.
        'Nd 08 DecimalDigitNumber 	8 	Decimal digit character, that Is, a character representing an integer in the range 0 through 9. "Nd" (number, decimal digit). 8.
        'Nl 09 LetterNumber 	9 	Number represented by a letter, instead of a decimal digit, for example, the Roman numeral for five, which Is "V". The indicator Is "Nl" (number, letter). 9.
        'No 10 OtherNumber 	10 	Number that Is neither a decimal digit nor a letter number, for example, the fraction 1/2. The indicator Is "No" (number, other). 10.
        'Zs 11 SpaceSeparator 	11 	Space character, which has no glyph but Is Not a control Or format character. "Zs" (separator, space). 11.
        'Zl 12 LineSeparator 	12 	Character that Is used to separate Lines of text. "Zl" (separator, Line). 12.
        'Zp 13 ParagraphSeparator 	13 	Character used to separate paragraphs. "Zp" (separator, paragraph). 13.
        'Cc 14 Control 	14 	Control code character, with a Unicode value of U+007F Or in the range U+0000 through U+001F Or U+0080 through U+009F. "Cc" (other, control). 14.
        'Cf 15 Format 	15 	Format character that affects the layout of text Or the operation of text processes, but Is Not normally rendered. "Cf" (other, format). 15.
        'Cs 16 Surrogate 	16 	High surrogate Or a low surrogate character. Surrogate code values are in the range U+D800 through U+DFFF. "Cs" (other, surrogate). 16.
        'Co 17 PrivateUse 	17 	Private-use character, with a Unicode value in the range U+E000 through U+F8FF. "Co" (other, private use). 17.
        'Pc 18 ConnectorPunctuation 	18 	Connector punctuation character that connects two characters. "Pc" (punctuation, connector). 18.
        'Ps 19 DashPunctuation 	19 	Dash Or hyphen character. "Pd" (punctuation, dash). 19.
        'Pe 20 OpenPunctuation 	20 	Opening character of one of the paired punctuation marks, such as parentheses, square brackets, And braces. "Ps" (punctuation, open). 20.
        'Pi 21 ClosePunctuation 	21 		Closing character of one of the paired punctuation marks, such as parentheses, square brackets, And braces. "Pe" (punctuation, close). 21.
        'Pf 22 InitialQuotePunctuation 	22 	Opening Or initial quotation mark character. "Pi" (punctuation, initial quote). 22.
        'Po 23 FinalQuotePunctuation 	23 	Closing Or final quotation mark character. "Pf" (punctuation, final quote). 23.
        'Sm 24 OtherPunctuation 	24 	Punctuation character that Is Not a connector, a dash, open punctuation, close punctuation, an initial quote, Or a final quote. "Po" (punctuation, other). 24.
        'Sc 25 MathSymbol 	25 	Mathematical Symbol character, such as "+" Or "= ". "Sm" (Symbol, math). 25.
        'Sk 26 CurrencySymbol 	26 	Currency Symbol character. "Sc" (Symbol, currency). 26.
        'So 27 ModifierSymbol 	27 	Modifier Symbol character, which indicates modifications of surrounding characters. For example, the fraction slash indicates that the number to the left Is the numerator And the number to the right Is the denominator. The indicator Is signified by the   Unicode designation "Sk" (Symbol, modifier). 27.
        'Cn 28 OtherSymbol 	28 	Symbol character that Is Not a mathematical Symbol, a currency Symbol Or a modifier Symbol. "So" (Symbol, other). 28.
        'Ws 29 OtherNotAssigned 	29 	Character that Is Not assigned to any Unicode category. "Cn" (other, Not assigned). 29.
    End Class


    'Friend Class MyDrawing
    Friend Sub MyDrawAtom(ByRef SourceForm As Source, S As String)
        Dim A1, A2, A3, A4, A5, X As String
        Dim GrammarRuleName, Grammar2FlowchartRule As String
        A1 = "" : A2 = "" : A3 = "" : A4 = "" : A5 = "" : X = ""
        'Every flowchart command hase to have had a /grammar defined for it.
        GrammarRuleName = Parse(X, ",", "")
        X = DropParse(X, ",", "")
        Grammar2FlowchartRule = Parse(X, ",", "")
        X = DropParse(X, ",", "")
        While Len(X) > 0
            A4 = Parse(X, ",", "")
            X = DropParse(X, ",", "")
            If IsThisAColor(A4) <> "" Then
                A1 = A4
                A2 = Parse(X, ",", "")
                X = DropParse(X, ",", "")
                A3 = Parse(X, ",", "")
                X = DropParse(X, ",", "")
            Else
                A4 = A2
                A5 = A3
                A2 = Parse(X, ",", "")
                X = DropParse(X, ",", "")
                A3 = Parse(X, ",", "")
                X = DropParse(X, ",", "")
                Log("", 1080, Err6, A1 & "," & A2 & "," & A3 & ",  " & A4 & "," & A5)
            End If
        End While

    End Sub
    Friend Sub MyDraw(ByRef SourceForm As Source, SSS As SourceSyntaxStructure, FC As String)
        Dim Kounter As Integer = 0
        While Len(FC) > 0
            Debug.Print(Parse(FC, vbCrLf, ""))
            ConverLineNumber2FlowChartXY(SourceForm, SSS.Source_Line_Number)
            MyUniverse.X = (SSS.Source_Line_Number Mod MyUniverse.MaxRows) * MyUniverse.Spacing1
            MyUniverse.Y = (SSS.Source_Line_Number - (SSS.Source_Line_Number Mod MyUniverse.MaxRows)) * MyUniverse.Spacing1
            Debug.Print("MyDraw at " & Kounter.ToString & ", ( " & MyUniverse.X.ToString & ", " & MyUniverse.Y.ToString & ")")
            FCcmd(SourceForm, Parse(FC, vbCrLf, ""), SSS.Source_Line_Number)
            FC = DropParse(FC, vbCrLf)
            MyDrawLine(SourceForm,
                       10 + Kounter,
                       SSS.Source_Line_Number * 10,
                       20 + Kounter,
                       SSS.Source_Line_Number * 10 - 10, "Red")
            Kounter += 10
        End While
    End Sub

    Friend Sub MyDrawLine(ByRef SourceForm As Source, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, CLR As String)
        Dim XY1, XY2 As Point
        Dim MyColor, Style, Start, MyEnd As String
        Dim Red, Green, Blue, Alpha As Byte
        MyColor = BinarySearchList(Options.ListBoxColors, CLR) 'name, Red, Green, Blue, Alpha, line style, start cap, end cap
        Debug.Print("Color name " & Parse(MyColor, ",", ""))
        MyColor = DropParse(MyColor, ",", "")
        Red = CByte(Str2Bin(Parse(MyColor, ",", ""))) : MyColor = DropParse(MyColor, ",", "")
        Green = CByte(Str2Bin(Parse(MyColor, ",", ""))) : MyColor = DropParse(MyColor, ",", "")
        Blue = CByte(Str2Bin(Parse(MyColor, ",", ""))) : MyColor = DropParse(MyColor, ",", "")
        Alpha = CByte(Str2Bin(Parse(MyColor, ",", ""))) : MyColor = DropParse(MyColor, ",", "")
        Style = Parse(MyColor, ",", "") : MyColor = DropParse(MyColor, ",", "")
        Start = Parse(MyColor, ",", "") : MyColor = DropParse(MyColor, ",", "")
        MyEnd = Parse(MyColor, ",", "") : MyColor = DropParse(MyColor, ",", "")

        'Dim MyBrush As Brush = New System.Drawing.SolidBrush(Color.FromArgb(Alpha, Red, Green, Blue))
        Dim MyPen As New Pen(Color.FromArgb(Alpha, Red, Green, Blue))
        MyPen.Width = 5
        MyPen.Brush = New System.Drawing.SolidBrush(Color.FromArgb(Alpha, Red, Green, Blue))
        XY1.Y = y1 + MyUniverse.Y
        XY1.X = x1 + MyUniverse.X
        XY2.Y = y2 + MyUniverse.Y
        XY2.X = x2 + MyUniverse.X
        Debug.Print(XY1.X.ToString & ", " & XY1.Y.ToString & ", " & XY2.X.ToString & ", " & XY2.Y.ToString & ", " & CLR)
        'todo fix this it does not work to draw with the right colors
        'SourceForm.FlowChartPictureBox.CreateGraphics.DrawLine(MyPen, XY1, XY2)
        'MyPen.Brush.Dispose()
        'MyPen.Brush = New System.Drawing.SolidBrush(Color.FromArgb(Alpha, Blue, Red, Green)) 'crazy color 
        'SourceForm.FlowChartPictureBox.CreateGraphics.DrawLine(MyPen, XY1, XY2)
        'MyPen.Brush.Dispose()
        MyPen.Brush = Brushes.DarkGreen
        MyPen.Width = 5 'todo the line width should come from the datatype of the path (lines(symbols) should always be 1)
        'draws green line width of 5
        SourceForm.FlowChartPictureBox.CreateGraphics.DrawLine(MyPen, XY1, XY2)
        MyPen.Brush.Dispose()
        'draw black line width of 1
        SourceForm.FlowChartPictureBox.CreateGraphics.DrawLine(Pens.Black, XY1, XY2)
        MyPen.Dispose()
        'MyBrush.Dispose()
    End Sub


    Friend Sub MyDrawText(ByRef SourceForm As Source, S As String, x As Integer, Y As Integer, CLR As String)
        Dim XY1 As Point
        Dim MyColor, Style, Start, MyEnd As String
        Dim Red, Green, Blue, Alpha As Byte
        MyColor = BinarySearchList(Options.ListBoxColors, CLR) 'name, Red, Green, Blue, Alpha, line style, start cap, end cap
        MyColor = DropParse(MyColor, ",", "")
        Red = CByte(Str2Bin(Parse(MyColor, ",", ""))) : MyColor = DropParse(MyColor, ",", "")
        Green = CByte(Str2Bin(Parse(MyColor, ",", ""))) : MyColor = DropParse(MyColor, ",", "")
        Blue = CByte(Str2Bin(Parse(MyColor, ",", ""))) : MyColor = DropParse(MyColor, ",", "")
        Alpha = CByte(Str2Bin(Parse(MyColor, ",", ""))) : MyColor = DropParse(MyColor, ",", "")
        Style = Parse(MyColor, ",", "") : MyColor = DropParse(MyColor, ",", "")
        Start = Parse(MyColor, ",", "") : MyColor = DropParse(MyColor, ",", "")
        MyEnd = Parse(MyColor, ",", "") : MyColor = DropParse(MyColor, ",", "")


        Dim MyBrush As Brush = New System.Drawing.SolidBrush(Color.FromArgb(Alpha, Red, Green, Blue))
        MyBrush = New System.Drawing.SolidBrush(Color.FromArgb(Alpha, Red, Green, Blue))
        XY1.Y = MyUniverse.Y + Y
        XY1.X = MyUniverse.X + x
        SourceForm.FlowChartPictureBox.CreateGraphics.DrawString(S, System.Drawing.SystemFonts.DefaultFont, MyBrush, XY1)
        SourceForm.FlowChartPictureBox.CreateGraphics.DrawString(S, System.Drawing.SystemFonts.CaptionFont, Brushes.Black, XY1)
        MyBrush.Dispose()
    End Sub




    Friend Sub DisplayStatus(A As Integer, B As String,
                             Optional C As String = "",
                             Optional D As String = "",
                             Optional E As String = "",
                             Optional F As String = "",
                             Optional G As String = "")
        Log("", 1082, Err1, A.ToString & ") " & B & C & D & E & F & G)
    End Sub





    'todo this needs to be deleted (changes from the old arrays to listbox ....
    Friend Function MySizeOf(Strs() As String) As Integer
        Return UBound(Strs)
    End Function

    Friend Function MySizeOf(Ints() As Integer) As Integer
        Return UBound(Ints)
    End Function


    Friend Function FlowChart_tablePathSymbolName(ByRef SourceForm As Source, IndexFlowChart As Integer) As String  ' The name of the /use, the variable name of /Path & /Constant
        Return GetMyField(SourceForm.ListBoxFlowChart, "FlowChart:SymbolName", IndexFlowChart)
    End Function
    Friend Sub FlowChart_tablePathSymbolName(ByRef SourceForm As Source, IndexFlowChart As Integer, Value As String) ' The name of the /use, the variable name of /Path & /Constant
    End Sub


    Friend Function FlowChart_tableCode(ByRef SourceForm As Source, IndexFlowChart As Integer) As Integer  ' The codes /Use, /Path, /Constant
        Return CInt(Val(GetMyField(SourceForm.ListBoxFlowChart, "FlowChart:Code", IndexFlowChart)))
    End Function



    Friend Function FlowChart_FileX1(ByRef SourceForm As Source, IndexFlowChart As Integer) As Integer
        Return CInt(Val(GetMyField(SourceForm.ListBoxFlowChart, "FlowChart:X1", IndexFlowChart)))
    End Function
    Friend Sub FlowChart_FileX1(ByRef SourceForm As Source, IndexFlowChart As Integer, Value As Integer)
    End Sub



    Friend Function FlowChart_FileY1(ByRef SourceForm As Source, IndexFlowChart As Integer) As Integer
        Return CInt(Val(GetMyField(SourceForm.ListBoxFlowChart, "FlowChart:Y1", IndexFlowChart)))
    End Function
    Friend Sub FlowChart_FileY1(ByRef SourceForm As Source, IndexFlowChart As Integer, Value As Integer)
    End Sub

    Friend Sub FlowChart_FileX2(ByRef SourceForm As Source, IndexFlowChart As Integer, Value As Integer)
    End Sub
    Friend Function FlowChart_FileX2(ByRef SourceForm As Source, IndexFlowChart As Integer) As Integer
        Return CInt(Val(GetMyField(SourceForm.ListBoxFlowChart, "FlowChart:X2", IndexFlowChart)))
    End Function


    Friend Function FlowChart_FileY2(ByRef SourceForm As Source, IndexFlowChart As Integer) As Integer   'Y2 for path, future options for /use
        Return CInt(Val(GetMyField(SourceForm.ListBoxFlowChart, "FlowChart:Y2", IndexFlowChart)))
    End Function
    Friend Sub FlowChart_FileY2(ByRef SourceForm As Source, IndexFlowChart As Integer, Value As Integer)
    End Sub


    'todo WHY is this a string for the data type instead of a pointer to the datatype array? (or int to datatype's) incase the datatype list changes
    Friend Function FlowChart_File_DataType(ByRef SourceForm As Source, IndexFlowChart As Integer) As String  'The DataType for /Path /Constant
        Return GetMyField(SourceForm.ListBoxFlowChart, "FlowChart:DataType", IndexFlowChart)
    End Function
    Friend Sub FlowChart_File_DataType(ByRef SourceForm As Source, IndexFlowChart As Integer, Value As String) 'The DataType for /Path /Constant
    End Sub


    'clear out listbox and fill with what ever it is I am looking for (Unless it already has the right table:field)
    'List box should have Value, pointer to flowchartlistbox
    'note when building the listbox, if numbers then it shoule have spaced(??, value) for ?? digits long
    Friend Sub ISAMOpen(ByRef SourecForm As Source, MyTableField As String)
        MsgBox("Write this")
    End Sub
    Friend Function ISAMFindFirst(ByRef SourecForm As Source, Index As Integer) As Integer
        MsgBox("Write this")
        Return 1
    End Function
    Friend Function ISAMNext(ByRef SourecForm As Source, Index As Integer) As Integer
        MsgBox("Write this")
        Return Index + 1
    End Function







    Friend Sub Named_FileSymbolName_ISAM(Index As Integer, Value As String)    'sorted Indexes to MyArrays
        SetMyRecord(Options.ListBoxSymbols, Index, Value)
    End Sub
    Friend Sub Named_FileSyntax_ISAM(Index As Integer, Value As String)      ' Only used during Decompile and reset to length of one afterwards
        SetMyRecord(Options.ListBoxSymbols, Index, Value)
    End Sub
    Friend Sub Named_FileSymbolName(Index As Integer, Value As String)   'Name of the Symbol
        SetMyRecord(Options.ListBoxSymbols, Index, Value)
    End Sub
    Friend Sub Named_FileSyntax(Index As Integer, Value As String) ' The Syntax for the deCompiler made from the program test 2020/6/22
        SetMyRecord(Options.ListBoxSymbols, Index, Value)
    End Sub
    '    friend   sub   Named_FileSymbolIndexes(Index as integer, Value as string) ' holes the index of the start (/name) of the symbol graphics
    Friend Sub Named_FileMicroCodeText(Index As Integer, Value As String)   'The actual program MicroCodeText to be 'fixed'
        SetMyRecord(Options.ListBoxSymbols, Index, Value)
    End Sub
    Friend Sub Named_FileOpCode(Index As Integer, Value As String) 'The Machine code of this assemble Symbol
        SetMyRecord(Options.ListBoxSymbols, Index, Value)
    End Sub
    Friend Sub Named_FileNotes(Index As Integer, Value As String)  'Notes for this Symbol
        SetMyRecord(Options.ListBoxSymbols, Index, Value)
    End Sub
    Friend Sub Named_FileDescription(Index As Integer, Value As String)  'Symbol Description
        SetMyRecord(Options.ListBoxSymbols, Index, Value)
    End Sub
    Friend Sub Named_FileNameOfFile(Index As Integer, Value As String)   'The device:/path/Filename where this came from 
        SetMyRecord(Options.ListBoxSymbols, Index, Value)
    End Sub
    Friend Sub Named_FileAuthor(Index As Integer, Value As String) 'Who wrote or responsible for this Symbol
        SetMyRecord(Options.ListBoxSymbols, Index, Value)
    End Sub
    Friend Sub Named_FileVersion(Index As Integer, Value As String) ' the date of the latest update
        SetMyRecord(Options.ListBoxSymbols, Index, Value)
    End Sub
    Friend Sub Named_FileStroke(Index As Integer, Value As String) 'The movement of the mouse that id's this Symbol
        SetMyRecord(Options.ListBoxSymbols, Index, Value)
    End Sub
    Friend Sub Symbol_FileSymbolName(Index As Integer, Value As String) 'The name of this Symbol for /Name code
        SetMyRecord(Options.ListBoxSymbolKeyWordGraphics, Index, Value)
    End Sub
    Friend Sub Symbol_File_NameOfPoint(Index As Integer, Value As String) 'name of points and color of Lines
        SetMyRecord(Options.ListBoxSymbolKeyWordGraphics, Index, Value)
    End Sub
    Friend Sub Symbol_FileCoded(Index As Integer, Value As String)  'The code /Line /point etc 
        SetMyRecord(Options.ListBoxSymbolKeyWordGraphics, Index, Value)
    End Sub
    Friend Sub Symbol_FileX1(Index As Integer, Value As String)
        SetMyRecord(Options.ListBoxSymbolKeyWordGraphics, Index, Value)
    End Sub
    Friend Sub Symbol_FileY1(Index As Integer, Value As String)
        SetMyRecord(Options.ListBoxSymbolKeyWordGraphics, Index, Value)
    End Sub
    Friend Sub Symbol_FileX2_io(Index As Integer, Value As String)   'Used also as Enum Input-Output/bot/optional IO  ... 
        SetMyRecord(Options.ListBoxSymbolKeyWordGraphics, Index, Value)
    End Sub
    Friend Sub Symbol_FileY2_dt(Index As Integer, Value As String)  ' Also used as the index to the data type
        SetMyRecord(Options.ListBoxSymbolKeyWordGraphics, Index, Value)
    End Sub
    Friend Sub Net_FilePathNames(ByRef SourceForm As Source, Index As Integer, Value As String)  ' This hold the name of the paths
        SetMyRecord(SourceForm.ListBoxVariables, Index, Value)
    End Sub
    Friend Sub Net_FileLinks(ByRef SourceForm As Source, Index As Integer, Value As String)  ' This holds all of the link numbers that are connected together.
        SetMyRecord(SourceForm.ListBoxVariables, Index, Value)
    End Sub
    Friend Sub Net_FileDataType(ByRef SourceForm As Source, Index As Integer, Value As String)  ' This holds datatype(s) defined for this pathname  that are connected together.
        SetMyRecord(SourceForm.ListBoxVariables, Index, Value)
    End Sub
    Friend Sub DataType_FileName(Index As Integer, Value As String)            'Name of the DataType
        SetMyRecord(Options.ListBoxDataTypes, Index, Value)
    End Sub
    Friend Sub DataType_FileDescription(Index As Integer, Value As String)
        SetMyRecord(Options.ListBoxDataTypes, Index, Value)
    End Sub
    Friend Sub DataType_FileNumberOfBytes(Index As Integer, Value As String)  ' size in bytes of the data
        SetMyRecord(Options.ListBoxDataTypes, Index, Value)
    End Sub
    Friend Sub DataType_FileColorIndex(Index As Integer, Value As String)      'number of the color in color_file ... to use
        SetMyRecord(Options.ListBoxDataTypes, Index, Value)
    End Sub
    Friend Sub DataType_FileWidth(Index As Integer, Value As String)          'Width of the /Path and diameter of the /Points
        SetMyRecord(Options.ListBoxDataTypes, Index, Value)
    End Sub






    Friend Function Named_FileSymbolName_ISAM(Index As Integer) As Integer    'sorted Indexes to MyArrays
        Return CInt(Val(GetMyRecord(Options.ListBoxSymbols, Index)))
    End Function
    Friend Function Named_FileSyntax_ISAM(Index As Integer) As Integer      ' Only used during Decompile and reset to length of one afterwards
        Return CInt(Val(GetMyRecord(Options.ListBoxSymbols, Index)))
    End Function
    Friend Function Named_FileSymbolName(Index As Integer) As String   'Name of the Symbol
        Return GetMyRecord(Options.ListBoxSymbols, Index)
    End Function
    Friend Function Named_FileSyntax(Index As Integer) As String ' The Syntax for the deCompiler made from the program test 2020/6/22
        Return GetMyRecord(Options.ListBoxSymbols, Index)
    End Function
    '    friend   function   Named_FileSymbolIndexes(Index as integer) As Integer ' holes the index of the start (/name) of the symbol graphics
    Friend Function Named_FileMicroCodeText(Index As Integer) As String   'The actual program MicroCodeText to be 'fixed'
        Return GetMyRecord(Options.ListBoxSymbols, Index)
    End Function
    Friend Function Named_FileOpCode(Index As Integer) As String 'The Machine code of this assemble Symbol
        Return GetMyRecord(Options.ListBoxSymbols, Index)
    End Function
    Friend Function Named_FileNotes(Index As Integer) As String  'Notes for this Symbol
        Return GetMyRecord(Options.ListBoxSymbols, Index)
    End Function
    Friend Function Named_FileDescription(Index As Integer) As String  'Symbol Description
        Return GetMyRecord(Options.ListBoxSymbols, Index)
    End Function
    Friend Function Named_FileNameOfFile(Index As Integer) As String   'The device:/path/Filename where this came from 
        Return GetMyRecord(Options.ListBoxSymbols, Index)
    End Function
    Friend Function Named_FileAuthor(Index As Integer) As String 'Who wrote or responsible for this Symbol
        Return GetMyRecord(Options.ListBoxSymbols, Index)
    End Function
    Friend Function Named_FileVersion(Index As Integer) As String ' the date of the latest update
        Return GetMyRecord(Options.ListBoxSymbols, Index)
    End Function
    Friend Function Named_FileStroke(Index As Integer) As String 'The movement of the mouse that id's this Symbol
        Return GetMyRecord(Options.ListBoxSymbols, Index)
    End Function
    Friend Function Symbol_FileSymbolName(Index As Integer) As String 'The name of this Symbol for /Name code
        Return GetMyRecord(Options.ListBoxSymbolKeyWordGraphics, Index)
    End Function
    Friend Function Symbol_File_NameOfPoint(Index As Integer) As String 'name of points and color of Lines
        Return GetMyRecord(Options.ListBoxSymbolKeyWordGraphics, Index)
    End Function
    Friend Function Symbol_FileCoded(Index As Integer) As Integer  'The code /Line /point etc 
        Return CInt(Val(GetMyRecord(Options.ListBoxSymbolKeyWordGraphics, Index)))
    End Function
    Friend Function Symbol_FileX1(Index As Integer) As Integer
        Return CInt(Val(GetMyRecord(Options.ListBoxSymbolKeyWordGraphics, Index)))
    End Function
    Friend Function Symbol_FileY1(Index As Integer) As Integer
        Return CInt(Val(GetMyRecord(Options.ListBoxSymbolKeyWordGraphics, Index)))
    End Function
    Friend Function Symbol_FileX2_io(Index As Integer) As String  'Used also as Enum Input-Output/bot/optional IO  ... 
        Return GetMyRecord(Options.ListBoxSymbolKeyWordGraphics, Index)
    End Function
    Friend Function Symbol_FileY2_dt(Index As Integer) As String ' Also used as the index to the data type
        Return GetMyRecord(Options.ListBoxSymbolKeyWordGraphics, Index)
    End Function
    Friend Function Net_FilePathNames(ByRef SourceForm As Source, Index As Integer) As String ' This hold the name of the paths
        Return GetMyRecord(SourceForm.ListBoxVariables, Index)
    End Function
    Friend Function Net_FileLinks(ByRef SourceForm As Source, Index As Integer) As String ' This holds all of the link numbers that are connected together.
        Return GetMyRecord(SourceForm.ListBoxVariables, Index)
    End Function
    Friend Function Net_FileDataType(ByRef SourceForm As Source, Index As Integer) As String ' This holds datatype(s) defined for this pathname  that are connected together.
        Return GetMyRecord(SourceForm.ListBoxVariables, Index)
    End Function
    Friend Function DataType_FileName(Index As Integer) As String            'Name of the DataType
        Return GetMyRecord(Options.ListBoxDataTypes, Index)
    End Function
    Friend Function DataType_FileDescription(Index As Integer) As String
        Return GetMyRecord(Options.ListBoxDataTypes, Index)
    End Function
    Friend Function DataType_FileNumberOfBytes(Index As Integer) As Integer  ' size in bytes of the data
        Return CInt(Val(GetMyRecord(Options.ListBoxDataTypes, Index)))
    End Function
    Friend Function DataType_FileColorIndex(Index As Integer) As String      'number of the color in color_file ... to use
        Return GetMyRecord(Options.ListBoxDataTypes, Index)
    End Function
    Friend Function DataType_FileWidth(Index As Integer) As Integer          'Width of the /Path and diameter of the /Points
        Return CInt(Val(GetMyRecord(Options.ListBoxDataTypes, Index)))
    End Function


    'Imports System.Diagnostics

    Friend Function CPUID() As String
        Dim myuuid As Guid = Guid.NewGuid()
        Dim myuuidAsString As String = myuuid.ToString()
        Return myuuidAsString
    End Function


End Module



