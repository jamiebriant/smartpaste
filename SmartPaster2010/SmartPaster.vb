Imports Microsoft.VisualStudio.CommandBars
Imports System.Windows.Forms
Imports Extensibility
Imports System.Runtime.InteropServices
Imports EnvDTE
Imports EnvDTE80

''' -----------------------------------------------------------------------------
''' Project	 : SmartPaster
''' Class	 : SmartPaster
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     Class responsible for doing the pasting/manipulating of clipdata.
''' </summary>
''' <remarks>
''' </remarks>
''' -----------------------------------------------------------------------------
Friend NotInheritable Class SmartPaster
    Public Sub New()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''    Convient property to retrieve the clipboard text from the clipboard
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private ReadOnly Property ClipboardText() As String
        Get
            ClipboardText = String.Empty
            Dim iData As IDataObject = Clipboard.GetDataObject()
            If iData.GetDataPresent(DataFormats.Text) Then Return CType(iData.GetData(DataFormats.Text), String)
        End Get
    End Property

#Region "Stringinize"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Stringinizes text passed to it for use in C#
    ''' </summary>
    ''' <param name="txt">Text to be Stringinized</param>
    ''' <returns>C# Stringinized text</returns>
    ''' <seealso cref='Stringinize' />
    ''' <seealso cref='StringinizeInVB' />
    ''' -----------------------------------------------------------------------------
    Private Function StringinizeInCS(ByVal txt As String) As String

        'c# quote character -- really just a "
        Dim qChr As String = """"
        Dim csNewLineEsc As String = Configuration.CSNewLine
        Dim csTabEsc As String = Configuration.CSTab

        'sb to work with
        Dim sb As New Text.StringBuilder(txt)

        'escape appropriately
        If Configuration.CSVerbatimLiterals Then
            'escape the quotes with ""
            sb.Replace(qChr, qChr & qChr)
        Else
            'escape \ and "
            sb.Replace("\", "\\")
            sb.Replace(qChr, "\" & qChr)
        End If

        'if we're allowed to span ...
        If Configuration.CSVerbatimLiteralsSpanLines Then
            'insert " at beginning and end
            sb.Insert(0, "@" & qChr) : sb.Append(qChr)
            Return sb.ToString
        End If



        '''''''''''''
        'process the passed string (txt), one line at a time

        'dump the stringbuilder into a temp string
        Dim tmp As String = sb.ToString : sb.Remove(0, sb.Length)

        'pointers that will isolate an individual line in the string
        Dim L, R As Integer
        Do
            'R should be where a linebreak occurs after L, or the end of the string
            R = tmp.IndexOf(Environment.NewLine, L)
            If R = -1 Then R = tmp.Length

            'now we'll have an individual line to work with, excluding the linebreak
            Dim ln As String = tmp.Substring(L, R - L)

            'if wordwrapping is enabled
            If Configuration.EnableWordWrap Then
                'wrap it
                ln = WordWrap(ln, Configuration.WrapAt)

                'and replace the linebreaks
                If Configuration.CSVerbatimLiterals Then
                    ln = ln.Replace(Environment.NewLine, qChr & " +" & Environment.NewLine & "@" & qChr)
                Else
                    ln = ln.Replace(Environment.NewLine, qChr & " +" & Environment.NewLine & qChr)
                End If

            End If

            'if R is not at the end, then a linebreak must follow
            If (R < tmp.Length - 1) Then
                'append lb appropriately
                If Configuration.CSEscapeLinebreaks Then
                    If Configuration.CSVerbatimLiterals Then
                        'if the newline string starts with a \, it must be an escape seq
                        If csNewLineEsc.StartsWith("\") Then
                            ln = ln & qChr & " + " & qChr & csNewLineEsc & qChr & " +" & Environment.NewLine & "@" & qChr
                        Else
                            ln = ln & qChr & " + " & csNewLineEsc & " +" & Environment.NewLine & "@" & qChr
                        End If

                    ElseIf (Not Configuration.CSVerbatimLiterals) Then
                        'if the newline string starts with a \, it must be an escape seq
                        If csNewLineEsc.StartsWith("\") Then
                            ln = ln & csNewLineEsc & qChr & " +" & Environment.NewLine & qChr
                        Else
                            ln = ln & qChr & " + " & csNewLineEsc & " +" & Environment.NewLine & qChr
                        End If
                    End If 'Configuration.CSVerbatimLiterals

                ElseIf (Not Configuration.CSEscapeLinebreaks) Then
                    'don't escape, just enquote
                    If Configuration.CSVerbatimLiterals Then
                        ln = ln & qChr & " +" & Environment.NewLine & "@" & qChr
                    Else
                        ln = ln & qChr & " +" & Environment.NewLine & qChr
                    End If

                End If
            End If

            'append line to stringbuilder
            sb.Append(ln)

            'move L accordingly
            L = R + Environment.NewLine.Length

        Loop While (R < tmp.Length) And (L < tmp.Length)

        'escape the tabs
        If Configuration.CSEscapeTabs Then
            If Configuration.CSVerbatimLiterals Then
                'if tabesc sstarts with a \, must be escape
                If csTabEsc.StartsWith("\") Then
                    sb.Replace(ControlChars.Tab, qChr & " + " & qChr & csTabEsc & qChr & " + " & "@" & qChr)
                Else
                    sb.Replace(ControlChars.Tab, qChr & " + " & csTabEsc & " + " & "@" & qChr)
                End If


            ElseIf (Not Configuration.CSVerbatimLiterals) Then
                'if tabesc sstarts with a \, must be escape
                If csTabEsc.StartsWith("\") Then
                    sb.Replace(ControlChars.Tab, csTabEsc)
                Else
                    sb.Replace(ControlChars.Tab, qChr & " + " & csTabEsc & " + " & qChr)
                End If

            End If 'Configuration.CSVerbatimLiterals

        End If '(Configuration.CSEscapeTabs )

        'get rid of lines that begin with '"" + '
        If Configuration.CSVerbatimLiterals Then
            sb.Replace(Environment.NewLine & "@" & qChr & qChr & " + ", Environment.NewLine)
        Else
            sb.Replace(Environment.NewLine & qChr & qChr & " + ", Environment.NewLine)
        End If


        'insert a quote at the beginning and end of string
        If Configuration.CSVerbatimLiterals Then
            sb.Insert(0, "@" & qChr)
        Else
            sb.Insert(0, qChr)
        End If
        sb.Append(qChr)

        'and return
        Return sb.ToString
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Stringinizes text passed to it for use in VB
    ''' </summary>
    ''' <param name="txt">Text to be Stringinized</param>
    ''' <returns>VB Stringinized text</returns>
    ''' <seealso cref='Stringinize' />
    ''' <seealso cref='StringinizeInVB' />
    ''' -----------------------------------------------------------------------------
    Private Function StringinizeInVB(ByVal txt As String) As String
        'c# quote character -- really just a "
        Dim qChr As String = """"
        Dim vbNewLineEsc As String = Configuration.VBNewLine
        Dim vbTabEsc As String = Configuration.VBTab

        'sb to work with
        Dim sb As New Text.StringBuilder(txt)

        'escape quotes
        sb.Replace(qChr, qChr & qChr)

        '''''''''''''
        'process the passed string (txt), one line at a time

        'dump the stringbuilder into a temp string
        Dim tmp As String = sb.ToString : sb.Remove(0, sb.Length)

        'pointers that will isolate an individual line in the string
        Dim L, R As Integer
        Do
            'R should be where a linebreak occurs after L, or the end of the string
            R = tmp.IndexOf(Environment.NewLine, L)
            If R = -1 Then R = tmp.Length

            'now we'll have an individual line to work with, excluding the linebreak
            Dim ln As String = tmp.Substring(L, R - L)

            'if wordwrapping is enabled
            If Configuration.EnableWordWrap Then
                'wrap it
                ln = WordWrap(ln, Configuration.WrapAt)

                'and replace the linebreaks
                ln = ln.Replace(Environment.NewLine, qChr & " & _" & Environment.NewLine & qChr)
            End If

            'if R is not at the end, then a linebreak must follow
            If (R < tmp.Length - 1) Then
                'append lb appropriately
                If Configuration.VBEscapeLinebreaks Then
                    ln = ln & qChr & " & " & vbNewLineEsc & " & _" & Environment.NewLine & qChr
                Else
                    'don't escape, just enquote
                    ln = ln & qChr & " & _" & Environment.NewLine & qChr
                End If
            End If

            'append line to stringbuilder
            sb.Append(ln)

            'move L accordingly
            L = R + Environment.NewLine.Length

        Loop While (R < tmp.Length) And (L < tmp.Length)

        'escape the tabs
        If Configuration.VBEscapeTabs Then
            'if tabesc sstarts with a \, must be escape
            sb.Replace(ControlChars.Tab, qChr & " & " & vbTabEsc & " & " & qChr)
        End If

        'get rid of lines that begin with '"" + '
        sb.Replace(Environment.NewLine & qChr & qChr & " & ", Environment.NewLine)

        'insert a quote at the beginning and end of string
        sb.Insert(0, qChr)
        sb.Append(qChr)

        'and return
        Return sb.ToString
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Stringinizes text using the default language set in configuration.
    ''' </summary>
    ''' <param name="txt">Text to be Stringinized</param>
    ''' <returns>Stringinized text</returns>
    ''' <seealso cref='StringinizeInCS' />
    ''' <seealso cref='StringinizeInVB' />
    ''' <seealso class='Configuration' />
    ''' -----------------------------------------------------------------------------
    Private Function Stringinize(ByVal txt As String) As String
        If Configuration.DefaultLanguage = "Visual Basic" Then
            Return StringinizeInVB(txt)
        Else
            Return StringinizeInCS(txt)
        End If
    End Function
#End Region

#Region "Commentize"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Commentizes text passed to it for use in VB
    ''' </summary>
    ''' <param name="txt">Text to be Stringinized</param>
    ''' <returns>VBCommentized text</returns>
    ''' <seealso cref='Commentize' />
    ''' <seealso cref='CommentizeInCS' />
    ''' -----------------------------------------------------------------------------
    Private Function CommentizeInVB(ByVal txt As String) As String

        Dim cmtChar As String = "'"

        Dim sb As New Text.StringBuilder(txt.Length)

        '''''''''''''
        'process the passed string (txt), one line at a time

        'pointers that will isolate an individual line in the string
        Dim L, R As Integer
        Do
            'R should be where a linebreak occurs after L, or the end of the string
            R = txt.IndexOf(Environment.NewLine, L)
            If R = -1 Then R = txt.Length

            'now we'll have an individual line to work with, excluding the linebreak
            Dim ln As String = txt.Substring(L, R - L)

            'if wordwrapping is enabled
            If Configuration.EnableWordWrap Then
                'wrap it
                ln = WordWrap(ln, Configuration.WrapAt)

                'and replace the linebreaks
                ln = ln.Replace(Environment.NewLine, Environment.NewLine & cmtChar)
            End If

            'append line to stringbuilder
            sb.Append(cmtChar & ln & Environment.NewLine)

            'move L accordingly
            L = R + Environment.NewLine.Length

        Loop While (R < txt.Length) And (L < txt.Length)

        Return sb.ToString
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Commentizes text passed to it for use in C#
    ''' </summary>
    ''' <param name="txt">Text to be Stringinized</param>
    ''' <returns>C# Commentized text</returns>
    ''' <seealso cref='Commentize' />
    ''' <seealso cref='CommentizeInVB' />
    ''' -----------------------------------------------------------------------------
    Private Function CommentizeInCS(ByVal txt As String) As String
        Dim cmtChar As String = "//"

        Dim sb As New Text.StringBuilder(txt.Length)

        '''''''''''''
        'process the passed string (txt), one line at a time

        'pointers that will isolate an individual line in the string
        Dim L, R As Integer
        Do
            'R should be where a linebreak occurs after L, or the end of the string
            R = txt.IndexOf(Environment.NewLine, L)
            If R = -1 Then R = txt.Length

            'now we'll have an individual line to work with, excluding the linebreak
            Dim ln As String = txt.Substring(L, R - L)

            'if wordwrapping is enabled
            If Configuration.EnableWordWrap Then
                'wrap it
                ln = WordWrap(ln, Configuration.WrapAt)

                'and replace the linebreaks
                ln = ln.Replace(Environment.NewLine, Environment.NewLine & cmtChar)
            End If

            'append line to stringbuilder
            sb.Append(cmtChar & ln & Environment.NewLine)

            'move L accordingly
            L = R + Environment.NewLine.Length

        Loop While (R < txt.Length) And (L < txt.Length)

        Return sb.ToString
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Commentizes text using the default language set in configuration.
    ''' </summary>
    ''' <param name="txt">Text to be Stringinized</param>
    ''' <returns>Commentized text</returns>
    ''' <seealso cref='CommentizeInVB' />
    ''' <seealso cref='CommentizeInCS' />
    ''' <seealso class='Configuration' />
    ''' -----------------------------------------------------------------------------
    Private Function Commentize(ByVal txt As String) As String
        If Configuration.DefaultLanguage = "Visual Basic" Then
            Return CommentizeInVB(txt)
        Else
            Return CommentizeInCS(txt)
        End If

    End Function
#End Region

#Region "Stringbuilderize"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Stringinizes text passed to it for use in C#
    ''' </summary>
    ''' <param name="txt">Text to be Stringinized</param>
    ''' <returns>C# Stringinized text</returns>
    ''' <seealso cref='Stringinize' />
    ''' <seealso cref='StringinizeInVB' />
    ''' -----------------------------------------------------------------------------
    Private Function StringbuilderizeInCS(ByVal txt As String, ByVal sbName As String) As String

        'c# quote character -- really just a "
        Dim qChr As String = """"
        Dim csNewLineEsc As String = Configuration.CSNewLine
        Dim csTabEsc As String = Configuration.CSTab

        'sb to work with
        Dim sb As New Text.StringBuilder(txt)

        'escape \,", and {}
        If Configuration.CSVerbatimLiterals Then
            sb.Replace(qChr, qChr & qChr)
        Else
            sb.Replace("\", "\\") : sb.Replace(qChr, "\" & qChr)
        End If

        If Configuration.AppendFormatOnStringBuilder Then
            sb.Replace("{", "{{") : sb.Replace("}", "}}")
        End If

        '''''''''''''
        'process the passed string (txt), one line at a time

        'dump the stringbuilder into a temp string
        Dim fullString As String = sb.ToString : sb.Remove(0, sb.Length)

        'pointers that will isolate an individual line in the string
        Dim leftIdx As Integer = 0, rightIdx As Integer = 0
        Do
            'R should be where a linebreak occurs after L, or the end of the string
            rightIdx = fullString.IndexOf(Environment.NewLine, leftIdx)
            If rightIdx = -1 Then rightIdx = fullString.Length

            'now we'll have an individual line to work with, excluding the linebreak
            Dim fullLine As String = fullString.Substring(leftIdx, rightIdx - leftIdx)

            'if wordwrapping is enabled then word wrap it
            If Configuration.EnableWordWrap Then fullLine = WordWrap(fullLine, Configuration.WrapAt)

            'since ln may contain multiple lines, it's necessary to handle each line
            'at once - this loop is very similar to the outerloop
            Dim innerLeftIdx As Integer = 0, innerRightIdx As Integer = 0
            Do
                'same concept as above to extract the inner string
                innerRightIdx = fullLine.IndexOf(Environment.NewLine, innerLeftIdx)
                If innerRightIdx = -1 Then innerRightIdx = fullLine.Length
                Dim innerLine As String = fullLine.Substring(innerLeftIdx, innerRightIdx - innerLeftIdx)


                'these will help us in the output portion
                Dim hasBr, hasTab As Boolean

                'if the R pointer (for the original string) is not at the end, but
                'the R2 Pointer (for the original string defined linebreaks) is
                ' then we should add a BR
                hasBr = _
                    ((rightIdx < fullString.Length) And (innerRightIdx = fullLine.Length)) _
                    And Configuration.CSEscapeLinebreaks

                'this is mostly important for using AppendFormat 
                hasTab = _
                    (innerLine.IndexOf(ControlChars.Tab) > -1) _
                    And Configuration.CSEscapeTabs

                'now build up the line we'll append
                If Configuration.AppendFormatOnStringBuilder Then

                    If (Configuration.CSVerbatimLiterals) Then

                        sb.Append(sbName & ".AppendFormat(@" & qChr)  'mySb.AppendFormat(@"

                        If hasTab Then
                            sb.Append(innerLine.Replace(ControlChars.Tab, "{0}")) '{0}myPastedString
                        Else
                            sb.Append(innerLine) 'myPastedString
                        End If


                        If hasBr And hasTab Then
                            sb.Append("{1}" & qChr & ", " & csTabEsc & ", " & csNewLineEsc)
                        ElseIf hasBr And (Not hasTab) Then
                            sb.Append("{0}" & qChr & ", " & csNewLineEsc)
                        ElseIf (Not hasBr) And hasTab Then
                            sb.Append(qChr & ", " & csNewLineEsc)
                        ElseIf (Not hasBr) And (Not hasTab) Then
                            sb.Append(qChr)
                        End If

                        sb.Append(");" & Environment.NewLine) ');<br>

                    ElseIf (Not Configuration.CSVerbatimLiterals) Then

                        sb.Append(sbName & ".AppendFormat(" & qChr) 'mySb.AppendFormat("

                        If hasTab Then
                            If csTabEsc.StartsWith("\") Then '\tmyPastedString
                                sb.Append(innerLine.Replace(ControlChars.Tab, "\t"))
                                hasTab = False 'since we won't have it as part of the args
                            Else
                                sb.Append(innerLine.Replace(ControlChars.Tab, "{0}")) '{0}myPastedString
                            End If
                        ElseIf (Not hasTab) Then
                            sb.Append(innerLine) 'myPastedString
                        End If

                        If hasBr Then
                            If hasTab Then
                                If csNewLineEsc.StartsWith("\") Then
                                    sb.Append(csNewLineEsc & qChr) '\n"
                                    hasBr = False 'since we won't have it as part of the args
                                Else
                                    sb.Append("{1}" & qChr) '{1}"
                                End If
                            ElseIf (Not hasTab) Then
                                If csNewLineEsc.StartsWith("\") Then
                                    sb.Append(csNewLineEsc & qChr) '\n"
                                    hasBr = False 'since we won't have it as part of the args
                                Else
                                    sb.Append("{0}" & qChr) '{0}"
                                End If
                            End If
                        ElseIf (Not hasBr) Then
                            sb.Append(qChr) '"
                        End If

                        If hasTab And hasBr Then
                            sb.Append("," & csTabEsc & "," & csNewLineEsc)
                        ElseIf hasTab And (Not hasBr) Then
                            sb.Append("," & csTabEsc)
                        ElseIf (Not hasTab) And hasBr Then
                            sb.Append("," & csNewLineEsc)
                        End If

                        sb.Append(");" & Environment.NewLine) ');<br>

                    End If '(Configuration.CSVerbatimLiterals)

                ElseIf (Not Configuration.AppendFormatOnStringBuilder) Then

                    sb.Append(sbName & ".Append(") 'mySb.Append(

                    If (Configuration.CSVerbatimLiterals) Then
                        If hasTab Then
                            If csTabEsc.StartsWith("\") Then
                                sb.Append("@" & qChr & innerLine.Replace(ControlChars.Tab, qChr & " + " & qChr & csTabEsc & qChr & " + @" & qChr)) '"" + "\t" + "myPastedString
                            Else
                                sb.Append("@" & qChr & innerLine.Replace(ControlChars.Tab, qChr & " + " & csTabEsc & " + @" & qChr)) '"" + tab + "myPastedString
                            End If
                        Else
                            sb.Append("@" & qChr & innerLine) '@"myPastedString
                        End If
                    ElseIf (Not Configuration.CSVerbatimLiterals) Then
                        If hasTab Then
                            If csTabEsc.StartsWith("\") Then
                                sb.Append(qChr & innerLine.Replace(ControlChars.Tab, csTabEsc)) '"\tmyPastedString
                            Else
                                sb.Append(qChr & innerLine.Replace(ControlChars.Tab, qChr & " + " & csTabEsc & " + " & qChr)) '"" + tab + "myPastedString
                            End If
                        Else
                            sb.Append(qChr & innerLine) '"myPastedString
                        End If
                    End If

                    If hasBr Then
                        If csNewLineEsc.StartsWith("\") Then
                            If Configuration.CSVerbatimLiterals Then
                                sb.Append(qChr & " + " & qChr & csNewLineEsc & qChr) '" + "\n"
                            Else
                                sb.Append(csNewLineEsc & qChr) '\n"
                            End If
                        Else
                            sb.Append(qChr & " + " & csNewLineEsc) '" + newline
                        End If
                    ElseIf Not hasBr Then
                        sb.Append(qChr) '"
                    End If

                    sb.Append(");" & Environment.NewLine)  ');<br>
                End If '(Configuration.AppendFormatOnStringBuilder)


                'move L2
                innerLeftIdx = innerRightIdx + Environment.NewLine.Length

            Loop While (innerRightIdx < fullLine.Length) And (innerLeftIdx < fullLine.Length)


            'move L accordingly
            leftIdx = rightIdx + Environment.NewLine.Length

        Loop While (rightIdx < fullString.Length) And (leftIdx < fullString.Length)

        'get rid of "" + tokens
        If Configuration.CSVerbatimLiterals Then
            'TODO: Better '@"" + ' replacement to not cover inside strings
            sb.Replace("@" & qChr & qChr & " + ", "")
        Else
            sb.Replace(qChr & qChr & " + ", "")
        End If

        'add the dec statement
        sb.Insert(0, "StringBuilder " & sbName & " = new StringBuilder(" & txt.Length & ");" & Environment.NewLine)

        'and return
        Return sb.ToString

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Stringinizes text passed to it for use in VB
    ''' </summary>
    ''' <param name="txt">Text to be Stringinized</param>
    ''' <returns>VB Stringinized text</returns>
    ''' <seealso cref='Stringinize' />
    ''' <seealso cref='StringinizeInCS' />
    ''' -----------------------------------------------------------------------------
    Private Function StringbuilderizeInVB(ByVal txt As String, ByVal sbName As String) As String

        'VB quote character -- really just a "
        Dim qChr As String = """"
        Dim vbNewLineEsc As String = Configuration.VBNewLine
        Dim vbTabEsc As String = Configuration.VBTab

        'sb to work with
        Dim sb As New Text.StringBuilder(txt)

        'escape " and {}
        sb.Replace(qChr, qChr & qChr)
        If Configuration.AppendFormatOnStringBuilder Then
            sb.Replace("{", "{{") : sb.Replace("}", "}}")
        End If

        '''''''''''''
        'process the passed string (txt), one line at a time

        'dump the stringbuilder into a temp string
        Dim fullString As String = sb.ToString : sb.Remove(0, sb.Length)

        'pointers that will isolate an individual line in the string
        Dim leftIdx As Integer = 0, rightIdx As Integer = 0
        Do
            'R should be where a linebreak occurs after L, or the end of the string
            rightIdx = fullString.IndexOf(Environment.NewLine, leftIdx)
            If rightIdx = -1 Then rightIdx = fullString.Length

            'now we'll have an individual line to work with, excluding the linebreak
            Dim fullLine As String = fullString.Substring(leftIdx, rightIdx - leftIdx)

            'if wordwrapping is enabled then word wrap it
            If Configuration.EnableWordWrap Then fullLine = WordWrap(fullLine, Configuration.WrapAt)

            'since ln may contain multiple lines, it's necessary to handle each line
            'at once - this loop is very similar to the outerloop
            Dim innerLeftIdx As Integer = 0, innerRightIdx As Integer = 0
            Do
                'same concept as above to extract the inner string
                innerRightIdx = fullLine.IndexOf(Environment.NewLine, innerLeftIdx)
                If innerRightIdx = -1 Then innerRightIdx = fullLine.Length
                Dim innerLine As String = fullLine.Substring(innerLeftIdx, innerRightIdx - innerLeftIdx)


                'these will help us in the output portion
                Dim hasBr, hasTab As Boolean

                'if the R pointer (for the original string) is not at the end, but
                'the R2 Pointer (for the original string defined linebreaks) is
                ' then we should add a BR
                hasBr = _
                    ((rightIdx < fullString.Length) And (innerRightIdx = fullLine.Length)) _
                    And Configuration.VBEscapeLinebreaks

                'this is mostly important for using AppendFormat 
                hasTab = _
                    (innerLine.IndexOf(ControlChars.Tab) > -1) _
                    And Configuration.VBEscapeTabs

                'now build up the line we'll append
                If Configuration.AppendFormatOnStringBuilder Then
                    sb.Append(sbName & ".AppendFormat(" & qChr) 'mySb.AppendFormat("

                    If hasTab Then
                        sb.Append(innerLine.Replace(ControlChars.Tab, "{0}")) '{0}myPastedString
                    ElseIf (Not hasTab) Then
                        sb.Append(innerLine) 'myPastedString
                    End If

                    If hasBr Then
                        If hasTab Then
                            sb.Append("{1}" & qChr) '{1}"
                        ElseIf (Not hasTab) Then
                            sb.Append("{0}" & qChr) '{0}"
                        End If
                    ElseIf (Not hasBr) Then
                        sb.Append(qChr) '"
                    End If

                    If hasTab And hasBr Then
                        sb.Append("," & vbTabEsc & "," & vbNewLineEsc)
                    ElseIf hasTab And (Not hasBr) Then
                        sb.Append("," & vbTabEsc)
                    ElseIf (Not hasTab) And hasBr Then
                        sb.Append("," & vbNewLineEsc)
                    End If

                    sb.Append(")" & Environment.NewLine) ')<br>

                ElseIf (Not Configuration.AppendFormatOnStringBuilder) Then

                    sb.Append(sbName & ".Append(") 'mySb.Append(

                    If hasTab Then
                        sb.Append( _
                            qChr & _
                            innerLine.Replace( _
                                ControlChars.Tab, _
                                qChr & " & " & vbTabEsc & " & " & qChr)) '" & tab & "myPastedString
                    Else
                        sb.Append(qChr & innerLine)   '"myPastedString
                    End If

                    If hasBr Then
                        sb.Append(qChr & " & " & vbNewLineEsc) '" & newline
                    ElseIf Not hasBr Then
                        sb.Append(qChr) '"
                    End If

                    sb.Append(")" & Environment.NewLine)  ')<br>
                End If '(Configuration.AppendFormatOnStringBuilder)


                'move L2
                innerLeftIdx = innerRightIdx + Environment.NewLine.Length

            Loop While (innerRightIdx < fullLine.Length) And (innerLeftIdx < fullLine.Length)


            'move L accordingly
            leftIdx = rightIdx + Environment.NewLine.Length

        Loop While (rightIdx < fullString.Length) And (leftIdx < fullString.Length)

        'add the dec statement
        sb.Insert(0, "Dim " & sbName & " As New StringBuilder(" & txt.Length & ")" & Environment.NewLine)

        'and return
        Return sb.ToString
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Stringinizes text using the default language set in configuration.
    ''' </summary>
    ''' <param name="txt">Text to be Stringinized</param>
    ''' <returns>Stringinized text</returns>
    ''' <seealso cref='StringinizeInCS' />
    ''' <seealso cref='StringinizeInVB' />
    ''' <seealso class='Configuration' />
    ''' -----------------------------------------------------------------------------
    Private Function Stringbuilderize(ByVal txt As String, ByVal sbName As String) As String
        If Configuration.DefaultLanguage = "Visual Basic" Then
            Return StringbuilderizeInVB(txt, sbName)
        Else
            Return StringbuilderizeInCS(txt, sbName)
        End If
    End Function
#End Region

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Adds linebreaks around the value specified in the configuration
    ''' </summary>
    ''' <param name="txt">Text to be wrapped</param>
    ''' <returns>Line-wrapped string</returns>
    ''' -----------------------------------------------------------------------------
    Private Function WordWrap(ByVal txt As String, ByVal WrapAt As Integer) As String
        'only bother doing it if the string is 5 over wrapAt
        If txt.Length < WrapAt + 5 Then Return txt

        Dim ptr, break As Integer
        Dim sb As New Text.StringBuilder

        'loop through our string, making sure we have a good lengthed substring to work with
        While (ptr + WrapAt) < txt.Length - 1

            'find right-most space in our substring
            break = txt.Substring(ptr, WrapAt).LastIndexOf(" "c)

            'no luck, look five spaces ahead, but only if there are five spaces left
            If (break = -1) And ((ptr + WrapAt + 5) < txt.Length) Then
                'find the offest after ptr+WrapAt
                break = txt.Substring(ptr + WrapAt, 5).IndexOf(" "c)
                'if found, add the offset to ptr+WrapAt
                If break <> -1 Then break += ptr + WrapAt
            End If

            'if still no luck, hard wrap
            If break = -1 Then break = Math.Min(WrapAt, (txt.Length - 1) - ptr)

            break += 1

            'add the substring and linebreak
            sb.Append(txt.Substring(ptr, break))
            sb.Append(Environment.NewLine)

            'advance left 
            ptr = ptr + break
        End While
        'since there is most likely a chunk left, handle it
        If ptr < txt.Length - 1 Then
            sb.Append(txt.Substring(ptr, txt.Length - ptr))
        End If

        Return sb.ToString
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Inserts text at current cursor location in the application
    ''' </summary>
    ''' <param name="application">application with activewindow</param>
    ''' <param name="text">text to insert</param>
    ''' -----------------------------------------------------------------------------
    Private Sub Paste(ByVal application As EnvDTE80.DTE2, ByVal text As String)
        'get the text document
        Dim txt As TextDocument = CType(application.ActiveDocument.Object("TextDocument"), TextDocument)

        'get an edit point
        Dim ep As EditPoint = txt.Selection.ActivePoint.CreateEditPoint

        'get a start point
        Dim sp As EditPoint = txt.Selection.ActivePoint.CreateEditPoint

        'open the undo context
        Dim isOpen As Boolean = application.UndoContext.IsOpen
        If Not isOpen Then application.UndoContext.Open("SmartPaster", False)

        'clear the selection
        If Not txt.Selection.IsEmpty Then txt.Selection.Delete()

        'insert the text
        'ep.Insert(Indent(text, ep.LineCharOffset))
        ep.Insert(text)

        'smart format
        If Configuration.AutoFormatAfterPaste Then sp.SmartFormat(ep)

        'close the context
        If Not isOpen Then application.UndoContext.Close()
    End Sub

#Region "Paste As ..."
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Public method to paste and format clipboard text as string the cursor 
    '''   location for the configured or active window's langage .
    ''' </summary>
    ''' <param name="application">application to insert</param>
    ''' -----------------------------------------------------------------------------
    Public Sub PasteAsString(ByVal application As EnvDTE80.DTE2)
        If application.ActiveWindow.Caption.EndsWith(".vb") Then
            Paste(application, StringinizeInVB(ClipboardText))
        ElseIf application.ActiveWindow.Caption.EndsWith(".cs") Then
            Paste(application, StringinizeInCS(ClipboardText))
        Else
            Paste(application, Stringinize(ClipboardText))
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Public method to paste and format clipboard text as comment the cursor 
    '''   location for the configured or active window's langage .
    ''' </summary>
    ''' <param name="application">application to insert</param>
    ''' -----------------------------------------------------------------------------
    Public Sub PasteAsComment(ByVal application As EnvDTE80.DTE2)
        If application.ActiveWindow.Caption.EndsWith(".vb") Then
            Paste(application, CommentizeInVB(ClipboardText))
        ElseIf application.ActiveWindow.Caption.EndsWith(".cs") Then
            Paste(application, CommentizeInCS(ClipboardText))
        Else
            Paste(application, Commentize(ClipboardText))
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Public method to paste format clipboard text into a specified region
    ''' </summary>
    ''' <param name="application">application to insert</param>
    ''' <param name="region">region name to use</param>
    ''' -----------------------------------------------------------------------------
    Public Sub PasteAsRegion(ByVal application As EnvDTE80.DTE2)
        'get the region name
        Dim region As String = InputBox("Name of region to wrap clip text:", , "myRegion")
        If region Is Nothing Or region = String.Empty Then Exit Sub

        'it's so simple, we really don't need a function
        Dim VBRegionized As String = _
                "#Region """ & region & """" & Environment.NewLine & _
                ClipboardText & Environment.NewLine & _
                "#End Region"
        Dim CSRegionized As String = _
                "#region " & region & Environment.NewLine & _
                ClipboardText & Environment.NewLine & _
                "#endregion"

        'detect language
        Dim Regionized As String
        If application.ActiveWindow.Caption.EndsWith(".vb") Then
            Regionized = VBRegionized
        ElseIf application.ActiveWindow.Caption.EndsWith(".cs") Then
            Regionized = CSRegionized
        Else
            If Configuration.DefaultLanguage = "Visual Basic" Then
                Regionized = VBRegionized
            Else
                Regionized = CSRegionized
            End If
        End If

        'and paste
        Paste(application, Regionized)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Public method to paste and format clipboard text as stringbuilder the cursor 
    '''   location for the configured or active window's langage .
    ''' </summary>
    ''' <param name="application">application to insert</param>
    ''' -----------------------------------------------------------------------------
    Public Sub PasteAsStringBuilder(ByVal application As EnvDTE80.DTE2)
        Dim stringbuilder As String = InputBox("Name of StringBuilder to use:", , "myStringBuilder")
        If stringbuilder Is Nothing Or stringbuilder = String.Empty Then Exit Sub

        If application.ActiveWindow.Caption.EndsWith(".vb") Then
            Paste(application, StringbuilderizeInVB(ClipboardText, stringbuilder))
        ElseIf application.ActiveWindow.Caption.EndsWith(".cs") Then
            Paste(application, StringbuilderizeInCS(ClipboardText, stringbuilder))
        Else
            Paste(application, Stringbuilderize(ClipboardText, stringbuilder))
        End If
    End Sub

#End Region

End Class


