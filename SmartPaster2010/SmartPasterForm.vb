Imports System.Windows.Forms

Friend Class SmartPasterForm
    Inherits Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        'ToolTip.SetToolTip(DefaultLang, "Language to use when the file does not end in .vb or .cs")
        'ToolTip.SetToolTip(DefaultLangVB, ToolTip.GetToolTip(DefaultLang))
        'ToolTip.SetToolTip(DefaultLangCS, ToolTip.GetToolTip(DefaultLang))

        'ToolTip.SetToolTip(VBNewLine, "String to insert for a New Line (e.g., vbCrLf, CotnrolChars.NewLine, Enviornment.NewLine)")
        'ToolTip.SetToolTip(VBNewLineLabel, ToolTip.GetToolTip(VBNewLine))

        'ToolTip.SetToolTip(VBTab, "String to insert for a Tab (e.g., vbTab, CotnrolChars.Tab)")
        'ToolTip.SetToolTip(VBTabLabel, ToolTip.GetToolTip(VBTab))

        ''ToolTip.SetToolTip(GenEscTabs, "Replace tab characters with '\t' or the default tab character")

        'ToolTip.SetToolTip(GenWrapAt, "Approximage length at which to wrap lines. Set 0 to disable wrapping")
        'ToolTip.SetToolTip(GenWrapAtLabel, ToolTip.GetToolTip(GenWrapAt))
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents Save As System.Windows.Forms.Button
    Friend WithEvents Cancel As System.Windows.Forms.Button
    Friend WithEvents CSGroup As System.Windows.Forms.GroupBox
    Friend WithEvents CSLinebreakEscape As System.Windows.Forms.ComboBox
    Friend WithEvents CSTabEscape As System.Windows.Forms.ComboBox
    Friend WithEvents CSDefaultLanguage As System.Windows.Forms.CheckBox
    Friend WithEvents VBGroup As System.Windows.Forms.GroupBox
    Friend WithEvents VBDefaultLanguage As System.Windows.Forms.CheckBox
    Friend WithEvents VBLinebreakEscape As System.Windows.Forms.ComboBox
    Friend WithEvents VBTabEscape As System.Windows.Forms.ComboBox
    Friend WithEvents CSVerbatimLiterals As System.Windows.Forms.CheckBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents AppendFormatOnStringBuilder As System.Windows.Forms.CheckBox
    Friend WithEvents About As System.Windows.Forms.Button
    Friend WithEvents EnablePasteAsComment As System.Windows.Forms.CheckBox
    Friend WithEvents EnablePasteAsRegion As System.Windows.Forms.CheckBox
    Friend WithEvents EnablePasteAsStringBuilder As System.Windows.Forms.CheckBox
    Friend WithEvents EnablePasteAsString As System.Windows.Forms.CheckBox
    Friend WithEvents EnableWordWrap As System.Windows.Forms.CheckBox
    Friend WithEvents WrapAt As System.Windows.Forms.TextBox
    Friend WithEvents AutoFormatAfterPaste As System.Windows.Forms.CheckBox
    Friend WithEvents CSVerbatimLiteralsSpanLines As System.Windows.Forms.CheckBox
    Friend WithEvents GeneralGroup As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents WrapAtLabel As System.Windows.Forms.Label
    Friend WithEvents CSLinebreakEscapeLabel As System.Windows.Forms.Label
    Friend WithEvents CSTabEscapeLabel As System.Windows.Forms.Label
    Friend WithEvents VBLinebreakEscapeLabel As System.Windows.Forms.Label
    Friend WithEvents VBTabEscapeLabel As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Save = New System.Windows.Forms.Button()
        Me.GeneralGroup = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.AppendFormatOnStringBuilder = New System.Windows.Forms.CheckBox()
        Me.WrapAt = New System.Windows.Forms.TextBox()
        Me.EnableWordWrap = New System.Windows.Forms.CheckBox()
        Me.WrapAtLabel = New System.Windows.Forms.Label()
        Me.AutoFormatAfterPaste = New System.Windows.Forms.CheckBox()
        Me.EnablePasteAsComment = New System.Windows.Forms.CheckBox()
        Me.EnablePasteAsRegion = New System.Windows.Forms.CheckBox()
        Me.EnablePasteAsStringBuilder = New System.Windows.Forms.CheckBox()
        Me.EnablePasteAsString = New System.Windows.Forms.CheckBox()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.CSVerbatimLiteralsSpanLines = New System.Windows.Forms.CheckBox()
        Me.CSDefaultLanguage = New System.Windows.Forms.CheckBox()
        Me.CSLinebreakEscape = New System.Windows.Forms.ComboBox()
        Me.CSTabEscape = New System.Windows.Forms.ComboBox()
        Me.CSVerbatimLiterals = New System.Windows.Forms.CheckBox()
        Me.VBDefaultLanguage = New System.Windows.Forms.CheckBox()
        Me.VBLinebreakEscape = New System.Windows.Forms.ComboBox()
        Me.VBTabEscape = New System.Windows.Forms.ComboBox()
        Me.Cancel = New System.Windows.Forms.Button()
        Me.CSLinebreakEscapeLabel = New System.Windows.Forms.Label()
        Me.CSTabEscapeLabel = New System.Windows.Forms.Label()
        Me.VBLinebreakEscapeLabel = New System.Windows.Forms.Label()
        Me.VBTabEscapeLabel = New System.Windows.Forms.Label()
        Me.About = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CSGroup = New System.Windows.Forms.GroupBox()
        Me.VBGroup = New System.Windows.Forms.GroupBox()
        Me.GeneralGroup.SuspendLayout()
        Me.CSGroup.SuspendLayout()
        Me.VBGroup.SuspendLayout()
        Me.SuspendLayout()
        '
        'Save
        '
        Me.Save.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Save.Location = New System.Drawing.Point(10, 180)
        Me.Save.Name = "Save"
        Me.Save.Size = New System.Drawing.Size(65, 25)
        Me.Save.TabIndex = 3
        Me.Save.Text = "Save"
        Me.ToolTip.SetToolTip(Me.Save, "Save changes and close window")
        '
        'GeneralGroup
        '
        Me.GeneralGroup.Controls.Add(Me.Label6)
        Me.GeneralGroup.Controls.Add(Me.AppendFormatOnStringBuilder)
        Me.GeneralGroup.Controls.Add(Me.WrapAt)
        Me.GeneralGroup.Controls.Add(Me.EnableWordWrap)
        Me.GeneralGroup.Controls.Add(Me.WrapAtLabel)
        Me.GeneralGroup.Controls.Add(Me.AutoFormatAfterPaste)
        Me.GeneralGroup.Controls.Add(Me.EnablePasteAsComment)
        Me.GeneralGroup.Controls.Add(Me.EnablePasteAsRegion)
        Me.GeneralGroup.Controls.Add(Me.EnablePasteAsStringBuilder)
        Me.GeneralGroup.Controls.Add(Me.EnablePasteAsString)
        Me.GeneralGroup.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GeneralGroup.Location = New System.Drawing.Point(5, 5)
        Me.GeneralGroup.Name = "GeneralGroup"
        Me.GeneralGroup.Size = New System.Drawing.Size(225, 160)
        Me.GeneralGroup.TabIndex = 4
        Me.GeneralGroup.TabStop = False
        Me.GeneralGroup.Text = "General"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(10, 20)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(150, 15)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Enable Paste As ..."
        Me.ToolTip.SetToolTip(Me.Label6, "Enables or Disables the option on the Context Menu")
        '
        'AppendFormatOnStringBuilder
        '
        Me.AppendFormatOnStringBuilder.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AppendFormatOnStringBuilder.Location = New System.Drawing.Point(10, 85)
        Me.AppendFormatOnStringBuilder.Name = "AppendFormatOnStringBuilder"
        Me.AppendFormatOnStringBuilder.Size = New System.Drawing.Size(205, 20)
        Me.AppendFormatOnStringBuilder.TabIndex = 2
        Me.AppendFormatOnStringBuilder.Text = "AppendFormat on StringBuilder"
        Me.ToolTip.SetToolTip(Me.AppendFormatOnStringBuilder, "Indicates whether the AppendFormat method (along with {0} and {1} replacements) s" & _
                "hould be used.")
        '
        'WrapAt
        '
        Me.WrapAt.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.WrapAt.Location = New System.Drawing.Point(110, 110)
        Me.WrapAt.MaxLength = 3
        Me.WrapAt.Name = "WrapAt"
        Me.WrapAt.Size = New System.Drawing.Size(35, 21)
        Me.WrapAt.TabIndex = 1
        Me.ToolTip.SetToolTip(Me.WrapAt, "Approximate column at which to wrap text at")
        '
        'EnableWordWrap
        '
        Me.EnableWordWrap.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EnableWordWrap.Location = New System.Drawing.Point(10, 110)
        Me.EnableWordWrap.Name = "EnableWordWrap"
        Me.EnableWordWrap.Size = New System.Drawing.Size(110, 20)
        Me.EnableWordWrap.TabIndex = 2
        Me.EnableWordWrap.Text = "Word Wrap At "
        Me.ToolTip.SetToolTip(Me.EnableWordWrap, "Indicates that text should be wrapped")
        '
        'WrapAtLabel
        '
        Me.WrapAtLabel.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.WrapAtLabel.Location = New System.Drawing.Point(150, 110)
        Me.WrapAtLabel.Name = "WrapAtLabel"
        Me.WrapAtLabel.Size = New System.Drawing.Size(70, 20)
        Me.WrapAtLabel.TabIndex = 0
        Me.WrapAtLabel.Text = "Characters"
        Me.WrapAtLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ToolTip.SetToolTip(Me.WrapAtLabel, "Approximate column at which to wrap text at")
        '
        'AutoFormatAfterPaste
        '
        Me.AutoFormatAfterPaste.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AutoFormatAfterPaste.Location = New System.Drawing.Point(10, 135)
        Me.AutoFormatAfterPaste.Name = "AutoFormatAfterPaste"
        Me.AutoFormatAfterPaste.Size = New System.Drawing.Size(205, 20)
        Me.AutoFormatAfterPaste.TabIndex = 2
        Me.AutoFormatAfterPaste.Text = "Auto Format After Paste"
        Me.ToolTip.SetToolTip(Me.AutoFormatAfterPaste, "Indicates whether the pasted text should be autoformatted after a paste.")
        '
        'EnablePasteAsComment
        '
        Me.EnablePasteAsComment.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EnablePasteAsComment.Location = New System.Drawing.Point(30, 35)
        Me.EnablePasteAsComment.Name = "EnablePasteAsComment"
        Me.EnablePasteAsComment.Size = New System.Drawing.Size(80, 20)
        Me.EnablePasteAsComment.TabIndex = 2
        Me.EnablePasteAsComment.Text = "Comment"
        Me.ToolTip.SetToolTip(Me.EnablePasteAsComment, "Enables or Disables the option on the Context Menu")
        '
        'EnablePasteAsRegion
        '
        Me.EnablePasteAsRegion.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EnablePasteAsRegion.Location = New System.Drawing.Point(30, 55)
        Me.EnablePasteAsRegion.Name = "EnablePasteAsRegion"
        Me.EnablePasteAsRegion.Size = New System.Drawing.Size(80, 20)
        Me.EnablePasteAsRegion.TabIndex = 2
        Me.EnablePasteAsRegion.Text = "Region"
        Me.ToolTip.SetToolTip(Me.EnablePasteAsRegion, "Enables or Disables the option on the Context Menu")
        '
        'EnablePasteAsStringBuilder
        '
        Me.EnablePasteAsStringBuilder.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EnablePasteAsStringBuilder.Location = New System.Drawing.Point(110, 55)
        Me.EnablePasteAsStringBuilder.Name = "EnablePasteAsStringBuilder"
        Me.EnablePasteAsStringBuilder.Size = New System.Drawing.Size(100, 20)
        Me.EnablePasteAsStringBuilder.TabIndex = 2
        Me.EnablePasteAsStringBuilder.Text = "StringBuilder"
        Me.ToolTip.SetToolTip(Me.EnablePasteAsStringBuilder, "Enables or Disables the option on the Context Menu")
        '
        'EnablePasteAsString
        '
        Me.EnablePasteAsString.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EnablePasteAsString.Location = New System.Drawing.Point(110, 35)
        Me.EnablePasteAsString.Name = "EnablePasteAsString"
        Me.EnablePasteAsString.Size = New System.Drawing.Size(80, 20)
        Me.EnablePasteAsString.TabIndex = 2
        Me.EnablePasteAsString.Text = "String"
        Me.ToolTip.SetToolTip(Me.EnablePasteAsString, "Enables or Disables the option on the Context Menu")
        '
        'ToolTip
        '
        Me.ToolTip.AutoPopDelay = 5000
        Me.ToolTip.InitialDelay = 100
        Me.ToolTip.ReshowDelay = 100
        '
        'CSVerbatimLiteralsSpanLines
        '
        Me.CSVerbatimLiteralsSpanLines.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CSVerbatimLiteralsSpanLines.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CSVerbatimLiteralsSpanLines.Location = New System.Drawing.Point(145, 45)
        Me.CSVerbatimLiteralsSpanLines.Name = "CSVerbatimLiteralsSpanLines"
        Me.CSVerbatimLiteralsSpanLines.Size = New System.Drawing.Size(85, 20)
        Me.CSVerbatimLiteralsSpanLines.TabIndex = 3
        Me.CSVerbatimLiteralsSpanLines.Text = "Span Lines"
        Me.ToolTip.SetToolTip(Me.CSVerbatimLiteralsSpanLines, "Indicates that Verbatim Literals should span multiple lines and not escape any co" & _
                "ntrol characters.")
        '
        'CSDefaultLanguage
        '
        Me.CSDefaultLanguage.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CSDefaultLanguage.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CSDefaultLanguage.Location = New System.Drawing.Point(10, 20)
        Me.CSDefaultLanguage.Name = "CSDefaultLanguage"
        Me.CSDefaultLanguage.Size = New System.Drawing.Size(125, 20)
        Me.CSDefaultLanguage.TabIndex = 2
        Me.CSDefaultLanguage.Text = "Default Language"
        Me.ToolTip.SetToolTip(Me.CSDefaultLanguage, "Indicates that C# should be used when the extension of the active file is not "".C" & _
                "S"" or "".VB""")
        '
        'CSLinebreakEscape
        '
        Me.CSLinebreakEscape.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CSLinebreakEscape.Font = New System.Drawing.Font("Trebuchet MS", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CSLinebreakEscape.Items.AddRange(New Object() {"<none>", "\n", "\r\n", "ControlChars.NewLine", "Environment.NewLine"})
        Me.CSLinebreakEscape.Location = New System.Drawing.Point(120, 65)
        Me.CSLinebreakEscape.Name = "CSLinebreakEscape"
        Me.CSLinebreakEscape.Size = New System.Drawing.Size(130, 24)
        Me.CSLinebreakEscape.TabIndex = 0
        Me.ToolTip.SetToolTip(Me.CSLinebreakEscape, "String to insert when escaping a linebreak. Strings starting with ""\"" are interpr" & _
                "eted as escape sequences.")
        '
        'CSTabEscape
        '
        Me.CSTabEscape.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CSTabEscape.Font = New System.Drawing.Font("Trebuchet MS", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CSTabEscape.Items.AddRange(New Object() {"<none>", "\t", "ControlChars.Tab"})
        Me.CSTabEscape.Location = New System.Drawing.Point(120, 90)
        Me.CSTabEscape.Name = "CSTabEscape"
        Me.CSTabEscape.Size = New System.Drawing.Size(130, 24)
        Me.CSTabEscape.TabIndex = 0
        Me.ToolTip.SetToolTip(Me.CSTabEscape, "String to insert when escaping a tab. Strings starting with ""\"" are interpreted a" & _
                "s escape sequences.")
        '
        'CSVerbatimLiterals
        '
        Me.CSVerbatimLiterals.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.CSVerbatimLiterals.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CSVerbatimLiterals.Location = New System.Drawing.Point(10, 45)
        Me.CSVerbatimLiterals.Name = "CSVerbatimLiterals"
        Me.CSVerbatimLiterals.Size = New System.Drawing.Size(125, 20)
        Me.CSVerbatimLiterals.TabIndex = 2
        Me.CSVerbatimLiterals.Text = "Verbatim Literals"
        Me.ToolTip.SetToolTip(Me.CSVerbatimLiterals, "Indicates whether string literals should be prefixed with ""@"" and not use any esc" & _
                "ape sequences. Note that escape strings will be concatenated if selected.")
        '
        'VBDefaultLanguage
        '
        Me.VBDefaultLanguage.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.VBDefaultLanguage.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VBDefaultLanguage.Location = New System.Drawing.Point(10, 20)
        Me.VBDefaultLanguage.Name = "VBDefaultLanguage"
        Me.VBDefaultLanguage.Size = New System.Drawing.Size(126, 20)
        Me.VBDefaultLanguage.TabIndex = 2
        Me.VBDefaultLanguage.Text = "Default Language"
        Me.ToolTip.SetToolTip(Me.VBDefaultLanguage, "Indicates that Visual Basic should be used when the extension of the active file " & _
                "is not "".CS"" or "".VB""")
        '
        'VBLinebreakEscape
        '
        Me.VBLinebreakEscape.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.VBLinebreakEscape.Font = New System.Drawing.Font("Trebuchet MS", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VBLinebreakEscape.Items.AddRange(New Object() {"<none>", "ControlChars.NewLine", "Environment.NewLine", "vbCrLf"})
        Me.VBLinebreakEscape.Location = New System.Drawing.Point(120, 40)
        Me.VBLinebreakEscape.Name = "VBLinebreakEscape"
        Me.VBLinebreakEscape.Size = New System.Drawing.Size(128, 24)
        Me.VBLinebreakEscape.TabIndex = 0
        Me.ToolTip.SetToolTip(Me.VBLinebreakEscape, "String to insert when escaping a linebreak")
        '
        'VBTabEscape
        '
        Me.VBTabEscape.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.VBTabEscape.Font = New System.Drawing.Font("Trebuchet MS", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VBTabEscape.Items.AddRange(New Object() {"<none>", "ControlChars.Tab", "vbTab"})
        Me.VBTabEscape.Location = New System.Drawing.Point(120, 65)
        Me.VBTabEscape.Name = "VBTabEscape"
        Me.VBTabEscape.Size = New System.Drawing.Size(128, 24)
        Me.VBTabEscape.TabIndex = 0
        Me.ToolTip.SetToolTip(Me.VBTabEscape, "String to insert when escaping a Tab")
        '
        'Cancel
        '
        Me.Cancel.Location = New System.Drawing.Point(85, 180)
        Me.Cancel.Name = "Cancel"
        Me.Cancel.Size = New System.Drawing.Size(65, 25)
        Me.Cancel.TabIndex = 3
        Me.Cancel.Text = "Cancel"
        Me.ToolTip.SetToolTip(Me.Cancel, "Cancel changes and close window")
        '
        'CSLinebreakEscapeLabel
        '
        Me.CSLinebreakEscapeLabel.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CSLinebreakEscapeLabel.Location = New System.Drawing.Point(10, 70)
        Me.CSLinebreakEscapeLabel.Name = "CSLinebreakEscapeLabel"
        Me.CSLinebreakEscapeLabel.Size = New System.Drawing.Size(105, 15)
        Me.CSLinebreakEscapeLabel.TabIndex = 1
        Me.CSLinebreakEscapeLabel.Text = "Linebreak Escape"
        Me.ToolTip.SetToolTip(Me.CSLinebreakEscapeLabel, "String to insert when escaping a linebreak. Strings starting with ""\"" are interpr" & _
                "eted as escape sequences.")
        '
        'CSTabEscapeLabel
        '
        Me.CSTabEscapeLabel.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CSTabEscapeLabel.Location = New System.Drawing.Point(10, 95)
        Me.CSTabEscapeLabel.Name = "CSTabEscapeLabel"
        Me.CSTabEscapeLabel.Size = New System.Drawing.Size(105, 15)
        Me.CSTabEscapeLabel.TabIndex = 1
        Me.CSTabEscapeLabel.Text = "Tab Escape"
        Me.ToolTip.SetToolTip(Me.CSTabEscapeLabel, "String to insert when escaping a tab. Strings starting with ""\"" are interpreted a" & _
                "s escape sequences.")
        '
        'VBLinebreakEscapeLabel
        '
        Me.VBLinebreakEscapeLabel.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VBLinebreakEscapeLabel.Location = New System.Drawing.Point(10, 45)
        Me.VBLinebreakEscapeLabel.Name = "VBLinebreakEscapeLabel"
        Me.VBLinebreakEscapeLabel.Size = New System.Drawing.Size(105, 15)
        Me.VBLinebreakEscapeLabel.TabIndex = 1
        Me.VBLinebreakEscapeLabel.Text = "Linebreak Escape"
        Me.ToolTip.SetToolTip(Me.VBLinebreakEscapeLabel, "String to insert when escaping a linebreak")
        '
        'VBTabEscapeLabel
        '
        Me.VBTabEscapeLabel.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VBTabEscapeLabel.Location = New System.Drawing.Point(10, 70)
        Me.VBTabEscapeLabel.Name = "VBTabEscapeLabel"
        Me.VBTabEscapeLabel.Size = New System.Drawing.Size(105, 16)
        Me.VBTabEscapeLabel.TabIndex = 1
        Me.VBTabEscapeLabel.Text = "Tab Escape"
        Me.ToolTip.SetToolTip(Me.VBTabEscapeLabel, "String to insert when escaping a Tab")
        '
        'About
        '
        Me.About.Location = New System.Drawing.Point(160, 180)
        Me.About.Name = "About"
        Me.About.Size = New System.Drawing.Size(65, 25)
        Me.About.TabIndex = 3
        Me.About.Text = "About"
        Me.ToolTip.SetToolTip(Me.About, "Show the About Dialog")
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Trebuchet MS", 8.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 210)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(210, 15)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Note that all options have tooltip text"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ToolTip.SetToolTip(Me.Label1, "Even this one :-)")
        '
        'CSGroup
        '
        Me.CSGroup.Controls.Add(Me.CSVerbatimLiteralsSpanLines)
        Me.CSGroup.Controls.Add(Me.CSDefaultLanguage)
        Me.CSGroup.Controls.Add(Me.CSLinebreakEscapeLabel)
        Me.CSGroup.Controls.Add(Me.CSLinebreakEscape)
        Me.CSGroup.Controls.Add(Me.CSTabEscapeLabel)
        Me.CSGroup.Controls.Add(Me.CSTabEscape)
        Me.CSGroup.Controls.Add(Me.CSVerbatimLiterals)
        Me.CSGroup.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CSGroup.Location = New System.Drawing.Point(240, 110)
        Me.CSGroup.Name = "CSGroup"
        Me.CSGroup.Size = New System.Drawing.Size(255, 120)
        Me.CSGroup.TabIndex = 9
        Me.CSGroup.TabStop = False
        Me.CSGroup.Text = "C#"
        '
        'VBGroup
        '
        Me.VBGroup.Controls.Add(Me.VBDefaultLanguage)
        Me.VBGroup.Controls.Add(Me.VBLinebreakEscapeLabel)
        Me.VBGroup.Controls.Add(Me.VBLinebreakEscape)
        Me.VBGroup.Controls.Add(Me.VBTabEscapeLabel)
        Me.VBGroup.Controls.Add(Me.VBTabEscape)
        Me.VBGroup.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VBGroup.Location = New System.Drawing.Point(240, 5)
        Me.VBGroup.Name = "VBGroup"
        Me.VBGroup.Size = New System.Drawing.Size(255, 95)
        Me.VBGroup.TabIndex = 9
        Me.VBGroup.TabStop = False
        Me.VBGroup.Text = "Visual Basic"
        '
        'SmartPasterForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 14)
        Me.ClientSize = New System.Drawing.Size(499, 233)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CSGroup)
        Me.Controls.Add(Me.GeneralGroup)
        Me.Controls.Add(Me.Save)
        Me.Controls.Add(Me.Cancel)
        Me.Controls.Add(Me.VBGroup)
        Me.Controls.Add(Me.About)
        Me.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "SmartPasterForm"
        Me.ShowInTaskbar = False
        Me.Text = "Smart Paster 1.1 Add-In Configuration"
        Me.GeneralGroup.ResumeLayout(False)
        Me.GeneralGroup.PerformLayout()
        Me.CSGroup.ResumeLayout(False)
        Me.VBGroup.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region


    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.Click
        Try
            Dim val As Integer = CInt(WrapAt.Text)
            If val < 5 Then val = 100
        Catch ex As Exception
        End Try

        'would do "With Configuration" except it's a singleton class and the
        'With statement requires a non-type identifier

        'general options
        Configuration.EnablePasteAsComment = EnablePasteAsComment.Checked
        Configuration.EnablePasteAsRegion = EnablePasteAsRegion.Checked
        Configuration.EnablePasteAsString = EnablePasteAsString.Checked
        Configuration.EnablePasteAsStringBuilder = EnablePasteAsStringBuilder.Checked
        Configuration.AppendFormatOnStringBuilder = AppendFormatOnStringBuilder.Checked
        Configuration.EnableWordWrap = EnableWordWrap.Checked
        Configuration.WrapAt = CInt(WrapAt.Text)
        Configuration.AutoFormatAfterPaste = AutoFormatAfterPaste.Checked
        If VBDefaultLanguage.Checked _
            Then Configuration.DefaultLanguage = "Visual Basic" _
            Else Configuration.DefaultLanguage = "C#"

        'vb
        Configuration.VBNewLine = VBLinebreakEscape.Text
        Configuration.VBEscapeLinebreaks = Not ( _
            VBLinebreakEscape.Text.StartsWith("<none>") _
            Or VBLinebreakEscape.Text = String.Empty)
        Configuration.VBTab = VBTabEscape.Text
        Configuration.VBEscapeTabs = Not ( _
            VBTabEscape.Text.StartsWith("<none>") _
            Or VBTabEscape.Text = String.Empty)

        'c#
        Configuration.CSNewLine = CSLinebreakEscape.Text
        Configuration.CSEscapeLinebreaks = Not ( _
            CSLinebreakEscape.Text.StartsWith("<none>") _
            Or CSLinebreakEscape.Text = String.Empty)
        Configuration.CSTab = CSTabEscape.Text
        Configuration.CSEscapeTabs = Not ( _
            CSTabEscape.Text.StartsWith("<none>") _
            Or CSTabEscape.Text = String.Empty)
        Configuration.CSVerbatimLiterals = CSVerbatimLiterals.Checked
        Configuration.CSVerbatimLiteralsSpanLines = CSVerbatimLiteralsSpanLines.Checked

        Close()
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Close()
    End Sub

    Private Sub About_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles About.Click
        Dim about As New AboutForm
        about.ShowDialog()
    End Sub

    Private Sub SmartPasterForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'general options
        EnablePasteAsComment.Checked = Configuration.EnablePasteAsComment
        EnablePasteAsRegion.Checked = Configuration.EnablePasteAsRegion
        EnablePasteAsString.Checked = Configuration.EnablePasteAsString
        EnablePasteAsStringBuilder.Checked = Configuration.EnablePasteAsStringBuilder
        AppendFormatOnStringBuilder.Checked = Configuration.AppendFormatOnStringBuilder
        EnableWordWrap.Checked = Configuration.EnableWordWrap
        WrapAt.Text = Configuration.WrapAt.ToString
        AutoFormatAfterPaste.Checked = Configuration.AutoFormatAfterPaste
        VBDefaultLanguage.Checked = (Configuration.DefaultLanguage = "Visual Basic")
        CSDefaultLanguage.Checked = Not VBDefaultLanguage.Checked

        'vb
        VBLinebreakEscape.Text = Configuration.VBNewLine
        If Not Configuration.VBEscapeLinebreaks Then VBLinebreakEscape.Text = "<none>"
        VBTabEscape.Text = Configuration.VBTab
        If Not Configuration.VBEscapeTabs Then VBTabEscape.Text = "<none>"

        'c#
        CSLinebreakEscape.Text = Configuration.CSNewLine
        If Not Configuration.CSEscapeLinebreaks Then CSLinebreakEscape.Text = "<none>"
        CSTabEscape.Text = Configuration.CSTab
        If Not Configuration.CSEscapeTabs Then CSTabEscape.Text = "<none>"
        CSVerbatimLiterals.Checked = Configuration.CSVerbatimLiterals
        CSVerbatimLiteralsSpanLines.Checked = Configuration.CSVerbatimLiteralsSpanLines

        WrapAt.Enabled = EnableWordWrap.Checked
        CSVerbatimLiteralsSpanLines.Enabled = CSVerbatimLiterals.Checked
    End Sub

    Private Sub GenWrapAt_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles WrapAt.KeyPress
        If Not [Char].IsDigit(e.KeyChar) Then e.Handled = False
    End Sub

    Private Sub CSVerbatimLiterals_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CSVerbatimLiterals.CheckedChanged
        CSVerbatimLiteralsSpanLines.Enabled = CSVerbatimLiterals.Checked
    End Sub


    Private Sub VBDefaultLanguage_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VBDefaultLanguage.CheckedChanged
        CSDefaultLanguage.Checked = Not VBDefaultLanguage.Checked
    End Sub

    Private Sub CSDefaultLanguage_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CSDefaultLanguage.CheckedChanged
        VBDefaultLanguage.Checked = Not CSDefaultLanguage.Checked
    End Sub

    Private Sub EnableWordWrap_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EnableWordWrap.CheckedChanged
        WrapAt.Enabled = EnableWordWrap.Checked
    End Sub

    Private Sub WrapAt_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles WrapAt.Validating
        Try
            Dim val As Integer = CInt(WrapAt.Text)
            If val < 5 Then val = 100
        Catch ex As Exception
            WrapAt.Text = "100"
        End Try
    End Sub
End Class
