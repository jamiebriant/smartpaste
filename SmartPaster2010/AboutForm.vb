Public Class AboutForm
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.AboutText.Text = _
            "Smart Paster Add-In 1.1 " & Environment.NewLine & _
            "(C) 2010 Inedo " & Environment.NewLine & _
            "www.inedo.com" & Environment.NewLine & _
            Environment.NewLine & _
            "If you find this tool useful, feel free to drop us an email and " & _
            "let us know. And remember, if you're thinking of using this to " & _
            "easily paste you 100-line SQL Statements, I encourage you to " & _
            "strongly consider stored procedures :-)"
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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents AboutText As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button
        Me.AboutText = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(208, 144)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(72, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "OK"
        '
        'AboutText
        '
        Me.AboutText.Location = New System.Drawing.Point(8, 8)
        Me.AboutText.Multiline = True
        Me.AboutText.Name = "AboutText"
        Me.AboutText.Size = New System.Drawing.Size(272, 128)
        Me.AboutText.TabIndex = 3
        Me.AboutText.Text = ""
        '
        'AboutForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(290, 178)
        Me.Controls.Add(Me.AboutText)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "AboutForm"
        Me.Text = "About Smart Paster Add-In"
        Me.ResumeLayout(False)

    End Sub

#End Region



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub AboutText_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles AboutText.KeyPress
        e.Handled = False
    End Sub

    Private Sub AboutForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
