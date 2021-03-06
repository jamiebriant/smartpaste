Imports System
Imports Microsoft.VisualStudio.CommandBars
Imports Extensibility
Imports EnvDTE
Imports EnvDTE80

Imports System.Runtime.InteropServices


<ProgId("SmartPaster2008.Connect")> _
Public Class Connect

    Implements IDTExtensibility2
    Implements IDTCommandTarget

#Region "Members"
    Dim _applicationObject As DTE2
    Dim _addInInstance As AddIn


    'needed to save the references, or the handlers dissaper :-(
    Private pasteAsButtons As ArrayList
    Private pasteAsPopup As CommandBarPopup

    'to store our instance of the SP
    Private smartPaster As SmartPaster

    'event capturers
#End Region

#Region "EnableContextMenuButtons"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Shows/Hides context menu buttons based on configuration settings
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub EnableContextMenuButtons()
        For Each pasteAsButton As CommandBarButton In pasteAsButtons
            Select Case pasteAsButton.Caption
                Case "Comment" : pasteAsButton.Visible = Configuration.EnablePasteAsComment
                Case "String" : pasteAsButton.Visible = Configuration.EnablePasteAsString
                Case "StringBuilder" : pasteAsButton.Visible = Configuration.EnablePasteAsStringBuilder
                Case "Region" : pasteAsButton.Visible = Configuration.EnablePasteAsRegion
            End Select
        Next
    End Sub

#End Region

    '''<summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
    Public Sub New()

    End Sub

    '''<summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
    '''<param name='application'>Root object of the host application.</param>
    '''<param name='connectMode'>Describes how the Add-in is being loaded.</param>
    '''<param name='addInInst'>Object representing this Add-in.</param>
    '''<remarks></remarks>
    Public Sub OnConnection(ByVal application As Object, ByVal connectMode As ext_ConnectMode, ByVal addInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection
        Try
            _applicationObject = CType(application, DTE2)
            _addInInstance = CType(addInInst, AddIn)

            smartPaster = New SmartPaster
            pasteAsButtons = New ArrayList

            'check for the commands
            Dim cmdExists As Boolean
            For Each cmd As Command In _applicationObject.Commands
                If cmd.Name = "SmartPaster2008.Connect.Configure" Then
                    cmdExists = True
                    Exit For
                End If
            Next

            If Not cmdExists Then
                _applicationObject.Commands.AddNamedCommand( _
                    _addInInstance, "Configure", "Configure SmartPaster", _
                    "Opens the Smart Paster Configuration Dialog", _
                    True, 548, , _
                    CType(vsCommandStatus.vsCommandStatusSupported, Integer) _
                        + CType(vsCommandStatus.vsCommandStatusEnabled, Integer))

                _applicationObject.Commands.AddNamedCommand( _
                    _addInInstance, "PasteAsComment", "Paste As Comment", _
                    "Pastes clipboard text as a comment.", _
                    True, 22, , _
                    CType(vsCommandStatus.vsCommandStatusSupported, Integer) _
                        + CType(vsCommandStatus.vsCommandStatusEnabled, Integer))

                _applicationObject.Commands.AddNamedCommand( _
                    _addInInstance, "PasteAsString", "Paste As String", _
                    "Pastes clipboard text as a string literal.", _
                    True, 22, , _
                    CType(vsCommandStatus.vsCommandStatusSupported, Integer) _
                        + CType(vsCommandStatus.vsCommandStatusEnabled, Integer))

                _applicationObject.Commands.AddNamedCommand( _
                    _addInInstance, "PasteAsStringBuilder", "Paste As StringBuilder", _
                    "Pastes clipboard text as a stringbuilder.", _
                    True, 22, , _
                    CType(vsCommandStatus.vsCommandStatusSupported, Integer) _
                        + CType(vsCommandStatus.vsCommandStatusEnabled, Integer))

                _applicationObject.Commands.AddNamedCommand( _
                    _addInInstance, "PasteAsRegion", "Paste As Region", _
                    "Pastes clipboard text in a specified region.", _
                    True, 22, , _
                    CType(vsCommandStatus.vsCommandStatusSupported, Integer) _
                        + CType(vsCommandStatus.vsCommandStatusEnabled, Integer))
            End If


            If connectMode = ext_ConnectMode.ext_cm_Startup Then

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Add items to the Context (Right-Click) Menu
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'find the position of the &Paste command
                Dim position As Integer = 0

                Dim bars As CommandBars = CType(_applicationObject.CommandBars, CommandBars)
                Dim codeWindow As CommandBar = bars.Item("Code Window")

                For i As Integer = 1 To codeWindow.Controls.Count
                    If codeWindow.Controls(i).Caption = "&Paste" Then
                        position = i
                        Exit For
                    End If
                Next

                'add the popup menu "Paste As...", which may already be on the menu
                pasteAsPopup = CType(codeWindow.Controls.Add(MsoControlType.msoControlPopup, 1, Type.Missing, position + 1, Type.Missing), CommandBarPopup)
                pasteAsPopup.Caption = "Paste As ..."

                'now the buttons
                Dim pasteAsButton As CommandBarButton

                'add "Comment"
                pasteAsButton = CType(pasteAsPopup.Controls.Add(MsoControlType.msoControlButton), CommandBarButton)
                pasteAsButton.Caption = "Comment"
                pasteAsButton.TooltipText = "Inserts clipboard with each line prefixed with a comment character"
				'pasteAsButton.FaceId = 22
                AddHandler pasteAsButton.Click, AddressOf PasteAsComment
                pasteAsButtons.Add(pasteAsButton)

                'add "String"
                pasteAsButton = CType(pasteAsPopup.Controls.Add(MsoControlType.msoControlButton), CommandBarButton)
                pasteAsButton.Caption = "String"
                pasteAsButton.TooltipText = "Inserts enquoted clipboard text with line breaks and other characters escaped"
				'pasteAsButton.FaceId = 22
                AddHandler pasteAsButton.Click, AddressOf PasteAsString
                pasteAsButtons.Add(pasteAsButton)

                'add "StringBuilder"
                pasteAsButton = CType(pasteAsPopup.Controls.Add(MsoControlType.msoControlButton), CommandBarButton)
                pasteAsButton.Caption = "StringBuilder"
                pasteAsButton.TooltipText = "Inserts enquoted clipboard text built up by a stringbuilder."
				'pasteAsButton.FaceId = 22
                AddHandler pasteAsButton.Click, AddressOf PasteAsStringBuilder
                pasteAsButtons.Add(pasteAsButton)

                'add "Region"
                pasteAsButton = CType(pasteAsPopup.Controls.Add(MsoControlType.msoControlButton), CommandBarButton)
                pasteAsButton.Caption = "Region"
                pasteAsButton.TooltipText = "Inserts clipboard wrapped around a specified region"
				'pasteAsButton.FaceId = 22
                AddHandler pasteAsButton.Click, AddressOf PasteAsRegion
                pasteAsButtons.Add(pasteAsButton)

                'add "Configure ..."
                pasteAsButton = CType(pasteAsPopup.Controls.Add(MsoControlType.msoControlButton), CommandBarButton)
                pasteAsButton.Caption = "Configure..."
				'pasteAsButton.FaceId = 548
                pasteAsButton.BeginGroup = True
                AddHandler pasteAsButton.Click, AddressOf PasteAsConfigure
                pasteAsButtons.Add(pasteAsButton)

                'now enable them based on configuration
                EnableContextMenuButtons()

            End If
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.ToString)
        End Try
    End Sub

#Region "PasteAs Handlers"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Occurs when the user clicks the PasteAsString button.
    ''' </summary>
    ''' <param name="ctrl">
    '''     Denotes the CommandBarButton control that initiated the event. 
    ''' </param>
    ''' <param name="cancelDefault">
    '''     False if the default behavior associated with the CommandBarButton control occurs, unless its canceled by another process or add-in. 
    ''' </param>
    ''' -----------------------------------------------------------------------------
    Private Sub PasteAsString(ByVal ctrl As CommandBarButton, ByRef cancelDefault As Boolean)
        smartPaster.PasteAsString(_applicationObject)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Occurs when the user clicks the PasteAsComment button.
    ''' </summary>
    ''' <param name="ctrl">
    '''     Denotes the CommandBarButton control that initiated the event. 
    ''' </param>
    ''' <param name="cancelDefault">
    '''     False if the default behavior associated with the CommandBarButton control occurs, unless its canceled by another process or add-in. 
    ''' </param>
    '''     ''' -----------------------------------------------------------------------------
    Private Sub PasteAsComment(ByVal ctrl As CommandBarButton, ByRef cancelDefault As Boolean)
        smartPaster.PasteAsComment(_applicationObject)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Occurs when the user clicks the PasteAsComment button.
    ''' </summary>
    ''' <param name="ctrl">
    '''     Denotes the CommandBarButton control that initiated the event. 
    ''' </param>
    ''' <param name="cancelDefault">
    '''     False if the default behavior associated with the CommandBarButton control occurs, unless its canceled by another process or add-in. 
    ''' </param>
    ''' -----------------------------------------------------------------------------
    Private Sub PasteAsRegion(ByVal ctrl As CommandBarButton, ByRef cancelDefault As Boolean)
        smartPaster.PasteAsRegion(_applicationObject)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Occurs when the user clicks the PasteAsConfigure button.
    ''' </summary>
    ''' <param name="ctrl">
    '''     Denotes the CommandBarButton control that initiated the event. 
    ''' </param>
    ''' <param name="cancelDefault">
    '''     False if the default behavior associated with the CommandBarButton control occurs, unless its canceled by another process or add-in. 
    ''' </param>
    ''' -----------------------------------------------------------------------------
    Private Sub PasteAsConfigure(ByVal ctrl As CommandBarButton, ByRef cancelDefault As Boolean)
        'show the config form
        Dim myForm As New SmartPasterForm
        myForm.ShowDialog()
        'since config may have changed, show/hide buttons
        EnableContextMenuButtons()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Occurs when the user clicks the PasteAsStringBuilder button.
    ''' </summary>
    ''' <param name="ctrl">
    '''     Denotes the CommandBarButton control that initiated the event. 
    ''' </param>
    ''' <param name="cancelDefault">
    '''     False if the default behavior associated with the CommandBarButton control occurs, unless its canceled by another process or add-in. 
    ''' </param>
    ''' -----------------------------------------------------------------------------
    Private Sub PasteAsStringBuilder(ByVal ctrl As CommandBarButton, ByRef cancelDefault As Boolean)

        smartPaster.PasteAsStringBuilder(_applicationObject)
    End Sub
#End Region

#Region "Exec"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''      Implements the Exec method of the IDTCommandTarget interface.
    '''      This is called when the command is invoked.
    ''' </summary>
    ''' <param name='commandName'>
    '''		The name of the command to execute.
    ''' </param>
    ''' <param name='executeOption'>
    '''		Describes how the command should be run.
    ''' </param>
    ''' <param name='varIn'>
    '''		Parameters passed from the caller to the command handler.
    ''' </param>
    ''' <param name='varOut'>
    '''		Parameters passed from the command handler to the caller.
    ''' </param>
    ''' <param name='handled'>
    '''		Informs the caller if the command was handled or not.
    ''' </param>
    ''' <seealso class='Exec' />
    ''' -----------------------------------------------------------------------------
    Public Sub Exec(ByVal commandName As String, ByVal executeOption As vsCommandExecOption, ByRef varIn As Object, ByRef varOut As Object, ByRef handled As Boolean) Implements IDTCommandTarget.Exec
        handled = False
        If (executeOption = vsCommandExecOption.vsCommandExecOptionDoDefault) Then
            handled = True
            Select Case commandName
                Case "SmartPaster.Connect.Configure"
                    'show the config form
                    Dim myForm As New SmartPasterForm
                    myForm.ShowDialog()
                    'since config may have changed, show/hide buttons
                    EnableContextMenuButtons()

                Case "SmartPaster.Connect.PasteAsComment"
                    smartPaster.PasteAsComment(_applicationObject)
                Case "SmartPaster.Connect.PasteAsString"
                    smartPaster.PasteAsString(_applicationObject)
                Case "SmartPaster.Connect.PasteAsStringBuilder"
                    smartPaster.PasteAsStringBuilder(_applicationObject)
                Case "SmartPaster.Connect.PasteAsRegion"
                    smartPaster.PasteAsRegion(_applicationObject)
                Case Else
                    handled = False
            End Select
        End If
    End Sub
#End Region

#Region "Unimplemented Event Handlers"


    '''<summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
    '''<param name='disconnectMode'>Describes how the Add-in is being unloaded.</param>
    '''<param name='custom'>Array of parameters that are host application specific.</param>
    '''<remarks></remarks>
    Public Sub OnDisconnection(ByVal disconnectMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection
    End Sub

    '''<summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification that the collection of Add-ins has changed.</summary>
    '''<param name='custom'>Array of parameters that are host application specific.</param>
    '''<remarks></remarks>
    Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate
    End Sub

    '''<summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
    '''<param name='custom'>Array of parameters that are host application specific.</param>
    '''<remarks></remarks>
    Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete
    End Sub

    '''<summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
    '''<param name='custom'>Array of parameters that are host application specific.</param>
    '''<remarks></remarks>
    Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown
    End Sub

#End Region

    Public Sub QueryStatus(ByVal commandName As String, ByVal neededText As vsCommandStatusTextWanted, ByRef statusOption As vsCommandStatus, ByRef commandText As Object) Implements IDTCommandTarget.QueryStatus
        If neededText = EnvDTE.vsCommandStatusTextWanted.vsCommandStatusTextWantedNone Then
            If commandName.StartsWith("SmartPaster.Connect") Then
                If (Not (_applicationObject.ActiveDocument Is Nothing)) AndAlso (Not (_applicationObject.ActiveDocument.Object("TextDocument") Is Nothing)) Then
                    statusOption = CType(vsCommandStatus.vsCommandStatusEnabled + vsCommandStatus.vsCommandStatusSupported, vsCommandStatus)
                Else
                    statusOption = vsCommandStatus.vsCommandStatusSupported
                End If
            Else
                statusOption = vsCommandStatus.vsCommandStatusUnsupported
            End If
        End If
    End Sub
End Class
