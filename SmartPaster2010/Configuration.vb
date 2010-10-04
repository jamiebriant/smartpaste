''' -----------------------------------------------------------------------------
''' Project	 : SmartPaster
''' Class	 : Configuration
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''   Singleton class for storing configuration options for the application.
''' </summary>
''' <remarks>
'''   Based somewhat on the Singleton design pattern.
''' </remarks>
''' -----------------------------------------------------------------------------
Friend NotInheritable Class Configuration

    'Singleton pattern: private constructor and private instance
    Private Shared ReadOnly Instance As Configuration = New Configuration
    Private Sub New()
        InitializeRegistryStore()
    End Sub

#Region "Shared Settings Properties"

#Region "Paste As ..."
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Specifies whether the option will be displayed on the context menu
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property EnablePasteAsComment() As Boolean
        Get
            Return (Instance.GetSetting("EnablePasteAsComment", [Boolean].TrueString) = [Boolean].TrueString)
        End Get
        Set(ByVal Value As Boolean)
            Instance.SetSetting("EnablePasteAsComment", Value.ToString)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Specifies whether the option will be displayed on the context menu
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property EnablePasteAsString() As Boolean
        Get
            Return (Instance.GetSetting("EnablePasteAsString", [Boolean].TrueString) = [Boolean].TrueString)
        End Get
        Set(ByVal Value As Boolean)
            Instance.SetSetting("EnablePasteAsString", Value.ToString)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Specifies whether the option will be displayed on the context menu
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property EnablePasteAsStringBuilder() As Boolean
        Get
            Return (Instance.GetSetting("EnablePasteAsStringBuilder", [Boolean].TrueString) = [Boolean].TrueString)
        End Get
        Set(ByVal Value As Boolean)
            Instance.SetSetting("EnablePasteAsStringBuilder", Value.ToString)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Specifies whether the option will be displayed on the context menu
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property EnablePasteAsRegion() As Boolean
        Get
            Return (Instance.GetSetting("EnablePasteAsRegion", [Boolean].TrueString) = [Boolean].TrueString)
        End Get
        Set(ByVal Value As Boolean)
            Instance.SetSetting("EnablePasteAsRegion", Value.ToString)
        End Set
    End Property
#End Region

#Region "VB"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Indicates if Tabs are escaped or not
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property VBEscapeTabs() As Boolean
        Get
            Return (Instance.GetSetting("VBEscapeTabs", [Boolean].TrueString) = [Boolean].TrueString)
        End Get
        Set(ByVal Value As Boolean)
            Instance.SetSetting("VBEscapeTabs", Value.ToString)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Indicates if Tabs are escaped or kept in the string
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property VBEscapeLinebreaks() As Boolean
        Get
            Return (Instance.GetSetting("VBEscapeLinebreaks", [Boolean].TrueString) = [Boolean].TrueString)
        End Get
        Set(ByVal Value As Boolean)
            Instance.SetSetting("VBEscapeLinebreaks", Value.ToString)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   String to use in Visual Basic for New Lines.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property VBNewLine() As String
        Get
            Return Instance.GetSetting("VBNewLine", "Environment.NewLine")
        End Get
        Set(ByVal Value As String)
            Instance.SetSetting("VBNewLine", Value)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   String to use in Visual Basic for Tabs.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property VBTab() As String
        Get
            Return Instance.GetSetting("VBTab", "ControlChars.Tab")
        End Get
        Set(ByVal Value As String)
            Instance.SetSetting("VBTab", Value)
        End Set
    End Property
#End Region

#Region "C#"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Indicates if Tabs are escaped or not
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property CSEscapeTabs() As Boolean
        Get
            Return (Instance.GetSetting("CSEscapeTabs", [Boolean].TrueString) = [Boolean].TrueString)
        End Get
        Set(ByVal Value As Boolean)
            Instance.SetSetting("CSEscapeTabs", Value.ToString)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Indicates if Tabs are escaped or kept in the string
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property CSEscapeLinebreaks() As Boolean
        Get
            Return (Instance.GetSetting("CSEscapeLinebreaks", [Boolean].TrueString) = [Boolean].TrueString)
        End Get
        Set(ByVal Value As Boolean)
            Instance.SetSetting("CSEscapeLinebreaks", Value.ToString)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   String to use in for new Lines.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property CSNewLine() As String
        Get
            Return Instance.GetSetting("CSNewLine", "Environment.NewLine")
        End Get
        Set(ByVal Value As String)
            Instance.SetSetting("CSNewLine", Value)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   String to use in tabs.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property CSTab() As String
        Get
            Return Instance.GetSetting("CSTab", "ControlChars.Tab")
        End Get
        Set(ByVal Value As String)
            Instance.SetSetting("CSTab", Value)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Indicates if Verbatim Literals should be used isntead of escapes
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property CSVerbatimLiterals() As Boolean
        Get
            Return (Instance.GetSetting("CSVerbatimLiterals", [Boolean].TrueString) = [Boolean].TrueString)
        End Get
        Set(ByVal Value As Boolean)
            Instance.SetSetting("CSVerbatimLiterals", Value.ToString)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Indicates if Verbatim Literals should be allowed to span multiple lines
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property CSVerbatimLiteralsSpanLines() As Boolean
        Get
            Return (Instance.GetSetting("CSVerbatimLiteralsSpanLines", [Boolean].FalseString) = [Boolean].TrueString)
        End Get
        Set(ByVal Value As Boolean)
            Instance.SetSetting("CSVerbatimLiteralsSpanLines", Value.ToString)
        End Set
    End Property

#End Region
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Specifies whether a StringBuilder Paste will use the Format method for
    '''  escaped characters
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property AppendFormatOnStringBuilder() As Boolean
        Get
            Return (Instance.GetSetting("AppendFormatOnStringBuilder", [Boolean].TrueString) = [Boolean].TrueString)
        End Get
        Set(ByVal Value As Boolean)
            Instance.SetSetting("AppendFormatOnStringBuilder", Value.ToString)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Specifies whether text will be wrapped on paste
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property EnableWordWrap() As Boolean
        Get
            Return (Instance.GetSetting("EnableWordWrap", [Boolean].TrueString) = [Boolean].TrueString)
        End Get
        Set(ByVal Value As Boolean)
            Instance.SetSetting("EnableWordWrap", Value.ToString)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Indicates the approximate image length at which to wrap lines
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property WrapAt() As Integer
        Get
            Dim ret As Integer
            Try
                ret = CInt(Instance.GetSetting("WrapAt", "100"))
            Catch ex As Exception
                ret = 100
                Instance.SetSetting("WrapAt", "100")
            End Try
            If (ret > 0) And (ret < 5) Then
                Instance.SetSetting("WrapAt", "100")
                ret = 100
            End If

            Return ret
        End Get
        Set(ByVal Value As Integer)
            Instance.SetSetting("WrapAt", Value.ToString)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Specifies the default language to use when filetype cannot be determined.
    '''   Accepted values are "Visual Basic" or "C#".
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property DefaultLanguage() As String
        Get
            Return Instance.GetSetting("DefaultLanguage", "C#")
        End Get
        Set(ByVal Value As String)
            Instance.SetSetting("DefaultLanguage", Value)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Specifies whether text will be autoformatted after paste
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property AutoFormatAfterPaste() As Boolean
        Get
            Return (Instance.GetSetting("AutoFormatAfterPaste", [Boolean].TrueString) = [Boolean].TrueString)
        End Get
        Set(ByVal Value As Boolean)
            Instance.SetSetting("AutoFormatAfterPaste", Value.ToString)
        End Set
    End Property







    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Specifies whether tabs should be escaped as "\t" or "ControlChars.Tab"
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Property EscapeTabs() As Boolean
        Get
            Return (Instance.GetSetting("EscapeTabs", [Boolean].TrueString) = [Boolean].TrueString)
        End Get
        Set(ByVal Value As Boolean)
            Instance.SetSetting("EscapeTabs", Value.ToString)
        End Set
    End Property
#End Region

#Region "Registry Stuff"

    Protected SettingsCache As System.Collections.Specialized.StringDictionary
    Protected regKey As Microsoft.Win32.RegistryKey

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Initializes the cache and registry key to use.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected Sub InitializeRegistryStore()
        Dim subkey As String = "Software\SmartPaster"

        SettingsCache = New System.Collections.Specialized.StringDictionary
        regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(subkey, True)

        ' create if it did not exist.
        If regKey Is Nothing Then
            regKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(subkey)

        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Looks for the key in the cache and returns the value if found. Otherwise,
    '''  looks in the registry and adds the value to the cache and returns the value.
    '''  If the key isn't in the registry, adds the default value to the registry and
    '''  cache, and returns the default value.
    ''' </summary>
    ''' <param name="key">key for the value to retrive</param>
    ''' <param name="defaultValue">default value of the setting to retrive</param>
    ''' <returns>value retrieved from registry or default value</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function GetSetting(ByVal key As String, ByVal defaultValue As String) As String
        Dim value As String

        If Not SettingsCache.ContainsKey(key) Then
            ' not in cache, so look in registry
            Dim obj As Object = regKey.GetValue(key)

            If obj Is Nothing Then
                ' save default to registry
                regKey.SetValue(key, defaultValue)
                value = defaultValue
            Else
                value = CStr(obj)
            End If

            ' save to cache for next time
            SettingsCache.Add(key, value)
        Else
            ' get from cache
            value = SettingsCache.Item(key)
        End If

        Return value
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Sets the value in the registry and cache
    ''' </summary>
    ''' <param name="key">key value to set</param>
    ''' <param name="value">value to store</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[alex]	5/20/2004	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Protected Sub SetSetting(ByVal key As String, ByVal value As String)
        ' save to cache and add to reg
        SettingsCache.Item(key) = value
        regKey.SetValue(key, value)
    End Sub


#End Region

End Class
