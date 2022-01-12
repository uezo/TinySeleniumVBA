VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Capabilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' TinySeleniumVBA v0.1.2
' A tiny Selenium wrapper written in pure VBA
'
' (c)2021 uezo
'
' Mail: uezo@uezo.net
' Twitter: @uezochan
' https://github.com/uezo/TinySeleniumVBA
'
' ==========================================================================
' MIT License
'
' Copyright (c) 2021 uezo
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ==========================================================================

Option Explicit

Public Root As New Dictionary
Private BrowserOptionKey As String

' Spec of capabilities:
' - Chrome: https://chromedriver.chromium.org/capabilities
' - Edge: https://docs.microsoft.com/en-us/microsoft-edge/webdriver-chromium/capabilities-edge-options

' Getters and Setters
Public Property Let args(value() As String)
    SetOption "args", value
End Property

Public Property Get args() As String()
    args = GetOption("args")
End Property

Public Property Let binary(ByVal value As String)
    SetOption "binary", value
End Property

Public Property Get binary() As String
    binary = GetOption("binary")
End Property

Public Property Let extensions(value() As String)
    SetOption "extensions", value
End Property

Public Property Get extensions() As String()
    extensions = GetOption("extensions")
End Property

Public Property Set localState(ByVal value As Dictionary)
    SetOption "localState", value
End Property

Public Property Get localState() As Dictionary
    Set localState = GetOption("localState")
End Property

Public Property Set prefs(value As Dictionary)
    SetOption "prefs", value
End Property

Public Property Get prefs() As Dictionary
    Set prefs = GetOption("prefs")
End Property

Public Property Let detach(value As Boolean)
    SetOption "detach", value
End Property

Public Property Get detach() As Boolean
    detach = GetOption("detach")
End Property

Public Property Let debuggerAddress(ByVal value As String)
    SetOption "debuggerAddress", value
End Property

Public Property Get debuggerAddress() As String
    debuggerAddress = GetOption("debuggerAddress")
End Property

Public Property Let excludeSwitches(value() As String)
    SetOption "excludeSwitches", value
End Property

Public Property Get excludeSwitches() As String()
    excludeSwitches = GetOption("excludeSwitches")
End Property

Public Property Let minidumpPath(ByVal value As String)
    SetOption "minidumpPath", value
End Property

Public Property Get minidumpPath() As String
    minidumpPath = GetOption("minidumpPath")
End Property

Public Property Set mobileEmulation(value As Dictionary)
    SetOption "mobileEmulation", value
End Property

Public Property Get mobileEmulation() As Dictionary
    Set mobileEmulation = GetOption("mobileEmulation")
End Property

Public Property Set perfLoggingPrefs(value As Dictionary)
    SetOption "perfLoggingPrefs", value
End Property

Public Property Get perfLoggingPrefs() As Dictionary
    Set perfLoggingPrefs = GetOption("perfLoggingPrefs")
End Property

Public Property Let windowTypes(ByVal value As String)
    SetOption "windowTypes", value
End Property

Public Property Get windowTypes() As String
    windowTypes = GetOption("windowTypes")
End Property

Public Sub SetOption(ByVal key As String, value As Variant)
    If IsObject(value) Then
        Set Root("alwaysMatch")(BrowserOptionKey)(key) = value
    Else
        Root("alwaysMatch")(BrowserOptionKey)(key) = value
    End If
End Sub

Public Function GetOption(ByVal key As String) As Variant
    If IsObject(Root("alwaysMatch")(BrowserOptionKey)(key)) Then
        Set GetOption = Root("alwaysMatch")(BrowserOptionKey)(key)
    Else
        GetOption = Root("alwaysMatch")(BrowserOptionKey)(key)
    End If
End Function

' Helper method to add argument
Public Sub AddArgument(ByVal argument As String)
    Dim arguments() As String
    Dim idx As Integer: idx = -1

    ' idx will not be updated when args is Nothing or args has no items
    On Error Resume Next
    idx = UBound(args)
    On Error GoTo 0
    
    If idx >= 0 Then
        arguments = args
    End If
        
    ReDim Preserve arguments(idx + 1)
    arguments(UBound(arguments)) = argument
    
    args = arguments
End Sub

' Helper method to set arguments
Public Sub SetArguments(ByVal value As String, Optional ByVal delimiter As String = " ")
    SetOption "args", Split(value, delimiter)
End Sub

' Helper method to add prefs
Public Sub AddPref(ByVal key As String, ByVal value As String)
    If prefs Is Nothing Then
        SetOption "prefs", New Dictionary
    End If
    prefs(key) = value
End Sub

' Convert to JSON string for debugging
Public Function ToJson() As String
    ToJson = JsonConverter.ConvertToJson(Root)
End Function

' Initializer for each browser
Public Sub InitializeFor(ByVal BrowserName As String, Optional ByVal optionKey As String)
    Set Root = New Dictionary
    Root.Add "browserName", BrowserName
    Root.Add "alwaysMatch", New Dictionary
    
    ' Add browser option keys for chromium like browsers. If value is not set, the option will not appear in JSON
    Dim browserOptions As New Dictionary
    browserOptions.Add "args", Nothing
    browserOptions.Add "binary", Nothing
    browserOptions.Add "extensions", Nothing
    browserOptions.Add "localState", Nothing
    browserOptions.Add "prefs", Nothing
    browserOptions.Add "detach", Nothing
    browserOptions.Add "debuggerAddress", Nothing
    browserOptions.Add "excludeSwitches", Nothing
    browserOptions.Add "minidumpPath", Nothing
    browserOptions.Add "mobileEmulation", Nothing
    browserOptions.Add "perfLoggingPrefs", Nothing
    browserOptions.Add "windowTypes", Nothing
    
    Select Case LCase(BrowserName)
        Case "chrome"
            BrowserOptionKey = "goog:chromeOptions"

        Case "microsoftedge"
            BrowserOptionKey = "ms:edgeOptions"

        Case Else
            BrowserOptionKey = optionKey
        
        'TODO: Add default options for browsers
    End Select
    
    Root("alwaysMatch").Add BrowserOptionKey, browserOptions
End Sub

' Shortcut initializer for chrome
Public Sub InitializeForChrome()
    InitializeFor "chrome"
End Sub

' Shortcut initializer for edge
Public Sub InitializeForEdge()
    InitializeFor "MicrosoftEdge"
End Sub
