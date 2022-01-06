VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Capabilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'References: Microsoft Scripting Runtime

'This class was written by GCUser
'This is a first pass at a Capabilities class for TinySeleniumVBA
'This was written to be exposed by the WebDriver class
'Add/modify following lines of code in WebDriver class:
'
'Private objCapabilities As Capabilities 'module level declare
'
'Public Function Capabilities() As Capabilities 'this exposes the Capability object
'    Set Capabilities = objCapabilities
'End Function
'
'Private Sub Class_Initialize()
'    Set objCapabilities = New Capabilities
'End Sub
'
'Add following lines to the Edge and Chrome methods, respectively:
'objCapabilities.BrowserName = "msedge"
'objCapabilities.BrowserName = "chrome"

'Modify the following line in OpenBrowser method:
'Set resp = Execute(CMD_NEW_SESSION, Params("capabilities", objCapabilities.Dictionary))
'
'Usage:
'Capabilities must be defined prior to Driver.OpenBrowser call as in the following usage:
'
'Driver.Edge
'Driver.Capabilities.Add "acceptInsecureCerts", False
'Driver.Capabilities.AddArgument "--headless"
'Driver.Capabilities.AddArgument "--window-size=1080,400"
'Driver.Capabilities.AddPref "download.default_directory", "C:\Users\....\download"
'Driver.Capabilities.AddPref "download.directory_upgrade", True
'Driver.Capabilities.AddPref "download.prompt_for_download", False
'Driver.Capabilities.AddPref "safebrowsing.enabled", True
'Debug.Print Driver.Capabilities.JsonString
'Driver.OpenBrowser
'Driver.Navigate "https://www.google.co.jp/?q=selenium"
'
'The following is printed in the immediate window:
'{
'    "alwaysMatch": {
'        "acceptInsecureCerts": false,
'        "ms:edgeOptions": {
'            "args": [
'                "--headless",
'                "--window-size=1080,400"
'            ],
'            "prefs": {
'                "download.default_directory": "C:\\Users\\....\\download",
'                "download.directory_upgrade": true,
'                "download.prompt_for_download": false,
'                "safebrowsing.enabled": true
'            }
'        }
'    }
'}

Private extensions() As Variant
Private args() As Variant
Private prefsDict As Dictionary
Private optionsDict As Dictionary
Private mBrowserName As String
Private alwaysMatchDict As Dictionary

'BrowserName passed by Webdriver.Edge or Webdriver.Chrome
Friend Property Let BrowserName(ByVal str As String)
    mBrowserName = str
End Property

Public Sub AddArgument(ByVal argumentstring As String)
    argumentstring = Trim(argumentstring)
    If Left(argumentstring, 2) <> "--" Then argumentstring = "--" & argumentstring
    If UBound(args) > -1 Then
        If Not IsError(Application.match(argumentstring, args, 0)) Then Exit Sub
    End If
    ReDim Preserve args(UBound(args) + 1)
    args(UBound(args)) = argumentstring
End Sub

Public Sub AddOption(ByVal key As String, ByVal data As Variant)
    If Not optionsDict.Exists(key) Then
        optionsDict.Add key, data
    Else
        optionsDict(key) = data
    End If
End Sub

Public Sub AddPref(ByVal key As String, ByVal data As Variant)
    If Not prefsDict.Exists(key) Then
        prefsDict.Add key, data
    Else
        prefsDict(key) = data
    End If
End Sub

Public Sub Add(ByVal key As String, ByVal data As Variant)
    If Not alwaysMatchDict.Exists(key) Then
        alwaysMatchDict.Add key, data
    Else
        alwaysMatchDict(key) = data
    End If
End Sub

Public Sub AddExtension(ByVal path As String)
    If UBound(extensions) > -1 Then
        If Not IsError(Application.match(path, extensions, 0)) Then Exit Sub
    End If
    ReDim Preserve extensions(UBound(extensions) + 1)
    extensions(UBound(extensions)) = path
End Sub

Private Function BuildCapabilities() As Dictionary
    Dim optionsCloneDict As Dictionary, alwaysMatchCloneDict As Dictionary
    Dim browserOptions_key As String

    Select Case mBrowserName
    Case "msedge"
        browserOptions_key = "ms:edgeOptions"
    Case "chrome"
        browserOptions_key = "goog:chromeOptions"
    Case Else
        MsgBox "something went wrong"
        Exit Function
    End Select

    Set optionsCloneDict = CloneDictionary(optionsDict)
    Set alwaysMatchCloneDict = CloneDictionary(alwaysMatchDict)
    
    If UBound(extensions) > -1 Then optionsCloneDict("extensions") = extensions
    If UBound(args) > -1 Then optionsCloneDict("args") = args
    
    If prefsDict.Count > 0 Then
        optionsCloneDict.Add "prefs", prefsDict
    End If
    
    If optionsCloneDict.Count > 0 Then
        alwaysMatchCloneDict.Add browserOptions_key, optionsCloneDict
    End If
    
    Set BuildCapabilities = New Dictionary
    If alwaysMatchCloneDict.Count > 0 Then
        BuildCapabilities.Add "alwaysMatch", alwaysMatchCloneDict
    End If
    
End Function

Public Function JsonString() As String
    JsonString = JsonConverter.ConvertToJson(BuildCapabilities, 4)
End Function

Public Function Dictionary() As Dictionary
    Set Dictionary = BuildCapabilities
End Function

Private Function CloneDictionary(dict As Dictionary) As Dictionary
    Dim newDict As Dictionary
    Set newDict = New Dictionary
    For Each key In dict.Keys
        newDict.Add key, dict(key)
    Next key
    newDict.CompareMode = dict.CompareMode
    Set CloneDictionary = newDict
End Function

Private Sub Class_Initialize()
    extensions = Array()
    args = Array()
    Set optionsDict = New Dictionary
    Set prefsDict = New Dictionary
    Set alwaysMatchDict = New Dictionary
End Sub

Sub Clear()
    Set optionsDict = Nothing
    Set optionsDict = New Dictionary
    Set prefsDict = Nothing
    Set prefsDict = New Dictionary
    Set alwaysMatchDict = Nothing
    Set alwaysMatchDict = New Dictionary
    extensions = Array()
    args = Array()
End Sub

Private Sub Class_Terminate()
    Set optionsDict = Nothing
    Set alwaysMatchDict = Nothing
    Set prefsDict = Nothing
End Sub
