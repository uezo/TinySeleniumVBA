# TinySeleniumVBA

Um pequeno Selenium wrapper escrito em puro VBA.

[🇬🇧English version is here](https://github.com/uezo/TinySeleniumVBA/blob/main/README.md)

[🇯🇵日本語のREADMEはこちら](https://github.com/uezo/TinySeleniumVBA/blob/main/README.ja.md)

# ✨ Características

- Sem Instalação: Qualquer pessoa mesmo que não tenha permissões de instalação pode começar a automatizar as operações de navegador.
- Inclui métodos úteis: FindElment(s)By*, Get/Set value a um form, click e muito mais.
- Open spec: Basicamente este wrapper é um cliente HTTP de um servidor Webdriver. Aprender sobre este wrapper é o mesmo que aprender sobre Webdriver em geral.
https://www.w3.org/TR/webdriver/


# 📦 Configuração Inicial

1. No editor de VBA em referências selecione: `Microsoft Scripting Runtime`

1. Adicione os módulos`WebDriver.cls`, `WebElement.cls`, `Capabilities.cls` e `JsonConverter.bas` a seu projeto VBA
    - Última versão (v0.1.2): https://github.com/uezo/TinySeleniumVBA/releases/tag/v0.1.2

1. Faça o Download do WebDriver de acordo com o navegador (Aviso: o Webdriver e o navegador devem corresponder a mesma versão)
    - Edge: https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/
    - Chrome: https://chromedriver.chromium.org/downloads

# 🪄 Exemplo de uso

```vb
Public Sub main()
    ' Start WebDriver (Edge)
    Dim Driver As New WebDriver
    Driver.Edge "path\to\msedgedriver.exe"
    
    ' Open browser
    Driver.OpenBrowser
    
    ' Navigate to Google
    Driver.Navigate "https://www.google.co.jp/?q=selenium"

    ' Get search textbox
    Dim searchInput
    Set searchInput = Driver.FindElement(By.Name, "q")
    
    ' Get value from textbox
    Debug.Print searchInput.GetValue
    
    ' Set value to textbox
    searchInput.SetValue "yomoda soba"
    
    ' Click search button
    Driver.FindElement(By.Name, "btnK").Click
    
    ' Refresh - you can use Execute with driver command even if the method is not provided
    Driver.Execute Driver.CMD_REFRESH
End Sub
```

# 🐙 BrowserOptions

Utilize `Capabilities` para configurar as opções do navegador. Este é um exemplo para lançar o browser como modo sem cabeça (invisível).

```vb
' Start web driver
Dim Driver As New WebDriver
Driver.Chrome "C:\Users\uezo\Desktop\chromedriver.exe"

' Configure Capabilities
Dim cap As Capabilities
Set cap = Driver.CreateCapabilities()
cap.AddArgument "--headless"
' Use SetArguments if you want to add multiple arguments
' cap.SetArguments "--headless --xxx -xxx"

' Show Capabilities as JSON for debugging
Debug.Print cap.ToJson()

' Open browser
Driver.OpenBrowser cap
```

Ver também os sites abaixo para compreender as especificações de `Capabilities` para cada navegador.
- Chrome: https://chromedriver.chromium.org/capabilities
- Edge: https://docs.microsoft.com/en-us/microsoft-edge/webdriver-chromium/capabilities-edge-options


# ⚡️ Execute JavaScript

Utilize `ExecuteScript()` para executar JavaScript no browser.

```vb
' Start web driver
Dim Driver As New WebDriver
Driver.Chrome "C:\Users\uezo\Desktop\chromedriver.exe"

' Open browser
Driver.OpenBrowser

' Navigate to Google
Driver.Navigate "https://www.google.co.jp/?q=liella"

' Show alert
Driver.ExecuteScript "alert('Hello TinySeleniumVBA')"

' === Use breakpoint to CLOSE ALERT before continue ===

' Pass argument
Driver.ExecuteScript "alert('Hello ' + arguments[0] + ' as argument')", Array("TinySeleniumVBA")

' === Use breakpoint to CLOSE ALERT before continue ===

' Pass element as argument
Dim searchInput
Set searchInput = Driver.FindElement(By.Name, "q")
Driver.ExecuteScript "alert('Hello ' + arguments[0].value + ' ' + arguments[1])", Array(searchInput, "TinySeleniumVBA")

' === CLOSE ALERT and continue ===

' Get return value from script
Dim retStr As String
retStr = Driver.ExecuteScript("return 'Value from script'")
Debug.Print retStr

' Get WebElement as return value from script
Dim firstDiv As WebElement
Set firstDiv = Driver.ExecuteScript("return document.getElementsByTagName('div')[0]")
Debug.Print firstDiv.GetText()

' Get complex structure as return value from script
Dim retArray
retArray = Driver.ExecuteScript("return [['a', '1'], {'key1': 'val1', 'key2': document.getElementsByTagName('div'), 'key3': 'val3'}]")

Debug.Print retArray(0)(0)  ' a
Debug.Print retArray(0)(1)  ' 1

Debug.Print retArray(1)("key1") ' val1
Debug.Print retArray(1)("key2")(0).GetText()    ' Inner Text
Debug.Print retArray(1)("key2")(1).GetText()    ' Inner Text
Debug.Print retArray(1)("key3") ' val3
```

# ❤️ Agradecimentos

[VBA-JSON](https://github.com/VBA-tools/VBA-JSON) de Tim Hall, um conversor de JSON para VBA que auxilia muito ao fazer um cliente HTTP. Esta valiosa biblioteca está inclusa nesta versão junto com sua respectiva licença. Muito obrigado!
