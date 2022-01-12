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

# ❤️ Agradecimentos

[VBA-JSON](https://github.com/VBA-tools/VBA-JSON) de Tim Hall, um conversor de JSON para VBA que auxilia muito ao fazer um cliente HTTP. Esta valiosa biblioteca está inclusa nesta versão junto com sua respectiva licença. Muito obrigado!
