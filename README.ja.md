# TinySeleniumVBA

A tiny Selenium wrapper written in pure VBA.

[🇬🇧English version is here](https://github.com/uezo/TinySeleniumVBA/blob/main/README.md) : 

[🇧🇷Versão em Português](https://github.com/tdmsoares/TinySeleniumVBA/blob/ReadmeInPortuguese/README.pt.md)


# ✨ 特長

- インストール不要: VBAのみで書かれているので、インストール権限のない人でもすぐにブラウザ自動操作に取り掛かることができます
- 便利なヘルパー機能: FindElement(s)By*、フォームへの値の入出力、クリックやその他便利な機能を提供しています
- オープンな仕様: 基本的にこのラッパーはWebDriverのHTTPクライアントですので、ラッパーの使い方を学ぶことはWebDriverの仕様を知ることと同義です。無駄になるものはありません
https://www.w3.org/TR/webdriver/


# 📦 セットアップ

1. ツール＞参照設定から `Microsoft Scripting Runtime` に参照を通してください

1. `WebDriver.cls`、`WebElement.cls`、`Capabilities.cls`、`JsonConverter.bas`をプロジェクトに追加してください
    - 最新版 (v0.1.2): https://github.com/uezo/TinySeleniumVBA/releases/tag/v0.1.2

1. WebDriverをダウンロードしてください（ブラウザのメジャーバージョンと同じもの）
    - Edge: https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/
    - Chrome: https://chromedriver.chromium.org/downloads

# 🪄 使い方

```vb
Public Sub main()
    ' WebDriverの開始 (Edge)
    Dim Driver As New WebDriver
    Driver.Edge "path\to\msedgedriver.exe"
    
    ' ブラウザを開く
    Driver.OpenBrowser
    
    ' Googleへ移動
    Driver.Navigate "https://www.google.co.jp/?q=selenium"

    ' 検索テキストボックスを取得
    Dim searchInput
    Set searchInput = Driver.FindElement(By.Name, "q")
    
    ' テキストボックスの値を取得
    Debug.Print searchInput.GetValue
    
    ' テキストボックスに値を入力
    searchInput.SetValue "yomoda soba"
    
    ' 検索ボタンのクリック
    Driver.FindElement(By.Name, "btnK").Click
    
    ' 再読み込み - ヘルパーメソッドを提供していない場合でも、ドライバーコマンドを直接実行することができます
    Driver.Execute Driver.CMD_REFRESH
End Sub
```

# 🐙 ブラウザーオプション

ブラウザーオプションを指定するには `Capabilities` を使うと便利です。以下はヘッドレスモード（非表示モード）でブラウザを起動する例です。

```vb
' Start web driver
Dim Driver As New WebDriver
Driver.Chrome "C:\path\to\chromedriver.exe"

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

`Capabilities`の仕様はブラウザ毎に異なりますので、以下のWebサイト等にてご確認ください。
- Chrome: https://chromedriver.chromium.org/capabilities
- Edge: https://docs.microsoft.com/en-us/microsoft-edge/webdriver-chromium/capabilities-edge-options


# ⚡️ JavaScriptの実行

`ExecuteScript()`を利用することで、ブラウザ上で任意のJavaScriptコードを実行することができます。

```vb
' Start web driver
Dim Driver As New WebDriver
Driver.Chrome "C:\path\to\chromedriver.exe"

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

# ❤️ 謝辞

[VBA-JSON](https://github.com/VBA-tools/VBA-JSON) という Tim Hall さんが開発したVBA用JSONコンバーターはHTTPクライアントを作る上でとても役に立ちました。このすばらしいライブラリは当該ライブラリのライセンスのもとでリリースに含まれています。ありがとうございます！
