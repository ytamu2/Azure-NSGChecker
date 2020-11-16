# Azure-NSGChecker
  
# 概要
Excelのシートを読込み、1行ずつNetwork WatcherのIPフローでNSGの通信チェックを行います。  
PowerShell 7で動作します。
  
# 準備
- NSGの疎通チェックをするための各値を記述したエクセルファイルを用意します。  
    <img src="https://github.com/ytamu2/Azure-NSGChecker/blob/images/images/Excel_sample.png" width=80%>
- NSGChackConfig.jsonに各種情報を記述します。（ファイル名は変更可能です）
    - エクセルの値が記述してる列や、AzureのサブスクリプションID等
- フォルダ構成は以下のようにして、各ファイルを配置します。  
    ├───bin  
    │       AzureNSGCheck.ps1  
    │       LogController.ps1  
    │  
    ├───etc  
    │       NSGCheckConfig.json  
    │  
    └───log
  
# 使い方
1. PowerShellでAzureにログインします。
    ```
    Login-AzAccount
    ```
2. `AzureNSGCheck.ps1`を実行します。
3. ダイアログが表示されるので、読込むエクセルファイルを選択します。
    <img src="https://github.com/ytamu2/Azure-NSGChecker/blob/images/images/opnefiledialog.png" width=90%>
4. `NSGChackConfigjson`内のOutput->Directoryに指定したディレクトリに通信結果がcsvファイルで出力されます。
    <img src="https://github.com/ytamu2/Azure-NSGChecker/blob/images/images/json-outputdirecroty.png" width=60%>
