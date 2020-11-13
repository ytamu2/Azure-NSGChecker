<################################################################################
## @author:Yasutoshi Tamura
## @name: AzureNSGCheck.ps1
## @summary: Network WatcherのIPフローの確認を使用してNSGの疎通確認
## @since:2020/11/13
## @version:1.0
## @see:
## @parameter
##  1:ConfigFileName：パラメータファイル名
## @return: 0:Success 1:Error 99:Exception
################################################################################>

param (
    [parameter()][string]$ConfigFileName = "NSGCheckConfig.json"
)

# シートの読込
function InvokeCheckNSG {
    param(
        $returnCode,
        $Log,
        $Sheet,
        $ExcelSheetjson,
        $NetworkWatcher,
        $OutDirectory
    )
    
    $Column = $ExcelSheetjson.Column
    $Row = $ExcelSheetjson.Row
    $Cells = $Sheet.Cells

    $Log.info("シート「$($Sheet.Name)」の読込を開始します。")

    $MaxRow = $Cells.Item($Sheet.Rows.Count, $Column.VMName).End(-4162).Row
    $Log.info("対象レコードは$($Row.Start)～$($MaxRow)行目です。")

    $ExcelRange = $Cells.Range($Cells.Item(1, 1), $Cells.Item($MaxRow, $Column.Max)).value()
    $OutputNSGCheck = @("Access,Rule,VMName,Direction,Protocol,LocalIP,LocalPort,RemoteIP,RemotePort")
    $resultCheck = $resturnCode.Success
    foreach ($eachRow in $($Row.Start)..($MaxRow)) {
        # ブランク項目存在チェック
        $blankCol = $ExcelRange[$eachRow, $Column.VMName]
        $blankCol = $blankCol -and $ExcelRange[$eachRow, $Column.ResourceGroup]
        $blankCol = $blankCol -and $ExcelRange[$eachRow, $Column.Direction]
        $blankCol = $blankCol -and $ExcelRange[$eachRow, $Column.Protocol]
        $blankCol = $blankCol -and $ExcelRange[$eachRow, $Column.LocalPort]
        $blankCol = $blankCol -and $ExcelRange[$eachRow, $Column.RemoteIPAddress]
        $blankCol = $blankCol -and $ExcelRange[$eachRow, $Column.RemotePort]

        if (!$blankCol) {
            $Log.Warn("空白の項目が存在するため、次の行にスキップします。")
            continue
        }
        
        # VM情報の取得
        if ($VirtualMachine.Name -ne $ExcelRange[$eachRow, $Column.VMName]) {
            $VirtualMachine = $null
            $VirtualMachine = Get-AzVM -ResourceGroupName $ExcelRange[$eachRow, $Column.ResourceGroup] -Name $ExcelRange[$eachRow, $Column.VMName]

            if (!$VirtualMachine) {
                $Log.warn("仮想マシン「$($ExcelRange[$eachRow, $Column.VMName])」の取得に失敗しました。")
                $resultCheck = $returnCode.Error
                continue
            }

            # VMステータス取得
            $VMStatus = ($VirtualMachine | Get-AzVM -Status).Statuses | where { $_.Code -match "PowerState" }
        }

        if ($VMStatus.Code -ne "PowerState/running") {
            $Log.Warn("仮想マシン「$($ExcelRange[$eachRow, $Column.VMName])」を起動してください。現在のステータスは「$($VMStatus.Code)」です。")
            $resultCheck = $returnCode.Error
            continue
        }

        # プライベートIPアドレス（NIC）情報の取得
        $NetworkInterface = $null
        $NetworkInterface = Get-AzNetworkInterface -ResourceGroupName $VirtualMachine.ResourceGroupName | where { $_.VirtualMachine.Id -eq $VirtualMachine.Id }
        $LocalIP = $NetworkInterface.IpConfigurations.PrivateIpAddress
        
        $Log.info("測定内容： 仮想マシン：$($VirtualMachine.Name) 方向：$($ExcelRange[$eachRow, $Column.Direction]) プロトコル：$($ExcelRange[$eachRow, $Column.Protocol]) ローカルIP：${LocalIP} ローカルポート：$($ExcelRange[$eachRow, $Column.LocalPort]) リモートIP：$($ExcelRange[$eachRow, $Column.RemoteIPAddress]) リモートポート：$($ExcelRange[$eachRow, $Column.RemotePort])")

        # ローカルIPとリモートIPの同一チェック
        if ($LocalIP -eq $ExcelRange[$eachRow, $Column.RemoteIPAddress]) {
            $Log.Warn("リモートIPがローカルIPと同一のため、次の行にスキップします。")
            $resultCheck = $returnCode.Error
            continue
        }

        # IPFlow実行
        $resultNWIPFlow = $null
        $resultNWIPFlow = Test-AzNetworkWatcherIPFlow -NetworkWatcher $NetworkWatcher -TargetVirtualMachineId $VirtualMachine.Id -Direction $ExcelRange[$eachRow, $Column.Direction] -Protocol $ExcelRange[$eachRow, $Column.Protocol] -LocalIPAddress $LocalIP -RemoteIPAddress $ExcelRange[$eachRow, $Column.RemoteIPAddress] -LocalPort $ExcelRange[$eachRow, $Column.LocalPort] -RemotePort $ExcelRange[$eachRow, $Column.RemotePort]
        
        if (!$resultNWIPFlow) {
            $Log.Error("NSGチェック処理に失敗しました。")
            $resultCheck = $returnCode.Error
            continue
        }

        $Log.Info("結果： アクセス：$($resultNWIPFlow.Access) ルール：$($resultNWIPFlow.RuleName)")
        $OutputNSGCheck += "$($resultNWIPFlow.Access),$($resultNWIPFlow.RuleName),$($VirtualMachine.Name),$($ExcelRange[$eachRow, $Column.Direction]),$($ExcelRange[$eachRow, $Column.Protocol]),$($LocalIP),$($ExcelRange[$eachRow, $Column.LocalPort]),$($ExcelRange[$eachRow, $Column.RemoteIPAddress]),$($ExcelRange[$eachRow, $Column.RemotePort])"

    }

    $Log.info("シート「$($Sheet.Name)」の読込が完了しました。")

    $OutFile = "AzureNSGCheck_$(Get-Date -Format 'yyyyMMddHHmmss')_$($Sheet.Name).csv"
    $OutputNSGCheck | Out-File -FilePath (Join-Path $OutDirectory -ChildPath $OutFile) -Encoding utf8

    return $resultCheck 
}

# スクリプト格納ディレクトリを取得
$scriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent

# モジュールのロード
. (Join-Path $scriptDir -childPath "LogController.ps1")

# エラー、リターンコード設定
$error.Clear()
Set-Variable -Name returnCode -Value @{Success = 0; Error = 1; Exception = 99 } -Option ReadOnly

$resultCode = $returnCode.Success
# 警告の表示抑止
Set-Item -Path Env:\SuppressAzurePowerShellBreakingChangeWarnings -Value "true"

# LogController オブジェクト生成
if ($Stdout) {
    $Log = New-Object LogController
}
else {
    $LogFilePath = Split-Path $scriptDir -Parent | Join-Path -ChildPath log -Resolve
    $LogFile = (Get-ChildItem $MyInvocation.MyCommand.Path).BaseName + ".log"
    $Log = New-Object LogController($(Join-Path $LogFilePath -ChildPath $LogFile), $false)
}

try {

    $SettingFilePath = Split-Path $MyInvocation.MyCommand.Path -Parent | Split-Path -Parent | Join-Path -ChildPath etc -Resolve
    $SettingFilePath = Join-Path -Path $SettingFilePath -ChildPath $ConfigFileName
    
    # 設定ファイル存在確認
    if (!(Test-Path $SettingFilePath)) {
        $Log.Error("${SettingFilePath}が存在しません")
        Exit $returnCode.Error
    }
    
    # 設定ファイルの読込
    $Settingjson = Get-Content $SettingFilePath | ConvertFrom-Json
    $Azurejson = $Settingjson.Azure

    $OutDirectory = $Settingjson.Output.Directory

    if (!(Test-Path -Path $OutDirectory -PathType Container)) {
        $Log.info("出力先フォルダ「$($OutDirectory)」が存在しないため、フォルダを作成します。")
        $resultNewDirectory = New-Item -Path $OutDirectory -ItemType Directory
        if (!$resultNewDirectory) {
            $Log.Error("出力先フォルダの作成に失敗しました。")
            exit $returnCode.Error
        }
    }

    # ファイルを開くダイアログ
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")

    $Dialog = New-Object System.Windows.Forms.OpenFileDialog
    $Dialog.Filter = "エクセルファイル(*.xlsx;*.xlsm)|*.xlsx;*.xlsm"
    $Dialog.InitialDirectory = $Settingjson.Excel.InitialDirectory
    $Dialog.Title = "ファイルを選択してください"
    $Dialog.Multiselect = $false

    # ダイアログを表示
    if ($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::Cancel) {
        $Log.info("Cancelが選択されたため、処理を終了します。")
        exit $returnCode.Success
    }

    $Log.info("読込ファイル「$($Dialog.FileName)」")
    # Excelインスタンス生成
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $Workbook = $Excel.Workbooks.Open($Dialog.FileName)

    $resultSubscription = Select-AzSubscription -SubscriptionId $Azurejson.SubscriptionID

    if (!$resultSubscription) {
        $Log.Error("サブスクリプション「$($Azurejson.SubscriptionID)」の選択に失敗しました。")
        exit $returnCode.Error
    }
    $Log.info("サブスクリプション「$($Azurejson.SubscriptionID)」を選択しました。")

    # Network Watcherの取得
    $NetworkWatcher = Get-AzNetworkWatcher -ResourceGroupName $Azurejson.NetworkWatcher.ResourceGroup -Name $Azurejson.NetworkWatcher.Name

    if (!$NetworkWatcher) {
        $Log.Error("NetworkWatcherの取得に失敗しました。")
        exit $returnCode.Error
    }

    foreach ($ExcelSheetjson in $Settingjson.Excel.Sheet) {
        $Sheet = $Workbook.Sheets.Item("$($ExcelSheetjson.Name)")
        
        if (!$Sheet) {
            $Log.Error("シート「$($ExcelSheetjson.Name)」の取得に失敗しました。")
            $resultCode = $returnCode.Error
            continue
        }
        
        # シートの読込＆通信チェック
        $resultCheckNSG = InvokeCheckNSG -returnCode $returnCode -Log $Log -Sheet $Sheet -ExcelSheetjson $ExcelSheetjson -NetworkWatcher $NetworkWatcher -OutDirectory $OutDirectory
        $resultCode = $resultCode -bor $resultCheckNSG

    }

}
#################################################
# エラーハンドリング
#################################################
catch {
    $Log.Error("予期しないエラーが発生しました。")
    $Log.Error($_.Exception)

    $resultCode = $returnCode.Exception
}
finally {
    if ($Workbook -ne $null) {
        $Workbook.Close($false)
        $Workbook = $null
    }
    if ($Excel -ne $null) {
        $Excel.Quit()
        $Excel = $null
    }
}

$Log.Info("NSGチェック処理が終了しました。")
exit $resultCode
