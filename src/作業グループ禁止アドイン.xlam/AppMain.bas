Attribute VB_Name = "AppMain"
Rem
Rem @appname WorkGroupBlocker - 作業グループ禁止アドイン
Rem
Rem @module AppMain
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2020/02/18 : 初回版
Rem    2020/02/27 : Git公開
Rem
Option Explicit
Option Private Module

Public Const APP_NAME = "作業グループ禁止アドイン"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.11"
Public Const APP_UPDATE = "2020/02/27"
Public Const APP_URL = "https://www.excel-chunchun.com/entry/work_group_blocker"

Public instBlockMultiSelectSheet As BlockMultiSelectSheet

'--------------------------------------------------
'アドイン実行時
Sub AddinStart()
    MsgBox "あなたの身を【作業グループ】から完全に護ります！！！" & vbLf & _
             vbLf & _
            "複数のシートを選びっぱなしにする" & vbLf & _
            "【作業グループ】は絶対に許しません！！！", _
                vbInformation + vbOKOnly, APP_NAME
    Call MonitorStart
End Sub

'アドイン一時停止時
Sub AddinStop()
    Dim item
    For Each item In Array( _
        "え～作業グループ許可しちゃうの～？", _
        "作業グループって複数シート選択のことだよ", _
        "複数シート選択したまま作業すると、データ壊しちゃうかもよ？", _
        "複数シート選択したまま保存すると、次に使う人がデータ壊しちゃうかもよ？", _
        "それでも作業グループ使いたいの？")
        If MsgBox(item, vbExclamation + vbYesNo, APP_NAME) = vbNo Then
            MsgBox "だよね～作業グループなんていらないよね～", vbOKOnly, APP_NAME
            Exit Sub
        End If
    Next
    MsgBox "もぉどうなっても知らないんだからっ！！！", vbOKOnly, APP_NAME
    Call MonitorStop
End Sub

'アドイン設定表示
Sub AddinConfig(): Call SettingForm.Show: End Sub

'アドイン情報表示
Sub AddinInfo()
    Select Case MsgBox(ThisWorkbook.Name & vbLf & vbLf & _
            "バージョン : " & APP_VERSION & vbLf & _
            "更新日　　 : " & APP_UPDATE & vbLf & _
            "開発者　　 : " & APP_CREATER & vbLf & _
            "実行パス　 : " & ThisWorkbook.Path & vbLf & _
            "公開ページ : " & APP_URL & vbLf & _
            vbLf & _
            "使い方や最新版を探しに公開ページを開きますか？" & _
            "", vbInformation + vbYesNo, "バージョン情報")
        Case vbNo
            '
        Case vbYes
            CreateObject("Wscript.Shell").Run APP_URL, 3
    End Select
End Sub

'アドイン完全終了
Sub AddinEnd(): ThisWorkbook.Close False: End Sub
'--------------------------------------------------

'監視開始
'Workbook_Openから呼ばれる
Sub MonitorStart(): Set instBlockMultiSelectSheet = New BlockMultiSelectSheet: End Sub

'監視停止
Sub MonitorStop(): Set instBlockMultiSelectSheet = Nothing: End Sub
