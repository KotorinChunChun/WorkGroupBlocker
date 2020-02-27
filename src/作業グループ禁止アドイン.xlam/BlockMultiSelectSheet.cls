VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BlockMultiSelectSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem
Rem 複数シート選択をした状態での操作を禁止する
Rem ・上書き保存しようとしたブックは解除してから保存させる
Rem ・新たに開いたブックなら真っ先に解除させる
Rem ・セルの編集なら取り消しを案内する
Rem
Rem　　インスタンスを生成するだけで監視します。
Rem
Rem @KotorinChunChun
Rem
Rem 2020/02/18
Rem
Option Explicit

Public WithEvents app As Application
Attribute app.VB_VarHelpID = -1

Private Sub app_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If ActiveWindow Is Nothing Then Exit Sub
    If ActiveWindow.SelectedSheets.Count = 1 Then Exit Sub
    
    'いいえを選択した場合は2度目以降の実行は行わない。
    Static MultiSelectedCounter As Long
    If MultiSelectedCounter > 0 Then MultiSelectedCounter = MultiSelectedCounter - 1: Exit Sub
    
    Debug.Print Sh.Name, Target.Address
    Select Case MsgBox("複数のシートが選択されたまま編集されようとしています。" & vbLf & _
                        "安全のため解除しませんか？" & vbLf & _
                        vbLf & _
                        "　　　はい：編集を取り消して単一シートを選択" & vbLf & _
                        "　　いいえ：無視してデータを書き込む", _
                        vbYesNoCancel, APP_NAME)
        Case vbYes
            ActiveWindow.SelectedSheets(1).Select
            'シートごとにこのイベントは呼ばれるが、最初の1回でUndoするともう呼ばれない。
            Application.Undo
        Case vbNo
            'いいえの場合、選択されたシートの数分Changeイベントが起こるのでその分だけイベントを取り消す
            If MultiSelectedCounter = 0 Then MultiSelectedCounter = ActiveWindow.SelectedSheets.Count - 1
    End Select
End Sub

Private Sub app_WorkbookBeforeSave(ByVal wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Call 作業グループ禁止
End Sub

Private Sub app_WorkbookOpen(ByVal wb As Workbook)
    Call 作業グループ禁止
End Sub

Sub 作業グループ禁止()
    If ActiveWindow Is Nothing Then Exit Sub
    If ActiveWindow.SelectedSheets.Count = 1 Then Exit Sub
    
    '保存と開くでは黙って解除しても良いと思う
    
    Select Case MsgBox("複数のシートが選択されています。" & vbLf & _
                        "安全のため解除しませんか？" & vbLf & _
                        vbLf & _
                        "　　　はい：解除する" & vbLf & _
                        "　　いいえ：解除しない" & vbLf & _
                        vbYesNoCancel, APP_NAME)
        Case vbYes: ActiveWindow.SelectedSheets(1).Select
        Case Else:  'vbNo = 何もしない。
    End Select
End Sub

'----------------------------------------
'コンストラクタ
Private Sub Class_Initialize()
    Set app = Application
End Sub

