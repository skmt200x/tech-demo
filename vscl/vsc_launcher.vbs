Option Explicit
' ------------------------------------------------------------------------------
' Node.js 開発サポートな Visual Studio Code ランチャースクリプト
' ------------------------------------------------------------------------------

''' 機能
'   Node.js のインストーラ版ではなくエンベデッド版（ZIP形式）を使い開発する
'   場合の Visual Studio Code ランチャーです。
'   エンベデッド版を使う場合、環境変数 PATH に Node.js（node.exe）が存在しない
'   ため、デバッガとの相性が悪いですが、ランチャーで環境変数を設定して起動
'   することで解決します。
'   *.cmdだとコンソールウィンドウが残ってしまうため、*.vbsで実装しています。



''' Visual Studio Code のパス
'   既定のパスにインストールしている場合は、変更する必要はありません。
Const DEFAULT_VSC_PATH = "%ProgramFiles%\Microsoft VS Code\Code.exe"


''' 作業ディレクトリ、スクリプト名、設定ファイル名の取得
Dim oFS, strWorkDir, strWorkFileName, strConfigFilePath
Set oFS = CreateObject("Scripting.FileSystemObject")

Sub Include(ByVal param)
    Call ExecuteGlobal(oFS.OpenTextfile(param).ReadAll())
End Sub

strWorkDir = oFS.GetAbsolutePathName(".")
strWorkFileName = oFS.GetFileName(WScript.ScriptFullName)
strConfigFilePath = oFS.BuildPath(strWorkDir, strWorkFileName & ".config")

''' 環境変数オブジェクトの準備
Dim oWS, oEnv
Set oWS = WScript.CreateObject("WScript.Shell")
Set oEnv = oWS.Environment("PROCESS")


''' 設定ファイルの読み込み処理
Dim oConfigs
Set oConfigs = CreateObject("Scripting.Dictionary")
If oFS.FileExists(strConfigFilePath) Then
    Call Include(strConfigFilePath)
Else
    Call WScript.Echo("ERR: " & strWorkFileName & ".config does not found.")
    Call WScript.Quit(1)
End If


''' NODE_HOMEの設定
If (oConfigs.Item("NODE_HOME") <> "") Then
    oEnv.Item("NODE_HOME") = oConfigs.Item("NODE_HOME")
    oEnv.Item("PATH") = oEnv.Item("NODE_HOME") & ";" _
        & oEnv.Item("PATH")
ElseIf (oEnv.Item("NODE_HOME") <> "") Then
    oEnv.Item("PATH") = oEnv.Item("NODE_HOME") & ";" _
        & oEnv.Item("PATH")
Else
    Call WScript.Echo("ERR: NODE_HOME does not found on config or environment variable.")
    Call WScript.Quit(2)
End If


''' NODE_PATHの設定
If (oConfigs.Item("NODE_PATH") <> "") Then
    oEnv.Item("NODE_PATH") = oConfigs.Item("NODE_PATH")
ElseIf (oEnv.Item("NODE_PATH") <> "") Then
    ''' None
Else
    oEnv.Item("NODE_PATH") = oEnv.Item("NODE_HOME") & ";" _
        & oEnv.Item("NODE_HOME") & "\node_modules"
End If



''' Visual Studio Codeのパスを取得～起動
Dim strVSCPath, strCmdLine, strCmdArgs
If (oConfigs.Item("VSC_PATH") <> "") Then
    strVSCPath = oConfigs.Item("VSC_PATH")
Else
    strVSCPath = DEFAULT_VSC_PATH
End If

Dim i
strCmdArgs = ""
For i = 0 To (WScript.Arguments.Count -1) Step 1
    strCmdArgs = strCmdArgs & " """ & WScript.Arguments(i) & """"
Next

strCmdLine = """" & strVSCPath & """" & strCmdArgs
Call oWS.Run(strCmdLine, 4, False)


''' 終了処理
Set oWS = Nothing
Set oFS = Nothing
Set oEnv = Nothing

Call WScript.Quit(0)