Option Explicit
' ------------------------------------------------------------------------------
' Node.js �J���T�|�[�g�� Visual Studio Code �����`���[�X�N���v�g
' ------------------------------------------------------------------------------

''' �@�\
'   Node.js �̃C���X�g�[���łł͂Ȃ��G���x�f�b�h�ŁiZIP�`���j���g���J������
'   �ꍇ�� Visual Studio Code �����`���[�ł��B
'   �G���x�f�b�h�ł��g���ꍇ�A���ϐ� PATH �� Node.js�inode.exe�j�����݂��Ȃ�
'   ���߁A�f�o�b�K�Ƃ̑����������ł����A�����`���[�Ŋ��ϐ���ݒ肵�ċN��
'   ���邱�Ƃŉ������܂��B
'   *.cmd���ƃR���\�[���E�B���h�E���c���Ă��܂����߁A*.vbs�Ŏ������Ă��܂��B



''' Visual Studio Code �̃p�X
'   ����̃p�X�ɃC���X�g�[�����Ă���ꍇ�́A�ύX����K�v�͂���܂���B
Const DEFAULT_VSC_PATH = "%ProgramFiles%\Microsoft VS Code\Code.exe"


''' ��ƃf�B���N�g���A�X�N���v�g���A�ݒ�t�@�C�����̎擾
Dim oFS, strWorkDir, strWorkFileName, strConfigFilePath
Set oFS = CreateObject("Scripting.FileSystemObject")

Sub Include(ByVal param)
    Call ExecuteGlobal(oFS.OpenTextfile(param).ReadAll())
End Sub

strWorkDir = oFS.GetAbsolutePathName(".")
strWorkFileName = oFS.GetFileName(WScript.ScriptFullName)
strConfigFilePath = oFS.BuildPath(strWorkDir, strWorkFileName & ".config")

''' ���ϐ��I�u�W�F�N�g�̏���
Dim oWS, oEnv
Set oWS = WScript.CreateObject("WScript.Shell")
Set oEnv = oWS.Environment("PROCESS")


''' �ݒ�t�@�C���̓ǂݍ��ݏ���
Dim oConfigs
Set oConfigs = CreateObject("Scripting.Dictionary")
If oFS.FileExists(strConfigFilePath) Then
    Call Include(strConfigFilePath)
Else
    Call WScript.Echo("ERR: " & strWorkFileName & ".config does not found.")
    Call WScript.Quit(1)
End If


''' NODE_HOME�̐ݒ�
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


''' NODE_PATH�̐ݒ�
If (oConfigs.Item("NODE_PATH") <> "") Then
    oEnv.Item("NODE_PATH") = oConfigs.Item("NODE_PATH")
ElseIf (oEnv.Item("NODE_PATH") <> "") Then
    ''' None
Else
    oEnv.Item("NODE_PATH") = oEnv.Item("NODE_HOME") & ";" _
        & oEnv.Item("NODE_HOME") & "\node_modules"
End If



''' Visual Studio Code�̃p�X���擾�`�N��
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


''' �I������
Set oWS = Nothing
Set oFS = Nothing
Set oEnv = Nothing

Call WScript.Quit(0)