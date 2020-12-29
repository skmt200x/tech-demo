
# 目的
Node.jsやJRE/JDKなど開発対象毎にバージョンが異なる場合に環境を切り替えながらVisual Studio Codeを起動するためのランチャープログラム（スクリプト）。

# 使い方
1. vsc_launcher.vbs を任意のディレクトリに配置する。
   ```
   %USERPROFILE%\VSCL\vsc_launcher.vbs
   ```
2. 作業フォルダに vsc_launcher.vbs.config を配置する。
   ```
   %USERPROFILE%\Workspace1\vsc_launcher.vbs.config
   ```
3. vsc_launcher.vbs のショートカットファイルを作成する。
4. ショートカット  (右クリック) > [プロパティ] > [作業フォルダー] に 「vsc_launcher.vbs.config」を配置したフォルダーを指定する。

# 構成
- vsc_launcher.vbs
    - プログラム（スクリプト）本体です。
    このファイルは直接使うのではなくWindowsショートカットを作成して使う前提です。
- vsc_launcher.vbs.config
    - 設定（構成）ファイルです。
    「vsc_launcher.vbs」ショートカットファイルの「作業ディレクトリ」以下にある「vsc_launcher.vbs.config」を読み込みます。

