# LittleToolVba

vbac のお試しと自分用 VBA の管理

## インストール

### アドインファイルの配置

```powershell
if(Test-Path ${env:APPDATA}\Microsoft\AddIns) {
    Copy-Item -Path bin\LittleToolVba.xlam -Destination ${env:APPDATA}\Microsoft\AddIns -Force
} else {
    Write-Output "フォルダがありません: ${env:APPDATA}\Microsoft\AddIns"
}

```

### Excel の設定

- ファイル ＞ オプション ＞
  - セキュリティ センター(トラスト センター) ＞ マクロの設定 ＞ ☑ VBA プロジェクト オブジェクト モデルへのアクセスを信頼する
  - アドイン ＞ <kbd>設定</kbd> ＞ ☑ LittleToolVba ＞ <kbd>OK</kbd>
  - <kbd>OK</kbd>

## 開発

Excel ファイルに VBA のソースをインポート/エクスポートするツールである [Ariawase](https://github.com/vbaidiot/Ariawase)
をサブモジュールとして参照し使用しています。

サブモジュールをチェックアウトするには `git submodule update` を実行してください。

ソースフォルダ `src` とアドインフォルダの `.xlam` ファイルを同期するタスクを `tasks.json` に定義してあります。
