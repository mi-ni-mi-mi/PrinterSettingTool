# PrinterSettingTool

## 概要
プリンタの追加や削除を行うアプリ。

## 使用方法
タスクスケジューラに、ユーザログオン時に最上位の特権で実行するよう、このアプリを登録する。これにより、管理者権限のないユーザでもプリンタの設定が可能になる。

## 特徴
- 多重起動を確認した時に初めてアプリが表示される。
- アプリを終了できない。（閉じるボタンをクリックすると最小化される。タスクバーにも表示されない。）
- タスクマネージャーには、バックグランドプロセスとして表示される。

## 参考
- https://dobon.net/vb/dotnet/graphics/defaultprinter.html
- https://dobon.net/vb/dotnet/graphics/printerport.html
- https://social.msdn.microsoft.com/Forums/ja-JP/bc667708-e8a5-4ea5-b88f-12458ff2bbe6/1250312522125311247912398ip12450124891252412473123982146224471
