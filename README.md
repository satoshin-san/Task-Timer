# タスク管理タイマーアプリ
このアプリケーションは、Tkinterを使用したタスク管理と時間追跡を行うデスクトップアプリです。複数のタスクを登録し、それぞれの作業時間を計測することができます。作業記録はExcelファイルにエクスポート可能です。

## 機能
- タスクの追加: 複数のタスクを一括で追加できます。
- タイマー機能: 選択したタスクに対して作業時間を計測します。
- 割り込みタスク対応: タイマーを一時停止して割り込みタスクを管理できます。
- タスク完了: 完了したタスクをリストから削除します。
- 作業記録: 作業時間をログとして保存し、Excel形式でエクスポート可能です。

## 使用方法
1. タスクを入力し、「一括タスク追加」ボタンでタスクリストに追加します。
2. リストから作業するタスクを選択し、目標作業時間（分）を入力します。
3. 「タイマー開始」ボタンで作業を開始し、終了時に「作業終了」ボタンを押します。
4. 必要に応じて、割り込みタスクを追加し、そのタスクに切り替えることができます。
5. 作業記録はアプリを閉じる際にExcelにエクスポートできます。

## ショートカットキー
- **タスクの選択:** タスクリストで ↑ / ↓ キーを使用
- **目標時間を入力後、Enterキーでタイマー開始**
- **作業終了:** Delete キー
- **ボタンにフォーカスがある場合、スペースキーでボタンを押せます**

## 必要なライブラリ
このアプリケーションを実行するには、次のPythonライブラリが必要です:
- `tkinter` （標準ライブラリ）
- `pandas` （インストールが必要）

## 注意事項
- エクスポート機能で保存されたExcelファイルは、デフォルトで`YYYY-MM-DD_作業記録.xlsx`という名前になります。
