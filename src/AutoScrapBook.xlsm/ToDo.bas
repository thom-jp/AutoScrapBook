Attribute VB_Name = "ToDo"
' すべてクリアでは他にシートがある場合に _
    シートごと削除されるようにする。

' ExportFileモジュールのGetSavePathでGoToを使いすぎて _
    一度バグを出したので反省した。構造を見直す。

' Relocation時に最もサイズの大きなものを基準にロケートするように変更する。

' Word出力時に必ずGroupingされるよう修正する。
