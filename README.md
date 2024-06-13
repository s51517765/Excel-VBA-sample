# Excel VBA sample

### ファイルを読み込む
- テキストファイルから読み込む
  - 改行コードCRLFの場合 [link](code/fileIO.bas#L25)
  - 改行コードLFの場合 [link](code/fileIO.bas#L55)
  - ファイルオープンダイアログを使う [link](code/fileIO.bas#L11)

### ファイルに書き込む
- Excelに出力
  - Thisbookに出力 [link](code/fileIO.bas#L44)
  - 新しいbookを作成し出力 [link](code/fileIO.bas#L123)
  - その他のbookに出力 [link](code/fileIO.bas#L137)
- テキストファイルに出力 [link](code/fileIO.bas#L155)

### 時刻を取得する
- 時間を計測する [link](code/time.bas#L7)
- 現在時刻を取得する [link](code/time.bas#L22)

### 文字列を扱う
- デリミタで分割して配列に格納
- 配列の文字列を結合する
- 文字列を数値に変換する
  - 10進数
  - 16進数

### 正規表現
- 文字列から数値を取得
- 文字列から16進文字列取得

### グラフの整形



### 一般的な高速化対応
- 画面更新の停止/復旧
- 再計算の停止
- VBAが応答なしとなったとき