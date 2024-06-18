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

### 時間・時刻を取得する
- 時間を計測する [link](code/time.bas#L7)
- 現在時刻を取得する [link](code/time.bas#L22)

### 文字列を扱う
- デリミタで分割して配列に格納 [link](code/strings.bas#L5)
- 文字列を結合する [link](code/strings.bas#L16)
- 配列の文字列を結合する [link](code/strings.bas#L24)
- 文字列を数値に変換する
  - 10進数 [link](code/strings.bas#L31)
  - 16進数 [link](code/strings.bas#L43)
- 16進数を10進数に変換 [link](code/strings.bas#L57)
- 10進数を16進数に変換 [link](code/strings.bas#L80)

### 正規表現
- 文字列から数値を取得
  - 整数 [link](code/regex.bas#L5)
  - 小数 [link](code/regex.bas#L21)
- 文字列から16進文字列取得
  - 0xがない場合 [link](code/regex.bas#L37)
  - 0xがある場合 [link](code/regex.bas#L53)
- 文字列を置換
  - 対称文字を削除 [link](code/regex.bas#L69)
  - 対称文字を置き換え [link](code/regex.bas#L78)

### グラフの整形



### 一般的な高速化対応
- 画面更新の停止/復旧
- 再計算の停止
- VBAが応答なしとなったとき