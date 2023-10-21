# holiday_japan_vba
VBAで祝日判定を行うためのプログラムです。VBEでクラスモジュールにインポートして使用します。

外部データは参照せず、計算によって一年間の祝日の辞書データ（Dictionaryオブジェクト）を作成し、
その祝日辞書データを元に、祝日かどうかの判定、祝日の名前の取得、祝日一覧のメッセージ表示などを行います。

祝日の計算は、Masahiro TANAKA氏の "holiday_japan" (https://github.com/masa16/holiday_japan) のロジックをそのまま使わせていただきました。
