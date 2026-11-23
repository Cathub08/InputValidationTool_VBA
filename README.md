# InputValidationTool_VBA
Excel VBA 入力チェックツール（v1.0）

## ツール概要
Excel VBAで作成した入力チェックツールです。
各項目に入力された内容の妥当性をチェックします。
入力エラーがある場合は、対象項目下のエラー表示ラベルへメッセージを表示し、
入力内容に問題がない場合は、正常メッセージを表示します。

## 作成目的
Excel VBAおよびUserFormを使用した入力チェック処理の学習を目的として作成しました。

メールアドレス、電話番号、郵便番号の形式判定やエラーメッセージ表示の実装を行い、
設計・実装・テストまでの一連の開発工程を経験することを目的としています。

## 開発環境
- Excel VBA
- UserForm

## 主な機能
- 氏名、メールアドレス、電話番号の必須入力チェック
- メールアドレス、電話番号、郵便番号の形式判定（VBScript.RegExpを使用）
- 年齢入力時の数値判定（IsNumericを使用）
- 入力エラー時のエラーメッセージ表示
- 入力正常時の完了メッセージ表示

## 画面イメージ・操作手順
### ①メイン画面
<img src=https://github.com/Cathub08/InputValidationTool_VBA/blob/main/images/%E2%91%A0%E3%83%A1%E3%82%A4%E3%83%B3%E7%94%BB%E9%9D%A2.png width="400">

### ②正常入力画面
<img src=https://github.com/Cathub08/InputValidationTool_VBA/blob/main/images/%E2%91%A1%E6%AD%A3%E5%B8%B8%E5%85%A5%E5%8A%9B.png width="600">

### ③エラー表示画面
<img src=https://github.com/Cathub08/InputValidationTool_VBA/blob/main/images/%E2%91%A2%E3%82%A8%E3%83%A9%E3%83%BC%E8%A1%A8%E7%A4%BA.png width="700">

## コード抜粋①
### 入力チェック処理
<img src=https://github.com/Cathub08/InputValidationTool_VBA/blob/main/images/Code%E6%8A%9C%E7%B2%8B%E2%91%A0.png width="900">

## コード抜粋②
### メールアドレス形式判定
<img src=https://github.com/Cathub08/InputValidationTool_VBA/blob/main/images/Code%E6%8A%9C%E7%B2%8B%E2%91%A1.png width="700">

## 工夫したポイント
- TrimおよびStrConv(vbNarrow)による入力値の正規化を実施
- StrConv(vbNarrow)では変換できない全角ハイフンが存在するため、
  UserForm上へ注釈ラベルを常時表示し、
  ハイフン入力時は半角「-」を使用するよう案内している
- 各入力項目ごとにエラーメッセージを個別表示
- VBScript.RegExpを利用した形式判定を実装
- モジュールを機能単位で分割し保守性を向上
- 要件定義、設計、実装、テストまで一連の開発工程を実施

## テスト
単体テストを実施しました。
- 正常系：各機能の動作確認
- 異常系：必須チェック、入力形式判定、数値チェック、対象シートなし等

✅ 単体テスト38件を実施

✅ 全テスト項目で期待結果を確認済み

## テスト結果

- 単体テスト：38件実施
- OK：38件
- NG：0件
- テスト成功率：100%

## フォルダ構成
```
docs/
├─要件定義.docx
├─設計書.docx
├─詳細設計書.docx
├─テスト仕様書.xlsx
└─バージョン管理メモ.txt

images/
├─①メイン画面.png
├─②正常入力.png
├─③エラー表示.png
├─Code抜粋①.png
└─Code抜粋②.png

README.md
LICENSE
```

## ドキュメント
- [要件定義](docs/要件定義.docx)
- [設計書](docs/設計書.docx)
- [詳細設計書](docs/詳細設計書.docx)
- [テスト仕様書](docs/テスト仕様書.xlsx)
- [バージョン管理メモ](docs/バージョン管理メモ.txt)

---

## License
このリポジトリは GNU GPL v3.0 ライセンスのもとで公開しています。


