# [VbaDevelopSupport - VBA開発支援アドイン](https://github.com/KotorinChunChun/VbaDevelopSupport)

* License: The MIT license
* Copyright (c) 2020 [KotorinChunChun](https://twitter.com/KotorinChunChun)
* 対象Excel: Microsoft® Excel® 2016 / Microsoft 365 32bit / 64bit
* 対象OS: Windows 10 1909 / 2004

## これは

ExcelのVBEに機能を追加して、VBA開発を楽にするためのアドインです。

![image](https://user-images.githubusercontent.com/55196383/94364650-f746ab00-0105-11eb-8e85-c9c1dd905211.png)



## ⚠ 利用条件 ⚠

* 作者自身の環境以外において、期待通りの動作となるかは一切保証できません。
* 一部の機能はモジュールを書き換えるため、ある日データが破損するかもしれません。当ツールのバックアップ機能に依存せず、自己責任でバックアップを取っておきましょう。
* ファイル→オプション→トラスト センター→[トラスト センターの設定(T)...]→マクロの設定→開発者向けマクロ設定→「☑ VBA プロジェクト オブジェクト モデルへのアクセスを信頼する(V)」にチェックを入れないと動作しません。この設定はExcelのセキュリティを著しく低下させた状態になることを理解した上で設定を変更してください。


## 使い方

アドインを実行するとVBEのメニューバーに「VBE開発支援」が増えるので、その中のコマンドを実行するだけ。

いずれのコマンドもアクティブなプロジェクトに対して実行されるので、処理したいプロジェクトを選択（またはモジュールを開いて）切り替えを忘れずに。

Excelを立ち上げたら自動で起動する常駐アドインとしたい場合は、自力でスクリプトを書くか手動で行ってください。



## 機能一覧

### ソースコード管理

* ソースをSRCにエクスポートする
* ソースをバックアップとエクスポートする
* ソースをYYYYMMDエクスポートする
* ソースコードのプロシージャ一覧を出力する
* ソースをSRCからインポートする（未実装）
* CustomUIをエクスポートする
* CustomUIをインポートする（未実装）

### コーディング支援

* Declareの生成
* Declareの変換

### VBE開発支援

* プロジェクトのパスワードを1234に変更する
* プロジェクトのフォルダを開く
* プロジェクトを閉じる
* ファイル化されていないブック全てを閉じる

### VBE機能追加
* 全てのコードウインドウを閉じる
* イミディエイトウィンドウを空にする

### 他
* VBA開発支援アドインを終了する



## エクスポートの基本ルール

### フォルダ構成

想定しているフォルダ構成は以下の通りです。

![image](./VBAソースコードバックアップ構成図.drawio.svg)



### 動作設定定義ファイル `kccsettings.json`

本アドインの動作を定義するファイルです。

実行対象ブックのプロジェクトフォルダの特定と、リリースフォルダ、ソースコードフォルダ、バックアップフォルダの定義などに使用しています。

たとえば、`./dev/実行対象プロジェクト.xlsm` があるとき、`./dev/kccsettings.json` に保存します。

もし、対象がサブフォルダ `./dev/hoge/実行対象プロジェクト.xlsm` の場合、上層に向かって `kccsettings.json` を検索して最初に発見したフォルダをプロジェクトルートフォルダと定めます。

算定で親フォルダを `./dev` としていますが、フォルダ名は違っても構いません。

設定ファイルは以下のように、UTF8のJSON形式で書きます。分からなければ、どのプロジェクトもこのままで構いません。本プロジェクトからコピーしてください。

￥マークに相当する部分は、エスケープのため `\\` となる点には注意

``` json
{
    // "#" で始まる行はコメント
    // .kccsettingファイルは .dev フォルダ直下に保存してください。
    // 書き出し時には再帰的に上位のフォルダを調べて読み込まれます。
    //
    // バックスラッシュ（￥記号）とスラッシュは、バックスラッシュによるエスケープが必要です。
    //
    // binPath リリースフォルダ
    "ExportBinFolder": ".\\..\\bin",
    // srcPath ソースコードフォルダ
    //"ExportSrcFolder": ".\\..\\src",
    "ExportSrcFolder": ".\\..\\src\\[FILENAME]",
    // src
    // backup binフォルダ
    "BackupBinFile": ".\\..\\backup\\bin\\[YYYYMMDD]_[HHMMSS]_[FILENAME]",
    // backup srcフォルダ
    //"BackupSrcFile": ".\\..\\backup\\src\\[YYYYMMDD]_[HHMMSS]_[FILENAME]"
    "BackupSrcFile": ".\\..\\backup\\src\\[FILENAME]\\[YYYYMMDD]_[HHMMSS]_[FILENAME]"
}
```



### 拡張子

ソースコードの拡張子は全て`.vba`です。

理由は次の通り

* 汎用エディタの関連付けを便利にするため。

* GitHubの自動言語判定でVBAと認識させるため。
* 拡張子が`.vba`でも、VBEへのインポートには支障がないため。

```
*.bas.vba 標準モジュール
*.cls.vba クラスモジュール／Sheet・Bookモジュール
*.frm.vba フォームモジュール
*.frm.frx フォームモジュールのバイナリ部分
CustomUI.xml   Excel 2007向けのリボンUI設定ファイル
CustomUI14.xml Excel 2010以降のリボンUI設定ファイル
```



### 除外ルール

以下の条件に当てはまるモジュールはエクスポートされません。

* 白紙のモジュール
* `Option Explicit`などの`Option`しか書かれていないモジュール
* 及び改行しか書かれていないモジュール
* 直前の`src`と容量の変化していない`.frm.frx`ファイル

これは、

* Excelの不要なシートモジュールの出力を防ぐため
* エクスポートのたびに必ず`.frm.frx`のバイナリが変化するため

です。



もし`.frm.frx`を更新したいのであれば、適当な図形を増やして容量を変化させるか既存の`../src`のファイルを消してからエクスポートしてください。



## ソースをエクスポートする

`../src/`にソースコードをエクスポートし、`../bin/`にプロジェクト一式を複製します。

GitHubで管理したい場合におすすめです。

#### 入力（プロジェクトの保存場所）

プロジェクトと同一フォルダのファイルすべてをbinに複製します。

```
/dev/AddinName.xlam
/dev/sample.sql
/dev/memo.txt
/dev/~$AddinName.xlam
```

ただし、同フォルダに `.kccignore` ファイルを保存することで、複製対象を制御することが可能です。



### リリース除外設定ファイル `.kccignore`

本アドインが実行（エクスポート・リリース）される際に、除外されるファイルを定義するファイルです。

`./dev/テストデータ` を `./bin/` へ複製されないようにしたりできます。

この設定ファイルは、以下のように.gitignore形式で書きます。詳しい文法は検索してください。（概ね互換性がありますが、独自実装なので不完全なところもあります）

``` txt
# "#" で始まる行はコメント
# https://qiita.com/anqooqie/items/110957797b3d5280c44f

# exe ファイルは要らない
# *.exe

# bin フォルダは要らない
# bin/

# addin.xlam は必要なファイル
# !/.temp/addin.xlam

# testフォルダの.xls ファイルは要らない
# test/**/*.xls

# 空のフォルダは維持できない

########## 以下に除外ルールを記述 ##########
~$*
.kccignore
要確認/*
```

例えば、 `~$ファイル名.xlam` というゴミファイルを除外するには、次のように書きます。

```
~$*
```



#### 出力

```
/bin/AddinName.xlam
/bin/sample.sql
/bin/memo.txt
/src/CodeName.bas.vba
```

※なお、bin、srcフォルダを削除してから、新たに作り直してエクスポートするため、大事なファイルをここに入れると消失します。



## ソースをバックアップとエクスポートする

上記と同じようにエクスポートしつつ、`../backup/`フォルダに日付情報付きでコピーします。

GitHub、WinMerge両方で管理したい場合におすすめです。

#### 入力（プロジェクトの保存場所）

プロジェクトと同一フォルダのファイルすべてをbinに複製します。

```
/任意のフォルダ/AddinName.xlam
```

#### 出力

```
/bin/AddinName.xlam
/src/CodeName.bas.vba
/backup/bin/YYYYMMDD_HHMMSS_AddinName.xlam
/backup/src/YYYYMMDD_HHMMSS/CodeName.bas.vba
```

※srcにエクスポートしたものが、backupに複製されます。



## ソースをYYYYMMDDにエクスポートする

`./src/`に日付フォルダを作成してソースコードをエクスポートします。

WinMergeによるバージョン管理をする場合におすすめです。

#### 入力（プロジェクトの保存場所）

```
/AddinName.xlam
```

#### 出力

```
/src/YYYYMMDD_HHMMSS/CodeName.bas.vba
```



## ソースコードのプロシージャ一覧を出力する

プロシージャ1件あたり1行となるように、新規エクセルブックにテーブル形式で出力します。

参考：https://qiita.com/Mikoshiba_Kyu/items/46b7243eb576848b3e55

### ユースケース

* 関数をデータベース化する場合
* 関数名の`Public/Private`つけ忘れをチェックする場合（テーブル右端参考）
* 引数や戻り値の型の統一をする場合



## Declareの生成

![image](https://user-images.githubusercontent.com/55196383/94357199-a3ba6a00-00d1-11eb-9045-0f47c1a2bc29.png)

2つのテキストボックスを備えたフォームが表示されます。

WinAPIの名前を入力すると、[Win32API_PtrSafe.TXT](https://www.microsoft.com/en-us/download/details.aspx?id=9970)に記載されているものであれば、自動的にDeclare文を生成してくれます。


#### 参考

* [VBAでWin32APIの64bit対応自動変換プログラムを作ってみた]([https://www.excel-chunchun.com/entry/vba-64bit-declare-convert
* [WinAPIの64bit化で出てくるPtrSafe､LongLong､LongPtrってなんなのさ？](https://www.excel-chunchun.com/entry/20200809-vba-declare-ptrsafe-longlong-longptr)



## Declareの変換

![image](https://user-images.githubusercontent.com/55196383/94357233-f8f67b80-00d1-11eb-9403-64d29dfbf811.png)

2つのテキストボックスを備えたフォームが表示されます。

ソースコードをペタっと貼ると、自動的に既存のDeclare文を検出して「Excel 2010 32bit/64bit両対応」と「Excel 2007以前対応」の宣言コードに置換します。

安全性を考慮するなら、両者のコードをWinMergeで差分を確認してから本番コードを置き換えてください。



## VbeDevelop.basについて

検証で使ったものや、開発中のもの、ネットからコピペしたものなど、ソースコードがぐちゃぐちゃです。

いつか整理します。



## 実行時エラーが出た場合や改造する人へ

あまり丁寧にエラー処理していません。フォルダ作成とかでエラーが出るかもしれません。

<br>

実行時エラーが出たら、メニューが動作しなくなります。いずれかの方法で復旧させてください。

* エクセルを再起動する
* F5を押して`AppMain.Reset_Addin`を実行する



## 今後の予定

* インポート対応
* プロシージャ別ファイル出力
* プロシージャ単位アップデート
* CustomUI XMLの読込
* テストの生成
* テストへのジャンプ
* テストの実行
* モジュールテストの実行
* マクロの記録で生成されたコードの最適化
* Declareの作成／変換に対応バージョンの切り替えオプションを追加



## 利用者へ

使用は自己責任でお願いします。何が起きても作者は保証できません。


