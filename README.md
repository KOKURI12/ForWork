#TaskExcelGenerator（JapaneseVer）

このプロジェクトは、JavaとApache POIライブラリを使って、業務用タスク管理用のExcelファイルを自動生成するツールです。  
Microsoft Teamsの情報整理、WBSタスクの可視化、現場作業の効率化に活用できます。

---

## 機能一覧

- Excelファイルを自動生成（日本語対応）
- ヘッダー行スタイル（色付き、太字）
- プルダウンメニュー（ステータス／優先度）
- 締切日が近い場合はセルを黄色にハイライト
- 未完了のタスクは赤字で表示
- Teamsリンクにハイパーリンク設定済み

---

## 必要な環境

| ツール       | 推奨バージョン           | 説明                   
|--------------|--------------------------|--------------------------
| Java JDK     | 17以上（例：22.0.1）     | Javaプログラムの実行用 
| Apache Maven | 3.6以上（例：3.9.9）     | プロジェクトの依存解決／ビルド 
| OS           | Windows / macOS          | どちらでも可 

---

## セットアップ手順（Windows前提）

### 1. Java の確認
cmd java -version
例:java 22.0.1

Oracle JDK ダウンロード先:https://www.oracle.com/java/technologies/javase-downloads.html

### 2. Maven の確認
mvn -version
例:Apache Maven 3.9.9

★インストールできていない場合★
Maven　ダウンロード先:https://maven.apache.org/download.cgi

環境変数設定
1.システム環境変数設定⇒新規⇒変数名:MAVEN_HOME　変数値:apache-mavenフォルダ格納先

2.パス⇒編集⇒新規⇒%MAVEN_HOME%\bin

3.cmd⇒mvn -version

### 3.プロジェクトをダウンロード＆解凍
TaskExcelGenerator_Ja.zip


## 実行手順
①TaskExcelGenerator_Ja.zip解凍フォルダ配置先⇒cmd

②プロジェクトのコンパイル　　　 ⇒mvn compile

③Javaアプリケーションの実行     ⇒mvn exec:java -Dexec.mainClass=com.example.TaskExcelGenerator

## 出力されるファイル
タスク管理テンプレート_日本語版.xlsx

## Excelファイルの仕様
項目	                         説明
タスクID / タスク名 / WBS番号	各タスク情報の基本項目
ステータス	                プルダウン（未開始、進行中、完了、確認待ち）
優先度	                        プルダウン（高、中、低）
締切日	                        2日以内で黄色にハイライト
Teams指示リンク	                自動ハイパーリンク形式
備考欄	                        メモ自由記入欄

## フォルダ構成
任务Excel生成器项目包_日本語版/
├── README.md
├── pom.xml
└── src/
    └── main/
        └── java/
            └── com/
                └── example/
                    └── TaskExcelGenerator.java

## 補足・サポート
このテンプレートは拡張可能です。以下のような改良も可能です：

・GUI付きアプリケーション化

・Excelテンプレートのカスタマイズ対応

・完了状況の自動統計表示など

