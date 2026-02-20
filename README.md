# AviUtl2-WhisperAutoSub

AviUtl2でWhisperを使用して動画・音声から自動で字幕を生成するプラグインです。
本プラグインは独自に開発されたものです。

## ダウンロード

右側の **Releases** から `WhisperAutoSub.aux2` をダウンロードしてください。

---

## はじめに（初めての方へ）

本プラグインを使うには **Python** と **ffmpeg** が必要です。
まだインストールしていない方は、以下の手順で準備してください。

### 1. Pythonのインストール

1. [Python公式サイト](https://www.python.org/downloads/) にアクセス
2. **「Download Python 3.12.x」** ボタンをクリック（3.10〜3.12 ならOK）
3. ダウンロードした `python-3.12.x-amd64.exe` を実行
4. ⚠️ **最初の画面で「Add python.exe to PATH」にチェックを入れる**（最重要）
5. 「Install Now」をクリック
6. インストール完了後、コマンドプロンプトで確認:
   ```
   python --version
   ```
   `Python 3.12.x` と表示されれば成功です。

> ⚠️ 「Add python.exe to PATH」を忘れた場合:
> インストーラーを再実行 →「Modify」→「Next」→「Add Python to environment variables」にチェック →「Install」
>
> ⚠️ Python 3.13 以降は PyTorch（GPU処理に必要）が未対応の場合があります。**3.12 を推奨**します。

### 2. ffmpegのインストール

1. [gyan.dev](https://www.gyan.dev/ffmpeg/builds/) にアクセス
2. **「release builds」** セクションにある `ffmpeg-release-essentials.zip` をダウンロード（約80MB）
3. ZIPを展開すると `ffmpeg-x.x-essentials_build` フォルダができます
4. その中の `bin\ffmpeg.exe` を以下のいずれかの方法で使えるようにします:
   - **おすすめ**: `ffmpeg.exe` を AviUtl2 の実行ファイルと同じフォルダにコピー
   - または、プラグインの初期設定タブで `ffmpeg.exe` のパスを直接指定
   - または、システム環境変数の PATH に `bin` フォルダのパスを追加

> 💡 `essentials` 版で十分です。`full` 版（約200MB）は不要な追加コーデックを含むため大きくなりますが、本プラグインでは違いはありません。

### 3. プラグインのインストール

`WhisperAutoSub.aux2` を AviUtl2 の Plugin フォルダに入れてください。

> Plugin フォルダの場所: AviUtl2 メニュー →「その他」→「アプリケーションデータ」→「プラグインフォルダ」

### 4. 初回セットアップ（任意）

セットアップは**必須ではありません**。字幕生成ボタンを押した際にバックエンドやモデルが未導入の場合、自動でインストールされます。

事前にまとめて準備しておきたい場合は:

1. AviUtl2 を起動し、メニューバーの **「表示」→「Whisper Subtitle」** にチェックを入れてウィンドウを表示
2. 「字幕生成」タブで使いたい **Backend** と **Model** を選ぶ
3. 「初期設定」タブを開く
4. Python と ffmpeg のパスが表示されていることを確認（自動検出されます）
5. **「セットアップ」ボタン** をクリック → 選択中のバックエンドとモデルがインストールされます

> 💡 セットアップは「字幕生成」タブで **現在選択中の Backend / Model の組み合わせだけ** をインストールします。
>
> 後からモデルを変更した場合（例: small → large-v3）、次に字幕生成ボタンを押したときに **初回だけ自動でモデルがダウンロード** されます（large-v3 は約3GB）。
> ダウンロード中は処理が始まるまで時間がかかりますが、2回目以降はローカルのモデルが使われるのですぐに開始されます。

---

## 必要環境

| 項目 | 要件 | 備考 |
|------|------|------|
| OS | Windows 10 / 11 | |
| AviUtl2 | 最新版推奨 | |
| Python | 3.10〜3.12 推奨 | [ダウンロード](https://www.python.org/downloads/) |
| ffmpeg | 最新の安定版推奨 | [ダウンロード](https://ffmpeg.org/download.html) |
| GPU (任意) | NVIDIA CUDA 12.1+ 対応 | なくてもCPUモードで動作します |

### 動作確認環境

| 項目 | バージョン |
|------|-----------|
| OS | Windows 10 |
| CPU | Intel Core i7-12700F |
| GPU | NVIDIA GeForce RTX 3080 (10GB) |
| RAM | 32GB |
| Python | 3.12.10 |
| FFmpeg | 7.1 |
| AviUtl2 | ExEdit2 2.0beta33 |

---

## 使い方

1. タイムラインに動画/音声クリップを配置
2. メニューバーの **「表示」→「Whisper Subtitle」** でウィンドウを表示
3. 「字幕生成」タブで Backend / Model / 言語等を設定
3. **「字幕生成」ボタン** をクリック
4. 自動で文字起こし → 字幕オブジェクトがタイムラインに配置されます

### 書式テンプレートを使う

字幕のフォント・色・サイズ等をカスタマイズしたい場合:

1. AviUtl2で好みのテキストオブジェクトを作成
2. 右クリック →「エイリアスとして保存」で `.object` ファイルを作成
3. 「字幕生成」タブの「書式: 選択」で `.object` ファイルを指定
4. 以降の字幕生成に、テンプレートの書式が自動で適用されます

---

## バックエンドの選び方

| バックエンド | 特徴 | おすすめ |
|---|---|---|
| **faster-whisper** | 高速・省メモリ（CTranslate2使用） | 通常はこちらがおすすめ |
| **openai-whisper** | OpenAI公式実装（PyTorch使用） | 互換性重視の場合 |

faster-whisper は同じモデルでも openai-whisper の2〜4倍高速です。
特にこだわりがなければ **faster-whisper をおすすめ** します。

## モデルの選び方

| モデル | サイズ | 精度 | 速度 | 用途 |
|---|---|---|---|---|
| tiny | 75MB | ★☆☆☆☆ | 最速 | テスト・動作確認向け |
| base | 150MB | ★★☆☆☆ | 高速 | 簡易な文字起こし |
| small | 500MB | ★★★☆☆ | 普通 | バランス型 |
| medium | 1.5GB | ★★★★☆ | やや遅い | 高精度 |
| large-v3 | 3GB | ★★★★★ | 遅い | 最高精度 |
| large-v3-turbo | 1.6GB | ★★★★☆ | 普通 | large-v3の高速版 |

> 💡 迷ったら **large-v3-turbo** がおすすめです。large-v3に近い精度で、速度は約2倍速いです。

---

## 主な機能一覧

- **2つのバックエンド**: faster-whisper / openai-whisper
- **6種類のモデル**: tiny / base / small / medium / large-v3 / large-v3-turbo
- **書式テンプレート**: .object ファイルでフォント・色・サイズを一括適用
- **言語選択**: 自動検出 / 日本語 / 英語 / 中国語 / 韓国語
- **CUDA GPU対応**: 自動検出（CPU fallback あり）
- **SRTエクスポート**: 字幕データを .srt ファイルに出力
- **テキスト処理**: 句読点削除 / !?削除 / 全半角正規化
- **字幕延長（linger）**: 発話終了後も指定秒数だけ字幕を表示し続ける
- **レイヤー自動シフト**: 配置先に既存オブジェクトがあれば自動で次のレイヤーへ
- **自動インストール**: セットアップボタンで選択中のバックエンド・モデルを導入

## フォルダ構成（自動生成）

```
data/Plugin/whisper_subtitle/
├── whisper_helper.py    … Pythonヘルパー（自動生成）
├── site-packages/       … pip パッケージ
├── models/              … Whisper モデル
└── temp/                … 一時ファイル
```

---

## トラブルシューティング

問題が発生した場合は、まずデバッグログを確認してください。

📄 ログの場所: `data/Plugin/whisper_subtitle/whisper_debug.log`

| 症状 | 原因 | 対処 |
|------|------|------|
| 「ffmpegが見つかりません」 | ffmpegが未インストールまたはパスが通っていない | 初期設定タブでffmpegのパスを指定 |
| 「Pythonが見つかりません」 | Pythonが未インストールまたはパスが通っていない | Python 3.10〜3.12をインストールし、パスを指定 |
| CUDAエラー | NVIDIA GPU非対応 or ドライバが古い | DeviceをCPUに変更 |
| 文字起こしが空 | 音声が無い/非常に短いクリップ | 音声付きのクリップか確認 |
| 処理が極端に遅い（数分以上） | GPUが使われていない | ログで `device=cpu` になっていないか確認。セットアップを再実行 |
| セットアップが途中で止まる | ネットワーク接続の問題 | インターネット接続を確認し、再度セットアップ |

### うまくいかないときは

1. `data/Plugin/whisper_subtitle/whisper_debug.log` の内容を確認
2. `site-packages` フォルダを削除してセットアップをやり直す
3. それでも解決しない場合は、GitHubの Issues にログを添えて報告してください

---

## 注意事項

- 初回の字幕生成時やモデル変更時は、バックエンドやモデルの自動ダウンロードが発生するため時間がかかります（large-v3で約3GB）
- 事前にセットアップボタンで済ませておくとスムーズです
- openai-whisper バックエンドでは PyTorch が自動インストールされます。CUDA版が必要な場合はセットアップボタンから再インストールしてください
- Python 3.13 以降は PyTorch が未対応の可能性があります。3.10〜3.12 を推奨します

---

## 使用ライブラリ・モデル

本プラグインは以下のオープンソースをユーザー環境に自動ダウンロードします（同梱・再配布はしていません）。

| ライブラリ/モデル | ライセンス | 用途 |
|---|---|---|
| [OpenAI Whisper](https://github.com/openai/whisper) | MIT | 音声認識バックエンド |
| [faster-whisper](https://github.com/SYSTRAN/faster-whisper) | MIT | 高速音声認識バックエンド |
| [PyTorch](https://pytorch.org/) | BSD | 機械学習フレームワーク |
| [CTranslate2](https://github.com/OpenNMT/CTranslate2) | MIT | faster-whisper 推論エンジン |
| Whisper モデル (tiny〜medium, turbo) | MIT | 音声認識モデル |
| Whisper モデル (large-v3) | Apache-2.0 | 音声認識モデル |
| [FFmpeg](https://ffmpeg.org/) | LGPL/GPL | 音声抽出（ユーザーが別途用意） |

各ライブラリ・モデルの利用条件は、それぞれのライセンスに従います。
開発者の皆様に感謝いたします。

## ライセンス

MIT License

## 免責事項

本ソフトウェアの使用によって生じたいかなる損害についても、作者は責任を負いません。

## 作者

GitHub: [nkopikaso](https://github.com/nkopikaso)
