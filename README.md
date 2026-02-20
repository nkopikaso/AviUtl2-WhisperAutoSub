# AviUtl2-WhisperAutoSub

AviUtl2でWhisperを使用して動画・音声から自動で字幕を生成するプラグインです。
本プラグインは独自に開発されたものです。

## ダウンロード

右側の **Releases** から `WhisperAutoSub.aux2` をダウンロードしてください。

## インストール

`WhisperAutoSub.aux2` を AviUtl2 の Plugin フォルダに入れてください。

> Plugin フォルダの場所: AviUtl2 メニュー →「その他」→「アプリケーションデータ」→「プラグインフォルダ」

## 必要環境

| 項目 | 要件 | 備考 |
|------|------|------|
| OS | Windows 10 / 11 | |
| AviUtl2 | 最新版推奨 | |
| Python | 3.10〜3.12 推奨 | 事前にインストールし、PATHを通すか初期設定で指定 |
| ffmpeg | 最新の安定版推奨 | 事前にインストールし、PATHを通すか初期設定で指定 |
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
| AviUtl2 | 最新版 |

## 主な機能

**2つのバックエンド対応**
- faster-whisper (CTranslate2) — 高速・省メモリ
- openai-whisper (PyTorch) — 公式実装

**6種類のモデル**
- tiny / base / small / medium / large-v3 / large-v3-turbo

**書式テンプレート**
- AviUtl2のエイリアス (.object) を読み込み、フォント・色・サイズ等を一括適用

**その他**
- 言語自動検出 / 日本語 / 英語 / 中国語 / 韓国語
- CUDA GPU 自動検出（CPU fallback あり）
- SRT ファイルエクスポート
- 句読点削除 / !?削除 / 全半角正規化
- 字幕延長（linger）— 発話終了後も指定秒数だけ字幕を表示
- レイヤー自動シフト — 配置先に既存オブジェクトがあれば次のレイヤーへ
- バックエンド・モデルの自動インストール（セットアップボタン）

## 使い方

1. タイムラインに動画/音声クリップを配置
2. メニューから「Whisper Subtitle」ウィンドウを開く
3. 初回のみ「初期設定」タブで ffmpeg / Python のパスを確認し、セットアップを実行
4. 「字幕生成」タブで Backend / Model / 言語等を設定
5. 「字幕生成」ボタンをクリック → 自動で字幕が配置されます

### 書式テンプレートの使い方

1. AviUtl2で好みのテキストオブジェクトを作成（フォント・色・サイズ等）
2. `.object` ファイルとしてエイリアス保存
3. 「字幕生成」タブの「書式: 選択」で指定
4. 以降の字幕生成にテンプレートの書式が適用されます

## フォルダ構成（自動生成）

```
data/Plugin/whisper_subtitle/
├── whisper_helper.py    … Pythonヘルパー（自動生成）
├── site-packages/       … pip パッケージ
├── models/              … Whisper モデル
└── temp/                … 一時ファイル
```

## トラブルシューティング

デバッグログ: `data/Plugin/whisper_subtitle/whisper_debug.log`

| 症状 | 対処 |
|------|------|
| ffmpegが見つからない | 初期設定タブでffmpegのパスを指定 |
| Pythonが見つからない | Python 3.10+をインストールし、パスを指定 |
| CUDAエラー / GPU非対応 | DeviceをCPUに変更 |
| 文字起こしが空 | 音声が無い/非常に短いクリップの可能性 |
| 処理が極端に遅い | デバッグログで device=cpu になっていないか確認。セットアップを再実行 |

## 注意事項

- 初回実行時はモデルダウンロードのため時間がかかります
- openai-whisper バックエンドでは pip install 時に PyTorch が入りますが、CUDA版が必要な場合はセットアップボタンから再インストールしてください

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
