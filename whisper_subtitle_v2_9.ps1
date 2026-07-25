##############################################################################
# Whisper Subtitle Plugin for AviUtl2 - v2.9.0
#
# v2.9.0: セットアップのチェック形式化。(A)起動時にpythonを1回だけ回してtorch/whisper/faster-whisper/fugashiの
# 導入状況を一括実測しキャッシュ(g_probe)、タブ切替やfugashi判定のたびにpythonを起動する重さを解消(表示専用、
# SetupThreadのskip判定は従来通りその場でPkgOkする)。(B)Python 3.10未満を検出したら警告ダイアログ(続行可)。
# (C)環境タブを再配置し「ライブラリ」欄にPyTorch行/モデル行を追加、最下端は350に収まる高さで再計算。
# (D)チェックボックス5個(torch/whisper/faster-whisper/fugashi/モデル)で「チェックした項目だけ強制再導入」に変更、
# 「強制再導入」チェックボックス1個は廃止。backend=openai-whisperのままfaster-whisperも導入できるようになり、
# kotoba-whisperを選んでも動かない問題の対処が可能に(InstallWhisperPkgラムダに切り出し)。
# (E)ffmpeg自動ダウンロードの独立大ボタンを廃止しffmpeg行内の[DL]ボタンに統合。
# (F)fugashi単独[再導入]ボタンとInstallFugashiThread()を削除(チェック方式に統合、機能重複を解消)。
# UpdateFugashiStatus()はpython同期実行(最大60秒)をやめg_probeキャッシュを読むだけに変更、
# 「タブ切替時にfugashi判定を呼ばない」制約を解除しRefreshSetupTabState()から呼べるように。
# (G)モデル存在判定をModelExists()に切り出しSetupThreadと環境タブ表示の両方から使用、
# kotoba-whisper選択時にbackend=openai-whisperだと動かない旨を警告表示。
# (H)PyTorch行を新設しg_probeの実測結果(CUDA/CPU/未導入)を表示。
# (I)字幕生成ボタン押下時に不足コンポーネントをファイル存在+probeキャッシュのみで検出し
# (pythonは起動しない)、導入が必要なら環境タブへ誘導するダイアログを表示。
# (J)セットアップボタンのラベルを「セットアップ (チェック項目は再導入)」に変更。
#
# v2.8.10: ステータス表示凍結バグ修正+環境タブの並び順変更。(I)SetupThread内でSetWindowTextW(g_status,...)を
# 直接呼んでいた6箇所が環境タブ専用ステータス(g_statusSetup)を素通りしていたため、セットアップ完了後も
# 「fugashi確認中...」等の表示のまま凍結する不具合があった(処理自体は正常完走、表示のみの問題)→
# g_statusとg_statusSetupを同時更新するSetStatusW()を新設し6箇所を置き換え。(J)環境タブ「外部ツール」欄を
# Python→ffmpeg順(Pythonが無いと何も始まらない依存関係順)に、「音声認識ライブラリ」欄をwhisper→faster-whisper順
# (実際に使っているのはwhisper側)に並び替え。ロジック・変数名は無変更、配置のみ入れ替え
#
# v2.8.9: 用語統一+強制再導入+fugashiボタン整理。(F)「プロ分割」という何がプロなのか伝わらない用語を「文節区切り」に統一(5箇所)、括弧書きの補足ラベルを廃止して短縮(「セグメント結合(短い字幕をまとめる)」→「短い字幕を結合」等)。(G)セットアップのskip判定はimportが通るかで見ているため中身が壊れていると再導入できない穴があった→「強制再導入」チェックボックスを追加し、ONならskip判定をすべて無視して入れ直す(実行後は自動でOFFに戻す、確認ダイアログ経由)。(H)fugashi単独導入ボタンは機能重複(PkgOkバグ解消済みのv2.8.7以降は独立ボタンである必要が無い)のため、音声認識ライブラリ欄の小型[再導入]ボタンへ格下げ移動
#
# v2.8.8: 環境タブの進捗可視化+ffmpeg再DL防止+UI再配置。(A)環境タブ専用の進捗バー/ステータス表示を追加(従来はg_tabSubCtrls側にのみ存在し環境タブでは非表示だった)。(B)自前DL済みffmpegをGetEffectiveFFmpeg()の探索順に追加+DLボタンに既導入確認ダイアログ。(C)ffmpegDL中にzipサイズをポーリングして実進捗を表示。(D)ffplay/ffprobe(計約204MB)を削除。(E)環境タブをセクション見出し+状態記号(●/○)付きレイアウトに再配置
#
# v2.8.7: 導入でつまずかないためのセットアップ改善3点。(1)PkgOk判定がspDirを通さず誤判定→数GB再DLされる不具合を修正。(2)torchのCUDA版をnvidia-smiから判定して導入(古い環境/GPU非搭載を考慮)、whisperより先に導入して二重DL回避。(3)ffmpeg未検出時に公式配布元から自動DL
#
# v2.8.6: 環境タブのフォルダ選択(faster-whisper/whisper)をIFileDialog+FOS_PICKFOLDERSでモダン化。ffmpeg/pythonと同じexplorer風の見た目に
#
# v2.8b: v2.8監査で確定した不具合の修正版 (v2.8本体は変更なし、別ファイルとしてビルド)
#   - [v2.8b] RunProcess(): timeoutMs が未使用だった問題を修正 (GetTickCount64で計測しTerminateProcess)
#   - [v2.8b] SplitText(): UTF-8境界後退がbp==posで止まり無限ループする問題を修正
#   - [v2.8b] ExtractAudio(): 固定長wchar_t[2048]バッファをstd::wstring組み立てに変更
#   - [v2.8b] g_busy を std::atomic<bool> 化、compare_exchange_strongでアトミックに開始判定
#             + 実行中は生成/セットアップボタンをEnableWindow(FALSE)で無効化
#   - [v2.8b] JSON生成で "" (ダブルクォート) のエスケープ漏れを修正
#
# v2.8.1: 高DPI (ディスプレイ拡大率) 対応 (v2.8bのロジックは変更なし、UI座標のみスケール)
#   - [v2.8.1] High-DPI support: fix UI layout breakage at 125%/150% display scaling
#
# v2.8.2a: [実験] プロ分割の時間割当を単語タイムスタンプ実測に変更 (話の間/話速変化に追従。フォールバック=文字数比按分)
# v2.8.2a: 先行表示(秒)設定を追加 (字幕を音声より早く出す。0=無効)
# v2.8.2b: [実験] 波形RMSスナップ (咳払い等の短い音を除外した発話区間の立ち上がりに字幕開始を吸着、±0.35s)
# v2.8.2c: [実験] 二段閾値VAD (息継ぎ等の低い持続音を発話区間から除外、区間先頭を声の立ち上がりへトリム)
# v2.8.2d: [実験] 先頭トリムを息継ぎ付き区間のみに限定 (v2.8.2cの一律10フレーム遅延を修正)
# v2.8.2e: [実験] 発話検出をSilero VAD(ニューラル)に変更。RMS閾値はフォールバックに降格 (クリップ依存の不安定さを根治)
# v2.8.3: 配布版統合。Silero VAD 2系統(faster-whisper同梱→silero-vadパッケージ)、RMSフォールバック廃止(スナップなしに降格)
# v2.8.5: 簡易スプリッタ(プロ分割OFF時)の複合助詞分割バグ修正。空白を区切り優先化+から/より/たんの分離禁止+前後空白トリム
# v2.8.2: セグメント結合オプション追加 (opt-in checkbox, default OFF = 既存挙動不変)
#   - [v2.8.2] 「セグメント結合」ON時、隣接セグメントを間(gap<=0.4s)が小さく合計がmaxchars以内なら結合。
#             maxchars を「上限」から「目標」に格上げし、2-5文字の断片化を解消 (海外ユーザーIssue対応)
#   - デフォルトOFFのため、既存の生成結果・安定版の挙動には一切影響しない
#
# 機能 (v2.8から継承):
#   - faster-whisper / openai-whisper 対応 (自動インストール)
#   - CUDA GPU加速 + CPU fallback
#   - 書式テンプレート (.object)
#   - 句読点削除、!?削除、全半角正規化
#   - 字幕延長 (発話後の表示維持)
#   - レイヤー自動シフト (既存オブジェクト回避)
#   - SRTエクスポート
#   - モデル: tiny/base/small/medium/large-v3/large-v3-turbo/kotoba-whisper
#   - [v2.8] ヒント文(initial_prompt)で固有名詞精度向上
#   - [v2.8] ハルシネーション対策 (condition_on_previous_text)
#   - [v2.8] temperature自動フォールバック
#   - [v2.8] word_timestamps切替 / repetition_penalty
#   - [v2.8] セグメント品質フィルタ (avg_logprob / no_speech_prob)
#   - [v2.8] hallucination_silence_threshold
#   - [v2.8] ホットワード (hotwords)
#   - [v2.8] kotoba-whisperモデル対応 (日本語高速)
#   - [v2.8] Batched推論 (GPU並列処理)
#   - [v2.8] SDK構造体をbeta40に更新 (EDIT_SECTION/HOST_APP_TABLE)
#   - [v2.8] ビルド安定化: /utf-8 /wd4828 コンパイルオプション追加
#   - [v2.8] ソース出力をBOMなしUTF-8に修正 (日本語環境対応)
#   - [v2.8] ビルドログの文字化け修正 (Console OutputEncoding)
#
# ビルド: .\whisper_subtitle_v2_9_0.ps1 [出力ディレクトリ]
# 要件: Visual Studio 2022, CMake 3.15+
# 出力: whisper_subtitle_v2_9_0.aux2 (既存の whisper_subtitle_v2_8_10.aux2 / whisper_subtitle.aux2 は上書きしない)
#
# Note: [v2.8.8] PyTorchのCUDA版はnvidia-smiのCUDA Versionから自動判定(cu121/cu118/cpu)して導入
##############################################################################

$d = if($args.Count -gt 0){ $args[0] } else { [Environment]::GetFolderPath("Desktop") }
$projDir = "$d\aviutl2_dev\whisper_subtitle_plugin_v2_9_0"
$src = "$projDir\src"
$logFile = "$d\whisper_build_log_v2_9_0.txt"
$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

"" | Out-File $logFile -Encoding UTF8
"Whisper Subtitle v2.9.0 Build $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File $logFile -Append -Encoding UTF8

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " Whisper Subtitle v2.9.0 - Build" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$cpp = @'
// Whisper Subtitle v2.8.2
// faster-whisper only, no whisper.h dependency
#include <windows.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <commdlg.h>
#include <commctrl.h>
#include <shellapi.h>
#include <string>
#include <vector>
#include <thread>
#include <fstream>
#include <sstream>
#include <algorithm>
#include <cstdio>
#include <atomic>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "comdlg32.lib")

// =========================================================================
// AviUtl2 Plugin SDK structures (from plugin2.h - exact match)
// =========================================================================

struct INPUT_PLUGIN_TABLE;
struct OUTPUT_PLUGIN_TABLE;
struct FILTER_PLUGIN_TABLE;
struct SCRIPT_MODULE_TABLE;
struct EDIT_HANDLE;
struct PROJECT_FILE;

typedef void* OBJECT_HANDLE;

struct OBJECT_LAYER_FRAME { int layer, start, end; };
struct MEDIA_INFO { int video_track_num, audio_track_num; double total_time; int width, height; };

struct MODULE_INFO {
    int type;
    static constexpr int TYPE_SCRIPT_FILTER  = 1;
    static constexpr int TYPE_SCRIPT_OBJECT  = 2;
    static constexpr int TYPE_SCRIPT_CAMERA  = 3;
    static constexpr int TYPE_SCRIPT_TRACK   = 4;
    static constexpr int TYPE_SCRIPT_MODULE  = 5;
    static constexpr int TYPE_PLUGIN_INPUT   = 6;
    static constexpr int TYPE_PLUGIN_OUTPUT  = 7;
    static constexpr int TYPE_PLUGIN_FILTER  = 8;
    static constexpr int TYPE_PLUGIN_COMMON  = 9;
    LPCWSTR name;
    LPCWSTR information;
};

struct EDIT_INFO {
    int width, height;
    int rate, scale;
    int sample_rate;
    int frame;
    int layer;
    int frame_max;
    int layer_max;
    int display_frame_start;
    int display_layer_start;
    int display_frame_num;
    int display_layer_num;
    int select_range_start;
    int select_range_end;
    float grid_bpm_tempo;
    int grid_bpm_beat;
    float grid_bpm_offset;
    int scene_id;
};

struct EDIT_SECTION {
    EDIT_INFO* info;
    OBJECT_HANDLE (*create_object_from_alias)(LPCSTR alias, int layer, int frame, int length);
    OBJECT_HANDLE (*find_object)(int layer, int frame);
    int (*count_object_effect)(OBJECT_HANDLE object, LPCWSTR effect);
    OBJECT_LAYER_FRAME (*get_object_layer_frame)(OBJECT_HANDLE object);
    LPCSTR (*get_object_alias)(OBJECT_HANDLE object);
    LPCSTR (*get_object_item_value)(OBJECT_HANDLE object, LPCWSTR effect, LPCWSTR item);
    bool (*set_object_item_value)(OBJECT_HANDLE object, LPCWSTR effect, LPCWSTR item, LPCSTR value);
    bool (*move_object)(OBJECT_HANDLE object, int layer, int frame);
    void (*delete_object)(OBJECT_HANDLE object);
    OBJECT_HANDLE (*get_focus_object)();
    void (*set_focus_object)(OBJECT_HANDLE object);
    PROJECT_FILE* (*get_project_file)(EDIT_HANDLE* edit);
    OBJECT_HANDLE (*get_selected_object)(int index);
    int (*get_selected_object_num)();
    bool (*get_mouse_layer_frame)(int* layer, int* frame);
    bool (*pos_to_layer_frame)(int x, int y, int* layer, int* frame);
    bool (*is_support_media_file)(LPCWSTR file, bool strict);
    bool (*get_media_info)(LPCWSTR file, MEDIA_INFO* info, int info_size);
    OBJECT_HANDLE (*create_object_from_media_file)(LPCWSTR file, int layer, int frame, int length);
    OBJECT_HANDLE (*create_object)(LPCWSTR effect, int layer, int frame, int length);
    void (*set_cursor_layer_frame)(int layer, int frame);
    void (*set_display_layer_frame)(int layer, int frame);
    void (*set_select_range)(int start, int end);
    void (*set_grid_bpm)(float tempo, int beat, float offset);
    LPCWSTR (*get_object_name)(OBJECT_HANDLE object);
    void (*set_object_name)(OBJECT_HANDLE object, LPCWSTR name);
    LPCWSTR (*get_layer_name)(int layer);
    void (*set_layer_name)(int layer, LPCWSTR name);
    LPCWSTR (*get_scene_name)();
    void (*set_scene_name)(LPCWSTR name);
    void (*set_scene_size)(int width, int height);
    void (*set_scene_frame_rate)(int rate, int scale);
    void (*set_scene_sample_rate)(int sample_rate);
};

struct EDIT_HANDLE {
    bool (*call_edit_section)(void (*func_proc_edit)(EDIT_SECTION* edit));
    bool (*call_edit_section_param)(void* param, void (*func_proc_edit)(void* param, EDIT_SECTION* edit));
    void (*get_edit_info)(EDIT_INFO* info, int info_size);
    void (*restart_host_app)();
    void (*enum_effect_name)(void* param, void (*func_proc_enum_effect)(void* param, LPCWSTR name, int type, int flag));
    static constexpr int EFFECT_TYPE_FILTER     = 1;
    static constexpr int EFFECT_TYPE_INPUT       = 2;
    static constexpr int EFFECT_TYPE_TRANSITION  = 3;
    static constexpr int EFFECT_FLAG_VIDEO       = 1;
    static constexpr int EFFECT_FLAG_AUDIO       = 2;
    static constexpr int EFFECT_FLAG_FILTER      = 4;
    void (*enum_module_info)(void* param, void (*func_proc_enum_module)(void* param, MODULE_INFO* info));
    HWND (*get_host_app_window)();
};

struct PROJECT_FILE {
    LPCSTR (*get_param_string)(LPCSTR key);
    void (*set_param_string)(LPCSTR key, LPCSTR value);
    bool (*get_param_binary)(LPCSTR key, void* data, int size);
    void (*set_param_binary)(LPCSTR key, void* data, int size);
    void (*clear_params)();
    LPCWSTR (*get_project_file_path)();
};

struct HOST_APP_TABLE {
    void (*set_plugin_information)(LPCWSTR information);
    void (*register_input_plugin)(INPUT_PLUGIN_TABLE* table);
    void (*register_output_plugin)(OUTPUT_PLUGIN_TABLE* table);
    void (*register_filter_plugin)(FILTER_PLUGIN_TABLE* table);
    void (*register_script_module)(SCRIPT_MODULE_TABLE* table);
    void (*register_import_menu)(LPCWSTR name, void (*func)(EDIT_SECTION* edit));
    void (*register_export_menu)(LPCWSTR name, void (*func)(EDIT_SECTION* edit));
    void (*register_window_client)(LPCWSTR name, HWND hwnd);
    EDIT_HANDLE* (*create_edit_handle)();
    void (*register_project_load_handler)(void (*func)(PROJECT_FILE* project));
    void (*register_project_save_handler)(void (*func)(PROJECT_FILE* project));
    void (*register_layer_menu)(LPCWSTR name, void (*func)(EDIT_SECTION* edit));
    void (*register_object_menu)(LPCWSTR name, void (*func)(EDIT_SECTION* edit));
    void (*register_config_menu)(LPCWSTR name, void (*func)(HWND hwnd, HINSTANCE dll_hinst));
    void (*register_edit_menu)(LPCWSTR name, void (*func)(EDIT_SECTION* edit));
    void (*register_clear_cache_handler)(void (*func)(EDIT_SECTION* edit));
    void (*register_change_scene_handler)(void (*func)(EDIT_SECTION* edit));
    void (*register_import_menu_param)(LPCWSTR name, void* param, void (*func)(void* param));
    void (*register_export_menu_param)(LPCWSTR name, void* param, void (*func)(void* param));
    void (*register_layer_menu_param)(LPCWSTR name, void* param, void (*func)(void* param));
    void (*register_object_menu_param)(LPCWSTR name, void* param, void (*func)(void* param));
    void (*register_edit_menu_param)(LPCWSTR name, void* param, void (*func)(void* param));
    void (*register_file_drop_handler)(LPCWSTR name, LPCWSTR filefilter, void (*func)(EDIT_SECTION* edit, LPCWSTR file));
    void (*register_file_drop_param_handler)(LPCWSTR name, LPCWSTR filefilter, void* param, void (*func)(void* param, LPCWSTR file));
};

// =========================================================================
// Globals
// =========================================================================

static HINSTANCE g_hInst = 0;
static HWND g_wnd = 0;
static HWND g_modelCombo = 0, g_deviceCombo = 0, g_backendCombo = 0;
static HWND g_langCombo = 0, g_qualityEdit = 0, g_tempEdit = 0;
static HWND g_chkRemovePunct = 0, g_chkNormalize = 0, g_chkRemoveExclam = 0;
static HWND g_chkMergeSeg = 0; // v2.8.2: セグメント結合 (opt-in, default OFF)
static HWND g_chkMorphSplit = 0; // v2.8.2: 形態素解析プロ分割 (opt-in, default OFF)
static HWND g_chkTwoLine = 0; // v2.8.5: 2行まで許容 (opt-in, default OFF)
static HWND g_fugashiStatus = 0; // v2.8.5: プロ分割ライブラリ状態表示
static HWND g_fwLocLabel = 0, g_owLocLabel = 0; // faster-whisper / openai-whisper location labels
static HWND g_layerEdit = 0, g_maxCharEdit = 0;
static HWND g_tab = 0;
static std::vector<HWND> g_tabSubCtrls, g_tabSettingsCtrls, g_tabAccCtrls, g_tabSetupCtrls;

static void RefreshSetupTabState(); // v2.8.8: 定義は後方(UpdateWhisperLocLabelsの近く)。SwitchTabより先に前方宣言が必要
static int SC(int v); // v2.9.0【E】: 定義は後方(High-DPI scaling helpers)。RefreshSetupTabState内のMoveWindowより先に前方宣言が必要
static void SwitchTab(int idx){
    for(auto h : g_tabSubCtrls) ShowWindow(h, idx == 0 ? SW_SHOW : SW_HIDE);
    for(auto h : g_tabSettingsCtrls) ShowWindow(h, idx == 1 ? SW_SHOW : SW_HIDE);
    for(auto h : g_tabAccCtrls) ShowWindow(h, idx == 2 ? SW_SHOW : SW_HIDE);
    for(auto h : g_tabSetupCtrls) ShowWindow(h, idx == 3 ? SW_SHOW : SW_HIDE);
    // v2.9.0: RefreshSetupTabState()はUpdateFugashiStatus()を呼ぶが、probeキャッシュ参照のみ(python同期実行なし)に
    // 変わったため「fugashi再判定は絶対に含めない」という旧制約は解除された。上の一括SW_SHOWより後に呼ぶことだけ守る(順序厳守)。
    if(idx == 3) RefreshSetupTabState();
}
static HWND g_templateLabel = 0, g_status = 0, g_progress = 0;
static HWND g_ffmpegLabel = 0, g_pythonLabel = 0;
// v2.8.8: 環境タブ専用の進捗バー/ステータス(【A】)。環境タブを開いている間はg_status/g_progressがSW_HIDEで見えないため専用に持つ
static HWND g_statusSetup = 0, g_progressSetup = 0;
static HWND g_ffmpegStatusLabel = 0, g_pythonStatusLabel = 0; // v2.8.8: 状態記号(●導入済み/○未導入)
// v2.9.4: g_btnDlFfmpeg(ffmpeg専用のDL/再取得ボタン)は廃止。チェック+セットアップに一本化した。
// v2.9.0【D】: 旧・全項目まとめて入れ直すチェックボックス1個を廃止し、項目別5個に置き換え。チェックされた項目だけ
// SetupThreadのskip判定を無視して入れ直す。SaveSettings/LoadSettingsには保存しない(毎回OFF始まりが安全)。
// v2.9.4: g_chkFfmpeg を追加。ffmpegだけ専用の[DL]/[再取得]ボタンという別動線になっており
// 「なぜこれだけ別なのか」とユーザーが戸惑う構造だった(開発者指摘)ため、他と同じチェック方式に統一した。
static HWND g_chkFfmpeg = 0, g_chkTorch = 0, g_chkWhisper = 0, g_chkFaster = 0, g_chkFugashi = 0; // v2.9.1: g_chkModelは廃止
static HWND g_torchStatusLabel = 0, g_modelStatusLabel = 0; // v2.9.0【H/G】: PyTorch行・モデル行の状態表示

// v2.9.0【A】: 起動時に python を1回だけ回して導入状況をまとめて実測しキャッシュする。
// 表示専用。SetupThread の skip 判定には使わない(実行時点の実態を見るべきなので従来の PkgOk を維持)。
struct PkgProbe {
    std::atomic<bool> done{false};   // 実測完了したか
    std::atomic<bool> pyFound{false};// python 自体が見つかったか
    int  pyMajor = 0, pyMinor = 0, pyPatch = 0;
    bool torch = false, torchCuda = false;
    bool whisper = false, fasterWhisper = false, fugashi = false;
};
static PkgProbe g_probe;

static HWND g_lingerEdit = 0;
static HWND g_leadEdit = 0; // v2.8.2a: 先行表示(秒)
static HWND g_promptEdit = 0; // v2.8: initial_prompt (hint text)
static HWND g_chkNoPrevText = 0; // v2.8: condition_on_previous_text=False
static HWND g_chkWordTs = 0; // v2.8: word_timestamps toggle
static HWND g_chkRepPenalty = 0; // v2.8: repetition_penalty
static HWND g_hotwordsEdit = 0; // v2.8: hotwords
static HWND g_chkBatched = 0; // v2.8: batched inference
// v2.9.5: faster-whisper の vad_filter。従来 True ハードコードで、実測すると「文字が前に出る」
// (開発者が実運用で観測していた傾向) の主因だった。既定OFF。openai-whisper では無関係なのでグレーアウトする。
static HWND g_chkVad = 0;
static EDIT_HANDLE* g_edit = nullptr;
// v2.8b: g_busy is now std::atomic<bool>. The check-and-set ("is a job already running? if not,
// start one") must be atomic, otherwise two threads (e.g. rapid double-click on Generate/Setup)
// can both observe g_busy==false and both proceed. compare_exchange_strong makes the whole
// check-and-set indivisible.
static std::atomic<bool> g_busy{false};
static HWND g_btnGenerate = 0, g_btnSetup = 0; // v2.8b: kept so we can EnableWindow(FALSE) while busy
static std::string g_templatePath, g_templateContent;
static std::string g_ffmpegPath, g_pythonPath;
static std::string g_fwSpPath, g_owSpPath; // custom site-packages dirs for faster-whisper / openai-whisper
static int g_projectRate = 30;

#define WM_UPDATE_STATUS (WM_USER + 100)
#define WM_UPDATE_PROGRESS (WM_USER + 101)
#define WM_PROBE_DONE (WM_USER + 102) // v2.9.0【A】: ProbeThread()完了通知
#define IDC_GENERATE 1001
#define IDC_TEMPLATE 1002
#define IDC_RESET_TPL 1003
#define IDC_EXPORT_SRT 1004
#define IDC_SETUP 1005
#define IDC_FFMPEG_BR 1006
#define IDC_PYTHON_BR 1007
#define IDC_FW_BR 1008
#define IDC_OW_BR 1009
#define IDC_FW_RESET 1010
#define IDC_OW_RESET 1011
// v2.9.0【F】: fugashi単独[再導入]ボタンの専用IDは廃止。チェックボックス方式に統合。
#define IDC_DL_FFMPEG 1013

// =========================================================================
// String conversion (must be first - used by everything)
// =========================================================================

static std::wstring Utf8ToWide(const std::string& u){
    if(u.empty()) return L"";
    int n = MultiByteToWideChar(CP_UTF8, 0, u.c_str(), -1, NULL, 0);
    std::wstring w(n, 0);
    MultiByteToWideChar(CP_UTF8, 0, u.c_str(), -1, &w[0], n);
    if(!w.empty() && w.back() == 0) w.pop_back();
    return w;
}
static std::string WideToUtf8(const std::wstring& w){
    if(w.empty()) return "";
    int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, NULL, 0, NULL, NULL);
    std::string u(n, 0);
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, &u[0], n, NULL, NULL);
    if(!u.empty() && u.back() == 0) u.pop_back();
    return u;
}
static int MsgBox(HWND hwnd, const std::string& text, const std::string& title, UINT type){
    return MessageBoxW(hwnd, Utf8ToWide(text).c_str(), Utf8ToWide(title).c_str(), type);
}

// =========================================================================
// Path helpers (portable)
// =========================================================================

static std::string GetDllDir(){
    wchar_t buf[MAX_PATH] = {};
    GetModuleFileNameW(g_hInst, buf, MAX_PATH);
    std::wstring s(buf);
    size_t pos = s.find_last_of(L"\\/");
    return WideToUtf8((pos != std::wstring::npos) ? s.substr(0, pos) : s);
}
static std::string GetPluginDir(){ return GetDllDir() + "\\whisper_subtitle"; }
static std::string GetExeDir(){
    std::string d = GetDllDir();
    for(int i = 0; i < 2; i++){
        size_t p = d.find_last_of("\\/");
        if(p != std::string::npos) d = d.substr(0, p);
    }
    return d;
}
static std::string GetTempDir(){ return GetPluginDir() + "\\temp"; }
static std::string GetModelsDir(){ return GetPluginDir() + "\\models"; }
static std::string GetSitePackagesDir(){ return GetPluginDir() + "\\site-packages"; }
static std::string GetIniPath(){ return GetPluginDir() + "\\whisper_subtitle.ini"; }

// =========================================================================
// UTF-8 file operation helpers (handles Japanese paths correctly)
// =========================================================================

static BOOL FileExistsU(const std::string& p){
    return GetFileAttributesW(Utf8ToWide(p).c_str()) != INVALID_FILE_ATTRIBUTES;
}
static BOOL CreateDirU(const std::string& p){
    return CreateDirectoryW(Utf8ToWide(p).c_str(), nullptr);
}
static BOOL DeleteFileU(const std::string& p){
    return DeleteFileW(Utf8ToWide(p).c_str());
}
// v2.9.9: フォルダを中身ごと削除する。ffmpeg.exe を本体フォルダへ移した後に、
// 展開作業用フォルダ(プラグインフォルダ\ffmpeg\<version>\...)を片付けるために使う。
// SHFileOperation は追加の依存(shell32)を増やさず既にリンク済み。
// pFrom はダブルNUL終端が必須なので注意(ここを間違えると別のパスを消しかねない)。
static void RemoveDirectoryTreeU(const std::string& p){
    if(p.empty()) return;
    std::wstring w = Utf8ToWide(p);
    if(GetFileAttributesW(w.c_str()) == INVALID_FILE_ATTRIBUTES) return;
    std::vector<wchar_t> from(w.begin(), w.end());
    from.push_back(L'\0');
    from.push_back(L'\0'); // ダブルNUL終端
    SHFILEOPSTRUCTW op = {};
    op.wFunc  = FO_DELETE;
    op.pFrom  = from.data();
    op.fFlags = FOF_NO_UI | FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT;
    SHFileOperationW(&op);
}

// =========================================================================
// Debug log
// =========================================================================

static void DebugLog(const std::string& msg){
    std::ofstream f(Utf8ToWide(GetPluginDir() + "\\whisper_debug.log"), std::ios::app);
    if(f.is_open()) f << msg << "\n";
}

static void SetPathLabel(HWND label, const std::string& path, const char* defaultText){
    if(!label) return;
    if(path.empty())
        SetWindowTextW(label, Utf8ToWide(defaultText).c_str());
    else{
        size_t p = path.find_last_of("\\/");
        std::string v = (p != std::string::npos) ? "..." + path.substr(p) : path;
        SetWindowTextW(label, Utf8ToWide(v).c_str());
    }
}

// =========================================================================
// Embedded Python helper (ALWAYS creates output file, even on error)
// =========================================================================

static const char* g_pyHelper = R"PYHELPER(# -*- coding: utf-8 -*-
import sys, json, os, traceback, glob

# Add local site-packages (whisper_subtitle/site-packages)
_script_dir = os.path.dirname(os.path.abspath(__file__))
_local_sp = os.path.join(_script_dir, "site-packages")
if os.path.isdir(_local_sp) and _local_sp not in sys.path:
    sys.path.insert(0, _local_sp)
# Also ensure system site-packages is in path
_sys_sp = os.path.join(os.path.dirname(sys.executable), "Lib", "site-packages")
if os.path.isdir(_sys_sp) and _sys_sp not in sys.path:
    sys.path.append(_sys_sp)

def add_cuda_paths():
    # Check both system and local site-packages
    sites = [_local_sp] if os.path.isdir(_local_sp) else []
    sys_site = os.path.join(os.path.dirname(sys.executable), "Lib", "site-packages")
    if os.path.isdir(sys_site):
        sites.append(sys_site)
    for site in sites:
        for pkg in ["nvidia/cublas/bin", "nvidia/cudnn/bin", "nvidia/cuda_runtime/bin",
                    "nvidia_cublas_cu12", "nvidia_cudnn_cu12", "ctranslate2"]:
            p = os.path.join(site, pkg.replace("/", os.sep))
            if os.path.isdir(p):
                os.environ["PATH"] = p + ";" + os.environ.get("PATH", "")
                if hasattr(os, "add_dll_directory"):
                    try: os.add_dll_directory(p)
                    except: pass
        for dll_dir in glob.glob(os.path.join(site, "nvidia*", "**", "bin"), recursive=True):
            os.environ["PATH"] = dll_dir + ";" + os.environ.get("PATH", "")
            if hasattr(os, "add_dll_directory"):
                try: os.add_dll_directory(dll_dir)
                except: pass
        ct2_lib = os.path.join(site, "ctranslate2", "lib")
        if os.path.isdir(ct2_lib):
            os.environ["PATH"] = ct2_lib + ";" + os.environ.get("PATH", "")
            if hasattr(os, "add_dll_directory"):
                try: os.add_dll_directory(ct2_lib)
                except: pass
        # v2.9.6【配布】torch/lib を追加。Windows版 torch は cublas64_12.dll / cudnn*.dll を
        # nvidia/* パッケージではなく torch\lib に同梱している。従来この場所を見ていなかったため、
        # openai-whisper用にtorchが入っている環境でも faster-whisper が CUDA で動かず、
        # 「cublas64_12.dll is not found」で転写の途中から落ちていた。
        torch_lib = os.path.join(site, "torch", "lib")
        if os.path.isdir(torch_lib):
            os.environ["PATH"] = torch_lib + ";" + os.environ.get("PATH", "")
            if hasattr(os, "add_dll_directory"):
                try: os.add_dll_directory(torch_lib)
                except: pass

def main():
    if len(sys.argv) < 3:
        print("Usage: whisper_helper.py <batch.json> <output.txt>")
        sys.exit(1)
    batch_path, output_path = sys.argv[1], sys.argv[2]
    err_path = output_path + ".err"
    with open(err_path, "w", encoding="utf-8") as ef:
        ef.write(f"Python: {sys.executable}\nArgs: {sys.argv}\n")
    # CRITICAL: Always create output file, even on error
    try:
        _run(batch_path, output_path, err_path)
    except Exception as e:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(f"ERROR|Unexpected error|{e}\n")
        with open(err_path, "a", encoding="utf-8") as ef:
            ef.write(f"FATAL: {e}\n{traceback.format_exc()}")
        sys.exit(1)

def _run(batch_path, output_path, err_path):
    try:
        with open(batch_path, "r", encoding="utf-8") as f:
            batch = json.load(f)
    except Exception as e:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(f"ERROR|JSON read error|{e}\n")
        return
    model_size = batch.get("model", "small")
    language = batch.get("language", "ja")
    if language == "auto":
        language = None
    device = batch.get("device", "cpu")
    # v2.9.8【重要】同梱ffmpegのフォルダをPATH先頭に足す。
    # openai-whisper の audio.py load_audio() は subprocess で "ffmpeg" を **PATHから**起動するため、
    # プラグインが data\Plugin\whisper_subtitle\ffmpeg\<ver>\bin\ に置いたffmpegを見つけられず
    # FileNotFoundError [WinError 2] で落ちる(音声を読む一歩手前まで全部成功しているのに字幕0件になる)。
    # faster-whisper は PyAV で自前デコードするので影響を受けない = 「fasterなら動くのにwhisperだと空」の原因。
    # C++側が GetEffectiveFFmpeg() で解決した実パスを渡してくるので、そのフォルダを載せる。
    _ffm = batch.get("ffmpeg", "")
    if _ffm and os.path.isfile(_ffm):
        _ffdir = os.path.dirname(_ffm)
        os.environ["PATH"] = _ffdir + os.pathsep + os.environ.get("PATH", "")
        with open(err_path, "a", encoding="utf-8") as ef:
            ef.write(f"ffmpeg dir added to PATH: {_ffdir}\n")

    # Add extra site-packages paths (from custom whisper locations)
    for sp in batch.get("extra_sp", []):
        if os.path.isdir(sp) and sp not in sys.path:
            sys.path.insert(0, sp)
    if device == "auto":
        try:
            import torch
            device = "cuda" if torch.cuda.is_available() else "cpu"
        except ImportError:
            device = "cpu"
    beam_size = batch.get("beam_size", 5)
    if isinstance(beam_size, str):
        beam_size = int(beam_size) if beam_size.isdigit() else 5
    if beam_size <= 0:
        beam_size = 5
    temperature = batch.get("temperature", 0)
    try:
        temperature = float(temperature)
    except (ValueError, TypeError):
        temperature = 0
    clips = batch.get("clips", [])
    model_dir = batch.get("model_dir", "")
    backend = batch.get("backend", "faster-whisper")
    if device == "cuda":
        add_cuda_paths()

    if backend == "whisper":
        _run_openai_whisper(model_size, language, device, clips, model_dir, output_path, err_path, beam_size, temperature, batch)
    else:
        _run_faster_whisper(model_size, language, device, clips, model_dir, output_path, err_path, beam_size, temperature, batch)

# v2.8.2: 形態素解析ベースの文節分割 (fugashi + unidic-lite, opt-in)。プロ品質の自然な区切り。
_morph_tagger = None
def _morph_get_tagger():
    global _morph_tagger
    if _morph_tagger is None:
        import fugashi
        _morph_tagger = fugashi.Tagger()
    return _morph_tagger

def _morph_group(text, maxc):
    """text を文節境界で maxc 文字以内のチャンクに分割。fugashi 未導入等は [text] を返す(安全)。"""
    try:
        tg = _morph_get_tagger()
    except Exception:
        return [text]
    # 文節へ: 自立語で新文節を開始、付属語(助詞/助動詞/記号/接尾辞)は前にくっつける
    bun = []
    cur = ""
    for w in tg(text):
        pos = w.feature.pos1
        if (not cur) or pos in ("助詞", "助動詞", "補助記号", "記号", "接尾辞"):
            cur += w.surface
        else:
            bun.append(cur)
            cur = w.surface
    if cur:
        bun.append(cur)
    if maxc <= 0:
        maxc = 20
    chunks = []
    cur = ""
    for b in bun:
        if cur and len(cur) + len(b) > maxc:
            chunks.append(cur)
            cur = b
        else:
            cur += b
    if cur:
        chunks.append(cur)
    return chunks if chunks else [text]

)PYHELPER" R"PYHELPER(
_vad_mode = "none"
_silero_model = None

def _speech_runs(path, hop=0.02, min_run=0.20, hang=0.12):
    """発話区間 [(開始秒,終了秒),...] をSilero VADで検出。
    v2.8.3: faster-whisper同梱→silero-vadパッケージの順に試行、どちらも無ければ[]
    (=スナップなし、素の単語時刻)。RMS方式は廃止(クリップ音量分布依存で不安定)。"""
    global _vad_mode, _silero_model
    try:
        import wave
        import numpy as np
        with wave.open(path, "rb") as wf:
            sr = wf.getframerate(); ch = wf.getnchannels(); sw = wf.getsampwidth()
            raw = wf.readframes(wf.getnframes())
        if sw != 2 or sr != 16000:
            return []
        x = np.frombuffer(raw, dtype=np.int16).astype(np.float32) / 32768.0
        if ch > 1:
            x = x.reshape(-1, ch).mean(axis=1)
    except Exception:
        return []
    # 1) faster-whisper 同梱の Silero
    try:
        from faster_whisper.vad import get_speech_timestamps, VadOptions
        opts = VadOptions(min_speech_duration_ms=int(min_run * 1000),
                          min_silence_duration_ms=100, speech_pad_ms=0)
        ts = get_speech_timestamps(x, vad_options=opts)
        _vad_mode = "silero-fw"
        return [(t["start"] / sr, t["end"] / sr) for t in ts]
    except Exception:
        pass
    # 2) silero-vad パッケージ (torchベース。whisperユーザーはtorch導入済み)
    try:
        import torch
        from silero_vad import load_silero_vad, get_speech_timestamps as _gst
        if _silero_model is None:
            _silero_model = load_silero_vad()
        ts = _gst(torch.from_numpy(x), _silero_model,
                  min_speech_duration_ms=int(min_run * 1000),
                  min_silence_duration_ms=100, speech_pad_ms=0)
        _vad_mode = "silero-pkg"
        return [(t["start"] / sr, t["end"] / sr) for t in ts]
    except Exception:
        pass
    _vad_mode = "none"
    return []

def _snap_start(t, runs, win=0.35):
    """字幕開始秒tを発話立ち上がりへ吸着。発話中に深く食い込んだ境界(連続トークの
    チャンク境界)は触らない。"""
    for (a, b) in runs:
        if a <= t <= b:
            return a if (t - a) <= win else t
    best = None; bd = None
    for (a, b) in runs:
        d = abs(a - t)
        if d <= win and (bd is None or d < bd):
            bd = d; best = a
    return best if best is not None else t

def _norm_len(s):
    return len("".join(s.split()))

def _group_spans(spans, ml):
    # v2.8.5: spans = [(cs, ce, text), ...] を ml 個ずつ束ねる
    if ml <= 1:
        return spans
    out = []
    for i in range(0, len(spans), ml):
        g = spans[i:i+ml]
        out.append((g[0][0], g[-1][1], "\\n".join(x[2] for x in g)))
    return out

def _emit_seg(results, ci, sf, ef, text, morph_split, maxchars, words=None, tl_start=0, fps=60, stats=None, runs=None, max_lines=1):
    """seg を results に追加。morph_split 時は文節分割して複数追加。
    v2.8.2a: words(単語タイムスタンプ)があれば実測時刻で割当。無ければ文字数比で按分。
    v2.8.5: max_lines>1 の場合、chunks を max_lines 個ずつ束ねて1テロップ(\\n連結)にする。"""
    if not (morph_split and text):
        results.append(f"{ci}|{sf}|{ef}|{text}")
        return
    chunks = _morph_group(text, maxchars)
    n = len(chunks)
    if n <= 1:
        results.append(f"{ci}|{sf}|{ef}|{text}")
        return
    # 1) 単語タイムスタンプによる実測割当 (v2.8.2a)
    try:
        if words:
            spans = []
            pos = 0
            for wt, ws, we in words:
                L = _norm_len(wt)
                if L <= 0 or ws is None or we is None:
                    continue
                spans.append((pos, pos + L, float(ws), float(we)))
                pos += L
            total_w = pos
            total_c = sum(_norm_len(c) for c in chunks)
            if spans and total_w > 0 and total_c > 0:
                scale = total_w / total_c
                out = []
                cpos = 0
                for ch in chunks:
                    cl = _norm_len(ch)
                    a = cpos * scale
                    b = (cpos + cl) * scale
                    cpos += cl
                    ov = [sp for sp in spans if sp[1] > a + 1e-6 and sp[0] < b - 1e-6]
                    if not ov:
                        out = None
                        break
                    out.append((ov[0][2], ov[-1][3], ch))
                if out:
                    if runs:
                        sn = []
                        for cs_s, ce_s, ch in out:
                            ns = _snap_start(cs_s, runs)
                            if ns < ce_s:
                                if stats is not None and abs(ns - cs_s) > 0.02:
                                    stats["snap"] = stats.get("snap", 0) + 1
                                sn.append((ns, ce_s, ch))
                            else:
                                sn.append((cs_s, ce_s, ch))
                        out = sn
                    out = _group_spans(out, max_lines)  # v2.8.5
                    prev_e = None
                    for cs_s, ce_s, ch in out:
                        cs = tl_start + int(cs_s * fps)
                        ce = tl_start + int(ce_s * fps)
                        if prev_e is not None and cs < prev_e:
                            cs = prev_e
                        if ce <= cs:
                            ce = cs + 1
                        prev_e = ce
                        if ch:
                            results.append(f"{ci}|{cs}|{ce}|{ch}")
                    if stats is not None:
                        stats["word"] = stats.get("word", 0) + 1
                    return
    except Exception:
        pass
    # 2) フォールバック: 文字数比の按分 (旧・等分割から改良)
    span = ef - sf
    total_c = sum(len(c) for c in chunks) or 1
    acc = 0
    prev = sf
    spans2 = []
    for i, ch in enumerate(chunks):
        acc += len(ch)
        ce = ef if i == n - 1 else sf + span * acc // total_c
        cs = prev
        if ce > cs and ch:
            spans2.append((cs, ce, ch))
            prev = ce
    spans2 = _group_spans(spans2, max_lines)  # v2.8.5
    for cs, ce, ch in spans2:
        results.append(f"{ci}|{cs}|{ce}|{ch}")
    if stats is not None:
        stats["prop"] = stats.get("prop", 0) + 1

def _run_faster_whisper(model_size, language, device, clips, model_dir, output_path, err_path, beam_size=5, temperature=0, batch=None):
    if batch is None: batch = {}
    try:
        from faster_whisper import WhisperModel
    except ImportError as e:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(f"ERROR|faster-whisper not installed|pip install faster-whisper\n")
        with open(err_path, "a", encoding="utf-8") as ef:
            ef.write(f"ImportError: {e}\n")
        return
    with open(err_path, "a", encoding="utf-8") as ef:
        ef.write(f"Loading: {model_size} device={device} (faster-whisper)\n")
    # v2.8: Map distil/kotoba model names to HuggingFace repo
    distil_map = {
        "kotoba-whisper": "kotoba-tech/kotoba-whisper-v2.0-faster",
    }
    model_path = distil_map.get(model_size, model_size)
    kwargs = {}
    if model_dir:
        kwargs["download_root"] = model_dir
        local_path = os.path.join(model_dir, model_size)
        if os.path.isdir(local_path):
            model_path = local_path
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Using local: {model_path}\n")
    model = None
    actual_device = device
    if device == "cuda":
        for ct in ["float16", "int8_float16", "int8"]:
            try:
                model = WhisperModel(model_path, device="cuda", compute_type=ct, **kwargs)
                with open(err_path, "a", encoding="utf-8") as ef:
                    ef.write(f"CUDA {ct} OK\n")
                break
            except Exception as e:
                with open(err_path, "a", encoding="utf-8") as ef:
                    ef.write(f"CUDA {ct} fail: {e}\n")
    if model is None:
        actual_device = "cpu"
        try:
            model = WhisperModel(model_path, device="cpu", compute_type="int8", **kwargs)
        except Exception as e:
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(f"ERROR|Model load failed|{e}\n")
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Model load error: {e}\n{traceback.format_exc()}")
            return
    # v2.8: Batched inference pipeline (GPU parallel processing)
    use_batched = batch.get("batched", False) and actual_device == "cuda"
    transcriber = model
    if use_batched:
        try:
            from faster_whisper import BatchedInferencePipeline
            transcriber = BatchedInferencePipeline(model=model)
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write("BatchedInferencePipeline enabled\n")
        except Exception as e:
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Batched init fail (using standard): {e}\n")
            transcriber = model
            use_batched = False
    # v2.8: Read new parameters from batch
    no_prev_text = batch.get("no_prev_text", False)
    word_timestamps = batch.get("word_timestamps", True)
    rep_penalty = batch.get("rep_penalty", False)
    # v2.9.5: 従来 vad_filter=True ハードコード。実測でONだと単語が40個消え開始時刻が早い側にズレる
    # (OFFにするとopenai-whisperとタイミング完全一致)。既定OFF、UIのチェックで切替。
    vad_filter = batch.get("vad_filter", False)
    prompt_text = batch.get("prompt", "")
    hotwords_text = batch.get("hotwords", "")
    morph_split = batch.get("morph_split", False)
    maxchars = batch.get("maxchars", 0)
    max_lines = int(batch.get("maxlines", 1) or 1)  # v2.8.5
    if morph_split:
        word_timestamps = True  # v2.8.2a: 実測割当に単語時刻が必須
    # v2.9.3: kotoba/distil 系モデルは config.json の alignment_heads が元モデル(large-v2/v3, デコーダ32層)
    # のままコピーされている。実際のデコーダは蒸留で2層しかないため、存在しない第7〜25層を参照して
    # CTranslate2 のネイティブ層で範囲外アクセスが起き、Pythonの例外にすらならずプロセスごと落ちる。
    # (実測: kotoba-whisper + 文節区切りON で「CUDA float16 OK」の直後に無言終了、結果ファイル未生成)
    # → 単語時刻を諦めて文字数比の按分にフォールバックする。morph_split の強制ONより後に置いて上書きすること。
    # seg.words は None になるが getattr(seg,"words",None) ガードで wl=[] となり按分経路に入る(既存動作)。
    # openai-whisper 側はそもそも kotoba/distil を弾いている(_run_openai_whisper 冒頭)ため対処不要。
    if word_timestamps and (model_size.startswith("distil-") or model_size.startswith("kotoba")):
        word_timestamps = False
        with open(err_path, "a", encoding="utf-8") as ef:
            ef.write(f"WARN: {model_size} has mismatched alignment_heads -> word_timestamps disabled (timing falls back to proportional)\n")
    morph_stats = {}
    # v2.8: Temperature fallback tuple (try 0 first, then escalate)
    if temperature == 0:
        temp_param = (0.0, 0.2, 0.4, 0.6, 0.8, 1.0)
    else:
        temp_param = float(temperature)
    with open(err_path, "a", encoding="utf-8") as ef:
        ef.write(f"v2.8 params: no_prev_text={no_prev_text} word_ts={word_timestamps} rep_penalty={rep_penalty} vad_filter={vad_filter} batched={use_batched} hotwords={hotwords_text[:50]} prompt={prompt_text[:50]}\n")
    results = []
    filtered_count = 0
    for ci, clip in enumerate(clips):
        wav_path = clip["wav"]
        tl_start = clip["timeline_start"]
        fps = clip["fps"]
        if not os.path.exists(wav_path):
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Clip {ci}: wav not found: {wav_path}\n")
            continue
        try:
            transcribe_kwargs = {
                "language": language,
                "beam_size": beam_size,
                "vad_filter": vad_filter,
                "word_timestamps": word_timestamps,
                "temperature": temp_param,
                "condition_on_previous_text": not no_prev_text,
                "vad_parameters": {"min_silence_duration_ms": 500, "speech_pad_ms": 300},  # vad_filter=False のときは無視される
            }
            if prompt_text:
                transcribe_kwargs["initial_prompt"] = prompt_text
            if rep_penalty:
                transcribe_kwargs["repetition_penalty"] = 1.2
                transcribe_kwargs["no_repeat_ngram_size"] = 3
            # v2.8: Hotwords (boost specific words in decoder)
            if hotwords_text:
                transcribe_kwargs["hotwords"] = hotwords_text
            # v2.8: Hallucination silence threshold
            transcribe_kwargs["hallucination_silence_threshold"] = 2.0
            # Use batched pipeline or standard model
            segments, info = transcriber.transcribe(wav_path, **transcribe_kwargs)
            runs = _speech_runs(wav_path) if morph_split else []
            for seg in segments:
                # v2.8: Segment quality filter
                # Skip low-confidence segments and likely non-speech
                if hasattr(seg, 'avg_logprob') and seg.avg_logprob < -1.0:
                    filtered_count += 1
                    continue
                if hasattr(seg, 'no_speech_prob') and seg.no_speech_prob > 0.6:
                    filtered_count += 1
                    continue
                sf = tl_start + int(seg.start * fps)
                ef2 = tl_start + int(seg.end * fps)
                text = seg.text.strip()
                wl = []
                if morph_split and getattr(seg, "words", None):
                    wl = [(w.word, w.start, w.end) for w in seg.words]
                if text and ef2 > sf:
                    _emit_seg(results, ci, sf, ef2, text, morph_split, maxchars, wl, tl_start, fps, morph_stats, runs, max_lines)
        except Exception as e:
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Clip {ci} err: {e}\n{traceback.format_exc()}")
    with open(output_path, "w", encoding="utf-8") as f:
        for line in results:
            f.write(line + "\n")
    with open(err_path, "a", encoding="utf-8") as ef:
        ef.write(f"Morph timing: word-aligned={morph_stats.get('word', 0)} proportional={morph_stats.get('prop', 0)} snap-moved={morph_stats.get('snap', 0)} vad={_vad_mode}\n")
        ef.write(f"Done: {len(results)} segs, {filtered_count} filtered ({actual_device}{'|batched' if use_batched else ''})\n")

)PYHELPER" R"PYHELPER(
def _run_openai_whisper(model_size, language, device, clips, model_dir, output_path, err_path, beam_size=5, temperature=0, batch=None):
    if batch is None: batch = {}
    # v2.8: Distil/Kotoba models not supported in openai-whisper
    if model_size.startswith("distil-") or model_size.startswith("kotoba"):
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(f"ERROR|{model_size} requires faster-whisper backend|Switch backend to faster-whisper\n")
        return
    try:
        import whisper
    except ImportError as e:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(f"ERROR|whisper not installed|pip install openai-whisper\n")
        with open(err_path, "a", encoding="utf-8") as ef:
            ef.write(f"ImportError: {e}\n")
        return
    with open(err_path, "a", encoding="utf-8") as ef:
        ef.write(f"Loading: {model_size} device={device} (openai-whisper)\n")
    # Map model names for openai-whisper
    model_name = model_size
    if model_size == "large-v3-turbo":
        model_name = "turbo"
    try:
        dl_dir = model_dir if model_dir else None
        model = whisper.load_model(model_name, device=device, download_root=dl_dir)
    except RuntimeError as e:
        if device == "cuda" and "CUDA" in str(e):
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"CUDA failed, falling back to CPU: {e}\n")
            device = "cpu"
            model = whisper.load_model(model_name, device="cpu", download_root=dl_dir)
        else:
            raise
    except Exception as e:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(f"ERROR|Model load failed|{e}\n")
        with open(err_path, "a", encoding="utf-8") as ef:
            ef.write(f"Model load error: {e}\n{traceback.format_exc()}")
        return
    # v2.8: Read new parameters
    no_prev_text = batch.get("no_prev_text", False)
    word_timestamps = batch.get("word_timestamps", True)
    prompt_text = batch.get("prompt", "")
    morph_split = batch.get("morph_split", False)
    maxchars = batch.get("maxchars", 0)
    max_lines = int(batch.get("maxlines", 1) or 1)  # v2.8.5
    if morph_split:
        word_timestamps = True  # v2.8.2a: 実測割当に単語時刻が必須
    morph_stats = {}
    # v2.8: Temperature fallback tuple
    if temperature == 0:
        temp_param = (0.0, 0.2, 0.4, 0.6, 0.8, 1.0)
    else:
        temp_param = float(temperature)
    results = []
    filtered_count = 0
    for ci, clip in enumerate(clips):
        wav_path = clip["wav"]
        tl_start = clip["timeline_start"]
        fps = clip["fps"]
        if not os.path.exists(wav_path):
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Clip {ci}: wav not found: {wav_path}\n")
            continue
        try:
            opts = {"language": language, "beam_size": beam_size, "verbose": False, "word_timestamps": word_timestamps, "temperature": temp_param, "condition_on_previous_text": not no_prev_text}
            if language is None:
                del opts["language"]
            if prompt_text:
                opts["initial_prompt"] = prompt_text
            result = model.transcribe(wav_path, **opts)
            runs = _speech_runs(wav_path) if morph_split else []
            for seg in result.get("segments", []):
                # v2.8: Segment quality filter
                if seg.get("avg_logprob", 0) < -1.0:
                    filtered_count += 1
                    continue
                if seg.get("no_speech_prob", 0) > 0.6:
                    filtered_count += 1
                    continue
                sf = tl_start + int(seg["start"] * fps)
                ef2 = tl_start + int(seg["end"] * fps)
                text = seg["text"].strip()
                wl = []
                if morph_split:
                    wl = [(w.get("word", ""), w.get("start"), w.get("end")) for w in seg.get("words", [])]
                if text and ef2 > sf:
                    _emit_seg(results, ci, sf, ef2, text, morph_split, maxchars, wl, tl_start, fps, morph_stats, runs, max_lines)
        except Exception as e:
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Clip {ci} err: {e}\n{traceback.format_exc()}")
    with open(output_path, "w", encoding="utf-8") as f:
        for line in results:
            f.write(line + "\n")
    with open(err_path, "a", encoding="utf-8") as ef:
        ef.write(f"Morph timing: word-aligned={morph_stats.get('word', 0)} proportional={morph_stats.get('prop', 0)} snap-moved={morph_stats.get('snap', 0)} vad={_vad_mode}\n")
        ef.write(f"Done: {len(results)} segs, {filtered_count} filtered (openai-whisper, {device})\n")

if __name__ == "__main__":
    main()
)PYHELPER";

static void EnsurePyHelper(){
    std::string p = GetPluginDir() + "\\whisper_helper.py";
    // Only overwrite if content changed (preserve user modifications if version matches)
    std::string existing;
    {
        std::ifstream rf(Utf8ToWide(p), std::ios::binary);
        if(rf.is_open()){
            existing = std::string((std::istreambuf_iterator<char>(rf)), std::istreambuf_iterator<char>());
        }
    }
    std::string embedded(g_pyHelper);
    if(existing != embedded){
        std::ofstream f(Utf8ToWide(p), std::ios::binary);
        if(f.is_open()){ f << g_pyHelper; f.close(); }
    }
}
static void EnsureDirectories(){
    CreateDirU(GetPluginDir());
    CreateDirU(GetTempDir());
    CreateDirU(GetModelsDir());
    CreateDirU(GetSitePackagesDir());
}

// =========================================================================
// Python / ffmpeg detection
// =========================================================================

static std::string FindPythonAuto(){
    std::string pp = GetPluginDir() + "\\python\\python.exe";
    if(FileExistsU(pp)) return pp;
    wchar_t up[MAX_PATH];
    if(GetEnvironmentVariableW(L"LOCALAPPDATA", up, MAX_PATH) > 0){
        // v2.9.0: v2.9.11 で「3.13は非推奨」として後回しにしたが、**その前提が誤りだった**ため撤回。
        // PyTorch 2.6以降(2025-01)で通常版の Python 3.13(cp313)は正式対応済み。
        // 入らなかった真因はプラグインが古い cu121 インデックスに固定されていたことで、
        // そちらは DetectCudaTag() の拡張で解消した。よって新しい版を優先する素直な降順に戻す。
        // (非対応なのは 3.13t = フリースレッド版だが、python.org の通常インストーラでは入らない)
        for(int v = 14; v >= 9; v--){
            std::wstring c = std::wstring(up) + L"\\Programs\\Python\\Python3" + std::to_wstring(v) + L"\\python.exe";
            if(GetFileAttributesW(c.c_str()) != INVALID_FILE_ATTRIBUTES) return WideToUtf8(c);
        }
    }
    wchar_t buf[MAX_PATH];
    if(SearchPathW(NULL, L"python.exe", NULL, MAX_PATH, buf, NULL) > 0){
        // v2.9.7【重要】Microsoft Store の「アプリ実行エイリアス」(プレースホルダ)を弾く。
        // Windows は %LOCALAPPDATA%\Microsoft\WindowsApps\python.exe という**中身の無いスタブ**を
        // 全ユーザーのPATHに標準で置いている。実行すると Store のページを開くだけで何もしない。
        // これを拾うと「Pythonは見つかったのに pip も -c も全部失敗する」という最悪の壊れ方をする
        // (実測: 新規アカウントでセットアップすると torch/whisper/fugashi/モデルが全滅するのに
        //  「初期設定完了」と表示され、ユーザーは永久に同じ導線を繰り返すことになった)。
        std::string found = WideToUtf8(buf);
        std::string lower = found;
        for(auto& c : lower) c = (char)tolower((unsigned char)c);
        if(lower.find("\\windowsapps\\") == std::string::npos)
            return found;
        // スタブだった場合は「見つからなかった」扱いにする(未検出として正しく案内する)
    }
    return "";
}
static std::string GetEffectivePython(){
    if(!g_pythonPath.empty() && FileExistsU(g_pythonPath))
        return g_pythonPath;
    return FindPythonAuto();
}
// v2.8.8【B】: DownloadFFmpeg()末尾のバージョン付きフォルダ再帰探索を切り出したもの。
// GetEffectiveFFmpeg()からも呼び、自前DL済み(GetPluginDir()\ffmpeg\<version>\bin\ffmpeg.exe)を再DL前に検出する。
static std::string FindBundledFFmpeg(){
    std::string dest = GetPluginDir() + "\\ffmpeg";
    std::string found;
    WIN32_FIND_DATAW fd;
    std::wstring pat = Utf8ToWide(dest) + L"\\*";
    HANDLE h = FindFirstFileW(pat.c_str(), &fd);
    if(h != INVALID_HANDLE_VALUE){
        do{
            if((fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && fd.cFileName[0] != L'.'){
                std::string cand = dest + "\\" + WideToUtf8(fd.cFileName) + "\\bin\\ffmpeg.exe";
                if(FileExistsU(cand)){ found = cand; break; }
            }
        } while(FindNextFileW(h, &fd));
        FindClose(h);
    }
    if(found.empty()){
        std::string flat = dest + "\\bin\\ffmpeg.exe";
        if(FileExistsU(flat)) found = flat;
    }
    return found;
}
static std::string GetEffectiveFFmpeg(){
    if(!g_ffmpegPath.empty() && FileExistsU(g_ffmpegPath))
        return g_ffmpegPath;
    std::string bundled = FindBundledFFmpeg(); // v2.8.8【B】: 自前DL済みffmpegを優先探索(再DL防止)
    if(!bundled.empty()) return bundled;
    std::string def = GetExeDir() + "\\ffmpeg.exe";
    if(FileExistsU(def)) return def;
    wchar_t buf[MAX_PATH];
    if(SearchPathW(NULL, L"ffmpeg.exe", NULL, MAX_PATH, buf, NULL) > 0) return WideToUtf8(buf);
    return "";
}

// =========================================================================
// RunProcess helper (captures stdout+stderr)
// =========================================================================

static bool RunProcess(const std::wstring& cmdLine, std::string& output, DWORD timeoutMs = 300000){
    SECURITY_ATTRIBUTES sa = {sizeof(sa), NULL, TRUE};
    HANDLE hReadOut, hWriteOut;
    CreatePipe(&hReadOut, &hWriteOut, &sa, 0);
    SetHandleInformation(hReadOut, HANDLE_FLAG_INHERIT, 0);
    STARTUPINFOW si = {sizeof(si)};
    si.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
    si.wShowWindow = SW_HIDE;
    si.hStdOutput = hWriteOut; si.hStdError = hWriteOut; si.hStdInput = NULL;
    PROCESS_INFORMATION pi = {};
    std::wstring cmd = cmdLine;
    if(!CreateProcessW(NULL, &cmd[0], NULL, NULL, TRUE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi)){
        CloseHandle(hReadOut); CloseHandle(hWriteOut);
        output = "CreateProcess failed: " + std::to_string(GetLastError());
        return false;
    }
    CloseHandle(hWriteOut);
    output.clear();
    char buf[4096]; DWORD br;
    // v2.8b: timeoutMs is now enforced via elapsed-time tracking (GetTickCount64).
    ULONGLONG startTick = GetTickCount64();
    bool timedOut = false;
    while(true){
        DWORD wr = WaitForSingleObject(pi.hProcess, 500);
        DWORD avail = 0;
        PeekNamedPipe(hReadOut, NULL, 0, NULL, &avail, NULL);
        while(avail > 0){
            DWORD toRead = (avail < sizeof(buf)-1) ? avail : sizeof(buf)-1;
            if(ReadFile(hReadOut, buf, toRead, &br, NULL) && br > 0){
                buf[br] = 0; output += buf; avail -= br;
            } else break;
            PeekNamedPipe(hReadOut, NULL, 0, NULL, &avail, NULL);
        }
        if(wr != WAIT_TIMEOUT) break;
        if(timeoutMs > 0 && (GetTickCount64() - startTick) >= (ULONGLONG)timeoutMs){
            timedOut = true;
            TerminateProcess(pi.hProcess, 1);
            WaitForSingleObject(pi.hProcess, 2000);
            break;
        }
    }
    // drain remaining
    DWORD avail2 = 0; PeekNamedPipe(hReadOut, NULL, 0, NULL, &avail2, NULL);
    while(avail2 > 0){
        if(ReadFile(hReadOut, buf, sizeof(buf)-1, &br, NULL) && br > 0){
            buf[br] = 0; output += buf;
        } else break;
        PeekNamedPipe(hReadOut, NULL, 0, NULL, &avail2, NULL);
    }
    CloseHandle(hReadOut);
    if(timedOut){
        CloseHandle(pi.hProcess); CloseHandle(pi.hThread);
        output += "\nTIMEOUT: process exceeded " + std::to_string(timeoutMs) + "ms and was terminated.\n";
        return false;
    }
    DWORD ec = 1; GetExitCodeProcess(pi.hProcess, &ec);
    CloseHandle(pi.hProcess); CloseHandle(pi.hThread);
    return (ec == 0);
}

// =========================================================================
// INI save/load
// =========================================================================

static void SaveSettings(){
    std::ofstream f(Utf8ToWide(GetIniPath()));
    if(!f.is_open()) return;
    f << "[Settings]\n";
    f << "template=" << g_templatePath << "\n";
    char buf[16];
    GetWindowTextA(g_layerEdit, buf, sizeof(buf)); f << "layer=" << buf << "\n";
    GetWindowTextA(g_maxCharEdit, buf, sizeof(buf)); f << "maxchars=" << buf << "\n";
    f << "model=" << SendMessageA(g_modelCombo, CB_GETCURSEL, 0, 0) << "\n";
    f << "device=" << SendMessageA(g_deviceCombo, CB_GETCURSEL, 0, 0) << "\n";
    f << "backend=" << SendMessageA(g_backendCombo, CB_GETCURSEL, 0, 0) << "\n";
    f << "language=" << SendMessageA(g_langCombo, CB_GETCURSEL, 0, 0) << "\n";
    char qBuf[16] = {}; GetWindowTextA(g_qualityEdit, qBuf, sizeof(qBuf));
    f << "quality=" << qBuf << "\n";
    char tBuf[16] = {}; GetWindowTextA(g_tempEdit, tBuf, sizeof(tBuf));
    f << "temperature=" << tBuf << "\n";
    f << "remove_punct=" << (SendMessageA(g_chkRemovePunct, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n";
    f << "remove_exclam=" << (SendMessageA(g_chkRemoveExclam, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n";
    f << "normalize=" << (SendMessageA(g_chkNormalize, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n";
    f << "merge_seg=" << (SendMessageA(g_chkMergeSeg, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n"; // v2.8.2
    f << "morph_split=" << (SendMessageA(g_chkMorphSplit, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n"; // v2.8.2
    f << "two_line=" << (SendMessageA(g_chkTwoLine, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n"; // v2.8.5
    f << "ffmpeg=" << g_ffmpegPath << "\n";
    f << "python=" << g_pythonPath << "\n";
    f << "fw_sp=" << g_fwSpPath << "\n";
    f << "ow_sp=" << g_owSpPath << "\n";
    char lBuf[16] = {}; GetWindowTextA(g_lingerEdit, lBuf, sizeof(lBuf));
    f << "linger=" << lBuf << "\n";
    char leadBuf2[16] = {}; GetWindowTextA(g_leadEdit, leadBuf2, sizeof(leadBuf2));
    f << "lead=" << leadBuf2 << "\n"; // v2.8.2a
    // v2.8 settings
    f << "no_prev_text=" << (SendMessageA(g_chkNoPrevText, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n";
    f << "word_ts=" << (SendMessageA(g_chkWordTs, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n";
    f << "rep_penalty=" << (SendMessageA(g_chkRepPenalty, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n";
    f << "vad=" << (SendMessageA(g_chkVad, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n"; // v2.9.5
    {
        int len = GetWindowTextLengthA(g_promptEdit);
        if(len > 0){
            std::string pBuf(len + 1, '\0');
            GetWindowTextA(g_promptEdit, &pBuf[0], len + 1);
            pBuf.resize(len);
            // Escape newlines for INI
            std::string esc;
            for(char c : pBuf){ if(c == '\n') esc += "\\n"; else if(c == '\r'){} else esc += c; }
            f << "prompt=" << esc << "\n";
        } else {
            f << "prompt=\n";
        }
    }
    // v2.8 settings
    f << "batched=" << (SendMessageA(g_chkBatched, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n";
    {
        int len = GetWindowTextLengthA(g_hotwordsEdit);
        if(len > 0){
            std::string hBuf(len + 1, '\0');
            GetWindowTextA(g_hotwordsEdit, &hBuf[0], len + 1);
            hBuf.resize(len);
            f << "hotwords=" << hBuf << "\n";
        } else {
            f << "hotwords=\n";
        }
    }
}
static void LoadSettings(){
    std::ifstream f(Utf8ToWide(GetIniPath()));
    if(!f.is_open()) return;
    std::string line;
    while(std::getline(f, line)){
        if(line.empty() || line[0] == '[') continue;
        size_t eq = line.find('='); if(eq == std::string::npos) continue;
        std::string key = line.substr(0, eq), val = line.substr(eq+1);
        while(!val.empty() && (val.back()=='\r'||val.back()=='\n')) val.pop_back();
        if(key=="template" && !val.empty()) g_templatePath = val;
        else if(key=="layer") SetWindowTextA(g_layerEdit, val.c_str());
        else if(key=="maxchars") SetWindowTextA(g_maxCharEdit, val.c_str());
        else if(key=="model") SendMessageA(g_modelCombo, CB_SETCURSEL, atoi(val.c_str()), 0);
        else if(key=="device") SendMessageA(g_deviceCombo, CB_SETCURSEL, atoi(val.c_str()), 0);
        else if(key=="backend") SendMessageA(g_backendCombo, CB_SETCURSEL, atoi(val.c_str()), 0);
        else if(key=="language") SendMessageA(g_langCombo, CB_SETCURSEL, atoi(val.c_str()), 0);
        else if(key=="quality") SetWindowTextA(g_qualityEdit, val.c_str());
        else if(key=="temperature") SetWindowTextA(g_tempEdit, val.c_str());
        else if(key=="remove_punct") SendMessageA(g_chkRemovePunct, BM_SETCHECK, atoi(val.c_str()) ? BST_CHECKED : BST_UNCHECKED, 0);
        else if(key=="remove_exclam") SendMessageA(g_chkRemoveExclam, BM_SETCHECK, atoi(val.c_str()) ? BST_CHECKED : BST_UNCHECKED, 0);
        else if(key=="normalize") SendMessageA(g_chkNormalize, BM_SETCHECK, atoi(val.c_str()) ? BST_CHECKED : BST_UNCHECKED, 0);
        else if(key=="merge_seg") SendMessageA(g_chkMergeSeg, BM_SETCHECK, atoi(val.c_str()) ? BST_CHECKED : BST_UNCHECKED, 0); // v2.8.2
        else if(key=="morph_split") SendMessageA(g_chkMorphSplit, BM_SETCHECK, atoi(val.c_str()) ? BST_CHECKED : BST_UNCHECKED, 0); // v2.8.2
        else if(key=="two_line") SendMessageA(g_chkTwoLine, BM_SETCHECK, atoi(val.c_str()) ? BST_CHECKED : BST_UNCHECKED, 0); // v2.8.5
        else if(key=="ffmpeg") g_ffmpegPath = val;
        else if(key=="python") g_pythonPath = val;
        else if(key=="fw_sp") g_fwSpPath = val;
        else if(key=="ow_sp") g_owSpPath = val;
        else if(key=="linger") SetWindowTextA(g_lingerEdit, val.c_str());
        else if(key=="lead") SetWindowTextA(g_leadEdit, val.c_str()); // v2.8.2a
        // v2.8 settings
        else if(key=="no_prev_text") SendMessageA(g_chkNoPrevText, BM_SETCHECK, atoi(val.c_str()) ? BST_CHECKED : BST_UNCHECKED, 0);
        else if(key=="word_ts") SendMessageA(g_chkWordTs, BM_SETCHECK, atoi(val.c_str()) ? BST_CHECKED : BST_UNCHECKED, 0);
        else if(key=="rep_penalty") SendMessageA(g_chkRepPenalty, BM_SETCHECK, atoi(val.c_str()) ? BST_CHECKED : BST_UNCHECKED, 0);
        else if(key=="vad") SendMessageA(g_chkVad, BM_SETCHECK, atoi(val.c_str()) ? BST_CHECKED : BST_UNCHECKED, 0); // v2.9.5
        else if(key=="prompt"){
            // Unescape newlines from INI
            std::string unesc;
            for(size_t i = 0; i < val.size(); i++){
                if(val[i] == '\\' && i+1 < val.size() && val[i+1] == 'n'){ unesc += '\n'; i++; }
                else unesc += val[i];
            }
            SetWindowTextA(g_promptEdit, unesc.c_str());
        }
        // v2.8 settings
        else if(key=="batched") SendMessageA(g_chkBatched, BM_SETCHECK, atoi(val.c_str()) ? BST_CHECKED : BST_UNCHECKED, 0);
        else if(key=="hotwords") SetWindowTextA(g_hotwordsEdit, val.c_str());
    }
}

// =========================================================================
// Setup thread (auto-install faster-whisper + model DL)
// =========================================================================

// Forward declarations (defined later, after timeline code)
static void SetStatus(const std::string& msg);
static void SetProgress(int val);
static void UpdateWhisperLocLabels();
static void UpdateFugashiStatus(); // v2.8.5
static void SetStatusW(const wchar_t* msg); // v2.8.10【I】: g_status/g_statusSetup同時更新版(定義はSetStatus/SetProgressの近く)
static void ProbeThread(); // v2.9.10: SetupThread末尾から再実測のため呼ぶ(定義は後方)

// v2.9.0【D】: g_chkForceReinstall(1個)を項目別チェック5個に置き換えたための共通ヘルパー。
static bool IsChecked(HWND h){ return h && SendMessageA(h, BM_GETCHECK, 0, 0) == BST_CHECKED; }
static void ClearSetupChecks(){ // v2.9.0【D】: SetupThreadの3つの終了経路すべてから呼ぶ(誤って2回目が走るのを防ぐ)
    HWND cs[] = {g_chkFfmpeg, g_chkTorch, g_chkWhisper, g_chkFaster, g_chkFugashi}; // v2.9.1: g_chkModelは廃止 / v2.9.4: g_chkFfmpeg追加
    for(HWND c : cs) if(c) SendMessageA(c, BM_SETCHECK, BST_UNCHECKED, 0);
}

// v2.9.0【G】: モデル名テーブル。従来 SetupThread 内にローカルで持っていたもの(const char* mn[])を
// ファイルスコープへ切り出し、SetupThread と ModelExists() の両方から同じ表を参照する。
static const char* kModelNames[] = {"tiny","base","small","medium","large-v3","large-v3-turbo","kotoba-whisper"};

// v2.9.0【G】: SetupThread 内にインラインで書かれていたモデル存在判定を、環境タブ表示からも
// 使えるように切り出したもの。判定内容は従来と完全に同一(faster-whisper の config.json / model.bin、
// openai-whisper の .pt、kotoba の HF キャッシュ)。
static bool ModelExists(const std::string& mName){
    std::string mDir = GetModelsDir();
    std::string localModel = mDir + "\\" + mName;
    bool exists = FileExistsU(localModel + "\\config.json")   // faster-whisper
        || FileExistsU(localModel + "\\model.bin")              // faster-whisper alt
        || FileExistsU(mDir + "\\" + mName + ".pt");            // openai-whisper
    // v2.8: distil/kotoba モデルの HuggingFace キャッシュ判定
    if(!exists && mName.find("kotoba") == 0){
        std::string hfCache = mDir + "\\models--kotoba-tech--kotoba-whisper-v2.0-faster";
        exists = FileExistsU(hfCache);
    }
    return exists;
}

// v2.8b: centralizes the g_busy flag + Generate/Setup button enable-state together so every
// exit path (including the four early-return guards in TranscribeThread) re-enables the UI.
// EnableWindow() is safe to call from a background thread (it just posts internally).
// v2.9.5: バックエンドに応じてVADチェックの有効/無効を切り替える。
// backend コンボの index 1 = openai-whisper (それ以外 = faster-whisper)。
static void UpdateVadEnable(){
    if(!g_backendCombo) return;
    int bi = SendMessageA(g_backendCombo, CB_GETCURSEL, 0, 0);
    BOOL isFaster = (bi == 1) ? FALSE : TRUE; // index 1 = openai-whisper
    if(g_chkVad) EnableWindow(g_chkVad, isFaster);
    // v2.9.6【配布】繰り返し抑制も faster-whisper 専用。Python側で rep_penalty を読むのは
    // _run_faster_whisper 内だけで、_run_openai_whisper は一度も参照しない
    // (repetition_penalty は CTranslate2 の機能で openai-whisper には無い)。
    // 既定バックエンドを openai-whisper にした以上、大半のユーザーが「押しても何も起きない
    // チェック」を触ることになるためグレーアウトする。
    if(g_chkRepPenalty) EnableWindow(g_chkRepPenalty, isFaster);
}

static void SetBusy(bool busy){
    g_busy = busy;
    if(g_btnGenerate) EnableWindow(g_btnGenerate, busy ? FALSE : TRUE);
    if(g_btnSetup) EnableWindow(g_btnSetup, busy ? FALSE : TRUE);
}

// v2.8.8: nvidia-smi の "CUDA Version: X.Y" を読んで、入れるべき torch のビルドタグを決める。
// 戻り値: "cu121" / "cu118" / "cpu"
static std::string DetectCudaTag(){
    std::string out;
    if(!RunProcess(L"nvidia-smi", out, 30000)) return "cpu"; // NVIDIA GPU 無し
    size_t p = out.find("CUDA Version:");
    if(p == std::string::npos) return "cpu";
    double ver = atof(out.c_str() + p + 13); // "CUDA Version: 13.2" → 13.2
    // v2.9.0【重要】新しいCUDAビルドまで選べるように拡張。
    // 従来は上限が cu121 で、**cu121 のホイールは PyTorch 2.5.x 世代で更新が止まっており
    // cp313 (Python 3.13) 用が存在しない**。そのため 3.13 環境では torch のインストールが
    // 必ず失敗していた。これを「Python 3.13が非推奨」と誤解していたが、原因は
    // プラグイン側が古いインデックスに固定されていたこと。実測(2026-07-26):
    //   cu121: cp38-cp312 / cu124: cp38-cp313 / cu126・cu128・cu129: cp39-cp314
    // ドライバが対応する範囲で新しい方を選べば 3.13/3.14 も普通に使える。
    // なお nvidia-smi の "CUDA Version" はドライバが対応する上限なので、
    // それ以下のビルドは下位互換で動く。
    if(ver >= 12.8) return "cu128";
    if(ver >= 12.6) return "cu126";
    if(ver >= 12.4) return "cu124";
    if(ver >= 12.1) return "cu121";
    if(ver >= 11.8) return "cu118";
    return "cpu";
}

// v2.8.8: ffmpeg を公式配布元から取得して プラグインフォルダ\ffmpeg\ に展開する。
// 成功したら ffmpeg.exe のフルパスを返す。失敗時は空文字。
static std::string DownloadFFmpeg(){
    std::string dest = GetPluginDir() + "\\ffmpeg";
    std::string zip  = GetTempDir() + "\\ffmpeg_dl.zip";
    std::string url  = "https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip";
    std::string ps = "powershell -NoProfile -ExecutionPolicy Bypass -Command \""
        "$ErrorActionPreference='Stop';"
        "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;"
        "Invoke-WebRequest -Uri '" + url + "' -OutFile '" + zip + "' -UseBasicParsing;"
        "Expand-Archive -Path '" + zip + "' -DestinationPath '" + dest + "' -Force;"
        "Remove-Item '" + zip + "' -Force\"";
    std::string out;
    if(!RunProcess(Utf8ToWide(ps), out, 900000)){
        DebugLog("ffmpeg download failed:\n" + out);
        return "";
    }
    // v2.8.8【B】: 再帰探索ロジックはFindBundledFFmpeg()に切り出し済み(GetEffectiveFFmpeg()と共有、重複排除)
    std::string found = FindBundledFFmpeg();
    if(!found.empty()){
        // v2.8.8【D】: ffplay.exe/ffprobe.exe(計約204MB)はソースから未使用のため削除。ffmpeg.exeは消さない。削除失敗は無視
        size_t p = found.find_last_of("\\/");
        if(p != std::string::npos){
            std::string binDir = found.substr(0, p);
            DeleteFileU(binDir + "\\ffplay.exe");
            DeleteFileU(binDir + "\\ffprobe.exe");
        }
        // v2.9.9【配布】ffmpeg.exe を aviutl2.exe と同じ場所へ移す。
        // 理由(開発者の実地観察): AviUtl2ユーザーの多くは元々 ffmpeg を本体と同じ場所に置いており、
        // それが事実上の慣例。そこに置けば他プラグイン(NVEnc/x264guiEx等)からも共有できる。
        // **上書きの心配は無い**: このDLが走るのは GetEffectiveFFmpeg() が空=どこにも
        // ffmpeg が無いときだけなので、移動先に既存ファイルは存在しない。
        // 失敗時(本体フォルダが書込不可など)は移動せず、従来どおりプラグインフォルダの物を使う
        // (GetEffectiveFFmpeg() は本体隣→PATH の前にプラグインフォルダを探すのでどちらでも動く)。
        std::string exeSide = GetExeDir() + "\\ffmpeg.exe";
        if(!FileExistsU(exeSide)){
            if(MoveFileW(Utf8ToWide(found).c_str(), Utf8ToWide(exeSide).c_str())){
                DebugLog("ffmpeg moved to exe dir: " + exeSide);
                // 展開に使った作業フォルダはもう不要
                RemoveDirectoryTreeU(GetPluginDir() + "\\ffmpeg");
                return exeSide;
            }
            DebugLog("ffmpeg move to exe dir failed, keeping: " + found);
        }
    }
    return found;
}

// v2.8.8【C】: zipファイルサイズをGetFileAttributesExWで覗いて実進捗を出す(ファイルは開かない=DL中の共有違反回避)
static bool GetFileSizeU(const std::string& path, unsigned long long& outSize){
    WIN32_FILE_ATTRIBUTE_DATA fad;
    if(!GetFileAttributesExW(Utf8ToWide(path).c_str(), GetFileExInfoStandard, &fad))
        return false;
    outSize = (((unsigned long long)fad.nFileSizeHigh) << 32) | (unsigned long long)fad.nFileSizeLow;
    return true;
}
// v2.8.8【C】: DownloadFFmpeg()を別スレッドで実行しつつ、呼び出し元スレッドで500ms毎にzipサイズをポーリングして
// SetStatus/SetProgressを更新する。PowerShell側(-UseBasicParsing)は無改造。0-85%=DL中、90%=展開中、100%=完了
static std::string DownloadFFmpegWithProgress(){
    std::string zip = GetTempDir() + "\\ffmpeg_dl.zip";
    const unsigned long long kExpected = 104ULL * 1024 * 1024; // 約104MB
    std::atomic<bool> done{false};
    std::string result;
    std::thread worker([&](){
        result = DownloadFFmpeg();
        done = true;
    });
    bool sawZip = false;
    while(!done.load()){
        Sleep(500);
        unsigned long long sz = 0;
        if(GetFileSizeU(zip, sz)){
            sawZip = true;
            double pct = (double)sz / (double)kExpected * 85.0;
            if(pct > 85.0) pct = 85.0;
            SetProgress((int)pct);
            SetStatus("ffmpeg \xe3\x83\x80\xe3\x82\xa6\xe3\x83\xb3\xe3\x83\xad\xe3\x83\xbc\xe3\x83\x89\xe4\xb8\xad... "
                + std::to_string((int)(sz / (1024 * 1024))) + "MB / 104MB");
        } else if(sawZip){
            // v2.8.8: zipが消滅 = Remove-Item済み(展開完了直後)か展開中。区別できないのでまとめて「展開中」表示
            SetProgress(90);
            SetStatus("ffmpeg \xe5\xb1\x95\xe9\x96\x8b\xe4\xb8\xad...");
        }
    }
    worker.join();
    return result;
}

// v2.9.7【進捗バー修正】SetupThread内の進捗更新は必ず SetProgress() を通すこと。
// SendMessageA(g_progress, ...) を直接呼ぶと**生成タブのバーだけ**が動き、環境タブ専用の
// g_progressSetup が取り残される。環境タブを開いている間 g_progress は SW_HIDE なので
// 「セットアップ中なのに進捗が全く出ない」「完了後もバーが途中で固まったまま」になる
// (実測: ffmpegだけ進捗が見えたのは DownloadFFmpegWithProgress() が SetProgress() を
//  使っていたため。終了時のリセットも g_progress にしか効かず80%で残っていた)。
// これは v2.8.10【I】で直したステータス凍結(SetWindowTextW(g_status,...)直接呼び)と同じ構造の穴。
static void SetupThread(){
    SetBusy(true);
    // v2.9.0【D】: 旧・全項目を無条件で入れ直すチェックボックス1個を廃止し、
    // torch/whisper/faster-whisper/fugashi/モデルの5項目チェックに置き換え。チェックした項目だけ
    // skip判定(importが通るか等)を無視して入れ直す。誤爆で複数GBの再ダウンロードが走るのを防ぐため、
    // 何が走るかを具体的に見せる確認ダイアログを経由する。
    bool fTorch   = IsChecked(g_chkTorch);
    bool fWhisper = IsChecked(g_chkWhisper);
    bool fFaster  = IsChecked(g_chkFaster);
    bool fFugashi = IsChecked(g_chkFugashi);
    bool fFfmpeg  = IsChecked(g_chkFfmpeg); // v2.9.4
    bool anyForce = fFfmpeg || fTorch || fWhisper || fFaster || fFugashi; // v2.9.1: モデルのチェックは廃止

    int mi = SendMessageA(g_modelCombo, CB_GETCURSEL, 0, 0);
    if(mi == CB_ERR) mi = 0;
    if(mi < 0 || mi >= (int)(sizeof(kModelNames)/sizeof(kModelNames[0]))) mi = 0;
    std::string mName = kModelNames[mi]; // v2.9.0【G】: ファイルスコープのkModelNames[]を参照(旧ローカルmn[]を廃止)

    if(anyForce){
        std::string items;
        if(fFfmpeg)  items += "\n  " "\xe3\x83\xbb" "ffmpeg (" "\xe7\xb4\x84" "104MB)"; // v2.9.4
        if(fTorch)   items += "\n  " "\xe3\x83\xbb" "PyTorch (" "\xe7\xb4\x84" "4.3GB)";
        if(fWhisper) items += "\n  " "\xe3\x83\xbb" "whisper";
        if(fFaster)  items += "\n  " "\xe3\x83\xbb" "faster-whisper";
        if(fFugashi) items += "\n  " "\xe3\x83\xbb" "\xe6\x96\x87\xe7\xaf\x80\xe5\x8c\xba\xe5\x88\x87\xe3\x82\x8a" "(fugashi)";
        int fr = MsgBox(g_wnd,
            "\xe4\xbb\xa5\xe4\xb8\x8b\xe3\x82\x92\xe5\xbc\xb7\xe5\x88\xb6\xe7\x9a\x84\xe3\x81\xab\xe5\x85\xa5\xe3\x82\x8c\xe7\x9b\xb4\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x99\xe3\x80\x82\n" + items +
            "\n\n\xe5\xb0\x8e\xe5\x85\xa5\xe6\xb8\x88\xe3\x81\xbf\xe3\x81\xa7\xe3\x82\x82\xe5\x86\x8d\xe3\x83\x80\xe3\x82\xa6\xe3\x83\xb3\xe3\x83\xad\xe3\x83\xbc\xe3\x83\x89\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x99\xe3\x80\x82\xe5\xae\x9f\xe8\xa1\x8c\xe3\x81\x97\xe3\x81\xa6\xe3\x82\x88\xe3\x82\x8d\xe3\x81\x97\xe3\x81\x84\xe3\x81\xa7\xe3\x81\x99\xe3\x81\x8b\xef\xbc\x9f",
            "\xe7\xa2\xba\xe8\xaa\x8d", MB_YESNO|MB_ICONWARNING);
        if(fr == IDNO){
            ClearSetupChecks(); // v2.9.0【D】: 誤って2回目が走るのを防ぐため毎回OFFに戻す
            SetBusy(false);
            return;
        }
    }
    SetStatusW(L"\x521d\x671f\x8a2d\x5b9a\x4e2d..."); // v2.8.10【I】: SetWindowTextW直接呼びを置き換え(環境タブ凍結対策)
    SetProgress(5);
    std::string report;
    std::string python = GetEffectivePython();
    if(python.empty()){
        MsgBox(g_wnd,
            "Python\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n\n"
            "python.org\xe3\x81\x8b\xe3\x82\x89Python 3.10+\xe3\x82\x92\xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82\n"
            "\xe3\x81\xbe\xe3\x81\x9f\xe3\x81\xaf\xe3\x80\x8cPython\xe9\x81\xb8\xe6\x8a\x9e\xe3\x80\x8d\xe3\x83\x9c\xe3\x82\xbf\xe3\x83\xb3\xe3\x81\xa7\xe6\x8c\x87\xe5\xae\x9a",
            "Python\xe3\x81\x8c\xe5\xbf\x85\xe8\xa6\x81", MB_OK|MB_ICONWARNING);
        SetStatusW(L"Ready (v2.9.0)"); // v2.8.10【I】: SetWindowTextW直接呼びを置き換え(環境タブ凍結対策)+バージョン更新
        SetProgress(0);
        ClearSetupChecks(); // v2.9.0【D】: 異常終了パスでもOFFに戻す
        SetBusy(false);
        return;
    }
    report += "Python: " + python + "\n";

    // v2.9.7【重要】Pythonが「実際に動くか」をここで必ず確かめる。
    // 従来はファイルが存在すればそのまま全工程を走らせていたため、Microsoft Storeのスタブや
    // 壊れたPythonを掴むと **torch/whisper/fugashi/モデルが全部WARNで終わるのに「初期設定完了」と
    // 表示され**、ユーザーは「生成→不足→セットアップ→完了→生成」を永久に繰り返すことになった。
    // 実測(新規アカウント)で確認したこのループを断つのが目的なので、失敗したら必ず中断する。
    {
        std::string vout;
        bool pyRuns = RunProcess(Utf8ToWide("\"" + python + "\" -c \"print(1)\""), vout, 30000);
        if(!pyRuns){
            DebugLog("Python is not executable: " + python + "\n" + vout.substr(0, 300));
            MsgBox(g_wnd,
                "Python\xe3\x81\x8c\xe6\xad\xa3\xe3\x81\x97\xe3\x81\x8f\xe5\x8b\x95\xe4\xbd\x9c\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n\n"
                + python + "\n\n"
                "Microsoft Store\xe3\x81\xae\xe3\x83\x97\xe3\x83\xac\xe3\x83\xbc\xe3\x82\xb9\xe3\x83\x9b\xe3\x83\xab\xe3\x83\x80\xe3\x83\xbc\xe3\x82\x92\xe6\xa4\x9c\xe5\x87\xba\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x84\xe3\x82\x8b\xe5\x8f\xaf\xe8\x83\xbd\xe6\x80\xa7\xe3\x81\x8c\xe3\x81\x82\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x99\xe3\x80\x82\n"
                "python.org \xe3\x81\x8b\xe3\x82\x89 Python 3.10 "
                "\xe4\xbb\xa5\xe4\xb8\x8a\xe3\x82\x92\xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe3\x81\x97\xe3\x80\x81\n"
                "\xe3\x80\x8cPython\xe9\x81\xb8\xe6\x8a\x9e\xe3\x80\x8d\xe3\x83\x9c\xe3\x82\xbf\xe3\x83\xb3\xe3\x81\xa7\xe6\x8c\x87\xe5\xae\x9a\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82",
                "Python\xe3\x81\x8c\xe4\xbd\xbf\xe7\x94\xa8\xe3\x81\xa7\xe3\x81\x8d\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93", MB_OK|MB_ICONERROR);
            SetStatusW(L"Ready (v2.9.0)");
            SetProgress(0);
            ClearSetupChecks();
            SetBusy(false);
            return;
        }
    }

    // v2.9.0【B】: Python 3.10 未満を検出したら警告(続行は可能)。probe未完了/不明(pyMajor==0)なら判定しない(待たせない)。
    // v2.9.0: v2.9.11 で入れた上限(3.13以上を警告)は前提が誤りだったので撤回し、下限のみに戻す。
    // 3.13 が入らなかったのは Python 側の問題ではなく cu121 固定が原因(DetectCudaTag()で解消済み)。
    // 万一 torch の導入に失敗した場合は、完了レポートに WARN として正直に出る(v2.9.7)。
    bool pyTooOld = (g_probe.pyMajor < 3) || (g_probe.pyMajor == 3 && g_probe.pyMinor < 10);
    if(g_probe.done && g_probe.pyMajor != 0 && pyTooOld){
        std::string pyVerStr = std::to_string(g_probe.pyMajor) + "." + std::to_string(g_probe.pyMinor) + "." + std::to_string(g_probe.pyPatch);
        int pv = MsgBox(g_wnd,
            "\xe6\xa4\x9c\xe5\x87\xba\xe3\x81\x95\xe3\x82\x8c\xe3\x81\x9f" " Python " "\xe3\x81\xaf" " " + pyVerStr +
            " " "\xe3\x81\xa7\xe3\x81\x99\xe3\x80\x82" "\n\n"
            "PyTorch / whisper " "\xe3\x81\xaf" " Python 3.10 " "\xe4\xbb\xa5\xe4\xb8\x8a\xe3\x82\x92\xe5\xbf\x85\xe8\xa6\x81\xe3\x81\xa8\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x99\xe3\x80\x82" "\n"
            "\xe3\x81\x93\xe3\x81\xae\xe3\x81\xbe\xe3\x81\xbe\xe7\xb6\x9a\xe3\x81\x91\xe3\x82\x8b\xe3\x81\xa8\xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe3\x81\xab\xe5\xa4\xb1\xe6\x95\x97\xe3\x81\x99\xe3\x82\x8b\xe5\x8f\xaf\xe8\x83\xbd\xe6\x80\xa7\xe3\x81\x8c\xe9\xab\x98\xe3\x81\x84\xe3\x81\xa7\xe3\x81\x99\xe3\x80\x82" "\n\n"
            "\xe7\xb6\x9a\xe8\xa1\x8c\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x99\xe3\x81\x8b\xef\xbc\x9f",
            "Python" "\xe3\x83\x90\xe3\x83\xbc\xe3\x82\xb8\xe3\x83\xa7\xe3\x83\xb3\xe3\x81\xae\xe7\xa2\xba\xe8\xaa\x8d", MB_YESNO|MB_ICONWARNING);
        if(pv == IDNO){
            ClearSetupChecks();
            SetBusy(false);
            return;
        }
    }

    // v2.8.8: 実行時(埋め込みPython helper)と同じく spDir を sys.path 先頭に置いて判定する。
    // 旧実装は spDir を通さなかったため --target で入れた物を毎回「未導入」と誤判定し、
    // 数GBの再ダウンロードが走っていた。UpdateFugashiStatus と同じ書式に揃える。
    std::string spDir = GetSitePackagesDir();
    auto PkgOk = [&](const std::string& imp) -> bool {
        std::string o;
        std::string code = "import sys; sys.path.insert(0, r'" + spDir + "'); import " + imp;
        return RunProcess(Utf8ToWide("\"" + python + "\" -c \"" + code + "\""), o, 60000);
    };

    // Install whisper backend to local site-packages
    int setupBi = SendMessageA(g_backendCombo, CB_GETCURSEL, 0, 0);
    if(setupBi == CB_ERR) setupBi = 0;
    // For openai-whisper, ensure CUDA torch. v2.8.2: 既にCUDA使えるならスキップ。
    // v2.8.8: whisper側の依存解決で torch を引き直させない(2-3GBの二重DL回避)ため、whisper導入より先に実行する。
    // v2.9.0【D】: g_chkTorch がONならバックエンド設定に関わらず導入する(チェックを付けた以上入れる)。
    if(setupBi == 1 || fTorch){
        SetStatus("PyTorch \xe7\xa2\xba\xe8\xaa\x8d\xe4\xb8\xad...");
        SetProgress(15);
        std::string to;
        std::string tcode = "import sys; sys.path.insert(0, r'" + spDir + "'); import torch; exit(0 if torch.cuda.is_available() else 1)";
        bool torchCudaOk = RunProcess(Utf8ToWide("\"" + python + "\" -c \"" + tcode + "\""), to, 120000);
        if(!fTorch && torchCudaOk){ // v2.9.0【D】: チェック時はskip判定を無視
            report += "PyTorch(CUDA): \xe5\xb0\x8e\xe5\x85\xa5\xe6\xb8\x88\xe3\x81\xbf (skip)\n";
        } else {
            std::string cudaTag = DetectCudaTag(); // v2.8.8
            std::string idx = (cudaTag == "cpu")
                ? "https://download.pytorch.org/whl/cpu"
                : ("https://download.pytorch.org/whl/" + cudaTag);
            SetStatus(cudaTag == "cpu"
                ? "PyTorch (CPU\xe7\x89\x88) \xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe4\xb8\xad..."
                : "PyTorch (CUDA) \xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe4\xb8\xad... (\xe6\x99\x82\xe9\x96\x93\xe3\x81\x8c\xe3\x81\x8b\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x99)");
            std::wstring torchCmd = Utf8ToWide("\"" + python + "\" -m pip install torch --upgrade --target=\"" + spDir + "\" --index-url " + idx);
            std::string torchOut; bool tok = RunProcess(torchCmd, torchOut, 1800000); // 2-3GBなので30分
            DebugLog("Torch install (" + cudaTag + "): " + torchOut.substr(0, 500));
            // v2.9.7: 従来は成否に関わらず "install" と書いていたため、**全部失敗しているのに
            // 完了レポートが成功したように見えた**(実測: スタブPythonで全滅時も "PyTorch(cu121): install")。
            // 実際に import できるかまで確かめて報告する。
            report += (tok && PkgOk("torch"))
                ? ("PyTorch(" + cudaTag + "): OK\n")
                : ("PyTorch(" + cudaTag + "): WARN\n" + torchOut.substr(0, 200) + "\n");
        }
    }

    SetProgress(20);
    // v2.9.0【D】: 従来は setupBi で1つに決め打ちだったため、backend=openai-whisper のままでは
    // faster-whisper を導入する手段が無かった(kotoba-whisper が選べるのに動かない原因のひとつ)。
    // 「現在のバックエンドに必要な方」は従来どおり不足時に導入し、「もう片方」はチェックが付いているときだけ導入する。
    auto InstallWhisperPkg = [&](const char* pipPkg, const char* pipImp, bool forceIt, bool withDeps){
        SetStatusW(Utf8ToWide(std::string(pipPkg) + " \xe7\xa2\xba\xe8\xaa\x8d\xe4\xb8\xad...").c_str());
        if(!forceIt && PkgOk(pipImp)){
            report += std::string(pipPkg) + ": \xe5\xb0\x8e\xe5\x85\xa5\xe6\xb8\x88\xe3\x81\xbf (skip)\n";
            return;
        }
        SetStatusW(Utf8ToWide(std::string(pipPkg) + " \xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe4\xb8\xad...").c_str());
        std::string out; bool ok;
        if(withDeps){
            // v2.8.8: torch は前段で適正版を導入済みなので、依存解決で引き直させない (2-3GBの二重DL回避)。
            // openai-whisper のみ --no-deps + 軽量依存明示を使う。
            std::wstring cmd = Utf8ToWide("\"" + python + "\" -m pip install " + pipPkg + " --target=\"" + spDir + "\" --upgrade --quiet --no-deps");
            ok = RunProcess(cmd, out, 600000);
            std::wstring dcmd = Utf8ToWide("\"" + python + "\" -m pip install numba numpy tqdm more-itertools tiktoken --target=\"" + spDir + "\" --upgrade --quiet");
            std::string dout; RunProcess(dcmd, dout, 600000);
            DebugLog(std::string(pipPkg) + " deps:\n" + dout);
            // v2.8.8: --no-deps で不足が出た場合の保険。import できなければ従来どおり依存込みで入れ直す
            if(!PkgOk(pipImp)){
                DebugLog("--no-deps install incomplete, retrying with deps");
                std::wstring rcmd = Utf8ToWide("\"" + python + "\" -m pip install " + pipPkg + " --target=\"" + spDir + "\" --upgrade --quiet");
                std::string rout; ok = RunProcess(rcmd, rout, 900000);
            }
        } else {
            // v2.9.0【D】: faster-whisper には --no-deps を使わない(v2.8.7からの積み残し。
            // --no-depsだとPkgOk失敗→フル再インストールの二度手間になっていたため、直接フル導入に変更)。
            std::wstring cmd = Utf8ToWide("\"" + python + "\" -m pip install " + pipPkg + " --target=\"" + spDir + "\" --upgrade --quiet");
            ok = RunProcess(cmd, out, 900000);
            // v2.9.6【配布】faster-whisper(CTranslate2)はCUDA実行に cuBLAS/cuDNN の DLL を必要とするが、
            // pip の依存には含まれないためWindowsでは自動で入らない。無いと**モデル構築は成功するのに
            // 転写の途中で「cublas64_12.dll is not found」で落ちる**(構築時ではないのでCPUフォールバックも
            // 効かない=字幕がゼロになる)。torch同梱の torch\lib でも代用できるが、torchを入れない
            // faster-whisper単体構成では存在しないので、CUDA機のときだけ明示的に導入する。
            // CPU機(DetectCudaTag()=="cpu")では不要なので入れない(約1GBの無駄を避ける)。
            if(ok && std::string(pipPkg) == "faster-whisper" && DetectCudaTag() != "cpu"){
                SetStatus("CUDA \xe3\x83\xa9\xe3\x83\xb3\xe3\x82\xbf\xe3\x82\xa4\xe3\x83\xa0 \xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe4\xb8\xad...");
                std::wstring ncmd = Utf8ToWide("\"" + python + "\" -m pip install nvidia-cublas-cu12 nvidia-cudnn-cu12 --target=\"" + spDir + "\" --upgrade --quiet");
                std::string nout; bool nok = RunProcess(ncmd, nout, 1800000);
                DebugLog("nvidia cuBLAS/cuDNN install: " + nout.substr(0, 500));
                report += nok ? "CUDA runtime (cuBLAS/cuDNN): OK\n" : "CUDA runtime: WARN\n";
            }
        }
        DebugLog(std::string(pipPkg) + " install:\n" + out);
        report += ok ? (std::string(pipPkg) + ": OK (" + spDir + ")\n") : (std::string(pipPkg) + ": WARN\n" + out + "\n");
    };
    if(setupBi == 1){
        InstallWhisperPkg("openai-whisper", "whisper", fWhisper, true);
        if(fFaster) InstallWhisperPkg("faster-whisper", "faster_whisper", true, false);
    } else {
        InstallWhisperPkg("faster-whisper", "faster_whisper", fFaster, false);
        if(fWhisper) InstallWhisperPkg("openai-whisper", "whisper", true, true);
    }
    // v2.8.2: プロ分割(形態素解析)用の fugashi + 辞書を同梱site-packagesへインストール。
    // 失敗しても致命的ではない (プロ分割OFFなら不要、Python側 _morph_group が未導入時は無分割にフォールバック)。
    {
        SetStatus("fugashi (\xe5\xbd\xa2\xe6\x85\x8b\xe7\xb4\xa0\xe8\xa7\xa3\xe6\x9e\x90) \xe7\xa2\xba\xe8\xaa\x8d\xe4\xb8\xad...");
        SetProgress(35);
        if(!fFugashi && PkgOk("fugashi") && PkgOk("unidic_lite")){ // v2.9.0【D】: チェック時はskip判定を無視
            report += "fugashi+unidic-lite: \xe5\xb0\x8e\xe5\x85\xa5\xe6\xb8\x88\xe3\x81\xbf (skip)\n";
        } else {
            SetStatus("fugashi (\xe5\xbd\xa2\xe6\x85\x8b\xe7\xb4\xa0\xe8\xa7\xa3\xe6\x9e\x90) \xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe4\xb8\xad...");
            std::wstring fcmd = Utf8ToWide("\"" + python + "\" -m pip install fugashi unidic-lite --target=\"" + spDir + "\" --upgrade --quiet");
            std::string fout; bool fok = RunProcess(fcmd, fout, 600000);
            DebugLog("fugashi install:\n" + fout);
            report += fok ? "fugashi+unidic-lite: OK\n" : ("fugashi+unidic-lite: WARN\n" + fout + "\n");
        }
    }

    // Download model
    SetProgress(50);
    SetStatusW(L"\x30e2\x30c7\x30eb DL\x4e2d..."); // v2.8.10【I】: SetWindowTextW直接呼びを置き換え(環境タブ凍結対策)
    std::string mDir = GetModelsDir();

    // v2.9.0【G】: kotoba-whisper は faster-whisper 専用モデル。backend=openai-whisper では動かないので警告のみで続行。
    if(mName == "kotoba-whisper" && setupBi == 1){
        MsgBox(g_wnd,
            "kotoba-whisper " "\xe3\x81\xaf" " faster-whisper " "\xe5\xb0\x82\xe7\x94\xa8\xe3\x83\xa2\xe3\x83\x87\xe3\x83\xab\xe3\x81\xa7\xe3\x81\x99\xe3\x80\x82" "\n\n"
            "\xe7\x8f\xbe\xe5\x9c\xa8\xe3\x81\xae\xe3\x83\x90\xe3\x83\x83\xe3\x82\xaf\xe3\x82\xa8\xe3\x83\xb3\xe3\x83\x89\xe8\xa8\xad\xe5\xae\x9a" "(openai-whisper)" "\xe3\x81\xa7\xe3\x81\xaf\xe5\x8b\x95\xe4\xbd\x9c\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82" "\n"
            "\xe3\x83\x90\xe3\x83\x83\xe3\x82\xaf\xe3\x82\xa8\xe3\x83\xb3\xe3\x83\x89\xe3\x82\x92" " faster-whisper " "\xe3\x81\xab\xe5\xa4\x89\xe6\x9b\xb4\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82",
            "\xe8\xad\xa6\xe5\x91\x8a", MB_OK|MB_ICONWARNING);
    }

    // Check if model already exists locally (skip expensive download+load)
    bool modelExists = ModelExists(mName); // v2.9.0【G】: SetupThread内インラインだった判定をModelExists()へ切り出し
    if(modelExists){ // v2.9.1: モデルのチェック廃止に伴い、存在すれば常にskip(v2.8.10以前と同じ挙動)
        DebugLog("Model already exists: " + mDir + "\\" + mName);
        report += "\xe3\x83\xa2\xe3\x83\x87\xe3\x83\xab(" + mName + "): \xe6\x97\xa2\xe3\x81\xab\xe5\xad\x98\xe5\x9c\xa8 (skip)\n";
    } else {
    std::string dlScript = GetTempDir() + "\\dl_model.py";
    {
        std::ofstream sf(Utf8ToWide(dlScript));
        sf << "import sys, os\n";
        sf << "sp = sys.argv[3] if len(sys.argv) > 3 else ''\n";
        sf << "if sp and os.path.isdir(sp): sys.path.insert(0, sp)\n";
        sf << "backend = sys.argv[4] if len(sys.argv) > 4 else 'faster-whisper'\n";
        sf << "model_dir, model_name = sys.argv[1], sys.argv[2]\n";
        sf << "try:\n";
        sf << "    if backend == 'whisper':\n";
        sf << "        import whisper\n";
        sf << "        mn = 'turbo' if model_name == 'large-v3-turbo' else model_name\n";
        sf << "        print(f'Downloading {mn} (openai-whisper) to {model_dir}')\n";
        sf << "        whisper.load_model(mn, download_root=model_dir)\n";
        sf << "    else:\n";
        sf << "        from faster_whisper import WhisperModel\n";
        sf << "        distil_map = {'kotoba-whisper': 'kotoba-tech/kotoba-whisper-v2.0-faster'}\n";
        sf << "        dl_name = distil_map.get(model_name, model_name)\n";
        sf << "        print(f'Downloading {dl_name} (faster-whisper) to {model_dir}')\n";
        sf << "        WhisperModel(dl_name, device='cpu', compute_type='int8', download_root=model_dir)\n";
        sf << "    print('OK')\n";
        sf << "except Exception as e:\n";
        sf << "    print(f'Error: {e}')\n";
        sf << "    import traceback; traceback.print_exc()\n";
        sf << "    sys.exit(1)\n";
    }
    {
        std::string backendName = (setupBi == 1) ? "whisper" : "faster-whisper";
        std::wstring cmd = Utf8ToWide("\"" + python + "\" \"" + dlScript + "\" \"" + mDir + "\" " + mName + " \"" + spDir + "\" " + backendName);
        std::string out; bool ok = RunProcess(cmd, out, 1200000);
        DebugLog("Model DL:\n" + out);
        report += ok ? ("\xe3\x83\xa2\xe3\x83\x87\xe3\x83\xab(" + mName + "): OK\n") : ("\xe3\x83\xa2\xe3\x83\x87\xe3\x83\xab: " + out + "\n");
    }
    DeleteFileU(dlScript);
    } // end else (model not exists)

    SetProgress(80);
    std::string ffmpeg = GetEffectiveFFmpeg();
    // v2.9.4: チェックが付いていれば導入済みでも取り直す(壊れたffmpegの入れ直し手段)。
    // 従来は環境タブの専用ボタンが唯一の再取得手段で、しかも「未導入時のみ表示」だったため
    // ffmpegが存在するが壊れている場合に入れ直せなかった。他の項目と同じチェック方式に統一した。
    if(fFfmpeg || ffmpeg.empty()){ // v2.8.8: 無ければ自動取得（1クリック導入の完成）
        SetStatus("ffmpeg \xe3\x83\x80\xe3\x82\xa6\xe3\x83\xb3\xe3\x83\xad\xe3\x83\xbc\xe3\x83\x89\xe4\xb8\xad... (\xe7\xb4\x84" "104MB)");
        std::string got = DownloadFFmpegWithProgress(); // v2.8.8【C】: 実進捗ポーリング版
        if(!got.empty()){
            g_ffmpegPath = got;
            SetPathLabel(g_ffmpegLabel, g_ffmpegPath, "(aviutl2.exe\xe3\x81\xae\xe5\xa0\xb4\xe6\x89\x80)");
            ffmpeg = got;
        }
    }
    report += ffmpeg.empty()
        ? "ffmpeg: \xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93 (\xe3\x80\x8c" "ffmpeg\xe9\x81\xb8\xe6\x8a\x9e\xe3\x80\x8d\xe3\x81\xa7\xe6\x8c\x87\xe5\xae\x9a)\n"
        : ("ffmpeg: " + ffmpeg + "\n");
    SetProgress(100);
    SaveSettings();
    MsgBox(g_wnd,
        "\xe5\x88\x9d\xe6\x9c\x9f\xe8\xa8\xad\xe5\xae\x9a\xe5\xae\x8c\xe4\xba\x86\n\n" + report,
        "\xe3\x82\xbb\xe3\x83\x83\xe3\x83\x88\xe3\x82\xa2\xe3\x83\x83\xe3\x83\x97", MB_OK|MB_ICONINFORMATION);
    SetStatusW(L"Ready (v2.9.0)"); // v2.8.10【I】: SetupThread正常終了時にg_statusSetupが取り残され「fugashi確認中...」等で凍結していたバグの本丸修正+バージョン更新
    SetProgress(0);
    UpdateWhisperLocLabels();
    // v2.9.10【重要】セットアップ直後に probe を実測し直す。
    // g_probe は**起動時に1回しか実測しない**ため、セットアップでtorch等を入れても
    // g_probe.torch は false のまま取り残される。結果、CheckMissingForGenerate() が
    // 「PyTorchが必要です」と言い続ける一方、SetupThread側は実行時に生で判定するので
    // 「導入済み」としてスキップし、**両者が食い違って永久に抜けられなくなる**
    // (実測: セットアップ→生成で毎回PyTorch要求、セットアップ押してもskip、AviUtl2再起動で解消)。
    // ここはすでにバックグラウンドスレッドなので同期呼び出しでよい(数秒かかるがUIは固まらない)。
    g_probe.done = false;
    ProbeThread(); // 完了時に WM_PROBE_DONE を投げて RefreshSetupTabState() まで走る
    ClearSetupChecks(); // v2.9.0【D】: 正常終了パスでもOFFに戻す(誤って2回目が走るのを防ぐ)
    SetBusy(false);
}

// v2.9.0【A】: 起動時に python を1回だけ回して torch/whisper/faster-whisper/fugashi の導入状況を
// まとめて実測しキャッシュする(g_probe)。表示専用。SetupThreadのskip判定には使わない
// (実行時点の実態を見るべきなので、従来どおりその場で PkgOk する)。
static void ProbeThread(){
    std::string py = GetEffectivePython();
    if(py.empty()){ g_probe.done = true; PostMessageW(g_wnd, WM_PROBE_DONE, 0, 0); return; }
    g_probe.pyFound = true;
    std::string sp = GetSitePackagesDir();
    std::string scr = GetTempDir() + "\\probe_env.py";
    {
        std::ofstream sf(Utf8ToWide(scr));
        sf << "import sys, os\n";
        sf << "sp = sys.argv[1] if len(sys.argv) > 1 else ''\n";
        sf << "if sp and os.path.isdir(sp): sys.path.insert(0, sp)\n";
        sf << "print('PY=%d.%d.%d' % (sys.version_info[0], sys.version_info[1], sys.version_info[2]))\n";
        sf << "def chk(m):\n";
        sf << "    try:\n";
        sf << "        __import__(m)\n";
        sf << "        return '1'\n";
        sf << "    except Exception:\n";
        sf << "        return '0'\n";
        sf << "print('TORCH=' + chk('torch'))\n";
        sf << "cu = '0'\n";
        sf << "try:\n";
        sf << "    import torch\n";
        sf << "    if torch.cuda.is_available(): cu = '1'\n";
        sf << "except Exception:\n";
        sf << "    pass\n";
        sf << "print('TORCHCUDA=' + cu)\n";
        sf << "print('WHISPER=' + chk('whisper'))\n";
        sf << "print('FASTER=' + chk('faster_whisper'))\n";
        sf << "fg = '1' if (chk('fugashi') == '1' and chk('unidic_lite') == '1') else '0'\n";
        sf << "print('FUGASHI=' + fg)\n";
        sf << "print('DONE')\n";
    }
    std::string out;
    RunProcess(Utf8ToWide("\"" + py + "\" \"" + scr + "\" \"" + sp + "\""), out, 180000);
    DeleteFileU(scr);
    // v2.9.0: out を行単位でパースして g_probe に格納。非atomicメンバは done=true にする前に書く
    // (読む側は done==true を確認してから読む、という既存コードと同水準の作法)。
    std::istringstream iss(out);
    std::string line;
    while(std::getline(iss, line)){
        while(!line.empty() && (line.back()=='\r'||line.back()=='\n')) line.pop_back();
        size_t eq = line.find('=');
        std::string key = (eq != std::string::npos) ? line.substr(0, eq) : line;
        std::string val = (eq != std::string::npos) ? line.substr(eq+1) : "";
        if(key == "PY"){
            int a=0,b=0,c=0;
            sscanf_s(val.c_str(), "%d.%d.%d", &a, &b, &c);
            g_probe.pyMajor = a; g_probe.pyMinor = b; g_probe.pyPatch = c;
        }
        else if(key == "TORCH") g_probe.torch = (val == "1");
        else if(key == "TORCHCUDA") g_probe.torchCuda = (val == "1");
        else if(key == "WHISPER") g_probe.whisper = (val == "1");
        else if(key == "FASTER") g_probe.fasterWhisper = (val == "1");
        else if(key == "FUGASHI") g_probe.fugashi = (val == "1");
    }
    g_probe.done = true;
    PostMessageW(g_wnd, WM_PROBE_DONE, 0, 0);
}

// v2.9.0【F】: fugashi単独[再導入]ボタン用のスレッド関数は削除。
// SetupThread内のfugashiブロックとpipコマンドが完全に同一で機能重複していたため、
// チェックボックス方式(g_chkFugashi)に統合した。

// v2.9.4: DownloadFFmpegThread() は削除。ffmpeg専用ボタン(IDC_DL_FFMPEG)の唯一の呼び出し元だったため、
// ボタン廃止で未使用関数(C4505)になる。DL処理本体の DownloadFFmpegWithProgress() は SetupThread が使うので残す。

// =========================================================================
// BrowseForFile
// =========================================================================

static std::string BrowseForFile(HWND parent, LPCWSTR filter, LPCWSTR title, const std::string& initialPath = ""){
    wchar_t fn[MAX_PATH] = {};
    std::wstring initDir; // v2.8.5: ofn 使用中は生存させる必要があるので関数スコープに置く
    if(!initialPath.empty()){
        std::wstring wp = Utf8ToWide(initialPath);
        // v2.8.5: ファイル名欄に現在のフルパスを入れる = そのファイルが選択された状態で開く
        if(wp.size() < MAX_PATH) wcscpy_s(fn, MAX_PATH, wp.c_str());
        size_t sl = wp.find_last_of(L"\\/");
        if(sl != std::wstring::npos) initDir = wp.substr(0, sl);
    }
    OPENFILENAMEW ofn = {sizeof(ofn)};
    ofn.hwndOwner = parent; ofn.lpstrFilter = filter;
    ofn.lpstrFile = fn; ofn.nMaxFile = MAX_PATH; ofn.lpstrTitle = title;
    if(!initDir.empty()) ofn.lpstrInitialDir = initDir.c_str();
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
    if(GetOpenFileNameW(&ofn)) return WideToUtf8(fn);
    return "";
}

// v2.8.5: フォルダ選択ダイアログを現在の参照先で開くためのコールバック (v2.8.6: 旧方式フォールバック用に残置)
static int CALLBACK BrowseSelCallback(HWND hwnd, UINT msg, LPARAM lp, LPARAM data){
    (void)lp;
    if(msg == BFFM_INITIALIZED && data)
        SendMessageW(hwnd, BFFM_SETSELECTIONW, TRUE, data);
    return 0;
}

// v2.8.6: 旧方式 (SHBrowseForFolderW / ツリー表示)。IFileDialog が使えない環境用のフォールバック
static std::string BrowseForFolderLegacy(HWND parent, LPCWSTR title, const std::string& initialPath){
    BROWSEINFOW bi = {};
    std::wstring wInit; // v2.8.5: SHBrowseForFolderW 実行中は生存させる
    bi.hwndOwner = parent;
    bi.lpszTitle = title;
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_USENEWUI;
    if(!initialPath.empty()){
        wInit = Utf8ToWide(initialPath);
        bi.lpfn = BrowseSelCallback;
        bi.lParam = (LPARAM)wInit.c_str();
    }
    PIDLIST_ABSOLUTE pidl = SHBrowseForFolderW(&bi);
    if(pidl){
        wchar_t path[MAX_PATH] = {};
        SHGetPathFromIDListW(pidl, path);
        CoTaskMemFree(pidl);
        return WideToUtf8(path);
    }
    return "";
}

// v2.8.6: モダンなフォルダ選択 (IFileDialog + FOS_PICKFOLDERS)。
// ffmpeg/python の GetOpenFileNameW と同じ explorer 風の見た目になる。
// handled=true は「モダンダイアログが正常に表示された」(ユーザーのキャンセル含む) の意味。
// この場合フォールバックを出してはいけない (キャンセルしたら旧ダイアログが開く、を防ぐ)。
static std::string BrowseForFolderModern(HWND parent, LPCWSTR title, const std::string& initialPath, bool& handled){
    handled = false;
    std::string result;
    HRESULT hrCo = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if(hrCo == RPC_E_CHANGED_MODE) return result; // 既にMTA = IFileDialog不可 → 旧方式へ
    bool needUninit = SUCCEEDED(hrCo);            // S_OK も S_FALSE も参照カウント+1なので対で戻す
    IFileDialog* pfd = NULL;
    if(SUCCEEDED(CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pfd))) && pfd){
        DWORD opts = 0;
        if(SUCCEEDED(pfd->GetOptions(&opts)))
            pfd->SetOptions(opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST);
        if(title) pfd->SetTitle(title);
        if(!initialPath.empty()){
            std::wstring wp = Utf8ToWide(initialPath);
            IShellItem* psi = NULL;
            if(SUCCEEDED(SHCreateItemFromParsingName(wp.c_str(), NULL, IID_PPV_ARGS(&psi))) && psi){
                pfd->SetFolder(psi); // v2.8.6: 現在の参照先を初期位置に (v2.8.5の意図を維持)
                psi->Release();
            }
        }
        HRESULT hr = pfd->Show(parent);
        if(SUCCEEDED(hr)){
            IShellItem* psiRes = NULL;
            if(SUCCEEDED(pfd->GetResult(&psiRes)) && psiRes){
                PWSTR pszPath = NULL;
                if(SUCCEEDED(psiRes->GetDisplayName(SIGDN_FILESYSPATH, &pszPath)) && pszPath){
                    result = WideToUtf8(pszPath);
                    CoTaskMemFree(pszPath);
                }
                psiRes->Release();
            }
            handled = true;
        } else if(hr == HRESULT_FROM_WIN32(ERROR_CANCELLED)){
            handled = true; // ユーザーキャンセル = 正常終了
        }
        pfd->Release();
    }
    if(needUninit) CoUninitialize();
    return result;
}

static std::string BrowseForFolder(HWND parent, LPCWSTR title, const std::string& initialPath = ""){
    bool handled = false;
    std::string r = BrowseForFolderModern(parent, title, initialPath, handled);
    if(handled) return r;
    return BrowseForFolderLegacy(parent, title, initialPath); // v2.8.6: モダン版が使えない環境のみ
}

// Resolve the site-packages dir for a whisper backend
// Custom path > local (plugin/site-packages) > system (python/Lib/site-packages)
static std::string GetEffectiveFwSpDir(){
    if(!g_fwSpPath.empty() && FileExistsU(g_fwSpPath + "\\faster_whisper\\__init__.py"))
        return g_fwSpPath;
    std::string local = GetSitePackagesDir();
    if(FileExistsU(local + "\\faster_whisper\\__init__.py"))
        return local;
    std::string python = GetEffectivePython();
    if(!python.empty()){
        size_t sl = python.find_last_of("\\/");
        if(sl != std::string::npos){
            std::string sys = python.substr(0, sl) + "\\Lib\\site-packages";
            if(FileExistsU(sys + "\\faster_whisper\\__init__.py"))
                return sys;
        }
    }
    return "";
}
static std::string GetEffectiveOwSpDir(){
    if(!g_owSpPath.empty() && FileExistsU(g_owSpPath + "\\whisper\\__init__.py"))
        return g_owSpPath;
    std::string local = GetSitePackagesDir();
    if(FileExistsU(local + "\\whisper\\__init__.py"))
        return local;
    std::string python = GetEffectivePython();
    if(!python.empty()){
        size_t sl = python.find_last_of("\\/");
        if(sl != std::string::npos){
            std::string sys = python.substr(0, sl) + "\\Lib\\site-packages";
            if(FileExistsU(sys + "\\whisper\\__init__.py"))
                return sys;
        }
    }
    return "";
}

// v2.9.0【I】: 字幕生成ボタン押下時に不足コンポーネントを検出する。判定は全てファイル存在 or
// probeキャッシュ参照のみで、pythonは起動しない(生成が遅くならないこと)。不足が無ければ空文字を返す。
static std::string CheckMissingForGenerate(){
    std::string lack;
    // v2.9.7: 実在するだけでは足りない。probeがバージョンを取れなかったPython(=実際には動かない。
    // Microsoft Storeのスタブ等)も「不足」として扱い、生成ボタンの導線で拾えるようにする。
    if(GetEffectivePython().empty() || (g_probe.done && g_probe.pyFound && g_probe.pyMajor == 0))
        lack += "\n" "\xe3\x83\xbb" "Python (3.10" "\xe4\xbb\xa5\xe4\xb8\x8a" ")";
    if(GetEffectiveFFmpeg().empty())
        lack += "\n" "\xe3\x83\xbb" "ffmpeg";
    int bi = SendMessageA(g_backendCombo, CB_GETCURSEL, 0, 0);
    if(bi == CB_ERR) bi = 0;
    if(bi == 1){ // openai-whisper
        if(g_probe.done && !g_probe.torch)
            lack += "\n" "\xe3\x83\xbb" "PyTorch (" "\xe7\xb4\x84" "4.3GB)";
        if(GetEffectiveOwSpDir().empty())
            lack += "\n" "\xe3\x83\xbb" "whisper";
    } else { // faster-whisper
        if(GetEffectiveFwSpDir().empty())
            lack += "\n" "\xe3\x83\xbb" "faster-whisper";
    }
    int mi = SendMessageA(g_modelCombo, CB_GETCURSEL, 0, 0);
    if(mi == CB_ERR) mi = 0;
    if(mi < 0 || mi >= (int)(sizeof(kModelNames)/sizeof(kModelNames[0]))) mi = 0;
    std::string mName = kModelNames[mi];
    if(!ModelExists(mName))
        lack += "\n" "\xe3\x83\xbb\xe3\x83\xa2\xe3\x83\x87\xe3\x83\xab" " " + mName;
    if(SendMessageA(g_chkMorphSplit, BM_GETCHECK, 0, 0) == BST_CHECKED){ // v2.8.5: 文節区切りONのときだけ判定
        if(g_probe.done && !g_probe.fugashi)
            lack += "\n" "\xe3\x83\xbb\xe6\x96\x87\xe7\xaf\x80\xe5\x8c\xba\xe5\x88\x87\xe3\x82\x8a" "(fugashi)";
    }
    if(!lack.empty()) lack = lack.substr(1); // 先頭の改行を除去
    return lack;
}

// =========================================================================
// Template handling
// =========================================================================

static bool LoadTemplate(const std::string& path){
    std::ifstream f(Utf8ToWide(path), std::ios::binary);
    if(!f.is_open()) return false;
    std::string content((std::istreambuf_iterator<char>(f)), std::istreambuf_iterator<char>());
    f.close();
    // BOM removal
    if(content.size()>=3 && (unsigned char)content[0]==0xEF &&
       (unsigned char)content[1]==0xBB && (unsigned char)content[2]==0xBF){
        content = content.substr(3);
    }
    // Normalize newlines
    std::string normalized;
    for(size_t i=0; i<content.size(); i++){
        if(content[i]=='\r'){
            normalized += '\n';
            if(i+1<content.size() && content[i+1]=='\n') i++;
        } else {
            normalized += content[i];
        }
    }
    // Strip frame info that causes SDK to override our length parameter
    std::string cleaned;
    std::istringstream ss(normalized);
    std::string line;
    while(std::getline(ss, line)){
        // Skip frame= line (e.g. "frame=0,159") - SDK uses this to override length
        if(line.find("frame=")==0) continue;
        // Skip [exedit] section entirely
        if(line == "[exedit]"){
            while(std::getline(ss, line)){
                if(!line.empty() && line[0] == '['){ cleaned += line + "\n"; break; }
            }
            continue;
        }
        cleaned += line + "\n";
    }
    // Verify text key exists
    std::string textKey = "\xe3\x83\x86\xe3\x82\xad\xe3\x82\xb9\xe3\x83\x88=";
    if(cleaned.find(textKey) == std::string::npos){
        DebugLog("Template has no text key, using raw: " + path);
        cleaned = normalized; // fallback to unstripped
    }
    g_templateContent = cleaned;
    g_templatePath = path;
    DebugLog("Template loaded: " + path + " (" + std::to_string(cleaned.size()) + " bytes)\nCONTENT:\n" + cleaned + "\nEND_CONTENT");
    return true;
}

static void UpdateTemplateLabel(){
    std::string nm = g_templatePath;
    size_t p = nm.find_last_of("\\/");
    if(p != std::string::npos) nm = nm.substr(p+1);
    SetWindowTextW(g_templateLabel, Utf8ToWide(nm).c_str());
}

// =========================================================================
// Status/Progress helpers (thread-safe)
// =========================================================================

static void SetStatus(const std::string& msg){
    char* buf = _strdup(msg.c_str());
    PostMessageW(g_wnd, WM_UPDATE_STATUS, 0, (LPARAM)buf);
}
static void SetProgress(int val){
    PostMessageW(g_wnd, WM_UPDATE_PROGRESS, val, 0);
}
// v2.8.10【I】: g_status と g_statusSetup(環境タブ) を同時に更新する。
// SetWindowTextW(g_status,...) を直接呼ぶと環境タブ側が取り残されて表示が凍結するため、
// SetupThread 等からの wide 文字列での更新は必ずこれを使うこと。
static void SetStatusW(const wchar_t* msg){
    if(g_status) SetWindowTextW(g_status, msg);
    if(g_statusSetup) SetWindowTextW(g_statusSetup, msg);
}

// =========================================================================
// Timeline clip structure
// =========================================================================

struct TimelineClip {
    std::string filePath;
    int timelineStart, timelineEnd;
    double sourceOffset;
};
static std::vector<TimelineClip> g_tlClips;

// =========================================================================
// Scan timeline (call_edit_section callback)
// =========================================================================

struct ScanParam {
    std::vector<TimelineClip>* clips;
    int rate;
    int maxLayer;
};

static void ScanCallback(void* param, EDIT_SECTION* es){
    ScanParam* sp = (ScanParam*)param;
    if(!es || !es->info) return;
    sp->rate = es->info->rate;
    sp->maxLayer = es->info->layer_max;
    int maxF = es->info->frame_max;
    int maxL = es->info->layer_max;
    // Scan all layers for media objects
    for(int lay = 0; lay <= maxL; lay++){
        for(int f = 0; f <= maxF; ){
            OBJECT_HANDLE obj = es->find_object(lay, f);
            if(!obj){ f++; continue; }
            OBJECT_LAYER_FRAME olf = es->get_object_layer_frame(obj);
            if(olf.layer != lay){ f++; continue; } // found object on different layer
            // Try video file effect, then audio file effect
            LPCSTR val = es->get_object_item_value(obj, L"\x52d5" L"\x753b\x30d5\x30a1\x30a4\x30eb", L"\x30d5\x30a1\x30a4\x30eb");
            if(!val) val = es->get_object_item_value(obj, L"\x97f3\x58f0\x30d5\x30a1\x30a4\x30eb", L"\x30d5\x30a1\x30a4\x30eb");
            if(val && val[0]){
                TimelineClip c;
                c.filePath = val;
                c.timelineStart = olf.start;
                c.timelineEnd = olf.end;
                LPCSTR offVal = es->get_object_item_value(obj, L"\x52d5" L"\x753b\x30d5\x30a1\x30a4\x30eb", L"\x518d\x751f\x4f4d\x7f6e");
                if(!offVal) offVal = es->get_object_item_value(obj, L"\x97f3\x58f0\x30d5\x30a1\x30a4\x30eb", L"\x518d\x751f\x4f4d\x7f6e");
                c.sourceOffset = offVal ? atof(offVal) : 0.0;
                sp->clips->push_back(c);
            }
            f = olf.end + 1;
        }
    }
}

// =========================================================================
// Extract audio via ffmpeg
// =========================================================================

static std::string ExtractAudio(const TimelineClip& clip, int fps, const std::string& out){
    std::string ffmpeg = GetEffectiveFFmpeg();
    if(ffmpeg.empty()) return "";
    double dur = (double)(clip.timelineEnd - clip.timelineStart) / fps;
    // v2.8b: build the command line with std::wstring instead of a fixed wchar_t[2048]
    // buffer. Long Japanese paths (multi-byte per char once escaped/concatenated) could
    // silently overflow/truncate the old fixed buffer via swprintf_s (or fail outright).
    std::wstring wFfmpeg = Utf8ToWide(ffmpeg);
    std::wstring wFile = Utf8ToWide(clip.filePath);
    std::wstring wOut = Utf8ToWide(out);
    wchar_t durBuf[64], ssBuf[64];
    swprintf_s(durBuf, 64, L"%.3f", dur);
    swprintf_s(ssBuf, 64, L"%.3f", clip.sourceOffset);
    std::wstring cmdStr;
    cmdStr.reserve(wFfmpeg.size() + wFile.size() + wOut.size() + 64);
    cmdStr += L"\""; cmdStr += wFfmpeg; cmdStr += L"\" -y -ss "; cmdStr += ssBuf;
    cmdStr += L" -i \""; cmdStr += wFile; cmdStr += L"\" -t "; cmdStr += durBuf;
    cmdStr += L" -vn -acodec pcm_s16le -ar 16000 -ac 1 \""; cmdStr += wOut; cmdStr += L"\"";
    std::string procOut;
    bool ok = RunProcess(cmdStr, procOut, 120000);
    if(!ok){
        DebugLog("ffmpeg fail: " + procOut);
        return "";
    }
    if(!FileExistsU(out)) return "";
    return out;
}

// =========================================================================
// Segment structure
// =========================================================================

struct Seg { int s, e; std::string text; };
static std::vector<Seg> g_segs;

// =========================================================================
// SplitText
// =========================================================================

// v2.8.2: UTF-8 文字数をカウント (以前は byte 数だったため日本語1字=3byteでズレていた)
static int Utf8CharCount(const std::string& s){
    int n = 0;
    for(size_t i = 0; i < s.size(); ){
        unsigned char c = (unsigned char)s[i];
        i += (c < 0x80 ? 1 : c < 0xe0 ? 2 : c < 0xf0 ? 3 : 4);
        n++;
    }
    return n;
}
// v2.8.5: 行の前後の半角/全角スペースを除去
static std::string TrimSpaces(const std::string& s){
    size_t b = 0, e = s.size();
    auto lead = [&](size_t i)->int{
        if(i < s.size() && s[i]==' ') return 1;
        if(i+3 <= s.size() && (unsigned char)s[i]==0xe3 && (unsigned char)s[i+1]==0x80 && (unsigned char)s[i+2]==0x80) return 3;
        return 0;
    };
    int t;
    while(b < e && (t=lead(b))>0) b += t;
    while(e > b){
        if(s[e-1]==' '){ e--; continue; }
        if(e-b>=3 && (unsigned char)s[e-3]==0xe3 && (unsigned char)s[e-2]==0x80 && (unsigned char)s[e-1]==0x80){ e-=3; continue; }
        break;
    }
    return s.substr(b, e-b);
}
static std::vector<std::string> SplitText(const std::string& text, int maxChars){
    std::vector<std::string> res;
    // v2.8.5: 既に改行(リテラル \n)を含む=プロ分割側で確定済みのテキストは再分割しない
    if(text.find("\\n") != std::string::npos){ res.push_back(text); return res; }
    // v2.8.2: maxChars を「文字数」として扱う (以前はバイト数=日本語で maxchars=40→13字と誤解を招いた)
    if(maxChars > 0 && maxChars < 1) maxChars = 1;
    if(maxChars <= 0 || Utf8CharCount(text) <= maxChars){
        res.push_back(text);
        return res;
    }
    // 各文字の開始バイト位置 (末尾に番兵 = text.size())
    std::vector<size_t> cp;
    for(size_t i = 0; i < text.size(); ){
        cp.push_back(i);
        unsigned char c = (unsigned char)text[i];
        i += (c < 0x80 ? 1 : c < 0xe0 ? 2 : c < 0xf0 ? 3 : 4);
    }
    cp.push_back(text.size());
    int total = (int)cp.size() - 1; // 文字数
    auto chAt = [&](int ci){ return text.substr(cp[ci], cp[ci+1] - cp[ci]); };
    // その文字の「後ろ」で切ると自然な度合い (句読点=3, 助詞/終助詞=2, それ以外=0)
    auto breakPri = [&](int ci)->int{
        std::string c = chAt(ci);
        if(c==" "||c=="\xe3\x80\x80") return 3; // v2.8.5: 空白(ポーズ)は自然な区切り
        if(c=="\xe3\x80\x81"||c=="\xe3\x80\x82"||c=="\xef\xbc\x81"||c=="\xef\xbc\x9f"||c=="!"||c=="?") return 3; // 、。！？!?
        static const char* joshi[] = {"\xe3\x82\x92","\xe3\x81\xab","\xe3\x81\xaf","\xe3\x81\x8c","\xe3\x81\xa7","\xe3\x81\xa8","\xe3\x82\x82","\xe3\x81\xae","\xe3\x81\xb8","\xe3\x82\x84","\xe3\x81\xad","\xe3\x82\x88","\xe3\x81\x8b","\xe3\x81\xa0","\xe3\x81\x9f"}; // を に は が で と も の へ や ね よ か だ た
        for(auto j : joshi) if(c==j) return 2;
        return 0;
    };
    // 行頭に来ると不自然 = この文字の「前」では切らない (禁則)
    auto noHead = [&](int ci)->bool{
        std::string c = chAt(ci);
        return c=="\xe3\x80\x81"||c=="\xe3\x80\x82"||c=="\xe3\x81\xa3"||c=="\xe3\x82\x83"||c=="\xe3\x82\x85"||c=="\xe3\x82\x87"||c=="\xe3\x83\xbc"||c=="\xef\xbc\x81"||c=="\xef\xbc\x9f"; // 、。っゃゅょー！？
    };
    // v2.8.5: 複合助詞の途中で切らない (だから/から/より/〜たん)。1文字照合の誤爆対策
    auto inseparable = [&](int ci)->bool{
        if(ci+1 >= total) return false;
        std::string a = chAt(ci), b = chAt(ci+1);
        if(a=="\xe3\x81\xa0" && b=="\xe3\x81\x8b") return true; // だ+か (だから)
        if(a=="\xe3\x81\x8b" && b=="\xe3\x82\x89") return true; // か+ら (から)
        if(a=="\xe3\x82\x88" && b=="\xe3\x82\x8a") return true; // よ+り (より)
        if(a=="\xe3\x81\x9f" && b=="\xe3\x82\x93") return true; // た+ん (〜たん)
        return false;
    };
    int start = 0;
    while(start < total){
        int hardEnd = start + maxChars; // ここを超えては切れない (文字index)
        if(hardEnd >= total){ res.push_back(text.substr(cp[start])); break; }
        int minCut = start + (maxChars/2 > 1 ? maxChars/2 : 1); // 短すぎ回避
        int bestCut = -1, bestPri = 0;
        for(int ci = start; ci < hardEnd; ci++){
            int pri = breakPri(ci);
            if(pri > 0 && !inseparable(ci) && (ci+1 >= total || !noHead(ci+1)) && (ci+1) >= minCut){
                if(pri >= bestPri){ bestPri = pri; bestCut = ci; } // 後ろの良い区切りを優先
            }
        }
        int cutAfter;
        if(bestCut >= 0) cutAfter = bestCut;
        else {
            cutAfter = hardEnd - 1; // 良い区切り無し=上限で切る。ただし禁則開始文字は避ける
            while(cutAfter > start && (cutAfter+1 < total) && noHead(cutAfter+1)) cutAfter--;
        }
        int nextStart = cutAfter + 1;
        if(nextStart <= start) nextStart = start + 1; // 進行保証
        res.push_back(text.substr(cp[start], cp[nextStart] - cp[start]));
        start = nextStart;
    }
    for(auto& s : res) s = TrimSpaces(s);
    res.erase(std::remove_if(res.begin(), res.end(), [](const std::string& s){ return s.empty(); }), res.end());
    if(res.empty()) res.push_back(TrimSpaces(text));
    return res;
}

// =========================================================================
// Run faster-whisper transcription
// =========================================================================

static bool RunFasterWhisper(int mi, int di, int bi, int beamSize, int li, float temp, bool noPrevText, bool wordTs, bool repPenalty, const std::string& promptText, const std::string& hotwordsText, bool batched, bool vadFilter){ // v2.9.5: vadFilter追加
    const char* mn[] = {"tiny","base","small","medium","large-v3","large-v3-turbo","kotoba-whisper"};
    const char* dn[] = {"auto","cpu","cuda"};
    const char* bn[] = {"faster-whisper","whisper"};
    const char* lc[] = {"auto","ja","en","zh","ko"};
    std::string tmp = GetTempDir() + "\\";
    std::string bp = tmp + "whisper_batch.json";
    std::string op = tmp + "whisper_results.txt";
    std::string errP = op + ".err";

    // Cleanup helper for all exit paths
    std::vector<std::string> wavs;
    auto cleanup = [&](){
        DeleteFileU(bp); DeleteFileU(op); DeleteFileU(errP);
        for(auto& w : wavs) if(!w.empty()) DeleteFileU(w);
    };

    // Extract audio for all clips
    for(size_t ci = 0; ci < g_tlClips.size(); ci++){
        std::string w = tmp + "whisper_fw_" + std::to_string(ci) + ".wav";
        wavs.push_back(ExtractAudio(g_tlClips[ci], g_projectRate, w));
    }
    bool anyW = false;
    for(auto& w : wavs) if(!w.empty()) anyW = true;
    if(!anyW){
        MsgBox(g_wnd,
            "\xe9\x9f\xb3\xe5\xa3\xb0\xe6\x8a\xbd\xe5\x87\xba\xe5\xa4\xb1\xe6\x95\x97\xe3\x80\x82" "ffmpeg\xe3\x82\x92\xe7\xa2\xba\xe8\xaa\x8d\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82",
            "Error", MB_OK|MB_ICONERROR);
        cleanup();
        return false;
    }

    std::string md = GetModelsDir();

    // Write batch JSON
    {
        std::ofstream bf(Utf8ToWide(bp));
        char tempStr[32]; sprintf_s(tempStr, "%.2f", temp);
        bf << "{\n  \"model\": \"" << mn[mi] << "\",\n  \"language\": \"" << lc[li] << "\",\n  \"device\": \"" << dn[di] << "\",\n  \"backend\": \"" << bn[bi] << "\",\n  \"beam_size\": " << beamSize << ",\n  \"temperature\": " << tempStr << ",\n  \"no_prev_text\": " << (noPrevText ? "true" : "false") << ",\n  \"word_timestamps\": " << (wordTs ? "true" : "false") << ",\n  \"rep_penalty\": " << (repPenalty ? "true" : "false") << ",\n  \"vad_filter\": " << (vadFilter ? "true" : "false") << ",\n  \"prompt\": \"";
        // JSON-escape prompt text
        for(char c : promptText){
            if(c == '"') bf << "\\\"";
            else if(c == '\\') bf << "\\\\";
            else if(c == '\n') bf << "\\n";
            else if(c == '\r'){}
            else bf << c;
        }
        bf << "\",\n  \"hotwords\": \"";
        // JSON-escape hotwords text
        for(char c : hotwordsText){
            if(c == '"') bf << "\\\"";
            else if(c == '\\') bf << "\\\\";
            else bf << c;
        }
        bf << "\",\n  \"batched\": " << (batched ? "true" : "false");
        {   // v2.8.2: 形態素分割ON/OFF と 文字数上限 を Python へ渡す
            int morphOn = (SendMessageA(g_chkMorphSplit, BM_GETCHECK, 0, 0) == BST_CHECKED) ? 1 : 0;
            char mcb[16] = {}; GetWindowTextA(g_maxCharEdit, mcb, sizeof(mcb)); int mcv = atoi(mcb); if(mcv < 0) mcv = 0;
            int twoLn = (SendMessageA(g_chkTwoLine, BM_GETCHECK, 0, 0) == BST_CHECKED) ? 2 : 1; // v2.8.5
            bf << ",\n  \"morph_split\": " << (morphOn ? "true" : "false") << ",\n  \"maxchars\": " << mcv << ",\n  \"maxlines\": " << twoLn;
        }
        bf << ",\n  \"model_dir\": \"";
        // v2.8b: escape " as well as \ (paths are unlikely to contain quotes, but a stray
        // one would previously have produced invalid/truncated JSON silently)
        for(char c : md){ if(c == '\\') bf << "\\\\"; else if(c == '"') bf << "\\\""; else bf << c; }
        bf << "\",\n  \"extra_sp\": [";
        {
            // Collect extra site-packages paths for Python
            std::vector<std::string> extraSp;
            std::string eFw = GetEffectiveFwSpDir();
            std::string eOw = GetEffectiveOwSpDir();
            if(!eFw.empty()) extraSp.push_back(eFw);
            if(!eOw.empty() && eOw != eFw) extraSp.push_back(eOw);
            for(size_t i = 0; i < extraSp.size(); i++){
                if(i > 0) bf << ", ";
                bf << "\"";
                for(char c : extraSp[i]){ if(c == '\\') bf << "\\\\"; else if(c == '"') bf << "\\\""; else bf << c; }
                bf << "\"";
            }
        }
        bf << "]"; // v2.9.8: 直前の配列を閉じる(この後に ffmpeg フィールドを挟むため分離した)
        // v2.9.8【重要】解決済み ffmpeg のパスを渡す。
        // openai-whisper は audio.py の load_audio() が **PATH から "ffmpeg" を外部プロセス起動**する。
        // 同梱ffmpegは data\Plugin\whisper_subtitle\ffmpeg\<ver>\bin\ にあってPATHに無いため、
        // 新規環境では import も torch もモデル読込も成功するのに **load_audio()だけが
        // FileNotFoundError [WinError 2] で落ち、字幕が0件になる**(実測)。
        // faster-whisper は PyAV で自前デコードするため影響を受けず、これが
        // 「fasterでは生成できるのにwhisperだと空になる」の正体だった。
        {
            std::string ffm = GetEffectiveFFmpeg();
            std::string fesc;
            for(char c : ffm){ if(c == '\\') fesc += "\\\\"; else if(c == '"') fesc += "\\\""; else fesc += c; }
            bf << ",\n  \"ffmpeg\": \"" << fesc << "\"";
        }
        bf << ",\n  \"clips\": [\n";
        bool first = true;
        for(size_t ci = 0; ci < g_tlClips.size(); ci++){
            if(wavs[ci].empty()) continue;
            if(!first) bf << ",\n"; first = false;
            std::string esc;
            for(char c : wavs[ci]){ if(c == '\\') esc += "\\\\"; else if(c == '"') esc += "\\\""; else esc += c; }
            bf << "    {\"wav\": \"" << esc << "\", \"timeline_start\": " << g_tlClips[ci].timelineStart
               << ", \"timeline_end\": " << g_tlClips[ci].timelineEnd << ", \"fps\": " << g_projectRate << "}";
        }
        bf << "\n  ]\n}\n";
    }

    SetStatus(std::string("[") + bn[bi] + "] \xe6\x96\x87\xe5\xad\x97\xe8\xb5\xb7\xe3\x81\x93\xe3\x81\x97\xe4\xb8\xad...");
    SetProgress(40);
    EnsurePyHelper();
    std::string python = GetEffectivePython();
    if(python.empty()){
        MsgBox(g_wnd, "Python\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93", "Error", MB_OK|MB_ICONERROR);
        cleanup(); return false;
    }
    std::string ps = GetPluginDir() + "\\whisper_helper.py";
    DebugLog("Python: " + python + "\nScript: " + ps + "\nBatch: " + bp + "\nOutput: " + op);

    // Run Python with pipes
    std::wstring wCmd = Utf8ToWide("\"" + python + "\" \"" + ps + "\" \"" + bp + "\" \"" + op + "\"");
    std::string pyOut;
    bool pyOk = RunProcess(wCmd, pyOut, 600000);
    DebugLog("Python exit=" + std::string(pyOk ? "0" : "nonzero") + "\n" + pyOut);

    // Read .err file
    std::string errC;
    {
        std::ifstream ef(Utf8ToWide(errP));
        if(ef.is_open()){
            std::string l;
            while(std::getline(ef, l)) errC += l + "\n";
        }
    }
    if(!errC.empty()) DebugLog(".err:\n" + errC);

    // Read result file
    std::ifstream rf(Utf8ToWide(op));
    if(!rf.is_open()){
        std::string msg = "\xe7\xb5\x90\xe6\x9e\x9c\xe3\x83\x95\xe3\x82\xa1\xe3\x82\xa4\xe3\x83\xab\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n\n";
        if(!pyOut.empty()){
            msg += "--- Python output ---\n";
            if(pyOut.size() > 500) msg += pyOut.substr(pyOut.size()-500); else msg += pyOut;
            msg += "\n";
        }
        if(!errC.empty()){
            msg += "--- Log ---\n";
            if(errC.size() > 300) msg += errC.substr(errC.size()-300); else msg += errC;
        }
        msg += "\n\n\xe3\x83\x87\xe3\x83\x90\xe3\x83\x83\xe3\x82\xb0\xe3\x83\xad\xe3\x82\xb0: " + GetPluginDir() + "\\whisper_debug.log";
        MsgBox(g_wnd, msg, "Error", MB_OK|MB_ICONERROR);
        cleanup(); return false;
    }
    std::string line;
    bool needAutoInstall = false;
    std::string installPkg;
    while(std::getline(rf, line)){
        if(line.empty()) continue;
        if(line.substr(0, 5) == "ERROR"){
            // Check for "not installed" errors -> auto-install
            if(line.find("not installed") != std::string::npos){
                if(line.find("faster-whisper") != std::string::npos) installPkg = "faster-whisper";
                else if(line.find("whisper") != std::string::npos) installPkg = "openai-whisper";
                if(!installPkg.empty()) needAutoInstall = true;
            }
            if(!needAutoInstall){
                MsgBox(g_wnd, line.substr(6), "Error", MB_OK|MB_ICONERROR);
                rf.close(); cleanup(); return false;
            }
            break;
        }
        size_t p1 = line.find('|'); if(p1 == std::string::npos) continue;
        size_t p2 = line.find('|', p1+1); if(p2 == std::string::npos) continue;
        size_t p3 = line.find('|', p2+1); if(p3 == std::string::npos) continue;
        Seg seg;
        int ci = atoi(line.substr(0, p1).c_str());
        seg.s = atoi(line.substr(p1+1, p2-p1-1).c_str());
        seg.e = atoi(line.substr(p2+1, p3-p2-1).c_str());
        seg.text = line.substr(p3+1);
        if(ci >= 0 && ci < (int)g_tlClips.size()){
            if(seg.s < g_tlClips[ci].timelineStart) seg.s = g_tlClips[ci].timelineStart;
            if(seg.e > g_tlClips[ci].timelineEnd) seg.e = g_tlClips[ci].timelineEnd;
        }
        if(seg.e > seg.s && !seg.text.empty()) g_segs.push_back(seg);
    }
    rf.close();

    // Auto-install missing backend and retry
    if(needAutoInstall && !installPkg.empty()){
        std::string spDir = GetSitePackagesDir();
        std::string python = GetEffectivePython();
        DebugLog("Auto-installing: " + installPkg);
        SetStatus(installPkg + " \xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe4\xb8\xad...");
        SetProgress(30);

        // Install backend first
        SetStatus(installPkg + " \xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe4\xb8\xad...");
        SetProgress(40);
        std::wstring pipCmd = Utf8ToWide("\"" + python + "\" -m pip install " + installPkg + " --target=\"" + spDir + "\" --upgrade --quiet");
        std::string pipOut; bool pipOk = RunProcess(pipCmd, pipOut, 600000);
        DebugLog("Auto pip install: " + pipOut);

        // For openai-whisper, re-install CUDA torch AFTER (pip pulls CPU version)
        if(pipOk && installPkg == "openai-whisper"){
            SetStatus("PyTorch (CUDA) \xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe4\xb8\xad...");
            SetProgress(50);
            // v2.9.6【配布】ここだけ cu121 ハードコードのまま取り残されていた(v2.8.7でSetupThread側には
            // DetectCudaTag()を入れたが、この復旧経路は未対応だった)。GPUの無いPCで踏むと
            // 200MBのCPU版で済むところに2.5GBのCUDA版を落としてしまう。SetupThreadと同じ判定に揃える。
            std::string autoTag = DetectCudaTag();
            std::string autoIdx = (autoTag == "cpu")
                ? "https://download.pytorch.org/whl/cpu"
                : ("https://download.pytorch.org/whl/" + autoTag);
            std::wstring torchCmd = Utf8ToWide("\"" + python + "\" -m pip install torch --upgrade --target=\"" + spDir + "\" --index-url " + autoIdx);
            std::string torchOut; RunProcess(torchCmd, torchOut, 1800000);
            DebugLog("Auto torch install (" + autoTag + "): " + torchOut.substr(0, 500));
        }

        if(!pipOk){
            MsgBox(g_wnd, installPkg + " \xe3\x81\xae\xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe3\x81\xab\xe5\xa4\xb1\xe6\x95\x97\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x97\xe3\x81\x9f\xe3\x80\x82\n\n" + pipOut, "Error", MB_OK|MB_ICONERROR);
            cleanup(); return false;
        }

        // Retry transcription
        SetStatus("\xe5\x86\x8d\xe8\xa9\xa6\xe8\xa1\x8c\xe4\xb8\xad...");
        SetProgress(60);
        std::string pyOut2;
        bool pyOk2 = RunProcess(wCmd, pyOut2, 600000);
        DebugLog("Retry Python exit=" + std::string(pyOk2 ? "0" : "nonzero") + "\n" + pyOut2);

        // Re-read results
        std::ifstream rf2(Utf8ToWide(op));
        if(!rf2.is_open()){
            MsgBox(g_wnd, "\xe3\x83\xaa\xe3\x83\x88\xe3\x83\xa9\xe3\x82\xa4\xe5\xa4\xb1\xe6\x95\x97", "Error", MB_OK|MB_ICONERROR);
            cleanup(); return false;
        }
        std::string line2;
        while(std::getline(rf2, line2)){
            if(line2.empty()) continue;
            if(line2.substr(0, 5) == "ERROR"){
                MsgBox(g_wnd, line2.substr(6), "Error", MB_OK|MB_ICONERROR);
                rf2.close(); cleanup(); return false;
            }
            size_t p1 = line2.find('|'); if(p1 == std::string::npos) continue;
            size_t p2 = line2.find('|', p1+1); if(p2 == std::string::npos) continue;
            size_t p3 = line2.find('|', p2+1); if(p3 == std::string::npos) continue;
            Seg seg;
            int ci = atoi(line2.substr(0, p1).c_str());
            seg.s = atoi(line2.substr(p1+1, p2-p1-1).c_str());
            seg.e = atoi(line2.substr(p2+1, p3-p2-1).c_str());
            seg.text = line2.substr(p3+1);
            if(ci >= 0 && ci < (int)g_tlClips.size()){
                if(seg.s < g_tlClips[ci].timelineStart) seg.s = g_tlClips[ci].timelineStart;
                if(seg.e > g_tlClips[ci].timelineEnd) seg.e = g_tlClips[ci].timelineEnd;
            }
            if(seg.e > seg.s && !seg.text.empty()) g_segs.push_back(seg);
        }
        rf2.close();

        // Re-read err file
        std::string errC2;
        {
            std::ifstream ef(Utf8ToWide(errP));
            if(ef.is_open()){
                std::string l;
                while(std::getline(ef, l)) errC2 += l + "\n";
            }
        }
        if(!errC2.empty()) DebugLog(".err (retry):\n" + errC2);
    }
    cleanup();
    return true;
}

// =========================================================================
// Main transcription thread
// =========================================================================

static void TranscribeThread(){
    SetBusy(true); SetProgress(0);
    DWORD startTick = GetTickCount();
    {std::ofstream f(Utf8ToWide(GetPluginDir() + "\\whisper_debug.log"), std::ios::trunc); f << "=== Whisper Subtitle v2.9.0 ===\n";}
    SaveSettings();
    char lt[16] = {}; GetWindowTextA(g_layerEdit, lt, sizeof(lt));
    int uiL = atoi(lt); if(uiL < 2 || uiL > 100) uiL = 2;
    int apiStartLayer = uiL - 1; // UI Layer 2 = API layer 1
    char mct[16] = {}; GetWindowTextA(g_maxCharEdit, mct, sizeof(mct));
    int maxC = atoi(mct); if(maxC < 0) maxC = 0;
    int mpOn = (SendMessageA(g_chkMorphSplit, BM_GETCHECK, 0, 0) == BST_CHECKED) ? 1 : 0; // v2.8.5: outer scope (fugashi check / simple-path grouping で再利用)
    int maxLines = (SendMessageA(g_chkTwoLine, BM_GETCHECK, 0, 0) == BST_CHECKED) ? 2 : 1; // v2.8.5
    bool fugashiMissing = false; // v2.8.5
    if(mpOn){
        std::string py = GetEffectivePython();
        std::string sp = GetSitePackagesDir();
        // helper と同じ様に site-packages を sys.path に載せて import を試す
        std::string chk = "\"" + py + "\" -c \"import sys,os; sys.path.insert(0, r'" + sp + "'); import fugashi\"";
        std::string o;
        fugashiMissing = !RunProcess(Utf8ToWide(chk), o, 60000);
    }
    // v2.8.2: 生成時の各設定をログに記録 (設定ごとの結果比較用)。どの設定でこの字幕群を作ったか一目で分かる。
    {
        int bkIdx = (int)SendMessageA(g_backendCombo, CB_GETCURSEL, 0, 0);
        int mdIdx = (int)SendMessageA(g_modelCombo, CB_GETCURSEL, 0, 0);
        int mgOn = (SendMessageA(g_chkMergeSeg, BM_GETCHECK, 0, 0) == BST_CHECKED) ? 1 : 0;
        char leadLogBuf[16] = {}; GetWindowTextA(g_leadEdit, leadLogBuf, sizeof(leadLogBuf)); // v2.8.2a
        DebugLog(std::string("========== SETTINGS: backend=") + std::to_string(bkIdx)
            + " model=" + std::to_string(mdIdx)
            + " maxchars=" + std::to_string(maxC)
            + " merge_seg=" + std::to_string(mgOn)
            + " morph_split(pro)=" + std::to_string(mpOn)
            + " two_line=" + std::to_string(maxLines>1?1:0)
            + " lead=" + leadLogBuf + " ==========");
    }
    int mi = SendMessageA(g_modelCombo, CB_GETCURSEL, 0, 0);
    int di = SendMessageA(g_deviceCombo, CB_GETCURSEL, 0, 0);
    int bi = SendMessageA(g_backendCombo, CB_GETCURSEL, 0, 0);
    char qBuf[16] = {}; GetWindowTextA(g_qualityEdit, qBuf, sizeof(qBuf));
    int beamSize = atoi(qBuf); if(beamSize <= 0) beamSize = 5;
    char tBuf2[16] = {}; GetWindowTextA(g_tempEdit, tBuf2, sizeof(tBuf2));
    float temp = (float)atof(tBuf2); if(temp < 0) temp = 0;
    int li = SendMessageA(g_langCombo, CB_GETCURSEL, 0, 0);
    if(mi == CB_ERR) mi = 0; if(di == CB_ERR) di = 0;
    if(bi == CB_ERR) bi = 0; if(li == CB_ERR) li = 1;
    bool removePunct = (SendMessageA(g_chkRemovePunct, BM_GETCHECK, 0, 0) == BST_CHECKED);
    bool removeExclam = (SendMessageA(g_chkRemoveExclam, BM_GETCHECK, 0, 0) == BST_CHECKED);
    bool normalizeText = (SendMessageA(g_chkNormalize, BM_GETCHECK, 0, 0) == BST_CHECKED);
    // v2.8 settings
    bool noPrevText = (SendMessageA(g_chkNoPrevText, BM_GETCHECK, 0, 0) == BST_CHECKED);
    bool wordTs = (SendMessageA(g_chkWordTs, BM_GETCHECK, 0, 0) == BST_CHECKED);
    bool repPenalty = (SendMessageA(g_chkRepPenalty, BM_GETCHECK, 0, 0) == BST_CHECKED);
    bool vadFilter = (SendMessageA(g_chkVad, BM_GETCHECK, 0, 0) == BST_CHECKED); // v2.9.5
    std::string promptText;
    {
        int pLen = GetWindowTextLengthA(g_promptEdit);
        if(pLen > 0){
            promptText.resize(pLen + 1);
            GetWindowTextA(g_promptEdit, &promptText[0], pLen + 1);
            promptText.resize(pLen);
        }
    }
    // v2.8 settings continued
    bool batched = (SendMessageA(g_chkBatched, BM_GETCHECK, 0, 0) == BST_CHECKED);
    std::string hotwordsText;
    {
        int hLen = GetWindowTextLengthA(g_hotwordsEdit);
        if(hLen > 0){
            hotwordsText.resize(hLen + 1);
            GetWindowTextA(g_hotwordsEdit, &hotwordsText[0], hLen + 1);
            hotwordsText.resize(hLen);
        }
    }
    DebugLog("Template: " + (g_templateContent.empty() ? std::string("none") : g_templatePath));

    std::string ffmpeg = GetEffectiveFFmpeg();
    if(ffmpeg.empty()){
        MsgBox(g_wnd,
            "ffmpeg.exe\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n\n"
            "\xe3\x80\x8c" "ffmpeg\xe9\x81\xb8\xe6\x8a\x9e\xe3\x80\x8d\xe3\x83\x9c\xe3\x82\xbf\xe3\x83\xb3\xe3\x81\xa7\xe6\x8c\x87\xe5\xae\x9a\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84",
            "Error", MB_OK|MB_ICONERROR);
        SetStatus("Ready (v2.9.0)"); SetProgress(0); SetBusy(false); return;
    }
    std::string python = GetEffectivePython();
    if(python.empty()){
        MsgBox(g_wnd,
            "Python\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n\n"
            "\xe3\x80\x8c\xe5\x88\x9d\xe6\x9c\x9f\xe8\xa8\xad\xe5\xae\x9a\xe3\x80\x8d\xe3\x81\xa7\xe3\x82\xbb\xe3\x83\x83\xe3\x83\x88\xe3\x82\xa2\xe3\x83\x83\xe3\x83\x97\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84",
            "Error", MB_OK|MB_ICONERROR);
        SetStatus("Ready (v2.9.0)"); SetProgress(0); SetBusy(false); return;
    }

    SetStatus("\xe3\x82\xbf\xe3\x82\xa4\xe3\x83\xa0\xe3\x83\xa9\xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x82\xad\xe3\x83\xa3\xe3\x83\xb3\xe4\xb8\xad...");
    SetProgress(10);
    g_tlClips.clear();
    g_segs.clear();
    ScanParam sp;
    sp.clips = &g_tlClips;
    sp.rate = 30;
    sp.maxLayer = 0;
    if(g_edit) g_edit->call_edit_section_param(&sp, ScanCallback);
    g_projectRate = sp.rate;
    if(g_tlClips.empty()){
        MsgBox(g_wnd,
            "\xe5\x8b\x95\xe7\x94\xbb/\xe9\x9f\xb3\xe5\xa3\xb0\xe3\x82\xaf\xe3\x83\xaa\xe3\x83\x83\xe3\x83\x97\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n"
            "\xe3\x82\xbf\xe3\x82\xa4\xe3\x83\xa0\xe3\x83\xa9\xe3\x82\xa4\xe3\x83\xb3\xe3\x81\xab\xe5\x8b\x95\xe7\x94\xbb/\xe9\x9f\xb3\xe5\xa3\xb0\xe3\x82\x92\xe9\x85\x8d\xe7\xbd\xae\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84",
            "Error", MB_OK|MB_ICONERROR);
        SetStatus("Ready (v2.9.0)"); SetProgress(0); SetBusy(false); return;
    }
    DebugLog("Clips: " + std::to_string(g_tlClips.size()) + " Rate: " + std::to_string(g_projectRate));
    SetProgress(20);

    if(!RunFasterWhisper(mi, di, bi, beamSize, li, temp, noPrevText, wordTs, repPenalty, promptText, hotwordsText, batched, vadFilter)){
        SetStatus("Ready (v2.9.0)"); SetProgress(0); SetBusy(false); return;
    }

    if(g_segs.empty()){
        MsgBox(g_wnd,
            "\xe6\x96\x87\xe5\xad\x97\xe8\xb5\xb7\xe3\x81\x93\xe3\x81\x97\xe7\xb5\x90\xe6\x9e\x9c\xe3\x81\x8c\xe7\xa9\xba\xe3\x81\xa7\xe3\x81\x99\xe3\x80\x82\n"
            "\xe9\x9f\xb3\xe5\xa3\xb0\xe3\x83\x95\xe3\x82\xa1\xe3\x82\xa4\xe3\x83\xab\xe3\x82\x92\xe7\xa2\xba\xe8\xaa\x8d\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82",
            "\xe7\xb5\x90\xe6\x9e\x9c", MB_OK|MB_ICONWARNING);
        SetStatus("Ready (v2.9.0)"); SetProgress(0); SetBusy(false); return;
    }

    // Create subtitle objects
    SetStatus("\xe5\xad\x97\xe5\xb9\x95\xe9\x85\x8d\xe7\xbd\xae\xe4\xb8\xad...");
    SetProgress(80);

    // Apply text processing to all segments
    for(auto& seg : g_segs){
        std::string t = seg.text;
        if(removePunct || removeExclam){
            std::string cleaned;
            for(size_t i = 0; i < t.size(); ){
                unsigned char c = (unsigned char)t[i];
                // ASCII: . , (punct)  ! ? (exclam)
                if(removePunct && (c == '.' || c == ',')){ i++; continue; }
                if(removeExclam && (c == '!' || c == '?')){ i++; continue; }
                // Japanese 。(E3 80 82) 、(E3 80 81) ・(E3 83 BB)
                if(c == 0xe3 && i+2 < t.size()){
                    unsigned char b1 = (unsigned char)t[i+1], b2 = (unsigned char)t[i+2];
                    if(removePunct && b1 == 0x80 && (b2 == 0x81 || b2 == 0x82)){ i += 3; continue; }
                    if(removePunct && b1 == 0x83 && b2 == 0xbb){ i += 3; continue; }
                }
                // Fullwidth ！(EF BC 81) ？(EF BC 9F) → exclam
                // Fullwidth ．(EF BC 8E) ，(EF BC 8C) → punct
                if(c == 0xef && i+2 < t.size()){
                    unsigned char b1 = (unsigned char)t[i+1], b2 = (unsigned char)t[i+2];
                    if(removeExclam && b1 == 0xbc && (b2 == 0x81 || b2 == 0x9f)){ i += 3; continue; }
                    if(removePunct && b1 == 0xbc && (b2 == 0x8e || b2 == 0x8c)){ i += 3; continue; }
                }
                // Copy character (handle multi-byte UTF-8)
                if(c < 0x80){ cleaned += t[i]; i++; }
                else if(c < 0xe0){ cleaned += t.substr(i, 2); i += 2; }
                else if(c < 0xf0){ cleaned += t.substr(i, 3); i += 3; }
                else { cleaned += t.substr(i, 4); i += 4; }
            }
            t = cleaned;
        }
        if(normalizeText){
            std::string tmp;
            for(size_t i = 0; i < t.size(); ){
                unsigned char c = t[i];
                if(c == 0xef && i+2 < t.size()){
                    unsigned char b1 = t[i+1], b2 = t[i+2];
                    if(b1 == 0xbc && b2 >= 0x81){ char a = (char)(b2 - 0x81 + 0x21); if(a >= 0x21 && a <= 0x5f){ tmp += a; i += 3; continue; } }
                    if(b1 == 0xbd && b2 <= 0x9e){ char a = (char)(b2 - 0x80 + 0x60); if(a >= 0x60 && a <= 0x7e){ tmp += a; i += 3; continue; } }
                }
                tmp += t[i]; i++;
            }
            t = tmp;
        }
        while(!t.empty() && (t[0] == ' ' || t[0] == '\t')) t.erase(0, 1);
        while(!t.empty() && (t.back() == ' ' || t.back() == '\t')) t.pop_back();
        seg.text = t;
    }
    g_segs.erase(std::remove_if(g_segs.begin(), g_segs.end(), [](const Seg& s){ return s.text.empty(); }), g_segs.end());

    // Sort segments by start frame
    std::sort(g_segs.begin(), g_segs.end(), [](const Seg& a, const Seg& b){ return a.s < b.s; });
    DebugLog("Segments before trim: " + std::to_string(g_segs.size()));

    // v2.8.2: Optional segment merging (opt-in checkbox, default OFF = original behavior).
    // Merges adjacent segments when the silence gap between them is small AND the combined
    // byte length stays within maxchars. This makes maxchars behave as a "target" instead of
    // only an "upper bound", fixing the 2-5 char fragmentation reported in the issue.
    // Byte-length is compared (same unit as SplitText) so no re-splitting happens downstream.
    // maxchars=0 (文字数で分割しない設定) でも結合は機能させる。その場合の結合上限は既定 20文字。
    // v2.8.2: 形態素分割ONのときは Python側で文節分割済みなので C++結合はスキップ(二重処理防止)。
    if(SendMessageA(g_chkMergeSeg, BM_GETCHECK, 0, 0) == BST_CHECKED &&
       SendMessageA(g_chkMorphSplit, BM_GETCHECK, 0, 0) != BST_CHECKED){
        int mergeLimit = (maxC > 0) ? maxC : 20; // 文字数 (v2.8.2で文字数ベースに統一)
        int mergeGapFrames = (int)(0.4 * g_projectRate); // gaps <= 0.4s treated as soft breaks
        std::vector<Seg> merged;
        for(auto& seg : g_segs){
            if(!merged.empty()){
                Seg& last = merged.back();
                int gap = seg.s - last.e;
                // space separator only between two ASCII alphanumerics (English); none for CJK
                std::string sep;
                if(!last.text.empty() && !seg.text.empty()){
                    unsigned char lc = (unsigned char)last.text.back(), fc = (unsigned char)seg.text.front();
                    bool la = (lc>='0'&&lc<='9')||(lc>='a'&&lc<='z')||(lc>='A'&&lc<='Z');
                    bool fa = (fc>='0'&&fc<='9')||(fc>='a'&&fc<='z')||(fc>='A'&&fc<='Z');
                    if(la && fa) sep = " ";
                }
                if(gap <= mergeGapFrames &&
                   (Utf8CharCount(last.text) + Utf8CharCount(seg.text) + (int)sep.size()) <= mergeLimit){
                    last.text += sep + seg.text;
                    if(seg.e > last.e) last.e = seg.e;
                    continue;
                }
            }
            merged.push_back(seg);
        }
        g_segs = merged;
        DebugLog("Merge(v2.8.2): " + std::to_string(g_segs.size()) + " segments after merging");
    }

    // Apply max char splitting FIRST (before trimming)
    struct PlaceItem { int s, e; std::string text; };
    std::vector<PlaceItem> items;
    int cLines = mpOn ? 1 : maxLines; // v2.8.5: morph ON は Python がグルーピング済み
    for(auto& seg : g_segs){
        if(seg.e <= seg.s || seg.text.empty()) continue;
        std::vector<std::string> parts = SplitText(seg.text, maxC);
        if(parts.size() <= 1){
            items.push_back({seg.s, seg.e, seg.text});
        } else {
            int total = seg.e - seg.s;
            int partLen = total / (int)parts.size();
            if(partLen < 1) partLen = 1;
            // v2.8.5: parts を cLines 個ずつ束ねて 1 テロップ(各行 \n 連結)にする
            for(size_t gi = 0; gi < parts.size(); gi += cLines){
                size_t gEnd = gi + cLines; if(gEnd > parts.size()) gEnd = parts.size();
                std::string joined;
                for(size_t pi = gi; pi < gEnd; pi++){
                    if(pi > gi) joined += "\\n"; // リテラル改行(0x5C 0x6E)
                    joined += parts[pi];
                }
                int ps = seg.s + (int)gi * partLen;
                int pe = (gEnd >= parts.size()) ? seg.e : (seg.s + (int)gEnd * partLen);
                if(pe > ps && !joined.empty())
                    items.push_back({ps, pe, joined});
            }
        }
    }

    // Extend subtitle display: keep showing after speech ends
    char lingerBuf[16] = {}; GetWindowTextA(g_lingerEdit, lingerBuf, sizeof(lingerBuf));
    double lingerSec = atof(lingerBuf);
    if(lingerSec < 0) lingerSec = 0;
    if(lingerSec > 10) lingerSec = 10;
    int lingerFrames = (int)(lingerSec * g_projectRate);
    if(lingerFrames > 0){
        int timelineEnd = 0;
        for(auto& cl : g_tlClips) if(cl.timelineEnd > timelineEnd) timelineEnd = cl.timelineEnd;
        for(size_t i = 0; i < items.size(); i++){
            items[i].e = items[i].e + lingerFrames;
            if(timelineEnd > 0 && items[i].e > timelineEnd)
                items[i].e = timelineEnd;
        }
    }

    // v2.8.2a: Lead time - show subtitle N seconds before speech starts
    char leadBuf[16] = {}; GetWindowTextA(g_leadEdit, leadBuf, sizeof(leadBuf));
    double leadSec = atof(leadBuf);
    if(leadSec < 0) leadSec = 0;
    if(leadSec > 5) leadSec = 5;
    int leadFrames = (int)(leadSec * g_projectRate);
    if(leadFrames > 0){
        for(size_t i = 0; i < items.size(); i++){
            items[i].s = items[i].s - leadFrames;
            if(items[i].s < 0) items[i].s = 0;
        }
    }

    // Clip overlapping items (e is used as: len = e - s)
    for(size_t i = 1; i < items.size(); i++){
        if(items[i].s < items[i-1].e){
            items[i-1].e = items[i].s;
        }
    }

    // Remove items that became too short (need at least 2 frames)
    std::vector<PlaceItem> finalItems;
    for(auto& it : items){
        if(it.e > it.s && !it.text.empty())
            finalItems.push_back(it);
    }
    DebugLog("Items after trim: " + std::to_string(finalItems.size()));

    int targetLayer = apiStartLayer;
    const int MIN_GAP = 0;

    // Log all items for debugging
    for(size_t i = 0; i < finalItems.size(); i++){
        DebugLog("Item " + std::to_string(i) + ": [" + std::to_string(finalItems[i].s) + "-" + std::to_string(finalItems[i].e) + "] \"" + finalItems[i].text.substr(0, 90) + "\"");
    }

    // Greedy bin-packing: track end frame per layer, assign each item to
    // the first layer where it fits (with MIN_GAP gap)
    std::vector<int> layerEnds; // layerEnds[i] = last used end frame on layer i
    std::vector<int> itemLayers(finalItems.size()); // which layer each item goes to

    for(size_t i = 0; i < finalItems.size(); i++){
        int assigned = -1;
        for(size_t li = 0; li < layerEnds.size(); li++){
            if(finalItems[i].s >= layerEnds[li] + MIN_GAP){
                assigned = (int)li;
                layerEnds[li] = finalItems[i].e;
                break;
            }
        }
        if(assigned < 0){
            assigned = (int)layerEnds.size();
            layerEnds.push_back(finalItems[i].e);
        }
        itemLayers[i] = assigned;
    }
    DebugLog("Packing: " + std::to_string(finalItems.size()) + " items into " + std::to_string(layerEnds.size()) + " layers (gap=" + std::to_string(MIN_GAP) + ")");

    // === PRE-CHECK: Find empty layer range for placement ===
    int numLayersNeeded = (int)layerEnds.size();
    struct LayerCheckParam {
        std::vector<PlaceItem>* items;
        int startLayer;
        int numLayers;
        bool hasConflict;
    };
    // Check if any existing objects overlap with our planned items
    auto findFreeLayer = [&](){
        LayerCheckParam lc;
        lc.items = &finalItems;
        lc.startLayer = targetLayer;
        lc.numLayers = numLayersNeeded;
        lc.hasConflict = false;

        auto checkCallback = [](void* param, EDIT_SECTION* es){
            LayerCheckParam* lc = (LayerCheckParam*)param;
            if(!es) return;
            // Sample several frame positions across our items to check for existing objects
            for(int lay = 0; lay < lc->numLayers; lay++){
                int apiLayer = lc->startLayer + lay;
                for(auto& item : *lc->items){
                    // Check start, middle, and a few points
                    int checkFrames[] = {item.s, (item.s + item.e) / 2, item.e - 1};
                    for(int f : checkFrames){
                        if(f < 0) continue;
                        OBJECT_HANDLE obj = es->find_object(apiLayer, f);
                        if(obj){
                            lc->hasConflict = true;
                            return;
                        }
                    }
                }
            }
        };
        if(g_edit) g_edit->call_edit_section_param(&lc, checkCallback);
        return lc.hasConflict;
    };

    // Shift targetLayer until we find a clear range (max 50 layers)
    int maxShift = 50;
    for(int shift = 0; shift < maxShift; shift++){
        if(!findFreeLayer()) break;
        DebugLog("Layer " + std::to_string(targetLayer + 1) + " occupied, shifting...");
        targetLayer++;
    }
    DebugLog("Target layer: " + std::to_string(targetLayer + 1) + " (API " + std::to_string(targetLayer) + ")");

    // === PASS 1: Place ALL with create_object (100% reliable) ===
    struct Pass1Param {
        std::vector<PlaceItem>* items;
        std::vector<int>* layers;
        int targetLayer;
        int placed;
        int failed;
    };
    Pass1Param p1;
    p1.items = &finalItems;
    p1.layers = &itemLayers;
    p1.targetLayer = targetLayer;
    p1.placed = 0;
    p1.failed = 0;

    auto pass1Callback = [](void* param, EDIT_SECTION* es){
        Pass1Param* p = (Pass1Param*)param;
        if(!es) return;
        const wchar_t *wT=L"\x30c6\x30ad\x30b9\x30c8", *wD=L"\x6a19\x6e96\x63cf\x753b",
            *wF=L"\x30d5\x30a9\x30f3\x30c8", *wS=L"\x30b5\x30a4\x30ba",
            *wC=L"\x6587\x5b57\x8272", *wA=L"\x6587\x5b57\x63c3\x3048";
        for(size_t idx = 0; idx < p->items->size(); idx++){
            auto& item = (*p->items)[idx];
            int apiLayer = p->targetLayer + (*p->layers)[idx];
            int len = item.e - item.s; if(len <= 0) len = 1;
            OBJECT_HANDLE obj = es->create_object(wT, apiLayer, item.s, len);
            if(obj){
                es->set_object_item_value(obj, wT, wT, item.text.c_str());
                es->set_object_item_value(obj, wT, wF, "Yu Gothic UI");
                es->set_object_item_value(obj, wT, wS, "60.00");
                es->set_object_item_value(obj, wT, wC, "ffffff");
                es->set_object_item_value(obj, wT, wA,
                    "\xe4\xb8\xad\xe5\xa4\xae\xe6\x8f\x83\xe3\x81\x88[\xe4\xb8\x8b]");
                es->set_object_item_value(obj, wD, L"Y", "400.00");
                p->placed++;
            } else {
                p->failed++;
            }
        }
    };

    if(g_edit) g_edit->call_edit_section_param(&p1, pass1Callback);
    int placed = p1.placed;
    DebugLog("Pass1 placed: " + std::to_string(placed) + " failed: " + std::to_string(p1.failed));

    // === PASS 2: If template, replace each one-by-one (separate edit section) ===
    if(!g_templateContent.empty() && placed > 0){
        struct Pass2Param {
            std::vector<PlaceItem>* items;
            std::vector<int>* layers;
            int targetLayer;
            std::string tplContent;
            int replaced;
            int rFailed;
            std::string logPath;
        };
        Pass2Param p2;
        p2.items = &finalItems;
        p2.layers = &itemLayers;
        p2.targetLayer = targetLayer;
        p2.tplContent = g_templateContent;
        p2.replaced = 0;
        p2.rFailed = 0;
        p2.logPath = GetPluginDir() + "\\whisper_debug.log";

        auto pass2Callback = [](void* param, EDIT_SECTION* es){
            Pass2Param* p = (Pass2Param*)param;
            if(!es) return;
            std::string textKey = "\xe3\x83\x86\xe3\x82\xad\xe3\x82\xb9\xe3\x83\x88=";
            const wchar_t *wT=L"\x30c6\x30ad\x30b9\x30c8", *wD=L"\x6a19\x6e96\x63cf\x753b",
                *wF=L"\x30d5\x30a9\x30f3\x30c8", *wS=L"\x30b5\x30a4\x30ba",
                *wC=L"\x6587\x5b57\x8272", *wA=L"\x6587\x5b57\x63c3\x3048";
            auto cbLog = [&](const std::string& msg){
                if(!p->logPath.empty()){
                    std::ofstream lf(p->logPath, std::ios::app);
                    lf << msg << "\n";
                }
            };

            // Step A: Delete ALL objects (make timeline completely empty)
            for(size_t idx = 0; idx < p->items->size(); idx++){
                auto& item = (*p->items)[idx];
                int apiLayer = p->targetLayer + (*p->layers)[idx];
                OBJECT_HANDLE existing = es->find_object(apiLayer, item.s);
                if(existing){
                    es->delete_object(existing);
                } else {
                    cbLog("PASS2_NOFIND #" + std::to_string(idx) + " L" + std::to_string(apiLayer) + " F" + std::to_string(item.s));
                }
            }

            // Step B: Recreate ALL with template (timeline is empty, no collisions)
            for(size_t idx = 0; idx < p->items->size(); idx++){
                auto& item = (*p->items)[idx];
                int apiLayer = p->targetLayer + (*p->layers)[idx];
                int len = item.e - item.s; if(len <= 0) len = 1;

                std::string alias = p->tplContent;
                size_t pos = alias.find(textKey);
                if(pos != std::string::npos){
                    size_t vs = pos + textKey.size();
                    size_t le = alias.find('\n', vs);
                    if(le == std::string::npos) le = alias.size();
                    alias = alias.substr(0, vs) + item.text + alias.substr(le);
                }

                OBJECT_HANDLE obj = es->create_object_from_alias(alias.c_str(), apiLayer, item.s, len);
                if(obj){
                    p->replaced++;
                } else {
                    if(p->rFailed == 0){
                        cbLog("FIRST_FAIL_ALIAS:\n" + alias.substr(0, 500) + "\nEND_ALIAS");
                    }
                    cbLog("PASS2_ALIAS_FAIL #" + std::to_string(idx) + " L" + std::to_string(apiLayer) + " [" + std::to_string(item.s) + "-" + std::to_string(item.e) + "] len=" + std::to_string(len));
                    // Restore with create_object
                    obj = es->create_object(wT, apiLayer, item.s, len);
                    if(obj){
                        es->set_object_item_value(obj, wT, wT, item.text.c_str());
                        es->set_object_item_value(obj, wT, wF, "Yu Gothic UI");
                        es->set_object_item_value(obj, wT, wS, "60.00");
                        es->set_object_item_value(obj, wT, wC, "ffffff");
                        es->set_object_item_value(obj, wT, wA,
                            "\xe4\xb8\xad\xe5\xa4\xae\xe6\x8f\x83\xe3\x81\x88[\xe4\xb8\x8b]");
                        es->set_object_item_value(obj, wD, L"Y", "400.00");
                    }
                    p->rFailed++;
                }
            }
        };

        if(g_edit) g_edit->call_edit_section_param(&p2, pass2Callback);
        DebugLog("Pass2 replaced: " + std::to_string(p2.replaced) + " failed: " + std::to_string(p2.rFailed));
    }
    DebugLog("Placed: " + std::to_string(placed) + " failed: " + std::to_string(p1.failed));

    SetProgress(100);
    DWORD elapsed = (GetTickCount() - startTick) / 1000;
    char timeBuf[64];
    if(elapsed >= 60) sprintf_s(timeBuf, " %dm%02ds", (int)(elapsed/60), (int)(elapsed%60));
    else sprintf_s(timeBuf, " %ds", (int)elapsed);
    std::string warn; // v2.8.5
    if(fugashiMissing) warn = "  \xe2\x9a\xa0 fugashi\xe6\x9c\xaa\xe5\xb0\x8e\xe5\x85\xa5\xe3\x81\xae\xe3\x81\x9f\xe3\x82\x81\xe6\x96\x87\xe7\xaf\x80\xe5\x8c\xba\xe5\x88\x87\xe3\x82\x8a\xe3\x81\x8c\xe8\xa1\x8c\xe3\x82\x8f\xe3\x82\x8c\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x81\xa7\xe3\x81\x97\xe3\x81\x9f"; // v2.8.9: F「プロ分割」→「文節区切り」用語統一
    SetStatus("Done! " + std::to_string(placed) + "\xe5\x80\x8b\xe3\x81\xae\xe5\xad\x97\xe5\xb9\x95\xe3\x82\x92\xe9\x85\x8d\xe7\xbd\xae (" + std::to_string(layerEnds.size()) + "Layer)" + timeBuf + warn);
    DebugLog("Placed: " + std::to_string(placed));
    SetBusy(false);
}

// =========================================================================
// SRT Export
// =========================================================================

static void ExportSRT(){
    if(g_segs.empty()){
        MsgBox(g_wnd, "\xe5\xad\x97\xe5\xb9\x95\xe3\x83\x87\xe3\x83\xbc\xe3\x82\xbf\xe3\x81\x8c\xe3\x81\x82\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93", "SRT Export", MB_OK|MB_ICONWARNING);
        return;
    }
    wchar_t fn[MAX_PATH] = L"subtitle.srt";
    OPENFILENAMEW ofn = {sizeof(ofn)};
    ofn.hwndOwner = g_wnd;
    ofn.lpstrFilter = L"SRT\0*.srt\0All\0*.*\0";
    ofn.lpstrFile = fn; ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_OVERWRITEPROMPT;
    ofn.lpstrDefExt = L"srt";
    if(!GetSaveFileNameW(&ofn)) return;

    // Apply linger to SRT (same as timeline placement)
    char lingerBuf[16] = {}; GetWindowTextA(g_lingerEdit, lingerBuf, sizeof(lingerBuf));
    double lingerSec = atof(lingerBuf);
    if(lingerSec < 0) lingerSec = 0;
    if(lingerSec > 10) lingerSec = 10;
    int lingerFrames = (int)(lingerSec * g_projectRate);

    struct SrtSeg { int s, e; std::string text; };
    std::vector<SrtSeg> srtSegs;
    for(auto& seg : g_segs) srtSegs.push_back({seg.s, seg.e, seg.text});
    // Extend
    if(lingerFrames > 0){
        for(auto& s : srtSegs) s.e += lingerFrames;
    }
    // v2.8.2a: Apply lead time to SRT (same as timeline placement)
    char leadBuf2b[16] = {}; GetWindowTextA(g_leadEdit, leadBuf2b, sizeof(leadBuf2b));
    double leadSec2 = atof(leadBuf2b);
    if(leadSec2 < 0) leadSec2 = 0;
    if(leadSec2 > 5) leadSec2 = 5;
    int leadFrames2 = (int)(leadSec2 * g_projectRate);
    if(leadFrames2 > 0){
        for(auto& s : srtSegs){
            s.s -= leadFrames2;
            if(s.s < 0) s.s = 0;
        }
    }
    // Clip overlaps
    for(size_t i = 1; i < srtSegs.size(); i++){
        if(srtSegs[i].s < srtSegs[i-1].e)
            srtSegs[i-1].e = srtSegs[i].s;
    }

    std::ofstream f(fn);
    int idx = 1;
    for(auto& seg : srtSegs){
        double ss = (double)seg.s / g_projectRate, se = (double)seg.e / g_projectRate;
        int sh = (int)(ss/3600), sm = (int)(fmod(ss,3600)/60), ssc = (int)fmod(ss,60), sms = (int)(fmod(ss,1)*1000);
        int eh = (int)(se/3600), em = (int)(fmod(se,3600)/60), esc2 = (int)fmod(se,60), ems = (int)(fmod(se,1)*1000);
        char buf[128];
        sprintf_s(buf, "%02d:%02d:%02d,%03d --> %02d:%02d:%02d,%03d", sh,sm,ssc,sms, eh,em,esc2,ems);
        f << idx++ << "\n" << buf << "\n" << seg.text << "\n\n";
    }
    MsgBox(g_wnd, "SRT\xe3\x82\xa8\xe3\x82\xaf\xe3\x82\xb9\xe3\x83\x9d\xe3\x83\xbc\xe3\x83\x88\xe5\xae\x8c\xe4\xba\x86", "SRT", MB_OK|MB_ICONINFORMATION);
}

// =========================================================================
// Window procedure
// =========================================================================

static LRESULT CALLBACK WndProc(HWND h, UINT m, WPARAM w, LPARAM l){
    if(m == WM_COMMAND){
        int id = LOWORD(w);
        if(HIWORD(w) == BN_CLICKED){
            if((HWND)l == g_chkMorphSplit){
                bool on = SendMessageA(g_chkMorphSplit, BM_GETCHECK, 0, 0) == BST_CHECKED;
                EnableWindow(g_chkMergeSeg, on ? FALSE : TRUE); // v2.8.5: プロ分割ON時はセグメント結合を無効化
            }
            if(id == IDC_GENERATE){
                // v2.9.0【I】: 生成を押した時点で不足を検出して導入へ誘導する。
                // 判定は全てファイル存在 or probe キャッシュ参照で、python は起動しない(生成が遅くならないこと)。
                std::string lack = CheckMissingForGenerate();
                if(!lack.empty()){
                    int r = MsgBox(g_wnd,
                        "\xe4\xbb\x8a\xe3\x81\xae\xe8\xa8\xad\xe5\xae\x9a\xe3\x81\xa7\xe3\x81\xaf\xe4\xbb\xa5\xe4\xb8\x8b\xe3\x81\x8c\xe5\xbf\x85\xe8\xa6\x81\xe3\x81\xa7\xe3\x81\x99\xe3\x80\x82\n\n"
                        + lack + "\n"
                        "\xe4\xbb\x8a\xe3\x81\x99\xe3\x81\x90\xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x99\xe3\x81\x8b\xef\xbc\x9f",
                        "\xe6\xba\x96\xe5\x82\x99\xe3\x81\x8c\xe5\xbf\x85\xe8\xa6\x81", MB_YESNO|MB_ICONINFORMATION);
                    if(r == IDYES){
                        TabCtrl_SetCurSel(g_tab, 3); SwitchTab(3); // 環境タブへ切り替えて進捗を見せる
                        bool exp = false;
                        if(g_busy.compare_exchange_strong(exp, true)){
                            if(g_btnGenerate) EnableWindow(g_btnGenerate, FALSE);
                            if(g_btnSetup) EnableWindow(g_btnSetup, FALSE);
                            std::thread(SetupThread).detach();
                        }
                    }
                    return 0; // v2.9.0【I】: WndProcはLRESULT関数のため戻り値が必要。IDNOのときはg_busyを一切触らずに抜ける
                }
                // v2.8b: atomic check-and-set. compare_exchange_strong only succeeds (and
                // flips g_busy to true) if g_busy was false at the time of the call, so two
                // near-simultaneous clicks can never both pass this gate.
                bool expected = false;
                if(g_busy.compare_exchange_strong(expected, true)){
                    if(g_btnGenerate) EnableWindow(g_btnGenerate, FALSE);
                    if(g_btnSetup) EnableWindow(g_btnSetup, FALSE);
                    std::thread(TranscribeThread).detach();
                }
            }
            else if(id == IDC_EXPORT_SRT) ExportSRT();
            else if(id == IDC_TEMPLATE){
                std::string p = BrowseForFile(g_wnd, L"Object\0*.object\0All\0*.*\0", L"\x66f8\x5f0f\x9078\x629e");
                if(!p.empty() && LoadTemplate(p)){ g_templatePath = p; UpdateTemplateLabel(); SaveSettings(); }
            }
            else if(id == IDC_RESET_TPL){
                g_templatePath.clear(); g_templateContent.clear();
                SetWindowTextW(g_templateLabel, L"\x30c7\x30d5\x30a9\x30eb\x30c8");
                SaveSettings();
            }
            else if(id == IDC_SETUP){
                bool expected = false;
                if(g_busy.compare_exchange_strong(expected, true)){
                    if(g_btnGenerate) EnableWindow(g_btnGenerate, FALSE);
                    if(g_btnSetup) EnableWindow(g_btnSetup, FALSE);
                    std::thread(SetupThread).detach();
                }
            }
            else if(id == IDC_FFMPEG_BR){
                std::string curFf = g_ffmpegPath.empty() ? (GetExeDir() + "\\ffmpeg.exe") : g_ffmpegPath; // v2.8.5
                std::string p = BrowseForFile(g_wnd, L"ffmpeg.exe\0ffmpeg.exe\0All\0*.*\0", L"ffmpeg.exe\x3092\x9078\x629e", curFf);
                if(!p.empty()){
                    g_ffmpegPath = p;
                    SetPathLabel(g_ffmpegLabel, g_ffmpegPath, "(aviutl2.exe\xe3\x81\xae\xe5\xa0\xb4\xe6\x89\x80)");
                    SaveSettings();
                    RefreshSetupTabState(); // v2.8.8: 状態記号/DLボタン表示を即反映
                }
            }
            else if(id == IDC_PYTHON_BR){
                std::string curPy = GetEffectivePython(); // v2.8.5: 手動未設定でも自動検出された実際のpythonが開く
                std::string p = BrowseForFile(g_wnd, L"python.exe\0python.exe\0All\0*.*\0", L"python.exe\x3092\x9078\x629e", curPy);
                if(!p.empty()){
                    g_pythonPath = p;
                    SetPathLabel(g_pythonLabel, g_pythonPath, "(\xe8\x87\xaa\xe5\x8b\x95\xe6\xa4\x9c\xe5\x87\xba)");
                    SaveSettings();
                    RefreshSetupTabState(); // v2.8.8: 状態記号を即反映
                }
            }
            else if(id == IDC_FW_BR){
                std::string curFw = g_fwSpPath.empty() ? GetEffectiveFwSpDir() : g_fwSpPath; // v2.8.5
                std::string p = BrowseForFolder(g_wnd, L"faster-whisper\x306esite-packages\x30d5\x30a9\x30eb\x30c0\x3092\x9078\x629e", curFw);
                if(!p.empty()){
                    // Validate: must contain faster_whisper/__init__.py
                    if(FileExistsU(p + "\\faster_whisper\\__init__.py")){
                        g_fwSpPath = p;
                    } else {
                        MsgBox(g_wnd, "\xe9\x81\xb8\xe6\x8a\x9e\xe3\x81\x97\xe3\x81\x9f\xe3\x83\x95\xe3\x82\xa9\xe3\x83\xab\xe3\x83\x80\xe3\x81\xab" "faster_whisper\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n" "faster_whisper\xe3\x83\x95\xe3\x82\xa9\xe3\x83\xab\xe3\x83\x80\xe3\x81\x8c\xe5\x85\xa5\xe3\x81\xa3\xe3\x81\xa6\xe3\x81\x84\xe3\x82\x8b" "site-packages\xe3\x82\x92\xe9\x81\xb8\xe6\x8a\x9e\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82", "Error", MB_OK|MB_ICONWARNING);
                        g_fwSpPath = p; // set anyway for user convenience
                    }
                    SaveSettings();
                    UpdateWhisperLocLabels();
                }
            }
            else if(id == IDC_OW_BR){
                std::string curOw = g_owSpPath.empty() ? GetEffectiveOwSpDir() : g_owSpPath; // v2.8.5
                std::string p = BrowseForFolder(g_wnd, L"whisper\x306esite-packages\x30d5\x30a9\x30eb\x30c0\x3092\x9078\x629e", curOw);
                if(!p.empty()){
                    if(FileExistsU(p + "\\whisper\\__init__.py")){
                        g_owSpPath = p;
                    } else {
                        MsgBox(g_wnd, "\xe9\x81\xb8\xe6\x8a\x9e\xe3\x81\x97\xe3\x81\x9f\xe3\x83\x95\xe3\x82\xa9\xe3\x83\xab\xe3\x83\x80\xe3\x81\xab" "whisper\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n" "whisper\xe3\x83\x95\xe3\x82\xa9\xe3\x83\xab\xe3\x83\x80\xe3\x81\x8c\xe5\x85\xa5\xe3\x81\xa3\xe3\x81\xa6\xe3\x81\x84\xe3\x82\x8b" "site-packages\xe3\x82\x92\xe9\x81\xb8\xe6\x8a\x9e\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82", "Error", MB_OK|MB_ICONWARNING);
                        g_owSpPath = p;
                    }
                    SaveSettings();
                    UpdateWhisperLocLabels();
                }
            }
            else if(id == IDC_FW_RESET){
                g_fwSpPath.clear();
                SaveSettings();
                UpdateWhisperLocLabels();
            }
            else if(id == IDC_OW_RESET){
                g_owSpPath.clear();
                SaveSettings();
                UpdateWhisperLocLabels();
            }
            // v2.9.0【F】: fugashi単独[再導入]ボタンのハンドラは削除(廃止、チェック方式に統合)。
            // v2.9.4: IDC_DL_FFMPEG ハンドラは廃止(ffmpeg専用ボタンをやめチェック+セットアップに統一)
        }
        else if(HIWORD(w) == CBN_SELCHANGE){
            if((HWND)l == g_modelCombo || (HWND)l == g_deviceCombo || (HWND)l == g_backendCombo || (HWND)l == g_langCombo) SaveSettings();
            // v2.9.5: vad_filter は faster-whisper 経路にしか無い設定。openai-whisper 選択時は
            // 触っても何も起きないのでグレーアウトする(効かないUIを出したままにしない)。
            if((HWND)l == g_backendCombo) UpdateVadEnable();
            if((HWND)l == g_modelCombo) RefreshSetupTabState(); // v2.9.0【G】: モデル切替で環境タブの状態表示を即追従させる
        }
        else if(HIWORD(w) == EN_CHANGE){
            if((HWND)l == g_layerEdit || (HWND)l == g_maxCharEdit || (HWND)l == g_lingerEdit || (HWND)l == g_leadEdit) SaveSettings();
        }
    }
    else if(m == WM_NOTIFY){
        NMHDR* nm = (NMHDR*)l;
        if(nm->hwndFrom == g_tab && nm->code == TCN_SELCHANGE){
            SwitchTab(TabCtrl_GetCurSel(g_tab));
        }
    }
    else if(m == WM_UPDATE_STATUS){
        SetWindowTextW(g_status, Utf8ToWide((char*)l).c_str());
        if(g_statusSetup) SetWindowTextW(g_statusSetup, Utf8ToWide((char*)l).c_str()); // v2.8.8【A】: 環境タブ専用ステータスにも反映
        free((void*)l);
    }
    else if(m == WM_UPDATE_PROGRESS){
        SendMessageA(g_progress, PBM_SETPOS, w, 0);
        if(g_progressSetup) SendMessageA(g_progressSetup, PBM_SETPOS, w, 0); // v2.8.8【A】: 環境タブ専用進捗バーにも反映
    }
    else if(m == WM_PROBE_DONE){ // v2.9.0【A】: ProbeThread完了。実測結果で環境タブの表示を最新化
        RefreshSetupTabState();
    }
    return DefWindowProcW(h, m, w, l);
}

// =========================================================================
// Detect whisper install locations
// =========================================================================

static void UpdateWhisperLocLabels(){
    std::string fwDir = GetEffectiveFwSpDir();
    std::string owDir = GetEffectiveOwSpDir();
    // v2.8.8【E】: 状態記号(●導入済み/○未導入 = U+25CF/U+25CB)を先頭に付与
    const char* markInstalled = "\xe2\x97\x8f\xe5\xb0\x8e\xe5\x85\xa5\xe6\xb8\x88\xe3\x81\xbf"; // "●導入済み" (v2.9.2: 末尾の空白を削除。所在タグを後ろに繋げるため)
    const char* markMissing   = "\xe2\x97\x8b\xe6\x9c\xaa\xe5\xb0\x8e\xe5\x85\xa5";            // "○未導入"

    // v2.9.2: ライブラリ名は別ラベル(固定表示)に移したので、ここは状態と所在タグだけを作る。
    // 所在は絶対パスをやめてタグ化した。理由: whisper も faster-whisper も通常は同じ
    // プラグイン同梱 site-packages に入るため、パスは両者の判別材料にならず幅を食うだけだった。
    // 実際のパスは [選択] ボタンのダイアログが現在地で開くので、そちらで確認できる(v2.8.5の仕様)。
    auto LocTag = [&](const std::string& dir, const std::string& manual) -> std::string {
        if(dir.empty()) return "";
        if(!manual.empty()) return " [\xe6\x89\x8b\xe5\x8b\x95]";        // [手動] ユーザーが明示指定
        if(dir == GetSitePackagesDir()) return " [\xe5\x90\x8c\xe6\xa2\xb1]"; // [同梱] プラグイン内
        return " [system]";                                              // Python本体側
    };
    std::string fwText = (!fwDir.empty() ? std::string(markInstalled) + LocTag(fwDir, g_fwSpPath) : markMissing);
    std::string owText = (!owDir.empty() ? std::string(markInstalled) + LocTag(owDir, g_owSpPath) : markMissing);

    if(g_fwLocLabel) SetWindowTextW(g_fwLocLabel, Utf8ToWide(fwText).c_str());
    if(g_owLocLabel) SetWindowTextW(g_owLocLabel, Utf8ToWide(owText).c_str());
}

static void UpdateFugashiStatus(){ // v2.9.0: python同期実行(最大60秒)をやめ probe キャッシュを読むだけにした
    if(!g_fugashiStatus) return;
    const wchar_t* t;
    if(!g_probe.done)          t = L"\x6587\x7bc0\x533a\x5207\x308a(fugashi): ...\x78ba\x8a8d\x4e2d";
    else if(g_probe.fugashi)   t = L"\x6587\x7bc0\x533a\x5207\x308a(fugashi): \x25cf\x5c0e\x5165\x6e08\x307f";
    else                       t = L"\x6587\x7bc0\x533a\x5207\x308a(fugashi): \x25cb\x672a\x5c0e\x5165";
    SetWindowTextW(g_fugashiStatus, t);
}

// v2.8.8【E】: 環境タブの状態表示(ffmpeg/Python/whisper系の●導入済み・○未導入・DLボタン表示)をまとめて更新。
// SwitchTab()の一括ShowWindowループより後に呼ばれる想定(そうでないとDLボタンのSW_HIDEが上書きされてしまう)。
// v2.9.0: UpdateFugashiStatus()がprobeキャッシュ参照のみ(python同期実行なし)になったため、
// 「fugashiの再判定はここでは絶対に呼ばない」という旧制約は解除。ここから安全に呼べる。
static void RefreshSetupTabState(){
    bool ffOk = !GetEffectiveFFmpeg().empty();
    if(g_ffmpegStatusLabel) SetWindowTextW(g_ffmpegStatusLabel, ffOk ? L"\x25cf\x5c0e\x5165\x6e08\x307f" : L"\x25cb\x672a\x5c0e\x5165"); // ●導入済み / ○未導入
    // v2.9.4: ffmpeg専用ボタンを廃止しチェック方式へ統一したため、ここでのShowWindow/ラベル切替と
    // パスラベル幅のMoveWindow補正(ボタンとの重なり回避用)は不要になった。幅は生成時に確定している。

    bool pyOk = !GetEffectivePython().empty();
    // v2.9.0【B】: probeキャッシュを使い、3.10未満なら▲+バージョンを表示する
    if(g_pythonStatusLabel){
        std::wstring t;
        if(!g_probe.done) t = L"..." L"\x78ba\x8a8d\x4e2d"; // ...確認中
        else if(!pyOk) t = L"\x25cb" L"\x672a\x691c\x51fa"; // ○未検出
        // v2.9.7: 旧実装は「バージョン不明でも実在は確認できている」として ●導入済み を出していたが、
        // **probeがバージョンを取れないPython = 実際には動かないPython** なので設計ミスだった。
        // (Microsoft Storeのスタブがこれに該当し、●導入済みと表示されて全処理が失敗していた)
        else if(g_probe.pyMajor == 0) t = L"\x25b2" L"\x4f7f\x7528\x4e0d\x53ef"; // ▲使用不可
        else {
            bool ok310 = (g_probe.pyMajor > 3) || (g_probe.pyMajor == 3 && g_probe.pyMinor >= 10);
            t = (ok310 ? std::wstring(L"\x25cf") : std::wstring(L"\x25b2"))
              + std::to_wstring(g_probe.pyMajor) + L"." + std::to_wstring(g_probe.pyMinor) + L"." + std::to_wstring(g_probe.pyPatch);
        }
        SetWindowTextW(g_pythonStatusLabel, t.c_str());
    }

    // v2.9.0【H】: PyTorch行。probeキャッシュのCUDA/CPU/未導入を表示
    if(g_torchStatusLabel){
        std::wstring t = L"PyTorch  ";
        if(!g_probe.done) t += L"..." L"\x78ba\x8a8d\x4e2d";
        else if(g_probe.torch && g_probe.torchCuda) t += L"\x25cf\x5c0e\x5165\x6e08\x307f" L" (CUDA)";
        else if(g_probe.torch) t += L"\x25cf\x5c0e\x5165\x6e08\x307f" L" (CPU)";
        else t += L"\x25cb" L"\x672a\x5c0e\x5165" L" (whisper" L"\x7528" L"/" L"\x7d04" L"4.3GB)";
        SetWindowTextW(g_torchStatusLabel, t.c_str());
    }

    // v2.9.0【G】: モデル行。ModelExists()はファイル存在チェックのみで軽量、probe不要
    if(g_modelStatusLabel){
        int mi2 = SendMessageA(g_modelCombo, CB_GETCURSEL, 0, 0);
        if(mi2 == CB_ERR) mi2 = 0;
        if(mi2 < 0 || mi2 >= (int)(sizeof(kModelNames)/sizeof(kModelNames[0]))) mi2 = 0;
        std::string curModel = kModelNames[mi2];
        int curBi = SendMessageA(g_backendCombo, CB_GETCURSEL, 0, 0);
        if(curBi == CB_ERR) curBi = 0;
        // v2.9.2: 「モデル (large-v3-turbo)」だと large-v3-turbo が固定の必須要件に読めてしまい、
        // 「これが無いと動かない」と誤解されるため「選択中のモデル: <name>」に変更。
        // 設定タブで選んだモデルの状態を映しているだけ、と一目で分かる文言にする。
        std::wstring t = L"\x9078\x629e\x4e2d\x306e\x30e2\x30c7\x30eb" L": " + Utf8ToWide(curModel) + L"  "; // 選択中のモデル: name
        if(curModel == "kotoba-whisper" && curBi == 1){
            t += L"\x25b2" L"faster-whisper" L"\x5c02\x7528"; // ▲faster-whisper専用: backend=openai-whisperでは動かない
        } else if(ModelExists(curModel)){
            t += L"\x25cf\x5c0e\x5165\x6e08\x307f";
        } else {
            t += L"\x25cb\x672a\x5c0e\x5165";
        }
        SetWindowTextW(g_modelStatusLabel, t.c_str());
    }

    UpdateWhisperLocLabels(); // faster-whisper/whisperの状態記号もここで最新化
    UpdateFugashiStatus(); // v2.9.0【F】: probeキャッシュ参照のみになったため、ここから安全に呼べる
}

// =========================================================================
// v2.8.2: High-DPI scaling helpers
// =========================================================================

static UINT g_dpi = 96;
static int SC(int v){ return MulDiv(v, (int)g_dpi, 96); }
static void InitDpi(){
    typedef UINT (WINAPI *PGETDPIFORSYSTEM)();
    HMODULE hUser = GetModuleHandleW(L"user32.dll");
    if(hUser){
        PGETDPIFORSYSTEM p = (PGETDPIFORSYSTEM)GetProcAddress(hUser, "GetDpiForSystem");
        if(p){ UINT d = p(); if(d >= 96){ g_dpi = d; return; } }
    }
    HDC dc = GetDC(NULL);
    if(dc){ int d = GetDeviceCaps(dc, LOGPIXELSX); ReleaseDC(NULL, dc); if(d >= 96) g_dpi = (UINT)d; }
}

// =========================================================================
// RegisterPlugin
// =========================================================================

extern "C" __declspec(dllexport) void __cdecl RegisterPlugin(HOST_APP_TABLE* host){
    host->set_plugin_information(L"Whisper Subtitle v2.9.0");
    InitCommonControls();
    InitDpi();
    EnsureDirectories(); EnsurePyHelper();

    WNDCLASSEXW wc = {sizeof(wc)};
    wc.lpszClassName = L"WhisperSub27";
    wc.lpfnWndProc = WndProc;
    wc.hInstance = g_hInst;
    wc.hbrBackground = GetSysColorBrush(COLOR_BTNFACE);
    // v2.9.1: hCursor未設定(NULL)だとWindowsはこのウィンドウ上でカーソルを設定しないため、
    // 直前のカーソル(ウィンドウ境目のサイズ変更用 ↔ 等)がそのまま残る不具合があった。
    // PresetBrowserで同現象が起きないのは、あちらがhCursorを設定しているため。
    wc.hCursor = LoadCursorW(NULL, IDC_ARROW);
    RegisterClassExW(&wc);

    g_wnd = CreateWindowExW(0, L"WhisperSub27", L"Whisper", WS_POPUP, 0, 0, SC(360), SC(400), 0, 0, g_hInst, 0);
    int W = 340;
    HWND hw;

    // === Tab control (4 tabs) ===
    g_tab = CreateWindowExW(0, WC_TABCONTROLW, L"", WS_VISIBLE|WS_CHILD|WS_CLIPSIBLINGS, SC(4), SC(2), SC(W+6), SC(360), g_wnd, 0, g_hInst, 0);
    TCITEMW ti = {TCIF_TEXT};
    ti.pszText = (LPWSTR)L"\x751f\x6210";
    TabCtrl_InsertItem(g_tab, 0, &ti);
    ti.pszText = (LPWSTR)L"\x8a2d\x5b9a";
    TabCtrl_InsertItem(g_tab, 1, &ti);
    ti.pszText = (LPWSTR)L"\x7cbe\x5ea6";
    TabCtrl_InsertItem(g_tab, 2, &ti);
    ti.pszText = (LPWSTR)L"\x74b0\x5883";
    TabCtrl_InsertItem(g_tab, 3, &ti);
    int tabY = 28;

    // ================================================================
    // Tab 0: 生成 (Generate) — compact, button near top
    // ================================================================
    int y = tabY + 8;
    hw = CreateWindowExW(0, L"STATIC", L"Model:", WS_CHILD|WS_VISIBLE, SC(14), SC(y+3), SC(42), SC(18), g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    g_modelCombo = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD|WS_VISIBLE|CBS_DROPDOWNLIST|WS_VSCROLL, SC(58), SC(y), SC(146), SC(120), g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_modelCombo);
    SendMessageW(g_modelCombo, CB_ADDSTRING, 0, (LPARAM)L"tiny");
    SendMessageW(g_modelCombo, CB_ADDSTRING, 0, (LPARAM)L"base");
    SendMessageW(g_modelCombo, CB_ADDSTRING, 0, (LPARAM)L"small");
    SendMessageW(g_modelCombo, CB_ADDSTRING, 0, (LPARAM)L"medium");
    SendMessageW(g_modelCombo, CB_ADDSTRING, 0, (LPARAM)L"large-v3");
    SendMessageW(g_modelCombo, CB_ADDSTRING, 0, (LPARAM)L"large-v3-turbo");
    SendMessageW(g_modelCombo, CB_ADDSTRING, 0, (LPARAM)L"kotoba-whisper (\x65e5\x82f1)");
    SendMessageA(g_modelCombo, CB_SETCURSEL, 5, 0);
    hw = CreateWindowExW(0, L"STATIC", L"\x8a00\x8a9e:", WS_CHILD|WS_VISIBLE, SC(212), SC(y+3), SC(36), SC(18), g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    g_langCombo = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD|WS_VISIBLE|CBS_DROPDOWNLIST|WS_VSCROLL, SC(250), SC(y), SC(90), SC(200), g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_langCombo);
    SendMessageW(g_langCombo, CB_ADDSTRING, 0, (LPARAM)L"\x81ea\x52d5");
    SendMessageW(g_langCombo, CB_ADDSTRING, 0, (LPARAM)L"ja");
    SendMessageW(g_langCombo, CB_ADDSTRING, 0, (LPARAM)L"en");
    SendMessageW(g_langCombo, CB_ADDSTRING, 0, (LPARAM)L"zh");
    SendMessageW(g_langCombo, CB_ADDSTRING, 0, (LPARAM)L"ko");
    SendMessageA(g_langCombo, CB_SETCURSEL, 1, 0);
    y += 30;
    hw = CreateWindowExW(0, L"BUTTON", L"\x5b57\x5e55\x751f\x6210", WS_CHILD|WS_VISIBLE|BS_DEFPUSHBUTTON, SC(14), SC(y), SC((W-24)/2 - 3), SC(36), g_wnd, (HMENU)IDC_GENERATE, g_hInst, 0); g_tabSubCtrls.push_back(hw); g_btnGenerate = hw;
    hw = CreateWindowExW(0, L"BUTTON", L"SRT\x30a8\x30af\x30b9\x30dd\x30fc\x30c8", WS_CHILD|WS_VISIBLE, SC(14 + (W-24)/2 + 6), SC(y), SC((W-24)/2 - 3), SC(36), g_wnd, (HMENU)IDC_EXPORT_SRT, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    y += 42;
    g_progress = CreateWindowExW(0, PROGRESS_CLASSW, L"", WS_CHILD|WS_VISIBLE, SC(14), SC(y), SC(W-24), SC(14), g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_progress);
    SendMessageA(g_progress, PBM_SETRANGE, 0, MAKELPARAM(0, 100));
    y += 18;
    g_status = CreateWindowExW(0, L"STATIC", L"Ready (v2.9.0)", WS_CHILD|WS_VISIBLE, SC(14), SC(y), SC(W-24), SC(20), g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_status);

    // ================================================================
    // Tab 1: 設定 (Settings) — Backend, Beam/Temp, Device, Layer, etc.
    // ================================================================
    y = tabY + 8;
    hw = CreateWindowExW(0, L"STATIC", L"Backend:", WS_CHILD, SC(14), SC(y+3), SC(52), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    g_backendCombo = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD|CBS_DROPDOWNLIST, SC(70), SC(y), SC(130), SC(80), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_backendCombo);
    SendMessageW(g_backendCombo, CB_ADDSTRING, 0, (LPARAM)L"faster-whisper");
    SendMessageW(g_backendCombo, CB_ADDSTRING, 0, (LPARAM)L"whisper");
    // v2.9.6【配布】既定を faster-whisper(0) → whisper(1) に変更。
    // 理由: ①日本語は openai-whisper + large-v3-turbo が実測で最良(memory whisper-backend-doctrine)
    //       ②faster-whisper は CUDA 実行に nvidia の cuBLAS/cuDNN DLL を必要とし、新規環境では
    //         それが無くて転写の途中で落ちる(モデル構築は成功するのでCPUフォールバックも効かない)。
    // 既存ユーザーは ini の backend= が読み込まれるのでこの既定値は影響しない。
    SendMessageA(g_backendCombo, CB_SETCURSEL, 1, 0);
    hw = CreateWindowExW(0, L"STATIC", L"Device:", WS_CHILD, SC(210), SC(y+3), SC(44), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    g_deviceCombo = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD|CBS_DROPDOWNLIST, SC(256), SC(y), SC(84), SC(80), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_deviceCombo);
    SendMessageW(g_deviceCombo, CB_ADDSTRING, 0, (LPARAM)L"\x81ea\x52d5");
    SendMessageW(g_deviceCombo, CB_ADDSTRING, 0, (LPARAM)L"CPU");
    SendMessageW(g_deviceCombo, CB_ADDSTRING, 0, (LPARAM)L"CUDA");
    // v2.9.6【配布】既定を CUDA(2) → 自動(0) に変更。GPUの有無を決め打ちしない
    // (Python側 525行で device=="auto" を実環境から判定している)。
    SendMessageA(g_deviceCombo, CB_SETCURSEL, 0, 0);
    y += 26;
    hw = CreateWindowExW(0, L"STATIC", L"Beam:", WS_CHILD, SC(14), SC(y+3), SC(38), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    g_qualityEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"5", WS_CHILD|ES_NUMBER|ES_CENTER, SC(52), SC(y), SC(30), SC(22), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_qualityEdit);
    hw = CreateWindowExW(0, L"STATIC", L"(1-10)", WS_CHILD, SC(85), SC(y+3), SC(42), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    hw = CreateWindowExW(0, L"STATIC", L"Temp:", WS_CHILD, SC(135), SC(y+3), SC(38), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    g_tempEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"0", WS_CHILD|ES_CENTER, SC(175), SC(y), SC(30), SC(22), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_tempEdit);
    hw = CreateWindowExW(0, L"STATIC", L"(0=\x56fa\x5b9a)", WS_CHILD, SC(208), SC(y+3), SC(80), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    y += 26;
    hw = CreateWindowExW(0, L"STATIC", L"Layer:", WS_CHILD, SC(14), SC(y+3), SC(40), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    g_layerEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"2", WS_CHILD|ES_NUMBER|ES_CENTER, SC(56), SC(y), SC(36), SC(22), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_layerEdit);
    hw = CreateWindowExW(0, L"STATIC", L"(\x5b57\x5e55\x914d\x7f6e\x5148)", WS_CHILD, SC(96), SC(y+3), SC(80), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    hw = CreateWindowExW(0, L"STATIC", L"\x6587\x5b57\x6570:", WS_CHILD, SC(185), SC(y+3), SC(48), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    g_maxCharEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"0", WS_CHILD|ES_NUMBER|ES_CENTER, SC(235), SC(y), SC(36), SC(22), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_maxCharEdit);
    hw = CreateWindowExW(0, L"STATIC", L"(0=\x7121\x5236\x9650)", WS_CHILD, SC(275), SC(y+3), SC(65), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    y += 28;
    hw = CreateWindowExW(0, L"STATIC", L"\x66f8\x5f0f:", WS_CHILD, SC(14), SC(y+3), SC(42), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    g_templateLabel = CreateWindowExW(0, L"STATIC", L"\x30c7\x30d5\x30a9\x30eb\x30c8", WS_CHILD, SC(58), SC(y+3), SC(110), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_templateLabel);
    hw = CreateWindowExW(0, L"BUTTON", L"\x9078\x629e", WS_CHILD, SC(175), SC(y), SC(55), SC(22), g_wnd, (HMENU)IDC_TEMPLATE, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    hw = CreateWindowExW(0, L"BUTTON", L"\x30ea\x30bb\x30c3\x30c8", WS_CHILD, SC(W-75), SC(y), SC(65), SC(22), g_wnd, (HMENU)IDC_RESET_TPL, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    y += 28;
    hw = CreateWindowExW(0, L"STATIC", L"\x5b57\x5e55\x5ef6\x9577:", WS_CHILD, SC(14), SC(y+3), SC(62), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    g_lingerEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"1.0", WS_CHILD|ES_CENTER, SC(78), SC(y), SC(40), SC(22), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_lingerEdit);
    hw = CreateWindowExW(0, L"STATIC", L"\x79d2 (0=\x306a\x3057)", WS_CHILD, SC(122), SC(y+3), SC(100), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    y += 28;
    hw = CreateWindowExW(0, L"STATIC", L"\x2500\x2500 \x30c6\x30ad\x30b9\x30c8\x51e6\x7406 \x2500\x2500", WS_CHILD, SC(14), SC(y), SC(W-24), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    y += 22;
    g_chkRemovePunct = CreateWindowExW(0, L"BUTTON", L"\x53e5\x8aad\x70b9\x524a\x9664", WS_CHILD|BS_AUTOCHECKBOX, SC(14), SC(y), SC(100), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_chkRemovePunct);
    g_chkRemoveExclam = CreateWindowExW(0, L"BUTTON", L"!?\x524a\x9664", WS_CHILD|BS_AUTOCHECKBOX, SC(116), SC(y), SC(72), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_chkRemoveExclam);
    g_chkNormalize = CreateWindowExW(0, L"BUTTON", L"\x5168\x534a\x89d2\x6b63\x898f\x5316", WS_CHILD|BS_AUTOCHECKBOX, SC(192), SC(y), SC(120), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_chkNormalize);
    // v2.8.2: セグメント結合 (短い字幕を maxchars 目標までまとめる。default OFF)
    y += 22;
    g_chkMergeSeg = CreateWindowExW(0, L"BUTTON", L"\x77ed\x3044\x5b57\x5e55\x3092\x7d50\x5408", WS_CHILD|BS_AUTOCHECKBOX, SC(14), SC(y), SC(W-28), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_chkMergeSeg); // v2.8.9: F-2 括弧書き廃止でラベル短縮
    // v2.8.2: 形態素解析によるプロ品質の文節分割 (fugashi, default OFF)
    y += 22;
    g_chkMorphSplit = CreateWindowExW(0, L"BUTTON", L"\x6587\x7bc0\x533a\x5207\x308a(fugashi)", WS_CHILD|BS_AUTOCHECKBOX, SC(14), SC(y), SC(W-28), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_chkMorphSplit); // v2.8.9: F-1「プロ分割」→「文節区切り」用語統一
    // v2.8.5: 2行まで許容 (1テロップを最大2行=maxchars×2までにまとめ、文節境界で改行)
    y += 22;
    g_chkTwoLine = CreateWindowExW(0, L"BUTTON", L"\x6700\x5927" L"2\x884c\x306b\x307e\x3068\x3081\x308b", WS_CHILD|BS_AUTOCHECKBOX, SC(14), SC(y), SC(W-28), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_chkTwoLine); // v2.8.9: F-2 括弧書き廃止でラベル短縮(\x5927直後の"2"はリテラル分割で\x59272化を回避)
    // v2.8.2a: 先行表示(秒) - 字幕を音声より早く出す (0=無効)
    y += 22;
    hw = CreateWindowExW(0, L"STATIC", L"\x5148\x884c\x8868\x793a:", WS_CHILD, SC(14), SC(y+3), SC(62), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);
    g_leadEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"0.0", WS_CHILD|ES_CENTER, SC(78), SC(y), SC(40), SC(22), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(g_leadEdit);
    hw = CreateWindowExW(0, L"STATIC", L"\x79d2 (0=\x7121\x52b9)", WS_CHILD, SC(122), SC(y+3), SC(100), SC(18), g_wnd, 0, g_hInst, 0); g_tabSettingsCtrls.push_back(hw);

    // ================================================================
    // Tab 2: 精度 (Accuracy / Speed)
    // ================================================================
    y = tabY + 8;
    g_chkNoPrevText = CreateWindowExW(0, L"BUTTON", L"\x5e7b\x8074\x9632\x6b62", WS_CHILD|BS_AUTOCHECKBOX, SC(14), SC(y), SC(82), SC(18), g_wnd, 0, g_hInst, 0); g_tabAccCtrls.push_back(g_chkNoPrevText);
    SendMessageA(g_chkNoPrevText, BM_SETCHECK, BST_CHECKED, 0);
    g_chkRepPenalty = CreateWindowExW(0, L"BUTTON", L"\x7e70\x8fd4\x3057\x6291\x5236", WS_CHILD|BS_AUTOCHECKBOX, SC(100), SC(y), SC(90), SC(18), g_wnd, 0, g_hInst, 0); g_tabAccCtrls.push_back(g_chkRepPenalty);
    SendMessageA(g_chkRepPenalty, BM_SETCHECK, BST_CHECKED, 0);
    g_chkWordTs = CreateWindowExW(0, L"BUTTON", L"\x5358\x8a9eTS", WS_CHILD|BS_AUTOCHECKBOX, SC(196), SC(y), SC(66), SC(18), g_wnd, 0, g_hInst, 0); g_tabAccCtrls.push_back(g_chkWordTs);
    g_chkBatched = CreateWindowExW(0, L"BUTTON", L"Batched", WS_CHILD|BS_AUTOCHECKBOX, SC(268), SC(y), SC(70), SC(18), g_wnd, 0, g_hInst, 0); g_tabAccCtrls.push_back(g_chkBatched);
    y += 22;
    // v2.9.5: VAD(無音カット)。1行目は x=14..338 で埋まっているので2行目に置く。
    // 既定OFF: 実測でONだと単語が40個消え、開始時刻が一貫して早い側にズレた(vad_OFFはopenai-whisperと完全一致)。
    // v2.9.6: 幅110では「VAD無音カッ」で切れた(BS_AUTOCHECKBOXはチェック枠が幅を食うため
    // 文字分だけでは足りない)。2行目はx=14..338が丸々空いているので150まで広げる。
    g_chkVad = CreateWindowExW(0, L"BUTTON", L"VAD\x7121\x97f3\x30ab\x30c3\x30c8", WS_CHILD|BS_AUTOCHECKBOX, SC(14), SC(y), SC(150), SC(18), g_wnd, 0, g_hInst, 0); g_tabAccCtrls.push_back(g_chkVad);
    y += 24;
    hw = CreateWindowExW(0, L"STATIC", L"\x30d2\x30f3\x30c8\x6587: (\x56fa\x6709\x540d\x8a5e\x3084\x5185\x5bb9\x3092\x5165\x529b)", WS_CHILD, SC(14), SC(y), SC(W-24), SC(18), g_wnd, 0, g_hInst, 0); g_tabAccCtrls.push_back(hw);
    y += 20;
    g_promptEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD|ES_MULTILINE|ES_AUTOVSCROLL|WS_VSCROLL, SC(14), SC(y), SC(W-24), SC(50), g_wnd, 0, g_hInst, 0); g_tabAccCtrls.push_back(g_promptEdit);
    y += 54;
    hw = CreateWindowExW(0, L"STATIC", L"\x30db\x30c3\x30c8\x30ef\x30fc\x30c9: (\x30ab\x30f3\x30de\x533a\x5207\x308a)", WS_CHILD, SC(14), SC(y), SC(W-24), SC(18), g_wnd, 0, g_hInst, 0); g_tabAccCtrls.push_back(hw);
    y += 20;
    g_hotwordsEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD|ES_AUTOHSCROLL, SC(14), SC(y), SC(W-24), SC(22), g_wnd, 0, g_hInst, 0); g_tabAccCtrls.push_back(g_hotwordsEdit);

    // ================================================================
    // Tab 3: 環境 (Setup) — v2.9.0【C】: セットアップのチェック形式化に伴い全面再配置。
    // PyTorch行/モデル行を新設(★新規)、ffmpeg自動DLボタンを独立大ボタンから行内[DL]ボタンへ統合、
    // 旧・全項目まとめて入れ直すチェックボックス(1個)を項目別チェック(torch/whisper/faster-whisper/fugashi/モデル)
    // 5個に置き換え。fugashi単独[再導入]ボタンは廃止(チェック方式に統合)。
    // y座標は仕様書§4の表のとおり足し算し、最下端が350(タブ下端362に収まる)になるよう再計算済み:
    //   22(見出し)+26(Python)+26(ffmpeg)+22(見出し)+24(torch)+24(whisper)+24(faster)+24(fugashi)
    //   +24(model)+22(見出し)+38(setup)+18(progress)+20(status) = 314、tabY+8=36 起点なので 36+314=350。
    // ================================================================
    y = tabY + 8; // y=36
    // ── 外部ツール ──
    hw = CreateWindowExW(0, L"STATIC", L"\x2500\x2500" L" " L"\x5916\x90e8\x30c4\x30fc\x30eb" L" " L"\x2500\x2500", WS_CHILD, SC(14), SC(y), SC(W-24), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    y += 22; // y=58
    hw = CreateWindowExW(0, L"STATIC", L"Python:", WS_CHILD, SC(14), SC(y+3), SC(46), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    g_pythonStatusLabel = CreateWindowExW(0, L"STATIC", L"", WS_CHILD, SC(62), SC(y+3), SC(70), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_pythonStatusLabel);
    g_pythonLabel = CreateWindowExW(0, L"STATIC", L"(\x81ea\x52d5\x691c\x51fa)", WS_CHILD|SS_PATHELLIPSIS, SC(136), SC(y+3), SC(140), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_pythonLabel);
    hw = CreateWindowExW(0, L"BUTTON", L"\x9078\x629e", WS_CHILD, SC(W-55), SC(y), SC(50), SC(22), g_wnd, (HMENU)IDC_PYTHON_BR, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    y += 26; // y=84
    // v2.9.4: ffmpegは「プラグインが導入するもの」なので、専用ボタンをやめてチェック一覧側へ移した。
    // 「外部ツール」欄に残るのはPythonだけ = ユーザーが自分で用意する前提のもの、という区分になる。
    // ── 導入するもの ──
    hw = CreateWindowExW(0, L"STATIC", L"\x2500\x2500" L" " L"\x5c0e\x5165\x3059\x308b\x3082\x306e" L" " L"\x2500\x2500", WS_CHILD, SC(14), SC(y), SC(W-24), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    y += 22; // y=106
    // v2.9.4★: ffmpeg行。専用[DL]/[再取得]ボタンは廃止し、チェック+セットアップで取り直す方式に統一。
    g_chkFfmpeg = CreateWindowExW(0, L"BUTTON", L"", WS_CHILD|BS_AUTOCHECKBOX, SC(14), SC(y+2), SC(16), SC(16), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_chkFfmpeg);
    hw = CreateWindowExW(0, L"STATIC", L"ffmpeg:", WS_CHILD, SC(34), SC(y+3), SC(46), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    g_ffmpegStatusLabel = CreateWindowExW(0, L"STATIC", L"", WS_CHILD, SC(82), SC(y+3), SC(70), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_ffmpegStatusLabel); // v2.8.8【E】: ●導入済み/○未導入
    // v2.9.4: DLボタンが消えた分パスの幅を広げた(156..280)。RefreshSetupTabStateのMoveWindowによる幅切替も不要になった
    g_ffmpegLabel = CreateWindowExW(0, L"STATIC", L"(aviutl2.exe\x306e\x5834\x6240)", WS_CHILD|SS_PATHELLIPSIS, SC(156), SC(y+3), SC(124), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_ffmpegLabel);
    hw = CreateWindowExW(0, L"BUTTON", L"\x9078\x629e", WS_CHILD, SC(W-55), SC(y), SC(50), SC(22), g_wnd, (HMENU)IDC_FFMPEG_BR, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    y += 26; // y=132
    // v2.9.0【H】★新規: PyTorch行。チェックすると入れ直し対象になる(SetupThread側でg_chkTorch参照)
    g_chkTorch = CreateWindowExW(0, L"BUTTON", L"", WS_CHILD|BS_AUTOCHECKBOX, SC(14), SC(y+2), SC(16), SC(16), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_chkTorch);
    g_torchStatusLabel = CreateWindowExW(0, L"STATIC", L"", WS_CHILD, SC(34), SC(y+3), SC(W-44), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_torchStatusLabel);
    y += 24; // y=156
    // v2.9.0【D】: whisper行にチェック追加。g_owLocLabelはUpdateWhisperLocLabels()が●/○込みで書き込む(既存のまま)
    g_chkWhisper = CreateWindowExW(0, L"BUTTON", L"", WS_CHILD|BS_AUTOCHECKBOX, SC(14), SC(y+2), SC(16), SC(16), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_chkWhisper);
    // v2.9.2: 名前と状態を別ラベルに分離。旧実装は「●導入済み whisper: <パス>」を1つのSTATICに入れ
    // SS_PATHELLIPSIS(=末尾を優先して残し真ん中を...で潰す)を掛けていたため、幅196pxでは末尾の
    // "\site-packages" だけが残り、whisperとfaster-whisperが同じ表示に見えていた(両者は同じ
    // site-packagesに入るためパスは判別材料にならない)。名前ラベルは省略無しで固定表示する。
    hw = CreateWindowExW(0, L"STATIC", L"whisper:", WS_CHILD, SC(34), SC(y+3), SC(96), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    g_owLocLabel = CreateWindowExW(0, L"STATIC", L"", WS_CHILD, SC(132), SC(y+3), SC(98), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_owLocLabel);
    hw = CreateWindowExW(0, L"BUTTON", L"\x9078\x629e", WS_CHILD, SC(W-110), SC(y), SC(50), SC(22), g_wnd, (HMENU)IDC_OW_BR, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    hw = CreateWindowExW(0, L"BUTTON", L"\x81ea\x52d5", WS_CHILD, SC(W-55), SC(y), SC(50), SC(22), g_wnd, (HMENU)IDC_OW_RESET, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    y += 24; // y=180
    g_chkFaster = CreateWindowExW(0, L"BUTTON", L"", WS_CHILD|BS_AUTOCHECKBOX, SC(14), SC(y+2), SC(16), SC(16), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_chkFaster);
    hw = CreateWindowExW(0, L"STATIC", L"faster-whisper:", WS_CHILD, SC(34), SC(y+3), SC(96), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(hw); // v2.9.2: 名前は固定表示(省略しない)
    g_fwLocLabel = CreateWindowExW(0, L"STATIC", L"", WS_CHILD, SC(132), SC(y+3), SC(98), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_fwLocLabel);
    hw = CreateWindowExW(0, L"BUTTON", L"\x9078\x629e", WS_CHILD, SC(W-110), SC(y), SC(50), SC(22), g_wnd, (HMENU)IDC_FW_BR, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    hw = CreateWindowExW(0, L"BUTTON", L"\x81ea\x52d5", WS_CHILD, SC(W-55), SC(y), SC(50), SC(22), g_wnd, (HMENU)IDC_FW_RESET, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    y += 24; // y=204
    // v2.9.0【F】: fugashi単独[再導入]ボタン(旧専用ID)は廃止し、チェック方式(g_chkFugashi)に統合。
    // 状態表示(g_fugashiStatus)はUpdateFugashiStatus()が「文節区切り(fugashi): ●導入済み/○未導入」を書き込む。
    g_chkFugashi = CreateWindowExW(0, L"BUTTON", L"", WS_CHILD|BS_AUTOCHECKBOX, SC(14), SC(y+2), SC(16), SC(16), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_chkFugashi);
    g_fugashiStatus = CreateWindowExW(0, L"STATIC", L"", WS_CHILD, SC(34), SC(y+3), SC(W-44), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_fugashiStatus);
    y += 24; // y=228
    // v2.9.0【G】★新規: モデル行
    // v2.9.1: チェックボックスは廃止。使うモデルはユーザーごとに違い(配布時にlarge-v3-turbo固定ではない)、
    // かつモデルは単なるファイルなのでtorchのような「importは通るが中身が壊れている」状態が起きない。
    // 強制再導入(1.5GB超の再DL)の価値が薄いためチェックを外し、状態表示のみ残す。
    // ラベルのx=34は他のライブラリ行(チェックボックス分インデント)と左端を揃えるため据え置き。
    g_modelStatusLabel = CreateWindowExW(0, L"STATIC", L"", WS_CHILD, SC(34), SC(y+3), SC(W-44), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_modelStatusLabel);
    y += 24; // y=252
    // ── 導入 ──
    hw = CreateWindowExW(0, L"STATIC", L"\x2500\x2500" L" " L"\x5c0e\x5165" L" " L"\x2500\x2500", WS_CHILD, SC(14), SC(y), SC(W-24), SC(18), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    y += 22; // y=274
    // v2.9.0【J】: ラベルを「セットアップ (不足分をまとめて導入)」→「セットアップ (チェック項目は再導入)」に変更
    hw = CreateWindowExW(0, L"BUTTON", L"\x30bb\x30c3\x30c8\x30a2\x30c3\x30d7" L" (" L"\x30c1\x30a7\x30c3\x30af\x9805\x76ee\x306f\x518d\x5c0e\x5165" L")", WS_CHILD, SC(14), SC(y), SC(W-24), SC(34), g_wnd, (HMENU)IDC_SETUP, g_hInst, 0); g_tabSetupCtrls.push_back(hw); g_btnSetup = hw;
    y += 38; // y=312
    // v2.8.8【A】: 環境タブ専用の進捗バー/ステータス。従来g_progress/g_statusはg_tabSubCtrls側にありSW_HIDEで非表示だった
    g_progressSetup = CreateWindowExW(0, PROGRESS_CLASSW, L"", WS_CHILD, SC(14), SC(y), SC(W-24), SC(14), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_progressSetup);
    SendMessageA(g_progressSetup, PBM_SETRANGE, 0, MAKELPARAM(0, 100));
    y += 18; // y=330
    g_statusSetup = CreateWindowExW(0, L"STATIC", L"Ready (v2.9.0)", WS_CHILD, SC(14), SC(y), SC(W-24), SC(20), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_statusSetup);
    // y+20 = 350 = 最下端(タブ下端362に収まる)

    SwitchTab(0);
    host->register_window_client(L"Whisper Subtitle", g_wnd);
    g_edit = host->create_edit_handle();
    LoadSettings();
    std::thread(ProbeThread).detach(); // v2.9.0【A】: 起動時に1回だけpythonを回して導入状況を一括実測(表示専用、UIは固まらない)
    EnableWindow(g_chkMergeSeg, (SendMessageA(g_chkMorphSplit, BM_GETCHECK, 0, 0) == BST_CHECKED) ? FALSE : TRUE); // v2.8.5
    UpdateVadEnable(); // v2.9.5: 起動時にも反映

    // Auto-detect ffmpeg on first run (if not already set)
    if(g_ffmpegPath.empty()){
        std::string def = GetExeDir() + "\\ffmpeg.exe";
        if(FileExistsU(def)){
            g_ffmpegPath = def;
            SaveSettings();
            DebugLog("Auto-detected ffmpeg: " + def);
        }
    }

    SetPathLabel(g_ffmpegLabel, g_ffmpegPath, "(aviutl2.exe\xe3\x81\xae\xe5\xa0\xb4\xe6\x89\x80)");
    SetPathLabel(g_pythonLabel, g_pythonPath, "(\xe8\x87\xaa\xe5\x8b\x95\xe6\xa4\x9c\xe5\x87\xba)");
    UpdateWhisperLocLabels();
    UpdateFugashiStatus(); // v2.8.5
    if(!g_templatePath.empty()){
        if(LoadTemplate(g_templatePath)) UpdateTemplateLabel();
        else g_templatePath.clear();
    }
}

BOOL APIENTRY DllMain(HINSTANCE h, DWORD r, LPVOID){
    if(r == DLL_PROCESS_ATTACH) g_hInst = h;
    return TRUE;
}
'@

# ===================================
# Write source file
# ===================================

Write-Host "Step 1: Creating project directory..." -ForegroundColor Yellow
New-Item -Path $src -ItemType Directory -Force | Out-Null
$cppPath = "$src\whisper_subtitle.cpp"
[System.IO.File]::WriteAllText($cppPath, $cpp, [System.Text.Encoding]::UTF8)
"Source: $cppPath" | Out-File $logFile -Append -Encoding UTF8
Write-Host "  Source: $cppPath" -ForegroundColor White

# ===================================
# Count braces
# ===================================
$openBraces = ([regex]::Matches($cpp, '\{')).Count
$closeBraces = ([regex]::Matches($cpp, '\}')).Count
Write-Host "  Brace check: open=$openBraces close=$closeBraces" -ForegroundColor $(if($openBraces -eq $closeBraces){"Green"}else{"Red"})
"Braces: open=$openBraces close=$closeBraces" | Out-File $logFile -Append -Encoding UTF8

# ===================================
# Generate CMakeLists.txt
# ===================================

Write-Host "Step 2: Generating CMakeLists.txt..." -ForegroundColor Yellow
$cmake = @"
cmake_minimum_required(VERSION 3.15)
project(whisper_subtitle_v2_9_0 LANGUAGES CXX)
set(CMAKE_CXX_STANDARD 17)
add_library(whisper_subtitle_v2_9_0 SHARED src/whisper_subtitle.cpp)
target_compile_definitions(whisper_subtitle_v2_9_0 PRIVATE UNICODE _UNICODE)
if(MSVC)
    target_compile_options(whisper_subtitle_v2_9_0 PRIVATE /utf-8 /wd4828)
endif()
# v2.9.6【配布】CRT を静的リンク(/MT)にする。
# 従来は CMake 既定の /MD で、配布先に VCRUNTIME140.dll / VCRUNTIME140_1.dll / MSVCP140.dll
# (Visual C++ 再頒布可能パッケージ) が無いと**起動すらしない**状態だった。
# 同プロジェクトの PresetBrowser / BatchEdit / WaveAmp / WaveProbe は全て cl.exe に /MT を
# 明示しており AviUtl2 上で本番稼働中 = このホストで静的CRTが安全なことは実証済み。
# Whisper だけ /MD だったのは CMake 既定のまま取り残されていただけで、設計判断ではない。
# CMP0091 は cmake_minimum_required(3.15) により NEW になるのでこのプロパティが効く。
set_property(TARGET whisper_subtitle_v2_9_0 PROPERTY MSVC_RUNTIME_LIBRARY "MultiThreaded")
set_target_properties(whisper_subtitle_v2_9_0 PROPERTIES
    OUTPUT_NAME "whisper_subtitle_v2_9_0"
    SUFFIX ".aux2"
    PREFIX ""
    RUNTIME_OUTPUT_DIRECTORY_RELEASE "`${CMAKE_BINARY_DIR}/Release"
)
target_link_libraries(whisper_subtitle_v2_9_0 PRIVATE comctl32 shell32 comdlg32 ole32)
"@
[System.IO.File]::WriteAllText("$projDir\CMakeLists.txt", $cmake, [System.Text.Encoding]::UTF8)
Write-Host "  CMakeLists.txt written" -ForegroundColor White

# ===================================
# Build
# ===================================

Write-Host "Step 3: Running CMake..." -ForegroundColor Yellow
$buildDir = "$projDir\build"
New-Item -Path $buildDir -ItemType Directory -Force | Out-Null

$cmakeExe = "cmake"
$cmakeGen = "Visual Studio 17 2022"

Push-Location $buildDir
try {
    $cmakeResult = & $cmakeExe -G $cmakeGen -A x64 .. 2>&1
    $cmakeResult | Out-String | Out-File $logFile -Append -Encoding UTF8
    if($LASTEXITCODE -ne 0){
        Write-Host "  CMake configure FAILED" -ForegroundColor Red
        $cmakeResult | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkRed }
        throw "CMake configure failed"
    }
    Write-Host "  CMake configure OK" -ForegroundColor Green

    Write-Host "Step 4: Building..." -ForegroundColor Yellow
    $buildResult = & $cmakeExe --build . --config Release 2>&1
    $buildResult | Out-String | Out-File $logFile -Append -Encoding UTF8
    if($LASTEXITCODE -ne 0){
        Write-Host "  Build FAILED" -ForegroundColor Red
        $buildResult | ForEach-Object { Write-Host "    $_" -ForegroundColor DarkRed }
        throw "Build failed"
    }
    Write-Host "  Build OK" -ForegroundColor Green

    # Find output
    $aux2 = Get-ChildItem -Path $buildDir -Recurse -Filter "whisper_subtitle_v2_9_0.aux2" | Select-Object -First 1
    if($aux2){
        # v2.8b: always deploy next to this script (..\sdk whisper開発場所\), NOT to $d/Desktop,
        # and NEVER to the production AviUtl2Portable folder. Filename is version-specific so the
        # existing whisper_subtitle_v2_8_9.aux2 / whisper_subtitle.aux2 (v2.8 stable) are never touched/overwritten.
        $destDir = $PSScriptRoot
        if([string]::IsNullOrEmpty($destDir)){ $destDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
        $dest = Join-Path $destDir "whisper_subtitle_v2_9_0.aux2"
        Copy-Item $aux2.FullName $dest -Force
        Write-Host ""
        Write-Host "============================================" -ForegroundColor Green
        Write-Host " BUILD SUCCESS! v2.8.2" -ForegroundColor Green
        Write-Host "============================================" -ForegroundColor Green
        Write-Host "Output: $dest" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "v2.8.2 Changes:" -ForegroundColor Cyan
        Write-Host "  + High-DPI UI scaling (125%/150% display scale) [NEW in v2.8.2]" -ForegroundColor White
        Write-Host "  + (v2.8b) RunProcess: timeoutMs now enforced (TerminateProcess on timeout)" -ForegroundColor White
        Write-Host "  + (v2.8b) SplitText: fixed UTF-8 boundary infinite loop" -ForegroundColor White
        Write-Host "  + (v2.8b) ExtractAudio: std::wstring command build (no fixed buffer)" -ForegroundColor White
        Write-Host "  + (v2.8b) g_busy: atomic<bool> + compare_exchange_strong, buttons disabled while busy" -ForegroundColor White
        Write-Host '  + (v2.8b) JSON generation: double-quote now escaped everywhere' -ForegroundColor White
        "BUILD SUCCESS" | Out-File $logFile -Append -Encoding UTF8
    } else {
        Write-Host "  .aux2 not found in build output" -ForegroundColor Red
        throw "Output file not found"
    }
} catch {
    Write-Host "BUILD FAILED: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Log: $logFile" -ForegroundColor Yellow
    "FAILED: $($_.Exception.Message)" | Out-File $logFile -Append -Encoding UTF8
    Start-Process notepad.exe $logFile
} finally {
    Pop-Location
}
Write-Host ""; Write-Host "Press Enter to exit..." -ForegroundColor Gray; Read-Host
