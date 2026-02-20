##############################################################################
# Whisper Subtitle Plugin for AviUtl2 - v2.5
#
# 機能:
#   - faster-whisper / openai-whisper 対応 (自動インストール)
#   - CUDA GPU加速 + CPU fallback
#   - 書式テンプレート (.object)
#   - 句読点削除、!?削除、全半角正規化
#   - 字幕延長 (発話後の表示維持)
#   - レイヤー自動シフト (既存オブジェクト回避)
#   - SRTエクスポート
#   - モデル: tiny/base/small/medium/large-v3/large-v3-turbo
#
# ビルド: .\whisper_subtitle_v2_5.ps1 [出力ディレクトリ]
# 要件: Visual Studio 2022, CMake 3.15+
#
# Note: PyTorch CUDA版はcu121 (CUDA 12.1)を使用
#       CUDA 11.x環境では --index-url を変更してください
##############################################################################

$d = if($args.Count -gt 0){ $args[0] } else { [Environment]::GetFolderPath("Desktop") }
$projDir = "$d\aviutl2_dev\whisper_subtitle_plugin"
$src = "$projDir\src"
$logFile = "$d\whisper_build_log.txt"
$ErrorActionPreference = "Continue"

"" | Out-File $logFile -Encoding UTF8
"Whisper Subtitle v2.5 Build $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File $logFile -Append -Encoding UTF8

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " Whisper Subtitle v2.5 - Build" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

$cpp = @'
// Whisper Subtitle v2.5
// faster-whisper only, no whisper.h dependency
#include <windows.h>
#include <shlobj.h>
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
};

// =========================================================================
// Globals
// =========================================================================

static HINSTANCE g_hInst = 0;
static HWND g_wnd = 0;
static HWND g_modelCombo = 0, g_deviceCombo = 0, g_backendCombo = 0;
static HWND g_langCombo = 0, g_qualityEdit = 0, g_tempEdit = 0;
static HWND g_chkRemovePunct = 0, g_chkNormalize = 0, g_chkRemoveExclam = 0;
static HWND g_fwLocLabel = 0, g_owLocLabel = 0; // faster-whisper / openai-whisper location labels
static HWND g_layerEdit = 0, g_maxCharEdit = 0;
static HWND g_tab = 0;
static std::vector<HWND> g_tabSubCtrls, g_tabSetupCtrls; // controls per tab

static void SwitchTab(int idx){
    for(auto h : g_tabSubCtrls) ShowWindow(h, idx == 0 ? SW_SHOW : SW_HIDE);
    for(auto h : g_tabSetupCtrls) ShowWindow(h, idx == 1 ? SW_SHOW : SW_HIDE);
}
static HWND g_templateLabel = 0, g_status = 0, g_progress = 0;
static HWND g_ffmpegLabel = 0, g_pythonLabel = 0;
static HWND g_lingerEdit = 0;
static EDIT_HANDLE* g_edit = nullptr;
static bool g_busy = false;
static std::string g_templatePath, g_templateContent;
static std::string g_ffmpegPath, g_pythonPath;
static std::string g_fwSpPath, g_owSpPath; // custom site-packages dirs for faster-whisper / openai-whisper
static int g_projectRate = 30;

#define WM_UPDATE_STATUS (WM_USER + 100)
#define WM_UPDATE_PROGRESS (WM_USER + 101)
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
        _run_openai_whisper(model_size, language, device, clips, model_dir, output_path, err_path, beam_size, temperature)
    else:
        _run_faster_whisper(model_size, language, device, clips, model_dir, output_path, err_path, beam_size, temperature)

def _run_faster_whisper(model_size, language, device, clips, model_dir, output_path, err_path, beam_size=5, temperature=0):
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
    model_path = model_size
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
    results = []
    for ci, clip in enumerate(clips):
        wav_path = clip["wav"]
        tl_start = clip["timeline_start"]
        fps = clip["fps"]
        if not os.path.exists(wav_path):
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Clip {ci}: wav not found: {wav_path}\n")
            continue
        try:
            segments, info = model.transcribe(wav_path, language=language, beam_size=beam_size, vad_filter=True, word_timestamps=True, temperature=temperature)
            for seg in segments:
                sf = tl_start + int(seg.start * fps)
                ef2 = tl_start + int(seg.end * fps)
                text = seg.text.strip()
                if text and ef2 > sf:
                    results.append(f"{ci}|{sf}|{ef2}|{text}")
        except Exception as e:
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Clip {ci} err: {e}\n{traceback.format_exc()}")
    with open(output_path, "w", encoding="utf-8") as f:
        for line in results:
            f.write(line + "\n")
    with open(err_path, "a", encoding="utf-8") as ef:
        ef.write(f"Done: {len(results)} segs ({actual_device})\n")

def _run_openai_whisper(model_size, language, device, clips, model_dir, output_path, err_path, beam_size=5, temperature=0):
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
    results = []
    for ci, clip in enumerate(clips):
        wav_path = clip["wav"]
        tl_start = clip["timeline_start"]
        fps = clip["fps"]
        if not os.path.exists(wav_path):
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Clip {ci}: wav not found: {wav_path}\n")
            continue
        try:
            opts = {"language": language, "beam_size": beam_size, "verbose": False, "word_timestamps": True, "temperature": temperature}
            if language is None:
                del opts["language"]
            result = model.transcribe(wav_path, **opts)
            for seg in result.get("segments", []):
                sf = tl_start + int(seg["start"] * fps)
                ef2 = tl_start + int(seg["end"] * fps)
                text = seg["text"].strip()
                if text and ef2 > sf:
                    results.append(f"{ci}|{sf}|{ef2}|{text}")
        except Exception as e:
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Clip {ci} err: {e}\n{traceback.format_exc()}")
    with open(output_path, "w", encoding="utf-8") as f:
        for line in results:
            f.write(line + "\n")
    with open(err_path, "a", encoding="utf-8") as ef:
        ef.write(f"Done: {len(results)} segs (openai-whisper, {device})\n")

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
        for(int v = 13; v >= 8; v--){
            std::wstring c = std::wstring(up) + L"\\Programs\\Python\\Python3" + std::to_wstring(v) + L"\\python.exe";
            if(GetFileAttributesW(c.c_str()) != INVALID_FILE_ATTRIBUTES) return WideToUtf8(c);
        }
    }
    wchar_t buf[MAX_PATH];
    if(SearchPathW(NULL, L"python.exe", NULL, MAX_PATH, buf, NULL) > 0) return WideToUtf8(buf);
    return "";
}
static std::string GetEffectivePython(){
    if(!g_pythonPath.empty() && FileExistsU(g_pythonPath))
        return g_pythonPath;
    return FindPythonAuto();
}
static std::string GetEffectiveFFmpeg(){
    if(!g_ffmpegPath.empty() && FileExistsU(g_ffmpegPath))
        return g_ffmpegPath;
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
    f << "ffmpeg=" << g_ffmpegPath << "\n";
    f << "python=" << g_pythonPath << "\n";
    f << "fw_sp=" << g_fwSpPath << "\n";
    f << "ow_sp=" << g_owSpPath << "\n";
    char lBuf[16] = {}; GetWindowTextA(g_lingerEdit, lBuf, sizeof(lBuf));
    f << "linger=" << lBuf << "\n";
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
        else if(key=="ffmpeg") g_ffmpegPath = val;
        else if(key=="python") g_pythonPath = val;
        else if(key=="fw_sp") g_fwSpPath = val;
        else if(key=="ow_sp") g_owSpPath = val;
        else if(key=="linger") SetWindowTextA(g_lingerEdit, val.c_str());
    }
}

// =========================================================================
// Setup thread (auto-install faster-whisper + model DL)
// =========================================================================

// Forward declarations (defined later, after timeline code)
static void SetStatus(const std::string& msg);
static void SetProgress(int val);
static void UpdateWhisperLocLabels();

static void SetupThread(){
    g_busy = true;
    SetWindowTextW(g_status, L"\x521d\x671f\x8a2d\x5b9a\x4e2d...");
    SendMessageA(g_progress, PBM_SETPOS, 5, 0);
    std::string report;
    std::string python = GetEffectivePython();
    if(python.empty()){
        MsgBox(g_wnd,
            "Python\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n\n"
            "python.org\xe3\x81\x8b\xe3\x82\x89Python 3.10+\xe3\x82\x92\xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82\n"
            "\xe3\x81\xbe\xe3\x81\x9f\xe3\x81\xaf\xe3\x80\x8cPython\xe9\x81\xb8\xe6\x8a\x9e\xe3\x80\x8d\xe3\x83\x9c\xe3\x82\xbf\xe3\x83\xb3\xe3\x81\xa7\xe6\x8c\x87\xe5\xae\x9a",
            "Python\xe3\x81\x8c\xe5\xbf\x85\xe8\xa6\x81", MB_OK|MB_ICONWARNING);
        SetWindowTextW(g_status, L"Ready (v2.5)");
        SendMessageA(g_progress, PBM_SETPOS, 0, 0);
        g_busy = false;
        return;
    }
    report += "Python: " + python + "\n";

    // Install whisper backend to local site-packages
    int setupBi = SendMessageA(g_backendCombo, CB_GETCURSEL, 0, 0);
    if(setupBi == CB_ERR) setupBi = 0;
    const char* pipPkg = (setupBi == 1) ? "openai-whisper" : "faster-whisper";
    std::string spDir = GetSitePackagesDir();
    // Install backend first
    SetWindowTextW(g_status, Utf8ToWide(std::string(pipPkg) + " \xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe4\xb8\xad...").c_str());
    SendMessageA(g_progress, PBM_SETPOS, 20, 0);
    {
        std::wstring cmd = Utf8ToWide("\"" + python + "\" -m pip install " + pipPkg + " --target=\"" + spDir + "\" --upgrade --quiet");
        std::string out; bool ok = RunProcess(cmd, out, 600000);
        DebugLog("pip install --target:\n" + out);
        report += ok ? (std::string(pipPkg) + ": OK (" + spDir + ")\n") : (std::string(pipPkg) + ": WARN\n" + out + "\n");
    }
    // For openai-whisper, re-install CUDA torch AFTER (pip may have overwritten with CPU version)
    if(setupBi == 1){
        SetStatus("PyTorch (CUDA) \xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe4\xb8\xad... (\xe6\x99\x82\xe9\x96\x93\xe3\x81\x8c\xe3\x81\x8b\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x99)");
        SendMessageA(g_progress, PBM_SETPOS, 40, 0);
        std::wstring torchCmd = Utf8ToWide("\"" + python + "\" -m pip install torch --upgrade --target=\"" + spDir + "\" --index-url https://download.pytorch.org/whl/cu121");
        std::string torchOut; RunProcess(torchCmd, torchOut, 900000);
        DebugLog("Torch CUDA install: " + torchOut.substr(0, 500));
    }

    // Download model
    SendMessageA(g_progress, PBM_SETPOS, 50, 0);
    SetWindowTextW(g_status, L"\x30e2\x30c7\x30eb DL\x4e2d...");
    int mi = SendMessageA(g_modelCombo, CB_GETCURSEL, 0, 0);
    if(mi == CB_ERR) mi = 0;
    const char* mn[] = {"tiny","base","small","medium","large-v3","large-v3-turbo"};
    if(mi < 0 || mi >= (int)(sizeof(mn)/sizeof(mn[0]))) mi = 0;
    std::string mName = mn[mi], mDir = GetModelsDir();

    // Check if model already exists locally (skip expensive download+load)
    std::string localModel = mDir + "\\" + mName;
    bool modelExists = FileExistsU(localModel + "\\config.json") // faster-whisper
        || FileExistsU(localModel + "\\model.bin")               // faster-whisper alt
        || FileExistsU(mDir + "\\" + mName + ".pt");              // openai-whisper
    if(modelExists){
        DebugLog("Model already exists: " + localModel);
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
        sf << "        print(f'Downloading {model_name} (faster-whisper) to {model_dir}')\n";
        sf << "        WhisperModel(model_name, device='cpu', compute_type='int8', download_root=model_dir)\n";
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

    SendMessageA(g_progress, PBM_SETPOS, 80, 0);
    std::string ffmpeg = GetEffectiveFFmpeg();
    report += ffmpeg.empty()
        ? "ffmpeg: \xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93 (\xe3\x80\x8c" "ffmpeg\xe9\x81\xb8\xe6\x8a\x9e\xe3\x80\x8d\xe3\x81\xa7\xe6\x8c\x87\xe5\xae\x9a)\n"
        : ("ffmpeg: " + ffmpeg + "\n");
    SendMessageA(g_progress, PBM_SETPOS, 100, 0);
    SaveSettings();
    MsgBox(g_wnd,
        "\xe5\x88\x9d\xe6\x9c\x9f\xe8\xa8\xad\xe5\xae\x9a\xe5\xae\x8c\xe4\xba\x86\n\n" + report,
        "\xe3\x82\xbb\xe3\x83\x83\xe3\x83\x88\xe3\x82\xa2\xe3\x83\x83\xe3\x83\x97", MB_OK|MB_ICONINFORMATION);
    SetWindowTextW(g_status, L"Ready (v2.5)");
    SendMessageA(g_progress, PBM_SETPOS, 0, 0);
    UpdateWhisperLocLabels();
    g_busy = false;
}

// =========================================================================
// BrowseForFile
// =========================================================================

static std::string BrowseForFile(HWND parent, LPCWSTR filter, LPCWSTR title){
    wchar_t fn[MAX_PATH] = {};
    OPENFILENAMEW ofn = {sizeof(ofn)};
    ofn.hwndOwner = parent; ofn.lpstrFilter = filter;
    ofn.lpstrFile = fn; ofn.nMaxFile = MAX_PATH; ofn.lpstrTitle = title;
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
    if(GetOpenFileNameW(&ofn)) return WideToUtf8(fn);
    return "";
}

static std::string BrowseForFolder(HWND parent, LPCWSTR title){
    BROWSEINFOW bi = {};
    bi.hwndOwner = parent;
    bi.lpszTitle = title;
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_USENEWUI;
    PIDLIST_ABSOLUTE pidl = SHBrowseForFolderW(&bi);
    if(pidl){
        wchar_t path[MAX_PATH] = {};
        SHGetPathFromIDListW(pidl, path);
        CoTaskMemFree(pidl);
        return WideToUtf8(path);
    }
    return "";
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
    wchar_t cmd[2048];
    swprintf_s(cmd, 2048,
        L"\"%s\" -y -ss %.3f -i \"%s\" -t %.3f -vn -acodec pcm_s16le -ar 16000 -ac 1 \"%s\"",
        Utf8ToWide(ffmpeg).c_str(), clip.sourceOffset,
        Utf8ToWide(clip.filePath).c_str(), dur, Utf8ToWide(out).c_str());
    std::wstring cmdStr(cmd);
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

static std::vector<std::string> SplitText(const std::string& text, int maxChars){
    std::vector<std::string> res;
    if(maxChars <= 0 || (int)text.size() <= maxChars){
        res.push_back(text);
        return res;
    }
    // UTF-8 aware split
    int len = (int)text.size(), pos = 0;
    while(pos < len){
        int end = pos + maxChars;
        if(end >= len){ res.push_back(text.substr(pos)); break; }
        // find break point
        int bp = end;
        while(bp > pos && ((unsigned char)text[bp] & 0xC0) == 0x80) bp--;
        res.push_back(text.substr(pos, bp - pos));
        pos = bp;
    }
    return res;
}

// =========================================================================
// Run faster-whisper transcription
// =========================================================================

static bool RunFasterWhisper(int mi, int di, int bi, int beamSize, int li, float temp){
    const char* mn[] = {"tiny","base","small","medium","large-v3","large-v3-turbo"};
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
        bf << "{\n  \"model\": \"" << mn[mi] << "\",\n  \"language\": \"" << lc[li] << "\",\n  \"device\": \"" << dn[di] << "\",\n  \"backend\": \"" << bn[bi] << "\",\n  \"beam_size\": " << beamSize << ",\n  \"temperature\": " << tempStr << ",\n  \"model_dir\": \"";
        for(char c : md){ if(c == '\\') bf << "\\\\"; else bf << c; }
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
                for(char c : extraSp[i]){ if(c == '\\') bf << "\\\\"; else bf << c; }
                bf << "\"";
            }
        }
        bf << "],\n  \"clips\": [\n";
        bool first = true;
        for(size_t ci = 0; ci < g_tlClips.size(); ci++){
            if(wavs[ci].empty()) continue;
            if(!first) bf << ",\n"; first = false;
            std::string esc;
            for(char c : wavs[ci]){ if(c == '\\') esc += "\\\\"; else esc += c; }
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
            std::wstring torchCmd = Utf8ToWide("\"" + python + "\" -m pip install torch --upgrade --target=\"" + spDir + "\" --index-url https://download.pytorch.org/whl/cu121");
            std::string torchOut; RunProcess(torchCmd, torchOut, 900000);
            DebugLog("Auto torch CUDA install: " + torchOut.substr(0, 500));
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
    g_busy = true; SetProgress(0);
    DWORD startTick = GetTickCount();
    {std::ofstream f(Utf8ToWide(GetPluginDir() + "\\whisper_debug.log"), std::ios::trunc); f << "=== Whisper Subtitle v2.5 ===\n";}
    SaveSettings();
    char lt[16] = {}; GetWindowTextA(g_layerEdit, lt, sizeof(lt));
    int uiL = atoi(lt); if(uiL < 2 || uiL > 100) uiL = 2;
    int apiStartLayer = uiL - 1; // UI Layer 2 = API layer 1
    char mct[16] = {}; GetWindowTextA(g_maxCharEdit, mct, sizeof(mct));
    int maxC = atoi(mct); if(maxC < 0) maxC = 0;
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
    DebugLog("Template: " + (g_templateContent.empty() ? std::string("none") : g_templatePath));

    std::string ffmpeg = GetEffectiveFFmpeg();
    if(ffmpeg.empty()){
        MsgBox(g_wnd,
            "ffmpeg.exe\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n\n"
            "\xe3\x80\x8c" "ffmpeg\xe9\x81\xb8\xe6\x8a\x9e\xe3\x80\x8d\xe3\x83\x9c\xe3\x82\xbf\xe3\x83\xb3\xe3\x81\xa7\xe6\x8c\x87\xe5\xae\x9a\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84",
            "Error", MB_OK|MB_ICONERROR);
        SetStatus("Ready (v2.5)"); SetProgress(0); g_busy = false; return;
    }
    std::string python = GetEffectivePython();
    if(python.empty()){
        MsgBox(g_wnd,
            "Python\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n\n"
            "\xe3\x80\x8c\xe5\x88\x9d\xe6\x9c\x9f\xe8\xa8\xad\xe5\xae\x9a\xe3\x80\x8d\xe3\x81\xa7\xe3\x82\xbb\xe3\x83\x83\xe3\x83\x88\xe3\x82\xa2\xe3\x83\x83\xe3\x83\x97\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84",
            "Error", MB_OK|MB_ICONERROR);
        SetStatus("Ready (v2.5)"); SetProgress(0); g_busy = false; return;
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
        SetStatus("Ready (v2.5)"); SetProgress(0); g_busy = false; return;
    }
    DebugLog("Clips: " + std::to_string(g_tlClips.size()) + " Rate: " + std::to_string(g_projectRate));
    SetProgress(20);

    if(!RunFasterWhisper(mi, di, bi, beamSize, li, temp)){
        SetStatus("Ready (v2.5)"); SetProgress(0); g_busy = false; return;
    }

    if(g_segs.empty()){
        MsgBox(g_wnd,
            "\xe6\x96\x87\xe5\xad\x97\xe8\xb5\xb7\xe3\x81\x93\xe3\x81\x97\xe7\xb5\x90\xe6\x9e\x9c\xe3\x81\x8c\xe7\xa9\xba\xe3\x81\xa7\xe3\x81\x99\xe3\x80\x82\n"
            "\xe9\x9f\xb3\xe5\xa3\xb0\xe3\x83\x95\xe3\x82\xa1\xe3\x82\xa4\xe3\x83\xab\xe3\x82\x92\xe7\xa2\xba\xe8\xaa\x8d\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82",
            "\xe7\xb5\x90\xe6\x9e\x9c", MB_OK|MB_ICONWARNING);
        SetStatus("Ready (v2.5)"); SetProgress(0); g_busy = false; return;
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

    // Apply max char splitting FIRST (before trimming)
    struct PlaceItem { int s, e; std::string text; };
    std::vector<PlaceItem> items;
    for(auto& seg : g_segs){
        if(seg.e <= seg.s || seg.text.empty()) continue;
        std::vector<std::string> parts = SplitText(seg.text, maxC);
        if(parts.size() <= 1){
            items.push_back({seg.s, seg.e, seg.text});
        } else {
            int total = seg.e - seg.s;
            int partLen = total / (int)parts.size();
            if(partLen < 1) partLen = 1;
            for(size_t pi = 0; pi < parts.size(); pi++){
                int ps = seg.s + (int)pi * partLen;
                int pe = (pi == parts.size()-1) ? seg.e : (ps + partLen);
                if(pe > ps && !parts[pi].empty())
                    items.push_back({ps, pe, parts[pi]});
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
        DebugLog("Item " + std::to_string(i) + ": [" + std::to_string(finalItems[i].s) + "-" + std::to_string(finalItems[i].e) + "] \"" + finalItems[i].text.substr(0, 30) + "\"");
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
    SetStatus("Done! " + std::to_string(placed) + "\xe5\x80\x8b\xe3\x81\xae\xe5\xad\x97\xe5\xb9\x95\xe3\x82\x92\xe9\x85\x8d\xe7\xbd\xae (" + std::to_string(layerEnds.size()) + "Layer)" + timeBuf);
    DebugLog("Placed: " + std::to_string(placed));
    g_busy = false;
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
            if(id == IDC_GENERATE && !g_busy) std::thread(TranscribeThread).detach();
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
            else if(id == IDC_SETUP && !g_busy) std::thread(SetupThread).detach();
            else if(id == IDC_FFMPEG_BR){
                std::string p = BrowseForFile(g_wnd, L"ffmpeg.exe\0ffmpeg.exe\0All\0*.*\0", L"ffmpeg.exe\x3092\x9078\x629e");
                if(!p.empty()){
                    g_ffmpegPath = p;
                    SetPathLabel(g_ffmpegLabel, g_ffmpegPath, "(aviutl2.exe\xe3\x81\xae\xe5\xa0\xb4\xe6\x89\x80)");
                    SaveSettings();
                }
            }
            else if(id == IDC_PYTHON_BR){
                std::string p = BrowseForFile(g_wnd, L"python.exe\0python.exe\0All\0*.*\0", L"python.exe\x3092\x9078\x629e");
                if(!p.empty()){
                    g_pythonPath = p;
                    SetPathLabel(g_pythonLabel, g_pythonPath, "(\xe8\x87\xaa\xe5\x8b\x95\xe6\xa4\x9c\xe5\x87\xba)");
                    SaveSettings();
                }
            }
            else if(id == IDC_FW_BR){
                std::string p = BrowseForFolder(g_wnd, L"faster-whisper\x306esite-packages\x30d5\x30a9\x30eb\x30c0\x3092\x9078\x629e");
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
                std::string p = BrowseForFolder(g_wnd, L"whisper\x306esite-packages\x30d5\x30a9\x30eb\x30c0\x3092\x9078\x629e");
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
        }
        else if(HIWORD(w) == CBN_SELCHANGE){
            if((HWND)l == g_modelCombo || (HWND)l == g_deviceCombo || (HWND)l == g_backendCombo || (HWND)l == g_langCombo) SaveSettings();
        }
        else if(HIWORD(w) == EN_CHANGE){
            if((HWND)l == g_layerEdit || (HWND)l == g_maxCharEdit || (HWND)l == g_lingerEdit) SaveSettings();
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
        free((void*)l);
    }
    else if(m == WM_UPDATE_PROGRESS){
        SendMessageA(g_progress, PBM_SETPOS, w, 0);
    }
    return DefWindowProcW(h, m, w, l);
}

// =========================================================================
// Detect whisper install locations
// =========================================================================

static void UpdateWhisperLocLabels(){
    std::string fwDir = GetEffectiveFwSpDir();
    std::string owDir = GetEffectiveOwSpDir();

    std::string fwText = "faster-whisper: ";
    if(!fwDir.empty()) fwText += fwDir;
    else fwText += "(\xe6\x9c\xaa\xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab)";
    if(!g_fwSpPath.empty()) fwText += " [\xe6\x89\x8b\xe5\x8b\x95]"; // [手動]

    std::string owText = "whisper: ";
    if(!owDir.empty()) owText += owDir;
    else owText += "(\xe6\x9c\xaa\xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab)";
    if(!g_owSpPath.empty()) owText += " [\xe6\x89\x8b\xe5\x8b\x95]"; // [手動]

    if(g_fwLocLabel) SetWindowTextW(g_fwLocLabel, Utf8ToWide(fwText).c_str());
    if(g_owLocLabel) SetWindowTextW(g_owLocLabel, Utf8ToWide(owText).c_str());
}

// =========================================================================
// RegisterPlugin
// =========================================================================

extern "C" __declspec(dllexport) void __cdecl RegisterPlugin(HOST_APP_TABLE* host){
    host->set_plugin_information(L"Whisper Subtitle v2.5");
    InitCommonControls();
    EnsureDirectories(); EnsurePyHelper();

    WNDCLASSEXW wc = {sizeof(wc)};
    wc.lpszClassName = L"WhisperSub25";
    wc.lpfnWndProc = WndProc;
    wc.hInstance = g_hInst;
    wc.hbrBackground = GetSysColorBrush(COLOR_BTNFACE);
    RegisterClassExW(&wc);

    g_wnd = CreateWindowExW(0, L"WhisperSub25", L"Whisper", WS_POPUP, 0, 0, 360, 480, 0, 0, g_hInst, 0);
    int W = 340;
    HWND hw;

    // === Tab control ===
    g_tab = CreateWindowExW(0, WC_TABCONTROLW, L"", WS_VISIBLE|WS_CHILD|WS_CLIPSIBLINGS, 4, 2, W+6, 440, g_wnd, 0, g_hInst, 0);
    TCITEMW ti = {TCIF_TEXT};
    ti.pszText = (LPWSTR)L"\x5b57\x5e55\x751f\x6210";
    TabCtrl_InsertItem(g_tab, 0, &ti);
    ti.pszText = (LPWSTR)L"\x521d\x671f\x8a2d\x5b9a";
    TabCtrl_InsertItem(g_tab, 1, &ti);
    int tabY = 28;

    // === Tab 0: Subtitle ===
    int y = tabY + 6;
    hw = CreateWindowExW(0, L"STATIC", L"Backend:", WS_CHILD|WS_VISIBLE, 14, y+3, 52, 18, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    g_backendCombo = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD|WS_VISIBLE|CBS_DROPDOWNLIST, 70, y, W-80, 80, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_backendCombo);
    SendMessageW(g_backendCombo, CB_ADDSTRING, 0, (LPARAM)L"faster-whisper");
    SendMessageW(g_backendCombo, CB_ADDSTRING, 0, (LPARAM)L"whisper");
    SendMessageA(g_backendCombo, CB_SETCURSEL, 0, 0);
    y += 24;
    hw = CreateWindowExW(0, L"STATIC", L"Model:", WS_CHILD|WS_VISIBLE, 14, y+3, 52, 18, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    g_modelCombo = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD|WS_VISIBLE|CBS_DROPDOWNLIST|WS_VSCROLL, 70, y, W-80, 120, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_modelCombo);
    SendMessageW(g_modelCombo, CB_ADDSTRING, 0, (LPARAM)L"tiny");
    SendMessageW(g_modelCombo, CB_ADDSTRING, 0, (LPARAM)L"base");
    SendMessageW(g_modelCombo, CB_ADDSTRING, 0, (LPARAM)L"small");
    SendMessageW(g_modelCombo, CB_ADDSTRING, 0, (LPARAM)L"medium");
    SendMessageW(g_modelCombo, CB_ADDSTRING, 0, (LPARAM)L"large-v3");
    SendMessageW(g_modelCombo, CB_ADDSTRING, 0, (LPARAM)L"large-v3-turbo");
    SendMessageA(g_modelCombo, CB_SETCURSEL, 5, 0);
    y += 24;
    hw = CreateWindowExW(0, L"STATIC", L"Beam:", WS_CHILD|WS_VISIBLE, 14, y+3, 38, 18, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    g_qualityEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"5", WS_CHILD|WS_VISIBLE|ES_NUMBER|ES_CENTER, 52, y, 30, 22, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_qualityEdit);
    hw = CreateWindowExW(0, L"STATIC", L"(1-10)", WS_CHILD|WS_VISIBLE, 85, y+3, 42, 18, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    hw = CreateWindowExW(0, L"STATIC", L"Temp:", WS_CHILD|WS_VISIBLE, 135, y+3, 38, 18, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    g_tempEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"0", WS_CHILD|WS_VISIBLE|ES_CENTER, 175, y, 30, 22, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_tempEdit);
    hw = CreateWindowExW(0, L"STATIC", L"(0=\x56fa\x5b9a)", WS_CHILD|WS_VISIBLE, 208, y+3, 80, 18, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    y += 24;
    hw = CreateWindowExW(0, L"STATIC", L"Device:", WS_CHILD|WS_VISIBLE, 14, y+3, 52, 18, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    g_deviceCombo = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD|WS_VISIBLE|CBS_DROPDOWNLIST, 70, y, W-80, 80, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_deviceCombo);
    SendMessageW(g_deviceCombo, CB_ADDSTRING, 0, (LPARAM)L"\x81ea\x52d5\x9078\x629e");
    SendMessageW(g_deviceCombo, CB_ADDSTRING, 0, (LPARAM)L"CPU");
    SendMessageW(g_deviceCombo, CB_ADDSTRING, 0, (LPARAM)L"CUDA (GPU)");
    SendMessageA(g_deviceCombo, CB_SETCURSEL, 2, 0);
    y += 24;
    hw = CreateWindowExW(0, L"STATIC", L"\x8a00\x8a9e:", WS_CHILD|WS_VISIBLE, 14, y+3, 52, 18, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    g_langCombo = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD|WS_VISIBLE|CBS_DROPDOWNLIST|WS_VSCROLL, 70, y, W-80, 200, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_langCombo);
    SendMessageW(g_langCombo, CB_ADDSTRING, 0, (LPARAM)L"\x81ea\x52d5\x5224\x5b9a");
    SendMessageW(g_langCombo, CB_ADDSTRING, 0, (LPARAM)L"\x65e5\x672c\x8a9e (ja)");
    SendMessageW(g_langCombo, CB_ADDSTRING, 0, (LPARAM)L"\x82f1\x8a9e (en)");
    SendMessageW(g_langCombo, CB_ADDSTRING, 0, (LPARAM)L"\x4e2d\x56fd\x8a9e (zh)");
    SendMessageW(g_langCombo, CB_ADDSTRING, 0, (LPARAM)L"\x97d3\x56fd\x8a9e (ko)");
    SendMessageA(g_langCombo, CB_SETCURSEL, 0, 0);
    y += 24;
    hw = CreateWindowExW(0, L"STATIC", L"Layer:", WS_CHILD|WS_VISIBLE, 14, y+3, 52, 18, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    g_layerEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"2", WS_CHILD|WS_VISIBLE|ES_NUMBER|ES_CENTER, 70, y, 40, 22, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_layerEdit);
    hw = CreateWindowExW(0, L"STATIC", L"(\x5b57\x5e55\x914d\x7f6e\x5148)", WS_CHILD|WS_VISIBLE, 115, y+3, 185, 18, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    y += 24;
    hw = CreateWindowExW(0, L"STATIC", L"\x6587\x5b57\x6570:", WS_CHILD|WS_VISIBLE, 14, y+3, 52, 18, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    g_maxCharEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"0", WS_CHILD|WS_VISIBLE|ES_NUMBER|ES_CENTER, 70, y, 40, 22, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_maxCharEdit);
    hw = CreateWindowExW(0, L"STATIC", L"(0=\x7121\x5236\x9650)", WS_CHILD|WS_VISIBLE, 115, y+3, 185, 18, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    y += 24;
    hw = CreateWindowExW(0, L"STATIC", L"\x66f8\x5f0f:", WS_CHILD|WS_VISIBLE, 14, y+3, 42, 18, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    g_templateLabel = CreateWindowExW(0, L"STATIC", L"\x30c7\x30d5\x30a9\x30eb\x30c8", WS_CHILD|WS_VISIBLE, 58, y+3, 110, 18, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_templateLabel);
    hw = CreateWindowExW(0, L"BUTTON", L"\x9078\x629e", WS_CHILD|WS_VISIBLE, 175, y, 55, 22, g_wnd, (HMENU)IDC_TEMPLATE, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    hw = CreateWindowExW(0, L"BUTTON", L"\x30ea\x30bb\x30c3\x30c8", WS_CHILD|WS_VISIBLE, W-75, y, 65, 22, g_wnd, (HMENU)IDC_RESET_TPL, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    y += 26;
    // (Insert mode removed - auto layer shift handles conflicts)
    y += 24;
    hw = CreateWindowExW(0, L"BUTTON", L"\x5b57\x5e55\x751f\x6210", WS_CHILD|WS_VISIBLE|BS_DEFPUSHBUTTON, 14, y, (W-24)/2 - 3, 34, g_wnd, (HMENU)IDC_GENERATE, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    hw = CreateWindowExW(0, L"BUTTON", L"SRT\x30a8\x30af\x30b9\x30dd\x30fc\x30c8", WS_CHILD|WS_VISIBLE, 14 + (W-24)/2 + 6, y, (W-24)/2 - 3, 34, g_wnd, (HMENU)IDC_EXPORT_SRT, g_hInst, 0); g_tabSubCtrls.push_back(hw);
    y += 40;
    g_progress = CreateWindowExW(0, PROGRESS_CLASSW, L"", WS_CHILD|WS_VISIBLE, 14, y, W-24, 14, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_progress);
    SendMessageA(g_progress, PBM_SETRANGE, 0, MAKELPARAM(0, 100));
    y += 18;
    g_status = CreateWindowExW(0, L"STATIC", L"Ready (v2.5)", WS_CHILD|WS_VISIBLE, 14, y, W-24, 20, g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_status);

    // === Tab 1: Setup ===
    y = tabY + 10;
    hw = CreateWindowExW(0, L"STATIC", L"ffmpeg:", WS_CHILD, 14, y+3, 46, 18, g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    g_ffmpegLabel = CreateWindowExW(0, L"STATIC", L"(aviutl2.exe\x306e\x5834\x6240)", WS_CHILD|SS_PATHELLIPSIS, 64, y+3, 190, 18, g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_ffmpegLabel);
    hw = CreateWindowExW(0, L"BUTTON", L"\x9078\x629e", WS_CHILD, W-55, y, 50, 22, g_wnd, (HMENU)IDC_FFMPEG_BR, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    y += 26;
    hw = CreateWindowExW(0, L"STATIC", L"Python:", WS_CHILD, 14, y+3, 46, 18, g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    g_pythonLabel = CreateWindowExW(0, L"STATIC", L"(\x81ea\x52d5\x691c\x51fa)", WS_CHILD|SS_PATHELLIPSIS, 64, y+3, 190, 18, g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_pythonLabel);
    hw = CreateWindowExW(0, L"BUTTON", L"\x9078\x629e", WS_CHILD, W-55, y, 50, 22, g_wnd, (HMENU)IDC_PYTHON_BR, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    y += 30;
    // faster-whisper location
    g_fwLocLabel = CreateWindowExW(0, L"STATIC", L"faster-whisper: ...", WS_CHILD|SS_PATHELLIPSIS, 14, y+3, W-130, 18, g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_fwLocLabel);
    hw = CreateWindowExW(0, L"BUTTON", L"\x9078\x629e", WS_CHILD, W-110, y, 50, 22, g_wnd, (HMENU)IDC_FW_BR, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    hw = CreateWindowExW(0, L"BUTTON", L"\x81ea\x52d5", WS_CHILD, W-55, y, 50, 22, g_wnd, (HMENU)IDC_FW_RESET, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    y += 24;
    // openai-whisper location
    g_owLocLabel = CreateWindowExW(0, L"STATIC", L"whisper: ...", WS_CHILD|SS_PATHELLIPSIS, 14, y+3, W-130, 18, g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_owLocLabel);
    hw = CreateWindowExW(0, L"BUTTON", L"\x9078\x629e", WS_CHILD, W-110, y, 50, 22, g_wnd, (HMENU)IDC_OW_BR, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    hw = CreateWindowExW(0, L"BUTTON", L"\x81ea\x52d5", WS_CHILD, W-55, y, 50, 22, g_wnd, (HMENU)IDC_OW_RESET, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    y += 32;
    hw = CreateWindowExW(0, L"BUTTON", L"\x30bb\x30c3\x30c8\x30a2\x30c3\x30d7 (\x30a4\x30f3\x30b9\x30c8\x30fc\x30eb + \x30e2\x30c7\x30eb" L"DL)", WS_CHILD, 14, y, W-24, 30, g_wnd, (HMENU)IDC_SETUP, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    y += 38;
    hw = CreateWindowExW(0, L"STATIC", L"\x203b Backend\x3068Model\x306f\x300c\x5b57\x5e55\x751f\x6210\x300d\x30bf\x30d6\x3092\x53c2\x7167", WS_CHILD, 14, y, W-24, 20, g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    y += 28;
    // --- Text processing settings ---
    hw = CreateWindowExW(0, L"STATIC", L"\x2500\x2500 \x30c6\x30ad\x30b9\x30c8\x51e6\x7406 \x2500\x2500", WS_CHILD, 14, y, W-24, 18, g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    y += 22;
    g_chkRemovePunct = CreateWindowExW(0, L"BUTTON", L"\x53e5\x8aad\x70b9\x524a\x9664", WS_CHILD|BS_AUTOCHECKBOX, 14, y, 100, 18, g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_chkRemovePunct);
    g_chkRemoveExclam = CreateWindowExW(0, L"BUTTON", L"!?\x524a\x9664", WS_CHILD|BS_AUTOCHECKBOX, 116, y, 72, 18, g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_chkRemoveExclam);
    g_chkNormalize = CreateWindowExW(0, L"BUTTON", L"\x5168\x534a\x89d2\x6b63\x898f\x5316", WS_CHILD|BS_AUTOCHECKBOX, 192, y, 120, 18, g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_chkNormalize);
    y += 24;
    // Linger time
    hw = CreateWindowExW(0, L"STATIC", L"\x5b57\x5e55\x5ef6\x9577:", WS_CHILD, 14, y+3, 62, 18, g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(hw);
    g_lingerEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"1.0", WS_CHILD|ES_CENTER, 78, y, 40, 22, g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_lingerEdit);
    hw = CreateWindowExW(0, L"STATIC", L"\x79d2 (0=\x306a\x3057)", WS_CHILD, 122, y+3, 100, 18, g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(hw);

    SwitchTab(0);
    host->register_window_client(L"Whisper Subtitle", g_wnd);
    g_edit = host->create_edit_handle();
    LoadSettings();

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
$cpp | Out-File $cppPath -Encoding UTF8
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
project(whisper_subtitle LANGUAGES CXX)
set(CMAKE_CXX_STANDARD 17)
add_library(whisper_subtitle SHARED src/whisper_subtitle.cpp)
target_compile_definitions(whisper_subtitle PRIVATE UNICODE _UNICODE)
set_target_properties(whisper_subtitle PROPERTIES
    OUTPUT_NAME "whisper_subtitle"
    SUFFIX ".aux2"
    PREFIX ""
    RUNTIME_OUTPUT_DIRECTORY_RELEASE "`${CMAKE_BINARY_DIR}/Release"
)
target_link_libraries(whisper_subtitle PRIVATE comctl32 shell32 comdlg32 ole32)
"@
$cmake | Out-File "$projDir\CMakeLists.txt" -Encoding UTF8
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
    $aux2 = Get-ChildItem -Path $buildDir -Recurse -Filter "whisper_subtitle.aux2" | Select-Object -First 1
    if($aux2){
        $dest = "$d\whisper_subtitle.aux2"
        Copy-Item $aux2.FullName $dest -Force
        Write-Host ""
        Write-Host "============================================" -ForegroundColor Green
        Write-Host " BUILD SUCCESS! v2.5" -ForegroundColor Green
        Write-Host "============================================" -ForegroundColor Green
        Write-Host "Output: $dest" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "v2.5 Features:" -ForegroundColor Cyan
        Write-Host "  + Setup (faster-whisper + model auto-DL)" -ForegroundColor White
        Write-Host "  + ffmpeg / Python path selector" -ForegroundColor White
        Write-Host "  + Fixed: result file always created" -ForegroundColor White
        Write-Host "  + Better error messages with debug log" -ForegroundColor White
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
