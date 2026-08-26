##############################################################################
# Whisper Subtitle Plugin for AviUtl2 - v2.9.75
# v2.9.66: 【AY1】**字幕が音より早く出すぎる**のを抑えた。開発者の指摘「すべて少し早め。字幕だから少し早めに出るのは良いんだが、それにしても速すぎる」。実データ228件を波形のエネルギーで測ったところ、**開始時点が無音の字幕が129件**あり、音が立つまで**中央値170ms / 最大1003ms**待たされていた(300ms以上が14件)。→ `_speech_env` で10ms刻みの短時間エネルギー包絡を作り、`_cap_lead` で **「音が立つ _MAX_LEAD(既定0.15秒) 前」より早くは出さない**。★**動かす向きは遅らせるだけ**。早める方向には一切動かさない(実測でも早まったもの0件)。★開始時点で既に音が出ているものは触らない(そこは合っている)。★遅らせた結果0.1秒未満になるなら触らない(字幕が潰れるほうが害が大きい)。★包絡が作れなければ素通り(従来と同一)。 実測: 早さの最大が **403〜1003ms → 163ms**、中央値は150ms前後を維持(自然なリードは残す)。**件数は不変**。★**VAD区間(`_speech_runs`)は使えない**。76秒に17区間しかなく多くの字幕が区間の内側に落ちるため、「最も近い発話開始」が数秒先になって judge できない。**粗い正典で測って誤診しかけた**。★163ms と 150ms の差は 60fps のフレーム量子化(16ms)と包絡の刻み(10ms)。
# v2.9.65: 【AX1/AX2】**字幕の切れ目が不自然**なのを直した。実機で開発者が挙げた違和感(「勉強して|きた!」「向いて|もらっても」「そうじゃ|ないのか」「静か|すぎるな」「おかし|すぎるなと」「もったいない|ことを」)は**すべて同じ1つの型**だった: **単独では意味を成さず前の語に寄りかかる成分が、行の先頭に落ちている**。 【AX1】文節を作るとき、`動詞/形容詞-非自立可能`(補助動詞「〜てくる」「〜ていく」/ 補助形容詞「〜ない」/ 程度「〜すぎる」/ サ変「真似する」)を、文節の途中なら前にくっつける。形式名詞(こと/もの/ため/はず/わけ 等)は**直前が動詞・形容詞・助動詞のときだけ**くっつける(そうしないと「事が起きた」の「事」まで巻き込む)。 【AX2】**貪欲詰め込み → 均等割り(DP)**。従来は maxc まで目一杯詰めたので2つ目が余り物になっていた(17字/maxc14 で 13+4)。個数は最小のまま、長さのばらつきが最小になる割り方を選ぶ(6+11)。切れ目の自然さ(句読点>助詞>その他)も加点する。★**字幕の数は変わらない**。実データ6本で件数は全て同一(45/48/37/52/46/38)、割る位置だけが変わる。★実データ104件のうち**2つ以上に割れる29件中25件**が自然な位置に変わり、**文字の欠落は0件**。★スイート39番に登録(v2.9.64/62/51 で9件落ちることを確認済み)。★n>400 の異常系は DP をやめて従来の貪欲に落ちる保険つき。
# v2.9.64: 【AV1】**詰める対象が狭すぎた**。v2.9.63 の AU2 は「ー〜～っッ」だけを見ていたが、whisper の繰り返しループは**毎回ちがう綴り**で出る。同じ音声を5回転写した実測: run1「ぉ」x222 / run2「お」x444・「え」x75 / run3「ー」x444 / run4「え」x10 / run5「ー」x5 → **5回中2回しか拾えていなかった**。実機(v2.9.63)では `squash=0` のまま「おぉぉぉ…」(小書きの ぉ)が **14文字x16件・13.7秒ぶん**並んだ。★「ぉ」は `noHead` に入っていない(禁則は 、。っゃゅょー！？ で **ぁぃぅぇぉ が無い**)ので AU1 も効かず、1フレーム化はしない代わりに**14文字ずつ16件**になった。→ 特定の記号ではなく **かな(ひらがな/カタカナ/長音記号/波ダッシュ)の同一文字連続**を対象にする。★数字と英字は入れない(「1000000」は 0 が6連続するので**数値が壊れる**)。★漢字も入れない(実データで「熱」x9 が出たが14文字に収まり実害が無く、正当な重ね字を壊す危険のほうが大きい)。 ★★**判断の型**: v2.9.63 で「別要因がもう1つある」と引き継ぎに書いた run2 の84件は、**同じ繰り返しループの別の綴り**だった。**症状の見た目が違うだけで同じ家族**。1つ直したら、**同じ原因が他の綴り/他の文字で出ていないかを実データで数える**(`検証ツール\choon_repro\scan_runs.py`)。
# v2.9.63: 【AU1/AU2】**ライブ配信素材で字幕が466件に膨張する**のを止めた。原因は当初「プロ分割+maxchars が切り刻んでいる」と記録していたが、**実物を走らせたら違った**。Python の `_morph_group` は長音の連続を最大2チャンクにしかしない(「あ」+「ー」x200 で2チャンク。1文節が maxc を超えると割らないため)。刻んでいたのは **C++ の SplitText の禁則処理**。「ー」は行頭禁止文字(3402行)で、良い区切りが無いときの後退ループ(3430行)が **cutAfter を start まで戻し切り**、進行保証 `nextStart = start + 1` に落ちて **1文字ずつ**に割れる。さらに `partLen = total / parts.size()` が 0 → 1 にクランプされ、**1フレームの「ー」が数百件**並ぶ。実物の SplitText を単体ビルドして確認: 「あ」+「ー」x200 → **187パート / 全部2文字以下 / 全部6フレーム未満 / 最短1フレーム**。★SplitText は文節区切り ON でも**無条件で全セグメントに走る**(4201行)ので、プロ分割の ON/OFF に関係なく踏む。 【AU1】禁則を満たす切り位置が無ければ **禁則を諦めて上限で切る**。窓の中が全部禁則文字のときだけ発火するので普通の文には効かない — **実データ6本303行のうち切り位置が変わったのは事故が起きている2行だけ**(過去の実データ38行/正常4ラン227行は1行も変わらず)。 【AU2】Python 側で、同じ引き伸ばし字(ー〜～っッ)が **4個以上続いたら3個に詰める**。落とした単語の**時間は直前の単語に足す**(捨てると4秒の叫びが0.3秒の点滅になる)。★「あー」「あーー」「あーーー」は一切変わらない。 ★★同一素材を実機と同じ引数で5回転写して**ブレ幅を実測**した: whisperセグ32〜45 / 出力38〜84件 / 長音の最長連続 1〜444。**膨張したのは5回中1回**。1回の実行で判断してはいけないことを数字で確認した。 ★オフライン検証(実データ6本): 現状455件 → AU1のみ68件 → AU1+AU2 **38件**。
# v2.9.62: 【AT1】**openai 経路の進捗バーがログとエラー表示を潰していた**。`verbose=False` は**進捗バーを消す指定ではない**。openai-whisper の transcribe.py を実物で確認したところ 263行「show the progress bar when verbose is False」/ 265行 `disable = verbose is not False` で、**False を渡すと disable=False = バーが出る**。直感と逆。実害は3つ: ①`whisper_debug.log` が `37%|###7 | 2832/7585 [00:10<00:17, 266.45frames/s]` で埋まる(実測) ②`RunProcess` の output に**無制限に溜まる** ③★**生成が失敗したとき、エラーダイアログは pyOut の末尾500文字を「--- Python output ---」として見せるので、肝心のエラーが進捗バーに押し出される**。今日ずっと直してきた「失敗の理由を正しく伝える」に真っ向から反していた。→ `verbose=None` にする。★**verbose は表示専用**で、デコードにも結果にも一切影響しない(147/154/478行はすべて print のみ。既定値も None)。★**前回この修正を「転写の引数を変える話なので要判断」として見送ったのは誤りだった**。実物を開いて確認したら表示専用だった。/ ★副作用の申し送り: tqdm が書かなくなるので「AviUtl2 を閉じると壊れたパイプで python が自滅する」偶然の保護が無くなる。ただし v2.9.52【AI2】の Job Object が正規の手段なので問題ない。/ ★スイート25 に検出器を追加。**この罠は「黙らせるつもりで False に戻す」が起きやすい**ので、そこを止めるのが狙い。v2.9.61 で落ち、v2.9.62 で合格を確認済み。
# v2.9.61: 【AS1】**編集できない理由を具体的に伝えるようにした**。v2.9.59【AO1】で `call_edit_section_param` の戻り値を見るようにしたが、**なぜ失敗したかは書いていなかった**(「編集が行える状態か確認してください」)。最新SDK(sdk2.1.4) の plugin2.h 688行に「編集が出来ない場合(**出力中等**)に失敗します」と明記されており、EDIT_HANDLE に `EDIT_STATE_PLAY`(プレビュー再生中) / `EDIT_STATE_SAVE`(ファイル出力中) が定義されている。★**プレビューを流したまま生成を押す**のはごく普通の操作で、そこで失敗したとき何を止めればよいか伝わっていなかった。スキャン/配置/書式適用の3メッセージに「プレビュー再生中・ファイル出力中は〜」「停止してから実行してください」を追加。/ ★**`get_edit_state()` は使わなかった**。SDK index 11 でプラグインの EDIT_HANDLE は 11個(0-10)、**あと1個で届く**が、この関数は 2026/4/12 追加なので**それより古い AviUtl2 では構造体がそこまで無く、範囲外を読んでクラッシュする**。`RequiredVersion` で下限を宣言する仕組みはあるが本プラグインは未使用。動く環境を壊す危険が、得られる情報に見合わない。**構造体には触らず、情報だけ足す**方式にした。/ ★**同時に HANDOVER の誤記録を訂正**: 「編集APIをワーカースレッドから叩いている(深層知識5-1違反)」は**誤り**だった。plugin2.h 685行「**コールバック関数はメインスレッドから呼ばれます**」。誤った原因は、私が読んでいたのが **sdk47(bool 14個)で3世代古かった**こと。★逆に `call_read_section` は 737行「**呼び出し元と同じスレッドで**呼ばれます」なので、読むだけだからと安易に切り替えると**そちらが本当の違反になる**。/ ★SDK は開発者が sdk2.1.4 に更新済み。**ABI 互換を実測**: プラグインの自前構造体は EDIT_SECTION 33個 / EDIT_HANDLE 11個で、最新(78個/32個)の**先頭と完全一致**(純粋な末尾追加)。
# v2.9.60: 【AP1】**複数オブジェクトのエイリアスを書式テンプレートに選べてしまい、字幕1つにつき全オブジェクトが作られていた**。SDK の契約(plugin2.h 125行)に「複数オブジェクトのエイリアスデータの場合は先頭のオブジェクトのハンドルが返却されます ※**オブジェクトは全て作成されます**」と明記されている。戻り値は1つなので Pass2 は replaced を1しか数えず、配置側は「字幕1つ=1オブジェクト」でレイヤーを詰めているため、**余分なオブジェクトが確保していないレイヤーに載って利用者の物と衝突する**。★**ログは「replaced: N failed: 0」と出るのにタイムラインは壊れる**。/ ★R1(テキスト項目チェック)では**捕まらない**。実測: 手元の `コメント.object` は**4オブジェクトだが `テキスト=` を含むので R1 を素通り**していた。開発者の Alias フォルダに実在するファイルで再現しうる。/ ★形式は**実物194件を分類して確定**した: 単一=`[Object]`+`[Object.N]`(N はエフェクト番号/191件) / 複数=`[0]`,`[0.N]`,`[1]`,`[1.N]`…(先頭の数字がオブジェクト番号/3件)。よって **`^[数字]$` の行数**を数えればオブジェクト数が分かる。判定は Python に移植して実物に当て、3件だけを拒否することを確認済み。/ ★R1 と同じ「**選択時に**受け付けず理由を伝える」形にした(生成をブロックする側ではないので詰まらせない)。`LoadTemplate` に `why` 出力引数を足し、既定 nullptr で既存呼び出しは不変。★スイート10 に検出器を追加。v2.9.59 で3件落ち、v2.9.60 で合格を確認済み。★**この発見は「plugin2.h のコメントを正典として持ち込んだ」ことによる**。ソースを別角度で眺める掃き出しは3連続で空振りしていた。
# v2.9.59: 【AO1/AO2 + 監査37】**SDK の戻り値を16箇所すべて捨てていた**。「失敗を成功として扱う」型の **SDK 境界版**。ファイル書き込み側は v2.9.55 の監査35 で塞いだが、**SDK 呼び出しは誰も見ていなかった**。/ 【AO1】`call_edit_section_param` は **bool を返す**(plugin2.h 296行)。従来は `if(g_edit)` でハンドルの null だけ見て戻り値を捨てていた。**false ならコールバックは一度も走らない**のに処理は続く。★最悪は Pass2 — 入れなかった場合、書式テンプレートが一切当たらないのにログは「Pass2 replaced: 0 failed: 0」= 置き換える物が無かったようにしか読めず、利用者には既定書式の字幕が残る。R1 と同じ「テンプレが効かないのに成功に見える」型。→ 4箇所すべてで受ける。スキャン/Pass1 は理由を出して中止、Pass2 は**字幕は置けているので止めず**に「書式だけ当たっていない」と伝える(K1 の側=詰まらせない)、レイヤー占有調べは記録のみ。/ 【AO2】`set_object_item_value` も **bool を返す**(plugin2.h 265行)。12箇所すべて捨てていた。★致命的なのは**本文**の設定で、失敗すると**中身が空の字幕オブジェクトが置かれ placed に数えられる**。項目名は AviUtl2 の日本語名を直接指定しているので、**本体側で名前が変わると全件が空になるのに「配置 N / 失敗 0」と出る**。→ 本文は成否を見て、失敗したらオブジェクトを消して failed に数える。装飾(フォント/サイズ/色/揃え/Y)は当たらなくても既定書式で字幕が出るので**数えて記録するだけ**。★Pass1 と Pass2 の復元経路で**同じ扱いに揃えた**(片方だけにしない)。★`styleFailed` が 0 でなくなったら AviUtl2 側の項目名が変わった合図。/ 【監査37】`audit_sdk_returns.py` を新設しランナーに登録。**正典は plugin2.h から自動で読む**ので、SDK 更新で bool 関数が増えても勝手に追従する(手で並べると必ず追従漏れになる)。★v2.9.58 と v2.9.51 で12件落ち、v2.9.59 で合格することを確認済み。
# v2.9.58: 【AN1】**テストログが UI を読み直していた**。`WriteTestLog()` の冒頭コメント自身が「設定は whisper_debug.log から抜き出す。UIから読み直すと二重管理になるうえ『実際にその実行で使われた値』とズレる危険がある」と書いているのに、**backend と beam だけがそれを破っていた**。実害は2つ: ①生成側は v2.9.14【監査③】で beam に**上限20**を足したが**こちらは上限が無い**ため、21以上を入れると**生成は 20 で走るのにテストログは入力値をそのまま記録する**。テストログは「実行間の比較」のためのものなので比較の前提が壊れる ②生成中に UI を変えられる(SetBusy が止めるのは Generate/Setup だけ)ので、AH4/AK1 と同じ形で「その回に使っていない値」を記録しうる。→ 上で集めた settings(= SETTINGS 行)から拾う方式に統一(runinfo と同じ出どころ)。★「同じ計算が2箇所」と「UIの読み直し」と「誤情報を出す(AG5)」が重なった形。★**数値入力のクランプを全欄で掃いて見つけた**(g_layerEdit/g_lingerEdit/g_leadEdit/g_qualityEdit/g_tempEdit/g_maxCharEdit)。残る g_maxCharEdit の上限なしは、大きい値=分割しないだけで実害が無いため見送り。★スイート29 に検出器を追加。v2.9.57 で2件落ち、v2.9.58 で合格を確認済み。
# v2.9.57: 【AM2/AM3/AM4】**「本実行 ↔ 再試行」の対を突き合わせて3件**。この対は AE1(タイムアウト)・AI1(理由の拾い直し)に続き**4回目**。/ 【AM2】再試行の ERROR 行に**日本語説明が無く英語だけが出ていた**。AG3 で「エラーを出す経路は2つある。両方に置くこと」と直したが、**再試行側は対象外のままだった**。しかも再試行は**パッケージを自動導入した直後**に走るので、numpy 不整合が最も出やすい場面。→ `NumpyHintIfNeeded()` に集約し**3箇所すべてから呼ぶ**(本実行の結果ファイル無し / 本実行の ERROR 行 / 再試行の ERROR 行)。★4箇所目を足すときにまたコピーすると必ず片方だけになるので集約が要点。/ 【AM3】**再試行前に結果と .err を消していなかった**。本実行は Q1 対策で消しているのに片方だけ。再試行の python が起動に失敗すると**本実行が書いた古い ERROR 行がそのまま読まれ**、既に解消した「〜 not installed」を今回の原因として表示する。Q1 と同じ形の片割れ。/ 【AM4】**再試行前に部分結果を捨てていなかった**。本実行のループは ERROR 行に当たるまでに一部を push しうるが、再試行はファイルを先頭から読み直すので**同じ字幕が二重に積まれる**。AC1(部分push後に別方式が再pushして二重)と同型。→ `g_segs.clear()`。/ ★スイート21 を更新。判定を集約したので**数えるものをリテラルから呼び出し箇所に変えた**(集約したとたんに空振りして「直したのに NG」になるため)。窓も 1200→4000 字(説明コメントを足しただけで取り切れなくなり誤判定した。今日3回目)。★v2.9.56 で7件・AG4 修正前の v2.9.44 で15件落ちることを確認済み。
# v2.9.56: 【AM1】**SRT が配置と食い違っていた**。配置側は延長(linger)を足したあと**タイムライン終端でクランプ**しているのに、SRT 側だけそれが無かった。そのため linger を使うと、タイムライン上の最後の字幕は動画の終わりで止まるのに **SRT では動画の終端を超えて続く**。同じ1回の生成の2つの出力が食い違う。★配置側(items)と SRT 側(srtSegs)は**独自に同じ計算を持っている**ので、片方だけ直すと必ずこうなる。**V2(先頭の潰れ) → V3(長さ0の除去) → AM1(終端クランプ) と3回目**。★終端の求め方も配置側と同じ(g_tlClips の最大 timelineEnd、0 なら掛けない)にした。★以後ずれないよう、スイート24 に**守りを対で検査する**行を6本追加(延長の終端クランプ/終端の求め方/先行表示の0クランプ/先頭の潰れ戻し/重なり解消/長さ0以下の除去)。v2.9.55 で2件落ち、v2.9.56 で合格することを確認済み。
# v2.9.55: 【AL1/AL2/AL3 + 監査35】「書けたかを見ていない」型を**家族ごと**塞いだ。AD1(SRT)・AJ1(ヘルパー)で2回直したが、全数を掃いたら**まだ4箇所**残っていた。★ofstream は**開いた時点で truncate する**ので、開けたかしか見ていないと「壊したのに成功として進む」。失敗が出るのは flush 時なので **close() まで済ませてから**確かめる必要がある。/ 【AL1】`SaveSettings()` — ini が途中まで書かれても黙って進み、次回起動で**設定の一部が既定値に戻る**。利用者には「設定が勝手に戻る」としか見えない。記録は毎回・通知はセッション中1回だけ(SaveSettings は設定変更のたびに呼ばれるため)。/ 【AL2】**バッチ JSON が is_open すら見ていなかった**。しかも `bp` は固定パスで、Q1 の対策(実行前に消す)は**結果ファイルと .err にしか掛かっていない**。書き込みに失敗すると**前回の whisper_batch.json が残り、Python が前の動画のクリップ一覧を読んで転写する**。★Q1(別の動画の字幕)と同じ形が、出力側ではなく**入力側**に残っていた。実行前に消す + 開けたか/書けたかを両方見る + 失敗したら止めて理由を出す。/ 【AL3】`dl_model.py` / `probe_env.py` も固定パス。書けないと**前の版のスクリプトが残って実行される**。消してから書き、成否を記録(こちらは止めない — セットアップは入れ直せる側、probe は表示専用)。/ 【監査35】`audit_file_writes.py` を新設しランナーに登録。全書き込みを列挙し、成否確認か `// best-effort: 理由` の明示かを**必ずどちらか選ばせる**。→ この型は今後黙って増えない。★v2.9.54 で13件落ち、v2.9.55 で合格することを確認済み。★検出器ミス18回目: 窓を40行固定にしたら SaveSettings(74行)を取り切れず**正しい実装を3件 NG と誤判定**した。窓は「次の書き込みまで」に変更。
# v2.9.54: 【AK1】**書式テンプレートが AH4 の取りこぼしだった**。AH4(v2.9.51)は 延長/先行表示/結合/文節 の4つを読み切りに直したが、テンプレートは対象外のままだった。配置(Pass2)は `g_templateContent` を**生成の最後**に読み、ログは**最初**に書く。生成中はテンプレの「選択」「解除」ボタンが押せる(`SetBusy()` が無効化するのは Generate/Setup だけ)ので、変えると ①その回の字幕が**新しいテンプレ**で配置され、ログには古い方が残る(AG5 と同じ誤情報) ②「解除」を押すと `g_templateContent` が空になり **Pass2 が丸ごと走らず字幕が素のまま**になる ③UI スレッドの代入中に Pass2 がコピーすると **std::string の data race**(未定義動作・クラッシュしうる)。★①②は競合を待つまでもなく操作すれば確実に起きる。★テンプレは**全字幕の見た目を決める**ので、混ざったときの影響は linger/lead より大きい。冒頭で `tplContentEarly` / `tplPathEarly` に読み切り、ログも配置もそれを使う(AH4 と同じ型)。★スイート29 に検出器を追加。/ 【検出器の修正】スイート29 の AH4 検査は `tb[3000:]` という**固定の文字数**で冒頭と後半を分けていた。読み切りブロックの手前にコメントを数行足しただけで境界を越え、**正しい実装を NG と誤判定した**(v2.9.54 で実際に誤爆。検出器ミス16回目)。境界を lead のクランプ位置に固定するアンカー方式へ変更。★v2.9.51/52/53/54 の4版で回し、**v2.9.51 の lead 誤りを引き続き捕まえること**を確認済み。
# v2.9.53: 【AJ1】`EnsurePyHelper()` が**書けたかを見ていなかった**。is_open() だけを見て、書き込みと flush の成否を確かめていない。★ofstream は**開いた時点で truncate する**ので、失敗するとそれまで正しかった whisper_helper.py が空か途中で壊れた状態に置き換わり、誰も気づかない。次の生成で python は壊れたスクリプトを実行し、利用者には無関係なエラーが出る。★SRT 側(AD1)は `if(!f)` → `close()` → `if(!f)` と**二重に**確かめて理由まで出しているのに、こちらは素通りだった。**「同じ処理が2箇所で片方だけ」の6回目**。同じ手順に揃え、bool を返して生成側で止める。★ここで止めるのは安全側 — 書き込みが走るのは `existing != embedded` のときだけで、失敗した時点でファイルは既に壊れている。走らせても必ず失敗する。K1(生成をブロックする判定を厳しくして詰まらせた)と違い誤検知の余地が無い。★**版を上げた直後が一番踏みやすい**(内容が変わるので書き込みが走る)。/ 【AJ2】`DownloadFFmpeg()` が PowerShell のシングルクォート文字列へ**パスを生のまま埋めていた**。PowerShell では ' を '' にしないと文字列が途中で切れる。**ユーザー名にアポストロフィを含む環境**(O'Brien / D'Angelo 等)では ffmpeg の自動取得が**必ず**失敗する。日本語環境では稀だが海外の利用者は踏む(Issue #4 は海外からだった)。`PsQuote()` を新設。★PowerShell を起動しているのはこの1箇所だけなので影響範囲も1箇所。★シングルクォート内では $ もバッククォートも literal なので、脱出が要るのは ' だけ。★スイート21 に検出器を追加し、**直す前の v2.9.52 で6件落ちることを確認**してから採用した。
# v2.9.52: 【AI1】転写が失敗したときに**原因と違う場所を案内していた**のを修正。Python は転写の例外をクリップ単位で握り潰し `Clip N err: <理由>` を .err に書くだけで exit 0 で終わるため、C++ からは成功に見えて「結果が空」の分岐に落ちる。そこで必ず「音声ファイルを確認してください」と出しており、実際の原因(CUDA のメモリ不足・モデルの破損など)は .err に**あるのに一度も表示していなかった**。利用者は関係のない音声ファイルを疑わされる。`CollectClipErrors()` を新設して理由を拾い、空結果ダイアログに出す。★本実行と再試行の**両方**が同じ1関数を呼ぶ(「同じ処理が2箇所で片方だけ」を作らない)。★状態を持つので clear は関数の冒頭に固定(前回状態の残留を作らない)。★足すのは情報だけで、判定は厳しくも緩くもしない。/ 【AI2】生成中に AviUtl2 を閉じると**子の python が残る**のを修正。RunProcess は CreateProcess するだけで Job に入れておらず、Windows は親が死んでも子を殺さない。openai-whisper は tqdm がパイプへ書くため壊れたパイプに気づいて自滅するが、**faster-whisper は転写中に一切書かないので気づかず走り続ける**。孤児は GPU を掴んだまま、固定パスの `temp\whisper_results.txt` / `whisper_fw_0.wav` を次の実行と取り合う(Q1 と同じ「別の動画の字幕」を再発させうる)。`JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE` の Job を1つ作り、**転写の子だけ**を入れる。★セットアップ(pip)には掛けない — 導入の途中で殺すと AG6 の「中途半端に入った状態」を新たに作るため。★引数の既定は false なので既存の呼び出しは挙動が完全に不変。★Job 割り当てに失敗しても ResumeThread して従来どおり続行する(詰まらせない)。
# v2.9.51: 【AH4】生成中に設定を変えると**同じ1回の生成で旧い値と新しい値が混ざる**問題を修正。生成は数分かかるのに、SetBusy() が無効化するのは Generate/Setup ボタンだけで設定は触れてしまう。TranscribeThread は冒頭で大半の設定を読む一方、**後半で 延長(linger) / 先行表示(lead) / 短い字幕を結合 / 文節区切り を読み直していた**。冒頭で読み切ってローカル変数(lingerSecEarly / leadSecEarly / mgOn / mpOn)を使うようにした。範囲のガードも読んだ直後に一度だけ掛ける。★SRT エクスポート側(BuildSrtText)は**エクスポート時点の値を読むのが正しい**ので変更しない。スイート29 に「後半で読み直していないか」を検査する行を追加。★gpt-5.4 のスレッド競合調査(観点B)が発見、こちらで裏取り。
# v2.9.50: 【AH3】ProbeThread が**二重に走る**のを防いだ。起動時は `std::thread(ProbeThread).detach()` で非同期に開始し(数秒かかる)、SetupThread は末尾で `ProbeThread()` を**同期呼び出し**する。**起動直後にセットアップを押すと2つが同時に g_probe を書く**。g_probe の torch / torchUsable / rawOut は atomic ではないので、bool が片方の値で上書きされたり、std::string(rawOut) の同時書き込みで**未定義動作**になる。★**AG1(セットアップ直後だけ環境タブが「未導入」のまま生成できない)の根本原因はこれだった可能性が高い** — 症状の発生条件(セットアップ直後にだけ起きる/後から単体で試すと再現しない)が完全に一致する。排他フラグ `g_probeRunning` を compare_exchange で取り、**先行分の完了を待ってから**実行する(単に return するとセットアップ後の再実測が行われない)。どの経路で抜けてもフラグを戻すよう RunGuard を置いた。スイート29 に検査を追加。★gpt-5.4 のスレッド競合調査が発見、こちらで裏取り。
# v2.9.49: 【AH1】cuBLAS/cuDNN の導入に**強制再導入の逃げ道が無かった**のを修正。faster-whisper のチェックを付けても haveCublas が真ならそのまま skip され、「cuBLAS はあるが cuDNN が壊れている/欠けている/中断で半端」を踏むと**押し直しても永久に直らない**(AG6 と同じ袋小路)。あわせて **cuDNN 側の存在も見る**ようにした(コメントには「cuBLAS/cuDNN の DLL が要る」と書いてあるのに片方しか確認していなかった)。 【AH2】モデルの**強制再DLチェックを復活**。v2.9.1 で廃止した根拠「モデルは単なるファイルなので中身が壊れる状態が起きない」は、HuggingFace 形式の中断DL(フォルダはあるが中身が不完全)によって**覆っている**(AF1 で実測)。AF1 の判定は K1 再発を避けるため**わざと緩く**してあるので、誤検知を踏んだ利用者が自力で入れ直せる道が要る。★**判定を厳しくするのではなく逃げ道を作る**のが、K1 と AG6 の両方を避ける形。スイート12 に AH1/AH2 の検査を追加。★gpt-5.4 の横断調査が発見、こちらで裏取り。
# v2.9.48: 【AG6】**セットアップを何度押しても直らない**問題を修正(まっさら環境で実測)。PyTorch の導入判定が `import torch` が通り `cuda.is_available()` が真なら「導入済み(skip)」だったが、システム側に古い torch(NumPy 1.x 時代)が残っていると **import は通り cuda も真なのに numpy と噛み合わず実際には使えない**。結果、環境タブは「▲要再セットアップ」と正しく警告しているのに、**その指示に従ってセットアップを押しても毎回 skip され、永久に復旧できなかった**。新規利用者がセットアップを中断すると必ず踏む。判定に `torch.from_numpy(numpy.zeros(1))` を加え、**使えるかどうか**で見る。★ここは**セットアップ側**の判定なので厳しくしてよい — 結果は「入れ直す」であって生成のブロックではない(AG2 で生成側を厳しくして K1 を再発させたのとは役割が違う)。スイート12 に検査を追加。
# v2.9.47: 【AG5】ProbeThread が**前回の実測値を持ち越す**のを修正。python が見つからない等で早期 return したとき g_probe の中身をクリアしておらず、前回の値がそのまま残っていた。特に v2.9.46 で足した rawOut は、**前回の probe 結果を今回のものとして生成ログへ再掲**してしまう(デバッグのために足したのに誤情報を出す)。torchUsable も前回値のまま環境タブへ出る。実測値を**どの脱出経路より前で**全部クリアするようにした。★Y2(リセットが early return より後ろ)・AD2(_vad_mode の持ち越し)と同じ「前回状態の残留」型。**新しく状態を持ったら、まずリセット位置を決めること。** スイート29 に「最初の return より前でクリアしているか」を位置関係で検査する行を追加。★gpt-5.4 の点検で発見、こちらで裏取り。
# v2.9.46: 【AG4】probe の記録が**生成すると消えていた**のを修正。v2.9.44 で probe の結果を whisper_debug.log に書くようにしたが、probe は**起動時にしか走らない**のに対しログは**生成開始時に truncate される**ため、生成した後にログを見ると probe の記録が無い。つまり「不具合が起きた後に原因を追う」という**いちばん必要な場面で手がかりがゼロ**だった(v2.9.43 の調査で実際に詰まり、原因を特定できないまま終わった)。probe の生出力を g_probe.rawOut に保持し、truncate の直後に再掲する。スイート21 に「再掲しているか」「それが truncate より後ろにあるか」を位置関係で検査する行を追加。
# v2.9.45: 【AG3】numpy 競合のときの日本語説明が**実際には一度も出ていなかった**のを修正。エラーを出す経路は2つある — (1)結果ファイルが開けない (2)結果ファイルに ERROR 行がある — のに、v2.9.43 で日本語説明を (1) にしか置いていなかった。ところが Python は例外を捕まえて結果ファイルへ `ERROR|Unexpected error|...` を**書く**ので、実際に通るのは (2)。結果、利用者には英語の "Unexpected error" だけが出ていた(まっさら環境で実測)。両方の経路に置き、スイート21 に「対になっているか」の検査を追加。★N1(merge_seg⇔two_line)・V2(配置⇔SRT)・W1(単語割当⇔按分)・AE1(本実行⇔再試行)と同じ**同じ処理が2箇所にあって片方だけ**の型。今日4回目。
# v2.9.44: 【AG2】**v2.9.43 で K1 を再発させてしまったのを修正**(本番で実際に踏んだ)。v2.9.43 は torch の導入判定そのものを chk_torch(from_numpy まで試す)に置き換えたが、この判定は CheckMissingForGenerate 経由で**生成をブロックする**side にあり、セットアップが「導入完了」と言うのに環境タブは「未導入」のまま生成できない、という K1 と同じ袋小路になった。★TORCH(ブロック判定)は従来の chk('torch') に戻し、「実際に使えるか」は **TORCHUSABLE として別に持つ**。使うのは環境タブの表示(▲要再セットアップ)と、生成が落ちたときの日本語説明だけ。**生成は止めない** — Z2・AA1 と同じ「知らせるが止めない」方針に揃えた。 ★あわせて **probe の出力を DebugLog に必ず残す**。v2.9.43 の不具合は probe が何を返したか記録が無く、スクリプト単体では正常なのに AviUtl2 内でだけ失敗する状態で**事後に追えなかった**。スイート21 に「TORCH を厳しくしていないか」「torchUsable をブロックに使っていないか」「probe をログに残しているか」の検査を追加。
# v2.9.43: 【AG1】セットアップを PyTorch のダウンロード中に閉じると、**字幕がゼロになるのに原因が分からない**問題を修正。中断すると numpy(2.4.6)だけが新しく入った状態になり、システム側に残っている古い torch(2.2.1+cu118 = NumPy 1.x 時代)と競合して、転写の直前 whisper.load_model で "FATAL: Numpy is not available" で死ぬ。★`import torch` は**警告だけで通る**ため probe の `__import__('torch')` では検出できず、「PyTorch 導入済み」と判定されて生成が走ってしまう。probe を `torch.from_numpy(numpy.zeros(1))` まで試す chk_torch() に変更し、実際に使えるかを見るようにした。あわせて生成が落ちたときのダイアログで、英語の NumPy メッセージの前に「セットアップが最後まで終わっていない可能性/環境タブで完了まで実行してください」と日本語で出す。★まっさら環境(aviutl2_v2.1.2 テスト用)で実際に踏んだ事例から。判定は Python 側の出力を拾うだけにして C++ に同じ判定を書かない(Z2 と同じ方針)。
# v2.9.42: 【AF1】モデルの導入判定が**中断DLの残骸を完成扱い**していた問題を修正。HuggingFace は DL 開始時にキャッシュのフォルダを作るため途中で止まると `models--<配布元>--<リポジトリ>` だけが残るが、ModelExists はフォルダ名が `-<モデル名>` で終わるかしか見ておらず「導入済み」と判定していた。結果、再DLも不足表示もされないまま生成時に失敗する(**新規利用者だけが踏む**。K1/L1/M1 と同じ系統)。`snapshots\<hash>\` に実体があるかを確認する HasSnapshotFile() を追加し、**一般経路と kotoba 特例の両方**に入れた(片方だけだと N1・V2・W1・AE1 と同じ轍)。★判定は **緩く** した — snapshots に何かファイルが1つでもあればよい。`model.bin` 等に限定すると配布形式が変わったとき **K1(入っているのに未導入 → 生成がブロックされ入れ直しても直らない)** が再発するため。監査(スイート12)にも「限定していないこと」を検査する行を入れた。
# v2.9.41: 【AE1】自動インストール後の**再試行だけタイムアウトが10分固定**だった問題を修正。本実行は tmo(音声長×2+5分、上限1時間)を計算して渡しているのに、再試行は 600000 のべた書きで、長尺だと10分では足りず「導入し直したのにまた失敗する」状態になっていた。同じ tmo を使う。★N1(merge_seg⇔two_line)・V2(配置⇔SRT)・W1(単語割当⇔按分)と同じ**同じ計算が2箇所にあって片方だけ**の型。スイート21 に対の監査を追加し、固定値を渡していたら NG にした。★gpt-5.4 の横断調査が発見、こちらで裏取り。
# v2.9.40: 【AD1】SRTエクスポートが**書けていなくても「完了」と表示**していた問題を修正。ofstream を開いた後 is_open()/good() を一切見ずに無条件で成功メッセージを出しており、保存先が書き込み不可(リムーバブル抜去/権限/パス長/空き容量)だとファイルが無いのに成功と表示されていた。open 直後と close 後の2箇所で確かめ、失敗時は理由つきのエラーを出して中止する。★U1(完了時の警告が画面に一度も出ていなかった)と同じ型 —「作った」と「届いた」は別。 【AD2】幻聴フィルタが**前のクリップの状態に汚染されて誤爆**する問題を修正。_speech_runs の早期 return 2箇所がモジュールグローバル _vad_mode を更新せず、前クリップの "silero-fw" が残ると、v2.9.13 が「VADが使えない」と「VADは動いたが発話0個」を区別するために入れた安全弁 `if not runs and _vad_mode == "none"` が False になって判定を続行。runs が空なので発話の重なりが 0 と計算され、**辞書に載っている語が無条件に落ちる**(本当に喋っていても消える)。★Y2 と同じ「前回状態の残留」型。★どちらも監査を既存スイート(24/29)に対で追加し、未修正版で落ちることを対照実験で確認済み。
# v2.9.39: 【AC1】単語割当の途中で例外が出ると、**同じ区間の字幕が二重に出る**問題を修正。_push(results,...) が try の内側にあり、results へ何件か追記した後で例外が出ると return に到達せず、下の按分フォールバックが同じセグメントをもう一度まるごと出していた。出力を try の外へ移し、失敗したら組み立てごと捨てる。★N1・V2・W1 が「同じ計算が2箇所にあって片方だけ直す」型だったのに対し、これは「片方が中途半端に走ったまま両方出る」型。監査は ast で構文木を見て『取り消せない出力(_push)が try の内側にないか』を固定した(正規表現ではネストした try を見誤る)。★2026-08-03 Copilot の横断調査が発見、こちらで裏取り。
# v2.9.38: 【AB1】言語=自動 を選んでいるとき、完了時に「想定外の言語が出たら ja 固定」と伝える。効果音や短いクリップは言語判定の手がかりが無く別言語へ転び、幻聴辞書は日本語と英語しか持たないため韓国語・ロシア語に化けた幻聴が素通りする(実測 2026-08-03: 効果音3つが韓/露の幻聴になり 0 filtered。同じ素材を ja 固定で回すと3件とも辞書で落ちた)。★どの言語に転ぶか事前に分からないので辞書では先回りできない。直し方は言語固定の一本。★既定は元から ja(CB_SETCURSEL 1)なので新規導入者は影響を受けない。踏むのは自動を選んだ人だけ。★生成はブロックしない。
# v2.9.37: 【AA1】旧版が **plugin 直下** に置いた HuggingFace キャッシュ(models--*)の残骸を検出して完了時に知らせる。正しい置き場は models\ 配下で、現行版はそちらしか見ない(GetModelsDir)。両方あると同じモデルが二重に置かれ、turbo だけで 1.5GB が黙って死蔵される。開発機で実測: 直下と models\ に snapshot ハッシュまで同一のものが各 1,546.5MB あった。★消しはしない(他人のディスクを無断で消さない)。あることと消してよいことだけ伝える。★生成はブロックしない(CheckMissingForGenerate に入れると K1 と同じく止まる。Z2 と同じ判断)。
# v2.9.36: 【Z2】旧版で導入した人に「GPUがあるのにCPUで動いている」ことを伝える。v2.9.35の修正はセットアップを実行し直さないと効かないため、何もしない人はずっと遅いまま(M1ガードで落ちないので気づけない)。Python が既に出しているログ(GPU found but cuBLAS DLL missing)を拾って完了時の警告欄に『環境タブでセットアップ』と直し方まで出す。★生成はブロックしない(CheckMissingForGenerate に入れると K1 と同じく止まる。cuBLAS は高速化であって必須ではない)。
# v2.9.35: 【Z1】faster-whisper が既に導入済みだと cuBLAS/cuDNN が**永久に入らない**問題を修正。導入処理が InstallWhisperPkg の中にあり、skip の早期 return で到達しなかった(システム側のPythonに faster-whisper があると PkgOk が true になる)。★M1ガードにより落ちずにCPUへ退避するため『エラーも出ずにずっと遅い』になる。独立したステップとして外に出し、既に cuBLAS があるときは入れない(1GBの無駄を避ける)。
# v2.9.34: 【Y2】ffmpeg/python が見つからずに失敗すると、g_segs が前回の実行のまま残り、SRTエクスポートが**前回の字幕を今回の結果として書き出す**問題を修正。リセットが early return より後ろにあった。TranscribeThread の冒頭へ移動。★Q1(結果ファイルの残骸)と同じ『前回の状態が残る』型。
# v2.9.33: 【X1】幻聴辞書(hallucination_phrases.txt)を Shift-JIS で保存するとUnicodeDecodeError で黙って組み込み辞書に戻り、利用者の編集が無視されていたのを修正。日本語Windowsのメモ帳は長らく ANSI が既定。cp932 でも読み直す(読めた文字コードは log の hal_phrases= に file / file(cp932) として出る)。★このファイルは幻聴フィルタの誤爆を利用者が止める唯一の手段なので、黙って効かないのが一番まずい。
# v2.9.32: 【W1】按分経路(単語タイムスタンプが無いとき)で、直前の字幕と重なるとチャンクのテキストが**黙って消えていた**のを修正。単語割当経路は同じ場面で最低1フレーム確保して必ず出力しており、片方だけ劣っていた。★按分経路は kotoba-whisper では常に通る(単語TSが強制OFF)。重なりは実測で確認済み(実音声353秒で0.28秒)。★今日3回目の『同じ計算が2箇所にあり片方だけ直した/劣っていた』型(N1・V2に続く)。
# v2.9.31: SRT出力の2件。【V1】2行モード/文節区切りで本文に入るリテラルの \n がSRTにそのまま出て改行にならなかった(テストログも同様)。本物の改行に変換する。【V2】SRT側の先行表示(lead)に P1 の順序回復が入っておらず、長さ0のSRTエントリ(00:00:00,000 --> 00:00:00,000)ができていた。★どちらも『同じ計算が2箇所にあり片方だけ直した』型。N1と同型。
# v2.9.30: 【U1】完了時の警告が**画面に一度も出ていなかった**問題を修正。警告はステータス欄に追記されるが、その STATIC が幅316pxの1行しかなく「Done! N個の字幕を配置」だけで埋まり、右端で切り捨てられていた。★今回の再生速度だけでなく、fugashi未導入(v2.8.5)と**既存オブジェクトと重なり配置できません(v2.9.21/H1)**も同じ理由で見えていなかった。ステータスを3行に広げ、警告ごとに改行を入れ、速度の文言を短縮した。
# v2.9.29: 【T1】v2.9.26 で自分が入れた不具合を修正。_cuda_dll_ok() の docstring に torch\\lib と書いたため Python が毎回 SyntaxWarning(invalid escape sequence) を出していた(動作はするがログが汚れ、エラー表示にも混ざる。公開中のv2.9.26に入っている)。 【T2】クリップごとに [開始-終了] offset= speed= とファイル名をログに出す。実機で再生速度の警告が出なかった原因(読めていないのか値が違うのか)を切り分けるため。F1/H1 の調査でも欲しかった情報なので常設する。
# v2.9.28: 【S1】再生速度が100%でないクリップを検出して警告する。プラグインは再生速度を一度も読んでおらず、尺をタイムライン長から計算するため音声だけ伸び縮みして字幕が動かず、クリップの後ろほどズレていた(実測: 速度70で100%と同じ位置に出た)。★この版は**警告のみ**で挙動は変えない。正しい対応(尺を速度倍+時間の逆写像)は別途。0は速度指定なしとみなし警告しない。
# v2.9.27: 【N1/N2】「最大2行にまとめる」が効かない2つの経路を修正。(N1)maxchars=0 だと SplitText が分割せずグルーピングが走らないため無反応だった(Python側は maxc<=0 を20文字に読み替えており経路で不一致)。2行ON時だけ既定20で分割する。(N2)mergeLimit が maxC のままで結合後が必ず maxC 以下になり、2行の対象が生まれなかった。maxLines 倍にする。★どちらも 2行OFF のときは従来と完全に同一。この2機能は16スイートのどれも見ておらず、開発者も morph_split=1 固定で通らない経路だった。 【O1】文節区切りONでも fugashi が無ければ Python は何も分割していないのにC++が譲っていたため2行が無反応だった(fugashiMissingを見て C++ 側でグルーピングする)。 【P1】先行表示(lead)で**タイムライン先頭付近の字幕が黙って消えていた**のを修正。全itemから一律に引いて0でクランプするため複数が0に潰れ、重なり解消で長さ0になり除去されていた(実測: 先頭3本が lead=1秒で2本、5秒で1本)。開始順序を回復して1フレーム残す。
# v2.9.26: 【L1】PyTorch のインデックス選択がドライバの CUDA 版だけで決まっており、Python 版数によっては存在しないホイールを指して pip が必ず失敗していた問題を修正(実測 win_amd64: cu118/cu121=cp312まで / cu124=cp313まで / cu126・cu128=cp314まで)。torch専用の PickTorchCudaTag() を新設し、ドライバ上限とPython版数の両方を満たさなければCPU版へ退避して理由を表示する。★DetectCudaTag() 自体は変更しない(faster-whisper の cuBLAS/cuDNN 要否判定にも使われており、そこを cpu 扱いにすると転写が cublas64_12.dll 不在で途中死して字幕ゼロになるため)。あわせて復旧経路で pip の失敗を見ていなかったのを修正。
# v2.9.25: テストログ(testlog_dir.txt 指定時)を固定名 whisper_testlog.txt の上書きに変更。従来は日時つきの名前で生成のたびに新規作成され、削除も世代管理も無いため配布先で溜まり続けていた。
# v2.9.24: 【K1/重大】faster-whisper で入れたモデルが「未導入」と誤判定され字幕生成が完全にブロックされる問題を修正(利用者report)。HuggingFaceキャッシュ形式 models--<配布元>--<リポジトリ> を見る処理が kotoba の特例にしか無かった。全モデルで後方一致により検出する。
# v2.9.23: 【J1】シンボリックリンク権限が無い環境(WinError 1314)でモデルDLが失敗する問題を修正。1314はPermissionErrorにならずhuggingface_hubのフォールバックに入らないため、踏んだときだけコピー経路を強制して再試行する。【J2】ffmpegのDLに $ProgressPreference の抑止・1回リトライ・タイムアウト明示を追加。
# v2.9.22: 【I1】幻聴フィルタが openai 経路にしか無く、faster-whisper を選ぶと一切動いていなかった問題を修正(実機ログで faster 6本すべて0件/openai 5本すべて1件を確認)。hallucination_phrases.txt の編集も faster では無効だった。両経路を対称化。
# v2.9.21: 【H1/重大】Pass2 が利用者の既存オブジェクトを削除しうる経路を修正。Pass1 の成否を per-item で記録し、Pass2 は自分が置いた item だけを削除/再作成するようにした。あわせてレイヤー探索の枯渇を記録し、配置失敗を完了メッセージに出すようにした。
# v2.9.20: クリップindexの位置結合バグを修正(音声なしクリップの後ろに本編があると字幕が全滅)。JSONにidxを明示して往復。あわせて同一クリップ(ファイル+区間+再生位置が完全一致)の重複転写を排除。
# v2.9.19: 字幕本文に実際の改行が入ると結果行/エイリアスが壊れて字幕が消える経路を塞いだ(プロンプトの反響で入りうる)。出力を _push() に集約し LF/CR を空白へ落とす。2行表示のリテラル\nはそのまま。
# v2.9.18: iniの保存/復元でプロンプトが壊れる問題を修正。バックスラッシュをエスケープしていなかったため、「C:\new」のような文字列が復元時に改行へ化けていた(実測8ケース中4ケース)。IniEscape/IniUnescape に集約し hotwords も同じ扱いに揃えた。
# v2.9.17: フェーズ別の所要時間をログに追加(抽出/Python/モデルロード/転写)。効率化を推測でやらないための計測。機能の変更は無し。
# v2.9.16: JSONエスケープの不備を修正。prompt はタブ等の制御文字を素通し、hotwords は" と \ しか見ておらず、貼り付けでタブが混入すると batch.json が壊れて字幕生成が丸ごと失敗していた。JsonEscape() に集約し4箇所すべてで使用。
# v2.9.15: 監査④。音声トラックの無い動画でも落ちないことを確認(ffmpegが-22で止まり正しく検出される)。ただしメッセージが「ffmpegを確認してください」で誤誘導だったため、原因別に出し分けるようにした。
# v2.9.14: 監査③(C++)の修正4件。プロジェクトfpsの無検証(ゼロ除算クラッシュ)、生成タイムアウトの固定600秒(長尺/CPUで全部失う)を音声長比例に、beam_sizeとtemperatureの上限追加。
# v2.9.13: 監査②の修正2件。無音ファイルで幻聴が残る問題(VAD不可と発話ゼロの取り違え)と、セグメント/ポーズ分割パートをまたいだ字幕の重なりを修正。
# v2.9.12: ★文節区切りON時に原文の空白が全部消えるバグを修正(英語字幕が単語連結で壊れていた)。形態素解析は空白をトークンにしないため表層形の連結で落ちていた。原文を走査して拾い直す。
# v2.9.11: 監査①(設定×バックエンド)の修正4件。ホットワードをopenai選択時にグレーアウト、単語タイムスタンプの強制ONをUIに反映、SRTエクスポートの生成中ガード、ffmpegDLボタンの二重実行防止。
# v2.9.10: 実機報告2件の修正。(1)「そういう」等が語の途中で割れる問題を文節規則で修正。(2)チャンクが1つのとき単語時刻もVADスナップも使わず即returnしていた問題を修正し、さらに単語間0.6秒以上の間でセグメントを分割(whisperがポーズをまたいで統合すると字幕が早く出るため)。
# v2.9.9: 配布対応。テストログを既定OFF(testlog_dir.txt がある場合のみ出力)にし、幻聴フレーズ辞書を hallucination_phrases.txt で差し替え・無効化できるようにした。
# v2.9.8: ★品質フィルタのしきい値を -1.0 から -3.0 に緩和。実測で正しい字幕が logprob=-1.029 で捨てられていた(ゴミは-5.865)。あわせてテストログ出力先を testlog_dir.txt で指定可能にし、kotoba-whisper のラベル表示切れを修正。
# v2.9.7: 温度フォールバックの上限を 1.0 から 0.4 に変更。実機で温度1.0まで上がった回にデコードが崩壊し(英語のゴミを出力・logprob=-5.865)、15〜45秒の本物の字幕が丸ごと失われたため。ループ脱出の能力は 0.4 までで残す。
# v2.9.6: 字幕生成完了時に「設定+実行情報+SRT」をデスクトップの SRTテストログ フォルダへ自動保存する機能を追加。毎回手でSRTエクスポートしなくても実行間の比較ができる。あわせて beam/device/VAD 等をログに記録するようにした(従来は残っていなかった)。
# v2.9.5: v2.9.4のエンベロープ方式を撤回(本物の字幕を3件消す事故)。幻聴判定はフレーズ辞書+VAD重なり+フレーズ占有率の3条件のみにし、辞書に無い語は絶対に落ちない設計に戻した。あわせて品質フィルタが捨てた内容をログに出すようにした(切り分け用)。
# v2.9.4: 幻聴除去を作り直し。v2.9.3 の hallucination_silence_threshold は幻聴を消さず 別の定型句に化けさせるだけ(実機で「ご視聴ありがとうございました」→「それではまた」)だったので撤去。判定を「発話エンベロープの外側 かつ VAD重なりが薄い」主体に変更し、文言列挙に依存しないようにした。
# v2.9.3: openai-whisper 側にも hallucination_silence_threshold=2.0 を渡すよう修正。faster-whisper にだけ渡していたため、whisper 選択時のみ無音区間に 「Thank you」「ご視聴ありがとうございました」等の幻聴字幕が生成されていた。
#
# v2.9.2: (1)Batched単独で字幕生成が失敗する不具合を修正。BatchedInferencePipelineは音声をVADで区切って
# バッチを作るためVADが無効だと成立しない。v2.8までvad_filterがTrue固定で隠れていたが、v2.9で既定OFFに
# したことで表面化した(v2.9で入れた回帰)。use_batched時はvad_filterを強制True、ログにNOTEを出す。
# ※Batchedを使うとVADが必ず効くので認識語数が減りタイミングも早くずれる=速度と引き換えに精度が落ちる。
# (2)精度タブの「Batched」ラベルが幅70で切れていたので2行目のVAD隣(x=174,幅100)へ移動。
# (3)Batchedもfaster-whisper専用(openai-whisper側はbatchedを一度も参照しない)なので、
# VAD・繰返し抑制と同様にbackend=whisper選択時はグレーアウトするよう統一。
#
# v2.9.1: ヒント文/ホットワードに日本語を入力すると字幕生成が失敗する不具合を修正(Issue #5)。
# GetWindowTextA等のANSI版APIはシステムコードページ(CP932)のバイト列を返すため、日本語がCP932のまま
# iniとwhisper_batch.jsonに書かれ、Python側のjson.load(encoding="utf-8")が
# 「'utf-8' codec can't decode byte 0x94」で失敗していた。GetWindowTextW+WideToUtf8()に変更(計6箇所:
# ini保存2/batch.json書き込み2/ini読込2)。数値専用欄はASCIIのみなので意図的に未変更。
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
# ビルド: .\whisper_subtitle_v2_9_75.ps1 [出力ディレクトリ]
# 要件: Visual Studio 2022, CMake 3.15+
# 出力: whisper_subtitle_v2_9_75.aux2 (既存の whisper_subtitle_v2_8_10.aux2 / whisper_subtitle.aux2 は上書きしない)
#
# Note: [v2.8.8] PyTorchのCUDA版はnvidia-smiのCUDA Versionから自動判定(cu121/cu118/cpu)して導入
##############################################################################

$d = if($args.Count -gt 0){ $args[0] } else { [Environment]::GetFolderPath("Desktop") }
# ★ビルド作業フォルダは ASCII パスのままにすること(日本語パスだと cl.exe が文字化けする)。
#   ここを「◯sdk whisper開発場所」配下に移してはいけない。
$projDir = "$d\aviutl2_dev\whisper_subtitle_plugin_v2_9_75"
$src = "$projDir\src"
# ★2026-08-04 開発者指示: ビルドログはデスクトップを圧迫するので開発フォルダへ出す。
#   ログはビルドに使われないので日本語パスでも安全(cl.exe は触らない)。
$logDir = "C:\Users\test\Desktop\AVIUTL２開発関連\◯sdk whisper開発場所\whisper build log"
if(-not (Test-Path $logDir)){ New-Item -ItemType Directory $logDir -Force | Out-Null }
$logFile = "$logDir\whisper_build_log_v2_9_75.txt"
$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

"" | Out-File $logFile -Encoding UTF8
"Whisper Subtitle v2.9.75 Build $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File $logFile -Append -Encoding UTF8

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " Whisper Subtitle v2.9.75 - Build" -ForegroundColor Cyan
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
static HWND g_chkFfmpeg = 0, g_chkTorch = 0, g_chkWhisper = 0, g_chkFaster = 0, g_chkFugashi = 0;
// v2.9.49【AH2】モデルの強制再DLチェックを **復活** させる。
// v2.9.1 で廃止した根拠は「モデルは単なるファイルなので torch のような
// 『import は通るが中身が壊れている』状態が起きない」だったが、**これは誤りだった**。
// HuggingFace 形式は DL 開始時にフォルダを作るため、中断すると
// 「フォルダはあるが中身が不完全」という、まさにその状態になる(AF1 で実測)。
// AF1 の判定は K1 再発を避けるため **わざと緩く**してある(snapshots に何か1つあれば導入済み)。
// 緩い判定には誤検知が残るので、**利用者が自分で入れ直せる逃げ道**が要る。
// これが無いと、誤判定を踏んだ利用者は Setup を押しても毎回 skip され復旧できない(AG6 と同じ袋小路)。
static HWND g_chkModel = 0;
static HWND g_torchStatusLabel = 0, g_modelStatusLabel = 0; // v2.9.0【H/G】: PyTorch行・モデル行の状態表示

// v2.9.0【A】: 起動時に python を1回だけ回して導入状況をまとめて実測しキャッシュする。
// 表示専用。SetupThread の skip 判定には使わない(実行時点の実態を見るべきなので従来の PkgOk を維持)。
struct PkgProbe {
    std::atomic<bool> done{false};   // 実測完了したか
    std::atomic<bool> pyFound{false};// python 自体が見つかったか
    int  pyMajor = 0, pyMinor = 0, pyPatch = 0;
    bool torch = false, torchCuda = false;
    // v2.9.44【AG2】「import できる」と「実際に使える」を **別のフラグ**にする。
    // ★v2.9.43 で torch の判定そのものを厳しくしたら、**K1 が再発した**
    //   (セットアップは「完了」と言うのに生成側が「PyTorch が必要」と言い続けて抜けられない)。
    //   生成をブロックする判定(CheckMissingForGenerate)は **従来どおり torch のまま**にして、
    //   使えるかどうかは torchUsable として持ち、**表示と警告にだけ使う**。
    //   Z2・AA1 と同じ「知らせるが止めない」方針に揃える。
    bool torchUsable = false;
    // v2.9.46【AG4】probe の**生出力**を保持する。
    // ★v2.9.44 で probe をログに書くようにしたが、**生成開始時にログが truncate される**ため、
    //   生成すると probe の記録は消えてしまい、「不具合が起きた後にログを見る」という
    //   肝心の場面で残っていなかった(2026-08-04 に実測)。生成のたびに再掲する。
    std::string rawOut;
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
static int g_cublasMissing = 0; // v2.9.36【Z2】GPUはあるが cuBLAS が無くCPUへ退避した
static int g_extractNoAudio = 0; // v2.9.15: 音声トラックが無くて抽出できなかったクリップ数

// v2.9.16: JSON文字列のエスケープ。従来は prompt / hotwords / ffmpegパス / wavパス の
// 4箇所にそれぞれ別々に書かれていて、抜けている文字が箇所ごとに違った。
//   prompt   : タブ等の制御文字を素通し
//   hotwords : " と \ しか見ていない
// Python の json.load は文字列中の生の制御文字を拒否するため(実測: タブで
// "Invalid control character")、混入すると batch.json が壊れて
// **字幕生成が丸ごと失敗**する。Excel等から貼り付けるとタブが入るので現実に起きうる。
// UTF-8のマルチバイト(0x80以上)はそのまま通す。
// v2.9.18: ini は行単位の key=value なので、値の改行を畳む必要がある。
// 従来は改行だけ畳んで **バックスラッシュ自体をエスケープしていなかった**ため、
// プロンプトに "C:\\new" のような文字列があると復元時に改行へ化けて壊れていた
// (実測: 8ケース中4ケースで往復が壊れる)。プロンプトは再起動のたびに復元されるので
// 毎回壊れ、しかも認識に効くパラメータなので静かに精度が落ちる。
static std::string IniEscape(const std::string& v){
    std::string o; o.reserve(v.size() + 8);
    for(char c : v){
        if(c == '\\')      o += "\\\\";
        else if(c == '\n') o += "\\n";
        else if(c == '\r') {}
        else               o += c;
    }
    return o;
}
static std::string IniUnescape(const std::string& v){
    std::string o; o.reserve(v.size());
    for(size_t i = 0; i < v.size(); i++){
        if(v[i] == '\\' && i + 1 < v.size()){
            if(v[i+1] == 'n'){  o += '\n'; i++; continue; }
            if(v[i+1] == '\\'){ o += '\\'; i++; continue; }
        }
        o += v[i];
    }
    return o;
}

static std::string JsonEscape(const std::string& in){
    std::string o;
    o.reserve(in.size() + 8);
    for(unsigned char c : in){
        switch(c){
            case '"':  o += "\\\""; break;
            case '\\': o += "\\\\"; break;
            case '\n': o += "\\n";  break;
            case '\r': o += "\\r";  break;
            case '\t': o += "\\t";  break;
            case '\b': o += "\\b";  break;
            case '\f': o += "\\f";  break;
            default:
                if(c < 0x20){
                    char b[8]; sprintf_s(b, "\\u%04x", (unsigned)c);
                    o += b;
                } else {
                    o += (char)c;
                }
        }
    }
    return o;
}

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
// v2.9.37【AA1】旧版は HuggingFace のキャッシュ(models--*)を **plugin 直下** に作っていた。
// 現行版は GetModelsDir() の下しか見ないため、直下のものは二度と読まれずに残り続ける。
// turbo だけで 1.5GB あり、models\ 側と snapshot ハッシュまで同一のものが二重に置かれる。
// ★数えるだけ。**消さない**(他人のディスクを無断で消さない。削除は利用者の判断)。
static int CountStrayModelDirs(){
    WIN32_FIND_DATAW fd;
    std::wstring pat = Utf8ToWide(GetPluginDir() + "\\models--*");
    HANDLE h = FindFirstFileW(pat.c_str(), &fd);
    if(h == INVALID_HANDLE_VALUE) return 0;
    int n = 0;
    do { if(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) n++; } while(FindNextFileW(h, &fd));
    FindClose(h);
    return n;
}

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
    // best-effort: ログなので失敗しても無視する(本処理を止める理由がない)
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
import sys, json, os, traceback, glob, re

# Add local site-packages (whisper_subtitle/site-packages)
_script_dir = os.path.dirname(os.path.abspath(__file__))
_local_sp = os.path.join(_script_dir, "site-packages")
if os.path.isdir(_local_sp) and _local_sp not in sys.path:
    sys.path.insert(0, _local_sp)
# Also ensure system site-packages is in path
_sys_sp = os.path.join(os.path.dirname(sys.executable), "Lib", "site-packages")
if os.path.isdir(_sys_sp) and _sys_sp not in sys.path:
    sys.path.append(_sys_sp)

# v2.9.67【AZ1】アライメント用の依存を**あらゆる import より前に**パスへ入れる。
# ★遅らせてはいけない理由(実機経路でだけ再現した):
#   _speech_runs が faster-whisper 経由で **グローバルの huggingface_hub(新しすぎる版)**を
#   先に sys.modules に載せてしまい、あとから読む transformers が非互換で落ちる。
#   同じ理由で HF_HOME も後から設定しても効かない(hub は import 時にキャッシュ位置を決める)。
# ★入れる位置は**同梱 site-packages の直後**。末尾に足すとグローバルが先に来て負ける。
# ★align-packages が無ければ何もしない = 従来どおりの動作。
_align_dir_boot = os.path.join(_script_dir, "align-packages")
if os.path.isdir(_align_dir_boot):
    if _align_dir_boot not in sys.path:
        sys.path.insert(1 if _local_sp in sys.path else 0, _align_dir_boot)
    _align_models_boot = os.path.join(_script_dir, "align-models")
    os.environ.setdefault("HF_HOME", _align_models_boot)
    os.environ.setdefault("HF_HUB_CACHE", os.path.join(_align_models_boot, "hub"))

def _cuda_dll_ok():
    """CTranslate2 が CUDA 実行に要求する cuBLAS の DLL が実在するか。

    ★v2.9.26【M1-guard】構築時ではなく**転写の途中**で要求されるため、無いまま cuda を
      選ぶと CPU フォールバックが効かず**字幕がゼロになる**(v2.9.6 のコメント参照)。
      「自動」で cuda を選ぶ前に必ずこれを通す。
    ★探す場所は add_cuda_paths() と同じに保つこと(片方だけ増やすとズレる)。
      実測: 開発機には nvidia/* が無く、torch/lib/cublas64_12.dll だけがあった。  # ※\\l は不正なエスケープなのでパスは/で書く
      つまり openai 用 torch の DLL を借りて faster-whisper が動いていた。
      torch を入れない faster-whisper 専用構成では nvidia-cublas-cu12 が要る。
    """
    sites = [_local_sp] if os.path.isdir(_local_sp) else []
    sys_site = os.path.join(os.path.dirname(sys.executable), "Lib", "site-packages")
    if os.path.isdir(sys_site):
        sites.append(sys_site)
    pats = []
    for site in sites:
        pats += [os.path.join(site, "torch", "lib", "cublas64_*.dll"),
                 os.path.join(site, "nvidia", "**", "cublas64_*.dll"),
                 os.path.join(site, "nvidia*", "**", "cublas64_*.dll")]
    for d in os.environ.get("PATH", "").split(os.pathsep):
        if d.strip():
            pats.append(os.path.join(d, "cublas64_*.dll"))
    for p in pats:
        try:
            if glob.glob(p, recursive=True):
                return True
        except Exception:
            pass
    return False

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
    # v2.9.26【M1】「自動」が torch でしか GPU を見ておらず、torch を使わない
    # faster-whisper(CTranslate2) では GPU があっても必ず CPU になっていた。
    # device の既定はコンボ先頭の「自動」なので **新規ユーザーは全員これを踏む**。
    # しかも README は「faster-whisper は PyTorch 不要」と勧めており、その構成ほど損をする。
    # 判定はバックエンドごとに正しい手段で行う(実測: ctranslate2 4.7.1 の
    # get_cuda_device_count() は RTX3080 で 1 を返す)。失敗時は従来どおり cpu。
    # cuda を選んでも _run_faster_whisper 側に float16->int8->cpu のフォールバックがある。
    # ★backend を読んだ後でないと分岐できないのでここに置く(add_cuda_paths より前)。
    if device == "auto":
        device = "cpu"
        if backend == "whisper":
            try:
                import torch
                if torch.cuda.is_available(): device = "cuda"
            except Exception: pass
        else:
            # ★GPUの存在だけで cuda にしないこと。cuBLAS の DLL が無いと
            #   転写の途中で落ち、CPUフォールバックも効かず字幕がゼロになる。
            try:
                import ctranslate2
                if ctranslate2.get_cuda_device_count() > 0:
                    add_cuda_paths()          # 先に DLL の在り処を PATH へ通してから確認する
                    if _cuda_dll_ok():
                        device = "cuda"
                    else:
                        with open(err_path, "a", encoding="utf-8") as ef:
                            ef.write("GPU found but cuBLAS DLL missing -> cpu"
                                     " (nvidia-cublas-cu12 / nvidia-cudnn-cu12 が未導入)\n")
            except Exception: pass
        with open(err_path, "a", encoding="utf-8") as ef:
            ef.write(f"device=auto resolved to {device} ({backend})\n")
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

)PYHELPER" R"PYHELPER(
# v2.9.65【AX1】文節に「寄りかかる語」を巻き込むための材料。
# 単独では意味を成さず前の語に寄りかかる成分が、**行の先頭に落ちる**のを防ぐ。
# 実機で開発者が挙げた違和感はすべてこの型だった:
#   「勉強して|きた」「向いて|もらっても」「そうじゃ|ないのか」(補助動詞/補助形容詞)
#   「静か|すぎるな」「おかし|すぎるなと」(程度の接尾)  「真似|する」「失礼|しても」(サ変)
#   「もったいない|ことを」(形式名詞)
# ★形式名詞は品詞が名詞なので、直前が動詞/形容詞/助動詞(=連体修飾を受けている)ときだけ。
#   そうしないと「事が起きた」の「事」まで巻き込む。
_KEISHIKI = ("こと", "もの", "ため", "はず", "わけ",
             "とき", "ところ", "つもり", "ほう",
             "よう", "せい", "くせ", "かぎり",
             "うえ", "あいだ", "かわり")

def _is_ja(ch):
    """日本語の文字(かな/漢字/半角カナ)か。空白を落としてよいかの判定に使う。"""
    o = ord(ch)
    return (0x3040 <= o <= 0x30FF) or (0x4E00 <= o <= 0x9FFF) or (0xFF66 <= o <= 0xFF9F)

# v2.9.75【BH1】文末記号のあとの空白が残っていた。
# 実測(2026-08-24 実機/89件中2件): 「おはよう! いってきます!」「こんにちは! また明日ね」。
# BA1(v2.9.68)は「両隣が日本語の文字」でしか空白を落とさないので、
# 片側が ! ? 。 、 だと素通りしていた。**同じ型の残り**。
# ★英文は守る: 少なくとも片側が**本物の日本語文字**であることを要求する。
#   "OK! Yes" は 次が Y なので落とさない。
_JA_PUNC = "\u3001\u3002\uff01\uff1f\u0021\u003f\u2026\u30fb"

def _is_ja_side(ch):
    """空白の隣として『日本語側』とみなしてよいか(文末記号を含む)。"""
    return _is_ja(ch) or (ch in _JA_PUNC)

def _brk_score(b):
    """その文節の直後で切る自然さ。句読点の後 > 助詞の後 > その他。"""
    t = b.rstrip()
    if not t:
        return 1
    if t[-1] in "、。！？!?…":
        return 3
    if t[-1] in "はがをにでともへやねよか":
        return 2
    return 1

def _morph_group(text, maxc):
    """text を文節境界で maxc 文字以内のチャンクに分割。fugashi 未導入等は [text] を返す(安全)。"""
    try:
        tg = _morph_get_tagger()
    except Exception:
        return [text]
    # 文節へ: 自立語で新文節を開始、付属語(助詞/助動詞/記号/接尾辞)は前にくっつける
    # v2.9.12: 形態素解析は空白をトークンとして返さないため、表層形を素朴に連結すると
    # 原文の空白が全部消える(実測: "Hello world this is" -> "Helloworldthisis")。
    # 原文を走査して、形態素の間に落ちた文字を拾い直す。
    bun = []
    bonus = []   # v2.9.68【BA1】その文節の直後で切る優先度(空白があった所)
    cur = ""
    prev = None
    pos = 0
    for w in tg(text):
        sfc = w.surface
        idx = text.find(sfc, pos) if sfc else pos
        if idx < 0:
            idx = pos          # 見つからない(正規化された等)ときは諦めて連結だけする
            gap = ""
        else:
            gap = text[pos:idx]  # 形態素の間に落ちた文字 = 主に空白
        pos = idx + len(sfc)
        p1 = w.feature.pos1
        p2 = str(getattr(w.feature, "pos2", "") or "")
        attach = (not cur) or p1 in ("助詞", "助動詞", "補助記号", "記号", "接尾辞")
        # v2.9.65【AX1】非自立可能な動詞/形容詞は、文節の途中なら前にくっつける
        if not attach and p2 == "非自立可能" and p1 in ("動詞", "形容詞"):
            attach = True
        # v2.9.65【AX1】形式名詞は、連体修飾を受けているときだけ前にくっつける
        if (not attach) and p1 == "名詞" and sfc in _KEISHIKI                 and prev is not None and prev.feature.pos1 in ("動詞", "形容詞", "助動詞"):
            attach = True
        # v2.9.10: 「そういう/こういう/どういう/ああいう/っていう」を語の途中で割らない。
        # unidic は そう(副詞) + いう(動詞,一般) と分けるため「いう」で新文節が始まり、
        # 実機で「ある文」/「その続き」と切れていた。
        # 直前が副詞・助詞で、今の語の語彙素が「言う」なら繋げる。
        # ★v2.9.69【BB1】**かな表記のときだけ**に限定した。効きすぎていて、
        #   「あのって|言わないでください」まで1文節にまとめてしまい、15文字の塊になって
        #   C++ の文字ベース分割に渡り、**「くだ|さいって」と割れた**(実機で発生)。
        #   かな「いう」= そういう/っていう(繋げたい) / 漢字「言わ・言っ」= 本物の動詞(割ってよい)。
        if not attach and prev is not None:
            lemma = str(getattr(w.feature, "lemma", "") or "")
            if lemma == "\u8a00\u3046" and prev.feature.pos1 in ("\u526f\u8a5e", "\u52a9\u8a5e") \
                    and not any(0x4E00 <= ord(_c) <= 0x9FFF for _c in sfc):
                attach = True
        # v2.9.68【BA1】whisper が句の切れ目に入れる空白を落とす。
        # 日本語は分かち書きしないので、**両隣が日本語なら空白は表示上のゴミ**。
        # 実測(開発者素材): 35件中5件(14%)に空白。全部「和+空白+和」だった。
        #   例)「きょうは いい天気」「あしたの 天気」「ちょっと まってね」
        # ★落とすだけで終わらせない。**whisper がそこを句の切れ目と見た証拠**なので、
        #   改行位置の候補として加点する(_brk_score より優先)。
        # ★英数字どうしの間(英語)は残す。C++ 側の結合処理と同じ規則。
        _sp_hint = False
        if gap and gap.strip() == "":
            _pc = cur[-1] if cur else (bun[-1][-1] if bun else "")
            _nc = sfc[0] if sfc else ""
            # v2.9.75【BH1】片側が文末記号でも落とす。ただし**必ず片方は本物の日本語**
            if _pc and _nc and _is_ja_side(_pc) and _is_ja_side(_nc) and (_is_ja(_pc) or _is_ja(_nc)):
                gap = ""
                _sp_hint = True
        if attach:
            cur = (cur + gap + sfc) if cur else sfc   # 先頭の空白は捨てる
        else:
            bun.append(cur + gap)   # 空白は前の文節の末尾に付ける
            bonus.append(3 if _sp_hint else 0)
            cur = sfc
        prev = w
    if cur or pos < len(text):
        bun.append(cur + text[pos:])   # 末尾に残った文字(空白等)も拾う
        bonus.append(0)
    if maxc <= 0:
        maxc = 20
    _keep = [(b, bonus[i] if i < len(bonus) else 0) for i, b in enumerate(bun) if b]
    bun = [x[0] for x in _keep]
    bonus = [x[1] for x in _keep]
    if not bun:
        return [text] if text else []
    # v2.9.65【AX2】貪欲詰め込み → **均等割り**。
    # 従来は maxc まで目一杯詰めたので、2つ目が余り物になっていた。
    #   17字/maxc14: 「そろそろ帰りたくなって」(13) + 「きたけど」(4)
    # 人は均等に割る:「そろそろ」(6) + 「帰りたくなってきたけど」(11)
    # 個数は最小のまま、長さのばらつきが最小になる割り方を DP で選ぶ。
    # 切れ目の自然さ(_brk_score)も加点するので、句読点や助詞の後が選ばれやすい。
    L = [len(b) for b in bun]
    total = sum(L)
    n = len(bun)
    if total <= maxc or n == 1:
        c = "".join(bun).strip()
        return [c] if c else ([text] if text else [])
    if n > 400:
        # ★保険: 文節が異常に多いときは DP をやめて従来の貪欲に落ちる。
        #   繰り返しループの残骸などで n が膨らんでも計算量が暴れないようにする。
        chunks, cur = [], ""
        for b in bun:
            if cur and len(cur) + len(b) > maxc:
                chunks.append(cur.rstrip()); cur = b.lstrip()
            else:
                cur += b
        if cur:
            chunks.append(cur.rstrip())
        chunks = [c for c in chunks if c]
        return chunks if chunks else ([text] if text else [])
    INF = float("inf")
    kmin = (total + maxc - 1) // maxc
    if kmin < 1:
        kmin = 1
    found = None
    for k in range(kmin, n + 1):
        target = total / float(k)
        dp = [[INF] * (k + 1) for _ in range(n + 1)]
        bk = [[-1] * (k + 1) for _ in range(n + 1)]
        dp[0][0] = 0.0
        for i in range(1, n + 1):
            jmax = k if k < i else i
            for j in range(1, jmax + 1):
                acc = 0
                m = i
                while m >= 1:
                    acc += L[m - 1]
                    if acc > maxc and (i - m + 1) > 1:
                        break   # 上限超過。ただし1文節だけなら割れないので許す
                    if dp[m - 1][j - 1] < INF:
                        c = dp[m - 1][j - 1] + (acc - target) * (acc - target)
                        if i < n:
                            # v2.9.68【BA1】空白があった所は whisper が句の切れ目と見た場所
                            _sc = _brk_score(bun[i - 1])
                            if bonus[i - 1] > _sc:
                                _sc = bonus[i - 1]
                            c += (3 - _sc) * 4.0
                        if c < dp[i][j]:
                            dp[i][j] = c
                            bk[i][j] = m - 1
                    m -= 1
        if dp[n][k] < INF:
            found = (k, bk)
            break
    if found is None:
        return [text] if text else []
    k, bk = found
    out, i, j = [], n, k
    while j > 0:
        m = bk[i][j]
        out.append("".join(bun[m:i]))
        i, j = m, j - 1
    out.reverse()
    # チャンク境界 = 改行位置なので、端の空白は残さない
    chunks = [c.strip() for c in out]
    chunks = [c for c in chunks if c]
    if chunks:
        return chunks
    return [text] if text else []

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
            # v2.9.40【AD2】前クリップの値を持ち越さない。理由は下の except を参照。
            _vad_mode = "none"
            return []
        x = np.frombuffer(raw, dtype=np.int16).astype(np.float32) / 32768.0
        if ch > 1:
            x = x.reshape(-1, ch).mean(axis=1)
    except Exception:
        # v2.9.40【AD2】ここと上の早期 return は _vad_mode を更新していなかった。
        # _vad_mode はモジュールグローバルなので、前のクリップで "silero-fw" になっていると
        # このクリップで VAD が動かなくても古い値が残る。すると _is_hallucination の
        #   if not runs and _vad_mode == "none": return False
        # (v2.9.13 が「VADが使えない」と「VADは動いたが発話0個」を区別するために入れた安全弁)
        # が **False になって判定を続行** し、runs が空なので発話の重なりが 0 と計算され、
        # **辞書に載っている語が無条件に落ちる**。本当に喋っていても消える。
        # ★Y2(リセットが early return より後ろ)と同じ「前回状態の残留」型。
        _vad_mode = "none"
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

# v2.9.3: whisper が無音区間で高確信度に生成する定型句 (学習データ由来の幻聴)。
# no_speech_prob も avg_logprob も正常値になるため既存の品質フィルタでは原理的に捕まらない。
#   実測値: 「次回予告」 no_speech_prob=0.000 / avg_logprob=-0.347 / VAD重なり=0%
# v2.9.5: whisper が無音区間で高確信度に生成する定型句 (学習データ由来の幻聴)。
# 実測: 開発者の動画で 45.55秒の幻聴は 7回中7回、下のどれかで必ず出た。
# no_speech_prob も avg_logprob も正常値になるため既存の品質フィルタでは捕まらない。
#   実測: 「次回予告」 no_speech_prob=0.000 / avg_logprob=-0.347
# ここに載っていない語は絶対に落ちない = 本物の字幕が消える事故が起きない設計。
_HAL_PHRASES_DEFAULT = (
    "ご視聴ありがとう", "ご視聴いただきありがとう", "ご覧いただきありがとう",
    "チャンネル登録", "次回予告", "ではまた", "またお会いしましょう",
    # v2.9.15: 別の動画の末尾に8回中6回出た(開発者が「言っていない」と確認済み)。
    # 日常会話でも使う語なので、実際に喋った回は VAD重なりとフレーズ占有率が守りになる。
    # これまで hallucination_phrases.txt にしか無く、ファイルを消すと復活する状態だった。
    "またね",
    "thankyou", "thanksforwatching", "wellberightback", "amaraorg",
)

def _hal_norm(t):
    """照合専用の正規化。英数字とかな漢字だけ残して記号や空白を落とす。
    ★出力される字幕テキストには一切手を加えない。ピリオドもアポストロフィも
      字幕にはそのまま残る (英語圏の利用者の表記を変えてしまわないため)。"""
    return re.sub(r"[^0-9a-z\u3040-\u30ff\u4e00-\u9fff]", "", t.lower())

def _load_hal_phrases():
    """v2.9.9: 辞書は whisper_subtitle/hallucination_phrases.txt で差し替えられる。
    1行1フレーズ、# 始まりはコメント。ファイルが無ければ上の既定値を使う。
    ★空ファイルを置けば幻聴フィルタを丸ごと無効化できる。
      配布先で誤爆したとき、リビルドを待たずに利用者自身が止められるようにするため。"""
    global _HAL_SRC
    path = os.path.join(_script_dir, "hallucination_phrases.txt")
    # v2.9.33【X1】従来は utf-8-sig 決め打ちで、失敗すると黙って組み込み辞書へ戻っていた。
    # 日本語 Windows のメモ帳は長らく ANSI(Shift-JIS) が既定だったため、利用者が
    # その形式で保存すると UnicodeDecodeError になり **編集が無視される**。
    # このファイルは「幻聴フィルタが誤爆したとき利用者が自分で止める」唯一の手段なので、
    # 黙って効かないのが一番まずい。cp932 でも読み直す。
    if os.path.exists(path):
        for _enc in ("utf-8-sig", "cp932"):
            try:
                out = []
                with open(path, encoding=_enc) as f:
                    for ln in f:
                        ln = ln.strip()
                        if ln and not ln.startswith("#"):
                            n = _hal_norm(ln)
                            if n:
                                out.append(n)
                _HAL_SRC = "file" if _enc == "utf-8-sig" else "file(cp932)"
                return tuple(out)
            except UnicodeDecodeError:
                continue
            except Exception:
                break
    _HAL_SRC = "builtin"
    return _HAL_PHRASES_DEFAULT

_HAL_SRC = "builtin"
_HAL_PHRASES = _load_hal_phrases()

def _is_hallucination(text, start, end, runs):
    """幻聴セグメントか判定する。落とす条件は3つすべてを満たす場合のみ。

      (1) 既知の定型幻聴フレーズを含む
      (2) そのフレーズが文の半分以上を占める
          "Thank you" 単独は落とすが "Thank you so much for watching everyone"
          のような実発話には反応しない (英語圏での誤爆防止)
      (3) 実際には発話されていない (VAD重なりが5割未満)
          本当に「ご視聴ありがとうございました」と喋った場合は重なりが濃いので残る

    v2.9.4 で試した「発話エンベロープの外側なら落とす」方式は撤回した。
    辞書に無い語まで落とせる代わりに、VADが後半の孤立した発話を拾えないと
    本物の字幕をまとめて消してしまい、実機で実際に3件消えた。
    VADが使えない環境 (runs が空) では判定しない = 従来どおりの挙動。
    """
    # v2.9.13: 「VADが使えない(none)」と「VADは動いたが発話が0個」を区別する。
    # 後者は無音ファイルなどで起きる、幻聴が確実な状況。実測で完全な無音20秒と
    # 極小ノイズ20秒のどちらでも "Thank you." が生成され、ここで見逃していた。
    if not runs and _vad_mode == "none":
        return False
    n = _hal_norm(text)
    if not n:
        return False
    hit = 0
    for p in _HAL_PHRASES:
        if p in n and len(p) > hit:
            hit = len(p)
    if hit == 0 or hit / len(n) < 0.5:
        return False
    dur = max(end - start, 1e-6)
    ov = 0.0
    for a, b in runs:
        ov += max(0.0, min(end, b) - max(start, a))
    return (ov / dur) < 0.5

# v2.9.66【AY1】字幕が「音より早く出すぎる」のを抑えるための材料。
# 実測(実データ228件): 開始時点が無音の字幕が129件あり、音が立つまで
#   中央値170ms / 最大1003ms 待たされていた。300ms以上が14件。
# 「字幕は少し早めに出てよい」が、1秒早いのは違和感になる。上限を設ける。
# ★VAD区間(_speech_runs)は粗すぎて使えない(76秒に17区間しかなく、多くの字幕は
#   区間の内側に落ちるため「最も近い発話開始」が数秒先になる)。10ms刻みの包絡で見る。
_MAX_LEAD = 0.15   # 字幕が音より早く出てよい上限(秒)。ここを変えれば効き方が変わる

def _speech_env(path, hop=0.010):
    """波形の短時間エネルギーとしきい値を返す。失敗したら None (機能を素通りさせる)。"""
    try:
        import wave
        import numpy as np
        with wave.open(path, "rb") as wf:
            sr = wf.getframerate(); ch = wf.getnchannels(); sw = wf.getsampwidth()
            raw = wf.readframes(wf.getnframes())
        if sw != 2 or sr <= 0:
            return None
        x = np.frombuffer(raw, dtype=np.int16).astype(np.float32) / 32768.0
        if ch > 1:
            x = x.reshape(-1, ch).mean(axis=1)
        hn = int(sr * hop)
        if hn <= 0:
            return None
        nw = len(x) // hn
        if nw < 10:
            return None
        e = np.sqrt(np.mean(x[:nw * hn].reshape(nw, hn) ** 2, axis=1))
        floor = float(np.percentile(e, 20)); peak = float(np.percentile(e, 95))
        if peak <= floor:
            return None
        return (e, hop, floor + (peak - floor) * 0.15)
    except Exception:
        return None

# v2.9.75【BE1】whisper の**セグメント開始が遅れて、直前の発話を丸ごと取りこぼす**型。
# 実測(開発者素材): 「ある区間」で 6.470-6.850 の 380ms の発話が
# どのセグメントにも属さず、字幕は 7.240 から出ていた(**770ms 遅い**)。
# 開発者の指摘「それでなんか がかなり遅く字幕が出てる」。
# ★アライメントでは直せない。CTC は **whisper が決めた区間の中しか探せない**ので、
#   区間の開始が遅れていればその前には置けない。
#   ★窓を前に1秒広げる案は試して**失敗**した。CTC は窓の左端から並べ始めるだけで、
#     「別の区間」が 4.780→3.780 と1秒早まって悪化した(立ち上がりまで 280→1280ms)。
# ★VAD区間(_speech_runs)も使えない。silero は 5.06-6.98 と前の発話ごと1本にまとめ、
#   さらに「なんか」(7.81-7.90)を**丸ごと見落として**いた。10ms刻みの包絡でないと見えない。
# ★動かす向きは「早める」だけ。かつ**無音から始まっている字幕にしか触らない**ので、
#   合っているものには当たらない(実測23セグメント中、動いたのは指摘の1件だけ)。
_ORPHAN_MAX_GAP = 1.0     # これ以上離れた発話は拾わない
# ★隙間だけを見ていると、**長い発話区間**に当たったときにその先頭まで戻ってしまう
#   (監査 A11 が検出: 2.86s の字幕が 1.48s まで 1.38秒も戻った)。移動量自体も縛る。
_ORPHAN_MAX_MOVE = 1.2    # 早める量の上限
_ORPHAN_MIN_LEN = 0.15    # これより短い音は拾わない(息継ぎ/物音)
_ORPHAN_HANG = 0.15       # 発話のかたまりを繋ぐ間合い

def _env_runs(env):
    """エネルギー包絡から発話のかたまりを作る。VAD より細かい。"""
    if not env:
        return []
    e, hop, th = env
    out = []
    i = 0
    n = len(e)
    hg = int(_ORPHAN_HANG / hop) or 1
    while i < n:
        if e[i] > th:
            j = i
            while j < n and (e[j] > th or (j + hg < n and max(e[j:j + hg]) > th)):
                j += 1
            if (j - i) * hop >= _ORPHAN_MIN_LEN:
                out.append((i * hop, j * hop))
            i = j
        else:
            i += 1
    return out


def _reclaim_orphan(t, prev_end, rr, env):
    """t の直前に誰も拾っていない発話があれば、そこまで**早める**。無ければ t のまま。"""
    if not rr or not env:
        return t
    e, hop, th = env
    # ★1コマだけ見ると取り逃す。立ち上がりの直前に置かれていることがあるので窓で見る
    #  (これが無いと「立ち上がり直前の区間」= 立ち上がり6ms手前 まで動かしてしまった)。
    lo = max(0, int((t - 0.05) / hop))
    hi = min(len(e), int((t + 0.15) / hop) + 1)
    if lo < hi and max(e[lo:hi]) > th:
        return t
    best = None
    for a, b in rr:
        if b > t + 1e-6:
            continue
        if b < t - _ORPHAN_MAX_GAP:
            continue
        if a < prev_end - 1e-6:
            continue
        if t - a > _ORPHAN_MAX_MOVE:
            continue
        if best is None or a < best:
            best = a
    if best is None or best >= t:
        return t
    return best


def _cap_lead(t, end, env):
    """字幕開始 t が無音から始まるなら、音が立つ _MAX_LEAD 秒前まで**遅らせる**。

    ★動かす向きは「遅らせる」だけ。早める方向には一切動かさない。
    ★開始時点で既に音が出ているなら触らない(そこは正しく合っている)。
    ★遅らせた結果 0.1秒未満になるなら触らない(字幕が潰れるほうが害が大きい)。
    """
    if not env:
        return t
    e, hop, th = env
    n = len(e)
    i = int(t / hop)
    if i < 0:
        i = 0
    if i >= n or e[i] > th:
        return t
    j = i
    while j < n and e[j] <= th:
        j += 1
    if j >= n:
        return t          # 以降ずっと無音 = 判断材料が無い
    onset = j * hop
    if onset - t <= _MAX_LEAD:
        return t
    ns = onset - _MAX_LEAD
    if end - ns < 0.1:
        return t
    return ns

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

# v2.9.10: 単語間がこの秒数以上あいていたらセグメントを分ける。
# 根拠: 実音声33セグメントの単語間ギャップは最大0.38秒(0.4秒以上は0件)だった。
# 0.6秒なら通常の息継ぎ・読点では発火せず、既存の切れ方を変えない。
# 発火数は whisper_debug.log の Morph timing 行 pause-split= で確認できる。
_PAUSE_GAP = 0.6

)PYHELPER" R"PYHELPER(
# ===== v2.9.67【AZ1】強制アライメント (WhisperX のアライメント段だけを使う) =====
# 目的: 単語タイムスタンプの精度。whisper は「何と言ったか」と「いつ言ったか」を同時に
#       推測するので時刻が粗い。**文字列が確定した状態で位置だけ探す**強制アライメントは
#       問題が簡単になるぶん精度が出る。
# 実測(実データ39セグメント / 発話の立ち上がりからの誤差):
#       いまの時刻 中央220ms / 50ms以内 5.1%  →  アライメント 中央17ms / 50ms以内 71.8%
# ★認識テキストは一切変えない。差し替えるのは**単語の時刻だけ**。
# ★依存は align-packages に隔離してある。無ければ黙って従来動作に落ちる。
_ALIGN_DIR = os.path.join(_script_dir, "align-packages")
_ALIGN_MODELS = os.path.join(_script_dir, "align-models")
_align_cache = None      # PER_RUN: (alignment_module, model, metadata, device, lang)

def _align_load(lang, device, err_path=None):
    """アライメント段を読み込む。失敗したら None(従来動作)。同じ言語なら使い回す。

    ★失敗の理由を必ず記録する。ここを黙って None にすると、
      「aligned=0 だが理由が分からない」状態になり切り分け不能になる(実際に踏んだ)。
    """
    global _align_cache
    if _align_cache is not None and _align_cache[4] == lang:
        return _align_cache
    def _note(msg):
        if err_path:
            try:
                with open(err_path, "a", encoding="utf-8") as ef:
                    ef.write("Align: " + msg + "\n")
            except Exception:
                pass
    if not lang:
        _note("language unknown -> skip")
        return None
    if not os.path.isdir(_ALIGN_DIR):
        _note("align-packages not found at " + _ALIGN_DIR)
        return None
    try:
        # ★同梱 site-packages の**直後**に入れる。末尾に足すとグローバル環境が先に来て
        #   古い tokenizers 等が勝つ(実測で踏んだ)。
        if _ALIGN_DIR not in sys.path:
            sys.path.insert(1, _ALIGN_DIR)
        os.environ.setdefault("HF_HOME", _ALIGN_MODELS)
        # whisperx/__init__ は話者分離(pyannote)を読むので、**__init__ を実行せずに**
        # パッケージだけ登録して alignment だけを取り込む。pyannote は同梱していない。
        import types as _t
        if "whisperx" not in sys.modules:
            _pkg = _t.ModuleType("whisperx")
            _pkg.__path__ = [os.path.join(_ALIGN_DIR, "whisperx")]
            sys.modules["whisperx"] = _pkg
        import whisperx.alignment as _A
        if lang not in _A.DEFAULT_ALIGN_MODELS_HF and lang not in _A.DEFAULT_ALIGN_MODELS_TORCH:
            _note("no alignment model for language '%s'" % lang)
            return None
        m, meta = _A.load_align_model(language_code=lang, device=device)
        _align_cache = (_A, m, meta, device, lang)
        return _align_cache
    except Exception as e:
        _note("load failed: %s: %s" % (type(e).__name__, e))
        return None

def _align_reattach(text, words):
    """アライメントが落とした文字(句読点など)を戻し、**連結 == text** を回復する。

    ★これが無いと _split_by_pause の一致チェックが永久に通らなくなり、ポーズ分割が黙って死ぬ
      (実測: 202セグメント中148件が不一致になっていた)。
    ★時刻は隣の単語から引き継ぐので精度は落ちない。
    ★追跡できなければ None(呼び側は元の words を使う = 安全側)。
    """
    out, ti = [], 0
    for w in words:
        wt = w.get("word") or ""
        if not wt:
            continue
        k = text.find(wt, ti)
        if k < 0:
            return None
        gap = text[ti:k]
        if gap:
            if out:
                out[-1][0] += gap
            else:
                wt = gap + wt
        out.append([wt, w.get("start"), w.get("end")])
        ti = k + len(w.get("word") or "")
    if not out:
        return None
    if ti < len(text):
        out[-1][0] += text[ti:]
    if "".join(x[0] for x in out) != text:
        return None
    return out

# v2.9.75【BD1】CTC アライメントが**崩壊したセグメント**の内部を割り直す。
# ★崩壊の指紋は「20ms(wav2vec2 CTC の最小1コマ)の単語が並ぶ」こと。
#   時刻ではなく、行き場を失った文字を最小幅で詰めた**詰め物**である。
#   実測(開発者素材/整列できた23セグメント)で7件が該当。
#   例:「ある行のテキスト」が13文字に 0.583秒 = **秒速22文字**。人間には不可能。
#   同じセグメントの「次の行」は 2.534s に置かれたが、whisper 自身の単語時刻(DTW)は 3.000s。
#   開発者の指摘「前の行…は遅い / 次の行は早すぎる」と両方一致した。
# ★先頭は動かさない。CTC の**セグメント先頭**は実測で優秀(エネルギー基準 220ms→17ms)。
#   壊れているのは内部の境界だけなので、そこだけ文字数で割り直す。
# ★健全なセグメントには触らない。実測では CTC の内部境界に比例配分より優位性は
#   見られなかったが(中央値同着 / 勝敗 2対3)、**n=5 と小さいので判断を急がない**。
# ★実測(判定 30% または3連続): 内部の切れ目が DTW から 中央117ms→67ms、
#   物理的に不可能な速さの字幕が 7件→3件。
_ALIGN_MIN_FRAME = 0.021   # これ以下 = 最小1コマに潰された印
_ALIGN_DEGEN_FRAC = 0.30   # 単語の何割が潰れていたら崩壊とみなすか
_ALIGN_DEGEN_RUN = 3       # 連続でこれだけ潰れていたら崩壊とみなす

def _align_undegen(words, seg_end):
    """崩壊していれば内部を割り直した単語列を返す。健全なら None(=触らない)。"""
    if len(words) < 3:
        return None
    n_min = 0
    run = best = 0
    for w in words:
        if (w[2] - w[1]) <= _ALIGN_MIN_FRAME:
            n_min += 1
            run += 1
            if run > best:
                best = run
        else:
            run = 0
    if not (n_min >= len(words) * _ALIGN_DEGEN_FRAC or best >= _ALIGN_DEGEN_RUN):
        return None
    t0 = words[0][1]
    span = float(seg_end) - t0
    total = 0
    for w in words:
        total += len(w[0])
    if span <= 0 or total <= 0:
        return None
    out = []
    acc = 0
    for w in words:
        s0 = t0 + span * (acc / total)
        acc += len(w[0])
        out.append((w[0], s0, t0 + span * (acc / total)))
    return out


# v2.9.75【BF1】whisperx が返すセグメントの**数が合わない**ことがある。
# 実測(別素材6本): `11 != 10` / `28 != 27` / `6 != 5` と**1個多く**返ってきた。
# whisperx はアライメント中にセグメントを**分割することがある**ため。
# ★これまでは数が違うだけで **そのクリップの整列を丸ごと捨てて**いた。
#   6本中2本(g3/g5)が全滅していた。壊れはしない(従来動作に落ちる)が、
#   素材によって効いたり効かなかったりする**最大の原因**がこれだった。
# ★セグメントの区切り方には依存させない。**単語の並び順**だけを頼りに割り当て直す。
# ★数が合っているときは**この関数を通さない**(実績のある経路を変えない)。

def _align_regroup(segs, items):
    """whisperx が返した単語を、こちらが渡した items に割り当て直す。
       返り値 {items の添字: [単語辞書, ...]}。対応づけできなかった添字は入れない。"""
    flat = []
    for sg in segs:
        for w in (sg.get("words") or []):
            if w.get("start") is None or w.get("end") is None:
                continue
            if not (w.get("word") or ""):
                continue
            flat.append(w)
    out = {}
    if not flat:
        return out
    j = 0
    for i, it in enumerate(items):
        t = it[2] or ""
        if not t or j >= len(flat):
            continue
        ti = 0
        k = j
        while k < len(flat):
            wt = flat[k].get("word") or ""
            p = t.find(wt, ti)
            if p < 0:
                break
            ti = p + len(wt)
            k += 1
            if ti >= len(t):
                break
        if k > j and _align_reattach(t, flat[j:k]) is not None:
            out[i] = flat[j:k]
            j = k
        elif k > j:
            # 対応づけに失敗。次の item を巻き添えにしないよう並びだけ進める
            j = k
    return out


def _align_words(items, wav_path, lang, device, err_path, stats=None):
    """items = [(開始秒, 終了秒, 本文), ...] を強制アライメントし
       {添字: [(単語, 開始, 終了), ...]} を返す。**失敗した添字は入れない**(呼び側が元を使う)。"""
    out = {}
    if not items:
        return out
    # ★初回はアライメントモデル(約1.2GB)の取得が走る。進捗が出ないので、
    #   何が起きているかをログに残さないと**フリーズに見える**。
    _first = _align_cache is None or _align_cache[4] != lang
    if _first:
        with open(err_path, "a", encoding="utf-8") as ef:
            ef.write("Align: loading model for '%s' (first use may download ~1.2GB)\n" % lang)
    got = _align_load(lang, device, err_path)
    if got is None:
        if _first:
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write("Align: not available for '%s' -> using whisper timestamps\n" % lang)
        return out
    if _first:
        with open(err_path, "a", encoding="utf-8") as ef:
            ef.write("Align: model ready (%s)\n" % lang)
    _A, model, meta, dev, _lg = got
    try:
        import wave as _w
        import numpy as _np
        with _w.open(wav_path, "rb") as wf:
            sr = wf.getframerate(); ch = wf.getnchannels(); sw = wf.getsampwidth()
            raw = wf.readframes(wf.getnframes())
        if sw != 2 or sr != 16000:
            return out
        audio = _np.frombuffer(raw, dtype=_np.int16).astype(_np.float32) / 32768.0
        if ch > 1:
            audio = audio.reshape(-1, ch).mean(axis=1)
        tr = [{"start": float(a), "end": float(b), "text": t} for (a, b, t) in items]
        res = _A.align(tr, model, meta, audio, dev, return_char_alignments=False)
        segs = res.get("segments") or []
        if len(segs) != len(items):
            # v2.9.75【BF1】捨てずに、単語の並びから割り当て直す
            _rg = _align_regroup(segs, items)
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write("Align: segment count mismatch %d != %d -> regrouped %d/%d\n"
                         % (len(segs), len(items), len(_rg), len(items)))
            if not _rg:
                return out
            segs = [{"words": _rg.get(_i) or []} for _i in range(len(items))]
        for i, (sg, it) in enumerate(zip(segs, items)):
            t = it[2]
            ws = [w for w in (sg.get("words") or [])
                  if w.get("start") is not None and w.get("end") is not None]
            if not ws:
                continue
            fixed = _align_reattach(t, ws)
            if fixed is None:
                continue
            out[i] = [(x[0], float(x[1]), float(x[2])) for x in fixed]
            _rep = _align_undegen(out[i], it[1])
            if _rep is not None:
                out[i] = _rep
                if stats is not None:
                    stats["degen"] = stats.get("degen", 0) + 1
        if stats is not None:
            stats["align"] = stats.get("align", 0) + len(out)
    except Exception as e:
        with open(err_path, "a", encoding="utf-8") as ef:
            ef.write("Align failed (falling back to whisper timestamps): %s\n" % e)
        return {}
    return out


)PYHELPER" R"PYHELPER(
# v2.9.63【AU2】引き伸ばし字の暴走した連続を、**分割にかける前に**詰める。
# whisper は叫び声や歌を「ー」の巨大連続として転写することがある(同一素材5回で1回、
# 最長444文字)。これを下流の SplitText が刻むと字幕が数百件に膨張する。
# ★AU1(C++の禁則フォールバック)で1文字化は止まるが、それでも14文字の「ーーーー…」が
#   30件ほど並ぶ。件数の問題ではなく**読めない**ので、元を減らすのはここでやる。
# ★消すのではなく詰める。A/B案(出力後に捨てる)と違い、字幕が丸ごと消えることはない。
_STRETCH_KEEP = 3   # ここまでは残す。「あー」「あーー」「あーーー」は一切変わらない

def _is_squashable(ch):
    """詰めてよい文字か。**かな(ひらがな/カタカナ/長音記号)と波ダッシュ**だけ。

    v2.9.64【AV1】v2.9.63 は「ー〜～っッ」だけを見ていたが、**繰り返しループの綴りは
    毎回変わる**。同じ音声を5回転写した実測:
        run1「ぉ」x222 / run2「お」x444・「え」x75 / run3「ー」x444 / run4「え」x10 / run5「ー」x5
    **5回中2回しか拾えていなかった**。実機では「おぉぉぉ…」(小書きの ぉ)が
    14文字x16件・13.7秒ぶん並んだ。「ぉ」は禁則文字でもないので AU1 も効かない。
    → 特定の記号ではなく **かな全般の同一文字連続** を対象にする。

    ★数字と英字は入れない。「1000000」は 0 が6連続するので**数値が壊れる**。
    ★漢字も入れない。実データで「熱」x9 が出たが 14文字に収まるので実害が無く、
      正当な重ね字を壊す危険のほうが大きい。**効果が薄いところまで広げない。**
    """
    o = ord(ch)
    return (0x3041 <= o <= 0x309F) or (0x30A0 <= o <= 0x30FF) or ch in "〜～"

def _squash_run(s, state):
    """同じ文字が _STRETCH_KEEP 個を超えて続くぶんを落とす(対象は _is_squashable のみ)。
    state=[直前の文字, 連続数] を外から渡すのは、**単語をまたいで数を引き継ぐ**ため。
    引き継がないと、単語境界で連続が切れたことになって詰まらない。"""
    out = []
    for ch in s:
        if ch == state[0] and _is_squashable(ch):
            state[1] += 1
            if state[1] > _STRETCH_KEEP:
                continue
        else:
            state[0] = ch
            state[1] = 1 if _is_squashable(ch) else 0
        out.append(ch)
    return "".join(out)

def _squash_stretch(text, words, stats=None):
    """text と words を同じ規則で詰めて返す。words が無い/text と食い違うときは text だけ。

    ★落とした単語の時間は**直前の残った単語の終わりに足す**。
      ここを捨てると、4秒の叫びが 0.3秒の点滅になる(単語を消すと時間も消えるため)。
    ★同じ状態機械を text と words に同じ順で通すので、詰めた後も連結は一致する。
      よって _split_by_pause の一致チェックは通ったままになる。
    """
    if not text:
        return text, words
    nt = _squash_run(text, [None, 0])
    if nt == text:
        return text, words
    if stats is not None:
        stats["squash"] = stats.get("squash", 0) + (len(text) - len(nt))
    if not words:
        return nt, words
    if "".join(w for (w, _, _) in words).strip() != text.strip():
        return nt, words   # 食い違うなら words は触らない(安全側)
    st = [None, 0]
    nw = []
    for (w, a, b) in words:
        s2 = _squash_run(w, st)
        if s2:
            nw.append([s2, a, b])
        elif nw and b is not None and (nw[-1][2] is None or b > nw[-1][2]):
            nw[-1][2] = b   # 落とした単語の持ち時間を直前へ渡す
    if not nw:
        return nt, words
    return nt, [tuple(x) for x in nw]

def _split_by_pause(text, words, sf, ef, tl_start, fps, stats):
    """whisper が間をまたいで2つの発話を1セグメントにまとめたとき、そこで分ける。

    まとめられると後半の字幕が実際に喋る前から表示される(実機で約1.3秒早く出た)。
    words の連結が text と一致することは実測で確認済みだが、
    **一致しない場合は分割しない**(安全側。テキストを壊すより早く出る方がまし)。
    """
    if not words:
        return [(text, words, sf, ef)]
    ws = [(w, s, e) for (w, s, e) in words if s is not None and e is not None]
    if len(ws) < 2:
        return [(text, words, sf, ef)]
    if "".join(w for w, _, _ in ws).strip() != text.strip():
        return [(text, words, sf, ef)]
    groups = [[ws[0]]]
    for prev, nxt in zip(ws, ws[1:]):
        if (nxt[1] - prev[2]) >= _PAUSE_GAP:
            groups.append([nxt])
        else:
            groups[-1].append(nxt)
    if len(groups) <= 1:
        return [(text, words, sf, ef)]
    out = []
    for g in groups:
        gt = "".join(w for w, _, _ in g).strip()
        if not gt:
            continue
        gs = tl_start + int(g[0][1] * fps)
        ge = tl_start + int(g[-1][2] * fps)
        if ge <= gs:
            ge = gs + 1
        out.append((gt, g, gs, ge))
    if len(out) <= 1:
        return [(text, words, sf, ef)]
    if stats is not None:
        stats["pause"] = stats.get("pause", 0) + (len(out) - 1)
    return out

def _last_end(results, ci):
    """同じクリップで直前に出力した字幕の終了フレーム。無ければ None。

    v2.9.13: _emit_part 内の単調クランプ(prev_e)は そのパート内でしか効かず、
    セグメント間や ポーズ分割のパート間で字幕が重なっていた
    (実測: 実音声353秒で 7092 -> 7075 = 0.28秒の重複)。
    C++側の重なり解消は「前の字幕の終わりを次の開始まで縮める」方式のため、
    縮めた結果が長さ0以下になると前の字幕が黙って消える。発生源で防ぐ。
    """
    if not results:
        return None
    p = results[-1].split("|", 3)
    if len(p) == 4 and p[0] == str(ci):
        try:
            return int(p[2])
        except ValueError:
            pass
    return None

def _push(results, ci, s0, e0, text):
    """結果を1行として積む。**実際の改行は必ず落とす**。

    v2.9.19: 結果ファイル(`ci|sf|ef|text`)も .object のエイリアスも行単位なので、
    本文に LF/CR が入ると
      - 結果行が2行に割れ、C++側のパース(パイプ3つを探す)で弾かれて字幕が丸ごと消える
      - エイリアスの「テキスト=」行が途中で切れてオブジェクト生成が壊れる
    実データ52セグメントでは0件だったが、**プロンプトの反響**(whisper が initial_prompt を
    そのまま出力する既知の失敗)で入りうる。プロンプト欄は複数行入力なので現実的。
    2行表示で使うのはリテラルの \\n (バックスラッシュ+n, 2文字)なので、そちらは触らない。
    """
    t = text.replace("\r\n", " ").replace("\r", " ").replace("\n", " ")
    if not t.strip():
        return
    results.append(f"{ci}|{s0}|{e0}|{t}")

def _emit_seg(results, ci, sf, ef, text, morph_split, maxchars, words=None, tl_start=0, fps=60, stats=None, runs=None, max_lines=1, env=None):
    """seg を results に追加。morph_split 時は文節分割して複数追加。
    v2.9.10: 先に大きなポーズでセグメントを分けてから、各パートを _emit_part に渡す。
    v2.9.63【AU2】文節区切りの ON/OFF に関わらず先に引き伸ばし字を詰める。
    膨張は C++ の SplitText で起きるので、OFF のときも同じ事故を踏む。"""
    text, words = _squash_stretch(text, words, stats)
    if not (morph_split and text):
        _push(results, ci, sf, ef, text) # v2.9.19
        return
    for pt, pw, ps, pe in _split_by_pause(text, words, sf, ef, tl_start, fps, stats):
        _emit_part(results, ci, ps, pe, pt, maxchars, pw, tl_start, fps, stats, runs, max_lines, env)

def _emit_part(results, ci, sf, ef, text, maxchars, words=None, tl_start=0, fps=60, stats=None, runs=None, max_lines=1, env=None):
    """v2.8.2a: words(単語タイムスタンプ)があれば実測時刻で割当。無ければ文字数比で按分。
    v2.8.5: max_lines>1 の場合、chunks を max_lines 個ずつ束ねて1テロップ(\\n連結)にする。
    v2.9.10: チャンクが1つでも単語時刻とVADスナップを通す。以前はここで即 return しており、
    短いセグメントだけ whisper のセグメント境界がそのまま採用され開始が早まっていた。"""
    if not text:
        return
    chunks = _morph_group(text, maxchars)
    n = len(chunks)
    # 1) 単語タイムスタンプによる実測割当 (v2.8.2a)
    # v2.9.39【AC1】results への追記は try の外でまとめて行う。理由は下の except を参照。
    emit = None
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
                    # v2.9.66【AY1】字幕が音より早く出すぎるのを抑える(遅らせる方向のみ)
                    if env:
                        cp = []
                        for cs_s, ce_s, ch in out:
                            ns = _cap_lead(cs_s, ce_s, env)
                            if ns > cs_s and stats is not None:
                                stats["lead"] = stats.get("lead", 0) + 1
                            cp.append((ns, ce_s, ch))
                        out = cp
                    out = _group_spans(out, max_lines)  # v2.8.5
                    prev_e = _last_end(results, ci)  # v2.9.13: セグメント/パートをまたいでクランプ
                    # v2.9.39【AC1】ここでは results に触らず、出すものを組み立てるだけにする
                    tmp = []
                    for cs_s, ce_s, ch in out:
                        cs = tl_start + int(cs_s * fps)
                        ce = tl_start + int(ce_s * fps)
                        if prev_e is not None and cs < prev_e:
                            cs = prev_e
                        if ce <= cs:
                            ce = cs + 1
                        prev_e = ce
                        if ch:
                            tmp.append((cs, ce, ch))
                    emit = tmp
    except Exception:
        # v2.9.39【AC1】従来は _push が try の内側にあり、results(呼び出し元と共有の
        # 可変リスト)へ**何件か追記した後**で例外が出ると、return に到達しないまま
        # 下の按分へ落ちて**同じセグメントをもう一度まるごと出していた**。
        # 同じ区間の字幕が二重に残る。N1・V2・W1 が「片方だけ直す」型だったのに対し、
        # これは「片方が中途半端に走ったまま両方出る」型。
        # 出力を try の外へ出し、失敗したら組み立てごと捨てる。
        emit = None
    # ★空リストでも「単語割当は成功した」なので按分へ落とさない。
    #   従来も out が真なら1件も push しないまま return していた。挙動を変えないための is not None。
    if emit is not None:
        for cs, ce, ch in emit:
            _push(results, ci, cs, ce, ch) # v2.9.19
        if stats is not None:
            stats["word"] = stats.get("word", 0) + 1
        return
    # 2) フォールバック: 文字数比の按分 (旧・等分割から改良)
    span = ef - sf
    total_c = sum(len(c) for c in chunks) or 1
    acc = 0
    # v2.9.13: 按分経路も直前の字幕と重ならないように開始を押し出す
    _le = _last_end(results, ci)
    prev = sf if (_le is None or sf >= _le) else _le
    spans2 = []
    for i, ch in enumerate(chunks):
        acc += len(ch)
        ce = ef if i == n - 1 else sf + span * acc // total_c
        cs = prev
        # v2.9.32【W1】従来は `if ce > cs` で、直前の字幕と重なると **chunk ごと捨てて**
        # テキストが黙って消えていた。単語割当経路は同じ場面で
        # `if ce <= cs: ce = cs + 1` として最低1フレーム確保し**必ず出力**している。
        # 同じ計算が2箇所にあって片方だけ劣っていた形。単語割当側に揃える。
        # ★按分経路は kotoba-whisper では常に通る(単語TSが強制OFFのため)。
        #   重なりは実在する(_last_end のコメント: 実音声353秒で0.28秒の重複を実測)。
        #   文節区切りでチャンクが細かいと、重なりぶんだけ先頭が落ちていた。
        if ce <= cs:
            ce = cs + 1
        if ch:
            spans2.append((cs, ce, ch))
            prev = ce
    spans2 = _group_spans(spans2, max_lines)  # v2.8.5
    for cs, ce, ch in spans2:
        _push(results, ci, cs, ce, ch) # v2.9.19
    if stats is not None:
        stats["prop"] = stats.get("prop", 0) + 1

)PYHELPER" R"PYHELPER(
def _is_symlink_privilege_error(e):
    """Windows の WinError 1314 (ERROR_PRIVILEGE_NOT_HELD) か判定する。

    v2.9.23【J1】シンボリックリンクの作成には開発者モードか管理者権限が要る。
    権限が無いと 1314 が出るが、Python ではこれが **PermissionError にならず素の OSError**
    になる(実測)。huggingface_hub の _create_symlink は FileExistsError と PermissionError
    しか捕まえないため、コピーへのフォールバックに入らず例外が上がってしまう。
    """
    if getattr(e, "winerror", None) == 1314:
        return True
    return "1314" in str(e)

def _force_hf_copy_instead_of_symlink():
    """huggingface_hub にシンボリックリンクではなくコピーを使わせる。

    v2.9.23【J1】are_symlinks_supported を False 固定にすると _create_symlink が
    コピー経路に入る。**常時これを有効にはしない**: コピーになると blob が二重化して
    ディスクを倍食う(large-v3-turbo で約1.5GB → 約3GB)。1314 を踏んだ環境だけ払う。
    """
    try:
        from huggingface_hub import file_download as _hfd
        _hfd.are_symlinks_supported = lambda *a, **k: False
        return True
    except Exception:
        return False

)PYHELPER" R"PYHELPER(
def _run_faster_whisper(model_size, language, device, clips, model_dir, output_path, err_path, beam_size=5, temperature=0, batch=None):
    import time as _time
    _t0 = _time.time()
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
    # v2.9.23【J1】シンボリックリンク権限が無い環境(WinError 1314)を踏んだら、
    # huggingface_hub にコピー経路を強制して1回だけ全体を再試行する。
    _sym_retry_done = False
    for _attempt in range(2):
        _sym_err = None
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
                    if _is_symlink_privilege_error(e):
                        _sym_err = e
                    with open(err_path, "a", encoding="utf-8") as ef:
                        ef.write(f"CUDA {ct} fail: {e}\n")
        if model is None:
            actual_device = "cpu"
            try:
                model = WhisperModel(model_path, device="cpu", compute_type="int8", **kwargs)
            except Exception as e:
                if _is_symlink_privilege_error(e):
                    _sym_err = e
                else:
                    with open(output_path, "w", encoding="utf-8") as f:
                        f.write(f"ERROR|Model load failed|{e}\n")
                    with open(err_path, "a", encoding="utf-8") as ef:
                        ef.write(f"Model load error: {e}\n{traceback.format_exc()}")
                    return
        if model is not None:
            break
        if _sym_err is not None and not _sym_retry_done and _force_hf_copy_instead_of_symlink():
            _sym_retry_done = True
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write("Symlink privilege missing (WinError 1314). "
                         "Retrying with copy instead of symlink.\n")
            continue
        # 1314 以外、または再試行しても駄目だった
        _msg = _sym_err if _sym_err is not None else "unknown"
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(f"ERROR|Model load failed|{_msg}\n")
        with open(err_path, "a", encoding="utf-8") as ef:
            ef.write(f"Model load error: {_msg}\n{traceback.format_exc()}")
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
    _tLoad = _time.time()  # v2.9.17: ここまでがモデルロード
    # v2.8: Read new parameters from batch
    no_prev_text = batch.get("no_prev_text", False)
    word_timestamps = batch.get("word_timestamps", True)
    rep_penalty = batch.get("rep_penalty", False)
    # v2.9.5: 従来 vad_filter=True ハードコード。実測でONだと単語が40個消え開始時刻が早い側にズレる
    # (OFFにするとopenai-whisperとタイミング完全一致)。既定OFF、UIのチェックで切替。
    vad_filter = batch.get("vad_filter", False)
    # v2.9.2【Batchedのエラー修正】BatchedInferencePipeline は音声をVADで区切ってバッチを作るため、
    # VADが無効だとバッチの単位を作れずエラーになる。v2.8までは vad_filter が True 固定だったので
    # 常にVADが効いており問題が表面化しなかったが、v2.9でVADを既定OFFにしたことで
    # 「Batchedだけチェックすると失敗し、VADと併用すると成功する」という症状が出ていた(開発者報告)。
    # Batched使用時はVADを強制的に有効にする。
    # ※VADをONにすると認識語数が減りタイミングも早くずれる副作用があるため、
    #   速度より精度を優先するなら Batched は使わない方がよい(README に明記)。
    if use_batched and not vad_filter:
        vad_filter = True
        with open(err_path, "a", encoding="utf-8") as ef:
            ef.write("NOTE: Batched requires VAD -> vad_filter forced to True\n")
    prompt_text = batch.get("prompt", "")
    hotwords_text = batch.get("hotwords", "")
    morph_split = batch.get("morph_split", False)
)PYHELPER" R"PYHELPER(
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
        # v2.9.7: 上限を 1.0 から 0.4 に下げた。
        # 温度0はビーム探索(決定的)だが、0より大きいとサンプリングになる。
        # 0.6以上は粗すぎて日本語では実質ゴミしか出ず、復帰する見込みが薄い。
        # 実機で1.0まで上がった回に「dead」「moderator」「excited」という英語のゴミが出て
        # logprob=-5.865 まで落ち、15〜45秒の本物の字幕が丸ごと失われた。
        # 繰り返しループから抜ける能力(フォールバック本来の目的)は 0.4 までで残す。
        temp_param = (0.0, 0.2, 0.4)
    else:
        temp_param = float(temperature)
    with open(err_path, "a", encoding="utf-8") as ef:
        ef.write(f"v2.8 params: no_prev_text={no_prev_text} word_ts={word_timestamps} rep_penalty={rep_penalty} vad_filter={vad_filter} batched={use_batched} hotwords={hotwords_text[:50]} prompt={prompt_text[:50]}\n")
    results = []
    filtered_count = 0
    for ci, clip in enumerate(clips):
        ci = clip.get("idx", ci)  # v2.9.20【F1】C++のg_tlClips上の元indexを使う(空wavで詰まるずれ対策)
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
            # v2.9.62【AT2】検出言語を記録する。**info を受け取っていながら一度も記録していなかった**。
            # 「言語=自動 + 短いクリップで韓国語に転ぶ」は実証済みの現象で、
            # 起きたときに**何語と判定されたかが唯一の証拠**なのに痕跡がゼロだった。
            # ★openai 経路と対で入れる(片方だけにしない)。
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Clip {ci}: detected language={getattr(info, 'language', '?')} "
                         f"prob={getattr(info, 'language_probability', 0):.2f}\n")
            # v2.9.22【I1】幻聴判定に VAD が常時必要。openai 経路と同じ形に揃える。
            # runs(スナップ用)は従来どおり morph_split のときだけ渡し、既存の配置挙動は変えない。
            vad_runs = _speech_runs(wav_path)
            env = _speech_env(wav_path) if morph_split else None
            _env_rr = _env_runs(env) if morph_split else []
            runs = vad_runs if morph_split else []
            # v2.9.67【AZ1】強制アライメントに全体が要るので材料化する(下のループで全部消費する)
            segments = list(segments)
            _aligned = _align_words(
                [(getattr(sg, "start", 0.0), getattr(sg, "end", 0.0), (getattr(sg, "text", "") or "").strip())
                 for sg in segments],
                # ★device ではなく actual_device。CUDA が使えず CPU に退避したとき、
                #   アライメントだけ CUDA を掴もうとして黙って無効化されるのを防ぐ。
                wav_path, getattr(info, "language", None) or language, actual_device, err_path, morph_stats
            ) if morph_split else {}
            for _si, seg in enumerate(segments):
                # v2.8: Segment quality filter
                # Skip low-confidence segments and likely non-speech
                # v2.9.5: openai経路と同様、捨てた内容を必ず記録する
                # v2.9.8: openai経路と同じ理由でしきい値を -3.0 に緩めた
                if (getattr(seg, 'avg_logprob', 0) < -3.0
                        or getattr(seg, 'no_speech_prob', 0) > 0.6):
                    filtered_count += 1
                    with open(err_path, "a", encoding="utf-8") as ef:
                        ef.write(f"Quality filtered: [{seg.start:.2f}-{seg.end:.2f}] "
                                 f"logprob={getattr(seg, 'avg_logprob', 0):.3f} "
                                 f"no_speech={getattr(seg, 'no_speech_prob', 0):.3f} "
                                 f"{seg.text.strip()[:40]}\n")
                    continue
                sf = tl_start + int(seg.start * fps)
                ef2 = tl_start + int(seg.end * fps)
                text = seg.text.strip()
                # v2.9.22【I1】幻聴フレーズ除去。従来この呼び出しは openai 経路にしか無く、
                # faster-whisper を選ぶとフィルタが一切動かなかった(実機ログ: faster 6本すべてで
                # Hallucination dropped が0件、openai 5本すべてで1件)。
                # hallucination_phrases.txt を編集しても faster では効かない状態だった。
                if _is_hallucination(text, seg.start, seg.end, vad_runs):
                    filtered_count += 1
                    with open(err_path, "a", encoding="utf-8") as ef:
                        ef.write(f"Hallucination dropped: [{seg.start:.2f}-{seg.end:.2f}] {text}\n")
                    continue
                wl = []
                _al = _aligned.get(_si) if morph_split else None
                if morph_split:
                    # v2.9.67【AZ1】整列できた分は差し替え、できなかった分は従来の時刻
                    wl = _al or ([(w.word, w.start, w.end) for w in seg.words]
                                 if getattr(seg, "words", None) else [])
                # v2.9.75【BC1】整列できたセグメントは VAD に吸着させない(openai 経路と同じ理由)
                # v2.9.75【BE1】直前の取りこぼしを拾う。★判定に渡すのは
                # **実際に表示される開始**(整列後の先頭単語)。whisper の区間開始を
                # 渡すと、整列でズレているぶん誤判定する(監査 A11 が検出)。
                # sf だけ動かしても
                # **単語時刻のほうが勝つ**ので、先頭の単語の開始も一緒に動かす。
                if morph_split and _env_rr:
                    _p_end = segments[_si - 1].end if _si > 0 else 0.0
                    _st0 = wl[0][1] if wl else seg.start
                    _st2 = _reclaim_orphan(_st0, _p_end, _env_rr, env)
                    if _st2 < _st0 - 1e-6:
                        sf = tl_start + int(_st2 * fps)
                        if wl:
                            wl = [(wl[0][0], _st2, wl[0][2])] + list(wl[1:])
                        morph_stats["orphan"] = morph_stats.get("orphan", 0) + 1
                _runs = [] if _al else runs
                # v2.9.75【BC1】上限(_cap_lead)も同じ理由で切る。エネルギー包絡は
                # 子音を拾えないので、整列が正しくても「音より早すぎる」と誤判定して
                # **後ろへ押す**(実測: 「ある区間」を 7.240s→7.650s と410ms遅らせていた)。
                _env = None if _al else env
                if text and ef2 > sf:
                    _emit_seg(results, ci, sf, ef2, text, morph_split, maxchars, wl, tl_start, fps, morph_stats, _runs, max_lines, _env)
        except Exception as e:
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Clip {ci} err: {e}\n{traceback.format_exc()}")
    with open(output_path, "w", encoding="utf-8") as f:
        for line in results:
            f.write(line + "\n")
    with open(err_path, "a", encoding="utf-8") as ef:
        ef.write(f"Morph timing: word-aligned={morph_stats.get('word', 0)} proportional={morph_stats.get('prop', 0)} snap-moved={morph_stats.get('snap', 0)} pause-split={morph_stats.get('pause', 0)} squash={morph_stats.get('squash', 0)} lead-capped={morph_stats.get('lead', 0)} aligned={morph_stats.get('align', 0)} degen-fixed={morph_stats.get('degen', 0)} orphan={morph_stats.get('orphan', 0)} vad={_vad_mode}\n")
        ef.write(f"Done: {len(results)} segs, {filtered_count} filtered ({actual_device}{'|batched' if use_batched else ''})\n")
        # v2.9.17: 内訳。モデルロードが支配的なのか転写なのかを毎回残す
        ef.write(f"Timing(py): load={_tLoad - _t0:.1f}s transcribe={_time.time() - _tLoad:.1f}s\n")

)PYHELPER" R"PYHELPER(
def _run_openai_whisper(model_size, language, device, clips, model_dir, output_path, err_path, beam_size=5, temperature=0, batch=None):
    import time as _time
    _t0 = _time.time()
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
    _tLoad = _time.time()  # v2.9.17: ここまでがモデルロード
    morph_stats = {}
    # v2.8: Temperature fallback tuple
    if temperature == 0:
        # v2.9.7: 上限を 1.0 から 0.4 に下げた。
        # 温度0はビーム探索(決定的)だが、0より大きいとサンプリングになる。
        # 0.6以上は粗すぎて日本語では実質ゴミしか出ず、復帰する見込みが薄い。
        # 実機で1.0まで上がった回に「dead」「moderator」「excited」という英語のゴミが出て
        # logprob=-5.865 まで落ち、15〜45秒の本物の字幕が丸ごと失われた。
        # 繰り返しループから抜ける能力(フォールバック本来の目的)は 0.4 までで残す。
        temp_param = (0.0, 0.2, 0.4)
    else:
        temp_param = float(temperature)
    results = []
    filtered_count = 0
    for ci, clip in enumerate(clips):
        ci = clip.get("idx", ci)  # v2.9.20【F1】C++のg_tlClips上の元indexを使う(空wavで詰まるずれ対策)
        wav_path = clip["wav"]
        tl_start = clip["timeline_start"]
        fps = clip["fps"]
        if not os.path.exists(wav_path):
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Clip {ci}: wav not found: {wav_path}\n")
            continue
        try:
            # v2.9.62【AT1】verbose=False は **進捗バーを消す指定ではない**。
            # openai-whisper の transcribe.py はこうなっている(実物で確認済み):
            #   263: # show the progress bar when verbose is False
            #   265: disable=verbose is not False   → False を渡すと disable=False = **バーが出る**
            # そのため tqdm の進捗が延々と stderr へ流れ、
            #  ① whisper_debug.log が `37%|###7 | 2832/7585 ...` で埋まる(実測)
            #  ② RunProcess の output に**無制限に溜まる**
            #  ③ ★生成が失敗したとき、エラーダイアログは pyOut の**末尾500文字**を
            #     「--- Python output ---」として見せるので、**肝心のエラーが進捗バーに
            #     押し出される**。今日直してきた「失敗の理由を正しく伝える」に真っ向から反する
            # → None にする。**verbose は表示専用**で、デコードにも結果にも一切影響しない
            #   (147/154/478行はすべて print のみ。既定値も None)。
            # ★副作用の申し送り: tqdm が書かなくなるので、AviUtl2 を閉じたときに
            #   「壊れたパイプで python が自滅する」偶然の保護が無くなる。
            #   ただし v2.9.52【AI2】で Job Object を入れたので、そちらが正規の手段。
            opts = {"language": language, "beam_size": beam_size, "verbose": None, "word_timestamps": word_timestamps, "temperature": temp_param, "condition_on_previous_text": not no_prev_text}
            if language is None:
                del opts["language"]
            if prompt_text:
                opts["initial_prompt"] = prompt_text
            # v2.9.4: hallucination_silence_threshold は渡さない。
            # v2.9.3 で試したが、本家の実装は is_segment_anomaly (単語確率/長さの異常) を
            # 通ったものしか消さず、whisper の定型幻聴は高確信度なので引っかからない。
            # 実機では幻聴が消えず「ご視聴ありがとうございました」→「それではまた」と
            # 別の定型句に化けただけで、ついでに他の字幕の時刻が最大84msずれた。
            # デコードには一切触れず、後段の _is_hallucination で決定的に落とす方式にする。
            result = model.transcribe(wav_path, **opts)
            # v2.9.62【AT2】検出言語を**自分で**記録する。従来は本体が出す
            # "Detected language: xx" に無自覚に頼っていたが、あれは transcribe.py 154行の
            # `if verbose is not None:` で守られており、**AT1 で verbose=None にした瞬間に消える**。
            # ★AT1 だけ入れてここを入れないと、診断材料を1つ失う退行になっていた。
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Clip {ci}: detected language={result.get('language', '?')}\n")
            # v2.9.3: 幻聴判定に VAD が常時必要。runs(スナップ用)は従来どおり
            # morph_split のときだけ渡し、既存の配置挙動を一切変えない。
            vad_runs = _speech_runs(wav_path)
            env = _speech_env(wav_path) if morph_split else None
            _env_rr = _env_runs(env) if morph_split else []
            runs = vad_runs if morph_split else []
            _segs = result.get("segments", []) or []
            # v2.9.67【AZ1】単語時刻を強制アライメントで取り直す(**テキストは触らない**)
            _aligned = _align_words(
                [(sg.get("start", 0.0), sg.get("end", 0.0), (sg.get("text") or "").strip()) for sg in _segs],
                wav_path, result.get("language") or language, device, err_path, morph_stats
            ) if morph_split else {}
            for _si, seg in enumerate(_segs):
                # v2.8: Segment quality filter
                # v2.9.5: 何を捨てたかを必ず記録する。ここが見えないと「字幕が減った」
                # ときに幻聴が消えたのか本物が消えたのか切り分けられない。
                # v2.9.8: しきい値を -1.0 から -3.0 に緩めた。
                # 実測で正しい字幕が logprob=-1.029 / -1.218 / -1.273 で捨てられていた
                # (「短い相づち」「固有名詞!」「固有名詞?」)。一方、本物のゴミは -5.865。
                # whisper 本体は -1.0 を「フォールバックを試す合図」として使っているだけで、
                # 「捨てる基準」ではない。温度を上げて得た結果は logprob が下がるのが当然で、
                # 同じ値で切ると正解まで巻き添えになる。
                if seg.get("avg_logprob", 0) < -3.0 or seg.get("no_speech_prob", 0) > 0.6:
                    filtered_count += 1
                    with open(err_path, "a", encoding="utf-8") as ef:
                        ef.write(f"Quality filtered: [{seg['start']:.2f}-{seg['end']:.2f}] "
                                 f"logprob={seg.get('avg_logprob', 0):.3f} "
                                 f"no_speech={seg.get('no_speech_prob', 0):.3f} "
                                 f"temp={seg.get('temperature', 0)} "
                                 f"{seg['text'].strip()[:40]}\n")
                    continue
                sf = tl_start + int(seg["start"] * fps)
                ef2 = tl_start + int(seg["end"] * fps)
                text = seg["text"].strip()
                # v2.9.3: 幻聴フレーズ除去 (無音区間に生える「ご視聴ありがとうございました」等)
                if _is_hallucination(text, seg["start"], seg["end"], vad_runs):
                    filtered_count += 1
                    with open(err_path, "a", encoding="utf-8") as ef:
                        ef.write(f"Hallucination dropped: [{seg['start']:.2f}-{seg['end']:.2f}] {text}\n")
                    continue
                wl = []
                _al = _aligned.get(_si) if morph_split else None
                if morph_split:
                    # v2.9.67【AZ1】整列できた分は差し替え、できなかった分は従来の時刻
                    wl = _al or [(w.get("word", ""), w.get("start"), w.get("end"))
                                 for w in seg.get("words", [])]
                # v2.9.75【BC1】整列できたセグメントは **VADに吸着させない**。
                # _snap_start は「whisper の粗い時刻を発話の頭へ寄せる」救済策で、
                # 時刻が不正確だった時代の遺物。強制アライメントは文字単位で±17msなのに、
                # VAD区間は数十秒に十数個しかない粗い情報で、**280msも後ろへ引っ張る**。
                # 実測(実機素材): 整列の時刻からのズレ 中央 92ms → 12ms。
                # ★以前「スナップは無害」と測ったのは、物差し自体がVAD由来で
                #   **スナップをスナップの基準で測っていた**ため。開発者が波形を見て気づいた。
                # v2.9.75【BE1】直前の取りこぼしを拾う。★判定に渡すのは
                # **実際に表示される開始**(整列後の先頭単語)。whisper の区間開始を
                # 渡すと、整列でズレているぶん誤判定する(監査 A11 が検出)。
                # sf だけ動かしても
                # **単語時刻のほうが勝つ**ので、先頭の単語の開始も一緒に動かす。
                if morph_split and _env_rr:
                    _p_end = _segs[_si - 1].get("end", 0.0) if _si > 0 else 0.0
                    _st0 = wl[0][1] if wl else seg.get("start", 0.0)
                    _st2 = _reclaim_orphan(_st0, _p_end, _env_rr, env)
                    if _st2 < _st0 - 1e-6:
                        sf = tl_start + int(_st2 * fps)
                        if wl:
                            wl = [(wl[0][0], _st2, wl[0][2])] + list(wl[1:])
                        morph_stats["orphan"] = morph_stats.get("orphan", 0) + 1
                _runs = [] if _al else runs
                # v2.9.75【BC1】上限(_cap_lead)も同じ理由で切る。エネルギー包絡は
                # 子音を拾えないので、整列が正しくても「音より早すぎる」と誤判定して
                # **後ろへ押す**(実測: 「ある区間」を 7.240s→7.650s と410ms遅らせていた)。
                _env = None if _al else env
                if text and ef2 > sf:
                    _emit_seg(results, ci, sf, ef2, text, morph_split, maxchars, wl, tl_start, fps, morph_stats, _runs, max_lines, _env)
        except Exception as e:
            with open(err_path, "a", encoding="utf-8") as ef:
                ef.write(f"Clip {ci} err: {e}\n{traceback.format_exc()}")
    with open(output_path, "w", encoding="utf-8") as f:
        for line in results:
            f.write(line + "\n")
    with open(err_path, "a", encoding="utf-8") as ef:
        ef.write(f"Morph timing: word-aligned={morph_stats.get('word', 0)} proportional={morph_stats.get('prop', 0)} snap-moved={morph_stats.get('snap', 0)} pause-split={morph_stats.get('pause', 0)} squash={morph_stats.get('squash', 0)} lead-capped={morph_stats.get('lead', 0)} aligned={morph_stats.get('align', 0)} degen-fixed={morph_stats.get('degen', 0)} orphan={morph_stats.get('orphan', 0)} vad={_vad_mode}\n")
        ef.write(f"Done: {len(results)} segs, {filtered_count} filtered (openai-whisper, {device})"
                 f" hal_phrases={len(_HAL_PHRASES)}({_HAL_SRC})\n")
        # v2.9.17: 内訳。モデルロードが支配的なのか転写なのかを毎回残す
        ef.write(f"Timing(py): load={_tLoad - _t0:.1f}s transcribe={_time.time() - _tLoad:.1f}s\n")

if __name__ == "__main__":
    main()
)PYHELPER";

// v2.9.53【AJ1】書けたかを見ていなかった。従来は is_open() だけを見て、
// 書き込みと flush の成否を一切確かめていない。
// ★ofstream は**開いた時点で truncate する**ので、失敗すると
//   それまで正しかったヘルパーが空か途中で壊れた状態に置き換わり、誰も気づかない。
//   次の生成で python は壊れたスクリプトを実行し、利用者には無関係なエラーが出る。
// ★SRT 側(AD1)は if(!f) → close() → if(!f) と**二重に**確かめて理由まで出しているのに、
//   こちらは素通りだった。「同じ処理が2箇所で片方だけ」型。同じ手順に揃える。
// 書き込みが走るのは existing != embedded のときだけ = **版を上げた直後が一番踏みやすい**。
static bool EnsurePyHelper(){
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
    if(existing == embedded) return true; // 書く必要がない = 既に正しい
    std::ofstream f(Utf8ToWide(p), std::ios::binary);
    if(!f){
        DebugLog("EnsurePyHelper: open failed: " + p);
        return false;
    }
    f << g_pyHelper;
    // ★close() まで済ませてから確かめる。バッファに載っただけで失敗が出るのは flush 時。
    f.close();
    if(!f){
        DebugLog("EnsurePyHelper: write failed (file is now broken): " + p);
        return false;
    }
    return true;
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

// v2.9.52【AI1】直近の実行で転写が投げた理由(表示用)。
// Python はクリップ単位で例外を握り潰し `Clip N err: <理由>` を .err に書くだけで
// exit 0 で終わるため、C++ からは成功に見えて「結果が空」の分岐に落ちる。
// ★理由は .err に**ある**のに一度も表示していなかった(AG3 と同じ型)。
// ★状態を新しく持つので、リセット位置を先に決める → CollectClipErrors() の冒頭で必ず clear。
//   本実行と再試行の両方がこの1つの関数を呼ぶ(「同じ処理が2箇所で片方だけ」を作らない)。
static std::string g_clipErr;

// v2.9.57【AM2】numpy と torch が噛み合っていないときの日本語説明。**3箇所から呼ぶ**。
// ★AG3 で「エラーを出す経路は2つある。両方に置くこと」と直したが、
//   **再試行側の ERROR 行は対象外のままだった**。しかも再試行は
//   「パッケージを自動導入した直後」に走るので、**numpy 不整合が最も出やすい場面**。
//   そこで英語だけが出ていた(AG3 と同じ症状が、対の片割れに残っていた)。
// ★3箇所目を足すときにまたコピーすると必ず片方だけになるので、ここに集約する。
// ★判定は Python 側の出力を拾うだけ(C++ に同じ判定を書かない。Z2 と同じ方針)。
static std::string NumpyHintIfNeeded(const std::string& out, const std::string& err, const std::string& line){
    if(out.find("Numpy is not available") == std::string::npos
       && err.find("Numpy is not available") == std::string::npos
       && out.find("compiled using NumPy 1.x") == std::string::npos
       && err.find("compiled using NumPy 1.x") == std::string::npos
       && line.find("Numpy is not available") == std::string::npos)
        return "";
    return "\x50\x79\x54\x6f\x72\x63\x68\x20\xe3\x81\xa8\x20\x4e\x75\x6d\x50\x79\x20\xe3\x81\xae\xe3\x83\x90\xe3\x83\xbc\xe3\x82\xb8\xe3\x83\xa7\xe3\x83\xb3\xe3\x81\x8c\xe5\x99\x9b\xe3\x81\xbf\xe5\x90\x88\xe3\x81\xa3\xe3\x81\xa6\xe3\x81\x84\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n\xe3\x82\xbb\xe3\x83\x83\xe3\x83\x88\xe3\x82\xa2\xe3\x83\x83\xe3\x83\x97\xe3\x81\x8c\xe6\x9c\x80\xe5\xbe\x8c\xe3\x81\xbe\xe3\x81\xa7\xe7\xb5\x82\xe3\x82\x8f\xe3\x81\xa3\xe3\x81\xa6\xe3\x81\x84\xe3\x81\xaa\xe3\x81\x84\xe5\x8f\xaf\xe8\x83\xbd\xe6\x80\xa7\xe3\x81\x8c\xe3\x81\x82\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x99\xe3\x80\x82\n\xe3\x80\x8c\xe7\x92\xb0\xe5\xa2\x83\xe3\x80\x8d\xe3\x82\xbf\xe3\x83\x96\xe3\x81\xae\xe3\x80\x8c\xe3\x82\xbb\xe3\x83\x83\xe3\x83\x88\xe3\x82\xa2\xe3\x83\x83\xe3\x83\x97\xe3\x80\x8d\xe3\x82\x92\xe3\x80\x81\xe5\xae\x8c\xe4\xba\x86\xe3\x81\x99\xe3\x82\x8b\xe3\x81\xbe\xe3\x81\xa7\xe9\x96\x89\xe3\x81\x98\xe3\x81\x9a\xe3\x81\xab\xe5\xae\x9f\xe8\xa1\x8c\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82\n\x28\x50\x79\x54\x6f\x72\x63\x68\x20\xe3\x81\xaf\xe7\xb4\x84\x34\x2e\x33\x47\x42\xe3\x81\x82\xe3\x82\x8a\xe3\x80\x81\xe6\x99\x82\xe9\x96\x93\xe3\x81\x8c\xe3\x81\x8b\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x99\x29\n\n";
}

static void CollectClipErrors(const std::string& errPath){
    g_clipErr.clear();
    std::ifstream ef(Utf8ToWide(errPath));
    if(!ef.is_open()) return;
    std::string l;
    while(std::getline(ef, l)){
        if(l.compare(0, 5, "Clip ") != 0) continue;
        size_t p = l.find(" err: ");
        if(p == std::string::npos) continue;
        std::string reason = l.substr(p + 6);
        while(!reason.empty() && (reason.back() == '\r' || reason.back() == '\n')) reason.pop_back();
        if(reason.empty()) continue;
        if(g_clipErr.find(reason) != std::string::npos) continue; // 同じ理由は1回だけ
        if(!g_clipErr.empty()) g_clipErr += "\n";
        g_clipErr += reason;
    }
}

// v2.9.52【AI2】転写の子プロセスを入れる Job。AviUtl2 が終了したら道連れで終了させる。
// ★掛けるのは転写だけ。セットアップ(pip)には掛けない — 導入の途中で殺すと
//   「中途半端に入った状態」を新たに作ることになり、AG6 で苦労した状態そのものになる。
// ★ハンドルは意図的に閉じない。プロセス終了時に OS が閉じ、そのとき子が終了する。
static HANDLE GetKillOnExitJob(){
    static HANDLE s_job = [](){
        HANDLE j = CreateJobObjectW(NULL, NULL);
        if(!j) return (HANDLE)NULL;
        JOBOBJECT_EXTENDED_LIMIT_INFORMATION li = {};
        li.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
        if(!SetInformationJobObject(j, JobObjectExtendedLimitInformation, &li, sizeof(li))){
            CloseHandle(j);
            return (HANDLE)NULL;
        }
        return j;
    }();
    return s_job;
}

static bool RunProcess(const std::wstring& cmdLine, std::string& output, DWORD timeoutMs = 300000, bool killWithParent = false){
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
    // v2.9.52【AI2】killWithParent のときだけ CREATE_SUSPENDED で起こす。
    // Job へ入れる前に子が動き出すと、その隙に孫プロセスが Job の外へ出るため。
    // ★既定は false なので、既存の呼び出しはフラグも含めて従来と完全に同じ。
    DWORD createFlags = CREATE_NO_WINDOW | (killWithParent ? CREATE_SUSPENDED : 0);
    if(!CreateProcessW(NULL, &cmd[0], NULL, NULL, TRUE, createFlags, NULL, NULL, &si, &pi)){
        CloseHandle(hReadOut); CloseHandle(hWriteOut);
        output = "CreateProcess failed: " + std::to_string(GetLastError());
        return false;
    }
    if(killWithParent){
        // 割り当てに失敗しても続行する(従来どおり動く)。詰まらせない。
        HANDLE job = GetKillOnExitJob();
        if(job) AssignProcessToJobObject(job, pi.hProcess);
        // ★どの経路でも必ず再開する。ここを飛ばすと子が永久に止まったままになる。
        ResumeThread(pi.hThread);
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

// v2.9.55【AL1】設定を保存できなかったことを知らせる。★毎回出すと鬱陶しいので1回だけ。
// 記録(DebugLog)は毎回残す。「黙って失い続ける」のを避けるのが目的。
static void NotifySettingsSaveFailed(){
    static bool s_told = false;
    if(s_told) return;
    s_told = true;
    MsgBox(g_wnd, "\xe8\xa8\xad\xe5\xae\x9a\xe3\x82\x92\xe4\xbf\x9d\xe5\xad\x98\xe3\x81\xa7\xe3\x81\x8d\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x81\xa7\xe3\x81\x97\xe3\x81\x9f\xe3\x80\x82\x0a\x0a\xe3\x81\x93\xe3\x81\xae\xe3\x81\xbe\xe3\x81\xbe\xe3\x81\xa0\xe3\x81\xa8\xe3\x80\x81\xe5\xa4\x89\xe6\x9b\xb4\xe3\x81\x97\xe3\x81\x9f\xe8\xa8\xad\xe5\xae\x9a\xe3\x81\xaf\xe6\xac\xa1\xe5\x9b\x9e\xe8\xb5\xb7\xe5\x8b\x95\xe6\x99\x82\xe3\x81\xab\xe5\x85\x83\xe3\x81\xab\xe6\x88\xbb\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x99\xe3\x80\x82\x0a\xe3\x83\x97\xe3\x83\xa9\xe3\x82\xb0\xe3\x82\xa4\xe3\x83\xb3\xe3\x83\x95\xe3\x82\xa9\xe3\x83\xab\xe3\x83\x80\xe3\x81\xae\xe7\xa9\xba\xe3\x81\x8d\xe5\xae\xb9\xe9\x87\x8f\xe3\x81\xa8\xe3\x82\xa2\xe3\x82\xaf\xe3\x82\xbb\xe3\x82\xb9\xe6\xa8\xa9\xe3\x82\x92\xe7\xa2\xba\xe8\xaa\x8d\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82\x0a\x0a\x28\xe3\x81\x93\xe3\x81\xae\xe9\x80\x9a\xe7\x9f\xa5\xe3\x81\xaf\x31\xe5\x9b\x9e\xe3\x81\xa0\xe3\x81\x91\xe8\xa1\xa8\xe7\xa4\xba\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x99\xe3\x80\x82\xe8\xa8\x98\xe9\x8c\xb2\xe3\x81\xaf\x20\x77\x68\x69\x73\x70\x65\x72\x5f\x64\x65\x62\x75\x67\x2e\x6c\x6f\x67\x20\xe3\x81\xab\xe6\xae\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x99\x29",
           "Whisper Subtitle", MB_OK|MB_ICONWARNING);
}

// v2.9.55【AL1】書けたかを見ていなかった。★ofstream は**開いた時点で truncate する**ので、
// 途中で失敗すると ini が**途中まで**の状態で残り、次回起動で設定の一部が既定値に戻る。
// しかも黙って起きるので、利用者には「設定が勝手に戻る」としか見えない。
// AD1(SRT)・AJ1(ヘルパー)と同じ「失敗を成功として扱う」型。3箇所目。
// ★毎回ダイアログを出すと鬱陶しい(SaveSettings は設定変更のたびに呼ばれる)ので、
//   記録は毎回・通知はセッション中1回だけにする。黙って失い続けるのを避けるのが目的。
static void SaveSettings(){
    std::ofstream f(Utf8ToWide(GetIniPath()));
    if(!f.is_open()){
        DebugLog("SaveSettings: open failed: " + GetIniPath());
        NotifySettingsSaveFailed();
        return;
    }
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
        // v2.9.1【文字化け修正】ANSI版(GetWindowTextA)はシステムコードページ(日本語環境ならCP932)の
        // バイト列を返すため、日本語を入力すると **CP932のまま ini と whisper_batch.json に書かれ**、
        // Python側の json.load(encoding="utf-8") が
        // 「'utf-8' codec can't decode byte 0x94 ... invalid start byte」で失敗していた(Issue #5)。
        // UTF-16で取得してから WideToUtf8() でUTF-8に変換する。
        int len = GetWindowTextLengthW(g_promptEdit);
        if(len > 0){
            std::wstring wBuf(len + 1, 0); // L'\0' と書くと \x エスケープ事故のもとなので 0 を使う
            GetWindowTextW(g_promptEdit, &wBuf[0], len + 1);
            wBuf.resize(len);
            std::string pBuf = WideToUtf8(wBuf);
            f << "prompt=" << IniEscape(pBuf) << "\n"; // v2.9.18
        } else {
            f << "prompt=\n";
        }
    }
    // v2.8 settings
    f << "batched=" << (SendMessageA(g_chkBatched, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0) << "\n";
    {
        // v2.9.1【文字化け修正】prompt と同じ理由でUTF-16経由に変更(Issue #5)
        int len = GetWindowTextLengthW(g_hotwordsEdit);
        if(len > 0){
            std::wstring wBuf(len + 1, 0);
            GetWindowTextW(g_hotwordsEdit, &wBuf[0], len + 1);
            wBuf.resize(len);
            std::string hBuf = WideToUtf8(wBuf);
            f << "hotwords=" << IniEscape(hBuf) << "\n"; // v2.9.18: prompt と同じ扱いに揃える
        } else {
            f << "hotwords=\n";
        }
    }
    // v2.9.55【AL1】★close() まで済ませてから確かめる。失敗が出るのは flush 時。
    // ここを見ないと、途中まで書かれた ini が残って次回起動で設定の一部が既定値に戻る。
    f.close();
    if(!f){
        DebugLog("SaveSettings: write failed (ini may be truncated): " + GetIniPath());
        NotifySettingsSaveFailed();
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
            // v2.9.1【文字化け修正】iniはUTF-8で書かれているので、ANSI版で戻すと化ける。
            // UTF-16に変換してからWide版で設定する(Issue #5)
            SetWindowTextW(g_promptEdit, Utf8ToWide(IniUnescape(val)).c_str()); // v2.9.18
        }
        // v2.8 settings
        else if(key=="batched") SendMessageA(g_chkBatched, BM_SETCHECK, atoi(val.c_str()) ? BST_CHECKED : BST_UNCHECKED, 0);
        else if(key=="hotwords") SetWindowTextW(g_hotwordsEdit, Utf8ToWide(IniUnescape(val)).c_str()); // v2.9.1【文字化け修正】/ v2.9.18: エスケープ対応
    }
}

// =========================================================================
// Setup thread (auto-install faster-whisper + model DL)
// =========================================================================

// Forward declarations (defined later, after timeline code)
static void SetStatus(const std::string& msg);
static void SetProgress(int val);
static void UpdateWhisperLocLabels();
static void WriteTestLog(); // v2.9.6
static void UpdateFugashiStatus(); // v2.8.5
static void SetStatusW(const wchar_t* msg); // v2.8.10【I】: g_status/g_statusSetup同時更新版(定義はSetStatus/SetProgressの近く)
static void ProbeThread(); // v2.9.10: SetupThread末尾から再実測のため呼ぶ(定義は後方)

// v2.9.0【D】: g_chkForceReinstall(1個)を項目別チェック5個に置き換えたための共通ヘルパー。
static bool IsChecked(HWND h){ return h && SendMessageA(h, BM_GETCHECK, 0, 0) == BST_CHECKED; }
static void ClearSetupChecks(){ // v2.9.0【D】: SetupThreadの3つの終了経路すべてから呼ぶ(誤って2回目が走るのを防ぐ)
    HWND cs[] = {g_chkFfmpeg, g_chkTorch, g_chkWhisper, g_chkFaster, g_chkFugashi, g_chkModel}; // v2.9.4: g_chkFfmpeg追加 / v2.9.49【AH2】: g_chkModel復活
    for(HWND c : cs) if(c) SendMessageA(c, BM_SETCHECK, BST_UNCHECKED, 0);
}

// v2.9.0【G】: モデル名テーブル。従来 SetupThread 内にローカルで持っていたもの(const char* mn[])を
// ファイルスコープへ切り出し、SetupThread と ModelExists() の両方から同じ表を参照する。
static const char* kModelNames[] = {"tiny","base","small","medium","large-v3","large-v3-turbo","kotoba-whisper"};

// v2.9.42【AF1】HuggingFace キャッシュが「中断DLの残骸」でないかを見る。
// HF は DL を始めた時点でキャッシュのフォルダを作るため、途中で止まると
// `models--<配布元>--<リポジトリ>` だけが残る。名前しか見ていないと**完成扱い**になり、
// 再DLも不足表示もされないまま生成時に失敗する(新規利用者だけが踏む)。
//
// ★判定は **緩く** すること。`snapshots\<hash>\` に **何かファイルが1つでもあれば** 導入済みとみなす。
//   `model.bin` の有無に限定してはいけない。配布形式が変わったときに
//   **K1(実際は入っているのに未導入と誤判定 → 生成がブロックされ入れ直しても直らない無限ループ)**
//   が再発する。K1 は利用者報告から直した箇所で、同じ穴に二度落ちないこと。
static bool HasSnapshotFile(const std::string& hfDir){
    std::string snapRoot = hfDir + "\\snapshots";
    WIN32_FIND_DATAW fd;
    HANDLE h = FindFirstFileW(Utf8ToWide(snapRoot + "\\*").c_str(), &fd);
    if(h == INVALID_HANDLE_VALUE) return false;   // snapshots が無い = DL が始まってすらいない
    bool found = false;
    do{
        if(!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) continue;
        if(fd.cFileName[0] == L'.') continue;      // "." と ".."
        std::string sub = snapRoot + "\\" + WideToUtf8(fd.cFileName);
        WIN32_FIND_DATAW fd2;
        HANDLE h2 = FindFirstFileW(Utf8ToWide(sub + "\\*").c_str(), &fd2);
        if(h2 != INVALID_HANDLE_VALUE){
            do{
                // ★ファイルなら何でもよい(symlink も FILE_ATTRIBUTE_DIRECTORY は立たない)
                if(!(fd2.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)){ found = true; break; }
            } while(FindNextFileW(h2, &fd2));
            FindClose(h2);
        }
        if(found) break;
    } while(FindNextFileW(h, &fd));
    FindClose(h);
    return found;
}

// v2.9.0【G】: SetupThread 内にインラインで書かれていたモデル存在判定を、環境タブ表示からも
// 使えるように切り出したもの。判定内容は従来と完全に同一(faster-whisper の config.json / model.bin、
// openai-whisper の .pt、kotoba の HF キャッシュ)。
static bool ModelExists(const std::string& mName){
    std::string mDir = GetModelsDir();
    std::string localModel = mDir + "\\" + mName;
    bool exists = FileExistsU(localModel + "\\config.json")   // faster-whisper
        || FileExistsU(localModel + "\\model.bin")              // faster-whisper alt
        || FileExistsU(mDir + "\\" + mName + ".pt");            // openai-whisper
    // v2.9.24【K1】faster-whisper は HuggingFace のキャッシュ形式
    // `models--<配布元>--<リポジトリ>` に落とすが、その形式を見る処理が下の kotoba 特例に
    // しか無く、他のモデルは実際には入っているのに「未導入」と誤判定していた(利用者report)。
    // 表示だけの問題ではない: CheckMissingForGenerate() 経由で**字幕生成がブロックされる**
    // (入れ直しても未導入のまま = 無限ループ)。v2.9.2 は faster が既定なので影響が大きい。
    // ★開発者環境では再現しない: openai-whisper も使っていると <名前>.pt があり
    //   上の3番目のチェックが通ってしまうため。faster しか使わない環境だけが踏む。
    // リポジトリ名は "faster-whisper-<モデル名>" のように末尾がモデル名になる規約なので
    // `-<モデル名>` で終わるフォルダを探す。配布元が変わっても効く
    // (large-v3-turbo は Systran → mobiuslabsgmbh へ移動した実績がある)。
    // ★前方一致ではなく**後方一致**にすること。前方一致だと "large-v3" が
    //   "...-large-v3-turbo" を誤って拾う。
    if(!exists){
        std::string tail = "-" + mName;
        WIN32_FIND_DATAW fd;
        std::wstring pat = Utf8ToWide(mDir) + L"\\models--*";
        HANDLE h = FindFirstFileW(pat.c_str(), &fd);
        if(h != INVALID_HANDLE_VALUE){
            do{
                if(!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) continue;
                std::string nm = WideToUtf8(fd.cFileName);
                if(nm.size() > tail.size()
                   && nm.compare(nm.size() - tail.size(), tail.size(), tail) == 0){
                    // v2.9.42【AF1】名前が合っても中身が空(中断DLの残骸)なら導入済みにしない。
                    // 判定は緩く — snapshots に何かファイルが1つでもあればよい(K1 再発の防止)。
                    if(HasSnapshotFile(mDir + "\\" + nm)){
                        exists = true;
                        break;
                    }
                }
            } while(FindNextFileW(h, &fd));
            FindClose(h);
        }
    }
    // kotoba はリポジトリ名が "kotoba-whisper-v2.0-faster" でモデル名で終わらないため個別に見る
    if(!exists && mName.find("kotoba") == 0){
        std::string hfCache = mDir + "\\models--kotoba-tech--kotoba-whisper-v2.0-faster";
        // v2.9.42【AF1】★上の一般経路と **対で** 直す。フォルダの存在だけだと
        // 中断DLの残骸を完成扱いする。片方だけ直すと N1・V2・W1・AE1 と同じ轍を踏む。
        exists = FileExistsU(hfCache) && HasSnapshotFile(hfCache);
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
    // v2.9.2: Batched も faster-whisper 専用。BatchedInferencePipeline は faster_whisper の機能で、
    // _run_openai_whisper 側は batched を一度も参照しない(grepで0件確認)。
    // VAD・繰返し抑制と同じく、効かないチェックを触れる状態で残さない。
    if(g_chkBatched) EnableWindow(g_chkBatched, isFaster);
    // v2.9.11【監査①】ホットワードも faster-whisper 専用だった。
    // _run_openai_whisper は hotwords を一度も参照しない(grepで0件確認)。
    // 入力できるのに黙って無視される状態だったので、他の3つと同じ扱いに揃える。
    // (openai-whisper には hotwords 相当のAPIが無い。initial_prompt で代用する案はあるが、
    //  プロンプトをそのまま字幕として出力してしまう既知の失敗があるため採用しない)
    if(g_hotwordsEdit) EnableWindow(g_hotwordsEdit, isFaster);
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

// v2.9.26【L1】GPU はあるのに、その Python 版に対応する GPU 版ホイールが無くて
// CPU 版へ落ちたときの説明文。空でなければセットアップ報告とログに必ず出す。
// 黙って CPU 版にして「なぜか遅い」状態にするのが最悪なので、理由を利用者に見せる。
// (setup と生成は g_busy で排他されるので同時に書かれることはない)
static std::string g_cudaNoWheelNote;

// v2.9.26【L1】**torch のインデックス選択専用**。DetectCudaTag() と分けてあるのは、
// あちらが faster-whisper 用 cuBLAS/cuDNN の要否判定にも使われているため。
// torch の都合(Python版数)であちらを "cpu" にすると cuDNN が入らなくなり、
// CUDA を試した転写が cublas64_12.dll 不在で途中死する(CPUフォールバックも効かず字幕ゼロ)。
// ★torch のホイールはドライバだけでなく Python 版数にも依存する。実測(2026-07-30, win_amd64):
//     cu118/cu121: cp38-cp312 / cu124: cp38-cp313 / cu126・cu128: cp39-cp314
//   従来はドライバ上限だけで選んでいたため、例えば「CUDA 12.1 + Python 3.14」で
//   cu121 を指し、cp314 が無いので pip が必ず失敗していた。
//   python.org の既定ボタンは常に最新版(3.14系)を配るので新規ユーザーが踏みやすい。
// ★CUDA版が上がるほど Python 対応も広いので、ドライバ上限のタグで足りない場合に
//   下位へ降りても解決しない。下方向の探索はせず CPU 版へ退避する
//   (CPU 版は cp314 まで存在する。PyPI torch 2.13.0 の cp314-cp314-win_amd64 で確認)。
static std::string PickTorchCudaTag(){
    g_cudaNoWheelNote.clear();
    std::string tag = DetectCudaTag();
    if(tag == "cpu") return "cpu";
    int pym = (g_probe.done && g_probe.pyMajor == 3) ? g_probe.pyMinor : 0;
    if(pym == 0) return tag; // Python版数が不明(probe未完了)。従来どおり制限しない
    int maxPy = 12;                                   // cu118 / cu121
    if(tag == "cu126" || tag == "cu128" || tag == "cu129") maxPy = 14;
    else if(tag == "cu124") maxPy = 13;
    if(pym <= maxPy) return tag;
    g_cudaNoWheelNote = "\x47\x50\x55\xe3\x81\xaf\xe6\xa4\x9c\xe5\x87\xba\xe3\x81\x95\xe3\x82\x8c\xe3\x81\xbe\xe3\x81\x97\xe3\x81\x9f\xe3\x81\x8c\xe3\x80\x81\xe3\x83\x89\xe3\x83\xa9\xe3\x82\xa4\xe3\x83\x90\xe3\x81\x8c\xe5\xaf\xbe\xe5\xbf\x9c\xe3\x81\x99\xe3\x82\x8b\x20\x43\x55\x44\x41\x20\xe3\x81\x8c\xe5\x8f\xa4\xe3\x81\x8f\xe3\x80\x81\x50\x79\x74\x68\x6f\x6e\x20\x33\x2e" + std::to_string(pym) + "\x20\xe7\x94\xa8\xe3\x81\xae\x20\x47\x50\x55\x20\xe7\x89\x88\x20\x50\x79\x54\x6f\x72\x63\x68\x20\xe3\x81\x8c\xe9\x85\x8d\xe5\xb8\x83\xe3\x81\x95\xe3\x82\x8c\xe3\x81\xa6\xe3\x81\x84\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\x43\x50\x55\x20\xe7\x89\x88\xe3\x82\x92\xe5\xb0\x8e\xe5\x85\xa5\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x99\x28\xe5\x8b\x95\xe4\xbd\x9c\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x99\xe3\x81\x8c\xe4\xbd\x8e\xe9\x80\x9f\xe3\x81\xa7\xe3\x81\x99\x29\xe3\x80\x82\x47\x50\x55\x20\xe3\x82\x92\xe4\xbd\xbf\xe3\x81\x86\xe3\x81\xab\xe3\x81\xaf\xe3\x82\xb0\xe3\x83\xa9\xe3\x83\x95\xe3\x82\xa3\xe3\x83\x83\xe3\x82\xaf\xe3\x83\x89\xe3\x83\xa9\xe3\x82\xa4\xe3\x83\x90\xe3\x82\x92\xe6\x9b\xb4\xe6\x96\xb0\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82\n";
    return "cpu";
}

// v2.8.8: ffmpeg を公式配布元から取得して プラグインフォルダ\ffmpeg\ に展開する。
// 成功したら ffmpeg.exe のフルパスを返す。失敗時は空文字。
// v2.9.53【AJ2】PowerShell のシングルクォート文字列へ埋めるための脱出。
// ' は '' にする。これをしないと**パスにアポストロフィを含む環境**
// (ユーザー名が O'Brien / D'Angelo など)でコマンドが途中で切れ、
// ffmpeg の自動取得が必ず失敗する。日本語環境では稀だが海外の利用者は踏む。
// ★シングルクォート内では $ もバッククォートも literal なので、脱出が要るのは ' だけ。
static std::string PsQuote(const std::string& s){
    std::string r;
    for(char c : s){
        if(c == '\'') r += "''";
        else r += c;
    }
    return r;
}

static std::string DownloadFFmpeg(){
    std::string dest = GetPluginDir() + "\\ffmpeg";
    std::string zip  = GetTempDir() + "\\ffmpeg_dl.zip";
    // v2.9.53【AJ2】PowerShell へ渡す分だけ脱出した控えを作る(実ファイル操作には使わない)
    std::string zipQ  = PsQuote(zip);
    std::string destQ = PsQuote(dest);
    std::string url  = "https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip";
    // v2.9.23【J2】実機で初回DLが非常に遅く70%で停止した。3点直す:
    //  1. $ProgressPreference='SilentlyContinue' — PowerShell 5.1 の Invoke-WebRequest は
    //     進捗バー描画のオーバーヘッドで大幅に遅くなる(既知の問題)
    //  2. 失敗時に1回リトライ。その前に中断した zip を必ず消す(再開はできないので消して取り直す)
    //  3. -TimeoutSec を明示して無限待ちを防ぐ
    // 進捗表示は呼び出し元が zip のサイズをポーリングしているのでこの変更の影響を受けない。
    std::string ps = "powershell -NoProfile -ExecutionPolicy Bypass -Command \""
        "$ErrorActionPreference='Stop';"
        "$ProgressPreference='SilentlyContinue';"
        "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12;"
        "if(Test-Path '" + zipQ + "'){Remove-Item '" + zipQ + "' -Force};"
        "$ok=$false;"
        "for($i=1; $i -le 2 -and -not $ok; $i++){"
        "  try{ Invoke-WebRequest -Uri '" + url + "' -OutFile '" + zipQ + "' -UseBasicParsing -TimeoutSec 600; $ok=$true }"
        "  catch{ if(Test-Path '" + zipQ + "'){Remove-Item '" + zipQ + "' -Force}; Start-Sleep -Seconds 3 }"
        "};"
        "if(-not $ok){ throw 'ffmpeg download failed after 2 attempts' };"
        "Expand-Archive -Path '" + zipQ + "' -DestinationPath '" + destQ + "' -Force;"
        "Remove-Item '" + zipQ + "' -Force\"";
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
    bool fModel   = IsChecked(g_chkModel);  // v2.9.49【AH2】: モデルの強制再DL(緩い判定の逃げ道)
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
        SetStatusW(L"Ready (v2.9.75)"); // v2.8.10【I】: SetWindowTextW直接呼びを置き換え(環境タブ凍結対策)+バージョン更新
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
            SetStatusW(L"Ready (v2.9.75)");
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
        // v2.9.48【AG6】**セットアップを何度押しても直らない**問題を修正。
        // 従来は `import torch` が通り `cuda.is_available()` が真なら「導入済み(skip)」にしていた。
        // ところがシステム側に古い torch(NumPy 1.x 時代)が残っていると、
        //   ・import は通る    ・cuda.is_available() も真
        // なのに **numpy と噛み合わず実際には使えない**(転写の直前で "Numpy is not available")。
        // 結果、環境タブは「▲要再セットアップ」と正しく警告しているのに、
        // **その指示に従ってセットアップを押しても毎回 skip され、永久に復旧できなかった**
        // (2026-08-04 にまっさら環境で実測。新規利用者がセットアップを中断すると必ず踏む)。
        // ★numpy を経由する操作を1回試して、**使えるかどうか**で判定する。
        // ★ここは **セットアップ側** の判定なので厳しくしてよい。厳しくした結果は「入れ直す」であって、
        //   生成がブロックされるわけではない(K1 の再発にはならない。AG2 と役割が違う)。
        std::string tcode = "import sys; sys.path.insert(0, r'" + spDir + "'); import torch, numpy; torch.from_numpy(numpy.zeros(1)); exit(0 if torch.cuda.is_available() else 1)";
        bool torchCudaOk = RunProcess(Utf8ToWide("\"" + python + "\" -c \"" + tcode + "\""), to, 120000);
        if(!fTorch && torchCudaOk){ // v2.9.0【D】: チェック時はskip判定を無視
            report += "PyTorch(CUDA): \xe5\xb0\x8e\xe5\x85\xa5\xe6\xb8\x88\xe3\x81\xbf (skip)\n";
        } else {
            std::string cudaTag = PickTorchCudaTag(); // v2.9.26【L1】torch用はPython版数も見る
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
            // v2.9.26【L1】GPU があるのに CPU 版へ落ちた理由を必ず見せる。
            if(!g_cudaNoWheelNote.empty()){ report += g_cudaNoWheelNote; DebugLog(g_cudaNoWheelNote); }
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
            // v2.9.35【Z1】cuBLAS/cuDNN の導入はここから外へ出した。
            // ここに置くと、faster-whisper が既に入っている(skip で早期 return)ときに
            // **一度も到達しない**。下の独立したステップで確保する。
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

    // v2.9.35【Z1】faster-whisper(CTranslate2) の CUDA 実行には cuBLAS/cuDNN の DLL が要るが
    // pip の依存には含まれない。従来は InstallWhisperPkg の中で導入していたため、
    // **faster-whisper が既に入っていると skip の早期 return で一度も到達しなかった**
    // (システム側の Python に入っていると PkgOk が true になる)。
    // 何度セットアップしても直らず、v2.9.28 の M1 ガードにより落ちずに CPU へ退避するので
    // 「エラーも出ずにずっと遅い」状態になる。独立したステップとして外に出す。
    // ★既にあるなら入れない(約1GBの無駄を避ける、という元の意図は保つ)。
    //   torch を入れている構成では torch\lib に同梱されているのでそれを使う。
    if((setupBi == 0 || fFaster) && DetectCudaTag() != "cpu"){
        // v2.9.49【AH1】ここには **強制再導入の逃げ道が無かった**。
        // faster-whisper のチェックを付けても haveCublas が真ならそのまま skip されるため、
        // 「cuBLAS はあるが cuDNN が壊れている/欠けている/中断で半端」を踏むと
        // **セットアップを押し直しても永久に直らない**(AG6 と同じ袋小路)。
        // ★判定そのものは変えない(厳しくすると誤検知で1GBを無駄に入れ直す)。
        //   **チェックが付いているときだけ判定を無視する**、という逃げ道を足す。
        // ★あわせて cuDNN 側も見る。従来は cuBLAS しか確認しておらず、
        //   コメントには「cuBLAS/cuDNN の DLL が要る」と書いてあるのに片方しか見ていなかった。
        bool haveCublas = FileExistsU(spDir + "\\torch\\lib\\cublas64_12.dll")
                       || FileExistsU(spDir + "\\nvidia\\cublas\\bin")
                       || FileExistsU(spDir + "\\nvidia_cublas_cu12");
        bool haveCudnn  = FileExistsU(spDir + "\\torch\\lib\\cudnn64_9.dll")
                       || FileExistsU(spDir + "\\torch\\lib\\cudnn64_8.dll")
                       || FileExistsU(spDir + "\\nvidia\\cudnn\\bin")
                       || FileExistsU(spDir + "\\nvidia_cudnn_cu12");
        if(fFaster || !haveCublas || !haveCudnn){
            SetStatus("CUDA \xe3\x83\xa9\xe3\x83\xb3\xe3\x82\xbf\xe3\x82\xa4\xe3\x83\xa0 \xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe4\xb8\xad...");
            std::wstring ncmd = Utf8ToWide("\"" + python + "\" -m pip install nvidia-cublas-cu12 nvidia-cudnn-cu12 --target=\"" + spDir + "\" --quiet");
            std::string nout; bool nok = RunProcess(ncmd, nout, 1800000);
            DebugLog("nvidia cuBLAS/cuDNN install: " + nout.substr(0, 500));
            report += nok ? "CUDA runtime (cuBLAS/cuDNN): OK\n" : "CUDA runtime: WARN\n";
        } else {
            DebugLog("cuBLAS already available -> skip nvidia-* install");
        }
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
    // v2.9.49【AH2】チェックが付いていたら判定を無視して入れ直す。
    // AF1 の判定は K1 再発を避けるためわざと緩い(snapshots に何か1つあれば導入済み)ので、
    // 誤検知を踏んだ利用者が自力で復旧できる道を必ず残しておくこと。
    if(modelExists && !fModel){
        DebugLog("Model already exists: " + mDir + "\\" + mName);
        report += "\xe3\x83\xa2\xe3\x83\x87\xe3\x83\xab(" + mName + "): \xe6\x97\xa2\xe3\x81\xab\xe5\xad\x98\xe5\x9c\xa8 (skip)\n";
    } else {
    std::string dlScript = GetTempDir() + "\\dl_model.py";
    // v2.9.55【AL3】固定パスなので、書けなかったときに**前の版のスクリプトが残って実行される**。
    // 版が変わると引数の意味も変わりうるので、古い物を新しい引数で動かすことになる。
    // AL2 と同じ理由で、実行前に消してから書き、成否を見る。
    DeleteFileU(dlScript);
    bool dlScriptOk = true;
    {
        std::ofstream sf(Utf8ToWide(dlScript));
        if(!sf) dlScriptOk = false;
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
        // v2.9.23【J1】シンボリックリンク権限が無い環境ではDLが WinError 1314 で失敗する。
        // セットアップ経路は「入れ直し」なのでコピー固定で構わない(確実性を優先)。
        sf << "        try:\n";
        sf << "            from huggingface_hub import file_download as _hfd\n";
        sf << "            _hfd.are_symlinks_supported = lambda *a, **k: False\n";
        sf << "        except Exception:\n";
        sf << "            pass\n";
        sf << "        distil_map = {'kotoba-whisper': 'kotoba-tech/kotoba-whisper-v2.0-faster'}\n";
        sf << "        dl_name = distil_map.get(model_name, model_name)\n";
        sf << "        print(f'Downloading {dl_name} (faster-whisper) to {model_dir}')\n";
        sf << "        WhisperModel(dl_name, device='cpu', compute_type='int8', download_root=model_dir)\n";
        sf << "    print('OK')\n";
        sf << "except Exception as e:\n";
        sf << "    print(f'Error: {e}')\n";
        sf << "    import traceback; traceback.print_exc()\n";
        sf << "    sys.exit(1)\n";
        sf.close();                       // v2.9.55【AL3】flush 時の失敗を拾う
        if(!sf) dlScriptOk = false;
    }
    if(!dlScriptOk){
        DebugLog("dl_model.py write failed: " + dlScript);
        // ★セットアップは「入れ直す」で復帰できる側なので、止めずに理由だけ残す。
        //   ここで走らせても下の RunProcess が失敗し、既存の失敗表示に乗る。
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
    SetStatusW(L"Ready (v2.9.75)"); // v2.8.10【I】: SetupThread正常終了時にg_statusSetupが取り残され「fugashi確認中...」等で凍結していたバグの本丸修正+バージョン更新
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
// v2.9.50【AH3】ProbeThread が **二重に走る** のを防ぐ。
// ★起動時は `std::thread(ProbeThread).detach()` で非同期に開始し(数秒かかる)、
//   SetupThread は末尾で `ProbeThread()` を**同期呼び出し**する。
//   起動直後にセットアップを押すと **2つが同時に g_probe を書く**。
//   g_probe の torch / torchUsable / rawOut は atomic ではないので、
//   ・bool が片方の値で上書きされる(表示が実態と食い違う)
//   ・std::string(rawOut) の同時書き込みは**未定義動作**
//   ★AG1(セットアップ直後だけ環境タブが「未導入」のまま生成できない)の
//     根本原因はこれだった可能性が高い。症状の発生条件が完全に一致する。
// ★先行分を **待ってから** 実行する(単に return すると、セットアップ後の再実測が行われない)。
//   probe は数秒なので、待っても実害は無い。
static std::atomic<bool> g_probeRunning{false};

static void ProbeThread(){
    { // 先行している probe があれば、その完了を待ってから始める
        bool expected = false;
        int spin = 0;
        while(!g_probeRunning.compare_exchange_strong(expected, true)){
            expected = false;
            Sleep(50);
            if(++spin > 400) return; // 20秒待っても空かないなら諦める(固まらせない)
        }
    }
    struct RunGuard { ~RunGuard(){ g_probeRunning = false; } } _rg; // どの経路で抜けても必ず戻す
    // v2.9.47【AG5】**実測値を最初に全部クリアする**。どの脱出経路より前に置くこと。
    // ★これが無いと前回の値が残る。特に危険なのは v2.9.46 で足した rawOut で、
    //   py が見つからない等で早期 return すると、**前回の probe 結果を今回のものとして
    //   生成ログへ再掲してしまう**(デバッグのために足したのに誤情報を出す)。
    //   torchUsable も同様に前回値のまま環境タブへ出る。
    // ★Y2(リセットが early return より後ろ)・AD2(_vad_mode の持ち越し)と同じ
    //   「前回状態の残留」型。同じ型を繰り返し踏んでいるので、**新しく状態を持ったら
    //   まずリセット位置を決める**こと。
    g_probe.pyFound = false;
    g_probe.pyMajor = 0; g_probe.pyMinor = 0; g_probe.pyPatch = 0;
    g_probe.torch = false; g_probe.torchUsable = false; g_probe.torchCuda = false;
    g_probe.whisper = false; g_probe.fasterWhisper = false; g_probe.fugashi = false;
    g_probe.rawOut.clear();
    std::string py = GetEffectivePython();
    if(py.empty()){ g_probe.done = true; PostMessageW(g_wnd, WM_PROBE_DONE, 0, 0); return; }
    g_probe.pyFound = true;
    std::string sp = GetSitePackagesDir();
    std::string scr = GetTempDir() + "\\probe_env.py";
    // v2.9.55【AL3】固定パス。書けないと**前の版の probe スクリプトが残って実行される**ため、
    // 実測結果が古い判定基準のものになる。実行前に消してから書き、成否を見る。
    DeleteFileU(scr);
    bool probeScriptOk = true;
    {
        std::ofstream sf(Utf8ToWide(scr));
        if(!sf) probeScriptOk = false;
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
        // v2.9.43【AG1】torch は **import が通っても使えないこと**がある。
        // 実測(2026-08-04 / まっさら環境): システム側の torch 2.2.1+cu118(NumPy 1.x 時代)に対し
        // プラグインが numpy 2.4.6 を入れると、`import torch` は警告だけで通るのに、
        // 実際にテンソルを作る段階で **"FATAL: Numpy is not available"** で落ちる。
        // 転写の直前(whisper.load_model)で死ぬので字幕はゼロ。しかも出るのは英語の NumPy 警告だけで
        // 利用者には原因が分からない。
        // ★__import__ だけでは検出できない。**実際に numpy を経由する操作を1回試す**。
        sf << "def chk_torch():\n";
        sf << "    try:\n";
        sf << "        import torch, numpy\n";
        sf << "        torch.from_numpy(numpy.zeros(1, dtype='float32'))\n";
        sf << "        return '1'\n";
        sf << "    except Exception:\n";
        sf << "        return '0'\n";
        // ★TORCH は従来どおり「import できるか」。これが生成のブロック判定に使われるので
        //   **厳しくしない**(v2.9.43 で厳しくして K1 が再発した)。
        sf << "print('TORCH=' + chk('torch'))\n";
        // ★TORCHUSABLE は「実際に numpy 経由で使えるか」。表示と警告にだけ使う。
        sf << "print('TORCHUSABLE=' + chk_torch())\n";
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
        sf.close();                       // v2.9.55【AL3】flush 時の失敗を拾う
        if(!sf) probeScriptOk = false;
    }
    if(!probeScriptOk){
        // ★probe は表示専用で、生成をブロックしない側。止めずに理由だけ残す。
        DebugLog("probe_env.py write failed: " + scr);
    }
    std::string out;
    RunProcess(Utf8ToWide("\"" + py + "\" \"" + scr + "\" \"" + sp + "\""), out, 180000);
    // v2.9.44【AG2】probe の結果を必ずログに残す。
    // ★これが無かったせいで v2.9.43 の不具合(環境タブが「未導入」のまま生成できない)を
    //   **事後に追えなかった**。スクリプト単体では正常なのに AviUtl2 内でだけ失敗する、
    //   という状況で手がかりがゼロになる。probe は起動のたびに1回だけなので出力は軽い。
    DebugLog("probe: python=" + py + "\n  sp=" + sp + "\n  out=\n" + out);
    // v2.9.46【AG4】生成のたびに再掲できるよう保持する(生成開始時に log が truncate されるため)
    g_probe.rawOut = "probe: python=" + py + "\n  sp=" + sp + "\n  out=\n" + out;
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
        else if(key == "TORCHUSABLE") g_probe.torchUsable = (val == "1"); // v2.9.44【AG2】表示と警告用
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

// v2.9.60【AP1】why を渡すと、受け付けなかった理由を返す(既定 nullptr = 従来と同じ)。
static bool LoadTemplate(const std::string& path, std::string* why = nullptr){
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
    // v2.9.60【AP1】複数オブジェクトのエイリアスを弾く。
    // ★SDK の契約(plugin2.h 125行): create_object_from_alias は
    //   「複数オブジェクトのエイリアスデータの場合は先頭のオブジェクトのハンドルが
    //     返却されます ※**オブジェクトは全て作成されます**」。
    //   つまり **字幕1つにつきエイリアス内の全オブジェクトが作られる**のに、
    //   戻り値は1つなので Pass2 は replaced を1しか数えない。
    //   配置側はレイヤーを「字幕1つ = 1オブジェクト」で詰めているので、
    //   余分なオブジェクトが確保していないレイヤーに載って利用者の物と衝突する。
    //   **ログは「replaced: 95 failed: 0」と出るのにタイムラインは壊れる。**
    // ★形式は実物で確認済み(手元の .object 194件を分類):
    //     単一 : [Object] + [Object.N]   … N はエフェクト番号 (191件)
    //     複数 : [0] [0.N] [1] [1.N] …   … 先頭の数字がオブジェクト番号 (3件・実在する)
    //   よって **^[数字]$ の行数**を数えればオブジェクト数が分かる。
    // ★R1 と同じ「選択時に受け付けず理由を伝える」形にする(生成をブロックする側ではない)。
    {
        int objCount = 0;
        std::istringstream os(cleaned);
        std::string ol;
        while(std::getline(os, ol)){
            while(!ol.empty() && (ol.back() == '\r' || ol.back() == ' ')) ol.pop_back();
            if(ol.size() >= 3 && ol.front() == '[' && ol.back() == ']'){
                std::string inner = ol.substr(1, ol.size() - 2);
                bool allDigit = !inner.empty();
                for(char c : inner) if(c < '0' || c > '9'){ allDigit = false; break; }
                if(allDigit) objCount++;
            }
        }
        if(objCount >= 2){
            DebugLog("Template REJECTED (multi-object alias: " + std::to_string(objCount) + "): " + path);
            if(why) *why = "\x0a\x0a\xe8\xa4\x87\xe6\x95\xb0\xe3\x81\xae\xe3\x82\xaa\xe3\x83\x96\xe3\x82\xb8\xe3\x82\xa7\xe3\x82\xaf\xe3\x83\x88\xe3\x82\x92\xe5\x90\xab\xe3\x82\x80\xe3\x82\xa8\xe3\x82\xa4\xe3\x83\xaa\xe3\x82\xa2\xe3\x82\xb9\xe3\x81\xa7\xe3\x81\x99\xe3\x80\x82\x0a\xe5\xad\x97\xe5\xb9\x95\x31\xe3\x81\xa4\xe3\x81\xab\xe3\x81\xa4\xe3\x81\x8d\xe5\x85\xa8\xe9\x83\xa8\xe3\x81\xae\xe3\x82\xaa\xe3\x83\x96\xe3\x82\xb8\xe3\x82\xa7\xe3\x82\xaf\xe3\x83\x88\xe3\x81\x8c\xe4\xbd\x9c\xe3\x82\x89\xe3\x82\x8c\xe3\x81\xa6\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x86\xe3\x81\x9f\xe3\x82\x81\xe3\x80\x81\xe5\x8f\x97\xe3\x81\x91\xe4\xbb\x98\xe3\x81\x91\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\x0a\xe3\x83\x86\xe3\x82\xad\xe3\x82\xb9\xe3\x83\x88\xe3\x82\xaa\xe3\x83\x96\xe3\x82\xb8\xe3\x82\xa7\xe3\x82\xaf\xe3\x83\x88\xe3\x82\x92\x31\xe3\x81\xa4\xe3\x81\xa0\xe3\x81\x91\xe9\x81\xb8\xe3\x82\x93\xe3\x81\xa7\xe4\xbf\x9d\xe5\xad\x98\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82";
            return false;
        }
    }

    // Verify text key exists
    std::string textKey = "\xe3\x83\x86\xe3\x82\xad\xe3\x82\xb9\xe3\x83\x88=";
    if(cleaned.find(textKey) == std::string::npos){
        // v2.9.27【R1】従来は raw にフォールバックして**受け入れていた**。しかし Pass2 は
        // テキスト項目が無いと差し替えを行わない(`if(pos != npos)` の外を通る)ため、
        // **全字幕がテンプレートの文言のまま**作られ、文字起こしが丸ごと捨てられる。
        // しかも Pass2 は成功扱い(replaced++)でログも failed:0 になり、**誰も気づけない**。
        // ファイル選択は "All *.*" も許すので、テキスト以外の .object や無関係な
        // ファイルを選べてしまう。受け入れずに false を返して理由を伝える。
        // Pass1 の素の字幕はそのまま残るので、内容は失われない。
        DebugLog("Template REJECTED (no text key): " + path);
        return false;
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
    // v2.9.28【S1】再生速度(%)。**警告に使うだけで、切り出しには一切使わない**。
    // 0 は「速度指定なし」とみなす(実データ496件中2件が 0.00)。
    double speed;
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
                // v2.9.28【S1】再生速度を読む。読めなければ 100(=等倍)として扱う。
                // 値は 再生位置 と同じく "100.00" または "100.00,...,直線移動,0" の形で、
                // atof がカンマで止まるので先頭値が取れる。
                LPCSTR spdVal = es->get_object_item_value(obj, L"\x52d5" L"\x753b\x30d5\x30a1\x30a4\x30eb", L"\x518d\x751f\x901f\x5ea6");
                if(!spdVal) spdVal = es->get_object_item_value(obj, L"\x97f3\x58f0\x30d5\x30a1\x30a4\x30eb", L"\x518d\x751f\x901f\x5ea6");
                c.speed = spdVal ? atof(spdVal) : 100.0;
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
        // v2.9.15【監査④】音声トラックが無いだけの動画を「ffmpegの不具合」と誤診させない。
        // 実測: 映像のみのmp4に対し ffmpeg は終了コード -22 と
        // "Output file does not contain any stream" を返す。
        // これを拾って呼び出し側に伝え、専用のメッセージを出す。
        if(procOut.find("does not contain any stream") != std::string::npos
           || procOut.find("Output file is empty") != std::string::npos)
            g_extractNoAudio++;
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
            // v2.9.63【AU1】禁則の後退が **進行を殺していた**。
            // 窓(maxChars文字)の中が全部禁則文字だと、この while が cutAfter を start まで
            // 戻し切り、下の進行保証 `nextStart = start + 1` に落ちて **1文字ずつ**に割れる。
            // 実測: 「あ」+「ー」x200 -> 187パート / 全部2文字以下 / 全部6フレーム未満 /
            //       最短1フレーム。実機のライブ配信素材で 455件の字幕に膨張した本体がこれ。
            // ★whisper は叫び声を「ー」の巨大連続として転写することがある(同一素材5回で
            //   1回、最長 444文字)。「ー」は noHead に入っているので必ずここを踏む。
            // → 禁則を満たす切り位置が無ければ **禁則を諦めて上限で切る**。
            //   禁則は「読みやすさ」の都合であって、守るために字幕を壊してよい規則ではない。
            // ★普通の文では発火しない: 窓の中の全文字が禁則文字のときだけ後退が start に届く。
            //   実データ6本303行で切り位置が変わったのは、事故が起きている2行だけだった。
            int back = cutAfter;
            while(back > start && (back+1 < total) && noHead(back+1)) back--;
            cutAfter = (back > start) ? back : (hardEnd - 1);
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

    // v2.9.17: フェーズ別の所要時間。どこに時間を使っているかが見えないと最適化を外す。
    ULONGLONG _tExtract0 = GetTickCount64();
    // Extract audio for all clips
    g_extractNoAudio = 0; // v2.9.15: 「音声トラックが無い」クリップ数を数え直す
    for(size_t ci = 0; ci < g_tlClips.size(); ci++){
        std::string w = tmp + "whisper_fw_" + std::to_string(ci) + ".wav";
        wavs.push_back(ExtractAudio(g_tlClips[ci], g_projectRate, w));
    }
    bool anyW = false;
    for(auto& w : wavs) if(!w.empty()) anyW = true;
    if(!anyW){
        // v2.9.15【監査④】原因で文言を分ける。音声が無いだけなのに ffmpeg を疑わせない。
        MsgBox(g_wnd,
            (g_extractNoAudio > 0)
              ? "\xe9\x81\xb8\xe6\x8a\x9e\xe3\x81\x97\xe3\x81\x9f\xe3\x82\xaf\xe3\x83\xaa\xe3\x83\x83\xe3\x83\x97\xe3\x81\xab\xe9\x9f\xb3\xe5\xa3\xb0\xe3\x81\x8c\xe5\x90\xab\xe3\x81\xbe\xe3\x82\x8c\xe3\x81\xa6\xe3\x81\x84\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n"
                "\xe9\x9f\xb3\xe5\xa3\xb0\xe3\x83\x88\xe3\x83\xa9\xe3\x83\x83\xe3\x82\xaf\xe3\x81\xae\xe3\x81\x82\xe3\x82\x8b\xe5\x8b\x95\xe7\x94\xbb\xe3\x81\x8b\xe3\x80\x81\xe9\x9f\xb3\xe5\xa3\xb0\xe3\x83\x95\xe3\x82\xa1\xe3\x82\xa4\xe3\x83\xab\xe3\x82\x92\xe9\x85\x8d\xe7\xbd\xae\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82"
              : "\xe9\x9f\xb3\xe5\xa3\xb0\xe6\x8a\xbd\xe5\x87\xba\xe5\xa4\xb1\xe6\x95\x97\xe3\x80\x82\x66\x66\x6d\x70\x65\x67\xe3\x82\x92\xe7\xa2\xba\xe8\xaa\x8d\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82",
            "Error", MB_OK|MB_ICONERROR);
        cleanup();
        return false;
    }

    std::string md = GetModelsDir();

    ULONGLONG _tExtract1 = GetTickCount64();
    // Write batch JSON
    // v2.9.55【AL2】従来はここが **is_open すら見ていなかった**。しかも bp は固定パスで、
    // Q1 の対策(実行前に消す)は**結果ファイルと .err にしか掛かっていない**。
    // そのため書き込みに失敗すると **前回の whisper_batch.json がそのまま残り**、
    // Python は**前の動画のクリップ一覧**を読んで転写する。Q1(別の動画の字幕)と同じ形が、
    // 出力側ではなく**入力側**に残っていた。
    // → ①実行前に必ず消す ②開けたか・書けたかを両方見る ③失敗したら止めて理由を出す。
    DeleteFileU(bp);
    bool batchOk = true;
    {
        std::ofstream bf(Utf8ToWide(bp));
        if(!bf) batchOk = false;
        char tempStr[32]; sprintf_s(tempStr, "%.2f", temp);
        bf << "{\n  \"model\": \"" << mn[mi] << "\",\n  \"language\": \"" << lc[li] << "\",\n  \"device\": \"" << dn[di] << "\",\n  \"backend\": \"" << bn[bi] << "\",\n  \"beam_size\": " << beamSize << ",\n  \"temperature\": " << tempStr << ",\n  \"no_prev_text\": " << (noPrevText ? "true" : "false") << ",\n  \"word_timestamps\": " << (wordTs ? "true" : "false") << ",\n  \"rep_penalty\": " << (repPenalty ? "true" : "false") << ",\n  \"vad_filter\": " << (vadFilter ? "true" : "false") << ",\n  \"prompt\": \"";
        // JSON-escape prompt text
        bf << JsonEscape(promptText); // v2.9.16
        bf << "\",\n  \"hotwords\": \"";
        // JSON-escape hotwords text
        bf << JsonEscape(hotwordsText); // v2.9.16
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
            fesc = JsonEscape(ffm); // v2.9.16
            bf << ",\n  \"ffmpeg\": \"" << fesc << "\"";
        }
        bf << ",\n  \"clips\": [\n";
        bool first = true;
        for(size_t ci = 0; ci < g_tlClips.size(); ci++){
            if(wavs[ci].empty()) continue;
            if(!first) bf << ",\n"; first = false;
            std::string esc;
            esc = JsonEscape(wavs[ci]); // v2.9.16
            // v2.9.20【F1】元の g_tlClips 上の index を明示する。空wavのクリップを詰めて書くため、
            // Python 側 enumerate の ci は「詰めた後」の index になり、それを C++ が元の配列
            // g_tlClips[ci] に当ててクランプしていた。音声の無いクリップの後ろに本編があると
            // 別クリップの範囲でクランプされ、範囲外の字幕が黙って全滅する(再現ハーネスで 0/2 件を確認)。
            bf << "    {\"idx\": " << ci << ", \"wav\": \"" << esc << "\", \"timeline_start\": " << g_tlClips[ci].timelineStart
               << ", \"timeline_end\": " << g_tlClips[ci].timelineEnd << ", \"fps\": " << g_projectRate << "}";
        }
        bf << "\n  ]\n}\n";
        // v2.9.55【AL2】★close() まで済ませてから確かめる。失敗が出るのは flush 時。
        bf.close();
        if(!bf) batchOk = false;
    }
    if(!batchOk){
        DebugLog("Batch JSON write failed: " + bp);
        MsgBox(g_wnd, "\xe9\x9f\xb3\xe5\xa3\xb0\xe8\xaa\x8d\xe8\xad\x98\xe3\x81\xb8\xe6\xb8\xa1\xe3\x81\x99\xe6\x83\x85\xe5\xa0\xb1\xe3\x82\x92\xe6\x9b\xb8\xe3\x81\x8d\xe8\xbe\xbc\xe3\x82\x81\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x81\xa7\xe3\x81\x97\xe3\x81\x9f\xe3\x80\x82\x0a\x0a\xe3\x83\x97\xe3\x83\xa9\xe3\x82\xb0\xe3\x82\xa4\xe3\x83\xb3\xe3\x83\x95\xe3\x82\xa9\xe3\x83\xab\xe3\x83\x80\xe3\x81\xae\xe7\xa9\xba\xe3\x81\x8d\xe5\xae\xb9\xe9\x87\x8f\xe3\x81\xa8\xe3\x82\xa2\xe3\x82\xaf\xe3\x82\xbb\xe3\x82\xb9\xe6\xa8\xa9\xe3\x82\x92\xe7\xa2\xba\xe8\xaa\x8d\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82\x0a\x28\xe3\x81\x9d\xe3\x81\xae\xe3\x81\xbe\xe3\x81\xbe\xe5\xae\x9f\xe8\xa1\x8c\xe3\x81\x99\xe3\x82\x8b\xe3\x81\xa8\xe3\x80\x81\xe5\x89\x8d\xe5\x9b\x9e\xe3\x81\xae\xe5\x8b\x95\xe7\x94\xbb\xe3\x81\xae\xe5\x86\x85\xe5\xae\xb9\xe3\x81\xa7\xe5\xad\x97\xe5\xb9\x95\xe3\x81\x8c\xe4\xbd\x9c\xe3\x82\x89\xe3\x82\x8c\xe3\x82\x8b\xe6\x81\x90\xe3\x82\x8c\xe3\x81\x8c\xe3\x81\x82\xe3\x82\x8b\xe3\x81\x9f\xe3\x82\x81\xe4\xb8\xad\xe6\xad\xa2\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x97\xe3\x81\x9f\x29",
               "Error", MB_OK|MB_ICONERROR);
        cleanup(); return false;
    }

    SetStatus(std::string("[") + bn[bi] + "] \xe6\x96\x87\xe5\xad\x97\xe8\xb5\xb7\xe3\x81\x93\xe3\x81\x97\xe4\xb8\xad...");
    SetProgress(40);
    // v2.9.53【AJ1】ヘルパーを書けなかったら、そのまま走らせない。
    // ★ここで止めるのは安全側 — 書き込みは「更新が必要だったのに失敗した」ときだけ失敗し、
    //   そのときファイルは既に truncate されて壊れている。走らせても必ず失敗し、
    //   しかも python 側の無関係なエラーとして出るので原因が分からなくなる。
    //   K1(生成をブロックする判定を厳しくして詰まらせた)とは違い、誤検知の余地が無い。
    if(!EnsurePyHelper()){
        MsgBox(g_wnd, "\xe9\x9f\xb3\xe5\xa3\xb0\xe8\xaa\x8d\xe8\xad\x98\xe7\x94\xa8\xe3\x81\xae\xe3\x82\xb9\xe3\x82\xaf\xe3\x83\xaa\xe3\x83\x97\xe3\x83\x88\xe3\x82\x92\xe6\x9b\xb8\xe3\x81\x8d\xe8\xbe\xbc\xe3\x82\x81\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x81\xa7\xe3\x81\x97\xe3\x81\x9f\xe3\x80\x82\x0a\x0a\xe3\x83\x97\xe3\x83\xa9\xe3\x82\xb0\xe3\x82\xa4\xe3\x83\xb3\xe3\x83\x95\xe3\x82\xa9\xe3\x83\xab\xe3\x83\x80\xe3\x81\xab\xe6\x9b\xb8\xe3\x81\x8d\xe8\xbe\xbc\xe3\x82\x81\xe3\x81\xaa\xe3\x81\x84\xe7\x8a\xb6\xe6\x85\x8b\xe3\x81\xa7\xe3\x81\x99\xe3\x80\x82\x0a\xe7\xa9\xba\xe3\x81\x8d\xe5\xae\xb9\xe9\x87\x8f\xe3\x80\x81\xe3\x82\xa2\xe3\x82\xaf\xe3\x82\xbb\xe3\x82\xb9\xe6\xa8\xa9\xe3\x80\x81\xe3\x82\xa6\xe3\x82\xa4\xe3\x83\xab\xe3\x82\xb9\xe5\xaf\xbe\xe7\xad\x96\xe3\x82\xbd\xe3\x83\x95\xe3\x83\x88\xe3\x81\xae\xe9\x99\xa4\xe5\xa4\x96\xe8\xa8\xad\xe5\xae\x9a\xe3\x82\x92\xe7\xa2\xba\xe8\xaa\x8d\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82",
               "Error", MB_OK|MB_ICONERROR);
        cleanup(); return false;
    }
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
    // v2.9.14【監査③】固定600秒では長尺やCPU実行で足りず、途中で強制終了して全部失う。
    // 実測: ffmpeg抽出は1284倍速なので抽出は問題ないが、推論はGPUのRTF 0.04に対し
    // CPUは1.0前後。30分動画をCPUで回すと1800秒かかり10分の上限を大きく超える。
    // 音声長の2倍 + 5分を目安にし、無限待ちを避けるため上限1時間で頭打ちにする。
    double totalSec = 0.0;
    for(size_t ti = 0; ti < g_tlClips.size(); ti++)
        totalSec += (double)(g_tlClips[ti].timelineEnd - g_tlClips[ti].timelineStart)
                    / (double)(g_projectRate > 0 ? g_projectRate : 30);
    double tmoD = totalSec * 2000.0 + 300000.0;
    if(tmoD < 600000.0) tmoD = 600000.0;
    if(tmoD > 3600000.0) tmoD = 3600000.0;
    DWORD tmo = (DWORD)tmoD;
    DebugLog("Transcribe timeout: " + std::to_string(tmo / 1000) + "s (audio "
             + std::to_string((int)totalSec) + "s)");
    // v2.9.27【Q1】結果ファイルと .err を **実行の前に** 消す。
    // Python は結果を**最後にしか書かない**(開始時に truncate しない)ため、途中で殺されると
    // (タイムアウト / AviUtl2 を閉じる / クラッシュ) 前回のファイルがそのまま残り、
    // 下の rf.is_open() が成功して **別の動画の字幕をそのまま配置してしまう**。
    // cleanup() は実行の「後」にしか呼ばれず、異常終了では走らない。
    // ★異常終了は実際に起きている: temp に中断された実行の whisper_fw_*.wav が残っていた。
    // .err も古い内容がエラー表示に混ざるので一緒に消す。
    DeleteFileU(op); DeleteFileU(errP);
    ULONGLONG _tPy0 = GetTickCount64();
    // v2.9.52【AI2】転写の子だけ Job に入れる(AviUtl2 が死んだら道連れ)。
    bool pyOk = RunProcess(wCmd, pyOut, tmo, true);
    ULONGLONG _tPy1 = GetTickCount64();
    DebugLog("Timing: extract=" + std::to_string(_tExtract1 - _tExtract0) + "ms"
             + " python=" + std::to_string(_tPy1 - _tPy0) + "ms"
             + " (json=" + std::to_string(_tPy0 - _tExtract1) + "ms)");
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
    CollectClipErrors(errP); // v2.9.52【AI1】握り潰された転写の理由を拾う(表示用)
    // v2.9.36【Z2】検出は Python 側で済んでいる(M1)。ここでは拾うだけで、
    // C++ 側に同じ判定を書かない(「同じ計算が2箇所」を避ける)。
    if(errC.find("GPU found but cuBLAS DLL missing") != std::string::npos)
        g_cublasMissing = 1;

    // Read result file
    std::ifstream rf(Utf8ToWide(op));
    if(!rf.is_open()){
        std::string msg;
        // v2.9.57【AM2】日本語説明は NumpyHintIfNeeded() に集約した(呼び出しは3箇所)。
        msg += NumpyHintIfNeeded(pyOut, errC, "");
        msg += "\xe7\xb5\x90\xe6\x9e\x9c\xe3\x83\x95\xe3\x82\xa1\xe3\x82\xa4\xe3\x83\xab\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n\n";
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
                // v2.9.45【AG3】numpy と torch が噛み合っていないときは、生の英語の前に
                // 何が起きたか・どう直すかを日本語で出す。
                // ★v2.9.43 では **結果ファイルが開けない経路**(!rf.is_open)にだけ入れていたが、
                //   Python は例外を捕まえて結果ファイルへ `ERROR|Unexpected error|...` を**書く**ので
                //   ファイルは開ける。よってあちらは通らず、利用者には英語だけが出ていた
                //   (実測: まっさら環境で "Unexpected error" のみ表示され、日本語の案内が出なかった)。
                //   **エラーを出す経路は2つある。両方に置くこと。**
                // v2.9.57【AM2】NumpyHintIfNeeded() に集約(呼び出しは3箇所)
                std::string em = NumpyHintIfNeeded(pyOut, errC, line);
                em += line.substr(6);
                MsgBox(g_wnd, em, "Error", MB_OK|MB_ICONERROR);
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
            std::string autoTag = PickTorchCudaTag(); // v2.9.26【L1】
            std::string autoIdx = (autoTag == "cpu")
                ? "https://download.pytorch.org/whl/cpu"
                : ("https://download.pytorch.org/whl/" + autoTag);
            std::wstring torchCmd = Utf8ToWide("\"" + python + "\" -m pip install torch --upgrade --target=\"" + spDir + "\" --index-url " + autoIdx);
            // v2.9.26【L1】従来は RunProcess の戻り値を見ておらず、pip が失敗しても素通りして
            // そのまま再試行し「リトライ失敗」とだけ出ていた(主経路 SetupThread は v2.9.7 で
            // 確認済みだが、この復旧経路だけ取り残されていた)。PkgOk は SetupThread 内の
            // ラムダでここからは使えないため、終了コードだけを見る。
            std::string torchOut; bool tok2 = RunProcess(torchCmd, torchOut, 1800000);
            DebugLog(std::string("Auto torch install (") + autoTag + "): "
                     + (tok2 ? "exit0 " : "FAILED ") + torchOut.substr(0, 500));
            if(!g_cudaNoWheelNote.empty()) DebugLog(g_cudaNoWheelNote);
        }

        if(!pipOk){
            MsgBox(g_wnd, installPkg + " \xe3\x81\xae\xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x83\x88\xe3\x83\xbc\xe3\x83\xab\xe3\x81\xab\xe5\xa4\xb1\xe6\x95\x97\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x97\xe3\x81\x9f\xe3\x80\x82\n\n" + pipOut, "Error", MB_OK|MB_ICONERROR);
            cleanup(); return false;
        }

        // Retry transcription
        SetStatus("\xe5\x86\x8d\xe8\xa9\xa6\xe8\xa1\x8c\xe4\xb8\xad...");
        SetProgress(60);
        std::string pyOut2;
        // v2.9.57【AM3】本実行は Q1 の対策で結果と .err を**実行前に消して**いるのに、
        // 再試行は消していなかった。再試行の python が起動に失敗すると、
        // **本実行が書いた古い ERROR 行がそのまま読まれ**、既に解消した
        // 「〜 not installed」を今回の原因として表示する。Q1 と同じ形の片割れ。
        DeleteFileU(op); DeleteFileU(errP);
        // v2.9.57【AM4】本実行のループは ERROR 行に当たるまでに一部の行を push しうる。
        // 再試行はファイルを**先頭から読み直す**ので、消さないと同じ字幕が二重に積まれる。
        // ★AC1(部分push後に別方式が再pushして二重になる)と同じ形。
        //   失敗した実行の部分結果は**まとめて破棄する**のが既定の方針。
        g_segs.clear();
        // v2.9.41【AE1】ここだけ 10分固定だった。本実行は上で tmo(音声長×2+5分、上限1時間)を
        // 計算しているのに、再試行はそれを使っていない。**同じ計算が2箇所にあって片方だけ**という
        // N1・V2・W1 と同じ型。長尺で1回目が落ちた場合、再試行は10分で確実に足りず、
        // 「導入し直したのにまた失敗する」状態になる。本実行と同じ tmo を使う。
        bool pyOk2 = RunProcess(wCmd, pyOut2, tmo, true); // v2.9.52【AI2】再試行も Job に入れる
        DebugLog("Retry Python exit=" + std::string(pyOk2 ? "0" : "nonzero") + "\n" + pyOut2);
        // v2.9.52【AI1】再試行側でも理由を拾い直す。ここを抜かすと本実行の理由が残り、
        // **前回状態の残留**(Y2/AD2/AG5 と同じ型)で誤った原因を表示することになる。
        CollectClipErrors(errP);

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
                // v2.9.57【AM2】ここに日本語説明が無く、**英語だけが出ていた**。
                // 再試行は自動導入の直後に走るので numpy 不整合が最も出やすい場面。
                std::string em2 = NumpyHintIfNeeded(pyOut2, errC, line2);
                em2 += line2.substr(6);
                MsgBox(g_wnd, em2, "Error", MB_OK|MB_ICONERROR);
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
    // v2.9.34【Y2】前回の実行の結果を**最初に**捨てる。
    // 従来はタイムラインスキャンの直前で clear していたが、それより前に
    // 「ffmpegが無い」「pythonが無い」の early return が2つあり、そこを通ると
    // g_segs が前回のまま残っていた。すると SRTエクスポートが g_segs.empty() を
    // false と見て**前回の字幕を今回の結果として書き出す**。
    // ★Q1(結果ファイルの残骸で別の動画の字幕が配置される)と同じ型。
    //   どの経路で抜けても残らないよう、判定より前に置く。
    g_tlClips.clear();
    g_segs.clear();
    g_cublasMissing = 0; // v2.9.36【Z2】前回の判定を持ち越さない(Y2 と同じ理由)
    DWORD startTick = GetTickCount();
    // best-effort: ログの truncate。失敗しても本処理に影響しない
    {std::ofstream f(Utf8ToWide(GetPluginDir() + "\\whisper_debug.log"), std::ios::trunc); f << "=== Whisper Subtitle v2.9.75 ===\n";}
    // v2.9.46【AG4】truncate の**直後**に probe の結果を再掲する。
    // ★これが無いと、probe は起動時にしか書かないので生成すると消え、
    //   「不具合が起きた後にログを見る」場面で手がかりがゼロになる(v2.9.43 の調査で実際に詰まった)。
    if(g_probe.done && !g_probe.rawOut.empty()) DebugLog(g_probe.rawOut);
    else DebugLog("probe: (未完了。起動直後に生成した可能性)");
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
    // v2.9.51【AH4】生成中に設定を変えられても混ざらないよう、ここで読み切って**関数スコープ**に持つ。
    // 生成は数分かかるのに SetBusy() が無効化するのは Generate/Setup ボタンだけなので、
    // 後半で UI を読み直すと同じ1回の生成で冒頭の値と後半の値が食い違う。
    // ★下のログ用ブロック `{ }` の**外**に置くこと(中に入れると後半から見えずコンパイルが通らない。実際に一度やった)。
    // ★範囲のガードも読んだ直後に一度だけ掛ける(後段では再クランプしない)。
    int mgOnEarly = (SendMessageA(g_chkMergeSeg, BM_GETCHECK, 0, 0) == BST_CHECKED) ? 1 : 0;
    // v2.9.54【AK1】書式テンプレートも冒頭で読み切る。AH4 の取りこぼし。
    // ★配置(Pass2)は g_templateContent を**生成の最後**に読む一方、ログは**最初**に書く。
    //   生成中はテンプレの「選択」「解除」ボタンが押せる(SetBusy が無効化するのは
    //   Generate/Setup だけ)ので、変えると:
    //     ・その回の字幕が**新しいテンプレ**で配置され、ログには古い方が残る(AG5 と同じ誤情報)
    //     ・「解除」を押すと Pass2 が丸ごと走らず**字幕が素のまま**になる
    //     ・UI スレッドの代入中に Pass2 がコピーすると **std::string の data race**(未定義動作)
    // ★テンプレは全字幕の見た目を決めるので、linger/lead より混ざったときの影響が大きい。
    std::string tplContentEarly = g_templateContent;
    std::string tplPathEarly    = g_templatePath;
    double lingerSecEarly = 0.0, leadSecEarly = 0.0;
    {
        char lb[16] = {}; GetWindowTextA(g_lingerEdit, lb, sizeof(lb));
        lingerSecEarly = atof(lb);
        if(lingerSecEarly < 0) lingerSecEarly = 0;
        if(lingerSecEarly > 10) lingerSecEarly = 10;
        char eb[16] = {}; GetWindowTextA(g_leadEdit, eb, sizeof(eb));
        leadSecEarly = atof(eb);
        if(leadSecEarly < 0) leadSecEarly = 0;
        if(leadSecEarly > 5) leadSecEarly = 5;
    }
    // v2.8.2: 生成時の各設定をログに記録 (設定ごとの結果比較用)。どの設定でこの字幕群を作ったか一目で分かる。
    {
        int bkIdx = (int)SendMessageA(g_backendCombo, CB_GETCURSEL, 0, 0);
        int mdIdx = (int)SendMessageA(g_modelCombo, CB_GETCURSEL, 0, 0);
        int mgOn = mgOnEarly; // v2.9.51【AH4】ログも読み切った値を使う(実処理と食い違わせない)
        // v2.9.52【AI3】ここだけ UI を読み直したままだった(AH4 の取りこぼし)。
        // すぐ上の mgOn は読み切った値に直っているのに lead だけ残っており、
        // ★実際に使う値(leadSecEarly)と**ログに出る値が食い違いうる**。
        //   AG5 と同じ「デバッグのために足したものが誤情報を出す」型なので読み切った値を出す。
        char leadLogBuf[32] = {};
        snprintf(leadLogBuf, sizeof(leadLogBuf), "%.1f", leadSecEarly);
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
    // v2.9.14【監査③】上限が無く、誤って大きな値を入れると事実上ハングしていた
    int beamSize = atoi(qBuf); if(beamSize <= 0) beamSize = 5; if(beamSize > 20) beamSize = 20;
    char tBuf2[16] = {}; GetWindowTextA(g_tempEdit, tBuf2, sizeof(tBuf2));
    // v2.9.14【監査③】whisper の想定は 0〜1。範囲外は丸める
    float temp = (float)atof(tBuf2); if(temp < 0) temp = 0; if(temp > 1.0f) temp = 1.0f;
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
        // v2.9.1【文字化け修正】ここが whisper_batch.json に直接書かれる本丸。
        // ANSI版だとCP932バイト列がJSONに入り、Python側の json.load で
        // 「'utf-8' codec can't decode byte 0x94」となって**字幕生成が丸ごと失敗**していた(Issue #5)。
        int pLen = GetWindowTextLengthW(g_promptEdit);
        if(pLen > 0){
            std::wstring w(pLen + 1, 0);
            GetWindowTextW(g_promptEdit, &w[0], pLen + 1);
            w.resize(pLen);
            promptText = WideToUtf8(w);
        }
    }
    // v2.8 settings continued
    bool batched = (SendMessageA(g_chkBatched, BM_GETCHECK, 0, 0) == BST_CHECKED);
    std::string hotwordsText;
    {
        // v2.9.1【文字化け修正】prompt と同じ(Issue #5)
        int hLen = GetWindowTextLengthW(g_hotwordsEdit);
        if(hLen > 0){
            std::wstring w(hLen + 1, 0);
            GetWindowTextW(g_hotwordsEdit, &w[0], hLen + 1);
            w.resize(hLen);
            hotwordsText = WideToUtf8(w);
        }
    }
    // v2.9.54【AK1】ログも読み切った値を使う(実際に配置で使う物と食い違わせない)
    DebugLog("Template: " + (tplContentEarly.empty() ? std::string("none") : tplPathEarly));

    std::string ffmpeg = GetEffectiveFFmpeg();
    if(ffmpeg.empty()){
        MsgBox(g_wnd,
            "ffmpeg.exe\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n\n"
            "\xe3\x80\x8c" "ffmpeg\xe9\x81\xb8\xe6\x8a\x9e\xe3\x80\x8d\xe3\x83\x9c\xe3\x82\xbf\xe3\x83\xb3\xe3\x81\xa7\xe6\x8c\x87\xe5\xae\x9a\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84",
            "Error", MB_OK|MB_ICONERROR);
        SetStatus("Ready (v2.9.75)"); SetProgress(0); SetBusy(false); return;
    }
    std::string python = GetEffectivePython();
    if(python.empty()){
        MsgBox(g_wnd,
            "Python\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n\n"
            "\xe3\x80\x8c\xe5\x88\x9d\xe6\x9c\x9f\xe8\xa8\xad\xe5\xae\x9a\xe3\x80\x8d\xe3\x81\xa7\xe3\x82\xbb\xe3\x83\x83\xe3\x83\x88\xe3\x82\xa2\xe3\x83\x83\xe3\x83\x97\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84",
            "Error", MB_OK|MB_ICONERROR);
        SetStatus("Ready (v2.9.75)"); SetProgress(0); SetBusy(false); return;
    }

    SetStatus("\xe3\x82\xbf\xe3\x82\xa4\xe3\x83\xa0\xe3\x83\xa9\xe3\x82\xa4\xe3\x83\xb3\xe3\x82\xb9\xe3\x82\xad\xe3\x83\xa3\xe3\x83\xb3\xe4\xb8\xad...");
    SetProgress(10);
    ScanParam sp; // v2.9.34【Y2】clear は関数冒頭へ移動した
    sp.clips = &g_tlClips;
    sp.rate = 30;
    sp.maxLayer = 0;
    // v2.9.59【AO1】SDK の call_edit_section_param は **bool を返す**(plugin2.h 296行)。
    // 従来は `if(g_edit)` でハンドルの null だけ見て、**戻り値を捨てていた**。
    // 編集セクションに入れなかった場合、コールバックは一度も走らないのに処理は続く。
    // ここでは clips が空のまま進み、下の「クリップが見つかりません」に落ちて
    // **実際とは違う理由**を利用者に案内することになる。
    bool scanOk = (g_edit && g_edit->call_edit_section_param(&sp, ScanCallback));
    if(!scanOk){
        DebugLog("call_edit_section_param(Scan) failed");
        // v2.9.61【AS1】失敗の**具体的な条件**を書く。SDK(plugin2.h 688行)に
        // 「編集が出来ない場合(**出力中等**)に失敗します」と明記されており、
        // EDIT_STATE_PLAY(プレビュー再生中) / EDIT_STATE_SAVE(ファイル出力中) が定義されている。
        // ★プレビューを流したまま生成を押す、は**ごく普通の操作**。
        //   従来は「編集が行える状態か確認してください」としか出ず、何を止めればよいか伝わらない。
        MsgBox(g_wnd, "\xe3\x82\xbf\xe3\x82\xa4\xe3\x83\xa0\xe3\x83\xa9\xe3\x82\xa4\xe3\x83\xb3\xe3\x82\x92\xe8\xaa\xad\xe3\x81\xbf\xe5\x8f\x96\xe3\x82\x8c\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x81\xa7\xe3\x81\x97\xe3\x81\x9f\xe3\x80\x82\x0a\x0a\xe3\x83\x97\xe3\x83\xac\xe3\x83\x93\xe3\x83\xa5\xe3\x83\xbc\xe5\x86\x8d\xe7\x94\x9f\xe4\xb8\xad\xe3\x83\xbb\xe3\x83\x95\xe3\x82\xa1\xe3\x82\xa4\xe3\x83\xab\xe5\x87\xba\xe5\x8a\x9b\xe4\xb8\xad\xe3\x81\xaf\xe8\xaa\xad\xe3\x81\xbf\xe5\x8f\x96\xe3\x82\x8c\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\x0a\xe5\x81\x9c\xe6\xad\xa2\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8b\xe3\x82\x89\xe3\x80\x81\xe3\x82\x82\xe3\x81\x86\xe4\xb8\x80\xe5\xba\xa6\xe5\xae\x9f\xe8\xa1\x8c\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82\x0a\x28\xe3\x83\x97\xe3\x83\xad\xe3\x82\xb8\xe3\x82\xa7\xe3\x82\xaf\xe3\x83\x88\xe3\x82\x92\xe9\x96\x8b\xe3\x81\x84\xe3\x81\xa6\xe3\x81\x84\xe3\x81\xaa\xe3\x81\x84\xe5\xa0\xb4\xe5\x90\x88\xe3\x82\x82\xe5\x90\x8c\xe3\x81\x98\xe3\x81\xa7\xe3\x81\x99\x29",
               "Error", MB_OK|MB_ICONERROR);
        SetStatus("Ready (v2.9.75)"); SetProgress(0); SetBusy(false); return;
    }
    // v2.9.14【監査③】es->info->rate を無検証で代入していた。0 が入ると
    // seg.s / g_projectRate でゼロ除算(AviUtl2ごとクラッシュ)し、JSON にも "fps": 0 が渡って
    // Python 側で全字幕がフレーム0に潰れる。異常値は既定の30に落とす。
    g_projectRate = (sp.rate > 0 && sp.rate <= 1000) ? sp.rate : 30;
    if(sp.rate <= 0 || sp.rate > 1000)
        DebugLog("WARN: invalid project rate " + std::to_string(sp.rate) + " -> using 30");
    if(g_tlClips.empty()){
        MsgBox(g_wnd,
            "\xe5\x8b\x95\xe7\x94\xbb/\xe9\x9f\xb3\xe5\xa3\xb0\xe3\x82\xaf\xe3\x83\xaa\xe3\x83\x83\xe3\x83\x97\xe3\x81\x8c\xe8\xa6\x8b\xe3\x81\xa4\xe3\x81\x8b\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n"
            "\xe3\x82\xbf\xe3\x82\xa4\xe3\x83\xa0\xe3\x83\xa9\xe3\x82\xa4\xe3\x83\xb3\xe3\x81\xab\xe5\x8b\x95\xe7\x94\xbb/\xe9\x9f\xb3\xe5\xa3\xb0\xe3\x82\x92\xe9\x85\x8d\xe7\xbd\xae\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84",
            "Error", MB_OK|MB_ICONERROR);
        SetStatus("Ready (v2.9.75)"); SetProgress(0); SetBusy(false); return;
    }
    // v2.9.20【F2/防御】同一 (ファイル, 区間, 再生位置) のクリップを1つに畳む。
    // ScanCallback は全レイヤーを走査するため、同じ動画をエフェクト用に複数レイヤーへ
    // 複製する編集では同一音声を複数回転写し、字幕オブジェクトが二重に積まれる。
    // 完全一致のみ畳むので誤爆はない(同一音声を2回転写して欲しい場面は存在しない)。
    {
        std::vector<TimelineClip> uniq;
        int dupDropped = 0;
        for(const auto& c : g_tlClips){
            bool dup = false;
            for(const auto& u : uniq){
                if(u.filePath == c.filePath && u.timelineStart == c.timelineStart
                   && u.timelineEnd == c.timelineEnd && u.sourceOffset == c.sourceOffset){ dup = true; break; }
            }
            if(dup){ dupDropped++; continue; }
            uniq.push_back(c);
        }
        if(dupDropped > 0){
            DebugLog("Clips dedup: removed " + std::to_string(dupDropped) + " duplicate clip(s)");
            g_tlClips = uniq;
        }
    }
    DebugLog("Clips: " + std::to_string(g_tlClips.size()) + " Rate: " + std::to_string(g_projectRate));
    // v2.9.29【T2】クリップごとの生値を出す。再生速度の警告が実機で出なかったとき、
    // 「読めていないのか」「読めているが値が違うのか」を切り分けられないと直せない。
    // F1(クリップindexのずれ)や H1 の調査でもこの情報が欲しかったので、診断用ではなく常設する。
    for(size_t ci = 0; ci < g_tlClips.size(); ci++){
        char cb[128];
        sprintf_s(cb, "Clip %d: [%d-%d] offset=%.3f speed=%.2f ",
                  (int)ci, g_tlClips[ci].timelineStart, g_tlClips[ci].timelineEnd,
                  g_tlClips[ci].sourceOffset, g_tlClips[ci].speed);
        DebugLog(std::string(cb) + g_tlClips[ci].filePath);
    }
    // v2.9.6: beam/device/VAD 等はこれまでログに残っておらず、実行間の比較ができなかった
    DebugLog(std::string("========== SETTINGS2: beam=") + std::to_string(beamSize)
        + " temp=" + tBuf2
        + " lang=" + std::to_string(li)
        + " device=" + std::to_string(di)
        + " vad=" + (vadFilter ? "1" : "0")
        + " no_prev_text=" + (noPrevText ? "1" : "0")
        + " word_ts=" + (wordTs ? "1" : "0")
        + " rep_penalty=" + (repPenalty ? "1" : "0")
        + " batched=" + (batched ? "1" : "0")
        + " prompt=" + (promptText.empty() ? std::string("-") : promptText)
        + " hotwords=" + (hotwordsText.empty() ? std::string("-") : hotwordsText)
        + " ==========");
    SetProgress(20);

    if(!RunFasterWhisper(mi, di, bi, beamSize, li, temp, noPrevText, wordTs, repPenalty, promptText, hotwordsText, batched, vadFilter)){
        SetStatus("Ready (v2.9.75)"); SetProgress(0); SetBusy(false); return;
    }

    if(g_segs.empty()){
        // v2.9.52【AI1】従来はここで必ず「音声ファイルを確認してください」と出していた。
        // しかし Python は転写の例外をクリップ単位で握り潰して exit 0 で終わるため、
        // **CUDA のメモリ不足やモデルの破損でもこの分岐に落ちる**。
        // 原因は .err の `Clip N err:` にあるのに一度も見せておらず、
        // 利用者は関係のない音声ファイルを疑わされていた(AG3 と同じ型)。
        // ★足すのは情報だけ。判定は厳しくも緩くもしない(ここは既に失敗した後)。
        std::string m = "\xe6\x96\x87\xe5\xad\x97\xe8\xb5\xb7\xe3\x81\x93\xe3\x81\x97\xe7\xb5\x90\xe6\x9e\x9c\xe3\x81\x8c\xe7\xa9\xba\xe3\x81\xa7\xe3\x81\x99\xe3\x80\x82\n";
        if(!g_clipErr.empty()){
            m += "\x0a\xe8\xbb\xa2\xe5\x86\x99\xe4\xb8\xad\xe3\x81\xab\xe3\x80\x81\xe6\xac\xa1\xe3\x81\xae\xe3\x82\xa8\xe3\x83\xa9\xe3\x83\xbc\xe3\x81\x8c\xe7\x99\xba\xe7\x94\x9f\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x97\xe3\x81\x9f\x3a\x0a\x0a";
            m += g_clipErr;
            m += "\x0a\x0a\xe3\x81\x93\xe3\x82\x8c\xe3\x81\xaf\xe9\x9f\xb3\xe5\xa3\xb0\xe3\x83\x95\xe3\x82\xa1\xe3\x82\xa4\xe3\x83\xab\xe3\x81\xa7\xe3\x81\xaf\xe3\x81\xaa\xe3\x81\x8f\xe3\x80\x81\xe8\xbb\xa2\xe5\x86\x99\xe3\x81\xae\xe5\xae\x9f\xe8\xa1\x8c\xe4\xb8\xad\xe3\x81\xab\xe8\xb5\xb7\xe3\x81\x8d\xe3\x81\x9f\xe5\x95\x8f\xe9\xa1\x8c\xe3\x81\xa7\xe3\x81\x99\xe3\x80\x82\x0a\xe8\xa9\xb3\xe7\xb4\xb0\xe3\x81\xaf\x20\x77\x68\x69\x73\x70\x65\x72\x5f\x64\x65\x62\x75\x67\x2e\x6c\x6f\x67\x20\xe3\x81\xab\xe6\xae\x8b\xe3\x81\xa3\xe3\x81\xa6\xe3\x81\x84\xe3\x81\xbe\xe3\x81\x99\xe3\x80\x82";
        } else {
            m += "\xe9\x9f\xb3\xe5\xa3\xb0\xe3\x83\x95\xe3\x82\xa1\xe3\x82\xa4\xe3\x83\xab\xe3\x82\x92\xe7\xa2\xba\xe8\xaa\x8d\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82";
        }
        MsgBox(g_wnd, m, "\xe7\xb5\x90\xe6\x9e\x9c", MB_OK|MB_ICONWARNING);
        SetStatus("Ready (v2.9.75)"); SetProgress(0); SetBusy(false); return;
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
    // v2.9.51【AH4】ここは **UI から読み直していた**。生成は数分かかるので、その間に
    // 利用者がチェックを変えると **同じ1回の生成で冒頭の値と後半の値が食い違う**。
    // (SetBusy が無効化するのは Generate/Setup ボタンだけで、設定は触れてしまう)
    // 冒頭(3587/3602行)で読んだ mpOn / mgOn をそのまま使う。
    if(mgOnEarly && !mpOn){
            // v2.9.27【N2】従来は maxC のままだったため、結合後は必ず maxC 以下になり
        // SplitText が分割しない = 「最大2行にまとめる」の対象が生まれなかった。
        // 「2行まで許容」は 2 x maxC まで表示できるという意味なので行数ぶん掛ける。
        // 2行 OFF のときは maxLines=1 なので従来と完全に同一。
        int mergeLimit = ((maxC > 0) ? maxC : 20) * maxLines; // 文字数 (v2.8.2で文字数ベースに統一)
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
    // v2.9.27【O1】文節区切り ON でも **fugashi が無ければ Python は何も分割していない**
    // (_morph_group は tagger を作れないと [text] をそのまま返す)。それなのに C++ 側が
    // 「Python がやったはず」と譲っていたため、両方ともグルーピングせず
    // 「最大2行にまとめる」が無反応だった。★2行のチェックは morph ON でもグレーアウト
    // されないので、利用者は設定できてしまう。fugashiMissing は上(mpOn のとき import 実測)
    // で既に確定しているので、無ければ C++ 側でグルーピングする。
    int cLines = (mpOn && !fugashiMissing) ? 1 : maxLines; // v2.8.5 / v2.9.27【O1】
    // v2.9.27【N1】maxchars=0 のとき SplitText は必ず1個しか返さないため、
    // 下のグルーピングが一度も走らず「最大2行にまとめる」が無反応だった
    // (Python の _morph_group は maxc<=0 を 20 文字に読み替えており、経路で挙動が違った)。
    // ★maxchars=0 の「分割しない」を全体に対して変えてはいけない。
    //   2行 OFF の利用者の字幕が突然 20 文字で切られてしまう。
    //   2行が ON のときだけ、分割幅の既定として Python と同じ 20 を使う。
    int splitC = maxC;
    if(splitC <= 0 && cLines > 1) splitC = 20;
    for(auto& seg : g_segs){
        if(seg.e <= seg.s || seg.text.empty()) continue;
        std::vector<std::string> parts = SplitText(seg.text, splitC); // v2.9.27【N1】
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
    // v2.9.51【AH4】冒頭で読み切った値を使う。ここで UI を読み直すと、
    // 生成中(数分)に利用者が値を変えた場合、**同じ1回の生成で設定が混ざる**。
    double lingerSec = lingerSecEarly;
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
    // v2.9.51【AH4】冒頭で読み切った値を使う(生成中に変えられても混ざらない)
    double leadSec = leadSecEarly;
    int leadFrames = (int)(leadSec * g_projectRate);
    if(leadFrames > 0){
        for(size_t i = 0; i < items.size(); i++){
            items[i].s = items[i].s - leadFrames;
            if(items[i].s < 0) items[i].s = 0;
        }
        // v2.9.27【P1】0 でクランプすると、タイムライン先頭付近の複数の item が
        // **同じ 0 に潰れる**。すると下の重なり解消 (items[i].s < items[i-1].e →
        // items[i-1].e = items[i].s) で前の item が長さ 0 になり、
        // 「短すぎ除去」で**黙って消える**。実測: 先頭に3本ある構成で
        // lead=1秒 → 2本、lead=5秒 → 1本まで減った。
        // 開始の順序を回復して、各 item に最低 1 フレーム残す。
        // (lead=0 のときはこのループに入らないので従来と完全に同一)
        for(size_t i = 1; i < items.size(); i++)
            if(items[i].s <= items[i-1].s) items[i].s = items[i-1].s + 1;
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

    // v2.9.75【BG1】**転写の暴走**を検出して知らせる。
    // 実測(2026-08-24): 同じ素材・同じ版でも、たまに whisper が長音の連続に落ちる。
    //   21:15 の回は 25秒ぶんの会話が「あーーー」1件になった(squash=443文字)。
    //   21:02 と 21:21 の回は正常(字幕90件/89件)。**版の違いではない** —
    //   アライメントは転写の後段なので、転写結果を変えることは原理的にない。
    // ★字幕は消さない。**知らせるだけ**。減る方向には一切動かさない
    //   (v2.9.4 で本物の字幕を3件消して撤回した前科がある)。
    // ★この場所に fps が無いので**単位に依存しない判定**にする:
    //   「1文字あたりの表示フレーム数」が全体の中央値の何倍か。
    //   実測の余裕: 正常な回の最大 5.0倍 / 暴走した字幕 42.2倍 → 境目は 10倍。
    int collapsed = 0;
    {
        std::vector<double> fpc;
        for(size_t i = 0; i < finalItems.size(); i++){
            int n = 0;
            for(size_t k = 0; k < finalItems[i].text.size(); k++)
                if((finalItems[i].text[k] & 0xC0) != 0x80) n++;
            if(n > 0) fpc.push_back((double)(finalItems[i].e - finalItems[i].s) / (double)n);
        }
        if(fpc.size() >= 8){
            std::vector<double> sortedFpc = fpc;
            std::sort(sortedFpc.begin(), sortedFpc.end());
            double med = sortedFpc[sortedFpc.size() / 2];
            if(med > 0.0){
                for(size_t i = 0; i < fpc.size(); i++)
                    if(fpc[i] >= med * 10.0) collapsed++;
            }
        }
    }
    DebugLog("Collapse check: " + std::to_string(collapsed) + " over 10x median frames/char");

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
        // v2.9.59【AO1】ここは占有調べなので**止めない**(取れなくても Pass1 が失敗を数える)。
    // ただし黙って進むと「なぜ重なったのか」が後から追えないので記録は残す。
    if(!(g_edit && g_edit->call_edit_section_param(&lc, checkCallback)))
        DebugLog("call_edit_section_param(LayerCheck) failed - occupancy unknown");
        return lc.hasConflict;
    };

    // Shift targetLayer until we find a clear range (max 50 layers)
    int maxShift = 50;
    bool layerClear = false;
    for(int shift = 0; shift < maxShift; shift++){
        if(!findFreeLayer()){ layerClear = true; break; }
        DebugLog("Layer " + std::to_string(targetLayer + 1) + " occupied, shifting...");
        targetLayer++;
    }
    // v2.9.21【H1】50回シフトしても空きが見つからないまま配置へ進んでいた。
    // その状態では create_object が既存オブジェクトと重なって失敗する。記録を残す。
    if(!layerClear)
        DebugLog("WARN: no free layer range after " + std::to_string(maxShift) + " shifts; some subtitles may fail to place");
    DebugLog("Target layer: " + std::to_string(targetLayer + 1) + " (API " + std::to_string(targetLayer) + ")");

    // === PASS 1: Place ALL with create_object (100% reliable) ===
    struct Pass1Param {
        std::vector<PlaceItem>* items;
        std::vector<int>* layers;
        int targetLayer;
        int placed;
        int failed;
        int styleFailed; // v2.9.59【AO2】装飾項目(フォント/サイズ/色/揃え/Y)の設定失敗数
        std::vector<char>* ok; // v2.9.21【H1】item ごとの配置成否
    };
    // v2.9.21【H1】Pass1 は成否をカウントしか持っておらず、Pass2 が「自分が置いた物か」を
    // 判別できなかった。create_object は既存オブジェクトと重なると失敗する(SDK仕様)ため、
    // 失敗した位置には利用者の既存オブジェクトが居る。そこを Pass2 が無条件に delete していた。
    std::vector<char> p1ok(finalItems.size(), 0);
    Pass1Param p1;
    p1.items = &finalItems;
    p1.layers = &itemLayers;
    p1.targetLayer = targetLayer;
    p1.placed = 0;
    p1.failed = 0;
    p1.styleFailed = 0; // v2.9.59【AO2】
    p1.ok = &p1ok;

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
                // v2.9.59【AO2】set_object_item_value も **bool を返す**(plugin2.h 265行)。
                // 従来は6箇所すべて捨てていた。★致命的なのは**テキスト本文**の設定で、
                // ここが失敗すると **中身が空の字幕オブジェクトが置かれ、placed に数えられる**。
                // 項目名は AviUtl2 の日本語名を直接指定しているので、本体側で名前が
                // 変わると**全件が空になるのに「配置 N 件 / 失敗 0 件」と出る**。
                // → 本文だけは成否を見て、失敗したらオブジェクトを消して失敗として数える。
                //   装飾(フォント/サイズ/色/揃え/Y)は当たらなくても既定書式で字幕は出るので、
                //   記録だけ残して止めない。
                if(!es->set_object_item_value(obj, wT, wT, item.text.c_str())){
                    es->delete_object(obj);
                    p->failed++;
                    continue;
                }
                int styleNg = 0;
                if(!es->set_object_item_value(obj, wT, wF, "Yu Gothic UI")) styleNg++;
                if(!es->set_object_item_value(obj, wT, wS, "60.00")) styleNg++;
                if(!es->set_object_item_value(obj, wT, wC, "ffffff")) styleNg++;
                if(!es->set_object_item_value(obj, wT, wA,
                    "\xe4\xb8\xad\xe5\xa4\xae\xe6\x8f\x83\xe3\x81\x88[\xe4\xb8\x8b]")) styleNg++;
                if(!es->set_object_item_value(obj, wD, L"Y", "400.00")) styleNg++;
                if(styleNg > 0) p->styleFailed += styleNg;
                p->placed++;
                (*p->ok)[idx] = 1; // v2.9.21【H1】この item は自分が置いた
            } else {
                p->failed++;
            }
        }
    };

    // v2.9.59【AO1】戻り値を見る。入れなかったら placed=0 のまま「0個配置」と出てしまう。
    if(!(g_edit && g_edit->call_edit_section_param(&p1, pass1Callback))){
        DebugLog("call_edit_section_param(Pass1) failed");
        // v2.9.61【AS1】具体的な条件を書く(上と同じ理由)
        MsgBox(g_wnd, "\xe5\xad\x97\xe5\xb9\x95\xe3\x82\x92\xe9\x85\x8d\xe7\xbd\xae\xe3\x81\xa7\xe3\x81\x8d\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x81\xa7\xe3\x81\x97\xe3\x81\x9f\xe3\x80\x82\x0a\x0a\xe3\x83\x97\xe3\x83\xac\xe3\x83\x93\xe3\x83\xa5\xe3\x83\xbc\xe5\x86\x8d\xe7\x94\x9f\xe4\xb8\xad\xe3\x83\xbb\xe3\x83\x95\xe3\x82\xa1\xe3\x82\xa4\xe3\x83\xab\xe5\x87\xba\xe5\x8a\x9b\xe4\xb8\xad\xe3\x81\xaf\xe9\x85\x8d\xe7\xbd\xae\xe3\x81\xa7\xe3\x81\x8d\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\x0a\xe5\x81\x9c\xe6\xad\xa2\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8b\xe3\x82\x89\xe3\x80\x81\xe3\x82\x82\xe3\x81\x86\xe4\xb8\x80\xe5\xba\xa6\xe5\xae\x9f\xe8\xa1\x8c\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82",
               "Error", MB_OK|MB_ICONERROR);
        SetStatus("Ready (v2.9.75)"); SetProgress(0); SetBusy(false); return;
    }
    int placed = p1.placed;
    DebugLog("Pass1 placed: " + std::to_string(placed) + " failed: " + std::to_string(p1.failed)
             // v2.9.59【AO2】装飾の設定失敗は止めないが、必ず数字を残す。
             // ★ここが 0 でなくなったら AviUtl2 側の項目名が変わった合図。
             + (p1.styleFailed > 0 ? "  styleFailed: " + std::to_string(p1.styleFailed) : ""));

    // === PASS 2: If template, replace each one-by-one (separate edit section) ===
    // v2.9.54【AK1】冒頭で読み切った値を使う。ここで g_templateContent を読み直すと、
    // 生成中に変更/解除された場合に**同じ1回の生成で開始時と配置時が食い違う**。
    if(!tplContentEarly.empty() && placed > 0){
        struct Pass2Param {
            std::vector<PlaceItem>* items;
            std::vector<int>* layers;
            int targetLayer;
            std::string tplContent;
            int replaced;
            int rFailed;
            std::string logPath;
            std::vector<char>* ok; // v2.9.21【H1】Pass1 が置けた item だけを触るため
        };
        Pass2Param p2;
        p2.items = &finalItems;
        p2.layers = &itemLayers;
        p2.targetLayer = targetLayer;
        p2.tplContent = tplContentEarly; // v2.9.54【AK1】冒頭で読み切った値
        p2.replaced = 0;
        p2.rFailed = 0;
        p2.ok = &p1ok;
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
                    // best-effort: Pass2 の経過ログ。失敗しても配置そのものには影響しない
                    std::ofstream lf(p->logPath, std::ios::app);
                    lf << msg << "\n";
                }
            };

            // Step A: Pass1 が置けた item のオブジェクトだけを削除する
            // v2.9.21【H1】以前は全 item について find_object したものを無条件に delete していた。
            // Pass1 が失敗した位置に居るのは利用者の既存オブジェクトなので、それを消していた。
            for(size_t idx = 0; idx < p->items->size(); idx++){
                if(!(*p->ok)[idx]) continue; // 自分が置いていない = 触らない
                auto& item = (*p->items)[idx];
                int apiLayer = p->targetLayer + (*p->layers)[idx];
                OBJECT_HANDLE existing = es->find_object(apiLayer, item.s);
                if(existing){
                    es->delete_object(existing);
                } else {
                    cbLog("PASS2_NOFIND #" + std::to_string(idx) + " L" + std::to_string(apiLayer) + " F" + std::to_string(item.s));
                }
            }

            // Step B: Step A で消した分だけを、テンプレートで作り直す
            // v2.9.21【H1】Pass1 が置けなかった item をここで作ると、利用者の既存オブジェクトと
            // 重なる位置に割り込むことになるので触らない(挙動は従来の「配置できず」と同じ)。
            for(size_t idx = 0; idx < p->items->size(); idx++){
                if(!(*p->ok)[idx]) continue; // Step A と対で揃える
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
                        // v2.9.59【AO2】Pass1 と同じ扱いに揃える(片方だけにしない)。
                        // 本文が入らなければ空のオブジェクトが残るので消す。
                        if(!es->set_object_item_value(obj, wT, wT, item.text.c_str())){
                            es->delete_object(obj);
                            cbLog("PASS2_RESTORE_TEXT_FAIL #" + std::to_string(idx));
                        } else {
                            // v2.9.59【AO2】装飾も Pass1 と同じく**数えて記録**する。
                            // 止めはしない(当たらなくても既定書式で字幕は出る)が、
                            // 0 でなくなったら AviUtl2 側の項目名が変わった合図になる。
                            int sNg = 0;
                            if(!es->set_object_item_value(obj, wT, wF, "Yu Gothic UI")) sNg++;
                            if(!es->set_object_item_value(obj, wT, wS, "60.00")) sNg++;
                            if(!es->set_object_item_value(obj, wT, wC, "ffffff")) sNg++;
                            if(!es->set_object_item_value(obj, wT, wA,
                                "\xe4\xb8\xad\xe5\xa4\xae\xe6\x8f\x83\xe3\x81\x88[\xe4\xb8\x8b]")) sNg++;
                            if(!es->set_object_item_value(obj, wD, L"Y", "400.00")) sNg++;
                            if(sNg > 0) cbLog("PASS2_RESTORE_STYLE_FAIL #" + std::to_string(idx)
                                              + " count=" + std::to_string(sNg));
                        }
                    }
                    p->rFailed++;
                }
            }
        };

        // v2.9.59【AO1】★ここが最悪だった。入れなかった場合、書式テンプレートは一切
        // 適用されないのにログは「Pass2 replaced: 0 failed: 0」= **何も置き換える物が
        // 無かった**ようにしか読めない。利用者には既定書式の字幕がそのまま残り、
        // なぜテンプレが効かないのか分からない。R1(テンプレ選び間違いを黙って受け入れる)
        // と同じ「テンプレが効かないのに成功に見える」型。
        bool p2Ok = (g_edit && g_edit->call_edit_section_param(&p2, pass2Callback));
        DebugLog("Pass2 replaced: " + std::to_string(p2.replaced) + " failed: " + std::to_string(p2.rFailed)
                 + (p2Ok ? "" : "  (call_edit_section_param failed)"));
        if(!p2Ok){
            // ★字幕そのものは Pass1 で置けているので、生成は成功として続ける。
            //   止めずに「書式だけ当たっていない」ことを伝える(K1 の側=詰まらせない)。
            // v2.9.61【AS1】具体的な条件を書く。ここは**止めない側**なので、
            // 「字幕は置けている」と「何を止めれば書式が当たるか」の両方を伝える。
            MsgBox(g_wnd, "\xe6\x9b\xb8\xe5\xbc\x8f\xe3\x83\x86\xe3\x83\xb3\xe3\x83\x97\xe3\x83\xac\xe3\x83\xbc\xe3\x83\x88\xe3\x82\x92\xe9\x81\xa9\xe7\x94\xa8\xe3\x81\xa7\xe3\x81\x8d\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x81\xa7\xe3\x81\x97\xe3\x81\x9f\xe3\x80\x82\x0a\xe5\xad\x97\xe5\xb9\x95\xe3\x81\xaf\xe9\x85\x8d\xe7\xbd\xae\xe3\x81\x95\xe3\x82\x8c\xe3\x81\xa6\xe3\x81\x84\xe3\x81\xbe\xe3\x81\x99\xe3\x81\x8c\xe3\x80\x81\xe6\x97\xa2\xe5\xae\x9a\xe3\x81\xae\xe6\x9b\xb8\xe5\xbc\x8f\xe3\x81\xae\xe3\x81\xbe\xe3\x81\xbe\xe3\x81\xa7\xe3\x81\x99\xe3\x80\x82\x0a\x0a\xe3\x83\x97\xe3\x83\xac\xe3\x83\x93\xe3\x83\xa5\xe3\x83\xbc\xe5\x86\x8d\xe7\x94\x9f\xe4\xb8\xad\xe3\x83\xbb\xe3\x83\x95\xe3\x82\xa1\xe3\x82\xa4\xe3\x83\xab\xe5\x87\xba\xe5\x8a\x9b\xe4\xb8\xad\xe3\x81\xab\xe8\xb5\xb7\xe3\x81\x8d\xe3\x82\x8b\xe3\x81\x93\xe3\x81\xa8\xe3\x81\x8c\xe3\x81\x82\xe3\x82\x8a\xe3\x81\xbe\xe3\x81\x99\xe3\x80\x82\x0a\xe5\x81\x9c\xe6\xad\xa2\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8b\xe3\x82\x89\xe3\x80\x81\xe3\x82\x82\xe3\x81\x86\xe4\xb8\x80\xe5\xba\xa6\xe7\x94\x9f\xe6\x88\x90\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82",
                   "\xe6\x9b\xb8\xe5\xbc\x8f", MB_OK|MB_ICONWARNING);
        }
    }
    DebugLog("Placed: " + std::to_string(placed) + " failed: " + std::to_string(p1.failed));

    SetProgress(100);
    DWORD elapsed = (GetTickCount() - startTick) / 1000;
    char timeBuf[64];
    if(elapsed >= 60) sprintf_s(timeBuf, " %dm%02ds", (int)(elapsed/60), (int)(elapsed%60));
    else sprintf_s(timeBuf, " %ds", (int)elapsed);
    std::string warn; // v2.8.5
    if(fugashiMissing) warn = "\n\xe2\x9a\xa0 fugashi\xe6\x9c\xaa\xe5\xb0\x8e\xe5\x85\xa5\xef\xbc\x9a\xe6\x96\x87\xe7\xaf\x80\xe5\x8c\xba\xe5\x88\x87\xe3\x82\x8a\xe3\x81\xaf\xe6\x9c\xaa\xe9\x81\xa9\xe7\x94\xa8"; // v2.8.9: F「プロ分割」→「文節区切り」用語統一
    // v2.9.21【H1】配置失敗は従来 debug ログにしか出ず、利用者は字幕が減ったことに気づけなかった
    // v2.9.28【S1】再生速度は切り出しにも配置にも使っていないため、100%でないクリップでは
    // 音声だけが伸び縮みして字幕は動かず、クリップの後ろほどズレる(実測: 速度70で
    // 100%と同じ位置に出た)。正しい対応は別途。まず黙って壊れるのを止める。
    // 0 は「速度指定なし」とみなして警告しない(実データ496件中2件)。
    {
        int spdOdd = 0;
        for(size_t si = 0; si < g_tlClips.size(); si++){
            double sv = g_tlClips[si].speed;
            if(sv != 0.0 && (sv < 99.99 || sv > 100.01)) spdOdd++;
        }
        if(spdOdd > 0){
            warn += "\n\xe2\x9a\xa0 " + std::to_string(spdOdd) + "\xe5\x80\x8b\xe3\x81\xaf\xe5\x86\x8d\xe7\x94\x9f\xe9\x80\x9f\xe5\xba\xa6\xe3\x81\x8c\x31\x30\x30\x25\xe4\xbb\xa5\xe5\xa4\x96\x28\xe5\xad\x97\xe5\xb9\x95\xe4\xbd\x8d\xe7\xbd\xae\xe3\x81\x8c\xe3\x81\x9a\xe3\x82\x8c\xe3\x81\xbe\xe3\x81\x99\x29";
            DebugLog("WARN: " + std::to_string(spdOdd) + " clip(s) with playback speed != 100%");
        }
    }
    // v2.9.36【Z2】旧版で導入した人は cuBLAS が無いまま CPU で動いている。
    // M1 ガードのおかげで落ちないぶん、遅い理由が分からない。直し方まで書く。
    // ★生成はブロックしない(CheckMissingForGenerate に入れると K1 と同じく止まる)。
    if(g_cublasMissing)
        warn += "\n\xe2\x9a\xa0 \x43\x55\x44\x41\xe6\x9c\xaa\xe5\xb0\x8e\xe5\x85\xa5\xe3\x81\xa7\x43\x50\x55\xe5\xae\x9f\xe8\xa1\x8c\xe3\x80\x82\xe7\x92\xb0\xe5\xa2\x83\xe3\x82\xbf\xe3\x83\x96\xe3\x81\xa7\xe5\x86\x8d\xe3\x82\xbb\xe3\x83\x83\xe3\x83\x88\xe3\x82\xa2\xe3\x83\x83\xe3\x83\x97";
    if(p1.failed > 0)
        warn += "\n\xe2\x9a\xa0 " + std::to_string(p1.failed) + "\xe5\x80\x8b\xe3\x81\xaf\xe6\x97\xa2\xe5\xad\x98\xe3\x82\xaa\xe3\x83\x96\xe3\x82\xb8\xe3\x82\xa7\xe3\x82\xaf\xe3\x83\x88\xe3\x81\xa8\xe9\x87\x8d\xe3\x81\xaa\xe3\x82\x8a\xe9\x85\x8d\xe7\xbd\xae\xe3\x81\xa7\xe3\x81\x8d\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93";
    // v2.9.38【AB1】言語=自動 のとき、効果音や短いクリップでは判定の手がかりが無く
    // 別言語に転ぶ。幻聴辞書は日本語と英語しか持たないので、韓国語・ロシア語に化けた
    // 幻聴は**素通りする**(実測 2026-08-03: 効果音3つが韓/露の幻聴になり 0 filtered。
    // 同じ素材を ja 固定で回すと「ご視聴ありがとうございました」として3件とも落ちた)。
    // ★どの言語に転ぶか事前に分からないので、辞書を伸ばしても先回りできない。
    //   直し方は言語固定の一本なので、それだけを伝える。
    // ★条件で絞らず、自動を選んでいるときは常に出す。結果から誤判定を推定しようとすると
    //   英語素材などで誤爆する。自動を選ぶのは意図的な行為なので毎回出ても筋が通る。
    if(li == 0)
        warn += "\n\xe2\x9a\xa0\x20\xe8\xa8\x80\xe8\xaa\x9e\x3d\xe8\x87\xaa\xe5\x8b\x95\xe3\x80\x82\xe6\x83\xb3\xe5\xae\x9a\xe5\xa4\x96\xe3\x81\xae\xe8\xa8\x80\xe8\xaa\x9e\xe3\x81\x8c\xe5\x87\xba\xe3\x81\x9f\xe3\x82\x89\x20\x6a\x61\x20\xe5\x9b\xba\xe5\xae\x9a";
    // v2.9.37【AA1】旧版が plugin 直下に置いた HuggingFace キャッシュの残骸。
    // 現行版は models\ しか読まないので、直下のものは二重に置かれたまま死蔵される
    // (実測: turbo が直下と models\ に snapshot ハッシュまで同一で各 1,546.5MB)。
    // ★知らせるだけで消さない。状態欄に完全パスは入らないので debug ログへ出す。
    {
        int stray = CountStrayModelDirs();
        if(stray > 0){
            warn += "\n\xe2\x9a\xa0\x20\xe6\x97\xa7\xe7\x89\x88\xe3\x83\xa2\xe3\x83\x87\xe3\x83\xab\xe3\x81\xae\xe6\xae\x8b\xe9\xaa\xb8\xe3\x81\x82\xe3\x82\x8a\x28\x6d\x6f\x64\x65\x6c\x73\x2d\x2d\x2a\x20\xe3\x81\xaf\xe5\x89\x8a\xe9\x99\xa4\xe5\x8f\xaf\x29";
            DebugLog("WARN: " + std::to_string(stray) + " stray model dir(s) directly under " + GetPluginDir()
                     + " (never read; the real location is " + GetModelsDir() + ") - safe to delete");
        }
    }
    // v2.9.75【BG1】暴走した回だと**気づけない**のが一番の害。ここで知らせる。
    if(collapsed > 0)
        warn += "\n\xe2\x9a\xa0\x20" + std::to_string(collapsed) + "\xe5\x80\x8b\xe3\x81\x8c\xe7\x95\xb0\xe5\xb8\xb8\xe3\x81\xab\xe9\x95\xb7\xe3\x81\x84\x28\xe8\xbb\xa2\xe5\x86\x99\xe3\x81\xae\xe6\x9a\xb4\xe8\xb5\xb0\x29\xe3\x80\x82\xe5\x86\x8d\xe5\xae\x9f\xe8\xa1\x8c\xe3\x82\x92\xe6\x8e\xa8\xe5\xa5\xa8";
    SetStatus("Done! " + std::to_string(placed) + "\xe5\x80\x8b\xe3\x81\xae\xe5\xad\x97\xe5\xb9\x95\xe3\x82\x92\xe9\x85\x8d\xe7\xbd\xae (" + std::to_string(layerEnds.size()) + "Layer)" + timeBuf + warn);
    DebugLog("Placed: " + std::to_string(placed));
    WriteTestLog(); // v2.9.6: 設定+結果+SRT をデスクトップへ自動保存
    SetBusy(false);
}

// =========================================================================
// SRT Export
// =========================================================================

// v2.9.6: SRT本文の組み立て。手動エクスポートと自動テストログで同じ処理を使う。
// 以前は ExportSRT の中にべた書きされていたため、片方だけ直すと食い違う恐れがあった。
static std::string BuildSrtText(){
    char lingerBuf[16] = {}; GetWindowTextA(g_lingerEdit, lingerBuf, sizeof(lingerBuf));
    double lingerSec = atof(lingerBuf);
    if(lingerSec < 0) lingerSec = 0;
    if(lingerSec > 10) lingerSec = 10;
    int lingerFrames = (int)(lingerSec * g_projectRate);

    struct SrtSeg { int s, e; std::string text; };
    std::vector<SrtSeg> srtSegs;
    for(auto& seg : g_segs) srtSegs.push_back({seg.s, seg.e, seg.text});
    if(lingerFrames > 0){
        // v2.9.56【AM1】配置側は延長後に**タイムライン終端でクランプ**しているのに、
        // SRT 側だけそれが無かった。そのため linger を使うと、
        // タイムライン上の最後の字幕は動画の終わりで止まるのに
        // **SRT では動画の終端を超えて続く**。同じ1回の生成の2つの出力が食い違う。
        // ★V2(先頭の潰れ)・V3(長さ0の除去)と**同じ対の3件目**。
        //   配置側(items)と SRT 側(srtSegs)は独自に同じ計算を持っているので、
        //   片方だけ直すと必ずこうなる。
        // ★終端の求め方も配置側と同じにする(g_tlClips の最大 timelineEnd、0 なら掛けない)。
        int timelineEnd = 0;
        for(auto& cl : g_tlClips) if(cl.timelineEnd > timelineEnd) timelineEnd = cl.timelineEnd;
        for(auto& s2 : srtSegs){
            s2.e += lingerFrames;
            if(timelineEnd > 0 && s2.e > timelineEnd) s2.e = timelineEnd;
        }
    }
    char leadBuf2b[16] = {}; GetWindowTextA(g_leadEdit, leadBuf2b, sizeof(leadBuf2b));
    double leadSec2 = atof(leadBuf2b);
    if(leadSec2 < 0) leadSec2 = 0;
    if(leadSec2 > 5) leadSec2 = 5;
    int leadFrames2 = (int)(leadSec2 * g_projectRate);
    if(leadFrames2 > 0){
        for(auto& s2 : srtSegs){
            s2.s -= leadFrames2;
            if(s2.s < 0) s2.s = 0;
        }
        // v2.9.31【V2】P1 と同じ問題がここにもあった。0 でクランプすると先頭付近の複数が
        // 同じ 0 に潰れ、下の重なり解消で長さ 0 になる。配置側(items)は v2.9.27 で直したが
        // **SRT側は独自に同じ計算を持っており取り残されていた**。
        // しかも SRT 側には「短すぎ除去」が無いので、消える代わりに
        // 00:00:00,000 --> 00:00:00,000 という長さ0のエントリができる。
        for(size_t i2 = 1; i2 < srtSegs.size(); i2++)
            if(srtSegs[i2].s <= srtSegs[i2-1].s) srtSegs[i2].s = srtSegs[i2-1].s + 1;
    }
    for(size_t i = 1; i < srtSegs.size(); i++){
        if(srtSegs[i].s < srtSegs[i-1].e)
            srtSegs[i-1].e = srtSegs[i].s;
    }
    // v2.9.31【V3】長さが 0 以下になったものを落とす。配置側(items)には同じ除去があるが
    // SRT側だけ無かった。重なり解消は**前の要素の e しか更新しない**ため、
    // 最後の要素は s > e のまま残りうる(実測: 同位置の超短尺5本で4件が不正になった)。
    // 不正なタイムコード(終了 <= 開始)は SRT として壊れているので出さない。
    srtSegs.erase(std::remove_if(srtSegs.begin(), srtSegs.end(),
                                 [](const SrtSeg& x){ return x.e <= x.s; }), srtSegs.end());
    std::string out;
    int idx = 1;
    for(auto& seg : srtSegs){
        double ss = (double)seg.s / g_projectRate, se = (double)seg.e / g_projectRate;
        int sh = (int)(ss/3600), sm = (int)(fmod(ss,3600)/60), ssc = (int)fmod(ss,60), sms = (int)(fmod(ss,1)*1000);
        int eh = (int)(se/3600), em = (int)(fmod(se,3600)/60), esc2 = (int)fmod(se,60), ems = (int)(fmod(se,1)*1000);
        char buf[160];
        sprintf_s(buf, "%d\n%02d:%02d:%02d,%03d --> %02d:%02d:%02d,%03d\n",
                  idx++, sh,sm,ssc,sms, eh,em,esc2,ems);
        // v2.9.31【V1】2行モード/文節区切りでは本文に**リテラルの \n (0x5C 0x6E)** が入る。
        // AviUtl2 のエイリアスはこれを改行として解釈するが、SRT ではただの2文字なので
        // `\n` と表示されてしまう。SRT では本物の改行に直す。
        // (0x5C を直接書くのは、生成スクリプト側のエスケープ地獄を避けるため)
        {
            std::string t = seg.text;
            for(size_t p = 0; p + 1 < t.size(); ){
                if((unsigned char)t[p] == 0x5C && t[p+1] == 'n'){ t.replace(p, 2, 1, '\n'); p++; }
                else p++;
            }
            out += buf; out += t; out += "\n\n";
        }
    }
    return out;
}

static void ExportSRT(){
    // v2.9.11【監査①】生成中は g_segs を TranscribeThread が書き換えている最中なので触らない。
    // このボタンはハンドルを保持しておらず EnableWindow で止められないため関数側で弾く。
    if(g_busy.load()){
        MsgBox(g_wnd, "\xe5\xad\x97\xe5\xb9\x95\xe7\x94\x9f\xe6\x88\x90\xe4\xb8\xad\xe3\x81\xa7\xe3\x81\x99\xe3\x80\x82\xe5\xae\x8c\xe4\xba\x86\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8b\xe3\x82\x89\xe5\xae\x9f\xe8\xa1\x8c\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84", "SRT Export", MB_OK|MB_ICONINFORMATION);
        return;
    }
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
    // v2.9.40【AD1】従来は「開けたか」も「書けたか」も見ずに **無条件で「完了」** を出していた。
    // 保存先が書き込み不可(リムーバブル抜去 / 権限なし / パス長超過 / 空き容量なし)だと
    // ファイルは作られないのに成功と表示され、利用者は SRT があるものとして先へ進む。
    // ★U1(完了時の警告が画面に一度も出ていなかった)と同じ型 —「作った」と「届いた」は別。
    std::ofstream f(fn, std::ios::binary);
    if(!f){
        MsgBox(g_wnd, "\x53\x52\x54\x20\xe3\x82\x92\xe4\xbf\x9d\xe5\xad\x98\xe3\x81\xa7\xe3\x81\x8d\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x81\xa7\xe3\x81\x97\xe3\x81\x9f\xe3\x80\x82\n\xe4\xbf\x9d\xe5\xad\x98\xe5\x85\x88\xe3\x81\xae\xe3\x83\x95\xe3\x82\xa9\xe3\x83\xab\xe3\x83\x80\xe3\x81\xa8\xe3\x82\xa2\xe3\x82\xaf\xe3\x82\xbb\xe3\x82\xb9\xe6\xa8\xa9\xe3\x82\x92\xe7\xa2\xba\xe8\xaa\x8d\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82", "SRT", MB_OK|MB_ICONERROR);
        return;
    }
    f << BuildSrtText();
    // ★close() まで済ませてから確かめる。バッファに載っただけで失敗が出るのは flush 時。
    f.close();
    if(!f){
        MsgBox(g_wnd, "\x53\x52\x54\x20\xe3\x81\xae\xe6\x9b\xb8\xe3\x81\x8d\xe8\xbe\xbc\xe3\x81\xbf\xe3\x81\xab\xe5\xa4\xb1\xe6\x95\x97\xe3\x81\x97\xe3\x81\xbe\xe3\x81\x97\xe3\x81\x9f\xe3\x80\x82\n\xe4\xbf\x9d\xe5\xad\x98\xe5\x85\x88\xe3\x81\xae\xe7\xa9\xba\xe3\x81\x8d\xe5\xae\xb9\xe9\x87\x8f\xe3\x82\x92\xe7\xa2\xba\xe8\xaa\x8d\xe3\x81\x97\xe3\x81\xa6\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82", "SRT", MB_OK|MB_ICONERROR);
        return;
    }
    MsgBox(g_wnd, "SRT\xe3\x82\xa8\xe3\x82\xaf\xe3\x82\xb9\xe3\x83\x9d\xe3\x83\xbc\xe3\x83\x88\xe5\xae\x8c\xe4\xba\x86", "SRT", MB_OK|MB_ICONINFORMATION);
}

// v2.9.6: 字幕生成が終わるたびに、設定と結果をデスクトップへ自動保存する。
// 毎回手で SRT エクスポートしなくても実行間の比較ができるようにするためのもの。
//   出力先  : デスクトップ\SRT テストログ\<日時>_<backend>_beam<N>.txt
//   中身    : 設定 / 実行情報(落とした字幕を含む) / SRT本文
// 設定は whisper_debug.log から抜き出す。UIから読み直すと二重管理になるうえ、
// 「実際にその実行で使われた値」とズレる危険があるため。
// 失敗しても字幕生成そのものには影響させない (握りつぶす)。
static void WriteTestLog(){
    if(g_segs.empty()) return;
    // v2.9.8: 出力先は whisper_subtitle\testlog_dir.txt の1行目で指定できる。
    // パスを直書きするとフォルダを移すたびにリビルドが要るため外出しにした。
    // ファイルが無い/空なら従来どおりデスクトップ。
    std::wstring dir;
    {
        std::ifstream cf(Utf8ToWide(GetPluginDir() + "\\testlog_dir.txt").c_str());
        std::string line;
        if(cf && std::getline(cf, line)){
            if(line.size() >= 3 && (unsigned char)line[0] == 0xEF) line = line.substr(3); // BOM
            while(!line.empty() && (line.back() == '\r' || line.back() == '\n'
                                    || line.back() == ' ' || line.back() == '\t')) line.pop_back();
            if(!line.empty()) dir = Utf8ToWide(line);
        }
    }
    // v2.9.9: testlog_dir.txt が無い/空なら何も出力しない (オプトイン)。
    // 以前はデスクトップへ出していたが、配布先の人のデスクトップに毎回ファイルが
    // 増えてしまうため既定を「出力しない」に変更した。開発機だけこのファイルを置く。
    if(dir.empty()) return;
    CreateDirectoryW(dir.c_str(), NULL);

    SYSTEMTIME st; GetLocalTime(&st);
    // v2.9.25: 以前は日時つきのファイル名だったため、字幕生成のたびに新しいファイルが
    // 増え続けていた。削除も世代管理も無いので配布先で際限なく溜まる。固定名1ファイルの
    // 上書きにする。日時/バックエンド/beam/字幕数は下の本文に書いており、
    // ファイル名から失われる情報は無い。
    // ★開発で各回の結果を残したいときは「ログを保存.ps1」で退避すること(次の生成で消える)。
    std::wstring path = dir + L"\\whisper_testlog.txt";

    // 直近の実行内容を whisper_debug.log から拾う
    std::string settings, runinfo;
    {
        std::ifstream lf(Utf8ToWide(GetPluginDir() + "\\whisper_debug.log").c_str());
        std::string line;
        while(std::getline(lf, line)){
            if(!line.empty() && line.back() == '\r') line.pop_back();
            if(line.find("SETTINGS") != std::string::npos){ settings += line; settings += "\n"; }
            else if(line.rfind("Clips: ", 0) == 0
                 || line.rfind("Template: ", 0) == 0
                 || line.rfind("Hallucination dropped", 0) == 0
                 || line.rfind("Quality filtered", 0) == 0
                 || line.rfind("Morph timing", 0) == 0
                 || line.rfind("Done: ", 0) == 0){ runinfo += line; runinfo += "\n"; }
        }
    }

    // v2.9.58【AN1】ここは **UI を読み直していた**。この関数の冒頭コメント自身が
    // 「UIから読み直すと二重管理になるうえ『実際にその実行で使われた値』とズレる」と
    // 書いているのに、backend と beam だけがそれを破っていた。実害は2つ:
    //  ① 生成側は v2.9.14【監査③】で beam に**上限20**を足したが、こちらは上限が無い。
    //     21以上を入れると **生成は 20 で走るのにテストログは入力値をそのまま記録する**。
    //     テストログは「実行間の比較」のためのものなので、比較の前提が壊れる。
    //  ② 生成中に UI を変えられる(SetBusy が止めるのは Generate/Setup ボタンだけ)。
    //     AH4/AK1 と同じ形で「その回に使っていない値」を記録しうる。
    // → 上で集めた settings(= whisper_debug.log の SETTINGS 行)から拾う。
    //   これがこの関数の既定の方式で、runinfo も同じところから取っている。
    auto pickInt = [&](const char* key, int fallback){
        size_t p = settings.find(key);
        if(p == std::string::npos) return fallback;   // 記録が無い古いログ向けの保険
        return atoi(settings.c_str() + p + strlen(key));
    };
    int bi   = pickInt("backend=", 0);
    int beam = pickInt("beam=", 5);

    // best-effort: テストログ。失敗しても字幕生成そのものには影響させない(既定 OFF の診断用)
    std::ofstream f(path.c_str(), std::ios::binary);
    if(!f) return;
    char dbuf[64];
    sprintf_s(dbuf, "%04d-%02d-%02d %02d:%02d:%02d",
              st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    f << "===== Whisper Subtitle v2.9.75 test log =====\n";
    f << "datetime  : " << dbuf << "\n";
    f << "backend   : " << (bi == 0 ? "faster-whisper" : "openai-whisper") << "\n";
    f << "beam      : " << beam << "\n";
    f << "subtitles : " << g_segs.size() << "\n";
    f << "\n--- settings ---\n" << settings;
    f << "\n--- run info ---\n" << runinfo;
    f << "\n--- srt ---\n" << BuildSrtText();
    f.close();
    DebugLog("Test log: " + WideToUtf8(path));
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
                // v2.9.11【監査①】文節区切りON時、Python側は word_timestamps を強制Trueにする
                // (実測割当に単語時刻が必須なため)。UIは触れるままだったので「外したのに効かない」
                // 状態だった。実態に合わせてONで固定表示し、操作不可にする。
                if(g_chkWordTs){
                    if(on) SendMessageA(g_chkWordTs, BM_SETCHECK, BST_CHECKED, 0);
                    EnableWindow(g_chkWordTs, on ? FALSE : TRUE);
                }
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
                if(!p.empty()){
                    // v2.9.60【AP1】受け付けない理由を受け取って、汎用文に足して出す。
                    std::string tplWhy;
                    if(LoadTemplate(p, &tplWhy)){ g_templatePath = p; UpdateTemplateLabel(); SaveSettings(); }
                    // v2.9.27【R1】黙って無視すると「選べたのに効かない」になる。理由を出す。
                    else MsgBox(g_wnd, std::string(
                        "\xe3\x81\x93\xe3\x81\xae\xe3\x83\x95\xe3\x82\xa1\xe3\x82\xa4\xe3\x83\xab\xe3\x81\xaf\xe6\x9b\xb8\xe5\xbc\x8f\xe3\x83\x86\xe3\x83\xb3\xe3\x83\x97\xe3\x83\xac\xe3\x83\xbc\xe3\x83\x88\xe3\x81\xa8\xe3\x81\x97\xe3\x81\xa6\xe4\xbd\xbf\xe3\x81\x88\xe3\x81\xbe\xe3\x81\x9b\xe3\x82\x93\xe3\x80\x82\n\n"
                        "\xe3\x83\x86\xe3\x82\xad\xe3\x82\xb9\xe3\x83\x88\xe3\x82\xaa\xe3\x83\x96\xe3\x82\xb8\xe3\x82\xa7\xe3\x82\xaf\xe3\x83\x88\xe3\x82\x92\xe5\x8f\xb3\xe3\x82\xaf\xe3\x83\xaa\xe3\x83\x83\xe3\x82\xaf\xe2\x86\x92"
                        "\xe3\x80\x8c\xe3\x82\xa8\xe3\x82\xa4\xe3\x83\xaa\xe3\x82\xa2\xe3\x82\xb9\xe3\x81\xa8\xe3\x81\x97\xe3\x81\xa6\xe4\xbf\x9d\xe5\xad\x98\xe3\x80\x8d"
                        "\xe3\x81\xa7\xe4\xbd\x9c\xe3\x81\xa3\xe3\x81\x9f .object \xe3\x82\x92\xe9\x81\xb8\xe3\x82\x93\xe3\x81\xa7\xe3\x81\x8f\xe3\x81\xa0\xe3\x81\x95\xe3\x81\x84\xe3\x80\x82") + tplWhy,
                        "\xe6\x9b\xb8\xe5\xbc\x8f", MB_OK|MB_ICONWARNING);
                }
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
        // v2.9.44【AG2】import はできるが **実際には使えない** 状態をここで知らせる。
        // (numpy と噛み合っていない = セットアップが最後まで終わっていない可能性)
        // ★ここは表示だけ。生成はブロックしない — ブロックすると K1 が再発する
        //   (セットアップは「完了」と言うのに生成できず、入れ直しても抜けられなくなる)。
        else if(g_probe.torch && !g_probe.torchUsable)
            t += L"\x25b2\x8981\x518d\x30bb\x30c3\x30c8\x30a2\x30c3\x30d7" L" (NumPy" L"\x3068\x5408\x3063\x3066\x3044\x306a\x3044)";
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
    host->set_plugin_information(L"Whisper Subtitle v2.9.2");
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
    // v2.9.8: "kotoba-whisper (日英)" はコンボ幅146pxに収まらず表示が切れていたので短縮。
    // faster-whisper 専用である旨は下の「選択中のモデル」ラベルが表示する。
    SendMessageW(g_modelCombo, CB_ADDSTRING, 0, (LPARAM)L"kotoba-whisper");
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
    // v2.9.30【U1】1行(20px)だと「Done! N個の字幕を配置 (1Layer) 12s」でほぼ埋まり、
    // 後ろに足した警告が右端で切り捨てられて**一度も見えていなかった**。3行に広げる。
    // ここは生成タブの最下段で、下に 200px 以上余っているので他の配置に影響しない。
    g_status = CreateWindowExW(0, L"STATIC", L"Ready (v2.9.75)", WS_CHILD|WS_VISIBLE, SC(14), SC(y), SC(W-24), SC(54), g_wnd, 0, g_hInst, 0); g_tabSubCtrls.push_back(g_status);

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
    y += 22;
    // v2.9.5: VAD(無音カット)。1行目は x=14..338 で埋まっているので2行目に置く。
    // 既定OFF: 実測でONだと単語が40個消え、開始時刻が一貫して早い側にズレた(vad_OFFはopenai-whisperと完全一致)。
    // v2.9.6: 幅110では「VAD無音カッ」で切れた(BS_AUTOCHECKBOXはチェック枠が幅を食うため
    // 文字分だけでは足りない)。2行目はx=14..338が丸々空いているので150まで広げる。
    g_chkVad = CreateWindowExW(0, L"BUTTON", L"VAD\x7121\x97f3\x30ab\x30c3\x30c8", WS_CHILD|BS_AUTOCHECKBOX, SC(14), SC(y), SC(150), SC(18), g_wnd, 0, g_hInst, 0); g_tabAccCtrls.push_back(g_chkVad);
    // v2.9.2: Batched を1行目(x=268 幅70)から2行目へ移動。幅70では「Batched」が切れていた(開発者報告)。
    // BS_AUTOCHECKBOX はチェック枠が幅を食うため、文字幅ぶんだけでは足りない。
    // 1行目は x=14..338 が埋まっていて広げられないので、余裕のある2行目に置く。
    g_chkBatched = CreateWindowExW(0, L"BUTTON", L"Batched", WS_CHILD|BS_AUTOCHECKBOX, SC(174), SC(y), SC(100), SC(18), g_wnd, 0, g_hInst, 0); g_tabAccCtrls.push_back(g_chkBatched);
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
    // v2.9.49【AH2】チェックボックスを復活。**y座標は変えない**(x=14 に置くだけなのでレイアウトは崩れない)。
    g_chkModel = CreateWindowExW(0, L"BUTTON", L"", WS_CHILD|BS_AUTOCHECKBOX, SC(14), SC(y+2), SC(16), SC(16), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_chkModel);
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
    g_statusSetup = CreateWindowExW(0, L"STATIC", L"Ready (v2.9.75)", WS_CHILD, SC(14), SC(y), SC(W-24), SC(20), g_wnd, 0, g_hInst, 0); g_tabSetupCtrls.push_back(g_statusSetup);
    // y+20 = 350 = 最下端(タブ下端362に収まる)

    SwitchTab(0);
    host->register_window_client(L"Whisper Subtitle", g_wnd);
    g_edit = host->create_edit_handle();
    LoadSettings();
    std::thread(ProbeThread).detach(); // v2.9.0【A】: 起動時に1回だけpythonを回して導入状況を一括実測(表示専用、UIは固まらない)
    EnableWindow(g_chkMergeSeg, (SendMessageA(g_chkMorphSplit, BM_GETCHECK, 0, 0) == BST_CHECKED) ? FALSE : TRUE); // v2.8.5
    // v2.9.11【監査①】起動時にも同じ状態にする(設定読み込み直後)
    {
        bool mOn = (SendMessageA(g_chkMorphSplit, BM_GETCHECK, 0, 0) == BST_CHECKED);
        if(g_chkWordTs){
            if(mOn) SendMessageA(g_chkWordTs, BM_SETCHECK, BST_CHECKED, 0);
            EnableWindow(g_chkWordTs, mOn ? FALSE : TRUE);
        }
    }
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
project(whisper_subtitle_v2_9_75 LANGUAGES CXX)
set(CMAKE_CXX_STANDARD 17)
add_library(whisper_subtitle_v2_9_75 SHARED src/whisper_subtitle.cpp)
target_compile_definitions(whisper_subtitle_v2_9_75 PRIVATE UNICODE _UNICODE)
if(MSVC)
    target_compile_options(whisper_subtitle_v2_9_75 PRIVATE /utf-8 /wd4828)
endif()
# v2.9.6【配布】CRT を静的リンク(/MT)にする。
# 従来は CMake 既定の /MD で、配布先に VCRUNTIME140.dll / VCRUNTIME140_1.dll / MSVCP140.dll
# (Visual C++ 再頒布可能パッケージ) が無いと**起動すらしない**状態だった。
# 同プロジェクトの PresetBrowser / BatchEdit / WaveAmp / WaveProbe は全て cl.exe に /MT を
# 明示しており AviUtl2 上で本番稼働中 = このホストで静的CRTが安全なことは実証済み。
# Whisper だけ /MD だったのは CMake 既定のまま取り残されていただけで、設計判断ではない。
# CMP0091 は cmake_minimum_required(3.15) により NEW になるのでこのプロパティが効く。
set_property(TARGET whisper_subtitle_v2_9_75 PROPERTY MSVC_RUNTIME_LIBRARY "MultiThreaded")
set_target_properties(whisper_subtitle_v2_9_75 PROPERTIES
    OUTPUT_NAME "whisper_subtitle_v2_9_75"
    SUFFIX ".aux2"
    PREFIX ""
    RUNTIME_OUTPUT_DIRECTORY_RELEASE "`${CMAKE_BINARY_DIR}/Release"
)
target_link_libraries(whisper_subtitle_v2_9_75 PRIVATE comctl32 shell32 comdlg32 ole32)
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
    $aux2 = Get-ChildItem -Path $buildDir -Recurse -Filter "whisper_subtitle_v2_9_75.aux2" | Select-Object -First 1
    if($aux2){
        # v2.8b: always deploy next to this script (..\sdk whisper開発場所\), NOT to $d/Desktop,
        # and NEVER to the production AviUtl2Portable folder. Filename is version-specific so the
        # existing whisper_subtitle_v2_8_9.aux2 / whisper_subtitle.aux2 (v2.8 stable) are never touched/overwritten.
        $destDir = $PSScriptRoot
        if([string]::IsNullOrEmpty($destDir)){ $destDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
        $dest = Join-Path $destDir "whisper_subtitle_v2_9_75.aux2"
        Copy-Item $aux2.FullName $dest -Force
        Write-Host ""
        Write-Host "============================================" -ForegroundColor Green
        Write-Host " BUILD SUCCESS! v2.9.75" -ForegroundColor Green
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
