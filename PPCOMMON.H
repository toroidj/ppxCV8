/*-----------------------------------------------------------------------------
	Paper Plane xUI	 commom library		header

	UNICODE		定義してあるときは UNICODE 版
	DEBUG		定義してあるときはデバッグコード
-----------------------------------------------------------------------------*/
#ifndef NODLL
#define NODLL 0
#define DLLDEFINED
#define DLLPARAM(value) = value
#else
#define DLLDEFINED extern
#define DLLPARAM(value)
#endif
#ifdef ONVFSDLL
#define ONPPXDLL
#endif
#if !NODLL && !defined(WINEGCC)
	#define EXTDLL __declspec(dllexport)
	#ifdef ONPPXDLL
		#define PPXDLL __declspec(dllexport)
	#else
		#define PPXDLL __declspec(dllimport)
	#endif
#else
	#define EXTDLL
	#define PPXDLL
#endif
#define PPXAPI WINAPI

#ifdef __cplusplus
extern "C" {
#endif

#define VFPS 600	// VFS パスのサイズ
#define OVER_VFPS_MSG "<too long>"	// VFPS を越えたときに代わりに得られる内容
#define VFSGETFFINFO_TYPESIZE 10


#ifdef RELEASE
#define PPXSYNCDEBUG 0
#else
#define PPXSYNCDEBUG 0 // UsePPx での同期待ちダイアログで問題箇所を表示する
#endif

/*-----------------------------------------------------------------------------
	PPx 本体, PPCOMMON, PPx Module 間の情報交換用構造体

Function の cmdID 割当て：記載がない限り、戻り値が 0 なら未実装

	置き換え	'0'→DLLのpath	'1'→カレントディレクトリ
				'F', 'X', 'T', 'R', 'Y', 't', 'L' → '1'と'C'から生成
*/
struct tagPPXAPPINFOA;
struct tagPPXAPPINFOW;

#define PPxInfoFunc(info, func, uptr) (info)->Function(info, func, (void *)uptr)
#define PPxInfoFunc32i(info, func, uptr) (int)(info)->Function(info, func, (void *)uptr)
#define PPxInfoFunc32u(info, func, uptr) (DWORD)(info)->Function(info, func, (void *)uptr)
#define PPxEnumInfoFunc(info, func, str, work) ((work)->buffer = str, PPxInfoFunc(info, func, work))

/* 0x0000～PPXCMDID_FILL:	PPXCMDENUMSTRUCT を使用したやりとり用。
		cmdid :		列挙対応版	'C', 'X', 'T'
					非対応		'0', '1', '2', 'R', 'Y', 't', 'L'

		uptr :		PPXCMDENUMSTRUCT へのポインタ
					未実装の場合でも、uptr->enums.buffer = '\0' の処理が必要
					また、パラメータを有するマクロ文字の場合は、
					そのパラメータへのポインタを buffer に格納する。

	0-100:マクロ文字の展開用に用いられる
*/
//##  0x100-1ff : 非列挙版非マクロ文字
//##  0x200-3ff : 列挙処理、非マクロ文字
//#define PPXCMDID_ENUMCOUNT	0x0200	// マクロ文字による列挙処理/対象数
						// 戻り値が 0:未実装、-1:対象無し 正:対象数
#define PPXCMDID_STARTENUM		0x0201	// マクロ文字による列挙処理/開始
						// 戻り値が偽:列挙対象無し(PPcのマーク無しは該当しない)
#define PPXCMDID_NEXTENUM		0x0202	// マクロ文字による列挙処理/次へ
						// 戻り値が偽:列挙対象無し
#define PPXCMDID_ENDENUM		0x0203	// マクロ文字による列挙処理/終了

#define PPXCMDID_STARTNOENUM	0x0204	// マーク無しはなしの列挙処理/開始
#define PPXCMDID_TRIMENUM		0x0205	// 処理済み(先頭)のマークを解除
#define PPXCMDID_ENUMFINDDATA	0x0300	// 該当エントリの WIN32_FIND_DATA
#define PPXCMDID_ENUMATTR		0x0301	// 該当エントリの属性(GetFileAttributes)
//#define PPXCMDID_ENUMSIZE		0x0302	// 該当エントリのファイルサイズ
//#define PPXCMDID_ENUMCOMMENT	0x0303	// 該当エントリのコメント

#define PPXCMDID_FILL			0x03ff	// 必ず戻り値が必要なコマンドの上限値

//##  0x0400～0x07ff:			各種コマンド(PPCOMMON内で解釈/共用)。

#define PPXCMDID_COMMAND		0x0400	// 未定義 *command
#define PPXCMDID_EXECUTE		0x0401	// 指定文字列を実行
#define PPXCMDID_EXTRACT		0x0402	// 指定文字列を展開して再格納
#define PPXCMDID_FUNCTION		0x0403	// 未定義 %*function
#define PPXCMDID_GETKEYNAME		0x0404	// key が示すキー名称をstrに出力
#define PPXCMDID_LONG_EXTRACT	0x0405	// PPXCMDID_EXTRACT のover CMDLINESIZE
#define PPXCMDID_LONG_EXTRACT_E	0x0406	// PPXCMDID_LONG_EXTRACT のエラー出力版

#define PPXCMDID_MENUONMENU		0x0700	// メニュー上のメニューのトリガ
#define PPXCMDID_ADDEXMENU		0x0701	// ポップアップ用追加メニュー(?xxx)生成
typedef struct {
	HMENU hMenu; // 追加元メニュー
	const TCHAR *exname; // 拡張名
	struct tagThSTRUCT *TH;
	DWORD *index;
} ADDEXMENUINFO;

#define PPXCMDID_EXECEXMENU		0x0702	// ポップアップ用追加メニュー
typedef struct {
	HMENU hMenu; // 追加元メニュー
	int index;
} EXECEXMENUINFO;

#define PPXCMDID_GETTMPFILENAME	0x07ef	// PP_ExtractMacro で %@ を使ったとき、その名前の通知時に使用

#define PPXCMDID_CHARTYPE		0x07fe	// 使用文字コード 0:Multibyte 1:UNICODE
#define PPXCMDID_VERSION		0x07ff	// PPx のバージョンを返す(10進 12300)

//##  0x0800～0x0bff:			各種コマンド(PPxで解釈)。
//   0x0800～0x09ff 取得系
// 窓情報 0x 2x 8x 9x ax
#define PPXCMDID_GETSUBID		0x0807	// 補助IDを取得
#define PPXCMDID_GETREQHWND		0x0808	// 呼び出し元HWND(PPvの依頼元)
#define PPXCMDID_GETFILEINFO	0x0820	// ファイル情報(VFSGetFileType)
#define PPXCMDID_WINDOWDIR		0x0881	// 連結状態における現在窓の位置

// 880, 888, 889 使用済み
//#define PPXCMDID_PAIRWINDOWC	0x0882	// 反対窓の ID
#define PPXCMDID_PAIRWINDOW		0x0883	// 反対窓の HWND(PPc)

#define PPXCMDID_COMBOIDNAME	0x0884	// 一体化窓の親窓ID
#define PPXCMDID_COMBOWINDOW	0x0885	// 一体化窓の親窓HWND(PPc)
#define PPXCMDID_COMBOSHOWPANES	0x0886	// 一体化窓の表示ペイン数, 非一体化:0
#define PPXCMDID_COMBOITEMS		0x0887	// 一体化窓に登録された窓数, 非一体化:0
#define PPXCMDID_COMBOIDCOUNT	0x088a	// 一体化窓の登録されたIDの窓数
#define PPXCMDID_COMBOGETPANE	0x088b	// 指定IDタブのペインindexを検索
#define PPXCMDID_COMBOGETHWNDEX	0x088c	// 指定IDタブのHWNDを取得

#define PPXCMDID_POPUPPOS		0x089f	// メニュー表示等に使う表示座標(POINT)
#define PPXCMDID_CSRDISPW		0x08a0	// 幅方向の表示可能数(PPc, PPv)
#define PPXCMDID_CSRDISPH		0x08a2	// 高さ方向の表示可能数(PPc, PPv)
#define PPXCMDID_CSRDINDEX		0x08a3	// 表示開始インデックス(PPc)
// 現在窓の情報
// PPc 4x 5x 6x
#define PPXCMDID_DIRTOTAL		0x0840	// 全エントリ数(ファイルマスクなし状態)
#define PPXCMDID_DIRTTOTAL		0x0841	// 利用可能エントリ数(マスクあり)
#define PPXCMDID_DIRTTOTALFILE	0x0842	// 利用可能ファイル数(マスクあり)
#define PPXCMDID_DIRTTOTALDIR	0x0843	// 利用可能ディレクトリ数(マスクあり)
#define PPXCMDID_DIRMARKS		0x0844	// 選択エントリ数
#define PPXCMDID_DIRMARKSIZE	0x0845	// 選択エントリサイズ
#define PPXCMDID_DIRLOADCOUNT	0x0846	// 読み込み/更新カウンタ
#define PPXCMDID_DIRTYPE		0x0848	// 現在のディレクトリの形式

#define PPXCMDID_COMBOWNDTYPE	0x0850	// 一体化窓上の現在・反対・プレビュー
#define PPXCMDID_COMBOSHOWINDEX	0x0851	// 一体化窓上のペイン位置
#define PPXCMDID_COMBOTABPANE	0x0852	// タブのペイン
#define PPXCMDID_COMBOTABNAME	0x0853	// タブの表示名
#define PPXCMDID_COMBOTABIDNAME	0x0854	// 該当タブのID(C_x)
#define PPXCMDID_DIRLOCK		0x0855	// ロック状態
#define PPXCMDID_TABTEXTCOLOR	0x0856	// タブの文字色
#define PPXCMDID_TABBACKCOLOR	0x0857	// タブの背景色
#define PPXCMDID_COMBOTABINDEX	0x0858	// 一体化窓上のタブ位置
#define PPXCMDID_COMBOTABCOUNT	0x0859	// 一体化窓上の該当ペインのタブ数
#define PPXCMDID_COMBOGETTAB	0x085a	// タブ位置取得
#define PPXCMDID_COMBOTABLONGID	0x085b	// 該当タブのID(CZxxx対応)
#define PPXCMDID_COMBOTABHWND	0x085c	// 該当タブのHWNDを取得

#define PPXCMDID_DRIVELABEL		0x085d	// ドライブのボリュームラベル
#define PPXCMDID_DRIVETOTALSIZE	0x085e	// ドライブの総容量
#define PPXCMDID_DRIVEFREE		0x085f	// ドライブの空き容量

#define PPXCMDID_SYNCVIEW		0x0860	// 連動ビュー状態
//#define PPXCMDID_SYNCINFO		0x0861	// 連動情報表示
#define PPXCMDID_SLOWMODE		0x0862	// 低速モード
#define PPXCMDID_POINTINFO		0x0863	// マウスカーソルが示す場所の情報 type,index/X,Y
// 選択情報 8x(他使用済み)
#define PPXCMDID_CSRINDEX		0x0880	// 該当エントリのカーソルindex(PPc)
#define PPXCMDID_CSRX			0x0888	// カーソル X 位置(PPv ※現在未使用)
#define PPXCMDID_CSRY			0x0889	// カーソル Y 位置(PPv ※現在未使用)
#define PPXCMDID_CSRLOCATE		0x088e	// カーソル位置 (x, y, w, h)
		// PPc x,y カーソル位置		w,h
		// PPv x,y カーソル位置		w,h 選択位置?
		// PPe x,y カーソル・選択位置

#define PPXCMDID_CSRRECT		0x088f	// メニュー表示用座標情報 (x, y, w, h)
		//	x,y メニュー表示スクリーン座標(ほぼ PPXCMDID_POPUPPOS)
		//	w,h PPc:エントリの幅、高さ PPv,PPe:基準フォントサイズ ※PPe x=1固定
// 選択物情報
// PPc
#define PPXCMDID_CSRATTR		0x0900	// 該当エントリの属性(GetFileAttributes
#define PPXCMDID_CSRMARK		0x0901	// 該当エントリのマーク
#define PPXCMDID_CSRMSIZE		0x0902	// 該当エントリのファイルサイズ
#define PPXCMDID_CSRCOMMENT		0x0903	// 該当エントリのコメント
#define PPXCMDID_CSRSTATE		0x0907	// 該当エントリの状態
#define PPXCMDID_CSREXTCOLOR	0x0908	// 該当エントリの拡張子色

// PPXCMDID_ENTRY* で文字列を扱うときの構造体
typedef union {
	int index; /* in */
	char path[MAX_PATH]; /* out */
} PPXUPTR_INDEX_UPATHA;
typedef union {
	int index;
	WCHAR path[MAX_PATH];
} PPXUPTR_INDEX_UPATHW;

typedef union {
	int index; /* in */
	char *result; /* out, HeapAlloc */
} PPXUPTR_ENTRYINFOA;
typedef union {
	int index;
	WCHAR *result;
} PPXUPTR_ENTRYINFOW;

#define PPXCMDID_ENTRYATTR		0x0910	// 指定エントリの属性(GetFileAttributes
#define PPXCMDID_ENTRYMARK		0x0911	// 指定エントリのマーク
#define PPXCMDID_ENTRYMSIZE		0x0912	// 指定エントリのファイルサイズ
#define PPXCMDID_ENTRYCOMMENT	0x0913	// 指定エントリのコメント
#define PPXCMDID_ENTRYMTIME		0x0914	// 指定エントリの更新時刻
#define PPXCMDID_ENTRYCTIME		0x0915	// 指定エントリの作成刻
#define PPXCMDID_ENTRYATIME		0x0916	// 指定エントリのアクセス時刻
#define PPXCMDID_ENTRYSTATE		0x0917	// 指定エントリの状態
#define PPXCMDID_ENTRYEXTCOLOR	0x0918	// 指定エントリの拡張子色
#define PPXCMDID_ENTRYNAME		0x0919	// 指定エントリの名前
#define PPXCMDID_ENTRYANAME		0x091a	// 指定エントリの短い名前

#define PPXCMDID_ENTRYINFO		0x091b	// 指定エントリの詳細情報(^[I])
#define PPXCMDID_ENTRY_HS_GETINDEX 0x091c	// PPXCMDID_ENTRYMARKFIRST_HS の
											// index からカーソル位置を得る
#define PPXCMDID_ENTRYEXTDATA_GETID 0x091d	// SETDATA/GETDATAに使うIDを取得
		// 指定エントリに関連づけたユーザデータを読み書きする
#define PPXCMDID_ENTRYEXTDATA_SETDATA 0x091e
#define PPXCMDID_ENTRYEXTDATA_GETDATA 0x091f
	// 拡張コメント用ID ID1=DFC_COMMENTEX_MAX, ID2=DFC_COMMENTEX_MAX-1...
	#define DFC_COMMENTEX_MAX 0x7f00 // PPX.H に他のIDを記載
	#define DFC_COMMENTEX (DFC_COMMENTEX_MAX - 10 + 1)
typedef struct {
	DWORD entry; // 対象となる entry を指定する。PPXCMDID_ENTRY_HS_GETINDEX で得た値を使用する
	DWORD id; // 識別ID。PPXCMDID_ENTRYEXTDATA_GETID で得た値を使用する
	SIZE32_T size; // dataの大きさ。64k - 1 以内の大きさまで対応
	BYTE *data;
} ENTRYEXTDATASTRUCT;
#define PPXCMDID_ENTRYFROMNAME	0x0920	// 指定 result のエントリ index を取得

#define PPXCMDID_GETFGWND		0x09f0	// PPtray, 直前の前面hWnd取得


//   0x0a00～0x0bff 設定、実行系
// コマンド類
#define PPXCMDID_CHDIR			0x0a00	// %j コマンド、path ジャンプ
#define PPXCMDID_PATHJUMP		0x0a01	// %J コマンド、path+file+ext ジャンプ
#define PPXCMDID_ENTRYJUMP		0x0a02	// %J コマンド、file+ext ジャンプ
#define PPXCMDID_MOVECSR		0x0a03	// *cursor コマンド
#define PPXCMDID_PPXCOMMAD		0x0a04	// %K コマンド
	// ※ 戻り値は、0:該当コマンドなし、1:正常実行完了 その他: (ERRORCODE ^ 1)
	// ※ ERRORCODE ^ 1 で OK
#define PPXCMDID_SETPOPLINE		0x0a05	// PopLineMessageを登録
#define PPXCMDID_SETSTATLINE	0x0a06	// PopLineMessage(作業中状況)を登録
#define PPXCMDID_EXTREPORTTEXT	0x0a07	// 報告窓があればテキストをログ表示
  #define PPXCMDID_EXTREPORTTEXT_CLOSE	0x0a071 // 戻値:報告窓は存在しない
  #define PPXCMDID_EXTREPORTTEXT_LOG	0x0a072 // 戻値:報告窓に書き込んだ
#define PPXCMDID_GETWNDVARIABLESTRUCT 0x0a08  // ZStringVariable で使用する窓別の構造体のアドレスを取得
#define PPXCMDID_GETCONTROLVARIABLESTRUCT 0x0a09 // ZStringVariable で使用するEdit/Tree内構造体のアドレスを取得
#define PPXCMDID_REPORTTEXT	0x0a0a	// テキストをログ表示
#define PPXCMDID_MESSAGE	0x0a0b	// メッセージダイアログを表示
#define PPXCMDID_DEBUGLOG	0x0a0c	// ログ記録(XM_DbgLOG)

//#define PPXCMDID_GETEXTRACTVARIABLESTRUCT 0x0a0b // ZStringVariable で使用するExract内構造体のアドレスを取得

#define PPXCMDID_SETPOPUPPOS	0x0a9f	// メニュー表示等に使う表示座標(POINT)

// 以下は num[0]: pane, num[1]:tab
#define PPXCMDID_COMBOTABEXTRACTLONG 0x0a4d	// 一体化窓の特定窓の取得(長文字列)
#define PPXCMDID_COMBOTABEXECUTE 0x0a4e	// 一体化窓の特定窓で実行
#define PPXCMDID_COMBOTABEXTRACT 0x0a4f	// 一体化窓の特定窓の取得
#define PPXCMDID_SETTABNAME		0x0a53	// タブの表示名 o

// 以下は PPXUPTR_TABINDEXSTR を使用
typedef struct {
	int pane, tab;
	char *str;
} PPXUPTR_TABINDEXSTRA;
typedef struct {
	int pane, tab;
	WCHAR *str;
} PPXUPTR_TABINDEXSTRW;

#define PPXCMDID_SETCOMBOWNDTYPE 0x0a50	// 一体化窓上の現在・反対・プレビュー
#define PPXCMDID_SETCOMBOTABPANE 0x0a52	// タブのペイン
#define PPXCMDID_SETDIRLOCK		0x0a55	// ロック状態
#define PPXCMDID_SETTABTEXTCOLOR 0x0a56	// タブの文字色
#define PPXCMDID_SETTABBACKCOLOR 0x0a57	// タブの背景色


// PPc
#define PPXCMDID_SETSYNCVIEW	0x0a60	// 連動ビュー状態
#define PPXCMDID_SETSYNCINFO	0x0a61	// 連動情報表示
#define PPXCMDID_SETSLOWMODE	0x0a62	// 低速モード
// 選択操作
#define PPXCMDID_CSRSETINDEX	0x0a80	// 該当エントリのインデックス(PPc)
#define PPXCMDID_CSRMARKFIRST	0x0a83	// 最初のマークエントリ
#define PPXCMDID_CSRMARKNEXT	0x0a84	// 次のマークエントリ
#define PPXCMDID_CSRMARKPREV	0x0a85	// 前のマークエントリ
#define PPXCMDID_CSRMARKLAST	0x0a86	// 最後のマークエントリ
//#define PPXCMDID_CSRSETX		0x0a88	// 該当エントリの X 位置(PPv)
//#define PPXCMDID_CSRSETY		0x0a89	// 該当エントリの Y 位置(PPv)
#define PPXCMDID_CSRSETLOCATE	0x0a8e	// カーソル位置 (x, y, w, h)

#define PPXCMDID_ENTRYMARKFIRST	0x0a93	// 最初のマークエントリ
#define PPXCMDID_ENTRYMARKNEXT	0x0a94	// 次のマークエントリ
#define PPXCMDID_ENTRYINSERT	0x0a95	// エントリを挿入
#define PPXCMDID_ENTRYINSERTMSG	0x0a96	// メッセージを挿入
#define PPXCMDID_ENTRYHIDEENTRY	0x0a97	// エントリを隠す(除去)
#define PPXCMDID_ENTRYMARKFIRST_HS	0x0a98	// 最初のマークエントリ(高速版)
#define PPXCMDID_ENTRYMARKNEXT_HS	0x0a99	// 次のマークエントリ(高速版)
#define PPXCMDID_ENTRYMARKPREV_HS	0x0a9a	// 前のマークエントリ(高速版)
#define PPXCMDID_ENTRYMARKLAST_HS	0x0a9b	// 最後のマークエントリ(高速版)

#define PPXCMDID_CSRSETDINDEX	0x0aa3	// 表示開始インデックス(PPc)

// 選択物操作
//PPc
#define PPXCMDID_CSRSETMARK		0x0b01	// 該当エントリのマーク
#define PPXCMDID_CSRSETCOMMENT	0x0b03	// 該当エントリのコメント
#define PPXCMDID_CSRSETSTATE	0x0b07	// 該当エントリの状態
#define PPXCMDID_CSRSETEXTCOLOR	0x0b08	// 該当エントリの拡張子色

#define PPXCMDID_ENTRYSETMARK		0x0b11	// 指定エントリのマーク
#define PPXCMDID_ENTRYSETCOMMENT	0x0b13	// 指定エントリのコメント
#define PPXCMDID_ENTRYSETSTATE		0x0b17	// 指定エントリの状態
#define PPXCMDID_ENTRYSETEXTCOLOR	0x0b18	// 指定エントリの拡張子色

// 機能拡張・フック

#define PPXCMDID_CLOSE_MODELESS_DIALOG	0x0bef	// モードレスダイアログの終了通知(hDlgWnd)
//機能差し替え（主にコンソール用との差し替え）
// #define PPXCMDID_HOOK_WAITSTATE	0x0bfa	// 待機通知用
// #define PPXCMDID_HOOK_LINEINPUT	0x0bfb	// tInputEx 差し替え用
#define PPXCMDID_HOOK_POPUPMENU	0x0bfc	// TrackPopupMenu 差し替え用
// #define PPXCMDID_HOOK_MSGBOX	0x0bfd	// PMessageBox 差し替え用
// #define PPXCMDID_HOOK_WAITKEY	0x0bfe	// １文字キー入力
// #define PPXCMDID_TEMP_WINDOW	0x0bff	// フォーカス制御用ウィンドウを取得

//## 0x0c00～0x0fff:	各種コマンド(PPx Moduleとのやりとりに使用する)。
#define PPXCMDID_REPORTPPCCOLUMN	0x0c02 // PPcustにPPcの拡張書式を通知
#define PPXCMDID_REQUIREKEYHOOK		0x0c03 // KeyHookEntryを登録
#define PPXCMDID_REPORTSEARCH		0x0c04 // 検索結果を通知(属性不明)
#define PPXCMDID_REPORTSEARCH_FILE		0x0c05 // 検索結果を通知(ファイル)
#define PPXCMDID_REPORTSEARCH_DIRECTORY	0x0c06 // 検索結果を通知(ディレクトリ)
#define PPXCMDID_REPORTSEARCH_FDATA		0x0c07 // 検索結果を通知(FINDDATA)
#define PPXCMDID_GETCONTROLVARIABLEDATA 0x0c0e  // Edit特殊環境変数の値を取得
#define PPXCMDID_SETCONTROLVARIABLEDATA 0x0c0f  // Edit特殊環境変数の値を設定
#define PPXCMDID_GETWNDVARIABLEDATA 0x0c10  // 窓別特殊環境変数の値を取得
#define PPXCMDID_SETWNDVARIABLEDATA 0x0c11  // 窓別特殊環境変数の値を設定
#define PPXCMDID_LONG_RESULT		0x0c12  // CMDLINESIZE を越える結果の受付
#define PPXCMDID_PPLIBHANDLE		0x0c13  // PPLIBxx.DLLのHMODULEを取得
#define PPXCMDID_PPLIBPATH			0x0c14  // PPLIBxx.DLLのPATHを取得
#define PPXCMDID_GETPROCVARIABLEDATA 0x0c15  // プロセス特殊環境変数の値を取得
#define PPXCMDID_SETPROCVARIABLEDATA 0x0c16  // プロセス特殊環境変数の値を設定
#define PPXCMDID_NEWAPPINFO			0x0c17	// 新しいPPXAPPINFOを要求
	typedef struct {
		int flags, datasize;
		void (PPXAPI *ThreadEntry)(struct tagPPXAPPINFOW *ppxa, void *data);
		void *data;
	} PPXCMDID_NEWAPPINFO_STRUCT;
	#define NEWAPPINFO_SIMPLE	0	// PPXAPPINFOW のメモリ確保のみ(HeapFreeで解放)。ウィンドウの関連付けなし、既存スレッド、親とのやりとり不可(PP_ExtractMacro内機能に限定)
	#define NEWAPPINFO_MESSAGE	B0	// メッセージやりとりウィンドウ付き
	#define NEWAPPINFO_PARENT	B1	// NEWAPPINFO_MESSAGEの親の子ウィンドウ関連づけ
	#define NEWAPPINFO_THREAD	B2	// 新規スレッド
//#define PPXCMDID_FREE_PPLIBINFO	0x0c18  // PPLIBxx.DLLのPPXAPPINFOを解放
#define PPXCMDID_GETRAWPARAM		0x0c19  // PPXCMDID_COMMAND等の未加工パラメータを取得
#define PPXCMDID_REQUIRE_CLOSETHREAD 0x0c1a	// PPXMEVENT_CLOSETHREAD を送信要求

//## 0xff000～0xfffff:	内部用
//-----------------------------------------------------------------------------
#define PPXARESULT(errorcode) ((errorcode) ^ 1)	// errorcode を返す時用
#define PPXA_INVALID_FUNCTION PPXARESULT(ERROR_INVALID_FUNCTION) // 0
#define PPXA_NO_ERROR PPXARESULT(NO_ERROR) // 1

typedef struct {
	char *buffer;	// 展開に必要なパラメータ／展開内容
	INT_PTR enumID;	// enum処理用バッファ
} PPXCMDENUMSTRUCTA;

typedef struct {
	WCHAR *buffer;	// 展開に必要なパラメータ／展開内容
	INT_PTR enumID;	// enum処理用バッファ
} PPXCMDENUMSTRUCTW;

typedef struct {
	const TCHAR *param; // コマンド名
	TCHAR *dest; // 結果保存先(Max CMDLINESIZE)
	const TCHAR *optparam; // パラメータ
} PPXMDLFUNCSTRUCT;

// モジュール共用体
typedef union tagPPXMODULEPARAM {
	struct tagPPXMINFOSTRUCT *info;
	struct tagPPXMCOMMANDSTRUCT *command;
	struct tagPPXMFILEDRAWSTRUCT *draw;
	struct tagPPXMKEYHOOKSTRUCT *keyhook;
	struct tagPPXMSEARCHSTRUCT *search;
} PPXMODULEPARAM;

typedef union {
	char str[1];
	WCHAR strW[1];
	WORD key;
	DWORD num;
	DWORD nums[1];
	int inums[1];
	INTHL numhl;
	UINTHL unumhl;
	DWORD_PTR dptrs[1];
	FILETIME ftime;
	COLORREF color;
	PPXMDLFUNCSTRUCT funcparam; // PPx独自のfunctionを処理するときに使用
	PPXCMDENUMSTRUCTA enums; // 列挙処理用
	PPXUPTR_INDEX_UPATHA indexpath;
	PPXUPTR_ENTRYINFOA entryinfo;
	PPXUPTR_TABINDEXSTRA tabidxstr;
} PPXAPPINFOUNIONA;

typedef union {
	WCHAR str[1];
	WCHAR strW[1];
	WORD key;
	DWORD num;
	DWORD nums[1];
	int inums[1];
	INTHL numhl;
	UINTHL unumhl;
	DWORD_PTR dptrs[1];
	FILETIME ftime;
	COLORREF color;
	PPXMDLFUNCSTRUCT funcparam;
	PPXCMDENUMSTRUCTW enums;
	PPXUPTR_INDEX_UPATHW indexpath;
	PPXUPTR_ENTRYINFOW entryinfo;
	PPXUPTR_TABINDEXSTRW tabidxstr;
} PPXAPPINFOUNIONW;

typedef DWORD_PTR(USECDECL * PPXAPPINFOFUNCTIONA)(struct tagPPXAPPINFOA *ppxa, DWORD cmdID, void *uptr);
typedef DWORD_PTR(USECDECL * PPXAPPINFOFUNCTIONW)(struct tagPPXAPPINFOW *ppxa, DWORD cmdID, void *uptr);

typedef struct tagPPXAPPINFOA {
	PPXAPPINFOFUNCTIONA Function;
	const char *Name;		// "PPc", "PPv", "PPTray" 等
	const char *RegID;	// "C_A" 等
	HWND hWnd;		// この構造体を所有する窓
} PPXAPPINFOA;

typedef struct tagPPXAPPINFOW {
	PPXAPPINFOFUNCTIONW Function;
	const WCHAR *Name;	// "PPc", "PPv", "PPTray" 等
	const WCHAR *RegID;	// "C_A" 等
	HWND hWnd;		// この構造体を所有する窓
} PPXAPPINFOW;

#ifdef UNICODE
	typedef PPXAPPINFOW			PPXAPPINFO;
	typedef PPXCMDENUMSTRUCTW	PPXCMDENUMSTRUCT;
	typedef PPXAPPINFOUNIONW	PPXAPPINFOUNION;
	typedef PPXAPPINFOFUNCTIONW	PPXAPPINFOFUNCTION;
	typedef PPXUPTR_INDEX_UPATHW PPXUPTR_INDEX_UPATH;
	typedef PPXUPTR_TABINDEXSTRW PPXUPTR_TABINDEXSTR;
	typedef PPXUPTR_ENTRYINFOW	PPXUPTR_ENTRYINFO;
#else
	typedef PPXAPPINFOA			PPXAPPINFO;
	typedef PPXCMDENUMSTRUCTA	PPXCMDENUMSTRUCT;
	typedef PPXAPPINFOUNIONA	PPXAPPINFOUNION;
	typedef PPXAPPINFOFUNCTIONA	PPXAPPINFOFUNCTION;
	typedef PPXUPTR_INDEX_UPATHA PPXUPTR_INDEX_UPATH;
	typedef PPXUPTR_TABINDEXSTRA PPXUPTR_TABINDEXSTR;
	typedef PPXUPTR_ENTRYINFOA	PPXUPTR_ENTRYINFO;
#endif

typedef struct {	// %F で使用する構造体
	const TCHAR *source;
	TCHAR dest[VFPS];
} PPXCMD_F;

#define SetLongResult(ppxa, pxs, text) \
	if ( 1 == (ppxa)->Function((ppxa), PPXCMDID_LONG_RESULT, (text)) ){ \
		(pxs).command->resultstring[0] = '\0'; \
	}else{ \
		if ( wcslen(text) < CMDLINESIZE ){ \
			wcscpy((pxs).command->resultstring, text); \
		}else{ \
			wcscpy((pxs).command->resultstring, L"#length over#"); \
		} \
	}

#define PPXRCDATA_CUSTOMIZE_LIST 10601 // モジュール毎の追加設定テキスト
#define TIMERID_TIMERMODULE 0xbf00 // - 0xbfff // timer module で使用
#define TIMERID_USER_MIN 0xC000
#define TIMERID_USER_MAX 0xDFFF
//===================================================================その他定数
/*-----------------------------------------------------------------------------
	PPXCOMMAND_WM : PPx の RegisterWindowMessage 用文字

	これによって、PPx にメッセージを送ることができる。
	送信方法は、通常 PostMessage を使用すること。
ex)
	WM_PPXCOMMAND = RegisterWindowMessage(PPXCOMMAND_WM);
	PostMessage(hPPxWnd, WM_PPXCOMMAND, keyname, param);
	keyname は tkey.h に定義された PPx のコマンド
	param は keyname に対応したパラメータ（通常は 0）
-----------------------------------------------------------------------------*/
#define PPXCOMMAND_WM	"PPxMess"

/*-----------------------------------------------------------------------------
	ウィンドウの管理に用いられる定義済み ID。PPxRegist に用いる

	ID は REGIDSIZE の char で、
	ID[0], ID[1] は PPx の種類を表す文字。以下の define を参照。
	ID[3] は 'A'～'Z' の同種 PPx の管理番号
	ID[4] = '\0' 固定。
-----------------------------------------------------------------------------*/
#define REGIDSIZE 4
#define REGEXTIDSIZE 8 // 一体化窓の C_Zxyy / Zxyy 形式用(x:一体化窓ID, y:追加ID)
#define PPB_REGID		"B_"	// PPb
#define PPC_REGID		"C_"	// PPc
#define PPV_REGID		"V_"	// PPv
#define PPTRAY_REGID	"TR"	// PPtray
#define PPCUST_REGID	"cs"	// PPcust

										// PPx の Window class
#ifdef UNICODE
	#define	PPCWinClass	"PaperPlaneCUIW"
	#define	PPVWinClass	"PaperPlaneVUIW"
	#define	PPTrayWinClass	"PPtrayW"
	#define	PPeditWinClass	"PPeditW"
#else
	#define	PPCWinClass	"PaperPlaneCUI"
	#define	PPVWinClass	"PaperPlaneVUI"
	#define	PPTrayWinClass	"PPtray"
	#define	PPeditWinClass	"PPedit"
#endif

#define PPx_DIRECTORY_COMMENTFILE "00_INDEX.TXT"

/*============================================================  Stack Heap 関連
	Heap 領域にデータを積み重ねて保管する関数群。
	一度に ThSTEP 分だけ確保するため、数K 単位の確保なら有利

・void	ThInit(ThSTRUCT *TH);							構造体の初期化を行う
・BOOL	ThFree(ThSTRUCT *TH);							確保メモリを解放
・BOOL	ThAppend(ThSTRUCT *TH, void *data, SIZE32_T size);	データを追加する
		data	追加するデータへのポインタ
		size	追加するデータの大きさ

・BOOL	ThAddString(ThSTRUCT *TH, char *data);			文字列を追加する
		data	追加する文字列へのポインタ
・BOOL	ThCatString(ThSTRUCT *TH, char *data);			文字列を連結する
		data	追加する文字列へのポインタ
※ThAddStringはASCIIZが列挙された形。
※ThCatStringはThAppend/ThAddStringと混同して使用できない
-----------------------------------------------------------------------------*/
#define ThSTEP 4096				// メモリの確保単位
#define ThSTRUCT_InitData {NULL, 0, 0}

typedef struct tagThSTRUCT {	// Stack Heap 管理構造体
	char *bottom;	// 確保したメモリの先頭
	DWORD top;		// 使用中のメモリの末端
	SIZE32_T size;	// 確保量
} ThSTRUCT;

extern PPXDLL void PPXAPI ThInit(ThSTRUCT *TH);
extern PPXDLL BOOL PPXAPI ThFree(ThSTRUCT *TH);
extern PPXDLL BOOL PPXAPI ThAppend(ThSTRUCT *TH, const void *data, SIZE32_T size);
extern PPXDLL BOOL PPXAPI ThAddString(ThSTRUCT *TH, const TCHAR *data);
extern PPXDLL BOOL PPXAPI ThCatString(ThSTRUCT *TH, const TCHAR *data);
extern PPXDLL BOOL PPXAPI ThSize(ThSTRUCT *TH, SIZE32_T size);
#define ThPointer(TH, offset) ((char *)((TH)->bottom + (offset)))
#define ThPointerT(TH, offset) ((TCHAR *)(char *)((TH)->bottom + (offset)))
#define ThLast(TH) ((char *)((TH)->bottom + (TH)->top))
#define ThNextAllocSizeS(maxsize) ThSTEP
// 16k
#define ThNextAllocSizeM(maxsize) ( (((maxsize) / 4) + ThSTEP) & ~(ThSTEP - 1) )
// 256k
#define ThNextAllocSizeL(maxsize) ( (((maxsize) / 2) + ThSTEP) & ~(ThSTEP - 1) )
#ifdef UNICODE
//	extern PPXDLL BOOL PPXAPI ThAddStringA(ThSTRUCT *TH, const char *data);
extern PPXDLL BOOL PPXAPI ThCatStringA(ThSTRUCT *TH, const char *data);
#else
//	#define ThAddStringA ThAddString
#define ThCatStringA ThCatString
#endif
extern PPXDLL BOOL PPXAPI ThSetString(ThSTRUCT *TH, const TCHAR *name, const TCHAR *str);
extern PPXDLL TCHAR * PPXAPI ThAllocString(ThSTRUCT *TH, const TCHAR *name, SIZE32_T strsize);
extern PPXDLL const TCHAR * PPXAPI ThGetString(ThSTRUCT *TH, const TCHAR *name, TCHAR *str, SIZE32_T strlength);
extern PPXDLL TCHAR * PPXAPI ThGetLongString(ThSTRUCT *TH, const TCHAR *name, TCHAR *str, SIZE32_T strlength);
extern PPXDLL BOOL PPXAPI ThEnumString(ThSTRUCT *TH, int index, TCHAR *name, TCHAR *str, SIZE32_T strlength);

// PP_ExtractMacro 内のThGetString (%so) で取得可能な文字列
#define StringVariable_Command_Title "title"		// %" の内容
#define StringVariable_Command_Response "response"	// %@, %a で生成したパス
#define StringVariable_Command_EditCache "EditText"	// %$ で編集した内容
//#define StringVariable_Command_MenuCacheItem "MenuCache1"	// %M で選択した項目
//#define StringVariable_Command_MenuCacheText "MenuCache2"	// %M で選択した内容

//============================================================== その他
									// PPx に SendMessage
#define PPxSendCommand(hWnd, wParam, lParam) SendMessage(hWnd, WM_PPXCOMMAND, wParam, lParam)
/* テンポラリを作成する
 attribute:
	0 : パスのみ	%temp%\ppxtemp\
	FILE_ATTRIBUTE_NORMAL	 : 0bytes ファイル作成	%temp%\ppxtemp\filename
	FILE_ATTRIBUTE_DIRECTORY : ディレクトリ作成		%temp%\ppxtemp\dirname
	FILE_ATTRIBUTE_LABEL     : パスの決定のみ		%temp%\ppxtemp\name
	FILE_ATTRIBUTE_COMPRESSED : %extracttemp% があれば %temp% の代わりに使用
*/
extern PPXDLL BOOL PPXAPI MakeTempEntry(SIZE32_T bufsize, TCHAR *tempath, DWORD attribute);

// レジストリの文字列を取得する
_Success_(return)
extern PPXDLL BOOL PPXAPI GetRegString(HKEY hKey, const TCHAR *path, const TCHAR *name, _Out_ TCHAR *dest, SIZE32_T destlen);
// PPX.HLP を開く
extern PPXDLL void PPXAPI PPxHelp(HWND hWnd, UINT command, DWORD_PTR data);

// ダイアログで Help 等の補助処理を行う
extern PPXDLL INT_PTR PPXAPI PPxDialogHelper(HWND hDlg, UINT iMsg, WPARAM wParam, LPARAM lParam);

// 同期オブジェクトに付加する文字列を返す（マルチユーザで同時実行するため）
extern PPXDLL const TCHAR * PPXAPI PPxGetSyncTag(void);

// param の算術を行う
#define GetCalc_ResultMaxLength 64
extern PPXDLL BOOL PPXAPI GetCalc(const TCHAR *param, TCHAR *resultstr, int *resultnum);

// CommonControls を使う前に呼び出す
extern PPXDLL HANDLE PPXAPI LoadCommonControls(DWORD usecontrol);

// PPx版 DialogBoxParam。ダイアログのフォント設定を行う
extern PPXDLL INT_PTR PPXAPI PPxDialogBoxParam(HANDLE hinst, const TCHAR *lpszTemplate, HWND hwndOwner, DLGPROC dlgprc, LPARAM lParamInit);
// ダイアログフォント設定済みの DLGITEMTEMPLATE を得る
extern PPXDLL LPDLGTEMPLATE PPXAPI GetDialogTemplate(HWND hParentWnd, HANDLE hinst, LPCTSTR lpszTemplate);

// RunAs 経由で起動していたら、ユーザ名を返す。そうでなければ NULL
extern PPXDLL const TCHAR * PPXAPI CheckRunAs(void);

// WINAPI をまとめて読み込む
typedef struct {
	void (WINAPI **APIptr)();	// 取得したAPIへのポインタ(最終なら NULL)
	const char *APIname;		// 取得するAPIの名前
} LOADWINAPISTRUCT;

// LoadWinAPI mode に使用する値
#define LOADWINAPI_HANDLE		0	// HANDLE を使って読み込み
#define LOADWINAPI_LOAD			1	// LoadLibrary を使って DLL をロード
#define LOADWINAPI_GETMODULE	2	// GetModuleHandle を使って DLL を取得
#define LOADWINAPI_LOAD_ERRMSG	3	// エラーメッセージ付き LOADWINAPI_LOAD
#define LOADWINAPI(ptr, name)	{(void (WINAPI **)())&ptr, name}
#define LOADWINAPI1(name)	{(void (WINAPI **)())&(D ## name), #name}
#define LOADWINAPI1ND(name)	{(void (WINAPI **)())&name, #name}
#define LOADWINAPI1T(name)	{(void (WINAPI **)())&(D ## name), #name TAPITAIL}
extern PPXDLL HMODULE PPXAPI LoadWinAPI(const char *DLLname, HMODULE hDLL, LOADWINAPISTRUCT *apis, int mode);

enum {
	SYSTEMDLL_COMDLG32 = 0,
	SYSTEMDLL_COMCTL32,
	SYSTEMDLL_OLEAUT32,
	SYSTEMDLL_MPR,
	SYSTEMDLL_WINMM,
	SYSTEMDLL_IMAGEHLP,
	SYSTEMDLL_PSAPI,
	SYSTEMDLL_SHCORE,
	SYSTEMDLL_SHLWAPI,
	SYSTEMDLL_RSTRTMGR,
	SYSTEMDLL_ATL,
	SYSTEMDLL_ATL100,
	SYSTEMDLL_UXTHEME,
	SYSTEMDLL_DWMAPI,
	SYSTEMDLL_MLANG,
};
PPXDLL HMODULE PPXAPI LoadSystemDLL(DWORD dllID);

extern PPXDLL int PPXAPI GetAccessApplications(const TCHAR *checkpath, TCHAR *text);

//============================================================== 文字列操作関連
// マルチバイト判定 Is????? 用のテーブルを初期化する
extern PPXDLL void PPXAPI FixCharlengthTable(char *table);

// str からオプションを一つ抽出する
// extern PPXDLL UTCHAR PPXAPI GetOption(LPCTSTR *str, TCHAR *param); 1.17 移行, 1.86 廃止
extern PPXDLL UTCHAR PPXAPI GetOptionParameter(LPCTSTR *commandline, TCHAR *optionname, TCHAR **optionparameter);

// str から１パラメータ分を抽出する
extern PPXDLL UTCHAR PPXAPI GetLineParam(LPCTSTR *str, TCHAR *param); // 1.86 移行
extern PPXDLL UTCHAR PPXAPI GetLineParamS(LPCTSTR *str, TCHAR *param, size_t length);
extern PPXDLL void PPXAPI GetQuotedParameter(LPCTSTR *commandline, TCHAR *param, TCHAR *parammax);

// 空白などを読み飛ばす
extern PPXDLL UTCHAR PPXAPI SkipSpace(LPCTSTR *str);

// , / 空白区切りを検出する
extern PPXDLL BOOL PPXAPI NextParameter(LPCTSTR *str);

// 数値を取得する。成功したかどうかは ptr が変化したかで判定する
extern PPXDLL INT_PTR PPXAPI GetNumber(LPCTSTR *ptr);
#define GetDwordNumber(ptr) (DWORD)GetNumber(ptr)
#define GetIntNumber(ptr) (int)GetNumber(ptr)
// 単位付きで取得する
extern PPXDLL BOOL PPXAPI GetSizeNumber(LPCTSTR *ptr, DWORD *Low, DWORD *High);

// 数値を文字列に変換する
// length は最低4+1文字(123k\0)確保する必要あり
#define XFN_SEPARATOR	B0	// 区切り文字「,」を追加
#define XFN_RIGHT		B1	// 右づめ
#define XFN_LEFT		B2	// 左づめ
#define XFN_UNITSPACE	B3	// 単位用に１桁確保する
#define XFN_HEXUNIT		B4	// 1024単位
#define XFN_MUL			B5	// low * high の値を用いる
#define XFN_MINKILO		B6	// 少なくとも「k」が付く
#define XFN_DECIMALPOINT B9 // 単位使用時小数点を使用(6桁=小数1桁,7以上=小2桁)
#define XFN_MINUNITMASK	(B6 | B7 | B8) // 1:k 2:M 3:G 4:T 5:P 6:E
#define XFN_MINUNITSHIFT 6

#define XFNW_FULL_SEP 26	// 64bit全桁 + XFN_SEPARATOR 12,345,678,901,234,567,890
#define XFNW_FULL_NOSEP 20	// 64bit全桁 12345678901234567890
#define XFNW_FULL32_SEP 13	// 32bit全桁 + XFN_SEPARATOR 1,234,567,890
#define XFNW_FULL32_NOSEP 10	// 32bit全桁 1234567890

extern PPXDLL void PPXAPI FormatNumber(TCHAR *str, DWORD flags, int length, DWORD low, DWORD high);
extern PPXDLL COLORREF PPXAPI GetColor(LPCTSTR *linesrc, BOOL usealias);

															// 全角・半角変換 -
extern PPXDLL TCHAR * PPXAPI Strsd(TCHAR *dststr, const TCHAR *srcstr); // 半角→全角
extern PPXDLL TCHAR * PPXAPI Strds(TCHAR *dststr, const TCHAR *srcstr); // 全角→半角
															// 大小変換 -------
extern PPXDLL TCHAR * PPXAPI Strupr(TCHAR *str); // 英大文字
extern PPXDLL TCHAR * PPXAPI Strlwr(TCHAR *str); // 英小文字
															// Key Code 変換 --
												// 「\, ^, &, @」を文字列化
extern PPXDLL TCHAR * PPXAPI PutShiftCode(TCHAR *str, int key);
extern PPXDLL int PPXAPI GetKeyCode(LPCTSTR *ptr);	// 文字列からkeycodeを得る
extern PPXDLL void PPXAPI PutKeyCode(TCHAR *str, int key);	// keycodeを文字列にする
extern PPXDLL int PPXAPI CheckRegistKey(const TCHAR *src, TCHAR *dest, const TCHAR *type); // キーの正当性をチェックして正規化する
#define CHECKREGISTKEY_RELOADKEY 2
#define CHECKREGISTKEY_OK 1
#define CHECKREGISTKEY_BADNAME -1
#define CHECKREGISTKEY_WARNKEY 0

// テキストからウィンドウハンドルを求める(*focus等で使用)。ないなら NULL
extern PPXDLL HWND PPXAPI GetWindowHandleByText(PPXAPPINFO *ppxa, const TCHAR *param);
//====================================================== 拡張エディットボックス
typedef struct tagTINPUT {
	HWND hOwnerWnd;	// 親のウィンドウハンドル
	WORD hRtype;		// 読込ヒストリの種類
	WORD hWtype;		// 書込ヒストリの種類
	const TCHAR *title;	// タイトルバーの内容
	TCHAR *buff;		// 編集文字列格納
	SIZE32_T size;		// 最大入力文字数

	int flag;
#define TIEX_USEREFLINE	B0	// 参照行を使う。参照行の内容は"buff"を使用
#define TIEX_USESELECT	B1	// 選択範囲の指定を使う
#define TIEX_USEPNBTN	B2	// Prev/Nextボタンを使う。要TIEX_USEREFLINE
#define TIEX_ID			B3	// %si"Input_ID" が ID
#define TIEX_EXECPRECMD	B4	// 予め %si"Input_FirstCmd" を実行する
#define TIEX_REFMODE	B5	// Refを始めから使用する
#define TIEX_TOP		B6	// 強制前面(SetForegroundWindowを明示指定)
#define TIEX_USEINFO	B7	// infoを使用する
#define TIEX_REFTREE	B8	// Refは Dir Tree(PPXEDIT_REFTREEと同じ)
#define TIEX_SINGLEREF	B9	// Refは すべていれかえ(PPXEDIT_SINGLEREFと同じ)
#define TIEX_REFEXT		B10	// 拡張子分離処理有り
#define TIEX_USEOPTBTN	B11	// Option Button を用意する
#define TIEX_INSTRSEL	B12	// "|"による範囲指定を有効にする(PPXEDIT_INSTRSELと同じ)
#define TIEX_NCHANGE	B13	// 更新通知有り。更新があったらK_EDITCHANGEが送信
#define TIEX_LINE_MULTI	B14	// １行編集の複数行版(PPXEDIT_LINE_MULTIと同じ)
#define TIEX_LEAVECANCEL	B15	// 非アクティブでキャンセル
#define TIEX_FIXFORPATH	B16	// ファイル名に使用できない文字を除去する
#define TIEX_FIXFORDIGIT	B17	// 全角数字を半角数字に
#define TIEX_LINE_MULTILINE	B18	// 改行を含むテキストの編集
	int firstC;		// 選択開始位置
	int lastC;		// 選択終了位置
	PPXAPPINFO *info;	// TIEX_USEINFOで有効 ( flag !TIEX_USEINFO のとき、tInputExで値がNULLに変更される) (0.37追加)
	ThSTRUCT *StringVariable;	// TIEX_EXECPRECMD で使用する文字列変数領域(1.64追加)
} TINPUT;
extern PPXDLL int PPXAPI tInput(HWND hWnd, const TCHAR *title, TCHAR *string, int maxlen, WORD readhist, WORD writehist);
extern PPXDLL int PPXAPI tInputEx(TINPUT *tinput);
/*-----------------------------------------------------------------------------
	エディットボックスを拡張する

	hED		拡張するエディットボックス／コンボボックスのハンドル
	maxlen	最大文字列数
	defstr	初期文字列
	rHist	Edit Boxのときに設定するヒストリの内容
	wHist	Edit Boxのときに設定するヒストリの内容(優先)
	flags	下記フラグを使用可能 */
#define PPXEDIT_COMBOBOX	B0	// 0:対象はEditBox		1:ComboBox
#define PPXEDIT_USEALT		B1	// Alt系の拡張コマンドを使用できるようにする
#define PPXEDIT_TABCOMP		B2	// TAB による補完を有効
#define PPXEDIT_JOINTTREE	B3	// Dir Tree は 一体型
#define PPXEDIT_NOINCLIST	B4	// 自動補完リストを無効にする
// B5 未使用
#define PPXEDIT_WANTENTER	B6	// Enter/ESCでWM_COMMANDに キーコード を通知
#define PPXEDIT_WANTEVENT	B7	// Event実行有り
#define PPXEDIT_REFTREE		B8	// Refは Dir Tree
#define PPXEDIT_SINGLEREF	B9	// Refは すべていれかえ
#define PPXEDIT_PANEMODE	B10	// マウスによるアクティブ時の動作を変える
#define PPXEDIT_WANTALLKEY	B11	// ダイアログ用のキーを無効にする
#define PPXEDIT_INSTRSEL	B12	// "|"による範囲指定を有効にする
#define PPXEDIT_NOWORDBREAK	B13	// ES_MULTILINE時に、単語の途中で改行する(Vista以降)
#define PPXEDIT_LINE_MULTI	B14	// １行編集の複数行版(要ES_MULTILINE)
#define PPXEDIT_ENABLE_WIDTH_CHANGE B15	// サイズ変更キー(&\←→)を有効に
#define PPXEDIT_ENABLE_HEIGHT_CHANGE B16	// サイズ変更キー(&\↑↓)を有効に
#define PPXEDIT_ENABLE_SIZE_CHANGE (B15 | B16)	// サイズ変更キーを有効に
#define PPXEDIT_DISABLE_CHECKSAVE B17	// (PPe)終了時の保存確認を無効にする
#define PPXEDIT_SAVE_BYCLOSE B18	// (PPe)PPXEDIT_DISABLE_CHECKSAVE有効時、終了時に保存する
#define PPXEDIT_LINE_MULTILINE	B19	// １行編集の複数行編集(要ES_MULTILINE)
// B20 以降は内部で利用
//-----------------------------------------------------------------------------
extern PPXDLL HWND PPXAPI PPxRegistExEdit(PPXAPPINFO *info, HWND hED, int maxlen, const TCHAR *defstr, WORD rHist, WORD wHist, int flags);
//---------------------------------------------------- ファイル検索
#define EXTPATHEXT ".COM;.EXE;.BAT;.CMD"
				// flags に指定する値
									// 実行ファイル優先処理の方法
#define CMDSEARCH_OFF		0	// 優先なし
#define CMDSEARCH_CURRENT	B0	// カレントファイルを優先
									// 動作指定
#define CMDSEARCH_MAKELIST	B6	// 一覧生成モード
#define CMDSEARCH_NOADDSEP	B7	// ディレクトリのとき、「\」を付加しない
#define CMDSEARCH_EDITBOX	B8	// エディットボックス用処理
#define CMDSEARCH_MULTI		B9	// 複数パラメータ指定のため「"」が必要
#define CMDSEARCH_DIRECTORY	B10	// ディレクトリ限定

#define CMDSEARCH_ONE		B11	// ループしない
#define CMDSEARCH_FLOAT		B12	// 部分一致検索
#define CMDSEARCH_NOUNC		B13	// UNCは検索しない
#define CMDSEARCH_ROMA		B14	// ローマ字検索(要 romahandle = 0 に初期化)
#define CMDSEARCH_WILDCARD	B15	// 「.」区切り検索(要 romahandle = 0 に初期化)
#define CMDSEARCH_REGEXP	B5	// 正規表現検索(要 romahandle = 0 に初期化)
// B16 以降は内部で使用する

typedef struct {	// 編集処理共通構造体
// 要設定
	PPXAPPINFO *info;
	HANDLE hF;				// ファイル検索用ハンドル
// 設定不要
	int cmdsearch;		// コマンド検索モード CMDSEARCH_
	TCHAR Fsrc[VFPS + 4];	// ファイル検索・連続検索用検索文字列
	TCHAR Fname[VFPS];	// ファイル検索・連続検索用キャッシュ
	TCHAR *FnameP;		// ファイル検索・連続検索用書込位置
	TCHAR *Fword;
	DWORD_PTR romahandle;	// CMDSEARCH_ROMA/CMDSEARCH_WILDCARD 指定時要 0 設定
} ESTRUCT;

typedef struct {
	DWORD start, end;
} ECURSOR;

#define GWSF_SPLIT_PARAM B0

#define GWS_BRAKET_NONE 0
#define GWS_BRAKET_LEFT 1
#define GWS_BRAKET_LEFTRIGHT 2
extern PPXDLL int PPXAPI GetWordStrings(TCHAR *line, ECURSOR *cursor, DWORD flags);
extern PPXDLL TCHAR * PPXAPI SearchFileIned(ESTRUCT *ED, TCHAR *line, ECURSOR *cursor, int flags);

															// テキストスタック
extern PPXDLL void PPXAPI PushTextStack(TCHAR mode, TCHAR *text);
extern PPXDLL void PPXAPI PopTextStack(TCHAR *mode, TCHAR *text);
															// システムフック
extern PPXDLL BOOL PPXAPI PPxHookEdit(int local);
// local: -1:フック解放、0:グローバルフック、1:プロセス内フック
//======================================================== 外部プロセス実行関連
//------------------------------ 実行関係のオプション( %O )
//						 0         1         2         3
//						 012345678 9012345678901234567890
#define XEO_STRINGS		"ABbkcCem\6EOsRrTWpDuPIiNSqxndalv"
#define XEO_STRLENGTH	31 // \0 除く
// B6 が未使用

// マクロ文字展開制御
#define XEO_DIRWILD		B0	// A:ディレクトリの末尾を \*.*
#define XEO_NOEXECMARK	B12	// R: r オプションを無効にする
#define XEO_EXECMARK	B13	// r:マーク全てを対象
#define XEO_DISPONLY	B17	// D:メニュー表示用処理
#define XEO_UTF8		B18	// u:レスポンスファイル等でutf-8を使用する
//#define XEO_PATHLHA		//  :UNLHA32.DLL向けに先頭「-」を変換
#define XEO_PATHSLASH	B23	// S:パス区切りを「/」に変更
#define XEO_PATHESCAPE	B16	// p:[,]を\[,\]にエスケープ
// 編集制御
#define XEO_INRETURN	B5	// C:改行をパラメータに含める
//#define XEO_NOEDITR	B6	// e:連続実行時に編集しない
#define XEO_NOEDIT		B22	// N:編集しない
// PP_ExtractMacro の extract パラメータ使用時の動作
#define XEO_EXTRACTLONG	B8	// -: LONGEXTRACTPARAM を使用している
#define XEO_EXTRACTEXEC	B9	// E:展開時に *command を実行する

// ファイル実行時制御
//起動時ウィンドウ
#define XEO_MAX			B25	// x:最大化状態で実行
#define XEO_MIN			B26	// n:最小化状態で実行
#define XEO_HIDE		B27	// d:表示しない
#define XEO_NOACTIVE	B28	// a:フォーカス変更無し
#define XEO_LOW			B29	// l:実行優先度がLOW
//コンソール関連
#define XEO_USEPPB		B1	// B:必ずPPBを使用
#define XEO_NOUSEPPB	B2	// b:PPBは使用しない
//コマンドラインシェル関連
#define XEO_NOCMDCMD	B3	// k:不明なコマンドをコマンドライン実行しない
#define XEO_USECMD		B4	// c:コマンドラインシェルを使用する
#define XEO_NOPIPERDIR	B19	// P:パイプ・リダイレクト処理を無効にする
//待機・実行後処理
#define XEO_EXTRACTLOG	B7	// m:展開結果をログ出力する
#define XEO_SEQUENTIAL	B11	// s:順次実行する
#define XEO_DOWNCSR		B10	// O:実行後カーソル移動
#define XEO_DELTEMP		B14	// T:レスポンスファイルを削除
#define XEO_WAITKEY		B15	// W:キー入力待ち
#define XEO_IGNOREERR	B20	// I:エラーを無視して続行する
#define XEO_WAITIDLE	B21	// i:起動が完了するまで待つ
#define XEO_WAITQUIET	B24	// q:待機中にダイアログを表示しない
//その他
//#define XEO_BATCH		B3	//  :バッチ実行
#define XEO_NOSCANEXE	B30	// v:X_execsによるチェックを行わない
#define XEO_CONSOLE		B31	// -:コンソール上で実行(PPb内部用)

extern PPXDLL BOOL PPXAPI ComExec(HWND hOwner, const TCHAR *line, const TCHAR *path);

// GetExecType の戻り値
enum {
	GTYPE_ERROR,	// 0:ファイルがない等のエラー
	GTYPE_GUI,		// 1:Win32 GUI
	GTYPE_CONSOLE,	// 2:Win32 Console/DOS
	GTYPE_DATA,		// 3:Data/Dir
	GTYPE_SHELLEXEC	// 4:ShellExe
};
extern PPXDLL int PPXAPI GetExecType(LPCTSTR *name, TCHAR *fpath, const TCHAR *path);
extern PPXDLL ERRORCODE PPXAPI GetExecuteErrorReason(const TCHAR *filename, TCHAR *reason);
extern PPXDLL BOOL PPXAPI CheckExebin(const TCHAR *path, int type);

/*-----------------------------------------------------------------------------
	ShellExecute を実行し、エラーがでたら文字列を取得する。
	失敗したら NULL を返す    ※HINSTANCE は解放不要である
-----------------------------------------------------------------------------*/
extern PPXDLL HANDLE PPXAPI PPxShellExecute(HWND hwnd, LPCTSTR lpOperation, LPCTSTR lpFile, LPCTSTR lpParameters, LPCTSTR lpDirectory, int flag, TCHAR *ErrMsg);

//================================================================== マクロ関連
#define CMDLINESIZE 0x400		// コマンドラインの文字数
#define MAXCMDLINESIZE 0x8000	// OSが受付可能なコマンドラインの最大文字数
#define CMDEXE_LENGTH 8000		// CME.EXE の受付可能な文字数

typedef union tagLONGEXTRACTPARAM { // XEO_EXTRACTLONG が有効のときに使用
	TCHAR text[CMDLINESIZE]; // CMDLINESIZE を越えない場合の保存先
	struct { // 戻り値が ERROR_PARTIAL_COPY のときに使用する
		union tagLONGEXTRACTPARAM *id; // 自分自身を示すポインタ(union有効id)
		ThSTRUCT th; // 展開先
	} longtext;
} LONGEXTRACTPARAM;
#define PP_InitLongParam(param) (((LONGEXTRACTPARAM *)(param))->longtext.id = ((LONGEXTRACTPARAM *)(param)))
#define PP_GetLongParamRAW(param) ((TCHAR *)((LONGEXTRACTPARAM *)(param))->longtext.th.bottom)
#define PP_GetLongParam(param, result) ( ((result) != ERROR_PARTIAL_COPY) ? (param) : PP_GetLongParamRAW(param) )
#define PP_FreeLongParamRAW(param) {ThFree(&((LONGEXTRACTPARAM *)(param))->longtext.th);}
#define PP_FreeLongParam(param, result) {if ( (result) == ERROR_PARTIAL_COPY ) PP_FreeLongParamRAW(param);}

extern PPXDLL void PPXAPI PPxRegGetIInfo(PPXAPPINFO *ptr);
extern PPXDLL ERRORCODE PPXAPI PP_ExtractMacro(HWND hWnd, PPXAPPINFO *ParentInfo, POINT *pos, const TCHAR *param, TCHAR *extract, int flag);


extern PPXDLL LRESULT PPXAPI AnswerExtractCall(PPXAPPINFO *info, WPARAM wParam, LPARAM lParam);
extern PPXDLL const TCHAR * PPXAPI SendExtractCall(HWND hTargetWnd, const TCHAR *src);

// PP_GetExtCommand 戻値
// -1 以外の負の値は、PPEXTRESULT_xxx において、該当実行内容がない場合
#define PPEXTRESULT_NONE	-1 // 該当無し
#define PPEXTRESULT_FILE	0
#define PPEXTRESULT_DIR		1 // PP_GetExtCommand では未使用
#define PPEXTRESULT_VFSDIR	2
#define PPEXTRESULT_LINK	3
#define PPEXTRESULT_CLSID	4

extern PPXDLL int PPXAPI PP_GetExtCommand(const TCHAR *src, const TCHAR *ID, TCHAR *cmdbuf, TCHAR *TypeName);

//================================================================ ヒストリ管理
// 書込み指定用
#define PPXH_GENERAL	B0	// g 汎用
#define PPXH_NUMBER		B1	// n 数字
#define PPXH_COMMAND	B10	// h コマンドライン

#define PPXH_DIR		B2	// d ディレクトリ
#define PPXH_FILENAME	B12	// c ディレクトリ無しのファイル名
#define PPXH_PATH		B11	// f フルパスファイル名

#define PPXH_SEARCH		B9	// s 検索用文字列
#define PPXH_MASK		B13	// m エントリ名マスク(PPCの[F])

#define PPXH_PPCPATH	B14	// p PPC ディレクトリヒストリ(バイナリ付き)
#define PPXH_PPVNAME	B8	// v PPV 表示ファイル名(バイナリ付き)
// ユーザ用(標準では使用しない)
#define PPXH_USER1		B5	// uU
#define PPXH_USER2		B4	// xX

// 特別指定
#define PPXH_NETPCNAME	B7	// ネットワークPC列挙用のキャッシュ
#define PPXH_ROMASTR	B6	// ローマ字検索用のキャッシュ

// 入力時の参照読込み指定用(バイナリ参照不可)
#define PPXH_GENERAL_R	(PPXH_GENERAL | PPXH_COMMAND | PPXH_FILENAME | PPXH_PATH)
#define PPXH_NUMBER_R	(PPXH_NUMBER)
#define PPXH_COMMAND_R	(PPXH_COMMAND)

#define PPXH_DIR_R		(PPXH_DIR | PPXH_PATH | PPXH_PPCPATH)
#define PPXH_FILENAME_R	(PPXH_FILENAME)
#define PPXH_PATH_R		(PPXH_PATH | PPXH_NAME_R)

#define PPXH_SEARCH_R	(PPXH_SEARCH | PPXH_MASK)
#define PPXH_MASK_R		(PPXH_SEARCH | PPXH_MASK)

#define PPXH_PPCPATH_R	(PPXH_PPCPATH)
#define PPXH_PPVNAME_R	(PPXH_PPVNAME)

#define PPXH_WILD_R	(PPXH_MASK_R)	// ワイルドカードを含むもの
												// ファイル名指定
#define PPXH_NAME_R	(PPXH_PPVNAME | PPXH_SEARCH | PPXH_FILENAME | PPXH_MASK)

#define PPXH_ALL_R	MAX16	// 全て
// 特殊
#define PPXH_NOOVERLAP	B31 // SearchPHistory でstr が一致しないものを優先する

// 情報取得
#define GetHistoryData(strptr) ((BYTE *)((BYTE *)(strptr) + *(WORD *)((BYTE *)(strptr) - 2)))
#define GetHistoryDataSize(strptr) ((WORD)(*(WORD *)((BYTE *)(strptr) - 6) - *(WORD *)((BYTE *)(strptr) - 2) - 6))

extern PPXDLL void PPXAPI InitHistory(void);
extern PPXDLL const TCHAR * PPXAPI SearchHistory(WORD type, const TCHAR *str);
extern PPXDLL const TCHAR * PPXAPI SearchPHistory(DWORD type, const TCHAR *str);
extern PPXDLL void PPXAPI WriteHistory(WORD type, const TCHAR *str, WORD b_size, void *bin);
extern PPXDLL const TCHAR * PPXAPI EnumHistory(WORD type, int No);	// No ≧ 0
extern PPXDLL BOOL PPXAPI DeleteHistory(WORD type, const TCHAR *str);
extern PPXDLL WORD PPXAPI GetHistoryType(const TCHAR **param);

//============================================================ カスタマイズ管理
#define CUST_NAME_LENGTH (VFPS + 8) // カスタマイズの name, subname の最大長

extern PPXDLL void PPXAPI InitCust(void);
extern PPXDLL int PPXAPI SetCustData(const TCHAR *name, const void *bin, size_t b_size); // 成功なら NO_ERROR
extern PPXDLL void *GetCustValue(const TCHAR *name, const TCHAR *subname, void *bin, size_t b_size);
extern PPXDLL int PPXAPI GetCustData(const TCHAR *name, void *bin, size_t b_size); // 成功なら NO_ERROR
#define IsExistCustData(name) (GetCustData(name, NULL, 0) >= 0)
#define GetCustDataSize(name) GetCustData(name, NULL, 0)
extern PPXDLL int PPXAPI DeleteCustData(const TCHAR *name); // 成功なら NO_ERROR
extern PPXDLL int PPXAPI EnumCustData(int offset, TCHAR *name, void *bin, size_t b_size);
extern PPXDLL int PPXAPI SetCustTable(const TCHAR *name, const TCHAR *subname, const void *bin, size_t b_size); // 成功なら NO_ERROR
extern PPXDLL int PPXAPI InsertCustTable(const TCHAR *name, const TCHAR *subname, DWORD index, const void *bin, size_t b_size); // 成功なら NO_ERROR
extern PPXDLL int PPXAPI SetCustStringTable(const TCHAR *name, const TCHAR *subname, const TCHAR *string, int keep_length);
extern PPXDLL int PPXAPI GetCustTable(const TCHAR *name, const TCHAR *subname, void *bin, size_t b_size); // 成功なら NO_ERROR
#define IsExistCustTable(name, subname) (GetCustTable(name, subname, NULL, 0) >= 0)
#define GetCustTableSize(name, subname) GetCustTable(name, subname, NULL, 0)
#define CountCustTable(name) EnumCustTable(-1, name, NULL, NULL, 0)

extern PPXDLL int PPXAPI DeleteCustTable(const TCHAR *name, const TCHAR *subname, int index); // 成功なら NO_ERROR
extern PPXDLL int PPXAPI EnumCustTable(int offset, const TCHAR *name, TCHAR *subname, void *bin, size_t b_size);
extern PPXDLL int PPXAPI SortCustTable(const TCHAR *name, int (PPXAPI *func)(const BYTE *cust1, const BYTE *cust2));
// SortCustTable の戻り値
#define SORTCUSTTABLE_SORT B30 // 今回ソートをした
#define SORTCUSTTABLE_EXTDATA B31 // 追加データを発見
//================================================= ヒストリ・カスタマイズ情報
// PPx の DB(ヒストリ／カスタマイズ領域)の状況を入手する
typedef struct {
	DWORD structsize;			// この構造体の大きさ/拡張用
	DWORD histsize, histfree;	// ヒストリDBの大きさと空き
	DWORD custsize, custfree;	// カスタマイズDBの大きさと空き
	TCHAR *custpath;			// カスタマイズファイルの既定位置
} PPXDBINFOSTRUCT;
extern PPXDLL void PPXAPI GetPPxDBinfo(PPXDBINFOSTRUCT *dbinfo);

// DB の大きさを変更する
#define PPXDB_HISTORY	1
#define PPXDB_CUSTOMIZE	2
extern PPXDLL void PPXAPI SetPPxDBsize(int mode, SIZE32_T size);

//======================================================================== 同期

#if PPXSYNCDEBUG
extern PPXDLL void PPXAPI UsePPxDebug(const char *filename, int fileline);
extern PPXDLL void PPXAPI FreePPxDebug(const char *filename, int fileline);
#define UsePPx() UsePPxDebug(__FILE__, __LINE__)
#define FreePPx() FreePPxDebug(__FILE__, __LINE__)
#else
extern PPXDLL void PPXAPI UsePPx(void);
extern PPXDLL void PPXAPI FreePPx(void);
#endif

/*-----------------------------------------------------------------------------
	filename で指定されたファイルをメモリに読み込む。ただし、10M 以上は読まない
	margin:		ファイルの末尾以降 0 padding する長さ
	image		読み込み先。NULL なら Systemheap を用い内部で確保。
							例)HeapFree(GetProcessHeap(), 0, image);
	imagesize	読み込んだイメージの大きさ。image != NULL なら確保済みの
				メモリの大きさをいれておく。
	filesize	ファイルのサイズ。4G over なら LFI_OVERSIZE(0xffffffff) が入る
				NULL, LFI_ALWAYSLIMIT, LFI_ALWAYSLIMITLESS 指定可
-----------------------------------------------------------------------------*/
#define LFI_ALWAYSLIMIT		(DWORD *)(DWORD_PTR)1	// 常にサイズ制限を行う
#define LFI_ALWAYSLIMITLESS	(DWORD *)(DWORD_PTR)2	// 常に制限しない
#define LFI_OVERSIZE	MAX32
extern PPXDLL ERRORCODE PPXAPI LoadFileImage(const TCHAR *filename, DWORD margin, char **image, SIZE32_T *imagesize, SIZE32_T *filesize);
// MultiByte, UNICODE の変換機能付き LoadFileImage
extern PPXDLL ERRORCODE PPXAPI LoadTextData(const TCHAR *filename, TCHAR **image, TCHAR **readpoint, TCHAR **maxptr, UINT codepage);
extern PPXDLL ERRORCODE PPXAPI LoadTextImage(const TCHAR *filename, TCHAR **image, TCHAR **readpoint, TCHAR **maxptr);
extern PPXDLL int PPXAPI GetTextCodeType(const BYTE *image, SIZE32_T size);
extern PPXDLL int PPXAPI GetFileCodeType(const TCHAR *filename, const BYTE *image, DWORD memsize);
// 新しいイメージを作成したときは、負の値
extern PPXDLL int PPXAPI FixTextImage(const char *src, DWORD memsize, TCHAR **dest, UINT usecp);
extern PPXDLL int PPXAPI FixTextFileImage(const TCHAR *filename, const char *src, DWORD memsize, TCHAR **dest, UINT usecp);

#define VERIFYZIP_SUCCEDD 0
#define VERIFYZIP_FAILED 1
#define VERIFYZIP_NOSUPPORTVERIFY 2
// ※ image の署名部分が破壊される
extern PPXDLL int PPXAPI VerifyImage(BYTE *image, DWORD imagesize);
extern PPXDLL BOOL PPXAPI GetFileHash(const TCHAR *filename, const TCHAR *hashtype, TCHAR *hashtext);

//================================================================== ウィンドウ
extern PPXDLL void PPXAPI CenterWindow(HWND hWnd);
extern PPXDLL void PPXAPI MoveCenterWindow(HWND hWnd, HWND hParentWnd);
extern PPXDLL void PPXAPI MoveWindowByKey(HWND hWnd, int offx, int offy);
extern PPXDLL void PPXAPI ForceSetForegroundWindow(HWND hWnd);
extern PPXDLL int PPXAPI FixUxTheme(HWND hWnd, const TCHAR *classname);

//==================================================================== メニュー
/*-----------------------------------------------------------------------------
int	PP_GetContextMenu(HMENU menu, char *fullpath, char *ext);

	ContextMenu を作成してメニューに登録する。

	HMENU menu		登録先
	char  *fullpath	対象ファイルのフルパス名
	char  *ext		対象ファイルの拡張子名(".ext"を示す）

戻り値：登録した個数
-----------------------------------------------------------------------------*/
typedef struct {
	ThSTRUCT th;
	DWORD id;
} PPXMENUDATAINFO;

extern PPXDLL int PPXAPI PP_GetContextMenu(HMENU menu, const TCHAR *ext, DWORD *ID);

extern PPXDLL int PPXAPI GetExtentionMenu(HMENU hSubMenu, const TCHAR *ext, PPXMENUDATAINFO *pmdi); // GetExtensionMenu の誤記

/*-----------------------------------------------------------------------------
		ポップアップメニューにカスタマイズ内容を設定する。

	hWnd	親ウィンドウ
	hMenu	挿入するメニューハンドル、NULL なら内部で生成
	id		登録を開始する識別番号、使用された分だけ増える
	Cname	メニュー定義がされているカスタマイズ名
	TH		登録した順に文字列が格納される、ThInit などで初期化しておくこと
			※先頭文字が 0x1 の場合は、メニューハンドルであるため、無視すること
			  '0x1' HMENU MenuNamez

->	NULL なら失敗
-----------------------------------------------------------------------------*/
extern PPXDLL HMENU PPXAPI PP_AddMenu(PPXAPPINFO *ParentInfo, HWND hWnd, HMENU hMenu, DWORD *id, const TCHAR *Cname, ThSTRUCT *TH);
extern PPXDLL UINT PPXAPI FindMenuItem(HMENU hMenu, const TCHAR *name, int *menuPosition);

#ifndef __cplusplus
#define MENUNAMEID_SIG '\x1'
typedef struct {
	TCHAR id; // MENUNAMEID_SIG
	HMENU hMenu;
	TCHAR menuname[];
} MENUNAMEID;

#define GetMenuDataMacro2(ptr, thmenu, id)\
{const TCHAR *maxptr;\
	(ptr) = (const TCHAR *)((thmenu)->bottom);\
	maxptr = (const TCHAR *)ThLast((thmenu));\
	GetMenuDataMacro((ptr), maxptr, id);\
}

#define GetMenuDataMacro(ptr, maxptr, id)\
{int search__id = 0, target__id = id;\
	for ( ; ; ){\
		if ( ptr >= maxptr ){\
			ptr = NULL;\
			break;\
		}\
		if ( *ptr == MENUNAMEID_SIG ){\
			ptr += TSIZEOF(MENUNAMEID);\
		}else{\
			if ( target__id <= search__id ) break;\
			search__id++;\
		}\
		ptr += tstrlen(ptr) + 1;\
	}\
}
#endif


// WM_MENUCHAR, WM_MENUDRAG, WM_MENURBUTTONUP のハンドラ
extern PPXDLL LRESULT PPXAPI PPxMenuProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);
typedef struct {
	PPXAPPINFO *info;
	HMENU hMenu;	// 使用するメニュー
	TCHAR *Command;	// 初期化不要/プロセス実行する必要があるならNULL以外
	int index;		// 初期化不要/キーで選択された ByCommand index(WM_MENUSELECTで取得)
	DWORD commandID; // 右クリックで呼び出すときに使用するコマンドのID
	ThSTRUCT th;
} PPXMENUINFO;
extern PPXDLL BOOL PPXAPI PPxSetMenuInfo(HMENU hMenu, PPXMENUINFO *xminfo);

typedef struct {
	DWORD_PTR key;		// 0x10000 未満…キーコード、0x10000 以上…コマンド
	const TCHAR *str;	// 表示文字列
} PPXINMENU;
typedef struct {
	const TCHAR *name; // 表示名 ? 指定可能。
	PPXINMENU *menus;
} PPXINMENUBAR;
#define PPXINMENY_SEPARATE PTRMAX // 行区切り
//#define PPXINMENY_BARBREAK	B15 // 桁区切り

typedef struct {
	HMENU hMenuBarMenu, hSystemMenu, *hMenuPopups;
	DWORD BarIDmin, BarIDmax, SysmenuOffset, PopupID;
	ThSTRUCT thMenuData;
	const TCHAR *MenuName;
	const PPXINMENUBAR *inmenu;
	BOOL Sysmenu;
} DYNAMICMENUSTRUCT;

extern PPXDLL HMENU PPXAPI InitDynamicMenu(DYNAMICMENUSTRUCT *dms, const TCHAR *barname /* static! */, const PPXINMENUBAR *inmenu);
extern PPXDLL void PPXAPI FreeDynamicMenu(DYNAMICMENUSTRUCT *dms);
extern PPXDLL void PPXAPI DynamicMenu_InitMenu(DYNAMICMENUSTRUCT *dms, HMENU hMenuBar, BOOL showbar);
extern PPXDLL BOOL PPXAPI DynamicMenu_InitPopupMenu(DYNAMICMENUSTRUCT *dms, HMENU hPopupMenu, PPXAPPINFO *info);
extern PPXDLL void PPXAPI InitSystemDynamicMenu(DYNAMICMENUSTRUCT *dms, HWND hWnd);
extern PPXDLL void PPXAPI SystemDynamicMenu(DYNAMICMENUSTRUCT *dms, PPXAPPINFO *info, WORD key);
extern PPXDLL void PPXAPI CommandDynamicMenu(DYNAMICMENUSTRUCT *dms, PPXAPPINFO *info, WPARAM wParam);

// カラム拡張の一覧メニューを hMenu に挿入。不要になったら、targetpath = NULL で実行することにより、カラムハンドラとThSTRUCTを解放できる
extern PPXDLL void PPXAPI GetColumnExtMenu(_Inout_ ThSTRUCT *thEcdata, const TCHAR *targetpath, HMENU hMenu, DWORD index);

/* GetColumnExtMenu の ThSTRUCT 内部構造
	{int} x 1 					項目へのオフセット
	{IColumnProvider *} x N		使用中のIColumnProviderの一覧。最後はNULL
	{COLUMNEXTDATAINFO} x N		項目, 次, そのID, 該当IColumnProvider。(menuナシの時は 各COLUMNEXTDATAINFO の次項目名の文字列が続く)
		COLUMNEXTDATAINFO には、COLUMNEXTDATAINFO_ICP がぶら下がる
*/

#if defined(PROPERTYKEY_DEFINED) && defined(SHCDF_UPDATEITEM)
typedef struct {
	DWORD next; // 同じコラム名(SHCOLUMNIDが一致)のCOLUMNEXTDATAINFO_ICP への位置
	IColumnProvider *Icp; // NULL のときは、IPropertyDescription
} COLUMNEXTDATAINFO_ICP;

typedef struct {
	DWORD next; // 次(SHCOLUMNIDが異なる)の COLUMNEXTDATAINFO への位置、末尾は 0x7fffffff
	union {
		SHCOLUMNID scid;
		PROPERTYKEY pkey;
	} id;
	COLUMNEXTDATAINFO_ICP icps;
	DWORD next_pkey;
} COLUMNEXTDATAINFO;
#endif

#define CAPTIONSEPARATOR T("        ")
extern PPXDLL void PPXAPI SetMessageOnCaption(HWND hWnd, const TCHAR *message);
extern PPXDLL void PPXAPI GetDesktopRect(HWND hWnd, RECT *desktop);

#define PPXFONT_F_mes	0	// Default_F_mes / GetSystemMetrics(SM_CYMENU)
#define PPXFONT_F_dlg	1	// Default_F_mes / GetSystemMetrics(SM_CYMENU)
#define PPXFONT_F_fix	2	// F_mes / GetSystemMetrics(SM_CYMENU)
#define PPXFONT_F_unfix	3	// Default_F_unfix / フォント高さは別設定
#define PPXFONT_F_tree	4	// Tree font を設定しない / T_DEFAULTFONTSIZE
#define PPXFONT_F_ctrl	5	// SystemParametersInfo(SPI_GETNONCLIENTMETRICS / T_DEFAULTFONTSIZE

typedef struct {
	LOGFONT font;
	DWORD dpi;
} LOGFONTWITHDPI;

extern PPXDLL void PPXAPI GetPPxFont(int type, DWORD dpi, LOGFONTWITHDPI *font);

//================================================================== キーボード
extern PPXDLL DWORD PPXAPI GetShiftKey(void);

typedef ERRORCODE (PPXAPI *EXECKEYCOMMANDFUNCTION)(PPXAPPINFO *info, WORD key);
typedef struct {
	EXECKEYCOMMANDFUNCTION Command;
	const TCHAR *CustName1;
	const TCHAR *CustName2;
} EXECKEYCOMMANDSTRUCT;

extern PPXDLL ERRORCODE PPXAPI ExecKeyCommand(const EXECKEYCOMMANDSTRUCT *ekcs, PPXAPPINFO *info, WORD key);
extern PPXDLL void PPXAPI SetIMEDefaultStatus(HWND hWnd);
extern PPXDLL void PPXAPI SetIMEStatus(HWND hWnd, int status);

//============================================================== ワイルドカード
typedef struct {
	BYTE b[0x20]; // 解析内容の保存場所、ここに収まらないときは別途自動確保
} FN_REGEXP;

#define REGEXPF_PATHMASK B0 // 'd' ディレクトリも検索対象
#define REGEXPF_NOEXT B1 // 'x' 拡張子を無視する
#define REGEXPF_SILENTERROR B2 // 'e' エラー表示なし
#define REGEXPF_WORDMATCH B3 // 'w' 単語検索(部分一致、空白区切り有り)
#define REGEXPF_REQ_ATTR B4 // 属性情報が必要
#define REGEXPF_REQ_SIZE B5 // 大きさ情報が必要
#define REGEXPF_REQ_TIME B6 // 時刻情報が必要
#define REGEXPF_BLANK B7 // 空欄だった
#define REGEXPF_STRING B8 // 's' パス区切りを意識しない(ディレクトリも対象に)
#define REGEXPF_ERROR B31 // エラー有り
/* ワイルドカードの解析を行う 戻り値:REGEXPF_ flags */
extern PPXDLL DWORD PPXAPI MakeFN_REGEXP(FN_REGEXP *fn, const TCHAR *src);
extern PPXDLL int PPXAPI FilenameRegularExpression(const TCHAR *src, FN_REGEXP *fn);
extern PPXDLL int PPXAPI FinddataRegularExpression(const WIN32_FIND_DATA *ff, FN_REGEXP *fn);
extern PPXDLL void PPXAPI FreeFN_REGEXP(FN_REGEXP *fn);

// mode の設定内容
#define ISEA_STRING		B0	// パス全体を検索(!B0:ファイル名のみ検索)
#define ISEA_ROMA		B1	// migemoでローマ字検索
#define ISEA_FLOAT		B16	// 部分一致検索が有効
#define ISEA_KEEPHL		B23	// 終了時にハイライトを残す／付ける(PPc[J]/\[J]内部使用)
#define ISEA_FNAME		B24	// ファイル名を検索(PPc[J]/\[J]内部使用)
#define ISEA_COMMENT	B25	// コメントを検索(PPc[J]/\[J]内部使用)
#define ISEA_FNAME_EXT	B26	// ファイル名時、拡張子も検索(PPc[J]/\[J]内部使用)
#define ISEA_COMMENTEX	B27	// 拡張コメントを検索(PPc[J]/\[J]内部使用)
#define ISEA_COLUMNS	B28	// カラム拡張を検索(PPc[J]/\[J]内部使用)
extern PPXDLL BOOL PPXAPI SearchRomaString(const TCHAR *text, const TCHAR *searchstr, DWORD mode, DWORD_PTR *handle);

//================================================================ スレッド管理
#define XTHREAD_ROOT		B0	// プロセスを起動したときに作成されたスレッド
#define XTHREAD_EXITENABLE	B1	// プロセス終了時に ExitThread を実行してもよい
#define XTHREAD_RESTARTREQUEST	B2	// 異常終了したときに、スレッドの再作成が可
#define XTHREAD_TERMENABLE	B3	// スレッドに異常が起きたときスレッド強制終了可
#define XTHREAD_REPORTTHREAD B6	// PPxUnRegisterThread 時に、
								// PPXCMDID_REQUIRE_CLOSETHREAD で要求された
								// PPXMEVENT_CLOSETHREAD を module へ送信
#define XTHREAD_CHECKQUIT	B7	// (PPc)終了できるかをメインループに通知

typedef struct TagTHREADSTRUCT {
	const TCHAR *ThreadName;		// スレッド名(要設定)
	DWORD flag;						// スレッドの属性(要設定)

	struct TagTHREADSTRUCT *next;	// 次のスレッド(設定不要)
	DWORD ThreadID;					// スレッドハンドル(設定不要)
	int PPxNo;						// 登録番号 PPxRegist が設定する
} THREADSTRUCT;

// XTHREAD_RESTARTREQUEST 使用時用
typedef struct {
	THREADSTRUCT threadinfo; // flag に XTHREAD_RESTARTREQUEST を指定する
	HWND hParentWnd; // 異常終了WM_PPXCOMMANDを通知する窓。(要設定)
	WPARAM wParam; // WM_PPXCOMMAND の wParam, lParam に格納される値(option)
	LPARAM lParam;
} RESTARTTHREADSTRUCT;

extern PPXDLL BOOL PPXAPI PPxRegisterThread(THREADSTRUCT *procstruct);
extern PPXDLL BOOL PPXAPI PPxUnRegisterThread(void);
extern PPXDLL BOOL PPXAPI PPxWaitExitThread(void);
extern PPXDLL HWND PPXAPI GetPPxhWndFromID(PPXAPPINFO *ppxa, const TCHAR **src, TCHAR *path);

#define XTHREADINFO_GETSTRUCT	0
#define XTHREADINFO_GETNAME		1
#define XTHREADINFO_GETFLAGS	64
#define XTHREADINFO_SETFLAGS	96
#define XTHREADINFO_SETFLAGS_CHECKQUIT	(XTHREADINFO_SETFLAGS + 7)
extern PPXDLL void * PPXAPI PPxThreadInfo(DWORD id, DWORD mode);

//=================================================================== Task 管理
// mode の選択肢
#define PPXREGIST_NORMAL	0	// A-Z 迄を自動割当、Z を超える時は Z
#define PPXREGIST_IDASSIGN	1	// A-Z のいずれかに限定、既に有るならエラー
#define PPXREGIST_MAX		2	// 指定の範囲内に登録できないならエラー

#define PPXREGIST_COMBO_FIRST		8
#define PPXREGIST_COMBO_IDASSIGN	9	// Combo 登録
#define PPXREGIST_COMBO_FREE		10	// Combo 解除

#define PPXREGIST_FREE		-1	// 登録解除
#define PPXREGIST_SETHWND	-2	// 指定RegID に対応する HWND を登録する
#define PPXREGIST_SETPPcFOCUS	-101	// 指定RegID, HWND を最終PPcフォーカスとして登録する

#define PPXREGIST_SEARCH_LIVEID	-100 // Reuse に使えるIDを探す
#define PPXREGIST_SEARCH_LIVEID_MAX -10

#define PPXREGIST_DUMMYHWND ((HWND)(LONG_PTR)-2) // hWnd が未確定のときに使用する
extern PPXDLL int PPXAPI PPxRegist(HWND hWnd, TCHAR *ID, int mode);
extern PPXDLL void PPXAPI PPxPostMessage(UINT uMsg, WPARAM wParam, LPARAM lParam);
extern PPXDLL void PPXAPI PPxSendMessage(UINT uMsg, WPARAM wParam, LPARAM lParam);
extern PPXDLL void PPXAPI PPxSetPath(int RegNo, TCHAR *path);
extern PPXDLL BOOL PPXAPI BusyPPx(HWND hWnd, HWND hUseWnd);	// 特定のPPxを占有する
extern PPXDLL HWND PPXAPI PPxGetHWND(TCHAR *id);

enum {
	JOBSTATE_UNDEFINED = 0,
	JOBSTATE_ATTRIBUTES,
	JOBSTATE_COMPARE,
	JOBSTATE_COUNTSIZE,
	JOBSTATE_RENAME,
	JOBSTATE_THREAD,
	JOBSTATE_ARC_PACK,
	JOBSTATE_ARC_EXTRACT,
	JOBSTATE_FOP_MOVE,
	JOBSTATE_FOP_COPY,
	JOBSTATE_FOP_MIRROR,
	JOBSTATE_FOP_SHORTCUT,
	JOBSTATE_FOP_LINK,
	JOBSTATE_FOP_DELETE,
	JOBSTATE_FOP_UNDO,
	JOBSTATE_FOP_SYMLINK,
	JOBSTATE_NORMAL,	// (cmd)
	JOBSTATE_COMMENT,
	JOBSTATE_MAX,
};
//==================================================================== 情報提示
// mode
#define AJI_SHOW 0
#define AJI_COMPLETE 1
#define AJI_SELECT 2
#define AJI_START 3
#define AJI_NOMODE 4
extern PPXDLL void PPXAPI ActionInfo(HWND hWnd, PPXAPPINFO *info, DWORD mode, const TCHAR *name);

/*-------------------------------------
	言語切替処理
	"ID|english text" で呼び出す
-------------------------------------*/
extern PPXDLL const TCHAR * PPXAPI MessageText(const TCHAR *text);
extern PPXDLL void PPXAPI MessageTextDialog(HWND hDlg);
extern PPXDLL void PPXAPI LocalizeDialogText(HWND hDlg, DWORD titleID);

/*-------------------------------------
	Beep を発声する

type:発声種別
	0:致命的エラー
	1:一般エラー
	3:重要な
	4:重要でない
	5:結果を表示する情報
	6:好意で出力する情報/作業完了
	7:その他
-------------------------------------*/
#define XB_NiERR	2	// 重要でないエラー
extern PPXDLL void PPXAPI XBeep(UINT type);

/*------------------------------------- 各種メッセージボックス
	情報提供を行う

hWnd, title:	ダイアログを表示する時の親ウィンドウ, タイトル

type:種別 ※ X_log & (1 << type) が true のとき、表示・ロギングする
	0-7:ダイアログ表示+ログ  8-15:ログのみ
		1,  9:一般エラー表示
		2, 10:重要でないエラー表示
		3, 11:重要な警告
		4, 12:重要でない警告
		5, 13:結果を表示する情報
		6, 14:好意で出力する情報/作業完了
		7, 15:その他
	16-23:エラー関係(ログ無し)
	24-31:デバック用(ログ有り)

message:出力文字列, % 対応
-------------------------------------*/
// ダイアログ表示+ログ
#define XM_FaERRld	0	// 致命的エラー
#define XM_GrERRld	1	// 一般エラー
#define XM_NiERRld	2	// 重要でないエラー
#define XM_ImWRNld	3	// 重要な警告
#define XM_NiWRNld	4	// 重要でない警告
#define XM_RESULTld	5	// 結果を表示する情報
#define XM_INFOld	6	// 好意で出力する情報/作業完了

// ログ
#define XM_FaERRl	8	// 致命的エラー
#define XM_INFOLLog	13	// 好意で出力する情報/作業完了(ログのみ, 通常 off)
#define XM_INFOlog	14	// 好意で出力する情報/作業完了(ログ窓＋ログ, 通常 off)
#define XM_ETCl		15	// その他(SetPopMsg, 通常 off)

// ダイアログ表示
#define XM_FaERRd	17	// 作業中止に至るエラー
#define XM_NsERRd	18	// 中止に至らないエラー

// コンソール出力
#define XM_MesCon	22	// StrErr に出力(現在使用例無し)
#define XM_MesConWait	23	// StrErr に出力(待機あり,現在使用例無し)

// デバッグ用
#define XM_DbgDIA	24	// ダイアログ表示
#define XM_BEEP		25	// Beep(現在使用例無し)
#define XM_DbgLOG	26	// ログのみ
#define XM_DUMP		27	// Dumpダイアログ message:ダンプ先頭 1st:DWORD, dump size
#define XM_DUMPLOG	28	// Dumpログのみ
#define XM_DbgCon	29	// StrErr に出力
#define XM_DbgConWait	30	// StrErr に出力(待機あり,現在使用例無し)

extern PPXDLL void USECDECL XMessage(HWND hWnd, const TCHAR *title, UINT type, const TCHAR *message, ...);
#define xmessage(type, message) XMessage(NULL, NULL, type, message)
#define MES_AbortCheck NULL // text に指定し、style が MB_QYES なら表示
#define MB_PPX_PLAIN_TEXT B24 // text の言語対応をなしにする
#define MB_PPX_PLAIN_TITLE B25 // title の言語対応をなしにする
#define MB_PPX_BUTTON_TEXT B26 // title に続いてボタンのテキストが列挙
extern PPXDLL int PPXAPI PMessageBox(HWND hWnd, const TCHAR *text, const TCHAR *title, UINT style);
//================================================================== エラー関係
#define PPERROR_GETLASTERROR 0
extern PPXDLL ERRORCODE PPXAPI PPErrorMsg(TCHAR *str, ERRORCODE code); //len:VFPS
extern PPXDLL ERRORCODE PPXAPI PPErrorBox(HWND hWnd, const TCHAR *title, ERRORCODE code);

//================================================================== ツールバー
extern PPXDLL HWND PPXAPI CreateToolBar(ThSTRUCT *thCmd, HWND hParentWnd, UINT *ID, const TCHAR *custname, const TCHAR *currentpath, DWORD style);
#ifdef _INC_COMMCTRL
extern PPXDLL void PPXAPI LoadToolBarBitmap(HWND hWnd, const TCHAR *bmpname, TBADDBITMAP *tbab, SIZE *bmpsize);
#endif
extern PPXDLL BOOL PPXAPI SetToolBarTipText(HWND hToolbarWnd, ThSTRUCT *thCmd, NMHDR *nmh);
extern PPXDLL TCHAR * PPXAPI GetToolBarCmd(HWND hToolbarWnd, ThSTRUCT *thCmd, UINT ID);
#define CloseToolBar(hWnd) DestroyWindow(hWnd)

#define ICONTYPE_TOOLBAR 0 // ツールバー用(絵文字フォント、背景ダイアログ)
#define ICONTYPE_FILEICON 1 // ファイルアイコン用(背景C_mes)
#define ICONTYPE_TOOLBAR_O 2 // ツールバー用(非透過)
extern PPXDLL HBITMAP PPXAPI CreateScriptBitmap(const TCHAR **source, UINT iconsize, UINT iconwidth, int icontype);

// DrawEdge 相当
extern PPXDLL void PPXAPI DrawSeparateLine(HDC hDC, const RECT *box, UINT flags);

//========================================================= 各種PPxのための処理
extern PPXDLL ERRORCODE PPXAPI PPxCommonCommand(HWND hWnd, LPARAM lParam, WORD key);
extern PPXDLL LRESULT PPXAPI PPxCommonExtCommand(WORD key, WPARAM wParam);

// X_win の設定内容
#define XWIN_MENUBAR	B0	// メニューバーを表示
#define XWIN_HIDETASK	B1	// タスクバーに表示させない
#define XWIN_HIDESCROLL	B2	// スクロールバーを表示させない
#define XWIN_SWAPSCROLL	B3	// スクロールバーの位置を交換。PPc専用
#define XWIN_TOOLBAR	B4	// ツールバーを表示
#define XWIN_NOSTATUS	B5	// ステータス行を非表示
#define XWIN_NOINFO		B6	// 情報行を非表示。PPc専用
#define XWIN_HEADER		B7	// ヘッダ表示。PPc専用
#define XWIN_NOTITLE	B8	// タイトルバー非表示。
//-------------------------------------------------------------------- PPB 関連
#define RECEIVESTRINGSLENGTH (CMDLINESIZE + 64)
extern PPXDLL int PPXAPI ReceiveStrings(const TCHAR *ID, TCHAR *str);

//-------------------------------------------------------------------- PPC 関連
// 指定するPPc窓の取得
#define CGETW_NEXTNOPREV	0	// 次(方向記憶なし), 要 RegNo
#define CGETW_NEXT	1	// 次(CGETW_PREVに使う方向も記憶), 要 RegNo
#define CGETW_BACK	-1	// 前(CGETW_PREVに使う方向も記憶), 要 RegNo
#define CGETW_PAIR	2	// 反対窓, 要 RegNo
#define CGETW_PREV	3	// 直前, 要 RegNo
#define CGETW_SAVEFOCUS	-2	// Focus 保存, 要 RegNo
#define CGETW_GETFOCUS	-3	// Focus 取得, RegNo 不要
#define CGETW_GETFOCUSPAIR	-4	// Focus 窓の反対窓取得, RegNo 不要
#define CGETW_GETCOMBOHWND	8	// 指定ID(0-25)のCombo hWnd取得
// 'A'-'Z':直接指定
extern PPXDLL HWND PPXAPI PPcGetWindow(int RegNo, int direction);
extern PPXDLL void PPXAPI PPcOldPath(INT_PTR RegNo, DWORD join, TCHAR *str);
extern PPXDLL HWND PPXAPI PPxCombo(HWND hWnd);

//-------------------------------------------------------------------- PPV 関連
// B16-B23 はPPv内部で使用
#define PPV_DISABLEBOOT		B0	// PPvが無ければそのまま(指定なければ起動)
#define PPV_NOFOREGROUND	B1	// PPvを前面にしない
#define PPV_NOENSURE		B2	// 必ず表示しなくてもよい(連動表示向け設定)
#define PPV_PARAM			B3	// オプション指定有り
#define PPV_BOOTID			B4	// B24～B31に使用するPPvのID(ASCII)あり
#define PPV_HEADVIEW		B5	// 途中まで読んで表示(PPV_NOENSUREと併用時)
#define PPV_SYNCVIEW		B6	// 連動表示設定
#define PPV_SETPARENTWND	B7	// hWnd を PPv に親として登録(-parent)
extern PPXDLL BOOL PPXAPI PPxView(HWND hWnd, const TCHAR *param, int flags);
extern PPXDLL int PPXAPI PPvViewName(TCHAR *param);

// 文字コード定義
enum {
	VTYPE_IBM = 0, // CP 437
	VTYPE_ANSI, // CP 1252 CP_LATIN1
	VTYPE_JIS,  // CP 50220
	VTYPE_SYSTEMCP, // CP_ACP 旧 VTYPE_SJIS
	VTYPE_EUCJP, // CP 51932 旧 VTYPE_EUC
	VTYPE_UNICODE, // CP 1200 UCS-2LE
	VTYPE_UNICODEB, // CP 1200? UCS-2BE
	VTYPE_SJISNEC, // CP 932?
	VTYPE_SJISB, // CP 932?
	VTYPE_KANA, // CP 932?
	VTYPE_UTF8, // CP 65001
	VTYPE_UTF7, // CP 65000
	VTYPE_RTF, // CP_ACP, PPv用
	VTYPE_OTHER, // CP_ACP, PPv用
	VTYPE_MAX,		// text code の種類の数
	// VTYPE_MAX 以上は codepage を使用(codepage の最小値は37)
};
#define VTypeToCPlist_max (VTYPE_UTF7 + 1)

enum {
	VTYPE_CRLF = 0,
	VTYPE_CR,
	VTYPE_LF
};
//----------------------------------------------------------------- PPCUST 関連
enum {
	// PPcustCStore 用設定
	PPXCUSTMODE_STORE = 0, // 保存
	PPXCUSTMODE_APPEND, // 全て上書き
	PPXCUSTMODE_UPDATE, // 必要に応じて追加・上書き

	PPXCUSTMODE_LANG_FILE = B16, // 言語ファイル。MesXXXXがあれば、X_LANGに反映
	PPXCUSTMODE_JPN_RESOURCE = B17, // DefCust専用、組込日本語リソース読込用

	// PPcust 用設定
	PPXCUSTMODE_DUMP_ALL = 0xffff, // 全表示項目を出力
	PPXCUSTMODE_DUMP_PART = 0xfffe, // ワイルドカード指定
	PPXCUSTMODE_DUMP_PART_NOBOM = 0xfffd, // log に BOM を出力しない

	PPXCUSTMODE_DUMP_TITLE = B16, // タイトルコメントだけ出力
	PPXCUSTMODE_DUMP_DISCOVER = B17, // 非表示項目も表示
	PPXCUSTMODE_DUMP_NO_COMMENT = B18, // コメントを無くす
	PPXCUSTMODE_DUMP_NO_SORT = B19 // ソートを省略して読み書きの高速化
};
#define PPCUSTRESULT_RELOAD B1
#define PPCUSTRESULT_RESTART B2
extern PPXDLL int PPXAPI PPcustCStore(TCHAR *mem, TCHAR *memmax, int appendmode, TCHAR **log, int (USECDECL *Sure)(const TCHAR *message));
extern PPXDLL TCHAR * PPXAPI PPcustCDump(void);
extern PPXDLL TCHAR * PPXAPI PPcust(int mode, const TCHAR *param);

/*-----------------------------------------------------------------------------
	内蔵テキストエディタ PPeui

	hPWnd		親となる HWND
	title & text	1) text != NULL
						title:ウィンドウタイトル text:編集する文字列
					2) text == NULL
						title:開くファイル名 text:NULL

	戻り値		開いた PPeui の HWND
-----------------------------------------------------------------------------*/
extern PPXDLL HWND PPXAPI PPEui(HWND hPWnd, const TCHAR *title, const TCHAR *text);

/*-----------------------------------------------------------------------------
	PPx Module 関連

・Module の要求仕様
	マルチスレッド・マルチプロセスを考慮する。
	ModuleEntry を実装すること
-----------------------------------------------------------------------------*/
#define PPXM_INFORMATION	0	// どの機能を提供できるかをPPXMINFOSTRUCTで通知
#define PPXMEVENT_CLEANUP	1	// 読み込んだModuleを解放するときに送られる
								//                (ppxa使用不可 , パラメータ無)
#define PPXMEVENT_CLOSETHREAD 2	// スレッドを終了するときに送られる。
								// ただし、PPXCMDID_REQUIRE_CLOSETHREAD で
								// 要求したときのみ
#define PPXMEVENT_KEYHOOK	5	// キー入力時に送られる(登録をしたときのみ)
								//							(PPXMKEYHOOKSTRUCT)
#define PPXMEVENT_COMMAND	6	// 未定義コマンドの解釈     (PPXMCOMMANDSTRUCT)
#define PPXMEVENT_FUNCTION	7	// 未定義関数の解釈         (PPXMCOMMANDSTRUCT)
#define PPXMEVENT_FILEDRAW	8	// PPcファイル情報の表示   (PPXMFILEDRAWSTRUCT)
#define PPXMEVENT_FILEDRAWINFO	9	// PPXMEVENT_FILEDRAW の情報送信を要求
		// 情報送信は、PPXCMDID_REPORTPPCCOLUMN で名前を通知する (パラメータ無)
#define PPXMEVENT_SEARCH	10 // ファイル等の検索をする   (PPXMSEARCHSTRUCT)
		// 検索結果は、PPXCMDID_REPORTSEARCH で通知する
//#define PPXMEVENT_CREATE		// アプリケーション作成時に送られる(未実装)
//#define PPXMEVENT_DESTROY		// アプリケーション終了時に送られる(未実装)
//#define PPXMEVENT_MOUSE		// マウスクリック時に送られる(未実装)
//#define PPXMEVENT_CURSOR		// カーソル位置変更時に送られる(未実装)

#define PPMTYPEFLAGS(ppm) (1 << (ppm))
typedef struct tagPPXMINFOSTRUCT {
	DWORD infotype;		// 必要な情報の種類 0:モジュール情報(typeflags, copyright)
	DWORD typeflags;	// 有効な PPMEVENT
						// PPMTYPEFLAGS(PPXMEVENT_FUNCTION) | … とする
	WCHAR *copyright;	// 著作権。MAXPATH以内。「モジュール名 Copyright (c)2010 TORO」などとする
} PPXMINFOSTRUCT;

// PPXMEVENT_COMMAND/PPXMEVENT_FUNCTION コマンドモジュール/関数モジュール
typedef struct tagPPXMCOMMANDSTRUCT {
#ifndef _WIN64
	const WCHAR *commandname;	// コマンドの名前、大文字化済み("COMMAND")
	const WCHAR *param;			// パラメータ
	DWORD paramcount;			// パラメータの数
	DWORD commandhash;			// *commandhash で得られる値
	WCHAR *resultstring;		// PPXMEVENT_FUNCTIONの戻り値。
								//  文字数CMDLINESIZEの大きさを確保済み
								//  PPXMEVENT_COMMAND では使用されない(NULL)
#else
	const WCHAR *commandname;	// コマンドの名前、大文字化済み("COMMAND")
	WCHAR *resultstring;		// PPXMEVENT_FUNCTIONの戻り値。
	const WCHAR *param;			// パラメータ
	DWORD paramcount;			// パラメータの数
	DWORD commandhash;			// *commandhash で得られる値
#endif
} PPXMCOMMANDSTRUCT;

// PPXMEVENT_FILEDRAW ファイル描画モジュール
typedef struct tagPPXMFILEDRAWSTRUCT {
	const WCHAR *modulename;	// コマンドの名前、大文字化済み、6文字以内
	DWORD commandhash;	// *commandhash で得られる値
	DWORD LoadCounter;	// ディレクトリを変更する度に1増加するカウンタ
						// ディレクトリ移動を検出するために使う。
	RECT DrawArea;		// 表示領域。背景表示していないときは、fontsize.y 未満
						// の範囲をhBackBrushで消去する必要がある
	HDC hDC;			// font 設定済みの HDC
	SIZE fontsize;		// 使用フォントの平均幅と高さ
	HBRUSH hBackBrush;	// 背景用ブラシ
	const WIN32_FIND_DATAW *finddata; // 表示する内容
	BOOL IsCursor;		// カーソル上のエントリなら !FALSE
	DWORD EntryIndex; // PPXCMDID_ENTRY_HS_GETINDEX
} PPXMFILEDRAWSTRUCT;

// PPXMEVENT_KEYHOOK - キー入力拡張モジュール
typedef struct tagPPXMKEYHOOKSTRUCT {
	WORD key;	// キーコード
} PPXMKEYHOOKSTRUCT;

// PPXMEVENT_SEARCH - 検索モジュール
#define PPXH_SEARCH_NAMEONLY B16 // PPXMEVENT_SEARCH 用追加指定。ファイル名情報のみを要求する(一行編集の補完処理用)
#define PPXMSEARCH_SHORTRESULT 100 // 短時間で結果を返すときの maxresults
typedef struct tagPPXMSEARCHSTRUCT {
	const WCHAR *keyword; // 検索文字列
	DWORD searchtype; // 検索種別。ヒストリの PPXH_ の値が格納される。
	DWORD maxresults; // 最大検索件数。主に自動補完のために検索時間を
		// 短縮するために用いるので、MAX32 でないとき(PPXMSEARCH_SHORTRESULT)は
		// 短時間(0.1秒程度)で結果を返すようにする
} PPXMSEARCHSTRUCT;

// ModuleEntry の戻り値
#define PPXMRESULT_SKIP 0	// 処理しなかった…次のモジュールを調べても良い
#define PPXMRESULT_DONE 1	// 処理した
#define PPXMRESULT_STOP -1	// 処理した、後続コマンドがあれば中止する

// PPx Module 側用の定義
#ifndef VC_DLL_EXPORTS
extern int PPXAPI ModuleEntry(PPXAPPINFOW *ppxa, DWORD cmdID, PPXMODULEPARAM pxp);
#endif

typedef int (PPXAPI *CALLBACKMODULEENTRY)(PPXAPPINFOW *ppxa, DWORD cmdID, PPXMODULEPARAM pxp);

// PPx Module を呼び出す
extern PPXDLL int PPXAPI CallModule(PPXAPPINFO *info, DWORD func, PPXMODULEPARAM ModuleParam, CALLBACKMODULEENTRY CallBackModule);

// Winelib 限定
#if NODLL
	extern void InitCommonDll(HINSTANCE hInst);
#endif
#ifdef __cplusplus
}
#endif
