import sys
import os
import datetime
import keyhac_ini

import pyauto
from keyhac import *

################################################################################
# 追加したい機能のメモ
#-------------------------------------------------------------------------------
# ・line単位でのコピー切り取り後の貼り付け処理が不十分              [line-ctrl]
# ・リーピート回数のMAX値を決める、コマンドから可変にする。         [rpt-max]
# ・左右のCtrlとShiftをうまく使い分けてC-SpaceやS-Spaceを復活させる [ctrl-sft-spc]
# ・IMEがONの時にテンキが入力出来ない時がある                       [ime-tenkey]
# ・クリックやダブルクリックを実現させる                            [mouse-clk]
# ・マウスポインタの位置を記録・再現させる。                        [mouse-clk]
# ・Shift-oによる改行-insertmodeを実装する。                        [main]
# ・エクセルの並べ替えのコマンドを追加する。                        [main]
# ・word用のテンキーマクロも作成する                                [tenkey-mcr]
# ・^M^L(C-v,C-m,C-l)で改ページの挿入                               [main]
# ・SkkIme以外のImeでもC-Jでime-on,lでime-offにする。               [like-skk]
# ・ノーマルモード(vim_mode)でのスペースの入力を阻止する。          [main]
# ・テンキー横の↓→キーをモディファイアにしてテンキーを拡張する    [tenkey-ctrl]
################################################################################

# 日時をペーストする機能
def dateAndTime(fmt):
    def _dateAndTime():
        return datetime.datetime.now().strftime(fmt)
    return _dateAndTime

## 処理時間計測のデコレータ
def profile(func):
    import functools

    @functools.wraps(func)
    def _profile(*args, **kw):
        import time
        timer = time.clock
        t0 = timer()
        ret = func(*args, **kw)
        if 1000*(timer()-t0) >100:
            print( '%s: %.3f [ms] elapsed %s' % (func.__name__, 1000 * (timer() - t0),dateAndTime("%H:%M:%S")()))
        return ret
    return _profile

## JobQueue/JobItem でサブスレッド処理にするデコレータ
def job_queue(func):
    import functools

    @functools.wraps(func)
    def _job_queue(*args, **kw):

        num_items = JobQueue.defaultQueue().numItems()
        if num_items:   # 処理待ちアイテムがある場合は、その数を表示
            print( "JobQueue.defaultQueue().numItems() :", num_items)

        def __job_queue_1(job_item):
            return func(*args, **kw)

#        def __job_queue_2(job_item):
#            # print( "job_queue : ", func.__name__, args, kw)
#            pass

#        job_item = JobItem(__job_queue_1, __job_queue_2)
        job_item = JobItem(__job_queue_1)
        JobQueue.defaultQueue().enqueue(job_item)

    return _job_queue



def configure(keymap):

    import pyauto
    #----------------------------------------------------------------------
    # keyhac起動時にSetCaretColorを起動させる。
    
    #SetCaretColorの起動状態の確認
    
    def isSetCaret(wnd):
        if (wnd.getProcessName() in ("SetCaretColor.exe")):      #SetCaretColor
            return True
        return False

    def isSetCaretOn():
        root = pyauto.Window.getDesktop()
        wnd = root.getFirstChild()
        while wnd:
            if isSetCaret(wnd):
                return True
            wnd = wnd.getNext()
        return False

    #SetCaretColorの起動
    def OnSetCaret():
        if not isSetCaretOn():
            executeFunc = keymap.command_ShellExecute( None, "..\SetCaretColor003\SetCaretColor.exe", "", "" )
            executeFunc()
    
    OnSetCaret()

    # --------------------------------------------------------------------
    # config.py編集用のテキストエディタの設定

    # プログラムのファイルパスを設定 (単純な使用方法)
    if 0:
        keymap.editor = "..\\portvim\\gvim.bat"

    # 呼び出し可能オブジェクトを設定 (高度な使用方法)
    if 1:
        @profile
        def editor(path):
            shellExecute( None, "..\\portvim\\gvim.exe", '--remote-silent "%s"'% path, "" )
        keymap.editor = editor

    # --------------------------------------------------------------------

    # キーの単純な置き換え
    keymap.replaceKey( "LWin", 235 )
#    keymap.replaceKey( "LShift", "LAlt" )
#    keymap.replaceKey( "RShift", "RAlt" )
#    keymap.replaceKey( "(240)", 255 )
#    keymap.replaceKey( "(242)", 255 )
#    keymap.replaceKey( "Apps", "RAlt" )

    # ユーザモディファイアキーの定義
    keymap.defineModifier( 235, "User0" )
#    keymap.defineModifier( 255, "User1" )

    # どのウインドウにフォーカスがあっても効くキーマップ
    if 1:
        keymap_global = keymap.defineWindowKeymap()

        # SandS
        keymap.replaceKey("Space", "RShift")
        # 親指コントロール
        keymap.replaceKey("(29)","LCtrl")
        keymap.replaceKey("(28)","RCtrl")

        keymap_global["O-RShift"] = "Space"

        # USER0-↑↓←→ : 10pixel単位のウインドウの移動
        keymap_global[ "U0-H" ] = keymap.command_MoveWindow( -20, 0 )
        keymap_global[ "U0-L" ] = keymap.command_MoveWindow( +20, 0 )
        keymap_global[ "U0-K" ] = keymap.command_MoveWindow( 0, -20 )
        keymap_global[ "U0-J" ] = keymap.command_MoveWindow( 0, +20 )

        # USER0-Shift-↑↓←→ : 1pixel単位のウインドウの移動
        keymap_global[ "U0-S-H" ] = keymap.command_MoveWindow( -1, 0 )
        keymap_global[ "U0-S-L" ] = keymap.command_MoveWindow( +1, 0 )
        keymap_global[ "U0-S-K" ] = keymap.command_MoveWindow( 0, -1 )
        keymap_global[ "U0-S-J" ] = keymap.command_MoveWindow( 0, +1 )

        # USER0-Ctrl-↑↓←→ : 画面の端まで移動
        keymap_global[ "U0-C-H" ] = keymap.command_MoveWindow_MonitorEdge(0)
        keymap_global[ "U0-C-L" ] = keymap.command_MoveWindow_MonitorEdge(2)
        keymap_global[ "U0-C-K" ] = keymap.command_MoveWindow_MonitorEdge(1)
        keymap_global[ "U0-C-J" ] = keymap.command_MoveWindow_MonitorEdge(3)


        keymap_global["U0-X"]        = keymap.defineMultiStrokeKeymap("Win")
        keymap_global["U0-E"]        = keymap.command_EditConfig
#        keymap_global["U0-R"]        = keymap.command_ReloadConfig

        # クリップボード履歴
        keymap_global[ "C-S-Z"   ] = keymap.command_ClipboardList     # リスト表示
        keymap_global[ "C-F7"    ] = keymap.command_ClipboardList     # リスト表示
        keymap_global[ "C-S-X"   ] = keymap.command_ClipboardRotate   # 直近の履歴を末尾に回す
        keymap_global[ "C-S-A-X" ] = keymap.command_ClipboardRemove   # 直近の履歴を削除
        keymap.quote_mark = "> "                                      # 引用貼り付け時の記号

        # キーボードマクロ
        keymap_global[ "U0-0" ] = keymap.command_RecordToggle
#        keymap_global[ "U0-R" ] = keymap.command_RecordToggle
        keymap_global[ "U0-1" ] = keymap.command_RecordStart
        keymap_global[ "U0-2" ] = keymap.command_RecordStop
        keymap_global[ "U0-3" ] = keymap.command_RecordPlay
#        keymap_global[ "U0-V" ] = keymap.command_RecordPlay
        keymap_global[ "U0-4" ] = keymap.command_RecordClear



        keymap_global[ "U0-F"  ] = "C-F"
        keymap_global[ "U0-P"  ] = "C-P"

        keymap_global[ "NumLock"  ] = "e","NumLock","NumLock"
#        keymap_global[ "C-NumLock"  ] = "NumLock"

# USER0-Rでconfig.pyのリロード
#  ウィンドウを使用せずみリロードするとリロードを実行したアプリでフック出来なくなるため
    if 1:
        @profile
        def command_PopApplicationList2():

            # すでにリストが開いていたら閉じるだけ
            if keymap.isListWindowOpened():
                keymap.cancelListWindow()
                return

            def popApplicationList2():

                test_items = [
                    ( "設定の再読込",  keymap.command_ReloadConfig )
                ]

                listers = [
                    ( "",     cblister_FixedPhrase(test_items) ),
                ]

                item, mod = keymap.popListWindow(listers)

                if item:
                    item[1]()

            # キーフックの中で時間のかかる処理を実行できないので、delayedCall() をつかって遅延実行する
            keymap.delayedCall( popApplicationList2, 0 )

        keymap_global[ "U0-R" ] = command_PopApplicationList2

# 操作するアプリケーションの登録

    if 1:


        @profile
        def isConsoleWindow(wnd):
            if wnd.getClassName() in (
#                                    "Vim",                                        #Vim
                                    "CfilerWindowClass",                            #内骨格
                                    "js:TARO10",                                    #一太郎
                                    "Afx:00400000:8:00010011:00000000:00010783",    #リキュール
                                    "TF8PPFPreviewForm",                            #Forum8のプレビューウィンドウ
                                    "TFormMemberResultViewer",                      #Forum8の解析結果ウィンドウ
                                    "SunAwtFrame",                                  #Android Studio
                                    "TFormOutlineElementWizard",                    #Forum8のアウトライン入力ウィンドウ
                                    "TFormOutNew",                                  #Forum8断面算定のプレビューウィンドウ
                                    "SmartVision Main Frame Window",                #SmartVision
                                    "TINP_Hotai_Kui_Form",                          #Forum8杭基礎のデータ入力ウィンドウ
                                    "TFormF3DModelEditor",                          #Forum8フレーム3Dのデータ入力ウィンドウ
                                    "TFormAOCollectionEditor",                      #Forum8フレーム3Dのデータ入力ウィンドウ
                                    "TFormSectionElementWizard",                    #Forum8フレーム3Dの断面編集ウィンドウ
                                    "TFormSectionEditor",                           #Forum8フレーム3Dの断面編集ウィンドウ
                                    "AcrobatSDIWindow",                             #Adobe Reader
                                    "ThunderRT6FormDC",                             #座標点プロット Mk_Protxy
                                    "Afx:00400000:b:00010003:00000006:024E0449",    #mp3tag
                                    "Afx:400000:8:10011:0:180573",                  #InstallShield
                                    "Afx:00400000:b:00010003:00000006:01520717",    #SoundEngine
                                    "Afx:00400000:b:00010003:00000006:000408F9",    #SoundEngine
                                    "Afx:400000:8:10011:0:2804b5",                  #VC++6
                                    "TFrmMainForm",                                 #FRAME-Ⅰ
                                    "Afx:400000:b:10011:6:86202b3",                 #PILE-1
                                    "TSDIAppForm",                                  #PILE-2(BCB全般共通の可能性あり)
                                    "{0F04AF43-7B85-46A5-A0A7-6D323A84AD7C}",       #Asr
                                    "ThunderRT6MDIForm",                            #A9CAD
                                    "TPreViewForm",                                 #COSMOのプレビューウィンドウ
#                                    "TMdenMainForm",                                #M電卓
                                    "CabinetWClass",                                #エクスプローラ ウィンドウ
                                    "TFormMain",                                    # フォーラム断面算定：
                                    "TextEditorWindow",                             # sakuraEditor
                                    "emo.system.d7h8",                              # EMOffice
                                    "TAppBuilder",                                  #BCB6
                                    "ConsoleWindowClass",
                                    "TkTopLevel",                                   #Git GUI
                                    "TAIMPMainForm",                                #AIMP(ポータブルミュージックプレーヤー)
                                    "WMP Skin Host",                                #Windows Media Player
                                    "Outlook Express Browser Class",                #Outlook Express
                                    "ATH_Note",                                     #Outlook Expressの新規メールWindow
                                    "DocuWorksViewerLightMainWindow",               #DocuWorksViewerLight
                                    "OpusApp",                                      #WORD
                                    "XLMAIN",                                       #EXCEL
                                    "TfrmAdvancedSystemCare7_Monitor",
                                    "wndclass_desked_gsk",                          #VBA
                                    "TFormIObitSD2",
                                    "LVM1414",                                      #LightVM
                                    "CkwWindowClass"):
                return True
            return False

        #########################################################################
        #各アプリケーション識別用関数
        #########################################################################

        def isExcel(wnd):
            if wnd.getClassName().startswith("EXCEL"):
                return True
            return False

        def isVim(wnd):
            if wnd.getClassName().startswith("Vim"):
                return True
            return False

        def isAndroidStudio(wnd):
            if wnd.getClassName().startswith("SunAwtFrame"):
                return True
            return False

        def isMdentaku(wnd):
            if keymap.getTopLevelWindow().getClassName().startswith("TMdenMainForm"):
                return True
            return False

        def isAfLogForm(wnd):
            if wnd.getClassName().startswith("TLogForm"):
                return True
            return False

        def isAIMP(wnd):
            if wnd.getClassName().startswith("TAIMPMainForm"):
                return True
            return False

        def isVba(wnd):
            if wnd.getClassName().startswith("VbaWindow"):
                return True
            return False

        def isWord(wnd):
            if wnd.getClassName().startswith("_WwG"):
                return True
            return False

        def isSkk(wnd):
            if wnd.getClassName().startswith("Skk"):
                return True
            return False

        def isForumPrev(wnd):
            if wnd.getClassName().startswith("TPPFPreviewFrame"):
                return True
            return False

        def isAutoCad(wnd):
            if (wnd.getClassName().startswith("Afx:400000") and
                    wnd.getProcessName().startswith("aclt")):
                return True
            return False

        def is32770Window(wnd):
            if (wnd.getClassName().startswith("#32770") and
                    (wnd.getProcessName() in ("TwinMain.exe",
                                              "taskmgr.exe",
                                              "MoviePhotoMenu.exe",
                                              "x-APPLICATION.exe",
                                              "TWinPost.exe"))):
                return True
            return False

        def isDocuWorks(wnd):
            if wnd.getClassName().startswith("AfxFrameOrView42"):
                return True
            return False

        def isChrome(wnd):
            if wnd.getClassName().startswith("Chrome_WidgetWin_1"):
                return True
            return False

        def isAfxWindow(wnd):
            if (wnd.getClassName().startswith("Afx:400000:") and
                    (wnd.getProcessName() in ("aclt.exe",
                                              "TVIEWER.EXE",
                                              "Dwviewer.exe",
                                              "Mp3tag.exe",
                                              "walldf.exe",
                                              "dwdesk.exe"))):
                return True
            return False

        def isVC6(wnd):
            if (wnd.getClassName().startswith("Afx:400000:") and
                    (wnd.getProcessName() in ("MSDEV.EXE"))):             #VC++6
                exetxt = wnd.getText()
                if exetxt.find("Microsoft Visual C++")> -1:
                    return True
            return False

        def isTMainForm(wnd):
            if (wnd.getClassName().startswith("TMainForm") and
                    (wnd.getProcessName() in ("ClsProj.exe",      #CalsManager
                                              "Foundation9.exe",  #Forum8 杭基礎 Ver9
                                              "ezhtml.exe",       #ezHTML
                                                ))):
                return True
            return False

        @profile
        def isCraftWare(wnd):
            if wnd.getClassName() in ("CfilerWindowClass",
#                                       "ClnchWindowClass",
                                       "keyhacWindowClass"):
                return True
            return False

        @profile
        def isAf(wnd):
            if wnd.getClassName() in ("TFileBox"):
                return True
            return False

        #あふwのトップレベルウィンドウ
        @profile
        def isAfwWindow(wnd):
            if wnd.getClassName() in ("TAfxWForm"):
                return True
            return False

        @profile
        def isTButton(wnd):
            if wnd.getClassName() in ("TButton"):
                return True
            return False

        @profile
        def isApp(wnd):
            if isConsoleWindow(wnd):
                return True
            elif isChrome(wnd):
                return True
            elif isTMainForm(wnd):
                return True
            elif is32770Window(wnd):
                return True
            elif isAfxWindow(wnd):
                return True
            elif isVC6(wnd):
                return True
            elif isAfwWindow(wnd):
                return True
            return False

# Ctrl-Tabで登録済みのアプリケーションの切り替え

    if 1:
        def ForegroundWindow(wnd):
            def _fanc():
                if wnd.isMinimized():
                    wnd.restore()
#                wnd = wnd.getLastActivePopup()
                wnd.setForeground()
                wnd.setActive()
            return _fanc

        def command_SwitchApplication():

            # すでにリストが開いていたら閉じるだけ
            if keymap.isListWindowOpened():
                keymap.cancelListWindow()
                return

            def SwitchApplication():

                root = pyauto.Window.getDesktop()

                wnd = root.getFirstChild()
                applications2 = None
                while wnd:
                    if isApp(wnd) and not isAfwWindow(wnd):
                        exetxt = wnd.getText()
                        if exetxt:
                            if wnd.getClassName()=="CabinetWClass":
                                exetxt = "\\"+exetxt
                            if applications2:
                                applications2 += [(exetxt, ForegroundWindow(wnd))]
                            else:
                                applications2 = [(exetxt, ForegroundWindow(wnd))]
                    wnd = wnd.getNext()


                listers = [
                    ( "アプリケーションリスト",     cblister_FixedPhrase(applications2) )
                ]

                item, mod = keymap.popListWindow(listers)

                if item:
                    item[1]()

            # キーフックの中で時間のかかる処理を実行できないので、delayedCall() をつかって遅延実行する
            keymap.delayedCall( SwitchApplication, 0 )

        keymap_global[ "C-Tab" ] = command_SwitchApplication


        def command_NextApplication(num):

            def _fanc():

                root = pyauto.Window.getDesktop()
                cnt =0

                #M電卓とvimから呼び出された時はカウントを1進めておく
                if (isMdentaku(keymap.getWindow()) or
                        isVim(keymap.getWindow())):
                    cnt = 1

                wnd = root.getFirstChild()
                while wnd:
                    if isApp(wnd):
                        exetxt = wnd.getText()
                        if exetxt:
                            if cnt == num:
                                 #あふwは無視する
                                if not (isAfwWindow(wnd)):
                                    ForegroundWindow(wnd)()
                                    return
                                else:
                                    cnt -=1
                            cnt += 1
                    wnd = wnd.getNext()
            return _fanc

        keymap_global["LC-0"] = command_NextApplication(1)
        keymap_global["LC-9"] = command_NextApplication(2)

        def command_OutAppList():

            root = pyauto.Window.getDesktop()

            wnd = root.getFirstChild()
            while wnd:
                appname = wnd.getClassName()
                print( appname)
                wnd = wnd.getNext()

        def command_TopLevelWindow():
            root = keymap.getTopLevelWindow()
            if root:
                print( root.getClassName())
                ckit.ckit_misc.setClipboardText(root.getClassName())
                keymap.popBalloon("mode",root.getClassName(),1000)

        def command_GetProcessName():
            root = keymap.getTopLevelWindow()
            if root:
                print( root.getProcessName())
                ckit.ckit_misc.setClipboardText(root.getProcessName())

        def command_ThisClass():
            root = Window.getFocus()
            if root:
                print( root.getClassName())
                ckit.ckit_misc.setClipboardText(root.getClassName())

        def command_DeskTop():

            root = pyauto.Window.getDesktop()
            if root:
#                root.setForeground()
                root.setActive()




################################################################################
# -*- mode: python; coding: utf-8-dos -*-
##
## Windows の操作を Vimのキーバインドで行う設定（keyhac版）
##
# このスクリプトは、keyhac で動作します。
#   https://sites.google.com/site/craftware/keyhac
# スクリプトですので、使いやすいようにカスタマイズしてご利用ください。
#
# この内容は、utf-8-dos の coding-system で config.py の名前でセーブして
# 利用してください。
################################################################################


    if 1:
        def classname_is_vim(window):
            return window.getClassName() not in ("mintty",             # Mintty
#                                                 "ConsoleWindowClass", # Cmd, Cygwin
                                                 "Emacs",              # NTEmacs
#                                                 "CfilerWindowClass",  # Cfiler
#                                                 "keyhacWindowClass",  # keyhac
                                                 "Vim",                # Vim
                                                 "PuTTY",              # PuTTY
                                                 "SWT_Windows0",        # Eclipse
                                                 "ConsoleWindowClass",        #ConsoleWindowClass
                                                 "SmartVision Main Frame Window",        #スマートビジョン(デレビ)
                                                 "SvUI SmartVision Parent Window"        #スマートビジョン(デレビ)
                                                 )

        keymap_vim = keymap.defineWindowKeymap(check_func=classname_is_vim)

    ############################################################################
    # Vim_Flags
    ############################################################################

        # メインのモードフラグ
        # 0:ノーマルモード (通常のキーボード)
        # 1:VimMode (Vimのノーマルモード)
        # 2:InsertMode (Vimの挿入モード)
        # 3:VisualMode (Vimのビジュアルモード)
        # 4:CommandMode (Vimのコマンドモード)
        # 5:SearchMode (検索モード EnterでVimModeに戻る)
        keymap_vim.mainmode = 1

        # VimMode でのコマンド入力中を示すフラグ
        keymap_vim.flg_mtd = 0

        # コマンドの実行回数
        keymap_vim.repeatN = 0

        # コマンドラインのコマンド
        keymap_vim.command_str =""

        # ビジュアルモードの状態を示すフラグ
        # 1:行単位の選択
        # 2:矩形単位の選択
        keymap_vim.flg_selmode = 0

        # スクロールバインドのオンオフ
        keymap_vim.flg_scroll = 0

        # 日本語入力固定モード
        keymap_vim.flg_imemode =1

        # Cfilerでの常態を示すフラグ
        keymap_vim.flg_cf_mode=0

        # EXCELLなどで日本語入力を固定する
        keymap_vim.flg_fixinput=0

        # M電卓で計算済みかどうか
        keymap_vim.flg_Mdentaku=0

        def vim_parm_reset():
            keymap_vim.flg_mtd = 0
            keymap_vim.command_str =""
            keymap_vim.repeatN = 0
            keymap_vim.flg_cf_mode =0
            keymap_vim.flg_selmode =0

        def vim_parm_AllReset():
            vim_parm_reset
            keymap_vim.flg_mcr =0
            keymap_vim.flg_scroll=0
            keymap_vim.flg_Mdentaku=0
            shellExecute( None, "..\\clnch\\clnch.exe",'--execute=setmcr;0', "" )

    ############################################################################
    # Vim_Init 変数の初期化
    ############################################################################

        # LC-Shiftの時間計測用変数
        keymap_vim.esc_tmr = 0
        # LC-Shiftが押されたかどうかのフラグ
        keymap_vim.esc_flg = 0

        ########################################################################
        # キーボードマクロの実装
        ########################################################################
        if 1:

            # キーボードマクロの記録の状態を示すフラグ
                # 0:マクロ OFF
                # 1:マクロ記録中
                # 2:マクロ実行中
            keymap_vim.flg_mcr = 0

            # 記録中のマクロ番号
            keymap_vim.mcr_num =0

            def start_rec(num):
                keymap_vim.mcr_num = num
                keymap_vim.mcr_string[num] = None
                keymap_vim.mcr_count[num] = 0
                keymap_vim.flg_mcr = 1
                shellExecute( None, "..\\clnch\\clnch.exe",'--execute=setmcr;1', "" )
#                keymap.popBalloon("mode","キーボードマクロ記録開始",1000)

            def stop_rec():
                keymap_vim.flg_mcr = 0
                write_ini_mcr(keymap_vim.mcr_num)
                shellExecute( None, "..\\clnch\\clnch.exe",'--execute=setmcr;0', "" )
#                keymap.popBalloon("mode","キーボードマクロ記録終了",1000)

            def add_macro(ckey):
                num = keymap_vim.mcr_num
                keymap_vim.mcr_count[num] += 1
                if keymap_vim.mcr_string[num]:
                    keymap_vim.mcr_string[num] += [(ckey)]
                else:
                    keymap_vim.mcr_string[num] = [(ckey)]

            @profile
            def execute_macro(num):
                def _fanc():
                    keymap_vim.flg_mcr = 2
                    i = 0
                    if keymap_vim.mcr_string[num]:
#                        keymap.popBalloon("mode","キーボードマクロ実行",1000)
                        for i in range(keymap_vim.mcr_count[num]-1):
                            send_vimmodekey(keymap_vim.mcr_string[num][i])
                    keymap_vim.flg_mcr = 0
                return _fanc

            # キーボードマクロのiniファイルからの読み込み
            def read_ini_mcr(num):
                sect = "mcr"+str(num)+"_"
                counter = keyhac_ini.getint("GLOBAL",sect+"cnt",0)
                i = 0
                if counter:
                    for i in range(counter):
                        add_macro(keyhac_ini.get("GLOBAL",sect+str(i),"null"))

            # キーボードマクロのiniファイルへの書き込み
            def write_ini_mcr(num):
                sect = "mcr"+str(num)+"_"
                delete_ini_mcr(sect)
                keyhac_ini.setint("GLOBAL",sect+"cnt",keymap_vim.mcr_count[num])
                i = 0
                if keymap_vim.mcr_count[num]:
                    for i in range(keymap_vim.mcr_count[num]):
                        keyhac_ini.set("GLOBAL",sect+str(i),keymap_vim.mcr_string[num][i])
                keyhac_ini.write()

            # キーボードマクロのiniファイルからの削除
            def delete_ini_mcr(sect):
                counter = keyhac_ini.getint("GLOBAL",sect+"cnt",0)
                if counter:
                    for i in range(counter):
                        keyhac_ini.remove_option("GLOBAL",sect+str(i))
                    keyhac_ini.remove_option("GLOBAL",sect+"cnt")

            # キーボードマクロの初期化
            keymap_vim.mcr_string = [(None)]
            keymap_vim.mcr_count = [(0)]
            for ic in range(30):
                keymap_vim.mcr_num = ic
                keymap_vim.mcr_string += [(None)]
                keymap_vim.mcr_count += [(0)]
                read_ini_mcr(ic)

        ########################################################################
        # IMEの切替え
        ########################################################################

        @profile
        def toggle_input_method():
            # keymap.command_InputKey("A-BackQuote")()
            keymap.command_InputKey("A-(243)")()

        @profile
        def toggle_imemode():
            if keymap_vim.flg_imemode:
                keymap_vim.flg_imemode=0
#                keymap.popBalloon("mode","日本語入力固定モード OFF",1000)
            else:
                keymap_vim.flg_imemode=1
#                keymap.popBalloon("mode","日本語入力固定モード ON",1000)
            if keymap_vim.flg_mcr != 2:
                shellExecute( None, "..\\clnch\\clnch.exe",'--execute=setime;%d'% keymap_vim.flg_imemode, "" )

        ############################################################################
        # 特定の機能を制御するクラス
        ############################################################################

        @profile
        def isEnterCanselClass(wnd):
            if ((wnd.getClassName() in (
                "Edit",
                "TEdit",
                )) or
                isExcel(wnd) or
                isAf(wnd) or
                isCraftWare(wnd)):
                return True
            return False

        def isTenKeyClass(wnd):
            if ((wnd.getClassName() in (
                "TRValGrid",
                "TF8Edit",
                "Edit",
                "F3 Server 60000000",       #Mightyのグリッド？
                "MSFlexGridWndClass",       #Mightyのグリッド？
                "RichTextWndClass",         #Mightyのグリッド？
                "TInplaceEdit",             #Forum8のグリッド
                "CRvgIntrEdit",             #TRValGridのセルのクラス
                )) or
                isExcel(wnd) or
                isWord(wnd)):
                return True
            return False

        def isEditorClass(wnd):
            if ((wnd.getClassName() in (
                "EditorClient",           #サクラエディタ
                "VbaWindow",              #VBAクラス
                )) or
                isWord(wnd)):
                return True
            return False
        ############################################################################
        # 文字変換
        ############################################################################

        @profile
        def exc_char(ichar, sft=0):
            if ichar=="Slash":
                return "/"
            if ichar=="Period":
                return "."
            if ichar=="Colon":
                return ":"
            if ichar=="semicolon":
                return ";"
            if ichar=="Minus":
                return "-"
            if ichar=="Caret":
                return "^"
            if ichar=="OpenBracket":
                return "["
            if ichar=="CloseBracket":
                return "]"
            if ichar=="Yen":
                return "\\"
            if ichar=="Atmark":
                return "@"
            if ichar=="Comma":
                return ","
            if ichar=="BackSlash":
                return "\\"
            if sft:
                if ichar=="Space":
                    return " "
                if ichar=="S-semicolon":
                    return "+"
                if ichar=="S-Underscore":
                    return "_"
                if ichar=="S-4":
                    return "$"
                if ichar=="S-8":
                    return "("
                if ichar=="S-9":
                    return ")"
                if ichar=="S-3":
                    return "#"


            if ichar=="/":
                return "Slash"
            if ichar==".":
                return "Period"
            if ichar==":":
                return "Colon"
            if ichar==";":
                return "semicolon"
            if ichar=="-":
                return "Minus"
            if ichar=="^":
                return "Caret"
            if ichar=="[":
                return "OpenBracket"
            if ichar=="]":
                return "CloseBracket"
            if ichar=="\\":
                return "Yen"
            if ichar=="@":
                return "Atmark"
            if ichar==",":
                return "Comma"
            if ichar=="\\":
                return "BackSlash"
            if sft:
                if ichar==" ":
                    return "Space"
                if ichar=="+":
                    return "S-semicolon"
                if ichar=="_":
                    return "S-Underscore"
                if ichar=="$":
                    return "S-4"
                if ichar=="(":
                    return "S-8"
                if ichar==")":
                    return "S-9"
                if ichar=="#":
                    return "S-3"

            return ichar

        @profile
        def CtoNum(nchar):
            cnt =0
            for cn in '0123456789':
                if cn == nchar:
                    return cnt
                cnt += 1
            cnt = 0
            for cn in 'abcdefghijklmnopqrstuvwxyz':
                if cn == nchar:
                    return cnt
                cnt += 1
            if nchar == "Enter":
                return 27
            elif nchar == "S-Enter":
                return 28
            return -1

        ########################################################################
        # Vim コマンド
        ########################################################################
        @profile
        def show_mode():
            if 0:
                mode = "Err"
                if keymap_vim.mainmode==0:
                    mode = "Nomal Mode"
                elif keymap_vim.mainmode==1:
                    mode = "Vim Mode"
                elif keymap_vim.mainmode==2:
                    mode = "Insert Mode"
                elif keymap_vim.mainmode==3:
                    mode = "Visual Mode"
                elif keymap_vim.mainmode==4:
                    mode = "Command Mode"
                elif keymap_vim.mainmode==5:
                    mode = "Search Mode"
                keymap.popBalloon("mode",mode,1000)
            else:
                if keymap_vim.flg_mcr != 2:
                    shellExecute( None, "..\\clnch\\clnch.exe",'--execute=setmod;%d'% keymap_vim.mainmode, "" )


#        @job_queue
        @profile
        def set_imeoff():
            if keymap.getWindow().getImeStatus():
                keymap.command_InputKey("A-(243)")()

        @profile
        def set_imeon():
            if not keymap.getWindow().getImeStatus():
                keymap.command_InputKey("A-(243)")()

        # コマンドラインテスト用の切り替えフラグ
        flg_commandline = 1

        if not flg_commandline:
            def show_command(cls=0):
                itime = None
                if cls:
                    itime=1000
                commandstr=keymap_vim.command_str
                keymap.popBalloon("test",":"+commandstr,itime)

        if flg_commandline:
            def show_command(cls=0):
                commandstr=keymap_vim.command_str
                if keymap_vim.flg_mcr != 2:
                    shellExecute( None, "..\\clnch\\clnch.exe",'--execute="inpcmd;%s"'% commandstr, "" )

#                # 一度閉じて再表示
#                if keymap.isListWindowOpened():
#                    keymap.cancelListWindow()
#                    return
#
#                if cls:
#                    return
#
##                command_NextApplication(0)()
#
#                def command_show_command():
#
##                    applications2 = None
#
#                    applications2 = [(":" + keymap_vim.command_str, execute_command)]
#
#                    listers = [
#                        ( "Command Line",     cblister_FixedPhrase(applications2) )
#                    ]
#
#                    item, mod = keymap.popListWindow(listers)
#
#                    if item:
#                        item[1]()
#
#                # キーフックの中で時間のかかる処理を実行できないので、delayedCall() をつかって遅延実行する
#                keymap.delayedCall( command_show_command, 0 )
#

        @profile
        def set_nomalmode():
            keymap_vim.mainmode =0
            show_mode()

        @profile
        def set_vimmode(flg=1):
            set_imeoff()
            vim_parm_reset()
            if keymap_vim.mainmode!=1:
                keymap_vim.mainmode =1
                if flg:
                    show_mode()

        @profile
        def set_insertmode():
            if (isExcel(keymap.getWindow()) and
                not keymap.getWindow().getClassName().startswith("EXCEL6")):
                keymap.command_InputKey("F2")()
            if keymap_vim.flg_imemode:
                set_imeon()
            keymap_vim.mainmode =2
            show_mode()

        @profile
        def set_visualmode():
            keymap_vim.mainmode =3
            show_mode()

        @profile
        def set_commandmode():
            keymap_vim.mainmode =4
            keymap_vim.command_str = ""
            show_command()
            #show_mode()

        def set_fixinputmode():
            keymap_vim.flg_fixinput=1

        def reset_fixinputmode():
            keymap_vim.flg_fixinput=0

        def switch_rect_sel():
            if isExcel(keymap.getWindow()):
                keymap.command_InputKey("C-Space")()
                keymap_vim.flg_selmode =2
            elif isWord(keymap.getWindow()):
                keymap.command_InputKey("C-S-F8")()
                keymap_vim.flg_selmode =2
            else:
                keymap.command_InputKey("S-F6")()
                keymap_vim.flg_selmode =1

        @profile
        def set_searchmode():
            keymap_vim.mainmode =5
            if keymap_vim.flg_imemode:
                set_imeon()
            show_mode()

        def move_up():
            keymap.command_InputKey("Up")()

        def move_down():
            keymap.command_InputKey("Down")()

        def move_left():
            keymap.command_InputKey("Left")()

        def move_right():
            keymap.command_InputKey("Right")()

        def move_next_word():
            keymap.command_InputKey("C-Right")()

        def move_back_word():
            keymap.command_InputKey("C-Left")()

        def move_pageup():
            keymap.command_InputKey("PageUp")()

        def move_pagedown():
            keymap.command_InputKey("PageDown")()

        def move_halfpageup():
            for i in range(15):
                keymap.command_InputKey("Up")()


        def move_halfpagedown():
            for i in range(15):
                keymap.command_InputKey("Down")()

        def move_buftop():
            if(isAf(keymap.getWindow())):
                keymap.command_InputKey("C-PageUp")()
            else:
                keymap.command_InputKey("C-Home")()

        def move_bufend():
            if(isAf(keymap.getWindow())):
                keymap.command_InputKey("C-PageDown")()
            else:
                keymap.command_InputKey("C-End")()

        def move_nexttab():
            keymap.command_InputKey("C-PageDown")()

        def move_prevtab():
            keymap.command_InputKey("C-PageUp")()

        def move_line_top():
            keymap.command_InputKey("Home")()

        def move_line_end():
            keymap.command_InputKey("End")()

        def input_esc():
            import time
            timer = time.clock
            dlt = 1000*(timer()-keymap_vim.esc_tmr)
            if dlt < 300 and keymap_vim.esc_flg:
                keymap.command_InputKey("Esc")()
            else:
                set_vimmode()
            keymap_vim.esc_tmr = timer()
            keymap_vim.esc_flg = 1

        def input_tab():
            keymap.command_InputKey("Tab")()

        def input_Space():
            keymap.command_InputKey("Space")()

        def input_stab():
            keymap.command_InputKey("S-Tab")()

        def input_backspace():
            keymap.command_InputKey("Back")()

        def input_enter():
            keymap.command_InputKey("Enter")()
            if isMdentaku(keymap.getWindow()):
                keymap_vim.flg_Mdentaku =1

        def concat_line():
            keymap.command_InputKey("End")()
            keymap.command_InputKey("Delete")()

        def open_line():
            keymap.command_InputKey("End","Enter")()
            set_insertmode()

        def undo():
            keymap.command_InputKey("C-z")()

        def redo():
            if isVba(keymap.getWindow()):
                keymap.command_InputKey("D-Alt")()
                keymap.command_InputKey("E","R")()
                keymap.command_InputKey("U-Alt")()
            else:
                keymap.command_InputKey("C-y")()

        def yank():
            keymap.command_InputKey("C-c")()
            set_vimmode()

        def kill():
            keymap.command_InputKey("C-x")()
            set_vimmode()

        def delete_char():
            keymap.command_InputKey("Delete")()

        @profile
        def paste(sft=0):
            if keymap_vim.mainmode != 3:
                if keymap_vim.flg_selmode==1:
                    if isExcel(keymap.getWindow()):
                        if sft==0:
                            keymap.command_InputKey("Down")()
                        select_lineVmode()
                    else:
                        if sft:
                            keymap.command_InputKey("Home","Enter","Up")()
                        else:
                            keymap.command_InputKey("End","Enter")()

            if keymap_vim.flg_Mdentaku == 1:
                keymap.command_InputKey("C-F11")()
            else:
                keymap.command_InputKey("C-v")()
            keymap_vim.flg_selmode = 0
            keymap_vim.flg_Mdentaku =0
            set_vimmode()

        @profile
        def search():
            keymap.command_InputKey("C-f")()
            set_searchmode()

        def replace():
            if (isExcel(keymap.getWindow()) or
                    isWord(keymap.getWindow())or
                    isVba(keymap.getWindow())):
                keymap.command_InputKey("C-h")()
            else:
                keymap.command_InputKey("C-r")()
            set_searchmode()

        def search_next():
            keymap.command_InputKey("F3")()

        def search_prev():
            keymap.command_InputKey("S-F3")()

        def scroll_up():
            keymap.command_MouseWheel(1.0)()

        def scroll_down():
            keymap.command_MouseWheel(-1.0)()

        def select_line():
            if isEditorClass(keymap.getWindow()):
                keymap.command_InputKey("S-Down")()
            else:
                keymap.command_InputKey("S-End")()

        def select_lineVmode():
            if isExcel(keymap.getWindow()):
                keymap.command_InputKey("S-Space")()
            else:
                keymap.command_InputKey("Home")()
                select_line()

        def print_out():
            keymap.command_InputKey("C-p")()

        def close_file():
            if isForumPrev(keymap.getWindow()):
                keymap.command_InputKey("A-x")()
            else:
                keymap.command_InputKey("C-F4")()

        def close_app():
            keymap.command_InputKey("A-F4")()

        def open_file():
            keymap.command_InputKey("C-o")()

        def open_newfile():
            if isAndroidStudio(keymap.getWindow()):
                keymap.command_InputKey("C-F2")()
            else:
                keymap.command_InputKey("C-n")()

        def save_file():
            keymap.command_InputKey("C-s")()

        def save_namingfile():
            keymap.command_InputKey("A-f", "A-a")()

        @profile
        def window_vs():
            if isExcel(keymap.getWindow()):
                keymap.command_InputKey("A-w","A-a","A-v")()
                keymap.command_InputKey("Enter")()
            elif isAutoCad(keymap.getWindow()):
                keymap.command_InputKey("A-W","A-T")()
                keymap.command_InputKey("Enter")()
            elif isVba(keymap.getWindow()):
                keymap.command_InputKey("A-W","A-V")()
                keymap.command_InputKey("Enter")()
            elif isWord(keymap.getWindow()):
                keymap.command_InputKey("A-W","A-B")()

        @profile
        def window_vs_this():
            if isExcel(keymap.getWindow()):
                keymap.command_InputKey("A-W","A-N")()
                keymap.command_InputKey("A-w","A-a","A-v")()
                keymap.command_InputKey("Enter")()
            elif isAutoCad(keymap.getWindow()):
                keymap.command_InputKey("A-W","A-T")()
                keymap.command_InputKey("Enter")()
            elif isVba(keymap.getWindow()):
                keymap.command_InputKey("A-W","A-V")()
                keymap.command_InputKey("Enter")()
            elif isWord(keymap.getWindow()):
                keymap.command_InputKey("A-W","A-N")()
                keymap.command_InputKey("C-F6")()
                keymap.command_InputKey("A-W","A-B")()

        @profile
        def window_sp():
            if isExcel(keymap.getWindow()):
                keymap.command_InputKey("D-Alt")()
                keymap.command_InputKey("W","A","O")()
                keymap.command_InputKey("U-Alt")()
                keymap.command_InputKey("Enter")()
            if isAutoCad(keymap.getWindow()):
                keymap.command_InputKey("D-Alt")()
                keymap.command_InputKey("W","H")()
                keymap.command_InputKey("U-Alt")()
                keymap.command_InputKey("Enter")()

        @profile
        def docu_warituke():
            if isDocuWorks(keymap.getWindow()):
                keymap.command_InputKey("A-t","A-0","A-2")()

        @profile
        def paste_value():
            keymap.command_InputKey("D-Alt")()
            keymap.command_InputKey("E","S","V")()
            keymap.command_InputKey("U-Alt")()
            keymap.command_InputKey("Enter")()

        @profile
        def print_preview():
            keymap.command_InputKey("D-Alt")()
            keymap.command_InputKey("F","V")()
            keymap.command_InputKey("U-Alt")()
            keymap.command_InputKey("Enter")()

        def insert_cells():
            keymap.command_InputKey("C-Add")()

        def delete_cells():
            keymap.command_InputKey("C-Subtract")()

        def window_only():
            keymap.command_InputKey("C-F10")()

        def next_window():
            if isWord(keymap.getWindow()):
                keymap.command_InputKey("C-F6")()
            else:
                keymap.command_InputKey("C-Tab")()

        def prev_window():
            if isWord(keymap.getWindow()):
                keymap.command_InputKey("C-S-F6")()
            else:
                keymap.command_InputKey("C-S-Tab")()

        def record_start():
            keymap_vim.flg_mcr = 1
            keymap.command_RecordStart()

        def record_stop():
            keymap_vim.flg_mcr=0
            keymap.command_RecordStop()

        def record_play():
            keymap_vim.flg_mcr=0
            repeat2(keymap.command_RecordPlay)()

        @profile
        def hold_on():
            if isExcel(keymap.getWindow()):
                if keymap_vim.flg_selmode==1:
                    keymap.command_InputKey("C-9")()
                elif keymap_vim.flg_selmode==2:
                    keymap.command_InputKey("C-0")()

        def hold_off():
            if isExcel(keymap.getWindow()):
                if keymap_vim.flg_selmode==1:
                    keymap.command_InputKey("C-S-9")()
                elif keymap_vim.flg_selmode==2:
                    keymap.command_InputKey("C-S-0")()

        @profile
        def select_move(_fanc):
            if keymap_vim.flg_selmode!=2 or isExcel(keymap.getWindow()):
                keymap.command_InputKey("D-LShift")()
                keymap.command_InputKey("D-RShift")()
                rtn = _fanc()
                keymap.command_InputKey("U-LShift")()
                keymap.command_InputKey("U-RShift")()
            else:
                rtn= _fanc()
            return rtn

        def set_scroll(swt):
            keymap_vim.flg_scroll = swt

        def Vimmode_and_Reset():
            vim_parm_AllReset()
            set_vimmode()

        def tag_stash():
            if isVba(keymap.getWindow()):
                keymap.command_InputKey("S-F2")()

        def tag_pop():
            if isVba(keymap.getWindow()):
                keymap.command_InputKey("C-S-F2")()

        def exl_hide_row():
            if isExcel(keymap.getWindow()):
                keymap.command_InputKey("A-O","A-R","A-H")()

        def exl_hide_col():
            if isExcel(keymap.getWindow()):
                keymap.command_InputKey("A-O","A-C","A-H")()

        def exl_show_row():
            if isExcel(keymap.getWindow()):
                keymap.command_InputKey("A-O","A-R","A-U")()

        def exl_show_col():
            if isExcel(keymap.getWindow()):
                keymap.command_InputKey("A-O","A-C","A-U")()

        ########################################################################
        # VimModeでのコマンド
        ########################################################################

        def repeat(fanc):
            def _fanc():
                N = keymap_vim.repeatN
                keymap_vim.repeatN =0
                if N==0:
                    N=1
                for i in range(N):
                    fanc()
            return _fanc

        def repeat2(fanc):
            def _fanc():
                N = keymap_vim.repeatN
                keymap_vim.repeatN =0
                if N==0:
                    N=1
                for i in range(N):
                    fanc()
            return _fanc

        def input_command(char):
            keymap_vim.command_str += exc_char(char,1)
            show_command()

        def back_commandchar():
            strn = len(keymap_vim.command_str)
            if strn==1:
                dummy=""
            else:
                dummy=keymap_vim.command_str[0:strn-1]
            keymap_vim.command_str = dummy
            show_command()

        @profile
        def execute_command():
            if keymap_vim.command_str == "s/":
                replace()
            elif keymap_vim.command_str == "q":
                close_file()
            elif keymap_vim.command_str == "qa":
                close_app()
            elif keymap_vim.command_str == "w":
                save_file()
            elif keymap_vim.command_str == "w f":
                save_namingfile()
            elif keymap_vim.command_str == "e.":
                open_file()
            elif keymap_vim.command_str == "new":
                open_newfile()
            elif keymap_vim.command_str == "ha":
                print_out()
            elif keymap_vim.command_str == "vs":
                window_vs()
            elif keymap_vim.command_str == "vs this":
                window_vs_this()
            elif keymap_vim.command_str == "sp":
                window_sp()
            elif keymap_vim.command_str == "wari":
                docu_warituke()
            elif keymap_vim.command_str == "pstv":
                paste_value()
            elif keymap_vim.command_str == "preview":
                print_preview()
            elif keymap_vim.command_str == "inscells":
                insert_cells()
            elif keymap_vim.command_str == "delcells":
                delete_cells()
            elif keymap_vim.command_str == "only":
                window_only()
            elif keymap_vim.command_str == "outapplist":
                command_OutAppList()
            elif keymap_vim.command_str == "toplevelwindow":
                command_TopLevelWindow()
            elif keymap_vim.command_str == "getprocessname":
                command_GetProcessName()
            elif keymap_vim.command_str == "thisclass":
                command_ThisClass()
            elif keymap_vim.command_str == "desktop":
                command_DeskTop()
            elif keymap_vim.command_str == "set scb":
                set_scroll(1)
            elif keymap_vim.command_str == "set noscb":
                set_scroll(0)
            elif keymap_vim.command_str == "set fixinput":
                set_fixinputmode()
            elif keymap_vim.command_str == "set nofixinput":
                reset_fixinputmode()
            elif keymap_vim.command_str == "hiderow":
                exl_hide_row()
            elif keymap_vim.command_str == "hidecol":
                exl_hide_col()
            elif keymap_vim.command_str == "showrow":
                exl_show_row()
            elif keymap_vim.command_str == "showcol":
                exl_show_col()

            #show_command(1)
            keymap_vim.command_str = ""
            if keymap_vim.mainmode==4:
                set_vimmode(1)


        @profile
        def ScrollBind(fanc):
            cnt = 1
            flg_sb = (keymap_vim.flg_scroll==1
                    and keymap_vim.mainmode==1
                    #エクセルのセル編集中は無効にする
                    and not keymap.getWindow().getClassName().startswith("EXCEL6"))
            if flg_sb:
                cnt = 2
                rptN = keymap_vim.repeatN

            for i in range(cnt):

                fanc()

                if flg_sb:
                    if i ==0:
                        next_window()
                        keymap_vim.repeatN = rptN
                    elif i ==1:
                        prev_window()

        ########################################################################
        # キーマップ
        ########################################################################
        # 移動関係
        @profile
        def move_method(ikey):
            def _fanc():
                rtn =1

                if ikey == "k":
                    ScrollBind(repeat(move_up))
                elif ikey == "j":
                    ScrollBind(repeat(move_down))
                elif ikey == "h":
                    ScrollBind(repeat(move_left))
                elif ikey == "l":
                    ScrollBind(repeat(move_right))
                elif ikey == "w":
                    ScrollBind(repeat(move_next_word))
                elif ikey == "b":
                    ScrollBind(repeat(move_back_word))
                elif ikey == "RC-f":
                    ScrollBind(repeat(move_pagedown))
                elif ikey == "LC-b":
                    ScrollBind(repeat(move_pageup))
                elif ikey == "RC-b":
                    ScrollBind(repeat(move_pageup))
                elif ikey == "LC-u":
                    ScrollBind(repeat(move_halfpageup))
                elif ikey == "RC-d":
                    ScrollBind(repeat(move_halfpagedown))
                elif ikey == "S-g":
                    ScrollBind(move_bufend)
                elif ikey == "Caret":
                    ScrollBind(move_line_top)
                elif ikey == "0":
                    ScrollBind(move_line_top)
                elif ikey == "S-4":
                    ScrollBind(move_line_end)
                else:
                    rtn = 0

                return rtn
            return _fanc

        def method_G(key):
            def _fanc():
                if key == "g":
                    ScrollBind(repeat(move_buftop))
                elif key == "t":
                    ScrollBind(repeat(move_nexttab))
                elif key == "S-t":
                    ScrollBind(repeat(move_prevtab))

                return 1
            return _fanc

        def method_D(key):
            def _fanc():
                rtn =1
                if key=="d":
                    keymap.command_InputKey("Home")()
                    repeat(select_line)()
                    if isEditorClass(keymap.getWindow()):
                        select_move(move_left)
                elif select_move(move_method(key)):
                    if key=="S-4" and isEditorClass(keymap.getWindow()):
                        select_move(move_left)
                else:
                    rtn = 0
                if rtn:
                    kill()
                return 1

            return _fanc

        def method_Y(key):
            def _fanc():
                rtn =1
                if key=="y":
                    keymap.command_InputKey("Home")()
                    repeat(select_line)()
                elif select_move(move_method(key)):
                    dummy =0
                else:
                    rtn = 0
                if rtn:
                    yank()
                return 1

            return _fanc

        @profile
        def method_C(key):
            def _fanc():
                rtn =1
                if key=="c":
                    keymap.command_InputKey("Home")()
                    repeat(select_line)()
                    if isEditorClass(keymap.getWindow()):
                        select_move(move_left)
                elif select_move(move_method(key)):
                    if key=="S-4" and isEditorClass(keymap.getWindow()):
                        select_move(move_left)
                else:
                    rtn = 0
                if rtn:
                    kill()
                    set_insertmode()
                return 1

            return _fanc

        def method_Q(key):
            def _fanc():
#                if key=="a":
#                    record_start()
                if key in "abcdefghijklmnopqrstuvwxyz":
                    start_rec(CtoNum(key))
                elif key == "Enter" or key == "S-Enter":
                    start_rec(CtoNum(key))
                return 1
            return _fanc

        def method_Z(key):
            def _fanc():
                if key=="f":
                    hold_on()
                    set_vimmode()
                    keymap_vim.flg_selmode=0
                elif key=="o":
                    hold_off()
                    set_vimmode()
                    keymap_vim.flg_selmode=0
                return 1
            return _fanc

        def method_Atmark(key):
            def _fanc():
#                if key=="a":
#                    record_play()
                if key in "abcdefghijklmnopqrstuvwxyz":
                    repeat(execute_macro(CtoNum(key)))()
                elif key == "Enter" or key == "S-Enter":
                    repeat(execute_macro(CtoNum(key)))()
                return 1
            return _fanc

        def method_CW(key):
            def _fanc():
                if key == "RC-w":
                    next_window()
                elif key == "l":
                    next_window()
                elif key == "h":
                    prev_window()

                return 1
            return _fanc

        def select_method(key):
            if key == "g":
                keymap_vim.method_fanc = method_G
            elif key == "d":
                keymap_vim.method_fanc = method_D
            elif key == "y":
                keymap_vim.method_fanc = method_Y
            elif key == "c":
                keymap_vim.method_fanc = method_C
            elif key == "q":
                keymap_vim.method_fanc = method_Q
            elif key == "z":
                keymap_vim.method_fanc = method_Z
            elif key == "Atmark":
                keymap_vim.method_fanc = method_Atmark
            elif key == "RC-w":
                keymap_vim.method_fanc = method_CW

            if key == "g":
                keymap_vim.flg_mtd = 2
            else:
                keymap_vim.flg_mtd =1

        # モード制御
        @profile
        def vim_command_InputKey(ikey):


            # VimMode(ノーマルモード)
            if keymap_vim.mainmode==1:
                if keymap_vim.flg_mtd == 0:

                    if (ikey == "w" and (isCraftWare(keymap.getWindow())or isAf(keymap.getWindow()))
                            and keymap_vim.flg_cf_mode != 1):
                        keymap.command_InputKey("w")()
                    elif move_method(ikey)():
                        dummy = 0
                    elif ikey == "S-Enter":
                        keymap.command_InputKey("S-Enter")()
                    elif ikey == "r":
                        if isCraftWare(keymap.getWindow()):
                            keymap_vim.flg_cf_mode=1
                            keymap.command_InputKey("r")()
                        elif isAf(keymap.getWindow()):
                            keymap_vim.flg_cf_mode=1
                            keymap.command_InputKey("r")()
                    elif ikey == "m":
                        if isCraftWare(keymap.getWindow()):
                            keymap_vim.flg_cf_mode=1
                            keymap.command_InputKey("m")()
                        elif isAf(keymap.getWindow()):
                            keymap.command_InputKey("m")()
                    elif ikey == "e":
                        if isAndroidStudio(keymap.getWindow()):
                            keymap.command_InputKey("F2")()
                        else:
                            keymap.command_InputKey("e")()
                    elif ikey == "S-h":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("h")()
                        elif isAf(keymap.getWindow()):
                            keymap.command_InputKey("h")()
                    elif ikey == "LC-j":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("j")()
                        elif isAf(keymap.getWindow()):
                            keymap.command_InputKey("j")()
                    elif ikey == "S-j":
                        if isCraftWare(keymap.getWindow()) or isAf(keymap.getWindow()):
                            keymap.command_InputKey("S-j")
                        else :
                            concat_line()
                    elif ikey == "LC-k":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("k")()
                        if isAf(keymap.getWindow()):
                            keymap.command_InputKey("d")()
                    elif ikey == "S-j":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("S-j")()
                            keymap_vim.flg_cf_mode = 1
                            set_searchmode()
                        elif isAf(keymap.getWindow()):
                            keymap.command_InputKey("S-j")()
                            keymap_vim.flg_cf_mode = 1
                            set_searchmode()
                    elif ikey=="C-S-c":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("S-c")()
                            keymap_vim.flg_cf_mode = 1
                            set_searchmode()
                    elif ikey=="S-c":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("C-S-c")()
                        elif isAf(keymap.getWindow()):
                            keymap.command_InputKey("A-BackSlash")()
                        else:
                            method_C("S-4")()
                    elif ikey=="f":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("Colon")()
                        elif isAf(keymap.getWindow()):
                            keymap.command_InputKey("Colon")()
                    elif ikey == "u":
                        repeat(undo)()
                    elif ikey == "RC-r":
                        repeat(redo)()
                    elif ikey == "n":
                        if isTButton(keymap.getWindow()):
                            keymap.command_InputKey("n")()
                        else:
                            ScrollBind(repeat(search_next))
                    elif ikey == "S-n":
                        if isTButton(keymap.getWindow()):
                            keymap.command_InputKey("S-n")()
                        else:
                            ScrollBind(repeat(search_prev))
                    elif ikey == "i":
                        if (isCraftWare(keymap.getWindow()) and
                            keymap_vim.flg_cf_mode==0):
                                keymap.command_InputKey("i")()
                        elif (isAf(keymap.getWindow()) and
                            keymap_vim.flg_cf_mode==0):
                                keymap.command_InputKey("i")()
                        else:
                            set_insertmode()
                    elif ikey == "a":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("a")()
                        elif isAf(keymap.getWindow()):
                            keymap.command_InputKey("a")()
                        elif isExcel(keymap.getWindow()):
                            keymap.command_InputKey("F2")()
                        else:
                            keymap.command_InputKey("Right")()
                            set_insertmode()
                    elif ikey == "S-a":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("S-a")()
                        elif isAf(keymap.getWindow()):
                            keymap.command_InputKey("S-a")()
                        elif isExcel(keymap.getWindow()):
                            keymap.command_InputKey("End")()
                            keymap.command_InputKey("Right")()
                            keymap.command_InputKey("F2")()
                        else:
                            keymap.command_InputKey("End")()
                            set_insertmode()
                    elif ikey == "S-i":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("S-i")()
                        elif isAf(keymap.getWindow()):
                            keymap.command_InputKey("S-i")()
                        elif isExcel(keymap.getWindow()):
                            keymap.command_InputKey("End")()
                            keymap.command_InputKey("Left")()
                            keymap.command_InputKey("F2")()
                        else:
                            keymap.command_InputKey("Home")()
                            set_insertmode()
                    elif ikey == "s":
                        if (isCraftWare(keymap.getWindow()) and
                            keymap_vim.flg_cf_mode==0):
                                keymap.command_InputKey("s")()
                        elif (isAf(keymap.getWindow()) and
                            keymap_vim.flg_cf_mode==0):
                                keymap.command_InputKey("s")()
                        else:
                            keymap.command_InputKey("Delete")()
                            set_insertmode()
                    elif ikey == "x":
                        repeat(delete_char)()
                    elif ikey == "LC-Caret":
                        toggle_imemode()
                    elif ikey == "LC-0":
                        command_NextApplication(1)()
                    elif ikey == "LC-9":
                        command_NextApplication(2)()
                    elif ikey == "p":
                        if (isCraftWare(keymap.getWindow()) and
                                keymap_vim.flg_cf_mode==0):
                            keymap.command_InputKey("p")()
                        elif (isAf(keymap.getWindow()) and
                                keymap_vim.flg_cf_mode==0):
                            keymap.command_InputKey("p")()
                        else:
                            paste()
                    elif ikey == "o":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("o")()
                        elif isAf(keymap.getWindow()):
                            keymap.command_InputKey("o")()
                        else:
                            open_line()
                    elif ikey == "S-p":
                        paste(1)
                    elif ikey == "LC-p":
                        print_out()
                    elif ikey == "Slash":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("f")()
                            keymap_vim.flg_cf_mode =1
                            set_searchmode()
                        elif isAf(keymap.getWindow()):
                            keymap.command_InputKey("f")()
                            keymap_vim.flg_cf_mode =1
                            set_searchmode()
                            set_imeoff()
                        else:
                            search()
                    elif ikey == "LC-y":
                        ScrollBind(scroll_up)
                    elif ikey == "RC-e":
                        ScrollBind(scroll_down)
                    elif ikey == "v":
                        set_visualmode()
                    elif ikey == "C-S-z":
                        keymap.command_ClipboardList()
                    elif ikey == "S-v":
                        keymap_vim.flg_selmode = 1
                        select_lineVmode()
                        set_visualmode()
                    elif ikey == "RC-v":
                        switch_rect_sel()
                        set_visualmode()
                    elif ikey == "Esc":
                        vim_parm_reset()
                        keymap.command_InputKey("Esc")()
#                        keymap_vim.flg_fixinput=0
                    elif ikey == "LC-RShift":
                        input_esc()
                    elif ikey == "Back":
                        repeat(input_backspace)()
                    elif ikey == "Tab":
                        repeat(input_tab)()
                    elif ikey == "Space":
                        repeat(input_Space)()
                    elif ikey == "S-Tab":
                        repeat(input_stab)()
                    elif ikey =="LC-CloseBracket":
                        tag_stash()
                    elif ikey =="RC-t":
                        tag_pop()
                    elif ikey == "Enter":
                        ScrollBind(repeat(tenkey_enter))
                    elif (ikey == "g" or ikey =="d" or ikey=="y" or ikey=="c" or
                          ikey == "z" or ikey=="RC-w" or ikey=="Atmark"):
                        if ((not ikey=="RC-w" and not ikey=="g") and
                             (isCraftWare(keymap.getWindow()) or
                             isAf(keymap.getWindow())) and
                             keymap_vim.flg_cf_mode==0):
                            if (isAf(keymap.getWindow()) and
                                ikey=="d"):
                                keymap.command_InputKey("l")()
                            else:
                                keymap.command_InputKey(ikey)()
                        elif (ikey=="y" and isTButton(keymap.getWindow())):
                            keymap.command_InputKey(ikey)()
                        else:
                            select_method(ikey)
                    elif ikey == "S-d":
                        method_D("S-4")()
                    elif ikey == "S-y":
                        method_Y("S-4")()
                    elif ikey == "q":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("q")()
                        elif isAf(keymap.getWindow()):
                            keymap.command_InputKey("q")()
                        elif keymap_vim.flg_mcr==1:
#                            record_stop()
                            stop_rec()
                        elif keymap_vim.flg_mcr==0:
                            select_method(ikey)
                    elif ikey == "Colon":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("x")()
                            set_searchmode()
                        else:
                            set_commandmode()
                    else:
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey(ikey)()
                        elif isAf(keymap.getWindow()):
                            keymap.command_InputKey(ikey)()
                        elif isCtrlAlt(ikey):
                            keymap.command_InputKey(ikey)()

                else:
                    keymap_vim.flg_mtd = 0
                    keymap_vim.method_fanc(ikey)()


            # InsertMode(挿入モード)
            elif keymap_vim.mainmode==2:
                if ikey == "RC-f":
                    move_right()
                elif ikey == "LC-b":
                    move_left()
                elif ikey == "LC-n":
                    move_down()
                elif ikey == "LC-p":
                    move_up()
                elif ikey =="RC-a":
                    move_line_top()
                elif ikey =="RC-e":
                    move_line_end()
                elif ikey == "Esc":
                    set_vimmode()
                    keymap_vim.flg_cf_mode = 1
                elif ikey == "LC-RShift":
                    input_esc()
                    keymap_vim.flg_cf_mode = 1
                elif ikey == "Enter":
                    if (isEnterCanselClass(keymap.getWindow())
                         and keymap_vim.flg_fixinput==0):
                        ScrollBind(tenkey_enter)
                        set_vimmode()
                    else:
                        ScrollBind(input_enter)
                    if isCraftWare(keymap.getWindow()):
                        keymap_vim.flg_cf_mode = 0
                elif ikey == "LC-y":
                    scroll_up()
                elif ikey == "RC-e":
                    scroll_down()
                elif ikey == "LC-0":
                    command_NextApplication(1)()
                elif ikey == "LC-9":
                    command_NextApplication(2)()
                elif ikey == "LC-j":
                    mWnd = keymap.getWindow()
                    if mWnd.getImeStatus()==0:
                        keymap.command_InputKey("A-(243)")()
                    else:
                        keymap.command_InputKey(ikey)()
                else:
                    keymap.command_InputKey(ikey)()

            # VisualMode(ビジュアルモード)
            elif keymap_vim.mainmode==3:
                if keymap_vim.flg_mtd == 0:

                    if select_move(move_method(ikey)):
                        dummy = 0
                    elif ikey == "y":
                        yank()
                    elif ikey == "d":
                        kill()
                    elif ikey == "p":
                        paste()
                    elif ikey == "x":
                        delete_char()
                        set_vimmode()
                    elif ikey == "Esc":
                        if keymap_vim.flg_selmode==2:
                            keymap.command_InputKey("Esc")()
                        set_vimmode()
                    elif ikey == "LC-RShift":
                        if keymap_vim.flg_selmode==2:
                            keymap.command_InputKey("Esc")()
                        input_esc()
                    elif ikey == "LC-y":
                        scroll_up()
                    elif ikey == "RC-e":
                        scroll_down()
                    elif ikey == "Enter":
                        set_vimmode()
                    elif ikey == "Slash":
                        if isCraftWare(keymap.getWindow()):
                            keymap.command_InputKey("f")()
                            keymap_vim.flg_cf_mode =1
                            set_searchmode()
                        else:
                            search()
                    elif ikey == "s":
                        if (isCraftWare(keymap.getWindow()) and
                            keymap_vim.flg_cf_mode==0):
                                keymap.command_InputKey("s")()
                        else:
                            keymap.command_InputKey("Delete")()
                            set_insertmode()
                    elif ikey == "Colon":
                        set_commandmode()
                    elif (ikey == "g" or ikey == "z" or ikey == "q" or ikey == "Atmark"):
                        select_method(ikey)
                    elif isCtrlAlt(ikey):
                        keymap.command_InputKey(ikey)()

                else:
                    if keymap_vim.flg_mtd==2:
                        keymap_vim.flg_mtd =0
                        select_move(keymap_vim.method_fanc(ikey))
                    else:
                        keymap_vim.flg_mtd = 0
                        keymap_vim.method_fanc(ikey)()

            # CommandMode(コマンドモード)
            elif keymap_vim.mainmode==4:
                if (ikey == "Esc" or ikey =="LC-RShift"):
                    keymap_vim.command_str = ""
                    show_command(1)
                    set_vimmode()
                elif ikey == "Enter":
                    execute_command()
                elif ikey == "Back":
                    back_commandchar()
                else:
                    input_command(ikey)

            # SearchMode(挿入モード)
            elif keymap_vim.mainmode==5:
                if ikey == "Enter":
                    keymap.command_InputKey(ikey)()
                    set_vimmode()
                elif ikey == "C-Enter":
                    if isAf(keymap.getWindow()):
                        keymap.command_InputKey("Enter")()
                    keymap.command_InputKey(ikey)()
                    set_vimmode()
                elif ikey =="LC-p":
                    move_up()
                elif ikey =="LC-n":
                    move_down()
                elif ikey =="RC-f":
                    move_right()
                elif ikey =="LC-b":
                    move_left()
                elif ikey =="RC-a":
                    move_line_top()
                elif ikey =="RC-e":
                    move_line_end()
                elif ikey == "LC-j":
                    mWnd = keymap.getWindow()
                    if mWnd.getImeStatus()==0:
                        keymap.command_InputKey("A-(243)")()
                    else:
                        keymap.command_InputKey(ikey)()
                elif ikey == "Esc":
                    set_vimmode()
                elif ikey == "LC-RShift":
                    input_esc()
                else:
                    keymap.command_InputKey(ikey)()

        ########################################################################
        #  共通関数
        ########################################################################

        def exc_shortcut(ichar):
            if ichar == "LC-h":
                return "Back"
            if ichar == "LC-i":
                return "Tab"
            if ichar == "LC-m":
                return "Enter"
            return ichar

        def isCtrlAlt(ichar):
            if( ichar.startswith("RC-") or
                ichar.startswith("LC-") or
                ichar.startswith("C-") or
                ichar.startswith("D-Alt") or
                ichar.startswith("U-Alt") or
                ichar.startswith("A-C-") or
                ichar.startswith("Alt-") ):
                return True
            return False

        def send_vim_key(ikey):
            def _fanc():
                if keymap_vim.mainmode==0:
                    ckey = exc_shortcut(ikey)
                    keymap.command_InputKey(ckey)()
                else:
                    send_vimmodekey(ikey)
            return _fanc

        def send_vimmodekey(ikey):
            if keymap_vim.flg_mcr == 1:
                add_macro(ikey)
            if ikey != "LC-RShift":
                keymap_vim.esc_flg=0
            if ikey in '1234567890':
                send_vim_num(CtoNum(ikey))()
            else:
                ckey = exc_shortcut(ikey)
                vim_command_InputKey(ckey)

        def send_vim_num(num):
            def _fanc():
                if (isCraftWare(keymap.getWindow()) and
                                keymap_vim.flg_cf_mode!=1):
                    keymap.command_InputKey(str(num))()
                    return
                if (isAf(keymap.getWindow()) and
                                keymap_vim.flg_cf_mode!=1):
                    keymap.command_InputKey(str(num))()
                    return
                if isAfLogForm(keymap.getWindow()):
                    keymap.command_InputKey(str(num))()
                    return
                if keymap_vim.mainmode==0 or keymap_vim.mainmode==2 or keymap_vim.mainmode==5:
                    keymap.command_InputKey(str(num))()
                else:
                    if num==0 and keymap_vim.repeatN==0:
                        vim_command_InputKey("0")
                    else:
                        keymap_vim.repeatN = keymap_vim.repeatN*10 + num
            return _fanc

        ########################################################################
        #  実装
        ########################################################################
#        for key in range(10):
#            keymap_vim[str(key)]      = send_vim_num(key)
        for key in 'abcdefghijklmnopqrstuvwxyz1234567890-^\\@[;:],./\\':
            keymap_vim[exc_char(key)]      = send_vim_key(exc_char(key))
            keymap_vim["S-"+exc_char(key)]      = send_vim_key("S-"+exc_char(key))
            keymap_vim["LC-"+exc_char(key)]      = send_vim_key("LC-"+exc_char(key))
            keymap_vim["RC-"+exc_char(key)]      = send_vim_key("RC-"+exc_char(key))
            keymap_vim["Alt-"+exc_char(key)]      = send_vim_key("Alt-"+exc_char(key))
            keymap_vim["C-S-"+exc_char(key)]      = send_vim_key("C-S-"+exc_char(key))
            keymap_vim["A-C-"+exc_char(key)]      = send_vim_key("A-C-"+exc_char(key))

        keymap_vim["O-RShift"] = send_vim_key("Space")
        keymap_vim["O-(236)"] = send_vim_key("Enter")
        keymap_vim["C-(236)"] = send_vim_key("C-Enter")
        keymap_vim["S-(236)"] = send_vim_key("S-Enter")
        keymap_vim["S-LC-i"] = send_vim_key("S-Tab")
        keymap_vim["D-Alt"] = send_vim_key("D-Alt")
        keymap_vim["U-Alt"] = send_vim_key("U-Alt")

        keymap_vim["Esc"] = send_vim_key("Esc")
        keymap_vim["LC-RShift"] = send_vim_key("LC-RShift")
        keymap_vim["LC-F9"]= set_nomalmode
        keymap_vim["LC-F10"]= Vimmode_and_Reset

################################################################################


    # USER0-Alt-↑↓←→/Space/C-B/PageDown : キーボードで擬似マウス操作
    if 1:
        keymap_global[ "U0-A-H" ] = keymap.command_MouseMove(-10,0)
        keymap_global[ "U0-A-L" ] = keymap.command_MouseMove(10,0)
        keymap_global[ "U0-A-K" ] = keymap.command_MouseMove(0,-10)
        keymap_global[ "U0-A-J" ] = keymap.command_MouseMove(0,10)
        keymap_global[ "U0-RShift" ] = keymap.command_MouseButtonClick('left')
#        keymap_global[ "U-U0-RShift" ] = keymap.command_MouseButtonUp('left')
        keymap_global[ "U0-U" ] = keymap.command_MouseWheel(1.0)
        keymap_global[ "U0-N" ] = keymap.command_MouseWheel(-1.0)

################################################################################

################################################################################
#  スクロールテスト
################################################################################


    if 1:
        def command_scroll_test():
            keymap.getWindow().sendMessage(WM_SYSCOMMAND, SC_VSCROLL, 1)

        keymap_global["C-F3"] = command_scroll_test

################################################################################
# clunch と 内骨格の起動

    # USER0-F : ウインドウのアクティブ化
#    if 1:
#        keymap_global[ "C-F6" ] = keymap.command_ActivateWindow( "cfiler.exe", "CfilerWindowClass" )


    # USER0-E : アクティブ化するか、まだであれば起動する
    if 1:
        def command_ActivateOrExecuteClunch(flg=1):
            wnd = Window.find( "ClnchWindowClass", "CraftLaunch" )
            if wnd and flg:
                if wnd.isMinimized():
                    wnd.restore()
                wnd = wnd.getLastActivePopup()
                wnd.setForeground()
            else:
                executeFunc = keymap.command_ShellExecute( None, "..\clnch\clnch.exe", "", "" )
                executeFunc()
            if flg:
                keymap_vim.mainmode=5

        keymap_global[ "RC-F4" ] = command_ActivateOrExecuteClunch

    if 0:
        def command_ActivateOrExecuteCfilter():
            wnd = Window.find( "CfilerWindowClass", None)
            if wnd:
                if wnd.isMinimized():
                    wnd.restore()
                wnd = wnd.getLastActivePopup()
                wnd.setForeground()
            else:
                executeFunc = keymap.command_ShellExecute( None, "..\cfiler\cfiler.exe", "", "" )
                executeFunc()

        keymap_global[ "LC-F5" ] = command_ActivateOrExecuteCfilter

    if 1:
        def command_ActivateOrExecuteVimFiler():
            wnd = Window.find( "Vim", None)
            if wnd:
                if wnd.isMinimized():
                    wnd.restore()
                wnd = wnd.getLastActivePopup()
                wnd.setForeground()
            else:
                executeFunc = keymap.command_ShellExecute( None, "..\\portvim\\gvim.exe", '-c ":VimFilerDouble"', "" )
                executeFunc()

        keymap_global[ "LC-F6" ] = command_ActivateOrExecuteVimFiler

    if 1:
        def command_ActivateOrExecuteAf():
            wnd = Window.find( "TAfxWForm", None)
            if wnd:
                if wnd.isMinimized():
                    wnd.restore()
                wnd = wnd.getLastActivePopup()
                wnd.setForeground()
            else:
                executeFunc = keymap.command_ShellExecute( None, "..\\afxw32_157\AFXW.EXE", "", "" )
                executeFunc()
            set_vimmode(0)

        keymap_global[ "LC-F5" ] = command_ActivateOrExecuteAf

        keymap_global[ "U0-c" ] = command_ActivateOrExecuteVimFiler

    if 1:
        def command_ActivateOrExecuteMCal():
            wnd = Window.find( "TMdenMainForm", None)
            if wnd:
                if wnd.isMinimized():
                    wnd.restore()
                wnd = wnd.getLastActivePopup()
                wnd.setForeground()
            else:
                executeFunc = keymap.command_ShellExecute( None, "..\programfiles\Mdentaku\Mdentaku.exe", "", "" )
                executeFunc()

        keymap_global[ "LC-F12" ] = command_ActivateOrExecuteMCal


################################################################################
#  自作関数
################################################################################

    # Window関係
    if 1:
        def command_ToggleActiveWindowSize():
            wnd = keymap.getTopLevelWindow()
            if wnd:
                if wnd.isMinimized():
                    wnd.restore()
                if wnd.isMaximized():
                    wnd.restore()
                else:
                    wnd.maximize()

    keymap_global["U0-W"] = command_ToggleActiveWindowSize



################################################################################
#   テンキーのカスタマイズ
################################################################################
    if 1:

        keymap_vim.tenkeymode =0
        keymap_vim.tenkeycount =0

        def set_tenkeyHmode():
            keymap_vim.tenkeymode = 1

        def set_tenkeyVmode():
            keymap_vim.tenkeymode = 2
            keymap_vim.tenkeycount =0

        def set_tenkeyMcrmode():
            keymap_vim.tenkeymode = 3

        def set_nonemode():
            keymap_vim.tenkeymode = 0
            keymap_vim.tenkeycount = 0

        def tenkey_enter():
            if isTenKeyClass(keymap.getWindow()):
                if keymap_vim.tenkeymode==1:
                    keymap.command_InputKey("Tab")()
                elif keymap_vim.tenkeymode==2:
                    keymap.command_InputKey("Enter")()
                    keymap_vim.tenkeycount += 1
                elif keymap_vim.tenkeymode==3:
                    if keymap_vim.flg_mcr:                 #マクロ実行中or記録中
                        keymap.command_InputKey("Enter")()
                    else:
                        method_Atmark("Enter")()
                else:
                    keymap.command_InputKey("Enter")()
            else:
                keymap.command_InputKey("Enter")()
                if isMdentaku(keymap.getWindow()):
                    keymap_vim.flg_Mdentaku=1

        def tenkey_reset():
            if keymap_vim.tenkeymode == 1:
                keymap.command_InputKey("Enter")()
            elif keymap_vim.tenkeymode ==2:
                N = keymap_vim.tenkeycount+1
                keymap_vim.tenkeycount =0
                keymap.command_InputKey("Enter")()
                for i in range(N):
                    keymap.command_InputKey("Up")()
                keymap.command_InputKey("Right")()
            elif keymap_vim.tenkeymode ==3:
                method_Atmark("S-Enter")()

        def tenkey_up():
            if keymap_vim.tenkeymode == 2 and keymap_vim.tenkeycount >0:
                keymap_vim.tenkeycount -= 1
            keymap.command_InputKey("Up")()

        def input_tenkey(ichar):
            def _fanc():
                ime = keymap.getWindow().getImeStatus()
                set_imeoff()
                keymap.command_InputKey(ichar)()
                if ime:
                    set_imeon()
            return _fanc

        # キーの単純な置き換え
        keymap.replaceKey( "Enter", 236 )

        # ユーザモディファイアキーの定義
        keymap.defineModifier( 236, "User1" )

#        keymap_vim["O-(236)"]=tenkey_enter

        keymap_vim["U1-Num4"]=tenkey_reset

#        keymap_vim["C-m"]=tenkey_enter

        keymap_vim["Up"]=tenkey_up

        keymap_vim["U1-Num1"]=set_tenkeyHmode

        keymap_vim["U1-Num2"]=set_tenkeyVmode

        keymap_vim["U1-Num3"]=set_tenkeyMcrmode

        keymap_vim["U1-Num0"]=set_nonemode

        for i in range(10):
            keymap_global["Num"+str(i)] = input_tenkey("Num"+str(i))

        keymap_global["Divide"] = input_tenkey("Divide")
        keymap_global["Multiply"] = input_tenkey("Multiply")
        keymap_global["Subtract"] = input_tenkey("Subtract")
        keymap_global["Add"] = input_tenkey("Add")
        keymap_global["Decimal"] = input_tenkey("Decimal")

        keymap_global["C-(236)"] = keymap.command_InputKey("C-Enter")
        keymap_global["S-(236)"] = keymap.command_InputKey("S-Enter")
        keymap_global["A-(236)"] = keymap.command_InputKey("A-Enter")


        # for Vim
        def vim_input_esc():
            def _fanc():
                keymap.command_InputKey("Esc")()
                if keymap.getWindow().getImeStatus():
                    keymap.command_InputKey("A-(243)")()
            return _fanc

        #skk.vimへの対応(SKKIMEを起動させない)
        if 0:
            def vim_set_skk():
                def _fanc():
                    keymap.command_InputKey("C-J")()
                    #IMEのoff
                    if keymap.getWindow().getImeStatus():
                        keymap.command_InputKey("A-(243)")()
                    keymap.command_InputKey("C-J")()
                return _fanc

            keymap_ovim["LC-J"] = vim_set_skk()

        keymap_ovim = keymap.defineWindowKeymap(class_name=u'Vim*')
        keymap_ovim["LC-RShift"] = vim_input_esc()
        keymap_ovim["O-RShift"] = keymap.command_InputKey("Space")
        keymap_ovim["O-(236)"] = keymap.command_InputKey("Enter")
        keymap_ovim["C-(236)"] = keymap.command_InputKey("C-Enter")
        keymap_ovim["C-h"] = keymap.command_InputKey("Back")

        # for Console
        keymap_cnsl = keymap.defineWindowKeymap(class_name='ConsoleWindowClass')
        keymap_cnsl["LC-RShift"] = keymap.command_InputKey("ESC")
        keymap_cnsl["O-RShift"] = keymap.command_InputKey("Space")
        keymap_cnsl["O-(236)"] = keymap.command_InputKey("Enter")
        keymap_cnsl["C-(236)"] = keymap.command_InputKey("C-Enter")
        keymap_cnsl["C-h"] = keymap.command_InputKey("Back")
################################################################################
    # クリップボード履歴の最大数 (デフォルト:1000)
    keymap.clipboard_history.maxnum = 10000


    # クリップボード履歴リスト表示のカスタマイズ
    if 1:

        # 定型文
        fixed_items = [
            ( "hi-ho.mail",               "naohide-h@sky.hi-ho.ne.jp" ),
            ( "HomeAddress",       "〒738-0005 広島県廿日市市桜尾本町13-33-204" ),
            ( "電話番号",                  "03-4567-8901" ),
            ( "Edit config.py",          keymap.command_EditConfig ),
            ( "ReLoad config.py",  keymap.command_ReloadConfig ),
        ]

        # 日時
        date_and_time_items = [
            ( "YYYY/MM/DD HH:MM:SS",   dateAndTime("%Y/%m/%d %H:%M:%S") ),
            ( "YYYY/MM/DD",            dateAndTime("%Y/%m/%d") ),
            ( "YYYYMMDD_HHMMSS",       dateAndTime("%Y%m%d_%H%M%S") ),
            ( "YYYYMMDD",              dateAndTime("%Y%m%d") ),
            ( "HHMMSS",                dateAndTime("%H%M%S") ),
        ]

        keymap.cblisters += [
            ( "定型文",         cblister_FixedPhrase(fixed_items) ),
            ( "日時",           cblister_FixedPhrase(date_and_time_items) ),
            ]


# Ctrl-F8でアクティブなアクティブなアプリケーションの情報を得る

    if 1:
        def setClipboard(txt):
            def _fanc():
                ckit.ckit_misc.setClipboardText(txt)
            return _fanc

        def setAllCopy(txt1,txt2,txt3):
            def _fanc():
                ckit.ckit_misc.setClipboardText(txt3)
                ckit.ckit_misc.setClipboardText(txt2)
                ckit.ckit_misc.setClipboardText(txt1)
            return _fanc

        def command_GetApplicationInfo():

            # すでにリストが開いていたら閉じるだけ
            if keymap.isListWindowOpened():
                keymap.cancelListWindow()
                return

            def GetApplicationInfo():

                twnd = keymap.getWindow()
                topname = keymap.getTopLevelWindow().getClassName()
                clsname = twnd.getClassName()
                exename = twnd.getProcessName()

                applications = [
                    ( "TopLevelWindow  : "+topname+" ", setClipboard(topname) ),
                    ( "ClassName       : "+clsname+" ", setClipboard(clsname) ),
                    ( "ApplicationName : "+exename+" ", setClipboard(exename) ),
                    ( "全てクリップボードにコピー", setAllCopy(topname,clsname,exename) ),
                ]


                listers = [
                    ( "Application Info",     cblister_FixedPhrase(applications) ),
                ]

                item, mod = keymap.popListWindow(listers)

                if item:
                    item[1]()

            # キーフックの中で時間のかかる処理を実行できないので、delayedCall() をつかって遅延実行する
            keymap.delayedCall( GetApplicationInfo, 0 )

        keymap_global[ "LC-F8" ] = command_GetApplicationInfo


# USER0-F1 : カスタムのリスト表示をつかったアプリケーション起動
    if 1:
        def command_PopApplicationList():

            # すでにリストが開いていたら閉じるだけ
            if keymap.isListWindowOpened():
                keymap.cancelListWindow()
                return

            def popApplicationList():

                applications = [
                    ( "Notepad", keymap.command_ShellExecute( None, "notepad.exe", "", "" ) ),
                    ( "Paint", keymap.command_ShellExecute( None, "mspaint.exe", "", "" ) ),
                ]

                websites = [
                    ( "Google", keymap.command_ShellExecute( None, "https://www.google.co.jp/", "", "" ) ),
                    ( "Facebook", keymap.command_ShellExecute( None, "https://www.facebook.com/", "", "" ) ),
                    ( "Twitter", keymap.command_ShellExecute( None, "https://twitter.com/", "", "" ) ),
                ]

                listers = [
                    ( "App",     cblister_FixedPhrase(applications) ),
                    ( "WebSite", cblister_FixedPhrase(websites) ),
                ]

                item, mod = keymap.popListWindow(listers)

                if item:
                    item[1]()

            # キーフックの中で時間のかかる処理を実行できないので、delayedCall() をつかって遅延実行する
            keymap.delayedCall( popApplicationList, 0 )

        keymap_global[ "U0-F1" ] = command_PopApplicationList


# keyhac起動後clunchも起動させる
    command_ActivateOrExecuteClunch(0)
    keymap.command_InputKey("ESC")()
