" ==============================================================================
" プラグイン設定(Vundle.vim)
" [使い方]
"  1. 導入したいプラグインパスを以下に列挙
"  2. 「:PluginInstall」を実行する
" [参考]
"  ・プラグイン https://qiita.com/tanabee/items/e2064c5ce59c85915940
"  ・プラグインサイト https://vimawesome.com/
" ==============================================================================
" {{{
"set nocompatible
"filetype off
"set rtp+=~/.vim/bundle/Vundle.vim
"call vundle#begin()
"
"Plugin 'VundleVim/Vundle.vim'
"
"" ▼▼▼インストールプラグインパスここから▼▼▼
"" ex. Plugin '[Github Author]/[Github repo]'
"Plugin 'roblillack/vim-bufferlist'
"Plugin 'Align'
"Plugin 'FavEx'
"Plugin 'gtags.vim'
"Plugin 'itchyny/lightline.vim'
"Plugin 'mbriggs/mark.vim'
"Plugin 'thinca/vim-qfreplace'
"Plugin 'anyakichi/vim-surround'
"Plugin 'vim-scripts/jellybeans.vim'
"Plugin 'fuenor/qfixgrep'
"" ▲▲▲インストールプラグインパスここまで▲▲▲
"
"call vundle#end()
"filetype plugin indent on
" }}}

" ==============================================================================
" ファイルパス存在チェック
" <usage>
"   call CheckPathExists( $CTAGS )
" ==============================================================================
" {{{
	function! CheckPathExists( sPath )
		if filereadable( a:sPath )
			echo "exists!"
		else
			echo "not exists!"
		endif
	endfunction
" }}}

" ==============================================================================
" ユーザー定義プラグインフォルダを runtimepath に追加
" ==============================================================================
" {{{
if has('unix')
	let $USERPLUGINS = $HOME . '/.vim/_plugins_user'
else
	let $USERPLUGINS = $VIM . '/_plugins_user'
endif
	for s:path in split(glob( $USERPLUGINS . '/*'), '\n')
		if s:path !~# '\~$' && isdirectory(s:path)
			let &runtimepath = &runtimepath.','.s:path
		endif
	endfor
" }}}

" ==============================================================================
" ユーザー定義エクステンションを path に追加
" ==============================================================================
" {{{
if has('unix')
	let $EXTENTION = $HOME . '/.vim/_extention'
else
	let $EXTENTION = $VIM . '/_extention'
endif
	for s:path in split(glob( $EXTENTION . '/*'), '\n')
		if s:path !~# '\~$' && isdirectory(s:path)
			let $PATH .= ';' . s:path
		endif
	endfor
" }}}

" ==============================================================================
" ユーザー定義フォルダ配下のヘルプファイルのタグを作成する
"	※処理が遅いため、プラグインインストール時のみ逐次実行すること
" ==============================================================================
" {{{
	function! CreateHelptags()
		for s:path in split(glob( $USERPLUGINS . '/**/doc'), '\n')
			let s:txtexists = 0
			let s:tagsexists = 0
			for s:subpath in split(glob( s:path . '/*'), '\n')
				if s:subpath =~? '.*.txt$'
					let s:txtexists = 1
				endif
				if s:subpath =~? '.*tags$'
					let s:tagsexists = 1
				endif
			endfor
			if s:txtexists == 1 && s:tagsexists == 0
				if isdirectory(s:path)
					execute "helptags " . s:path
				endif
			endif
		endfor
	endfunction
" }}}

" ==============================================================================
" neocomplete.vim 用設定
" [参照] https://github.com/Shougo/neocomplete.vim
" ==============================================================================
" {{{
"	let g:neocomplete#enable_at_startup = 1 " 起動時に有効化
"	let g:neocomplete#enable_smart_case = 1 " Use smartcase.
"	let g:neocomplete#sources#syntax#min_keyword_length = 3
"	let g:neocomplete#max_list = 50 "キーワードの長さ、デフォルトで80
"	let g:neocomplete#max_keyword_width = 80
"	let g:neocomplete#enable_ignore_case = 1
""	highlight Pmenu ctermbg=6
""	highlight PmenuSel ctermbg=3
""	highlight PMenuSbar ctermbg=0
"	highlight Pmenu ctermbg=71
"	highlight PmenuSel ctermbg=71
"	highlight PMenuSbar ctermbg=71
" }}}

" ==============================================================================
" editexisting-ext.vim 用設定
" (他のVimで開いているファイルを開こうとしたときポップアップさせる)
" [参照] http://vimwiki.net/?tips%2F94
" [参照] https://github.com/koron/vim-kaoriya/issues/9
" ==============================================================================
" {{{
if has('unix')
	"do nothing
else
	packadd! editexisting
endif
" }}}

" ==============================================================================
" テキストファイルの自動改行抑止設定
" ==============================================================================
" {{{
	autocmd BufRead *.txt set tw=0
" }}}

" ==============================================================================
" 折り畳みマーカー設定
" ==============================================================================
" {{{
	set foldmethod=marker
" }}}

" ==============================================================================
" GVim 起動時は同じウィンドウにまとめて起動する。
" [参考] http://tyru.hatenablog.com/entry/20130430/vim_resident
" ==============================================================================
" {{{
"	call singleton#enable()
" }}}

" ==============================================================================
" lightline 設定
" [参考] https://github.com/itchyny/lightline.vim#landscape-theme-with-the-patched-font
" [参考] http://itchyny.hatenablog.com/entry/20130917/1379369171
" ==============================================================================
" {{{
	let g:lightline = {}
"	let g:lightline.colorscheme = 'jellybeans'
	let g:lightline.enable = {
					\		'statusline': 1,
					\		'tabline': 0
					\	}
	let g:lightline.separator = { 'left': '', 'right': '' }
	let g:lightline.subseparator = { 'left': '', 'right': '|' }
if has('unix')
	"unix時は絶対パス(absolutepath)を表示する
	let g:lightline.active = {
					\		'left': [
					\			[ 'mode', 'paste' ],
					\			[ 'filename' ],
					\			[ 'absolutepath', 'bufnum', 'readonly', 'modified', 'anzu' ]
					\		],
					\		'right': [
					\			[ 'rangediff' ],
					\			[ 'lineinfo', 'percent' ],
					\			[ 'fileformat', 'fileencoding', 'filetype' ]
					\		]
					\	}
else
	let g:lightline.active = {
					\		'left': [
					\			[ 'mode', 'paste' ],
					\			[ 'filename' ],
					\			[ 'bufnum', 'readonly', 'modified', 'anzu' ]
					\		],
					\		'right': [
					\			[ 'rangediff' ],
					\			[ 'lineinfo', 'percent' ],
					\			[ 'fileformat', 'fileencoding', 'filetype' ]
					\		]
					\	}
endif
	let g:lightline.component_function = {
					\		'rangediff': 'GetSelRngDiff',
					\		'anzu': 'anzu#search_status'
					\	}
" }}}

" ==============================================================================
" ctrlp 設定
" [参照] https://github.com/ctrlpvim/ctrlp.vim
" ==============================================================================
" {{{
	let g:ctrlp_map = '<F4>'
	let g:ctrlp_cmd = 'CtrlP'
	let g:ctrlp_working_path_mode = 'ra'
"	let g:ctrlp_root_markers = ['pom.xml', '.p4ignore', 'tags']
	let g:ctrlp_root_markers = ['pom.xml', '.p4ignore']
	
	command! -nargs=? Cpa CtrlP '/mnt/c/codes_sample'
" }}}

" ==============================================================================
" osc52 設定（リモート越しのローカルコピー）
" <usage>
"   1. ターミナルソフトの設定を行う(以下はRLogin時の手順)
"     1-1. Server Edit Entryを開く
"	  1-2. [クリップボード]->[制御コードによるクリップボード操作]
"     1-3. [OSC 52 によるクリップボードの書き込みを許可する]にチェック
"   2. osc52.vimをプラグインフォルダに格納
" <url> http://tateren.hateblo.jp/entry/2017/07/21/213020
" ==============================================================================
" {{{
if has('unix')
	execute 'source '. expand( "$HOME/.vim/_plugins_user/osc52/plugin/osc52.vim" )
	vnoremap y y:call SendViaOSC52(getreg('"'))<cr>
	nnoremap yy Vy:call SendViaOSC52(getreg('"'))<cr>
endif
" }}}

" ==============================================================================
" ビジュアルモード文字数カウント
" 覚書：列番号取得には苦労した。。。
"		詳細は以下 URL 参照。
"		  http://www49.atwiki.jp/draemonash/pages/69.html
" ==============================================================================
" {{{
	set updatetime=100 "CursorHold の閾値[ms]
	autocmd CursorMoved,CursorHold * call s:CalcSelRngStrNum()
"	autocmd ModeChanged *:[vV\x16]* * call s:CalcSelRngStrNum()
	let s:lStrtClm = 0
	let s:lStrtRow = 0
	let s:lDiffClm = 0
	let s:lDiffRow = 0
	let g:sSelRowsRng = ""
	
	function! s:CalcSelRngStrNum()
		let l:sCurMode = mode()
		let l:lCurRow = line('.')
		if l:sCurMode ==# "v" || l:sCurMode ==# "V" || l:sCurMode ==# "\<C-v>"
			let l:lCurClm = virtcol('.')
			let s:lDiffClm = abs( l:lCurClm - s:lStrtClm ) + 1
			let s:lDiffRow = abs( l:lCurRow - s:lStrtRow ) + 1
			if l:lCurRow > s:lStrtRow
				let g:sSelRowsRng = s:lStrtRow . "-" . l:lCurRow
			else
				let g:sSelRowsRng = l:lCurRow . "-" . s:lStrtRow
			endif
		elseif l:sCurMode ==# "n"
			let l:lCurClm = virtcol( [line('.'), col('.') - 1, 0] ) + 1
			let s:lStrtClm = l:lCurClm
			let s:lStrtRow = l:lCurRow
			let s:lDiffClm = 0
			let s:lDiffRow = 0
			let g:sSelRowsRng = l:lCurRow
		else
			"Do Nothing
		endif
	endfunction
	
	function! GetSelRngDiff()
		if s:lDiffClm == 0 && s:lDiffRow == 0
			return "-"
		else
			let l:sCurMode = mode()
			if l:sCurMode ==# "v" || l:sCurMode ==# "\<C-v>"
				return "h" . s:lDiffRow . " w" . s:lDiffClm
			elseif l:sCurMode ==# "V"
				return "h" . s:lDiffRow
			elseif l:sCurMode ==# "n"
				return "-"
			else
				return "error"
			endif
		endif
	endfunction
" }}}

" **************************************************************************************************
" *****									キーバインド設定									   *****
" **************************************************************************************************
" <<説明>>
"	- map：キー入力を別のキーに割り当てる。
"	- noremap：キー入力を別のキーに割り当てる。
"			   ただし、再マップされないので、マップが入れ子になったり再帰的になることがない。
"		map と noremap の違いについては、以下 URL 参照
"		  http://cocopon.me/blog/?p=3871
"	- <sirent>：メッセージは表示されず、 メッセージ履歴に残らない
"	- <expr>：引数が式 (スクリプト) として扱われる。マップが実行されたときに、式が評価される。
" <<覚書>>
"	- Ctrl+Shift+アルファベットは Vim の仕様上、キーバインドに設定できない。
"		参考URL）https://github.com/vim-jp/issues/issues/756
"	- map 定義行にコメントは基本的に記載できないが、「|」を使えば可能。
"		参考URL）https://vim-jp.org/vimdoc-ja/map.html
"			マップコマンドでは '"' (ダブルクォート) も {lhs} や {rhs} の一部と見なされるため、
"			マップコマンドの後ろにコメントを置くことはできません。しかし、|" を使うことができます。
"			これはコメント付きの空の新しいコマンドを開始するからです。

"=== モード共通 ===
" {{{
	noremap		<silent>			x			"_x|											" 1文字削除（切り取りなし）
"	nnoremap	<silent>			dd			"_dd|											" 行削除（切取りなし）
"	nnoremap	<silent>			cc			dd|												" 行切取り
	noremap		<silent>			H			0|												" カーソルを先頭に移動
	noremap		<silent>			L			$|												" カーソルを末尾に移動
"	noremap							/			/\v|											" very magic で検索
	noremap							/			/\V|											" very nomagic で検索
"	noremap		<silent>	<expr>	gp			'`[' . strpart(getregtype(), 0, 1) . '`]'|		" ペーストしたテキストを再選択できるようにする
"	map								go			<Plug>(openbrowser-smart-search)|				"【open-browser】ブラウザで開く
	noremap		<silent>			<space>		<nop>|											" スペースキー無効
	noremap		<silent>			<c-j>		10jzz|											" カーソル移動＋画面移動（下）
	noremap		<silent>			<c-k>		10kzz|											" カーソル移動＋画面移動（上）
	noremap		<silent>			<c-h>		10zh10h|										" カーソル移動＋画面移動（左）
	noremap		<silent>			<bs>		10zh10h|										" カーソル移動＋画面移動（左） for tmux
	noremap		<silent>			<c-l>		10zl10l|										" カーソル移動＋画面移動（右）
" }}}

"=== ノーマルモード ===
" {{{
"	nnoremap	<silent>			<cr>		<nop>|
	nnoremap				<expr>	gc			":Gc " . expand('<cword>')|						" Grep＠現在ファイル
	nnoremap				<expr>	gp			":Gp " . expand('<cword>')|						" Grep＠上位階層「tags」格納ディレクトリ配下
	nnoremap				<expr>	gt			":Gt"|											" Grep(対話型)＠ファイルパスなどを指定
	nnoremap						ciy			ciw<c-r>0<esc>b|								" 単語コピー
	nnoremap	<silent>			<esc><esc>	:nohlsearch<cr>|								" ハイライト解除
	nnoremap	<silent>			<tab>		a<tab><esc>|									" タブ挿入
	nnoremap						<c-d>		:Gtags <C-r><C-w><CR>|							"【gtags(*1)】定義先ジャンプ
	nnoremap						<c-e>		:Gtags -r <C-r><C-w><CR>|						"【gtags(*1)】呼出し元ジャンプ
"	nnoremap						gs			:Gtags -s <C-r><C-w><CR>|						"【gtags(*1)】
"	nnoremap						gG			:Gtags -g <C-r><C-w><CR>|						"【gtags(*1)】文字列Grep
"	nnoremap						,d			:<C-u>Gtags -f %<CR>|							"【gtags(*1)】関数一覧表示
	nnoremap						<c-t>		<C-o>|											"【gtags(*1)】
	nnoremap	<silent>			<c-g>		:call ExecuteGuiGrep()<cr>|						" GUIのGrepソフト起動
	nnoremap	<silent>			<c-m>		oa<esc>"_x|										" インデントを維持して改行
	nnoremap	<silent>			<c-tab>		<c-w>w|											" ウィンドウ切替（次へ）
	nnoremap	<silent>			<c-s-tab>	<c-w>W|											" ウィンドウ切替（前へ）
"	nnoremap	<silent>			<c-tab>		gt|												" タブ切替（次へ）
"	nnoremap	<silent>			<c-s-tab>	gT|												" タブ切替（前へ）
	nnoremap	<silent>			<c-w><c-k>	5<c-w>+|										" ウィンドウ幅を変える 上:高さup
	nnoremap	<silent>			<c-w><c-j>	5<c-w>-|										" ウィンドウ幅を変える 下:高さdown
	nnoremap	<silent>			<c-w><c-h>	5<c-w>>|										" ウィンドウ幅を変える 右:幅up
	nnoremap	<silent>			<c-w><c-l>	5<c-w><|										" ウィンドウ幅を変える 左:幅down
	nnoremap	<silent>			<space>h	<c-w>h|											" ウィンドウフォーカス移動(左)
	nnoremap	<silent>			<space>j	<c-w>j|											" ウィンドウフォーカス移動(下)
	nnoremap	<silent>			<space>k	<c-w>k|											" ウィンドウフォーカス移動(上)
	nnoremap	<silent>			<space>l	<c-w>l|											" ウィンドウフォーカス移動(右)
	nnoremap	<silent>			<c-n><c-h>	<c-w>h|											" ウィンドウフォーカス移動(左)
	nnoremap	<silent>			<c-n><c-j>	<c-w>j|											" ウィンドウフォーカス移動(下)
	nnoremap	<silent>			<c-n><c-k>	<c-w>k|											" ウィンドウフォーカス移動(上)
	nnoremap	<silent>			<c-n><c-l>	<c-w>l|											" ウィンドウフォーカス移動(右)
	nnoremap	<silent>			<M-h>		<c-w>h|											" ウィンドウフォーカス移動(左)
	nnoremap	<silent>			<M-j>		<c-w>j|											" ウィンドウフォーカス移動(下)
	nnoremap	<silent>			<M-k>		<c-w>k|											" ウィンドウフォーカス移動(上)
	nnoremap	<silent>			<M-l>		<c-w>l|											" ウィンドウフォーカス移動(右)
	nnoremap						<c-]>		g<c-]>|											" tagsジャンプの時に複数ある時は一覧表示
"	nnoremap	<silent>			<c-w>		:tabclose<cr>|									" タブを閉じる
"	nnoremap	<silent>			<c-n>		:tabnew<cr>|									" 新規タブを開く
"	nmap							<c-o>		:GtagsCursor<CR>|								"【gtags(※1)】
	nnoremap	<silent>			<F1>		:call BufferList()<cr>|							" バッファリスト作成
	nnoremap	<silent>			<F2>		:TagbarToggle<cr>|								" タグリスト作成
"	nnoremap	<silent>			<F3>														" <F3>はペーストモード切替に割り当て
"	nnoremap	<silent>						a<C-R>=strftime("%Y/%m/%d (%a)")<cr><esc>|		" 現在日時出力
"	nnoremap	<silent>						a<C-R>=strftime("%H:%M:%S")<cr><esc>|			" 現在時刻出力
"	nnoremap	<silent>			<F4>														" <F4>はCtrlP起動に割り当て
	nnoremap	<silent>			<F5>		:execute ExecCurrentScript()<cr>|				" 現在のプログラムを実行
"	nnoremap	<silent>			<F6>		:vs<cr><c-w>wggVGy:q<cr><c-w>W|					" 全体をコピー
	nnoremap	<silent>			<F6>		:call OpenCurFileWithExternalEditer($MYEXEPATH_CURSOR)<cr>|	" Cursorで開く
"	nnoremap	<silent>			<F7>		:Vexplore<cr>|									" Explorerを起動
"	nnoremap	<silent>			<F7>		:NERDTreeToggle<CR>|							" 【NERDtree】起動
	nnoremap	<silent>			<F7>		:call OpenCurFileWithExternalEditer($MYEXEPATH_VSCODE)<cr>|	" VSCodeで開く
	nnoremap	<silent>			<F8>		:call SwitchFontSize()<cr>|						" フォントサイズをトグル
	nmap		<silent>			<F9>		kyiwjciw<c-r>0<esc>b<c-a>j|						" 前行の単語をコピーしてインクリメント
	nnoremap	<silent>			<F10>		:call ToggleWindowSize()<cr>|					" ウィンドウサイズをトグル
	nnoremap	<silent>			<F11>		:set expandtab<cr>:retab!<cr>|					" タブ⇒空白 変換
	nnoremap	<silent>			<F12>		:set noexpandtab<cr>:retab!<cr>|				" 空白⇒タブ 変換
	nnoremap	<silent>			<s-F1>		:call CopyCurPrjRltvFilePath()<cr>|				" 現在相対ファイルパスコピー
	nnoremap	<silent>			<c-F1>		:call CopyCurFilePath()<cr>|					" 現在ファイルパスコピー
	nnoremap	<silent>			<c-F2>		:call CopyCurFileName()<cr>|					" 現在ファイル名コピー
	nnoremap	<silent>			<c-F3>		:call CopyCurDirPath()<cr>|						" 現在フォルダパスコピー
	nnoremap	<silent>			<c-F4>		:call CopyCurFileExt()<cr>|						" 現在ファイル拡張子コピー
"	nnoremap	<silent>			<c-F5>		:call CopyFileLineNo()<cr>|						" 現在行番号コピー
"	nnoremap	<silent>			<c-F5>		:call CopyFileExtAndLineNo()<cr>|				" 現在ファイル種別＆行番号コピー
	nnoremap	<silent>			<c-F5>		:call CopyRltvFilePathAndLineNo()<cr>|			" 現在相対ファイルパス＆行番号コピー
	nnoremap	<silent>			<c-F8>		:call SwitchBirdEyesViewMode()<cr>|				" 俯瞰モード
	nnoremap	<silent>			<c-s-F5>	:call UpdateTagFile()<cr>|						" タグファイル更新
	nmap		<silent>			n			<Plug>(anzu-n)zz|								" 検索結果を画面中央に
	nmap		<silent>			N			<Plug>(anzu-N)zz|								" 検索結果を画面中央に
	nmap		<silent>			*			<Plug>(anzu-star)zz|							" 検索結果を画面中央に
"	nnoremap	<silent>			*			:call SearchStart()<cr>|						" 検索結果を画面中央に
	nmap		<silent>			#			<Plug>(anzu-sharp)zz|							" 検索結果を画面中央に
	nmap		<silent>			K			g*zz|											" 検索結果を画面中央に
	nmap		<silent>			g#			g#zz|											" 検索結果を画面中央に
	nmap		<silent>			<Esc><Esc>	<Plug>(anzu-clear-search-status):noh<cr>|		" 検索結果を画面中央に
	" (*1) Gtags用キーバインド [参照] http://cha.la.coocan.jp/doc/gnu_global.html#sec10
" }}}

"=== 挿入モード ===
" {{{
	inoremap	<silent>			<c-j>		<esc>|											" Ctrl+J でノーマルモードに移行
	imap		<silent>			<c-k>		<Plug>(neosnippet_expand_or_jump)|				" 【neosnippet】スニペットを展開
	imap					<expr>	<TAB>		neosnippet#expandable_or_jumpable() ? "\<Plug>(neosnippet_expand_or_jump)" : pumvisible() ? "\<C-n>" : "\<TAB>"|	" 【neosnippet】次の展開へ
if has('unix')
	imap		<silent>			<c-v>		<c-r>0|											" 貼り付け
else
	imap		<silent>			<c-v>		<S-Insert>|										" 貼り付け
endif
"	inoremap	<silent>			{			{}<LEFT>|										" 括弧とクォーテーションの自動補完
"	inoremap	<silent>			[			[]<LEFT>|										" 括弧とクォーテーションの自動補完
"	inoremap	<silent>			(			()<LEFT>|										" 括弧とクォーテーションの自動補完
"	inoremap	<silent>			"			""<LEFT>|										" 括弧とクォーテーションの自動補完
"	inoremap	<silent>			'			''<LEFT>|										" 括弧とクォーテーションの自動補完
"	inoremap	<silent>			【			【】<LEFT>|										" 括弧とクォーテーションの自動補完（全角文字はなぜかうまく動かない…）
"	inoremap	<silent>			「			「」<LEFT>|										" 括弧とクォーテーションの自動補完（全角文字はなぜかうまく動かない…）
"	inoremap	<silent>			『			『』<LEFT>|										" 括弧とクォーテーションの自動補完（全角文字はなぜかうまく動かない…）
"	inoremap	<silent>			“			“”<LEFT>|										" 括弧とクォーテーションの自動補完（全角文字はなぜかうまく動かない…）
"	inoremap	<silent>			‘			‘’<LEFT>|										" 括弧とクォーテーションの自動補完（全角文字はなぜかうまく動かない…）
"	inoremap	<silent>			（			（）<LEFT>|										" 括弧とクォーテーションの自動補完（全角文字はなぜかうまく動かない…）
"	inoremap	<silent>			｛			｛｝<LEFT>|										" 括弧とクォーテーションの自動補完（全角文字はなぜかうまく動かない…）
" }}}

"=== ヴィジュアルモード ===
" {{{
	vmap		<silent>			Y			<esc>:set expandtab<cr>gv:retab!<cr>gvyu|		" 選択範囲をタブ->空白変換後にコピー
	vnoremap	<silent>			T			<esc>:set expandtab<cr>gv:retab!<cr>gv:call CopyLineWithLineNo()<cr>u|	" 選択範囲をタブ->空白変換後に行番号付き行コピー
"	vnoremap	<silent>			p			"0p<cr>											" ヤンクせずに貼り付け
	vnoremap	<silent>			s			c|												" 削除＆挿入モード
	vnoremap	<silent>			<			<gv|											" インデント左シフト
	vnoremap	<silent>			>			>gv|											" インデント右シフト
	vnoremap	<silent>			<F3>		:s/\s*$//<cr>:nohlsearch<cr>|					" 末尾の空白を削除
	vnoremap	<silent>			<F4>		:s/\\/\//g<cr>|									" ファイルパス区切り変換（\->/）
	vnoremap	<silent>			<F5>		:s/\v_(.)/\u\1/g<cr>|							" スネークケース⇒キャメルケース変換
	vnoremap	<silent>			<F6>		:s/\v([A-Z])/_\L\1/g<cr>|						" キャメルケース⇒スネークケース変換
	vnoremap	<silent>			<F7>		:s/\v^(\w+).*/\1/g<cr>|							" 行頭文字列のみ抽出
	vnoremap	<silent>			<F8>		:s/\v^\!.*\!\n//g<cr>|							" コンソール文字削除
	vnoremap	<silent>			<F9>		:call ReplaceRelativePathFromCurrent()<cr>|		" 相対パスへ変換
	vnoremap	<silent>			<F11>		<esc>:set expandtab<cr>gv:retab!<cr>|			" タブ⇒空白 変換
	vnoremap	<silent>			<F12>		<esc>:set noexpandtab<cr>gv:retab!<cr>|			" 空白⇒タブ 変換
	vnoremap	<silent>			<c-F5>		:<bs><bs><bs><bs><bs>call CopyFileLineNo()<cr>| " 行番号コピー
"	vnoremap	<silent>			d			"_d|											" 削除（切取りなし）
"	vnoremap	<silent>			c			d|												" 切取り
"	vnoremap	<silent>			p			"0p|											" 上書きペースト(クリップボード上書きなし)
"	vnoremap				<expr>	gc			":Gc " . expand('<cword>')|						" Grep＠現在ファイル
"	vnoremap				<expr>	gp			":Gp " . expand('<cword>')|						" Grep＠上位階層「tags」格納ディレクトリ配下
"	vnoremap				<expr>	gt			":Gt " . expand('<cword>')|						" Grep(対話型)＠ファイルパスなどを指定
" }}}

"=== コマンドラインモード ===
" {{{
	cnoremap						<c-n>		<Down>|											" コマンド履歴進む(シェル同等)
	cnoremap						<c-p>		<Up>|											" コマンド履歴戻る(シェル同等)
	cnoremap						<c-b>		<Left>|											" カーソル移動 左(シェル同等)
	cnoremap						<c-f>		<Right>|										" カーソル移動 右(シェル同等)
	cnoremap						<c-a>		<Home>|											" カーソル移動 行頭(シェル同等)
	cnoremap						<c-e>		<End>|											" カーソル移動 行末(シェル同等)
	cnoremap						<c-d>		<Del>|											" 文字削除(シェル同等)
if has('unix')
	cmap							<c-v>		<c-r>0|											" 貼り付け
else
	cmap							<c-v>		<S-Insert>|										" 貼り付け
endif
" }}}

"=== その他モード ===
" {{{
	smap							<C-k>		<Plug>(neosnippet_expand_or_jump)|				" 【neosnippet】スニペットを展開
"	xmap							<C-k>		<Plug>(neosnippet_expand_target)|				" 【neosnippet】スニペットを展開
	smap					<expr>	<TAB>		neosnippet#expandable_or_jumpable() ?  "\<Plug>(neosnippet_expand_or_jump)" : "\<TAB>"|	" 【neosnippet】次の展開へ
" }}}

"=== ターミナルウィンドウ ===
" {{{
"	tnoremap						<c-n><c-h>	<c-w>h|											" ウィンドウフォーカス移動(左)
"	tnoremap						<c-n><c-j>	<c-w>j|											" ウィンドウフォーカス移動(下)
"	tnoremap						<c-n><c-k>	<c-w>k|											" ウィンドウフォーカス移動(上)
"	tnoremap						<c-n><c-l>	<c-w>l|											" ウィンドウフォーカス移動(右)
"	tnoremap						<c-i>		<c-w>N|											" Terminal-Normalモード切替え
"	tnoremap						<esc><esc>	<c-w>:q!<cr>|									" Terminalモード終了
" }}}

" **************************************************************************************************
" *****										  基本設定										   *****
" **************************************************************************************************
" ==============================================================================
" 基本設定
" ==============================================================================
" {{{
"	set guifont=MS_Gothic:h10:cSHIFTJIS					" フォントサイズ設定（フォントサイズ設定は以下の「フォントサイズ設定」参照）
"	set columns=82										" ウィンドウの横幅を～カラムにします
"	set lines=35										" ウィンドウの高さを～行にします
"	set winwidth=1										" カレントウィンドウ 最小の幅
"	set winminwidth=1									" カレントウィンドウ以外 最小の幅
"	set winheight=1										" カレントウィンドウ 最小の高さ
"	set winminheight=1									" カレントウィンドウ以外 最小の高さ
	set equalalways										" ウィンドウサイズの自動調整を有効にする
	set guioptions-=m									" メニューバーを非表示
	set guioptions+=M									" $VIMRUNTIME/menu.vim を読み込まない
	set guioptions-=T									" ツールバーを非表示
"	set guioptions+=a									" ビジュアルモードで選択している箇所をクリップボードにコピー
	set guioptions+=r									" 縦スクロールバー表示
	set guioptions+=b									" 横スクロールバー表示
	set guioptions-=e									" gVimでもテキストベースのタブページを使う
	set number											" 行番号表示
	set foldcolumn=3									" 折り畳み行のネスト表示
"	set relativenumber									" 相対行番号表示(7.3)
	set ruler											" ルーラーを表示 (noruler:非表示)
	set cmdheight=2										" コマンドラインの高さ (gvimはgvimrcで指定)
	set laststatus=2									" コマンドをステータス行に表示
"	set statusline=										" ステータスライン表示設定
"	set statusline+=\ [BufNo:\%03.3n]					" ステータスライン表示設定 buffer number
"	set statusline+=\ [Func:%{cfi#get_func_name()}]		" ステータスライン表示設定 func name（タグジャンプに異常をきたすため、使用しない）
"	set statusline+=\ %F								" ステータスライン表示設定 filepath
"	set statusline+=\ <<%t>>							" ステータスライン表示設定 filename
"	set statusline+=\ %-40(\ %m%r%h%w\ %)				" ステータスライン表示設定 status flags
"	set statusline+=\ %=								" ステータスライン表示設定 right align remainder
"	set statusline+=\ [Frmt:%{&fenc!=''?&fenc:&enc}]	" ステータスライン表示設定 file format
"	set statusline+=\ [CR/LF:%{&ff}]					" ステータスライン表示設定 line feed
"	set statusline+=\ [Asc:0x\%02.2B]					" ステータスライン表示設定 ascii code (hex)
"	set statusline+=\ [SelRng:%{GetSelRngClmDiff()}x%{GetSelRngRowDiff()}]	" ステータスライン表示設定 select width x height
"	set statusline+=\ [XPos:%03v]						" ステータスライン表示設定 cursor position
"	set showtabline=2									" タブページを常に表示
"	set tabline=%!MakeTabLine()							" タブページのラベルを設定
"	set cursorcolumn									" カーソル列を目立たせる
"	set cursorline										" カーソル行を目立たせる
	set title											" タイトルを表示
"	set fileformat=dos									" 改行をWindowsの形式に変更。
	set scrolloff=5										" カーソル周辺行数
"	set swapfile										" スワップファイル(.swpファイル)を作成する
"	set directory=$VIM/__swapfiles						" スワップファイル(.swpファイル)の出力先を指定
	set noswapfile										" スワップファイル(.swpファイル)を作成しない
"	set undofile										" アンドゥファイル(.un~ファイル)を作成する
"	set undodir=$VIM/__undofiles						" アンドゥファイル(.un~ファイル)の出力先を指定
	set noundofile										" アンドゥファイル(.un~ファイル)を作成しない
"	set backup											" バックアップファイル(~ファイル)を作成する
"	set backupdir=$VIM/__backupfiles					" バックアップファイル(~ファイル)の出力先を指定
	set nobackup										" バックアップファイル(~ファイル)を作成しない
	set writebackup										" ファイルの上書きの前にバックアップを作る
														" set writebackupを指定してもオプション 'backup' がオンでない限り、
														" バックアップは上書きに成功した後に削除される。
if has('unix')
	set viminfo+=n$HOME/.vim/.viminfo					" VIMINFO ファイル出力先設定
else
	set viminfo+=n$VIM/_viminfo							" VIMINFO ファイル出力先設定
endif
	set hidden											" 編集結果非保存のバッファから、新しいバッファを開くときに警告を出さない
	set history=50										" ヒストリの保存数
	set textwidth=0										" 一行が長くなった場合の自動改行を抑止する
"	set formatoptions+=mM								" 日本語の行の連結時には空白を入力しない
	set formatoptions-=tc								" 一行が長くなった場合の自動改行を抑止する
"	set formatoptions+=q								" コメントを整形する
	set virtualedit=block								" Visual blockモードでフリーカーソルを有効にする
	set backspace=indent,eol,start						" バックスペースでインデントや改行を削除できるようにする
	set ambiwidth=double								" □や○の文字があってもカーソル位置がずれないようにする
	set wildmenu										" コマンドライン補完するときに強化されたものを使う
	set timeout timeoutlen=3000 ttimeoutlen=100			" キーコードやマッピングされたキー列が完了するのを待つ時間(ミリ秒)
	set clipboard+=unnamed								" クリップボードを共有
	set pastetoggle=<F3>								" <F3>でペーストモード切替え
	set nrformats-=octal								" <C-a>,<C-x> 実行にて 8 進数を無効にする。
	set nrformats-=alpha								" <C-a>,<C-x> 実行にて アルファベットを無効にする。
"	set browsedir=buffer								" ファイル保存ダイアログの初期ディレクトリをバッファファイル位置に設定
	set nocompatible									" vi互換をオフ
	set noshowcmd										" 選択中の行数、列数を表示しない（自作したので不要）（compatible 以降に記載すること）
	set shellslash										" Windowsでディレクトリパスの区切り文字表示に / を使えるようにする
	set showmatch										" 括弧の対応を数秒（0.1秒単位）表示
	set smarttab										" 行頭の余白内で tab を打ち込むと、'shiftwidth' の数だけインデントする。
	set autoindent										" 自動的にインデントする
	set smartindent										" 新しい行を作ったときに高度な自動インデントを行う
	set cinoptions+=:0									" Cインデントの設定
	set whichwrap=b,s,h,l,<,>,[,]						" カーソルを行頭、行末で止まらないようにする
	set tags+=tags;										" タグファイルのパス指定(カレントディレクトリからさかのぼってtagsファイルを検索)
	set nowrap											" 次のスクロールが可能。 zhで左へスクロール、zlで右へスクロール。 zHで左へ半分スクロール、zLで右へ半分スクロール
	set isfname+=32										" スペースがファイル名に入っていても、gfコマンドで開ける
	set list											" どの文字でタブや改行を表示するかを設定
	set listchars=tab:^\ ,eol:$							"	参考：set listchars=tab:>-,extends:<,trail:-,eol:<
	set nowrapscan										" 検索時にファイルの最後まで行ったら最初に戻らない
	set incsearch										" インクリメンタルサーチ
	set hlsearch										" 検索文字の強調表示
	set smartcase										" 大文字小文字の両方が含まれている場合は大文字小文字を区別
	set ignorecase										" 検索の時に大文字小文字を区別しない
	set iskeyword=a-z,A-Z,48-57,_						" w,bの移動で認識する文字
	set grepprg=internal								" vimgrep をデフォルトのgrepとする場合internal
"	set shortmess+=I									" スプラッシュ(起動時のメッセージ)を表示しない
	set noerrorbells									" エラー時の音とビジュアルベルの抑制(gvimは.gvimrcで設定)
	set visualbell t_vb=								" Beep 音を鳴らなくする
"	set lazyredraw										" マクロ実行中などの画面再描画を行わない
	set display=lastline								" Tab、行末の半角スペースを明示的に表示する
	set softtabstop=4									" ソフトタブストップ（<Tab> の挿入や <BS> の使用等の編集操作をするときに、<Tab> が対応する空白の数）
	set shiftwidth=4									" シフト移動幅（自動インデントやコマンド "<<", ">>" でずれる幅）
	set tabstop=4										" タブストップ（画面上でタブ文字が占める幅）
"	set expandtab										" タブを空白にする
	autocmd FileType cpp,hpp set softtabstop=2
	autocmd FileType cpp,hpp set tabstop=2
	autocmd FileType cpp,hpp set shiftwidth=2
	autocmd FileType cpp,hpp set expandtab
	autocmd FileType cpp,hpp set colorcolumn=100
if has('mouse')
	set mouse=a											" マウスを有効にする
endif
	set iminsert=0										" 挿入モードでのデフォルトのIME状態設定（IM オフ）
	set imsearch=0										" 検索モードでのデフォルトのIME状態設定（IM オフ）
	set encoding=utf-8									" vim内部エンコーディング変更
if has('unix')
	set fencs=ucs-bom,utf-8,shift-jis,euc-jp,default,latin1	"自動判別対象文字コード設定
endif
if has('unix')
	set shell=bash										" ターミナルウィンドウ(:terminal)のデフォルトシェルをbashにする
endif
"	set diffopt+=internal								" diff時に内部diffライブラリを使用する
"	set diffopt+=filler									" diff時に片方に行挿入されていた場合に片方に追加行を表示する
"	set diffopt+=algorithm:histogram					" diff時のアルゴリズムをhistogram差分アルゴリズムに変更する
"	set diffopt+=indent-heuristic						" diff時に内部diffライブラリのインデントヒューリスティックを使用する
"	set diffopt+=iwhite									" diff時に空白数の違いを無視する
"	set diffopt+=icase									" diff時に大文字小文字の違いを無視する
"try
"	set termkey=<c-n>									" ターミナルウィンドウのターミナルキー変更
"catch
"	set termwinkey=<c-n>								" ターミナルウィンドウのターミナルキー変更
"endtry
"	set ambiwidth=double
	set autoread										" ファイル自動読み込み
	set binary noeol									" 保存時に実施されるファイル末尾に対する改行コード自動付与を抑制する
	let g:vim_json_conceal = 0							" Jsonファイルのシンタックス非表示設定解除（参考:Vim/vim82/syntax/json.vim）
	let g:markdown_syntax_conceal = 0					" Markdownファイルのシンタックス非表示設定解除（参考:Vim/vim82/syntax/markdown.vim）
" }}}

" ==============================================================================
" ファイルパス設定
" ==============================================================================
"  " {{{
	let $CTAGS = expand( "$MYEXEPATH_CTAGS" )
	let $GTAGS = expand( "$MYEXEPATH_GTAGS" )
	let $GUIGREP = expand( "$MYEXEPATH_TRESGREP" )
if has('unix')
	let $MARKVIM = expand( "$HOME/.vim/_plugins_user/mark.vim/plugin/mark.vim" )
	let g:sSysPathDlmtr = '/'
	let g:sSysPathPrefix = '/'
else
	let $MARKVIM = expand( "$VIM/_plugins_user/mark.vim/plugin/mark.vim" )
	let g:sSysPathDlmtr = '\'
	let g:sSysPathPrefix = ''
endif
	let g:sProjectRootFileName = "tags"
" }}}

" ==============================================================================
" 共通ユーザ定義関数
" ==============================================================================
" {{{
	function! GetCurFilePath()
		"expand()のパス区切りは、環境に関わらず"/"固定。そのため環境に合わせてパス区切りを変える。
		return substitute( expand('%:p'), "/", g:sSysPathDlmtr, "g" )
	endfunction
	function! GetCurDirPath()
		let l:sCurDirPathTmp = expand("%:p")[ 0:( len( expand("%:p") ) - len( "/" . expand("%:t") ) - 1 ) ]
		"expand()のパス区切りは、環境に関わらず"/"固定。そのため環境に合わせてパス区切りを変える。
		return substitute( l:sCurDirPathTmp, "/", g:sSysPathDlmtr, "g" )
	endfunction
	function! GetCurFileName()
		return expand( '%:t' )
	endfunction
	function! GetCurFileExt()
		return expand("%:e")
	endfunction
	function! GetCurFileType()
		let l:sFileName = expand( '%:t' )
		let l:asFileTypes = ["if_l_mat\.c","if_c_mat\.c","if_c\.h","if\.h","_l_mat\.c","_c_mat\.c","_c\.h","\.h","\.c"]
		let l:sRetFileType = l:sFileName
		for l:sFileType in l:asFileTypes
			if l:sFileName =~ l:sFileType . "$"
				if l:sFileType[0] == "_"
					let l:sRetFileType = l:sFileType[1:]
				else
					let l:sRetFileType = l:sFileType
				endif
				break
			endif
		endfor
		return l:sRetFileType
	endfunction
	
	" === ファイルパス検索（＋存在確認） ====
	function! SrchStoreDirPathToTop( sTrgtDirPath, sSrchFileName )
		" === tags 存在確認 ====
		let l:asDirNames = split( a:sTrgtDirPath, g:sSysPathDlmtr )
		let l:sDirPath = ""
		let l:bIsExist = 0
		for l:iDirMaxCnt in range( len( l:asDirNames ) - 1, 1, -1 )
			for l:iDirCnt in range( 0, l:iDirMaxCnt )
				if l:iDirCnt == 0
					let l:sDirPath = l:asDirNames[ l:iDirCnt ]
				else
					let l:sDirPath = l:sDirPath . g:sSysPathDlmtr . l:asDirNames[ l:iDirCnt ]
				endif
			endfor
			let l:sDirPath = g:sSysPathPrefix . l:sDirPath
			if filereadable( l:sDirPath . g:sSysPathDlmtr . a:sSrchFileName )
				let l:bIsExist = 1
				break
			else
				" Do Nothing
			endif
		endfor
		
		if l:bIsExist == 1
			return l:sDirPath
		else
			return ""
		endif
	endfunction
" }}}

" ==============================================================================
" カラースキーマ設定
" ==============================================================================
" {{{
	syntax on
	colorscheme jellybeans
	highlight NonText		guibg=NONE	guifg=#404040
	highlight SpecialKey	guibg=NONE	guifg=#707070
	autocmd BufNewFile,BufRead *.cls  set filetype=vb
	autocmd BufNewFile,BufRead *.frm  set filetype=vb
" }}}

" ==============================================================================
" window位置の保存と復帰
" ==============================================================================
" {{{
if has('unix')
	"do nothing
else
	let g:vimposfilepath = '$VIM/_vimpos'
	
	function! s:savewindowparam(filename)
		redir => pos
		exec 'winpos'
		
		redir end
		let pos = matchstr(pos, '\cx[-0-9 ]\+,\s*y[-0-9 ]\+$')
		let file = expand(a:filename)
		let str = []
		let cmd = 'winpos '.substitute(pos, '[^-0-9 ]', '', 'g')
		cal add(str, cmd)
		let l = &lines
		let c = &columns
		cal add(str, 'set lines='. l.' columns='. c)
		silent! let ostr = readfile(file)
		if str != ostr
			call writefile(str, file)
		endif
	endfunction
	
	augroup savewindowparam
		autocmd!
		execute 'autocmd savewindowparam vimleave * call s:savewindowparam("'.g:vimposfilepath.'")'
	augroup end
	
	if filereadable(expand(g:vimposfilepath))
		execute 'source '. g:vimposfilepath
	endif
endif
" }}}

" ==============================================================================
" カーソル位置記憶
" 参考: https://qiita.com/yahihi/items/4112ab38b2cc80c91b16
" 備考: うまく動かない場合は、~/.viminfoのオーナーを変更する。
"         (e.g. sudo chown <user>:<group> ~/.viminfo)
" ==============================================================================
" {{{
if has('unix')
	augroup vimrcEx
		au BufRead * if line("'\"") > 0 && line("'\"") <= line("$") |
		\ exe "normal g`\"" | endif
	augroup END
endif
" }}}

" ==============================================================================
" 挿入モード時、ステータスラインの色を変更
" ==============================================================================
" {{{
	let g:hi_insert = 'highlight StatusLine guifg=darkblue guibg=darkyellow gui=none ctermfg=blue ctermbg=yellow cterm=none'
	
	if has('syntax')
		augroup InsertHook
			autocmd!
			autocmd InsertEnter * call s:StatusLine('Enter')
			autocmd InsertLeave * call s:StatusLine('Leave')
		augroup end
	endif
	
	let s:slhlcmd = ''
	function! s:StatusLine(mode)
		if a:mode == 'Enter'
			silent! let s:slhlcmd = 'highlight ' . s:GetHighlight('StatusLine')
			silent exec g:hi_insert
		else
			highlight clear StatusLine
			silent exec s:slhlcmd
		endif
	endfunction
	
	function! s:GetHighlight(hi)
		redir => hl
		exec 'highlight '.a:hi
		redir END
		let hl = substitute(hl, '[\r\n]', '', 'g')
		let hl = substitute(hl, 'xxx', '', '')
		return hl
	endfunction
" }}}

" ==============================================================================
" 全角スペースを表示
"	コメント以外で全角スペースを指定しているので、scriptencodingと、
"	このファイルのエンコードが一致するよう注意！
"	強調表示されない場合、ここでscriptencodingを指定するとうまくいく事があります。
" ==============================================================================
" {{{
	scriptencoding utf-8
	function! ZenkakuSpace()
		silent! let hi = s:GetHighlight('ZenkakuSpace')
		if hi =~ 'E411' || hi =~ 'cleared$'
			highlight ZenkakuSpace cterm=underline ctermfg=darkgrey gui=underline guifg=darkgrey
		endif
	endfunction
	if has('syntax')
		augroup ZenkakuSpace
			autocmd!
			autocmd ColorScheme		  * call ZenkakuSpace()
			autocmd VimEnter,WinEnter * match ZenkakuSpace /　/
		augroup end
		call ZenkakuSpace()
	endif
" }}}

" ==============================================================================
" カレントディレクトリ変更（grep,tags用）
" ==============================================================================
" {{{
	if exists('+autochdir')
		"autochdirがある場合カレントディレクトリを移動
		set autochdir
	else
		"autochdirが存在しないが、カレントディレクトリを移動したい場合
		au BufEnter * execute ":silent! lcd " . escape(expand("%:p:h"), ' ')
	endif
" }}}

" ==============================================================================
" 外部で変更のあったファイルを自動的に読み直す(ウィンドウを移動するたび)
" [参照] http://vim-users.jp/2011/03/hack206/
" ==============================================================================
" {{{
	augroup vimrc-checktime
		autocmd!
		autocmd WinEnter * checktime
	augroup end
" }}}

" ==============================================================================
" IME カラー設定
" [参照] http://sites.google.com/site/fudist/Home/vim-nihongo-ban/vim-color
" ==============================================================================
" {{{
	if has('multi_byte_ime')
		highlight Cursor	guifg=NONE guibg=Yellow
		highlight CursorIM	guifg=NONE guibg=Green
	endif
" }}}

" ==============================================================================
" コメント中の特定の単語を強調表示する
" [ハイライト実行例] TODO: NOTE: INFO: XXX: TEMP: FIXME: ASK: QUESTION:
" ==============================================================================
" {{{
	augroup HilightsForce
		autocmd!
		autocmd WinEnter,BufRead,BufNew,Syntax * :silent! call matchadd('Todo', '\v(TODO|NOTE|INFO|XXX|TEMP|FIXME|ASK|QUESTION)(\([a-zA-Z0-9_-]+\))?:')
		autocmd WinEnter,BufRead,BufNew,Syntax * highlight Todo guibg=Red guifg=White ctermbg=Red ctermfg=White
	augroup END
" }}}

" ==============================================================================
" 行末の空白をハイライトする
" ==============================================================================
" {{{
	let g:bHilightTailSpaces = 0
	if g:bHilightTailSpaces == 1
		augroup highlightSpace
			autocmd!
			autocmd WinEnter,BufRead,BufNew,Syntax * :silent! call matchadd('highlightSpaces', '\S\+\zs\s\+\ze$')
			autocmd WinEnter,BufRead,BufNew,Syntax * highlight highlightSpaces ctermbg=Red guibg=Red
		augroup END
	endif
" }}}

" ==============================================================================
" 関数名の色付け設定
" [参照] http://ogawa.s18.xrea.com/tdiary/20070523.html
" ==============================================================================
" {{{
"	autocmd FileType ruby,c,cpp syntax match CFunction /\v[a-zA-Z_]\w*\s*!*((\[[^]]*\]\s*)?\(\s*[^\*])@=/
	autocmd FileType ruby,c,cpp syntax match CFunction /\v[a-zA-Z_]\w*\s*!*((\[[^]]*\]\s*)?\(\s*)@=/
	autocmd FileType ruby,c,cpp hi CFunction guifg=orange ctermfg=216
" }}}

" ==============================================================================
" Vim 内Grep の設定
" [grep option]
"	-i：大文字小文字を無視する
"	-E：拡張正規表現で検索を行う
" [参照] https://eng-entrance.com/linux-command-grep#-i
" ==============================================================================
" {{{
	let g:sGrepWord = ""
	let g:sGrepOpt = "-i"
"	let g:sGrepFileExt = "*.vim,*.c,*.h,*.C,*.H,*.cc,*.cpp,*.hpp,*.f,*.f90,*.ff90,*.F,*.F90,*.vbs,*.bas,*.cls,*.py,*.swift"
	let g:sGrepFileExt = "*"
	let g:sGrepPath = "" "パス区切りは'/'で指定すること！
	
	"------------------------------------------------------
	"現在のファイルに対して Grep (Grep Current)
	" :Gc <word>
	"------------------------------------------------------
	command! -nargs=? Gc call s:GrepAtCurFile(<f-args>)
	function! s:GrepAtCurFile(...)
		"### 対象単語 取得 ###
		if a:0 >= 1
			let g:sGrepWord = a:000[(a:0 - 1)]	"引数の単語で検索
		else
			let g:sGrepWord = expand("<cword>") "カーソル上の単語で検索
		endif
		
		"### 対象パス 取得 ###
		let l:sRootPath = GetCurFilePath()
		
		"### Grep 実行 ###
		execute "Grep " . g:sGrepOpt . " """ . g:sGrepWord . """ " . l:sRootPath
	endfunction
	
	"------------------------------------------------------
	"上位階層のルート識別ファイル格納ディレクトリをルートとして Grep (Grep Project)
	" :Gp <word> | :Gpa <word> | :Gpb <word> | :Gpc <word>
	"------------------------------------------------------
	command! -nargs=? Gp let s:sGrepRootPath = g:sProjectRootFileName | call s:GrepAtPrj(<f-args>)
	command! -nargs=? Gpa let s:sGrepRootPath = ".greproota" | call s:GrepAtPrj(<f-args>)
	command! -nargs=? Gpb let s:sGrepRootPath = ".greprootb" | call s:GrepAtPrj(<f-args>)
	command! -nargs=? Gpc let s:sGrepRootPath = ".greprootc" | call s:GrepAtPrj(<f-args>)
	function! s:GrepAtPrj(...)
		"### 対象単語 取得 ###
		if a:0 >= 1
			let g:sGrepWord = a:000[(a:0 - 1)]	"引数の単語で検索
		else
			let g:sGrepWord = expand("<cword>") "カーソル上の単語で検索
		endif
		
		"### 対象パス 取得 ###
		if g:sGrepPath == ""
			let l:sRootPath = SrchStoreDirPathToTop( GetCurDirPath(), s:sGrepRootPath )
			if l:sRootPath == ""
				let l:sRootPath = GetCurDirPath()
			endif
		else
			let l:sRootPath = substitute( g:sGrepPath, "/", g:sSysPathDlmtr, "g" )
		endif
		
		"### 拡張子 取得 ###
		let l:asFileTypes = split( g:sGrepFileExt, "," )
		let l:sFileTypeOpt	= ""
		for l:sFileType in l:asFileTypes
			if l:sFileTypeOpt == ""
				let l:sFileTypeOpt = "--include=" . l:sFileType
			else
				let l:sFileTypeOpt = l:sFileTypeOpt . " " . "--include=" . l:sFileType
			endif
		endfor
		
		"### Grep 実行 ###
	"	echom   "Grep -r " . g:sGrepOpt . " " . l:sFileTypeOpt . " """. g:sGrepWord . """ " . l:sRootPath . "/**"
		execute "Grep -r " . g:sGrepOpt . " " . l:sFileTypeOpt . " """. g:sGrepWord . """ " . l:sRootPath . "/**"
	endfunction
	
	"------------------------------------------------------
	"ファイルパスなどを指定して対話型 Grep (Grep Intaractive)
	" :Gi
	"------------------------------------------------------
	command! Gi call s:GrepAtInteractive()
	function! s:GrepAtInteractive()
		"### 対象単語 取得 ###
		let l:sGrepWord = input("grep target text : ", expand('<cword>'))
		if l:sGrepWord == ""
			echo "\n"
			echo "  grep target word is not specified!"
			echo "  this process is suspended!"
			return
		endif
		let g:sGrepWord = l:sGrepWord
		
		"### 対象パス 取得 ###
		"初期値判定
		if g:sGrepPath == ""
			let l:sRootPath = SrchStoreDirPathToTop( GetCurDirPath(), g:sProjectRootFileName )
			if l:sRootPath == ""
				let l:sRootPath = GetCurDirPath()
			endif
		else
			let l:sRootPath = g:sGrepPath
		endif
		"パス入力
		let l:sRootPath = input("grep target path : ", l:sRootPath)
		if l:sRootPath == ""
			echo "\n"
			echo "  grep target path is not specified!"
			echo "  this process is suspended!"
			return
		endif
		let l:sRootPath = substitute( l:sRootPath, "/", g:sSysPathDlmtr, "g" )
		let g:sGrepPath = l:sRootPath
		
		"### 拡張子 取得 ###
		let l:sGrepFileExt = input("grep target file ext(ex. *.c,*.h) : ", g:sGrepFileExt)
		if l:sGrepFileExt == ""
			echo "\n"
			echo "  grep target file extention is not specified!"
			echo "  this process is suspended!"
			return
		endif
		let g:sGrepFileExt = l:sGrepFileExt
		let l:asFileTypes = split( l:sGrepFileExt, "," )
		let l:sFileTypeOpt	= ""
		for l:sFileType in l:asFileTypes
			if l:sFileTypeOpt == ""
				let l:sFileTypeOpt = "--include=" . l:sFileType
			else
				let l:sFileTypeOpt = l:sFileTypeOpt . " " . "--include=" . l:sFileType
			endif
		endfor
		
		"### grepオプション 取得 ###
		let l:sGrepOpt = input("grep options : ", g:sGrepOpt)
		let g:sGrepOpt = l:sGrepOpt
		
		"### Grep 実行 ###
		execute "Grep -r " . l:sGrepOpt . " " . l:sFileTypeOpt . " """. l:sGrepWord . """ " . l:sRootPath . "/**"
	endfunction
" }}}

" ==============================================================================
" GUIのGrepソフトでGrepする
" [参照] http://thinca.hatenablog.com/entry/20111204/1322932585
" ==============================================================================
" {{{
	function! ExecuteGuiGrep()
		if has('unix')
			echo "[error] ExecuteGuiGrep() is windows only!"
		else
			let l:sGuiGrepSoftPath = substitute( $GUIGREP, "/", g:sSysPathDlmtr, "g" )
			let l:sRootDirPath = SrchStoreDirPathToTop( GetCurDirPath(), g:sProjectRootFileName )
			if l:sRootDirPath == ""
				let l:sRootDirPath = GetCurDirPath()
			endif
			let l:sSearchKeyword = expand('<cword>')
		"	execute "!cmd.exe /c start /max " . l:sGuiGrepSoftPath . " " . l:sRootDirPath . " /KEYWORD=" . l:sSearchKeyword
		"	execute "!cmd.exe /c start /b /max " . l:sGuiGrepSoftPath . " " . l:sRootDirPath . " /KEYWORD=" . l:sSearchKeyword
			execute "!start " . l:sGuiGrepSoftPath . " " . l:sRootDirPath . " /KEYWORD=""" . l:sSearchKeyword . """"
		endif
	endfunction
" }}}

" ==============================================================================
" MemoFile 書込設定
" Usage : :Cm 実行でデスクトップ配下に
"		  temp_XXX.txt を作成する
" ==============================================================================
" {{{
	command! Cm call CreateMemoFile()
	function! CreateMemoFile()
		let l:sFileNameIdx = 1
		let l:sMemoFilePath = 0
		while 1
			let l:sMemoFilePath = (expand('~/Desktop/temp_') . printf("%03d", l:sFileNameIdx) . '.txt')
			if filereadable(l:sMemoFilePath) == 1
				let l:sFileNameIdx = l:sFileNameIdx + 1
			else
				break
			endif
		endwhile
		execute "w " . l:sMemoFilePath
	endfunction
" }}}

" ==============================================================================
" 上位階層にあるタグファイルを探して、更新。
" 注意 : 事前にタグファイルを作成しておくこと。
" ==============================================================================
" {{{
	command! Utf call UpdateTagFile()
	function! UpdateTagFile()
		" === 上位階層ディレクトリ tags 存在確認 ====
		let l:sDirPath = SrchStoreDirPathToTop( GetCurDirPath(), g:sProjectRootFileName )
		
		" === tags ファイル更新 ====
		if l:sDirPath == ""
			echo "update tag file error!! tag file is missing ..."
		else
			if has('unix')
				execute "cd " . l:sDirPath
				execute "!ctags -R"
			else
				execute "cd " . l:sDirPath
				execute "!start " . $CTAGS . " -R"
				execute "!start " . $GTAGS . " -v"
				execute "redraw"
			endif
			echo "update tag file!!  " . l:sDirPath
		endif
	endfunction
" }}}

" ==============================================================================
" フォントサイズ設定
" ==============================================================================
" {{{
	let g:FontSizeLevel = 3
	let s:aiFontSizeList = [ 2, 8, 10, 11, 13, 16 ] "要素番号0は俯瞰モード用
	let s:bIsBirdEyesMode = 0
	
	"フォントサイズ更新
	function! UpdateFontSize()
		if has('unix')
			"do nothing
		else
			execute "set guifont=MS_Gothic:h" . s:aiFontSizeList[g:FontSizeLevel] . ":cSHIFTJIS"
			execute 'source '. g:vimposfilepath
		endif
	endfunction
	call UpdateFontSize() "初回読み込み時のフォントサイズ
	
	" フォントサイズをトグルする。
	function! SwitchFontSize()
		if g:FontSizeLevel < (len(s:aiFontSizeList) - 1)
			let g:FontSizeLevel = g:FontSizeLevel + 1
		else
			let g:FontSizeLevel = 1
		endif
		let g:FontSizeLevelOld = g:FontSizeLevel
		call UpdateFontSize()
	endfunction
	
	" 俯瞰モード
	function! SwitchBirdEyesViewMode()
		if g:bIsBirdEyesMode == 0
			let g:FontSizeLevelOld = g:FontSizeLevel
			let g:FontSizeLevel = 0
			let g:bIsBirdEyesMode = 1
		else
			let g:FontSizeLevel = g:FontSizeLevelOld
			let g:bIsBirdEyesMode = 0
		endif
		call UpdateFontSize()
	endfunction
" }}}

" ==============================================================================
" ファイル保存時に「タブ」を「空白」に変換するかを選択する
" ==============================================================================
" {{{
	let g:EnableTab2SpaceAtSave = 0
	autocmd BufWritePre * call Tab2SpaceAtSave()
	function! Tab2SpaceAtSave()
		if g:EnableTab2SpaceAtSave == 1
			set expandtab
			retab!
		else
			"Do Nothing
		endif
	endfunction
" }}}

" ==============================================================================
" 現在開いているファイル名・ファイルパスをクリップボードにコピーする
" 参考：copypath.vim v1.0
"		http://nanasi.jp/articles/vim/copypathim.html
" ==============================================================================
" {{{
	let g:lCopy2UnnamedRegister = 1 " 1）無名レジスタ＋＊レジスタにコピー、それ以外)＊レジスタにコピー
	
	command! Cfp call CopyCurFilePath()
	function! CopyCurFilePath()
		let l:sCurFilePath = GetCurFilePath()
		call SetToClipboard(l:sCurFilePath)
		echo "copy file path        : " . l:sCurFilePath
	endfunction
	
	command! Crp call CopyCurPrjRltvFilePath()
	function! CopyCurPrjRltvFilePath()
		"substitute関数は"\"のエスケープが必要なため、一旦"/"に置換してから処理する。
		let l:sCurFilePath = GetCurFilePath()
		let l:sCurFilePath = substitute( l:sCurFilePath, "\\", "/", "g" )
		let l:sPrjRootPath = SrchStoreDirPathToTop( GetCurDirPath(), g:sProjectRootFileName )
		if l:sPrjRootPath != ""
			let l:sPrjRootPath = l:sPrjRootPath . g:sSysPathDlmtr
		endif
		let l:sPrjRootPath = substitute( l:sPrjRootPath, "\\", "/", "g" )
		let l:sCurPrjRltvFilePath = substitute( l:sCurFilePath, l:sPrjRootPath, "", "g" )
		let l:sCurPrjRltvFilePath = substitute( l:sCurPrjRltvFilePath, "/", g:sSysPathDlmtr, "g" )
		call SetToClipboard(l:sCurPrjRltvFilePath)
		echo "copy rltv file path   : " . l:sCurPrjRltvFilePath
	endfunction
	
	command! Cdp call CopyCurDirPath()
	function! CopyCurDirPath()
		let l:sCurDirPath = GetCurDirPath()
		call SetToClipboard(l:sCurDirPath)
		echo "copy directory path   : " . l:sCurDirPath
	endfunction
	
	command! Cfn call CopyCurFileName()
	function! CopyCurFileName()
		let l:sCurFileName = GetCurFileName()
		call SetToClipboard(l:sCurFileName)
		echo "copy file name        : " . l:sCurFileName
	endfunction
	
	command! Cfe call CopyCurFileExt()
	function! CopyCurFileExt()
		let l:sCurFileExt = GetCurFileExt()
		call SetToClipboard(l:sCurFileExt)
		echo "copy file extention   : " . l:sCurFileExt
	endfunction
	
	command! Cpn call CopyProgramNo()
	function! CopyProgramNo()
		let l:sGetLine = getline(1)
		let l:sPrgNoTmp = l:sGetLine
		let l:sPrgNoTmp = substitute( l:sPrgNoTmp, "^\/\\* ", "", "g" )
		let l:sPrgNoTmp = substitute( l:sPrgNoTmp, " \\*\/$", "", "g" )
		let l:sPrgNo = l:sPrgNoTmp
		call SetToClipboard(l:sPrgNo)
		echo "copy program no       : " . l:sPrgNo
	endfunction
	
	command! Cln call CopyFileLineNo()
	function! CopyFileLineNo()
		let l:sLineNo = g:sSelRowsRng
		call SetToClipboard(l:sLineNo)
		echo "copy line no          : " . l:sLineNo
	endfunction
	
	command! Cel call CopyFileExtAndLineNo()
	function! CopyFileExtAndLineNo()
		let l:sCurFileName = GetCurFileType()
		"let l:sLineNo = line(".")
		let l:sLineNo = g:sSelRowsRng
		let l:sFileLineNo = l:sCurFileName . "/" . l:sLineNo
		call SetToClipboard(l:sFileLineNo)
		echo "copy fileext & line   : " . l:sFileLineNo
	endfunction
	
	command! Cpl call CopyRltvFilePathAndLineNo()
	function! CopyRltvFilePathAndLineNo()
		"substitute関数は"\"のエスケープが必要なため、一旦"/"に置換してから処理する。
		let l:sCurFilePath = GetCurFilePath()
		let l:sCurFilePath = substitute( l:sCurFilePath, "\\", "/", "g" )
		let l:sPrjRootPath = SrchStoreDirPathToTop( GetCurDirPath(), g:sProjectRootFileName )
		if l:sPrjRootPath != ""
			let l:sPrjRootPath = l:sPrjRootPath . g:sSysPathDlmtr
		endif
		let l:sPrjRootPath = substitute( l:sPrjRootPath, "\\", "/", "g" )
		let l:sCurFileName = substitute( l:sCurFilePath, l:sPrjRootPath, "", "g" )
		let l:sCurFileName = substitute( l:sCurFileName, "/", g:sSysPathDlmtr, "g" )
		"let l:sLineNo = line(".")
		let l:sLineNo = g:sSelRowsRng
		let l:sFileLineNo = l:sCurFileName . ":" . l:sLineNo
		call SetToClipboard(l:sFileLineNo)
		echo "copy rfilepath & line : " . l:sFileLineNo
	endfunction
	
"	command! Cll call CopyLineWithLineNo()
	function! CopyLineWithLineNo()
		normal gv
		let l:lRow1 = line(".")
		normal o
		let l:lRow2 = line(".")
		if l:lRow1 < l:lRow2
			let l:lStartRow = l:lRow1
			let l:lEndRow = l:lRow2
		else
			let l:lStartRow = l:lRow2
			let l:lEndRow = l:lRow1
		endif
		let l:lMaxDigit = strlen(string(l:lEndRow))
		let l:sCopyStr = ""
		for l:lCurRow in range( l:lStartRow, l:lEndRow, 1 )
			let l:sAddStr = printf(" %". l:lMaxDigit ."d %s", l:lCurRow, getline(l:lCurRow))
			if l:sCopyStr == ""
				let l:sCopyStr = l:sAddStr
			else
				let l:sCopyStr = l:sCopyStr . "\n" . l:sAddStr
			endif
		endfor
	"	echom l:lStartRow . " " . l:lEndRow
		redraw
		call SetToClipboard(l:sCopyStr)
		exec "normal \<esc>"
		"echo "copy line with lineno"
	endfunction
	
	command! Cpk call CopyCurRos2PkgName()
	function! CopyCurRos2PkgName()
		"substitute関数は"\"のエスケープが必要なため、一旦"/"に置換してから処理する。
		let l:sCurFilePath = GetCurFilePath()
		let l:sCurFilePath = substitute( l:sCurFilePath, "\\", "/", "g" )
		let l:sPkgRootPath = SrchStoreDirPathToTop( GetCurDirPath(), "package.xml" )
		if l:sPkgRootPath == ""
			echom "[error] package.xml is missing."
			return
		endif
		let l:sPkgFilePath = l:sPkgRootPath . g:sSysPathDlmtr . "package.xml"
		let l:sPkgName = ""
		let l:sPattern = '\v.*\<name\>(.*)\<\/name\>'
		for l:sLine in readfile(l:sPkgFilePath)
			let l:lsMatch = matchlist(l:sLine, l:sPattern)
			if !empty(l:lsMatch)
				let l:sPkgName = l:lsMatch[1]
			endif
		endfor
		call SetToClipboard(l:sPkgName)
		echo "copy ros2 package name : " . l:sPkgName
	endfunction
	
	function! GetFromClipboard()
		if g:lCopy2UnnamedRegister == 1
		"	let l:sClipbordStr = @* " * register
			let l:sClipbordStr = @" " unnamed register
		else
			let l:sClipbordStr = @* " * register
		endif
		return l:sClipbordStr
	endfunction
	function! SetToClipboard(sString)
		if has('unix')
			call SendViaOSC52(a:sString)
		else
			if g:lCopy2UnnamedRegister == 1
				let @* = a:sString " * register.
				let @" = a:sString " unnamed register.
			else
				let @* = a:sString " * register.
			endif
		endif
	endfunction
" }}}

" ==============================================================================
" 選択したファイルパスを、現在のファイルパスからの相対パスへ置き換える。
" [使い方例]
"	現在絶対パス:	c:\test\aaa\bbb\test.txt
"	選択絶対パス:	c:\test\ccc\test2.txt
"	出力パス:		..\..\ccc\test2.txt
" [注意事項]
"	・ヴィジュアルモードで実行すること！
"	   ex) vnoremap <silent> <F9> :call ReplaceRelativePathFromCurrent()<cr>
" ==============================================================================
" {{{
	let g:sOutPathDlmtr = '/'
	command! Rrp call ReplaceRelativePathFromCurrent()
	function! ReplaceRelativePathFromCurrent()
		if has('unix')
			echo "ReplaceRelativePathFromCurrent() is not supported on linux!"
		else
			"選択文字列 取得
			let tmp = @@
			silent normal gvy
			let l:sDstPath = @@
			let l:sSrcPath = GetCurFilePath()
			echo l:sDstPath
			
			"選択文字列 削除
			silent normal gv
			silent normal d
			
			"相対パス取得
			let l:sRelativePath = ConvRelativePath( l:sSrcPath, l:sDstPath, g:sOutPathDlmtr )
			
			"選択文字列 出力
			let @* = l:sRelativePath
			silent normal p
			
			echo "replace relative path : success!"
		endif
	endfunction
	
	function! ConvRelativePath( sSrcPath, sDstPath, sDlmtr )
		"substitute関数は"\"のエスケープが必要なため、一旦"/"に置換してから処理する。
		let l:sSrcPath = substitute( a:sSrcPath, "\\", "/", "g" )
		let l:sDstPath = substitute( a:sDstPath, "\\", "/", "g" )
		"echom l:sSrcPath . "★" . l:sDstPath
		
		"一致パスを削除
		let l:iChrIdx = 1
		while l:iChrIdx <= len( l:sSrcPath )
			if l:sSrcPath[l:iChrIdx] == '/'
				"echom l:sSrcPath . "★" . l:sDstPath
				if l:sSrcPath[:l:iChrIdx] == l:sDstPath[:l:iChrIdx]
					let l:sSrcPath = l:sSrcPath[l:iChrIdx+1:]
					let l:sDstPath = l:sDstPath[l:iChrIdx+1:]
					let l:iChrIdx = 1
				else
					let l:iChrIdx = l:iChrIdx + 1
				endif
			else
				let l:iChrIdx = l:iChrIdx + 1
			endif
		endwhile
		"echom l:sSrcPath . "★" . l:sDstPath
		
		"さかのぼりパス取得
		let l:iPathDepth = (len(l:sSrcPath) - len(substitute( l:sSrcPath, "/", "", "g" )))
		"echom l:iPathDepth
		let l:sRiseUpPath = ""
		if l:iPathDepth == 0
			let l:sRiseUpPath = "./"
		else
			for l:iIdx in range(1,l:iPathDepth)
				let l:sRiseUpPath = l:sRiseUpPath . "../"
			endfor
		endif
		
		"パス置換("/"から指定された区切り文字へ)
		let l:sRelativePath = l:sRiseUpPath . l:sDstPath
		let l:sRelativePath = substitute( l:sRelativePath, '/', a:sDlmtr, "g" )
		
		return l:sRelativePath
	endfunction
" }}}

" ==============================================================================
" 現在のスクリプトを実行する
"	<<デフォルト時>>
"	  :!%
"	<<「cmdpre = "ruby"」とした場合>>
"	  :!ruby %
"	<<「cmdpst = "-v"」とした場合>>
"	  :!% -v
" ==============================================================================
" {{{
	let g:cmdpre = ""
	let g:cmdpst = ""
	function! ExecCurrentScript()
		let l:sCurFileExt = GetCurFileExt()
		let l:sCurFileName = GetCurFileName()
		let l:sExecCmd = ""
		if stridx( l:sCurFileName, "vimrc" ) > 0
			let l:sExecCmd = "so %"
		else
			if l:sCurFileExt == "vim"
				let l:sExecCmd = "so %"
			elseif l:sCurFileExt == "py"
				let l:sExecCmd = "!python % " . g:cmdpst
			elseif l:sCurFileExt == "rb"
				let l:sExecCmd = "!ruby % " . g:cmdpst
			else
				let l:sExecCmd = "!" . g:cmdpre . " % " . g:cmdpst
			endif
		endif
		return l:sExecCmd
	endfunction
" }}}

" ==============================================================================
" ウィンドウタブ機能無効化
" ==============================================================================
" {{{
	autocmd! BufNewFile,BufRead,BufEnter,BufNew,BufWinEnter * call AutoTabPageOnly()
	let g:TabPageOnlyEnable = 1
	function! AutoTabPageOnly()
		if g:TabPageOnlyEnable == 1
			if tabpagenr() == 1
				"do nothing
			else
				execute ("tabonly")
			endif
		else
			"do nothing
		endif
	endfunction
" }}}

" ==============================================================================
" ウィンドウタブ表示を変更する
" [参考] http://thinca.hatenablog.com/entry/20111204/1322932585
" ==============================================================================
" {{{
"	function! MakeTabLine()
"		let l:asTitles = map(range(1, tabpagenr('$')), 'GetTabPageLabel(v:val)')
"		let l:sDelimiter = ' '	" タブ間の区切り
"		let l:sTabpages = join(l:asTitles, l:sDelimiter) . l:sDelimiter . '%#TabLineFill#%T'
"		let l:sInfo = ''  " 好きな情報を入れる
"		return l:sTabpages . '%=' . l:sInfo  " タブリストを左に、情報を右に表示
"	endfunction
"	" 各タブページのカレントバッファ名+αを表示
"	function! GetTabPageLabel( lTabPageNum )
"		" t:title と言う変数があったらそれを使う
"		let l:sTitle = gettabvar(a:lTabPageNum, 'l:sTitle')
"		if l:sTitle !=# ''
"			return l:sTitle
"		endif
"		
"		" タブページ内のバッファのリスト
"		let l:asBufList = tabpagebuflist(a:lTabPageNum)
"		
"		" カレントタブページかどうかでハイライトを切り替える
"		let l:sHilight = a:lTabPageNum is tabpagenr() ? '%#TabLineSel#' : '%#TabLine#'
"		
"		" バッファが複数あったらバッファ数を表示
"		let l:sBufNum = ''
"	"	let l:sBufNum = len(l:asBufList)
"	"	if l:sBufNum is 1
"	"		let l:sBufNum = ''
"	"	endif
"		
"		" タブページ内に変更ありのバッファがあったら '+' を付ける
"		let l:sModifyStat = len(filter(copy(l:asBufList), 'getbufvar(v:val, "&modified")')) ? '+' : ''
"		let l:sSpace = (l:sBufNum . l:sModifyStat) ==# '' ? '' : ' '  " 隙間空ける
"		
"		" カレントバッファ
"		let l:lCurBufIdx = l:asBufList[tabpagewinnr(a:lTabPageNum) - 1]  " tabpagewinnr() は 1 origin
"		let l:sCurBufAbsPath = bufname(l:lCurBufIdx)
"		if l:sCurBufAbsPath == ""
"			let l:sBufName = "(無題)"
"		else
"			if stridx( l:sCurBufAbsPath, '/' ) == -1
"				let l:sBufName = l:sCurBufAbsPath
"			else
"				let l:sBufName = strpart( l:sCurBufAbsPath, strridx( l:sCurBufAbsPath, '/') + 1 )
"			endif
"		endif
"		
"		return '%' . a:lTabPageNum . 'T' . l:sHilight . ' ' . l:sBufNum . l:sModifyStat . l:sSpace . l:sBufName . ' ' . '%T%#TabLineFill#'
"	endfunction
" }}}

" ==============================================================================
" 現在ファイル削除コマンド
" ==============================================================================
" {{{
	command! Delme call DeleteCurFile()
	function! DeleteCurFile()
		if has('unix')
			execute "!rm """ . GetCurFilePath() . """"
		else
			execute "!del """ . GetCurFilePath() . """"
		endif
	endfunction
" }}}

" ==============================================================================
" ウィンドウサイズ最大化
" [参考] http://d.hatena.ne.jp/akishin999/20090509/1241855699
" ==============================================================================
" {{{
	"au GUIEnter * simalt ~x
" }}}

" ==============================================================================
" 終了時 タブ⇔空白 自動置換
" 
" ★不具合ありのため、動作しない！
" AutoRepTabSpace()は動作するが、vim 終了時に実行してくれない…
" vim 終了時に実行される autocmd を要調査
" ==============================================================================
" {{{
	let g:AutoRepTabSpaceEnable = 0
	let g:AutoRepTabSpaceType = 0 "1:tab 2:space other:keep
	let g:AutoRepTabSpaceExt = 'c|h'
	
	if g:AutoRepTabSpaceEnable == 1
		autocmd BufWipeout * call AutoRepTabSpace() 
	"	execute 'autocmd! BufUnload ' . g:AutoRepTabSpaceExt ' call AutoRepTabSpace()'
	endif
	function! AutoRepTabSpace()
		"  <<本関数内で拡張子を判別する理由>>
		"	 「autocmd BufDelete *.h,*.c call AutoRepTabSpace()」でも拡張子を
		"	 指定して実行することができるが、たとえば「a.c」「a.c」「a.txt」を
		"	 読み込んだ上で「a.txt」を開き":q"を実行すると、終了時「a.txt」が
		"	 開かれているため、上記 autocmd は実行されない。
		"	 そのため、autocmdは常時実行できるよう拡張子を "*" で指定しておき、
		"	 本関数内で拡張子を判別する必要がある。
	"	redir! > redir.txt
		for l:iBufIdx in range( 1, bufnr("$") )
			if bufexists(l:iBufIdx)
				if bufname(l:iBufIdx) =~ '\.[' . g:AutoRepTabSpaceExt . ']$'
				"	echo l:iBufIdx . " " . bufname(l:iBufIdx)
					if g:AutoRepTabSpaceType == 1 "space => tab
						execute( l:iBufIdx . 'bufdo set noexpandtab' )
						execute( l:iBufIdx . 'bufdo retab!' )
						execute( l:iBufIdx . 'bufdo w' )
					elseif g:AutoRepTabSpaceType == 2 "tab => space
						execute( l:iBufIdx . 'bufdo set expandtab' )
						execute( l:iBufIdx . 'bufdo retab!' )
						execute( l:iBufIdx . 'bufdo w' )
					else
						"Do Nothing
					endif
				else
					"Do Nothing
				endif
			else
				"Do Nothing
			endif
		endfor
	"	redir END
	endfunction
	
"	"BufDelete の実行タイミング調査用コード
"	autocmd BufDelete * call TestFunc()
"	function! TestFunc()
"		redir! >> C:\Users\draem_000\Desktop\test2\redir.txt
"		echo "exec"
"		redir END
"	endfunction
" }}}

" ==============================================================================
" 終了時 改行コード 自動置換
" ★「終了時 タブ⇔空白 自動置換」と同様の理由で動作しない！
" ==============================================================================
" {{{
"	let g:AutoRepNewLineCodeEnable = 0
"	let g:AutoRepNewLineCodeType = 2 "0:Lf(unix)、1:Cr(mac)、Other:CrLf(dos)
"	let g:AutoRepNewLineCodeExt = '*.c,*.h'
"	
"	if g:AutoRepNewLineCodeEnable == 1
"		autocmd BufDelete *.c,*.h call AutoRepNewLineCode()
"	"	execute 'autocmd savewindowparam vimleave * call s:savewindowparam("'.g:vimposfilepath.'")'
"	"	execute 'autocmd! BufWinLeave ' . g:AutoRepNewLineCodeExt ' call AutoRepNewLineCode()'
"	endif
"	function! AutoRepNewLineCode()
"		if g:AutoRepNewLineCodeType == 0
"			set fileformat=unix
"		elseif g:AutoRepNewLineCodeType == 1
"			set fileformat=mac
"		else
"			set fileformat=dos
"		endif
"		execute 'w'
"	endfunction
" }}}

" ==============================================================================
" 文字コード/改行コード 再オープン
" Usage
"   :Reoenc [euc-jp|shift_jis|utf-8|..]
"   :Reoff [dos|mac|unix]
" ==============================================================================
" {{{
	command! -nargs=1 Reoenc call s:ReOpenAtNewCharCode(<f-args>)
	function! s:ReOpenAtNewCharCode(...)
		if a:0 == 1
			execute 'e ++enc=' . a:1
		endif
	endfunction
	
	command! -nargs=1 Reoff call s:ReOpenAtNewLineCode(<f-args>)
	function! s:ReOpenAtNewLineCode(...)
		if a:0 == 1
			execute 'e ++ff=' . a:1
		endif
	endfunction
" }}}

" ==============================================================================
" 文字コード/改行コード置換
" Usage
"   :Repenc [euc-jp|shift_jis|utf-8|..]
"   :Repff [dos|mac|unix]
" ==============================================================================
" {{{
	command! -nargs=? Repenc call s:ReplaceCharCode(<f-args>)
	function! s:ReplaceCharCode(...)
		if a:0 == 1
			execute 'set fenc=' . a:1
		endif
	endfunction
	
	command! -nargs=? Repff call s:ReplaceNewlineCode(<f-args>)
	function! s:ReplaceNewlineCode(...)
		if a:0 == 1
			execute 'set ff=' . a:1
		endif
	endfunction
" }}}

" ==============================================================================
" ウィンドウサイズトグル
" [参考] https://qiita.com/grohiro/items/e3dbcc93510bc8c4c812
" ==============================================================================
" {{{
	let g:toggle_window_size = 0
	function! ToggleWindowSize()
		if g:toggle_window_size == 1
			exec "normal \<C-w>="
			let g:toggle_window_size = 0
		else
			:resize
			:vertical resize
			let g:toggle_window_size = 1
		endif
	endfunction
" }}}

" ==============================================================================
" 辞書ファイル登録
" [参考] https://nanasi.jp/articles/howto/config/dictionary.html
" ==============================================================================
" {{{
if has('unix')
	autocmd FileType vb :set dictionary=$HOME/.vim/_dictionary/vbscript.dict
else
	autocmd FileType vb :set dictionary=$VIM/_dictionary/vbscript.dict
endif
" }}}

" ==============================================================================
" INSERT mode に入るときにカーソル形状を変える
" [参考] https://oki2a24.com/2019/02/19/how-to-set-terminal-vim-cursor-in-vimrc-as-i-leraned-from-mintty-wiki-tips/
" ==============================================================================
" {{{
if has('unix')
	let &t_ti .= "\e[2 q"	" [Vim 起動時]		 非点滅ブロック
	let &t_SI .= "\e[6 q"	" [挿入モード時]	 非点滅縦棒
	let &t_EI .= "\e[2 q"	" [ノーマルモード時] 非点滅ブロック
	let &t_te .= "\e[0 q"	" [vim 終了時]		 デフォルト
endif
" }}}

" ==============================================================================
" "{"と","と"}"区切りの選択文字列に対してインデント整形する
" ==============================================================================
" {{{
	command! -range Aibr call AutoIndentBrackets()
	function! AutoIndentBrackets()
		"選択文字列取得
		silent normal gvd
		let l:inputstr = @*
		
		"文字列整形
		let l:outputstr = ""
		let l:tablevel = 0
		let l:idx = 0
		let l:isStrMode = 0
		while l:idx < strlen(l:inputstr)
			"文字列モードの場合
			if l:isStrMode == 1
				if l:inputstr[l:idx] =~ "\""
					let l:isStrMode = 0
				endif
				let l:outputstr = l:outputstr . l:inputstr[l:idx]
			"文字列モードでない場合
			else
				if l:inputstr[l:idx] =~ "{" || l:inputstr[l:idx] =~ "(" || l:inputstr[l:idx] =~ "[" || l:inputstr[l:idx] =~ "<"
					let l:tablevel = l:tablevel + 1
					let l:outputstr = l:outputstr . l:inputstr[l:idx] . "\n" . repeat("\t", l:tablevel)
					let l:skipchar = 1
				elseif l:inputstr[l:idx] =~ "," || l:inputstr[l:idx] =~ ";"
					let l:outputstr = l:outputstr . l:inputstr[l:idx] . "\n" . repeat("\t", l:tablevel)
					let l:skipchar = 1
				elseif l:inputstr[l:idx] =~ "}" || l:inputstr[l:idx] =~ ")" || l:inputstr[l:idx] =~ "]" || l:inputstr[l:idx] =~ ">"
					let l:tablevel = l:tablevel - 1
					let l:outputstr = l:outputstr . "\n" . repeat("\t", l:tablevel) . l:inputstr[l:idx]
					let l:skipchar = 1
				elseif l:inputstr[l:idx] =~ "\""
					let l:isStrMode = 1
					let l:outputstr = l:outputstr . l:inputstr[l:idx]
				else
					let l:outputstr = l:outputstr . l:inputstr[l:idx]
				endif
			endif
			
			let l:idx = l:idx + 1
			"空白文字以外の文字まで進める
			if l:skipchar == 1
				while l:inputstr[l:idx] =~ " "
					let l:idx = l:idx + 1
				endwhile
				let l:skipchar = 0
			endif
		endwhile
		"貼り付け
		let @* = l:outputstr
		silent normal p
	endfunction
" }}}

" ==============================================================================
" 端末の Vim でも Alt キーを使う
" [参照] https://thinca.hatenablog.com/entry/20101215/1292340358
" ==============================================================================
" {{{
let s:alt_enable = 0 " 本設定はLinux環境にてVim起動時にカーソル位置の文字が"4"に変えられる問題を引き起こす。そのため、無効化する。
if s:alt_enable
	if has('unix') && !has('gui_running')
		" Use meta keys in console.
		function! s:use_meta_keys()
			for i in map(
			\   range(char2nr('a'), char2nr('z'))
			\ + range(char2nr('A'), char2nr('Z'))
			\ + range(char2nr('0'), char2nr('9'))
			\ , 'nr2char(v:val)')
				" <ESC>O do not map because used by arrow keys.
				if i != 'O'
					execute 'nmap <ESC>' . i '<M-' . i . '>'
				endif
			endfor
		endfunction
		
		call s:use_meta_keys()
		map <NUL> <C-Space>
		map! <NUL> <C-Space>
	endif
endif
" }}}

" ==============================================================================
" ターミナル起動
" ==============================================================================
" {{{
	command! -nargs=? Tc terminal ++curwin
	command! -nargs=? Tv vert terminal
	command! -nargs=? Th bo terminal
" }}}

" ==============================================================================
" クリップボードからの貼り付け時、自動インデント無効
" [参考] https://ttssh2.osdn.jp/manual/4/ja/usage/tips/vim.html
" ==============================================================================
" {{{
if has("patch-8.0.0238")
	" Bracketed Paste Mode対応バージョン(8.0.0238以降)では、特に設定しない
	" 場合はTERMがxtermの時のみBracketed Paste Modeが使われる。
	" tmux利用時はTERMがscreenなので、Bracketed Paste Modeを利用するには
	" 以下の設定が必要となる。
	if &term =~ ".*screen.*" || &term =~ ".*tmux.*"
		let &t_BE = "\e[?2004h"
		let &t_BD = "\e[?2004l"
		exec "set t_PS=\e[200~"
		exec "set t_PE=\e[201~"
	endif
else
	" 8.0.0210 ～ 8.0.0237 ではVim本体でのBracketed Paste Mode対応の挙動が
	" 望ましくない(自動インデントが無効にならない)ので、Vim本体側での対応を
	" 無効にする。
	if has("patch-8.0.0210")
		set t_BE=
	endif
	" Vim本体がBracketed Paste Modeに対応していない時の為の設定。
	if &term =~ ".*xterm.*" || &term =~ ".*screen.*" || &term =~ ".*tmux.*"
		let &t_ti .= "\e[?2004h"
		let &t_te .= "\e[?2004l"
		function XTermPasteBegin(ret)
			set pastetoggle=<Esc>[201~
			set paste
			return a:ret
		endfunction
		noremap <special> <expr> <Esc>[200~ XTermPasteBegin("0i")
		inoremap <special> <expr> <Esc>[200~ XTermPasteBegin("")
		vnoremap <special> <expr> <Esc>[200~ XTermPasteBegin("c")
		cnoremap <special> <Esc>[200~ <nop>
		cnoremap <special> <Esc>[201~ <nop>
	endif
endif
" }}}

" ==============================================================================
" 誤って作られた"]"ファイル自動削除
" （:w時にenterと間違えて]を押すことがある）
" ==============================================================================
" {{{
	augroup setAutoDelete
		autocmd!
		autocmd BufWritePost ] call DeleteWronglyMadeBrackets()
	augroup END
	function! DeleteWronglyMadeBrackets()
		call delete("]")
		echom "']' file is deleted."
	endfunction
" }}}

" ==============================================================================
" 現在ファイルを外部エディタで開く
" ==============================================================================
" {{{
	function! OpenCurFileWithExternalEditer( sEditerPath )
		if has('win32')
			let sFilePath = fnameescape( GetCurFilePath() )
		"	echo sFilePath
			execute "!start " . a:sEditerPath . " " . l:sFilePath
		else
			echo "[error] OpenCurFileWithExternalEditer() can only be executed on windows."
		endif
	endfunction
" }}}

" ==============================================================================
" Tmuxからコピーした標準出力について、左右ペインの標準出力を抽出する
" ==============================================================================
" {{{
	command! -range Tml call ExtractPromptTmux(1) " extract left
	command! -range Tmr call ExtractPromptTmux(2) " extract right
	function! ExtractPromptTmux( iExtractPain ) " iExtractPain: 1(left) or 2(right)
		"選択文字列取得
		silent normal gvd
		let l:sInputStr = @*
		
		"文字列整形
		let l:sOutputStr = ""
		let l:asLines = split( l:sInputStr, "\n" )
		let l:sPattern = ''
		if a:iExtractPain == 1 " Extract left
			let l:sPattern = '\v\\u2502.*'
		elseif a:iExtractPain == 2 " Extract right
			let l:sPattern = '\v^.*\\u2502'
		else " other
			let l:sPattern = ''
			echom "[error] ExtractPromptTmux() iExtractPain must be 1(left) or 2(right)."
		endif
		for l:sLine in l:asLines
			let l:replacedstr = substitute( l:sLine, l:sPattern, "", "g" )
			if l:sOutputStr == ""
				let l:sOutputStr = l:replacedstr
			else
				let l:sOutputStr = l:sOutputStr . "\n" . l:replacedstr
			endif
		endfor
		let l:sOutputStr = l:sOutputStr . "\n"
		
		"貼り付け
		let @* = l:sOutputStr
		exec "normal k"
		silent normal p
	endfunction
" }}}

" ==============================================================================
" 引用番号を更新する
" ==============================================================================
" {{{
	command! Sqn call SetQuoteNo()
	function! SetQuoteNo()
		let l:sInFilePath = GetCurFilePath()
		if !filereadable(l:sInFilePath)
			echoerr "[error] File does not exist : " . l:sInFilePath
			return 0
		endif
		if &modified
			echoerr "[error] Save before execution."
			return 0
		endif
		
		let l:sTmpFilePath = l:sInFilePath . ".tmp"
		let l:sBakFilePath = l:sInFilePath . ".setquote_bak"
		while filereadable(l:sBakFilePath)
			let l:sBakFilePath = l:sBakFilePath . "_"
		endwhile
		
		call delete(l:sTmpFilePath)
		
		let l:sPattern = '\v\[\[\d+\]\]'
		let l:iQuoteIdx = 1
		try
			let l:asLines = readfile(l:sInFilePath)
			for l:sLine in l:asLines
				let l:sLineRep1 = substitute(l:sLine, l:sPattern, "", "g")
				let l:sLineRep2 = substitute(l:sLine, l:sPattern, "1", "g")
				let l:iMatchNum = len(l:sLineRep2) - len(l:sLineRep1)
				let l:sReplaceKeyword = "!!!!!"
				let l:sReplacedLine = substitute(l:sLine, l:sPattern, l:sReplaceKeyword, "g")
				for l:iMatchIdx in range( 1, l:iMatchNum )
					let l:sReplacedLine = substitute(l:sReplacedLine, l:sReplaceKeyword, "[[" . l:iQuoteIdx . "]]" , "")
					let l:iQuoteIdx += 1
				endfor
				call writefile([l:sReplacedLine], l:sTmpFilePath, 'a')
			endfor
		catch
			echo v:exception
		finally
			let l:sCopyCmdName = ""
			if has('win32')
				let l:sCopyCmdName = "copy"
			else
				let l:sCopyCmdName = "\cp -f"
			endif
			call system(l:sCopyCmdName . ' "' . l:sInFilePath . '" "' . l:sBakFilePath . '"')
			call system(l:sCopyCmdName . ' "' . l:sTmpFilePath . '" "' . l:sInFilePath . '"')
			call delete(l:sTmpFilePath)
			execute 'edit! ' . l:sInFilePath
		endtry
	endfunction
" }}}
" 
" ==============================================================================
" 
" ==============================================================================
	function! SearchStart()
		let l:sCurWord = expand("<cword>")
	"	exec "normal /"
		let l:sLineNoOld = line(".")
		exec "normal *"
	"	let l:sLineNoNew = line(".")
	"	"echom l:sLineNoOld . " " . l:sLineNoNew
	"	if l:sLineNoOld != l:sLineNoNew
	"		normal N
	"	endif
	endfunction

" ==============================================================================
" 和訳のために英文を整形する
" ==============================================================================
" {{{
	command! -range Fet call FormatEnglishText()
	function! FormatEnglishText()
		"s/\v\. +([A-Z0-9])/\.\r\1/g
		"let l:sReplacedLine = substitute(l:sReplacedLine, l:sReplaceKeyword, "[[" . l:iQuoteIdx . "]]" , "")
		"TODO: https://nanasi.jp/articles/code/screen/visual.html
	endfunction
" }}}

" ==============================================================================
" diffopt設定トグル
" ==============================================================================
" {{{
	command! Tiw call ToggleDiffopt('iwhite')
	command! Tic call ToggleDiffopt('icase')
	function! ToggleDiffopt(opt)
		if &diffopt =~# '\<' . a:opt . '\>'
			execute 'set diffopt-=' . a:opt
			echo 'diffopt: removed ' . a:opt
		else
			execute 'set diffopt+=' . a:opt
			echo 'diffopt: added ' . a:opt
		endif
		execute 'set diffopt?'
	endfunction
" }}}

" **************************************************************************************************
" *****										プラグイン設定									   *****
" **************************************************************************************************
" ==============================================================================
" Taglist 設定
" ==============================================================================
" {{{
"	let Tlist_Show_One_File = 1		" アクティブバッファのみタグ表示
"	let Tlist_Use_Right_Window = 1	" 右ウィンドウ表示
"	let Tlist_Exit_OnlyWindow = 1	" taglistのウインドウだけならVimを閉じる
"	let Tlist_Display_Prototype = 0 " プロトタイプを非表示
"	let Tlist_Display_Tag_Scope = 0 " タグスコープを非表示
"	let Tlist_Auto_Open = 0			" 自動起動無効
" }}}

" ==============================================================================
" Tagbar 設定
" ==============================================================================
" {{{
"	let g:tagbar_ctags_bin = 
	let g:tagbar_type_vb = {
		\ 'ctagstype' : 'vb',
		\ 'kinds'     : [
			\ 'd:macros:1:0',
			\ 'p:prototypes:1:0',
			\ 'g:enums',
			\ 'e:enumerators:0:0',
			\ 't:typedefs:0:0',
			\ 'n:namespaces',
			\ 'c:classes',
			\ 's:structs',
			\ 'u:unions',
			\ 'f:functions',
			\ 'm:members:0:0',
			\ 'v:variables:0:0'
		\ ]
	\ }
	let g:tagbar_sort = 0	" ソートしない
	let g:tagbar_show_linenumbers = 1
	let g:tagbar_autopreview = 0
	let g:tagbar_autofocus = 0
	let g:tagbar_width = 60
	let g:tagbar_autoclose = 0
	let g:tagbar_case_insensitive = 0
	let g:tagbar_compact = 0
	let g:tagbar_indent = 2
	let g:tagbar_no_status_line = 1
" }}}

" ==============================================================================
" bufferlist 設定
" ==============================================================================
" {{{
	let g:BufferListWidth = 30
	let g:BufferListHideBufferList = 0
	let g:BufferListExpandBufName = 0
	let g:BufferListPreview = 0
	let g:BufferListTailWidth = 9
	let g:BufferListShortenChar = "..."
	hi BufferSelected guifg=black guibg=#9ad000 gui=bold
"	hi BufferNormal guifg=white
" }}}

" ==============================================================================
" align.vim 設定
" [参照] http://vim-users.jp/2009/09/hack77/
" ==============================================================================
" {{{
	let g:align_xstrlen = 3 " 日本語用
" }}}

" ==============================================================================
" code_overview 設定
" [参照] http://vim-users.jp/2009/09/hack77/
" ==============================================================================
" {{{
"	let g:code_overview_autostart = 1
"	let g:code_overview_use_colorscheme = 1
"	let g:codeoverview_autoupdate = 1
" }}}

" =======================================
" open-browser の設定
" =======================================
" {{{
	let g:netrw_nogx = 1 " disable netrw's gx mapping.
" }}}

" ==============================================================================
" mark.vim 設定
" ★カラースキーマ設定の後に記述すること！★
" [参照] http://nanasi.jp/articles/vim/mark_vim.html
" ==============================================================================
" {{{
	execute 'source ' . $MARKVIM
	command! -nargs=? M execute 'source ' . $MARKVIM
" }}}

" ==============================================================================
" winresizer.vim 設定
" [参照] https://github.com/simeji/winresizer
" ==============================================================================
" {{{
"	let g:winresizer_enable = 1
"	let g:winresizer_start_key = '<C-S-T>'
" }}}

" ==============================================================================
" qfixgrep 設定
" [参照] http://vim-users.jp/2009/09/hack77/
" ==============================================================================
" {{{
if has('unix')
	let QFix_PreviewHeight = 15
	let QFix_Height = 15
else
	let QFix_PreviewHeight = 12
endif
	let MyGrep_Commands = 1
	let mygrepprg = 'grep'
	let QFixWin_EnableMode = 1		" QuickFixウィンドウでもプレビューや絞り込みを有効化
	let QFix_UseLocationList = 0	" QFixHowm/QFixGrepの結果表示にロケーションリストを使用する/しない
"	set shellslash					" Windowsの場合は、shellslash を設定してやると、パスの表記が簡単になります。ex) let MyGrep_ExcludeReg = '/CVS/'
	let MyGrep_ExcludeReg = '[~#]$\|\.bak$\|\.o$\|\.obj$\|\.exe$\|[/\\]tags$\|^tags$'
"	let MyGrep_Encoding = 'cp932'	"cp932を扱えないGNU grepの場合
									"Windowsでcp932を扱えないGNU grepを使用する場合、以下の様に設定します。
	let QFix_CopenCmd = 'botright'	" Quickfixウィンドウを最も下側に表示
	
" More scrollbar-ish behavior
	let g:nanomap_auto_realign = 1
	let g:nanomap_auto_open_close = 1
	let g:nanomap_highlight_delay = 100
" }}}

" ==============================================================================
" vaffle 設定
" ==============================================================================
" {{{
	let g:vaffle_show_hidden_files = 1	" 隠しファイルを表示
" }}}

" ==============================================================================
" neosnippet設定
" ==============================================================================
" {{{
if has('unix')
	let g:neosnippet#snippets_directory = $HOME . '/.vim/_snipets'
else
	let g:neosnippet#snippets_directory = $VIM . '/_snipets'
endif
if has('conceal')
	set conceallevel=2 concealcursor=i
endif
" }}}

" ==============================================================================
" surround 設定
" ==============================================================================
" {{{
	let g:surround_{char2nr("「")} = "「\r」"
	let g:surround_{char2nr("」")} = "「\r」"
	let g:surround_{char2nr("【")} = "【\r】"
	let g:surround_{char2nr("】")} = "【\r】"
	let g:surround_{char2nr("（")} = "（\r）"
	let g:surround_{char2nr("）")} = "（\r）"
	let g:surround_{char2nr("＜")} = "＜\r＞"
	let g:surround_{char2nr("＞")} = "＜\r＞"
	let g:surround_{char2nr("｛")} = "｛\r｝"
	let g:surround_{char2nr("｝")} = "｛\r｝"
" }}}

" ==============================================================================
" linediff 設定
" ==============================================================================
" {{{
	let g:linediff_first_buffer_command  = 'leftabove new'
	let g:linediff_second_buffer_command = 'rightbelow vertical new'
" }}}

" ==============================================================================
" showmarks 設定
" ==============================================================================
" {{{
	autocmd VimEnter * DoShowMarks!
" }}}

" " TODO:
" " クイックフィックスウィンドウを開かないようにする
" let g:lsp_diagnostics_enabled = 1
" let g:lsp_signs_enabled = 1
" let g:lsp_diagnostics_echo_cursor = 1
" 
" " 自動補完のオプション
" set completeopt=menuone,noinsert,noselect

