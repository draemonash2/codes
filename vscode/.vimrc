"=== モード共通 ===
noremap							x			"_x
"nnoremap						dd			"_dd
"nnoremap						cc			dd
noremap							H			0
noremap							L			$
"noremap							/			/\v
noremap							/			/\V
"noremap					<expr>	gp			'`[' . strpart(getregtype(), 0, 1) . '`]'
"map								go			<Plug>(openbrowser-smart-search)
noremap							<space>		<nop>
noremap							<c-j>		10jzz
noremap							<c-k>		10kzz
noremap							<c-h>		10zh10h
noremap							<bs>		10zh10h
noremap							<c-l>		10zl10l

"=== ノーマルモード ===
"nnoremap						<cr>		<nop>
nnoremap				<expr>	gc			":Gc " . expand('<cword>')
nnoremap				<expr>	gp			":Gp " . expand('<cword>')
nnoremap				<expr>	gt			":Gt"
nnoremap						ciy			ciw<c-r>0<esc>b
nnoremap						<esc><esc>	:nohlsearch<cr>
nnoremap						<tab>		a<tab><esc>
nnoremap						<c-d>		:Gtags <C-r><C-w><CR>
nnoremap						<c-e>		:Gtags -r <C-r><C-w><CR>
"nnoremap						gs			:Gtags -s <C-r><C-w><CR>
"nnoremap						gG			:Gtags -g <C-r><C-w><CR>
"nnoremap						,d			:<C-u>Gtags -f %<CR>
nnoremap						<c-t>		<C-o>
"nnoremap						<c-g>		:call ExecuteGuiGrep()<cr>
nnoremap						<c-m>		oa<esc>"_x
nnoremap						<c-tab>		<c-w>w
nnoremap						<c-s-tab>	<c-w>W
"nnoremap						<c-tab>		gt
"nnoremap						<c-s-tab>	gT
nnoremap						<c-w><c-k>	5<c-w>+
nnoremap						<c-w><c-j>	5<c-w>-
nnoremap						<c-w><c-h>	5<c-w>>
nnoremap						<c-w><c-l>	5<c-w><
nnoremap						<space>h	<c-w>h
nnoremap						<space>j	<c-w>j
nnoremap						<space>k	<c-w>k
nnoremap						<space>l	<c-w>l
nnoremap						<c-n><c-h>	<c-w>h
nnoremap						<c-n><c-j>	<c-w>j
nnoremap						<c-n><c-k>	<c-w>k
nnoremap						<c-n><c-l>	<c-w>l
nnoremap						<M-h>		<c-w>h
nnoremap						<M-j>		<c-w>j
nnoremap						<M-k>		<c-w>k
nnoremap						<M-l>		<c-w>l
nnoremap						<c-]>		g<c-]>
"nnoremap						<c-w>		:tabclose<cr>
"nnoremap						<c-n>		:tabnew<cr>
"nmap							<c-o>		:GtagsCursor<CR>
"nnoremap						<F1>		:call BufferList()<cr>
"nnoremap						<F2>		:TagbarToggle<cr>
"nnoremap						<F3>														" <F3>はペーストモード切替に割り当て
"nnoremap									a<C-R>=strftime("%Y/%m/%d (%a)")<cr><esc>
"nnoremap									a<C-R>=strftime("%H:%M:%S")<cr><esc>
"nnoremap						<F4>														" <F4>はCtrlP起動に割り当て
"nnoremap						<F5>		:execute ExecCurrentScript()<cr>
"nnoremap						<F6>		:vs<cr><c-w>wggVGy:q<cr><c-w>W
"nnoremap						<F7>		:Vexplore<cr>
"nnoremap						<F7>		:NERDTreeToggle<CR>
"nnoremap						<F8>		:call SwitchFontSize()<cr>
nmap							<F9>		kyiwjciw<c-r>0<esc>b<c-a>j
nnoremap						<F10>		:call ToggleWindowSize()<cr>
nnoremap						<F11>		:set expandtab<cr>:retab!<cr>
nnoremap						<F12>		:set noexpandtab<cr>:retab!<cr>
"nnoremap						<s-F1>		:call CopyCurPrjRltvFilePath()<cr>
"nnoremap						<c-F1>		:call CopyCurFilePath()<cr>
"nnoremap						<c-F2>		:call CopyCurFileName()<cr>
"nnoremap						<c-F3>		:call CopyCurDirPath()<cr>
"nnoremap						<c-F4>		:call CopyCurFileExt()<cr>
""nnoremap						<c-F5>		:call CopyFileLineNo()<cr>
""nnoremap						<c-F5>		:call CopyFileExtAndLineNo()<cr>
"nnoremap						<c-F5>		:call CopyRltvFilePathAndLineNo()<cr>
"nnoremap						<c-F8>		:call SwitchBirdEyesViewMode()<cr>
"nnoremap						<c-s-F5>	:call UpdateTagFile()<cr>
"nmap							n			<Plug>(anzu-n)zz
"nmap							N			<Plug>(anzu-N)zz
"nmap							*			<Plug>(anzu-star)zz
"nmap							#			<Plug>(anzu-sharp)zz
nmap							K			g*zz
nmap							g#			g#zz
"nmap							<Esc><Esc>	<Plug>(anzu-clear-search-status):noh<cr>

"=== 挿入モード ===
inoremap						<c-j>		<esc>
"imap							<c-k>		<Plug>(neosnippet_expand_or_jump)
"imap					<expr>	<TAB>		neosnippet#expandable_or_jumpable() ? "\<Plug>(neosnippet_expand_or_jump)" : pumvisible() ? "\<C-n>" : "\<TAB>"
"if has('unix')
"imap							<c-v>		<c-r>0
"else
imap							<c-v>		<S-Insert>
"endif
"inoremap						{			{}<LEFT>
"inoremap						[			[]<LEFT>
"inoremap						(			()<LEFT>
"inoremap						"			""<LEFT>
"inoremap						'			''<LEFT>
"inoremap						【			【】<LEFT>
"inoremap						「			「」<LEFT>
"inoremap						『			『』<LEFT>
"inoremap						“			“”<LEFT>
"inoremap						‘			‘’<LEFT>
"inoremap						（			（）<LEFT>
"inoremap						｛			｛｝<LEFT>

"=== ヴィジュアルモード ===
vmap							Y			<esc>:set expandtab<cr>gv:retab!<cr>gvyu
"vnoremap						T			<esc>:set expandtab<cr>gv:retab!<cr>gv:call CopyLineWithLineNo()<cr>u
"vnoremap						p			"0p<cr>											" ヤンクせずに貼り付け
vnoremap						s			c
vnoremap						<			<gv
vnoremap						>			>gv
vnoremap						<F3>		:s/\s*$//<cr>:nohlsearch<cr>
vnoremap						<F4>		:s/\\/\//g<cr>
vnoremap						<F5>		:s/\v_(.)/\u\1/g<cr>
vnoremap						<F6>		:s/\v([A-Z])/_\L\1/g<cr>
vnoremap						<F7>		:s/\v^(\w+).*/\1/g<cr>
vnoremap						<F8>		:s/\v^\!.*\!\n//g<cr>
"vnoremap						<F9>		:call ReplaceRelativePathFromCurrent()<cr>
vnoremap						<F11>		<esc>:set expandtab<cr>gv:retab!<cr>
vnoremap						<F12>		<esc>:set noexpandtab<cr>gv:retab!<cr>
"vnoremap						<c-F5>		:<bs><bs><bs><bs><bs>call CopyFileLineNo()<cr>
"vnoremap						d			"_d
"vnoremap						c			d
"vnoremap						p			"0p
"vnoremap				<expr>	gc			":Gc " . expand('<cword>')
"vnoremap				<expr>	gp			":Gp " . expand('<cword>')
"vnoremap				<expr>	gt			":Gt " . expand('<cword>')

"=== コマンドラインモード ===
cnoremap						<c-n>		<Down>
cnoremap						<c-p>		<Up>
cnoremap						<c-b>		<Left>
cnoremap						<c-f>		<Right>
cnoremap						<c-a>		<Home>
cnoremap						<c-e>		<End>
cnoremap						<c-d>		<Del>
"if has('unix')
"cmap							<c-v>		<c-r>0
"else
cmap							<c-v>		<S-Insert>
"endif

"=== その他モード ===
"smap							<C-k>		<Plug>(neosnippet_expand_or_jump)
"xmap							<C-k>		<Plug>(neosnippet_expand_target)
"smap					<expr>	<TAB>		neosnippet#expandable_or_jumpable() ?  "\<Plug>(neosnippet_expand_or_jump)" : "\<TAB>"
