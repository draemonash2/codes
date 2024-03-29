" Script Name: mark.vim
" Version:     1.1.10 (global version)
" Last Change: January 16, 2015
" Author:      Yuheng Xie <thinelephant@gmail.com>
" Contributor: Luc Hermitte
"
" Description: a little script to highlight several words in different colors
"              simultaneously
"
" Usage:       :Mark regexp   to mark a regular expression
"              :Mark regexp   with exactly the same regexp to unmark it
"              :Mark          to clear all marks
"
"              You may map keys for the call in your vimrc file for
"              convenience. The default keys is:
"              Highlighting:
"                Normal \m  mark or unmark the word under or before the cursor
"                       \r  manually input a regular expression
"                       \n  clear current mark (i.e. the mark under the cursor),
"                           or clear all marks
"                Visual \m  mark or unmark a visual selection
"                       \r  manually input a regular expression
"              Searching:
"                Normal \*  jump to the next occurrence of current mark
"                       \#  jump to the previous occurrence of current mark
"                       \/  jump to the next occurrence of ANY mark
"                       \?  jump to the previous occurrence of ANY mark
"                        *  behaviors vary, please refer to the table on
"                        #  line 123
"                combined with VIM's / and ? etc.
"
"              The default colors/groups setting is for marking six
"              different words in different colors. You may define your own
"              colors in your vimrc file. That is to define highlight group
"              names as "MarkWordN", where N is a number. An example could be
"              found below.
"
" Bugs:        some colored words could not be highlighted
"
" Changes:
" 16th Jan 2015, Yuheng Xie: add auto event WinEnter
" (*) added auto event WinEnter for reloading highlights after :split, etc.
"
" 29th Jul 2014, Yuheng Xie: call matchadd()
" (*) added call to VIM 7.1 matchadd(), make highlighting keywords possible
"
" 10th Mar 2006, Yuheng Xie: jump to ANY mark
" (*) added \* \# \/ \? for the ability of jumping to ANY mark, even when the
"     cursor is not currently over any mark
"
" 20th Sep 2005, Yuheng Xie: minor modifications
" (*) merged MarkRegexVisual into MarkRegex
" (*) added GetVisualSelectionEscaped for multi-lines visual selection and
"     visual selection contains ^, $, etc.
" (*) changed the name ThisMark to CurrentMark
" (*) added SearchCurrentMark and re-used raw map (instead of VIM function) to
"     implement * and #
"
" 14th Sep 2005, Luc Hermitte: modifications done on v1.1.4
" (*) anti-reinclusion guards. They do not guard colors definitions in case
"     this script must be reloaded after .gvimrc
" (*) Protection against disabled |line-continuation|s.
" (*) Script-local functions
" (*) Default keybindings
" (*) \r for visual mode
" (*) uses <leader> instead of "\"
" (*) do not mess with global variable g:w
" (*) regex simplified -> double quotes changed into simple quotes.
" (*) strpart(str, idx, 1) -> str[idx]
" (*) command :Mark
"     -> e.g. :Mark Mark.\{-}\ze(

" default colors/groups
" you may define your own colors in you vimrc file, in the form as below:
"  custom mod <TOP>
let g:sColorPattern = 1
if g:sColorPattern == 0 "default
    hi MarkWord1  ctermbg=Cyan     ctermfg=Black  guibg=#8CCBEA    guifg=Black      gui=none
    hi MarkWord2  ctermbg=Green    ctermfg=Black  guibg=#A4E57E    guifg=Black      gui=none
    hi MarkWord3  ctermbg=Yellow   ctermfg=Black  guibg=#FFDB72    guifg=Black      gui=none
    hi MarkWord4  ctermbg=Red      ctermfg=Black  guibg=#FF7272    guifg=Black      gui=none
    hi MarkWord5  ctermbg=Magenta  ctermfg=Black  guibg=#FFB3FF    guifg=Black      gui=none
    hi MarkWord6  ctermbg=Blue     ctermfg=Black  guibg=#9999FF    guifg=Black      gui=none
    hi MarkWord7  ctermbg=White    ctermfg=Black  guibg=White      guifg=Black      gui=bold
elseif g:sColorPattern == 1
    hi MarkWord1  ctermbg=Cyan     ctermfg=Black  guibg=#8CCBEA    guifg=Black      gui=bold
    hi MarkWord2  ctermbg=Green    ctermfg=Black  guibg=#A4E57E    guifg=Black      gui=bold
    hi MarkWord3  ctermbg=Yellow   ctermfg=Black  guibg=#FFDB72    guifg=Black      gui=bold
    hi MarkWord4  ctermbg=Red      ctermfg=Black  guibg=#FF7272    guifg=Black      gui=bold
    hi MarkWord5  ctermbg=Magenta  ctermfg=Black  guibg=#FFB3FF    guifg=Black      gui=bold
    hi MarkWord6  ctermbg=Blue     ctermfg=Black  guibg=#9999FF    guifg=Black      gui=bold
    hi MarkWord7  ctermbg=White    ctermfg=Black  guibg=White      guifg=Black      gui=bold
elseif g:sColorPattern == 2
    hi MarkWord1  ctermbg=Cyan     ctermfg=Black  guibg=#66FFFF    guifg=#007D7A    gui=bold
    hi MarkWord2  ctermbg=Green    ctermfg=Black  guibg=#66FF33    guifg=#208600    gui=bold
    hi MarkWord3  ctermbg=Yellow   ctermfg=Black  guibg=#FFFF00    guifg=#7D7A00    gui=bold
    hi MarkWord4  ctermbg=Red      ctermfg=Black  guibg=#FFC000    guifg=#765A00    gui=bold
    hi MarkWord5  ctermbg=Magenta  ctermfg=Black  guibg=#FF99CC    guifg=#A20051    gui=bold
    hi MarkWord6  ctermbg=Blue     ctermfg=Black  guibg=#CC99FF    guifg=#4B0096    gui=bold
    hi MarkWord7  ctermbg=White    ctermfg=Black  guibg=White      guifg=Black      gui=bold
elseif g:sColorPattern == 3
    hi MarkWord1  ctermbg=Cyan     ctermfg=Black  guibg=#CCECFF    guifg=#366092    gui=bold
    hi MarkWord2  ctermbg=Magenta  ctermfg=Black  guibg=#FFCCFF    guifg=#BC005E    gui=bold
    hi MarkWord3  ctermbg=Yellow   ctermfg=Black  guibg=#FFFF99    guifg=#605E00    gui=bold
    hi MarkWord4  ctermbg=Blue     ctermfg=Black  guibg=#CC99FF    guifg=#6B00D6    gui=bold
    hi MarkWord5  ctermbg=Green    ctermfg=Black  guibg=#CCFF99    guifg=#478E00    gui=bold
    hi MarkWord6  ctermbg=Red      ctermfg=Black  guibg=#FFC000    guifg=#8E6C00    gui=bold
    hi MarkWord7  ctermbg=White    ctermfg=Black  guibg=White      guifg=Black      gui=bold
else
endif
"  custom mod <END>

" Anti reinclusion guards
if exists('g:loaded_mark') && !exists('g:force_reload_mark')
  finish
endif

" Support for |line-continuation|
let s:save_cpo = &cpo
set cpo&vim

" Default bindings

" custom mod <TOP>
if !hasmapto('<Plug>MarkSet', 'n')
  nmap <unique> <silent> <space><space> <Plug>MarkSet
endif
if !hasmapto('<Plug>MarkSet', 'v')
  vmap <unique> <silent> <space><space> <Plug>MarkSet
endif
if !hasmapto('<Plug>MarkRegex', 'n')
  nmap <unique> <silent> <space>r <Plug>MarkRegex
endif
if !hasmapto('<Plug>MarkRegex', 'v')
  vmap <unique> <silent> <space>r <Plug>MarkRegex
endif
if !hasmapto('<Plug>MarkClear', 'n')
  nmap <unique> <silent> <space>n <Plug>MarkClear
endif
" custom mod <END>
" custom add <TOP>
if !hasmapto('<Plug>MarkAllClear', 'n')
  nmap <unique> <silent> <space>c <Plug>MarkAllClear
endif
" custom add <END>

nnoremap <silent> <Plug>MarkSet   :call
\ <sid>MarkCurrentWord()<cr>
vnoremap <silent> <Plug>MarkSet   <c-\><c-n>:call
\ <sid>DoMark(<sid>GetVisualSelectionEscaped("enV"))<cr>
nnoremap <silent> <Plug>MarkRegex :call
\ <sid>MarkRegex()<cr>
vnoremap <silent> <Plug>MarkRegex <c-\><c-n>:call
\ <sid>MarkRegex(<sid>GetVisualSelectionEscaped("N"))<cr>
nnoremap <silent> <Plug>MarkClear :call
\ <sid>DoMark(<sid>CurrentMark())<cr>
" custom add <TOP>
nnoremap <silent> <Plug>MarkAllClear :call
\ <sid>ClearAllMark()<cr>
" custom add <END>

" Here is a sumerization of the following keys' behaviors:
" 
" First of all, \#, \? and # behave just like \*, \/ and *, respectively,
" except that \#, \? and # search backward.
"
" \*, \/ and *'s behaviors differ base on whether the cursor is currently
" placed over an active mark:
"
"       Cursor over mark                  Cursor not over mark
" ---------------------------------------------------------------------------
"  \*   jump to the next occurrence of    jump to the next occurrence of
"       current mark, and remember it     "last mark".
"       as "last mark".
"
"  \/   jump to the next occurrence of    same as left
"       ANY mark.
"
"   *   if \* is the most recently used,  do VIM's original *
"       do a \*; otherwise (\/ is the
"       most recently used), do a \/.

"nnoremap <silent> <leader>* :call <sid>SearchCurrentMark()<cr>                    " custom del
"nnoremap <silent> <leader># :call <sid>SearchCurrentMark("b")<cr>                 " custom del
"nnoremap <silent> <leader>/ :call <sid>SearchAnyMark()<cr>                        " custom del
"nnoremap <silent> <leader>? :call <sid>SearchAnyMark("b")<cr>                     " custom del
"nnoremap <silent> * :if !<sid>SearchNext()<bar>execute "norm! *"<bar>endif<cr>    " custom del
"nnoremap <silent> # :if !<sid>SearchNext("b")<bar>execute "norm! #"<bar>endif<cr> " custom del

command! -nargs=? Mark call s:DoMark(<f-args>)

autocmd! BufWinEnter,WinEnter * call s:UpdateMark()

" Functions

function! s:MarkCurrentWord()
let w = s:PrevWord()
if w != ""
  call s:DoMark('\<' . w . '\>')
endif
endfunction

function! s:GetVisualSelection()
let save_a = @a
silent normal! gv"ay
let res = @a
let @a = save_a
return res
endfunction

function! s:GetVisualSelectionEscaped(flags)
" flags:
"  "e" \  -> \\  
"  "n" \n -> \\n  for multi-lines visual selection
"  "N" \n removed
"  "V" \V added   for marking plain ^, $, etc.
let result = s:GetVisualSelection()
let i = 0
while i < strlen(a:flags)
  if a:flags[i] ==# "e"
    let result = escape(result, '\')
  elseif a:flags[i] ==# "n"
    let result = substitute(result, '\n', '\\n', 'g')
  elseif a:flags[i] ==# "N"
    let result = substitute(result, '\n', '', 'g')
  elseif a:flags[i] ==# "V"
    let result = '\V' . result
  endif
  let i = i + 1
endwhile
return result
endfunction

" manually input a regular expression
function! s:MarkRegex(...) " MarkRegex(regexp)
let regexp = ""
if a:0 > 0
  let regexp = a:1
endif
call inputsave()
let r = input("@", regexp)
call inputrestore()
if r != ""
  call s:DoMark(r)
endif
endfunction

" define variables if they don't exist
function! s:InitMarkVariables()
if !exists("g:mwHistAdd")
  let g:mwHistAdd = "/@"
endif
if !exists("g:mwCycleMax")
  let i = 1
  while hlexists("MarkWord" . i)
    let i = i + 1
  endwhile
  let g:mwCycleMax = i - 1
endif
if !exists("g:mwCycle")
  let g:mwCycle = 1
endif
let i = 1
while i <= g:mwCycleMax
  if !exists("g:mwWord" . i)
    let g:mwWord{i} = ""
  endif
  let i = i + 1
endwhile
if !exists("g:mwLastSearched")
  let g:mwLastSearched = ""
endif
endfunction

" return the word under or before the cursor
function! s:PrevWord()
let line = getline(".")
if line[col(".") - 1] =~ '\w'
  return expand("<cword>")
else
  return substitute(strpart(line, 0, col(".") - 1), '^.\{-}\(\w\+\)\W*$', '\1', '')
endif
endfunction

" mark or unmark a regular expression
function! s:DoMark(...) " DoMark(regexp)
" define variables if they don't exist
call s:InitMarkVariables()

" clear all marks if regexp is null
let regexp = ""
if a:0 > 0
  let regexp = a:1
endif
if regexp == ""
  call s:ClearAllMark() " custom add
  " custom del <TOP>
  "let i = 1
  "while i <= g:mwCycleMax
  "  if g:mwWord{i} != ""
  "    let g:mwWord{i} = ""
  "    let lastwinnr = winnr()
  "    let winview = winsaveview()
  "    if exists("*matchadd")
    "      windo silent! call matchdelete(3333 + i)
    "    else
    "      exe "windo syntax clear MarkWord" . i
    "    endif
    "    exe lastwinnr . "wincmd w"
    "    call winrestview(winview)
    "  endif
    "  let i = i + 1
    "endwhile
    "let g:mwLastSearched = ""
    " custom del <END>
    return 0
  endif

  " clear the mark if it has been marked
  let i = 1
  while i <= g:mwCycleMax
    if regexp == g:mwWord{i}
      if g:mwLastSearched == g:mwWord{i}
        let g:mwLastSearched = ""
      endif
      let g:mwWord{i} = ""
      let lastwinnr = winnr()
      let winview = winsaveview()
      if exists("*matchadd")
        windo silent! call matchdelete(3333 + i)
      else
        exe "windo syntax clear MarkWord" . i
      endif
      exe lastwinnr . "wincmd w"
      call winrestview(winview)
      return 0
    endif
    let i = i + 1
  endwhile

  " add to history
  if stridx(g:mwHistAdd, "/") >= 0
    call histadd("/", regexp)
  endif
  if stridx(g:mwHistAdd, "@") >= 0
    call histadd("@", regexp)
  endif

  " quote regexp with / etc. e.g. pattern => /pattern/
  let quote = "/?~!@#$%^&*+-=,.:"
  let i = 0
  while i < strlen(quote)
    if stridx(regexp, quote[i]) < 0
      let quoted_regexp = quote[i] . regexp . quote[i]
      break
    endif
    let i = i + 1
  endwhile
  if i >= strlen(quote)
    return -1
  endif

  " choose an unused mark group
  let i = 1
  while i <= g:mwCycleMax
    if g:mwWord{i} == ""
      let g:mwWord{i} = regexp
      if i < g:mwCycleMax
        let g:mwCycle = i + 1
      else
        let g:mwCycle = 1
      endif
      let lastwinnr = winnr()
      let winview = winsaveview()
      if exists("*matchadd")
        windo silent! call matchdelete(3333 + i)
        windo silent! call matchadd("MarkWord" . i, g:mwWord{i}, -10, 3333 + i)
      else
        exe "windo syntax clear MarkWord" . i
        " suggested by Marc Weber, use .* instead off ALL
        exe "windo syntax match MarkWord" . i . " " . quoted_regexp . " containedin=.*"
      endif
      exe lastwinnr . "wincmd w"
      call winrestview(winview)
      return i
    endif
    let i = i + 1
  endwhile

  " choose a mark group by cycle
  let i = 1
  while i <= g:mwCycleMax
    if g:mwCycle == i
      if g:mwLastSearched == g:mwWord{i}
        let g:mwLastSearched = ""
      endif
      let g:mwWord{i} = regexp
      if i < g:mwCycleMax
        let g:mwCycle = i + 1
      else
        let g:mwCycle = 1
      endif
      let lastwinnr = winnr()
      let winview = winsaveview()
      if exists("*matchadd")
        windo silent! call matchdelete(3333 + i)
        windo silent! call matchadd("MarkWord" . i, g:mwWord{i}, -10, 3333 + i)
      else
        exe "windo syntax clear MarkWord" . i
        " suggested by Marc Weber, use .* instead off ALL
        exe "windo syntax match MarkWord" . i . " " . quoted_regexp . " containedin=.*"
      endif
      exe lastwinnr . "wincmd w"
      call winrestview(winview)
      return i
    endif
    let i = i + 1
  endwhile
endfunction

" update mark colors
function! s:UpdateMark()
  " define variables if they don't exist
  call s:InitMarkVariables()

  let i = 1
  while i <= g:mwCycleMax
    exe "syntax clear MarkWord" . i
    if g:mwWord{i} != ""
      " quote regexp with / etc. e.g. pattern => /pattern/
      let quote = "/?~!@#$%^&*+-=,.:"
      let j = 0
      while j < strlen(quote)
        if stridx(g:mwWord{i}, quote[j]) < 0
          let quoted_regexp = quote[j] . g:mwWord{i} . quote[j]
          break
        endif
        let j = j + 1
      endwhile
      if j >= strlen(quote)
        continue
      endif

      if exists("*matchadd")
        silent! call matchadd("MarkWord" . i, g:mwWord{i}, -10, 3333 + i)
      else
        " suggested by Marc Weber, use .* instead off ALL
        exe "syntax match MarkWord" . i . " " . quoted_regexp . " containedin=.*"
      endif
    endif
    let i = i + 1
  endwhile
endfunction

" return the mark string under the cursor. multi-lines marks not supported
function! s:CurrentMark()
  " define variables if they don't exist
  call s:InitMarkVariables()

  let line = getline(".")
  let i = 1
  while i <= g:mwCycleMax
    if g:mwWord{i} != ""
      let start = 0
      while start >= 0 && start < strlen(line) && start < col(".")
        let b = match(line, g:mwWord{i}, start)
        let e = matchend(line, g:mwWord{i}, start)
        if b < col(".") && col(".") <= e
          let s:current_mark_position = line(".") . "_" . b
          return g:mwWord{i}
        endif
        let start = e
      endwhile
    endif
    let i = i + 1
  endwhile
  return ""
endfunction

" search current mark
function! s:SearchCurrentMark(...) " SearchCurrentMark(flags)
  let flags = ""
  if a:0 > 0
    let flags = a:1
  endif
  let w = s:CurrentMark()
  if w != ""
    let p = s:current_mark_position
    call search(w, flags)
    call s:CurrentMark()
    if p == s:current_mark_position
      call search(w, flags)
    endif
    let g:mwLastSearched = w
  else
    if g:mwLastSearched != ""
      call search(g:mwLastSearched, flags)
    else
      call s:SearchAnyMark(flags)
      let g:mwLastSearched = s:CurrentMark()
    endif
  endif
endfunction

" combine all marks into one regexp
function! s:AnyMark()
  " define variables if they don't exist
  call s:InitMarkVariables()

  let w = ""
  let i = 1
  while i <= g:mwCycleMax
    if g:mwWord{i} != ""
      if w != ""
        let w = w . '\|' . g:mwWord{i}
      else
        let w = g:mwWord{i}
      endif
    endif
    let i = i + 1
  endwhile
  return w
endfunction

" search any mark
function! s:SearchAnyMark(...) " SearchAnyMark(flags)
  let flags = ""
  if a:0 > 0
    let flags = a:1
  endif
  let w = s:CurrentMark()
  if w != ""
    let p = s:current_mark_position
  else
    let p = ""
  endif
  let w = s:AnyMark()
  call search(w, flags)
  call s:CurrentMark()
  if p == s:current_mark_position
    call search(w, flags)
  endif
  let g:mwLastSearched = ""
endfunction

" search last searched mark
function! s:SearchNext(...) " SearchNext(flags)
  let flags = ""
  if a:0 > 0
    let flags = a:1
  endif
  let w = s:CurrentMark()
  if w != ""
    if g:mwLastSearched != ""
      call s:SearchCurrentMark(flags)
    else
      call s:SearchAnyMark(flags)
    endif
    return 1
  else
    return 0
  endif
endfunction

" Restore previous 'cpo' value
let &cpo = s:save_cpo

" custom add <TOP>
function! s:ClearAllMark()
  let i = 1
  while i <= g:mwCycleMax
    if g:mwWord{i} != ""
      let g:mwWord{i} = ""
      let lastwinnr = winnr()
      let winview = winsaveview()
      if exists("*matchadd")
        windo silent! call matchdelete(3333 + i)
      else
        exe "windo syntax clear MarkWord" . i
      endif
      exe lastwinnr . "wincmd w"
      call winrestview(winview)
    endif
    let i = i + 1
  endwhile
  let g:mwLastSearched = ""
endfunction
" custom add <END>

" vim: ts=2 sw=2
