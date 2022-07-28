# ~/.bashrc: executed by bash(1) for non-login shells.
# see /usr/share/doc/bash/examples/startup-files (in the package bash-doc)
# for examples

# If not running interactively, don't do anything
case $- in
    *i*) ;;
      *) return;;
esac

# don't put duplicate lines or lines starting with space in the history.
# See bash(1) for more options
HISTCONTROL=ignoreboth

# append to the history file, don't overwrite it
shopt -s histappend

# for setting history length see HISTSIZE and HISTFILESIZE in bash(1)
HISTSIZE=1000
HISTFILESIZE=2000

# check the window size after each command and, if necessary,
# update the values of LINES and COLUMNS.
shopt -s checkwinsize

# If set, the pattern "**" used in a pathname expansion context will
# match all files and zero or more directories and subdirectories.
#shopt -s globstar

# make less more friendly for non-text input files, see lesspipe(1)
[ -x /usr/bin/lesspipe ] && eval "$(SHELL=/bin/sh lesspipe)"

# set variable identifying the chroot you work in (used in the prompt below)
if [ -z "${debian_chroot:-}" ] && [ -r /etc/debian_chroot ]; then
    debian_chroot=$(cat /etc/debian_chroot)
fi

# set a fancy prompt (non-color, unless we know we "want" color)
case "$TERM" in
    xterm-color|*-256color) color_prompt=yes;;
esac

# uncomment for a colored prompt, if the terminal has the capability; turned
# off by default to not distract the user: the focus in a terminal window
# should be on the output of commands, not on the prompt
force_color_prompt=yes

if [ -n "$force_color_prompt" ]; then
    if [ -x /usr/bin/tput ] && tput setaf 1 >&/dev/null; then
	# We have color support; assume it's compliant with Ecma-48
	# (ISO/IEC-6429). (Lack of such support is extremely rare, and such
	# a case would tend to support setf rather than setaf.)
	color_prompt=yes
    else
	color_prompt=
    fi
fi

if [ "$color_prompt" = yes ]; then
	#[参考URL]https://zenn.dev/kotokaze/articles/bash-console
#   PS1='${debian_chroot:+($debian_chroot)}\[\033[01;32m\]\u@\h\[\033[00m\]:\[\033[01;34m\]\w\[\033[00m\]\$ '
	PS1='\n\[\e[37;45m\]\u@\h \[\e[32;47m\] \[\e[30;47m\]\D{%m/%d %H:%M:%S} \[\e[37;44m\] \w \[\e[00;36;49m\] \[\e[00m\]\$ '
#	PS1='${debian_chroot:+($debian_chroot)}\[\033[01;32m\]\u@\h\[\033[00m\]:\[\033[01;34m\]\w \$\[\033[00m\] '
else
    PS1='${debian_chroot:+($debian_chroot)}\u@\h->\t->\w \$ '
fi
unset color_prompt force_color_prompt

# If this is an xterm set the title to user@host:dir
case "$TERM" in
xterm*|rxvt*)
    PS1="\[\e]0;${debian_chroot:+($debian_chroot)}\u@\h: \w\a\]$PS1"
    ;;
*)
    ;;
esac

# enable color support of ls and also add handy aliases
if [ -x /usr/bin/dircolors ]; then
    test -r ~/.dircolors && eval "$(dircolors -b ~/.dircolors)" || eval "$(dircolors -b)"
    alias ls='ls --color=auto'
    #alias dir='dir --color=auto'
    #alias vdir='vdir --color=auto'

    alias grep='grep --color=auto'
    alias fgrep='fgrep --color=auto'
    alias egrep='egrep --color=auto'
fi

# colored GCC warnings and errors
#export GCC_COLORS='error=01;31:warning=01;35:note=01;36:caret=01;32:locus=01:quote=01'

# some more ls aliases
#alias ll='ls -alF'
#alias la='ls -A'
#alias l='ls -CF'

# Add an "alert" alias for long running commands.  Use like so:
#   sleep 10; alert
#alias alert='notify-send --urgency=low -i "$([ $? = 0 ] && echo terminal || echo error)" "$(history|tail -n1|sed -e '\''s/^\s*[0-9]\+\s*//;s/[;&|]\s*alert$//'\'')"'

# Alias definitions.
# You may want to put all your additions into a separate file like
# ~/.bash_aliases, instead of adding them here directly.
# See /usr/share/doc/bash-doc/examples in the bash-doc package.

if [ -f ~/.bash_aliases ]; then
    . ~/.bash_aliases
fi

# enable programmable completion features (you don't need to enable
# this, if it's already enabled in /etc/bash.bashrc and /etc/profile
# sources /etc/bash.bashrc).
if ! shopt -oq posix; then
  if [ -f /usr/share/bash-completion/bash_completion ]; then
    . /usr/share/bash-completion/bash_completion
  elif [ -f /etc/bash_completion ]; then
    . /etc/bash_completion
  fi
fi

HISTTIMEFORMAT='%F %T '

function is_tail_char_slash() {
	if [ $# -ne 1 ]; then
		echo "[error] is_tail_char_slash() argument error."
		return 1
	fi
	srcchr=${1}
	srcchrtmp=`echo "${srcchr}" | sed "s/\/$//g"`
	#echo ${srcchr} : ${srcchrtmp}
	if [ "${srcchr}" == "${srcchrtmp}" ]; then
		return 0 # tail char is not slash
	else
		return 1 # tail char is slash
	fi
}
function is_tail_char_slash_test() { #{{{
	is_tail_char_slash ./aaa/aaaaa
	if [ $? -ne 0 ]; then
		echo "[error] is_tail_char_slash_test() test error 02"
	fi
	is_tail_char_slash ./aaa/aaaaa/
	if [ $? -ne 1 ]; then
		echo "[error] is_tail_char_slash_test() test error 01"
	fi
} #}}}

function gr() {
	grep -nrIR "$@" --exclude={tags,GTAGS*,GRTAGS*} .
}
function cdex() {
	\cd "$@"			# cdがaliasでループするので\をつける
	pwd
	ls -lFA --color=auto
}
function vimall() {
	list=`find . -type f`
	vim $list
}
function vimdiffdir() {
	if [ $# -ne 2 ]; then
		echo "[error] specify two arguments."
		echo "  usage : vimdiffdir <file1> <file2>"
		return 1
	fi
	dir1=${1}
	dir2=${2}
	echo '***** diff -rq *****'
	diff -rq ${dir1} ${dir2}
	echo ''
	echo '***** vimdiff (only files with differences) *****'
	difflist=(`diff -rq ${dir1} ${dir2} | grep "Files " | sed -e 's/Files //' | sed -e 's/ and /:::::/' | sed -e 's/ differ//'`)
	for diffline in "${difflist[@]}"
	do
		file1=${diffline%:::::*}
		file2=${diffline#*:::::}
		echo "==> ${file1} vs ${file2} <=="
		vimdiff ${file1} ${file2}
	done
}
function swap() {
	suffix=swaptmp
	if [ $# -ne 2 ]; then
		echo "[error] specify two arguments."
		echo "  usage : swap <file/dir1> <file/dir2>"
		return 1
	fi
	file1=$1
	file2=$2
	\mv ./${file1} ./${file2}.${suffix}
	\mv ./${file2} ./${file1}
	\mv ./${file2}.${suffix} ./${file2}
}
function bak() {
	mode=1 # 1:Alphabet other:Time
	delimiter=_bak
	if [ $# -ne 1 ]; then
		echo "[error] specify one arguments."
		echo "  usage : bak <file/dir>"
		return 1
	fi
	infile=${1}
	if [ ! -e ${infile} ]; then
		echo "[error] \"${infile}\" does not exists."
		return 1
	fi
	if [ ${mode} -eq 1 ]; then
		nowsuffix=$(date '+%s' | awk '{print strftime("%y%m%d", $1)}')
		idxa=$(printf "%d" \'a)
		idxz=$(printf "%d" \'z)
		for ((i = ${idxa}; i <= ${idxz}; i++)) {
			char=$(printf "\x$(printf "%x" ${i})")
			outfile=${infile}${delimiter}${nowsuffix}${char}
			#echo ${outfile}
			if [ ! -e ${outfile} ]; then
				\cp -rf ${infile} ${outfile}
				break
			fi
		}
	else
		nowsuffix=$(date '+%s' | awk '{print strftime("%y%m%d-%H%M%S", $1)}')
		\cp -f ${INFILE} ${INFILE}${delimiter}${nowsuffix}
	fi
}
function lndir() {
	if [ $# -ne 2 ]; then
		echo "[error] specify one arguments."
		echo "  usage : lnhdir <srcdir> <dstdir>"
		return 1
	fi
	srcdirroot=${1}
	dstdirroot=${2}
	if [ ! -e ${srcdirroot} ]; then
		echo "[error] source directory \"${srcdirroot}\" does not exists."
		return 1
	fi
	if [ -e ${dstdirroot} ]; then
		echo "[error] destination directory \"${dstdirroot}\" exists."
		return 1
	fi
	is_tail_char_slash ${srcdirroot}
	if [ $? -ne 0 ]; then
		echo "[error] remove tail charactor \"/\" from \"${srcdirroot}\""
		return 1
	fi
	is_tail_char_slash ${dstdirroot}
	if [ $? -ne 0 ]; then
		echo "[error] remove tail charactor \"/\" from \"${dstdirroot}\""
		return 1
	fi
	export filelist=`find ${srcdirroot} -type f`
	for file in ${filelist}
	do
		file=${file/${srcdirroot}\//}
		dir=${file%/*}
		srcfile=${srcdirroot}/${file}
		dstdir=${dstdirroot}/${dir}
		dstfile=${dstdirroot}/${file}
		#echo ${dir} : ${file} : ${dstdir}
		#echo ${srcfile} : ${dstfile}
		mkdir -p ${dstdir}
		ln ${srcfile} ${dstfile}
	done
}
function tma() {
	if [ ! -z "$TMUX" ]; then
		echo "[error] cannot be run on tmux."
		return 1
	fi
	if [ $# -eq 1 ]; then
		session_name=${1}
		tmux attach-session -t ${session_name} || tmux new-session -s ${session_name}
	else
		tmux attach-session || tmux new-session
	fi
}
function tmk() {
	if [ $# -ne 1 ]; then
		echo "[error] specify one arguments."
		echo "  usage : tmk <session_name>"
		return 1
	fi
	session_name=${1}
	tmux kill-session -t ${session_name}
}
function killall() {
	#MAX=1
	MAX=`jobs | wc -l`
	#echo ${MAX}
	JOBNOS=""
	for NUM in `seq 1 ${MAX}`
	do
		JOBNOS="${JOBNOS} %${NUM}"
	done
	if [ "${JOBNOS}" != "" ]; then
		#echo ${JOBNOS}
		kill ${JOBNOS}
	fi
}

alias ll='ls -lFA --color=auto'
alias lln='ls -lFAv1 --color=auto'
alias la='ls -AF --color=auto'
alias l='ls -CF --color=auto'
alias ff='find . -type f | grep '
alias fd='find . -type d | grep '
(diff --help | grep -- "--color") &> /dev/null
if [ $? -eq 0 ]; then
	alias diff='\diff --color'
fi

alias cp='cp -i'
alias mv='mv -i'
alias rm='rm -i'
alias rmi='rm -i'

alias cd=cdex
alias ..='cd ..;'
alias ...='cd ../..;'
alias ....='cd ../../..;'
alias .....='cd ../../../..;'

alias br='vim ~/.bashrc; . ~/.bashrc'
alias bre='. ~/.bashrc'
alias vr='vim ~/.vimrc'
alias ir='vim ~/.inputrc; bind -f ~/.inputrc'
alias sr='vim ~/.screenrc'
alias tmc='vim ~/.tmux.conf'

alias tml='tmux list-sessions'
alias tmatemp='tma temp'

alias gitlo="git log -40 --oneline --graph --pretty=format:\"%Cred%ad%Creset ::::: %Cblue%h%Creset ::::: %Cgreen%an%Creset ::::: %C(yellow)%s\""

#########################################################
# Environment dependent settings
#########################################################
alias cdw='cd /mnt/c/;'
alias exp='explorer.exe .'		# open current directory with explorer.exe
alias sht='sudo shutdown -h now'

