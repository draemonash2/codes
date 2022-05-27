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

function gr() {
	grep -nr "$@" .
}
function cdex() {
	\cd "$@"			# cdがaliasでループするので\をつける
	pwd
	ls -lFA --color=auto
}
function vimm() {
	sOpenPath=""
	for pathfile in ./$1*; do
		sOpenPath=$pathfile
	done
	vim $sOpenPath
}
function swap() {
	SUFFIX=swaptmp
	if [ $# -ge 2 ]; then
		FILE1=$1
		FILE2=$2
		\mv ./$1 ./$2.${SUFFIX}
		\mv ./$2 ./$1
		\mv ./$2.${SUFFIX} ./$2
	else
		echo "[error] Specify two or more arguments."
	fi
}
function tma() {
	if [ -z "$TMUX" ]; then
		if [ -n "${1}" ]; then
			tmux attach-session -t ${1} || tmux new-session -s ${1}
		else
			tmux attach-session || tmux new-session
		fi
	fi
}
function tmk() {
	if [ -n "${1}" ]; then
		tmux kill-session -t ${1}
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
alias vr='vim ~/.vimrc'
alias ir='vim ~/.inputrc; bind -f ~/.inputrc'
alias sr='vim ~/.screenrc'
alias tmc='vim ~/.tmux.conf'

alias tml='tmux list-sessions'

OSC52DIR="${HOME}/.vim/_plugins_user/osc52/plugin"
#echo ${TERMAPP}
if [ "${TERMAPP}" = "teraterm" ]; then
	\cp -f ${OSC52DIR}/osc52.vim{.teraterm,}
else
	\cp -f ${OSC52DIR}/osc52.vim{.org,}
fi
#ll ${OSC52DIR}

#########################################################
# Environment dependent settings
#########################################################
alias cdw='cd /mnt/c/;'
alias exp='explorer.exe .'		# open current directory with explorer.exe
alias sht='sudo shutdown -h now'

