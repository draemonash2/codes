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

# remove all aliases
unalias -a

# for PS1
function _update_ps1() { # {{{
	# [参考URL] https://zenn.dev/kotokaze/articles/bash-console
	color_black=0
	color_red=1
	color_green=2
	color_yellow=3
	color_blue=4
	color_magenta=5
	color_cyan=6
	color_white=7
	color_default=9
	if [ -f /.dockerenv ]; then
		bg_color=${color_magenta}
	else
		bg_color=${color_blue}
	fi
	PS1="\n"
	PS1="${PS1}\[\e[0;3${bg_color};04${bg_color}m\]!"					# head keywords
	PS1="${PS1}\[\e[0;37;04${bg_color}m\]\u@\h "						# user
	PS1="${PS1}\[\e[0;30;047m\] \D{%m/%d %H:%M:%S} "					# time
	PS1="${PS1}\[\e[0;37;04${bg_color}m\] \$(_puts_prompt_git_branch) "	# git branch
	PS1="${PS1}\[\e[0;30;047m\] \$(_puts_prompt_container_name) "		# docker container name
	PS1="${PS1}\[\e[0;97;100m\] \w "									# pwd
	PS1="${PS1}\[\e[0;30;040m\]!"										# tail keywords
	PS1="${PS1}\[\e[0;39;049m\]"										# reset
	PS1="${PS1}\n\$ "
} # }}}
show_prompt_branch_name=0
function _puts_prompt_git_branch() { # {{{
	if [ ${show_prompt_branch_name} -eq 1 ]; then
		branch_name=$(git branch --no-color 2>/dev/null | sed -ne "s/^\* \(.*\)$/\1/p")
		if [ ! "${branch_name}" = "" ]; then
			result_staged=`git status | grep "Changes to be committed:"`
			result_notstaged=`git status | grep "Changes not staged for commit:"`
			if [ "${result_staged}" != "" ] || [ "${result_notstaged}" != "" ]; then
				echo "${branch_name} +"
			else
				echo "${branch_name}"
			fi
		else
			echo "-"
		fi
	else
		echo " "
	fi
} # }}}
show_prompt_container_name=1
function _puts_prompt_container_name() { # {{{
	# Notes:
	#   This function requires placing a .dockercontainer with 
	#   the container name in the container's home directory
	if [ ${show_prompt_container_name} -eq 1 ]; then
		if [ -f /.dockerenv ]; then
			if [ -f ~/.dockercontainer ]; then
				echo "$(cat ~/.dockercontainer)"
			else
				echo "unknown container"
			fi
		else
			echo "host"
		fi
	else
		echo " "
	fi
} # }}}
if [ "$color_prompt" = yes ]; then
	_update_ps1
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
(diff --help | grep -- "--color") &> /dev/null
if [ $? -eq 0 ]; then
	alias diff='\diff --color'
fi

# colored GCC warnings and errors
#export GCC_COLORS='error=01;31:warning=01;35:note=01;36:caret=01;32:locus=01:quote=01'

# some more ls aliases
#alias ll='ls -alF'
#alias la='ls -A'
#alias l='ls -CF'
alias ll='ls -lFAv --color=auto --time-style="+%Y-%m-%d %H:%M"'
alias la='ls -AF --color=auto'
alias l='ls -CF --color=auto'
function _update_curdir() # {{{
{
	pwdold=${PWD}
	\cd
	\cd ${pwdold}
} # }}}
alias up='_update_curdir; ll'
function lll() # {{{
{
	if [ $# -ge 1 ]; then
		case $1 in
			"-fl")
				ls -F | grep -v / | sed s/\*$//g | sed s/\@$//g | sort
				;;
			"-lf")
				ls -F | grep -v / | sed s/\*$//g | sed s/\@$//g | sort
				;;
			"-f")
				ls -F | grep -v @ | grep -v / | sed s/\*$//g | sort
				;;
			"-l")
				ls -F | grep  @ | sed s/\@$//g | sort
				;;
			"-d")
				ls -F | grep / | sed s/\\/$//g | sort
				;;
			\*)
				echo "[error] unknown argument. $1"
				return 1
				;;
		esac
	else
		ls -F | sed s/\*$//g | sed s/\@$//g | sed s/\\/$//g | sort
	fi
}
	function _test_lll() # {{{
	{
		(
			# preprocess
			TEST_DIR=_test_lll_test
			mkdir ${TEST_DIR}
			cd ${TEST_DIR}
			
			mkdir 2_test_dir
			touch 3_test_file.md
			touch 4_test_file.sh
			ln -sf 3_test_file.md 1_test_symlink.md
			
			# test
			echo "### lll ###"
			lll
			echo "### lll -fl ###"
			lll -fl
			echo "### lll -lf ###"
			lll -lf
			echo "### lll -f ###"
			lll -f
			echo "### lll -l ###"
			lll -l
			echo "### lll -d ###"
			lll -d
			
			# post process
			cd ../
			rm -rf ${TEST_DIR}
		)
	} # }}}
# }}}
alias lllf="lll -lf"
alias llld="lll -d"

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
export LANG=en_US.UTF8

### Common
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

alias sht='sudo shutdown -h now'

function _is_tail_char_slash() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : _is_tail_char_slash <word>"
		return 1
	fi
	srcchr=${1}
	srcchrtmp=${srcchr: -1}
	if [ "${srcchrtmp}" == "/" ]; then
		return 1 # tail char is slash
	else
		return 0 # tail char is not slash
	fi
}
	function _test_is_tail_char_slash() { #{{{
		_is_tail_char_slash ./aaa/aaaaa
		if [ $? -ne 0 ]; then
			echo "[error] is_tail_char_slash_test() test error 02"
		fi
		_is_tail_char_slash ./aaa/aaaaa/
		if [ $? -ne 1 ]; then
			echo "[error] is_tail_char_slash_test() test error 01"
		fi
	} #}}}
 # }}}
function _remove_tail_slash() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : _remove_tail_slash <word>"
		return 1
	fi
	srcchr=${1}
	srcchrremoved=`echo "${srcchr}" | sed "s/\/$//g"`
	echo ${srcchrremoved}
}
	function _test_remove_char_slash() { #{{{
		str=./aaa/aaaaa
		result=`_remove_tail_slash ${str}`; echo ${result}
		result=`_remove_tail_slash ${str}`; echo ${result}
		str=/
		result=`_remove_tail_slash ${str}`; echo ${result}
		result=`_remove_tail_slash ${str}`; echo ${result}
		str=/aa
		result=`_remove_tail_slash ${str}`; echo ${result}
		result=`_remove_tail_slash ${str}`; echo ${result}
	} #}}}
 # }}}
function _output_ps1_color_palette() { # {{{
	printf "\n === PS1 color palette ===\n"
	type[0]="none        "
	type[1]="bold        "
	type[2]="half-bright "
	type[4]="underscore  "
	type[5]="blink       "
	type[7]="reverse     "
	echo ${seasons[3]}
	for i in 0 1 2 4 5 7 ; do
		for j in `seq 30 38` `seq 90 97` ; do
			printf "${type[i]} "
			for k in `seq 40 47` `seq 100 107` ; do
				printf "\033[${i};${j};${k}m"
				printf " ${i};${j};${k} "
				printf "\033[0;39;49m"  # set default
			done
			printf "\n"
		done
	done
	printf "\n"
} # }}}
function _output_color_palette_clridx() { # {{{
	bgclridx=${1:-0}
	echo "=== colour palette color idx ==="
	echo "  i.e. set -g status-style \"fg=colour???,bg=colour${bgclridx}\""
	for fgclridx in {0..255}; do
		idxmod=`expr $(expr ${fgclridx} - 16) % 36`
		#echo ${idxmod}
		if [ ${idxmod} -eq 0 ]; then
			printf "\n"
		fi
		bgclrstr="\x1b[48;5;${bgclridx}m"
		fgclrstr="\x1b[38;5;${fgclridx}m"
		fgoutstr="$(printf "%03d\n" "${fgclridx}")"
		printf "${bgclrstr}${fgclrstr} ${fgoutstr} \x1b[0m"
	done
} # }}}
function _is_exist_str() { # {{{
	# usage: _is_exist_str <target> <search_word>
	#          _is_exist_str aaa_atest ates  # -> 0 (exist)
	#          _is_exist_str aaa_atest btes  # -> 1 (not exist)
	target=$1
	searchword=$2
	if [[ "${target}" =~ "${searchword}" ]]; then
		return 0 # exist
	else
		return 1 # not exist
	fi
}
	function _test_is_exist_str() { #{{{
		_is_exist_str aaa_atest atest; echo $?  # -> 0 (exist)
		_is_exist_str aaa_atest btest; echo $?  # -> 1 (not exist)
		_is_exist_str aaa aaa; echo $?          # -> 0 (exist)
	} #}}}
# }}}
function _is_str_pos_head() { # {{{
	# usage: _is_str_pos_head <target> <search_word>
	#          _is_str_pos_head aaa_atest aaa_   # -> 0 (head)
	#          _is_str_pos_head aaa_atest aa_    # -> 1 (not head)
	#          _is_str_pos_head aaa_atest ab     # -> 1 (not exist)
	target=$1
	searchword=$2
	if [[ "${target}" =~ "${searchword}" ]]; then
		matchpos=$(expr length \( ${target} : "\(.*\)${searchword}" \))
		if [ ${matchpos} = 0 ]; then
			return 0 # head
		else
			return 1 # not head
		fi
	else
		return 1 # not exist
	fi
}
	function _test_is_str_pos_head() { #{{{
		_is_str_pos_head aaa_atest aaa_; echo $?  # -> 0 (head)
		_is_str_pos_head aaa_atest aa_; echo $?   # -> 1 (not head)
		_is_str_pos_head aaa_atest ab; echo $?    # -> 1 (not exist)
		_is_str_pos_head aaa aaa; echo $?         # -> 0 (head)
	} #}}}
# }}}
function _is_str_pos_tail() { # {{{
	# usage: _is_str_pos_tail <target> <search_word>
	#          _is_str_pos_tail aaa_atest atest  # -> 0 (tail)
	#          _is_str_pos_tail aaa_atest ates   # -> 1 (not tail)
	#          _is_str_pos_tail aaa_atest btest  # -> 1 (not exist)
	target=$1
	searchword=$2
	if [[ "${target}" =~ "${searchword}" ]]; then
		matchpos=$(expr length \( ${target} : "\(.*\)${searchword}" \))
		target_len=${#target}
		searchword_len=${#searchword}
		if [ ${target_len} = $((${matchpos} + ${searchword_len})) ]; then
			return 0 # tail
		else
			return 1 # not tail
		fi
	else
		return 1 # not exist
	fi
}
	function _test_is_str_pos_tail() { #{{{
		_is_str_pos_tail aaa_atest atest; echo $?  # -> 0 (tail)
		_is_str_pos_tail aaa_atest ates; echo $?   # -> 1 (not tail)
		_is_str_pos_tail aaa_atest btest; echo $?  # -> 1 (not exist)
		_is_str_pos_tail aaa aaa; echo $?          # -> 0 (tail)
	} #}}}
# }}}
function _get_scp_config() { # {{{
	# [config file format]
	#   hostname<tab>username<tab>password
	#   e.g.
	#     $ cat ~/_config_scp_a
	#     192.168.12.11<tab>endo<tab>pw1234
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : _get_scp_config <partnername>"
		echo "    <partnername> partner name"
		return 1
	fi
	partnername=$1
	config_file=~/_config_scp_${partnername}
	if [ ! -f ${config_file} ]; then
		echo "[error] ${config_file} does not exist."
		return 1
	fi
	host=$(cut -f 1 ${config_file} | head -n 1)
	user=$(cut -f 2 ${config_file} | head -n 1)
	password=$(cut -f 3 ${config_file} | head -n 1)
} # }}}
function gr() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : gr <keyword>"
		return 1
	fi
	grep --no-messages -nrIR "$@" --exclude={tags,GTAGS*,GRTAGS*} .
} # }}}
function grw() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : grw <keyword>"
		return 1
	fi
	grep --no-messages -nrIRw "$@" --exclude={tags,GTAGS*,GRTAGS*} .
} # }}}
function cdex() { # {{{
	\cd "$@"			# cdがaliasでループするので\をつける
	pwd
	ll
} # }}}
function vimall() { # {{{
	if [ $# -eq 0 ]; then
		list=`find . -type f`
	else
		list=`find . -type f -name $1`
	fi
	vim $list
} # }}}
function vimdiffdir() { # {{{
	if [ $# -ne 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : vimdiffdir <dir1> <dir2>"
		return 1
	fi
	dir1=${1}
	dir2=${2}
	echo '***** diff -rq *****'
	diff -rq ${dir1} ${dir2}
	echo ''
	echo '***** vimdiff (only files with differences) *****'
	difflist=(`LANG=C; diff -rq ${dir1} ${dir2} | grep "Files " | grep -v "/.git/" | sed -e 's/Files //' | sed -e 's/ and /:::::/' | sed -e 's/ differ//'`)
	for diffline in "${difflist[@]}"
	do
		file1=${diffline%:::::*}
		file2=${diffline#*:::::}
		echo "==> ${file1} vs ${file2} <=="
		vimdiff ${file1} ${file2}
		sleep 1
	done
} # }}}
function vimdiffcheck() { # {{{
	# Compare with vimdiff only when there is a difference
	if [ $# -lt 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : vimdiffcheck <file1> <file2> [<sleeptime>]"
		return 1
	fi
	file1=$1
	file2=$2
	sleeptime=${3:-0}
	echo "==> ${file1} vs ${file2} <=="
	diff ${file1} ${file2} &> /dev/null
	if [ $? -eq 1 ]; then
		vimdiff ${file1} ${file2}
		sleep ${sleeptime}
	fi
} # }}}
function swap() { # {{{
	suffix=swaptmp
	if [ $# -ne 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : swap <file/dir1> <file/dir2>"
		return 1
	fi
	file1=$1
	file2=$2
	\mv ./${file1} ./${file2}.${suffix}
	\mv ./${file2} ./${file1}
	\mv ./${file2}.${suffix} ./${file2}
} # }}}
function bak() { # {{{
	mode=1 # 1:Alphabet other:Time
	delimiter=.bak
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
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
} # }}}
function lndir() { # {{{
	if [ $# -ne 2 ]; then
		echo "[error] wrong number of arguments."
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
	_is_tail_char_slash ${srcdirroot}
	if [ $? -ne 0 ]; then
		echo "[error] remove tail charactor \"/\" from \"${srcdirroot}\""
		return 1
	fi
	_is_tail_char_slash ${dstdirroot}
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
} # }}}
function killjobsall() { # {{{
	jobidlist=$(jobs | cut -d] -f -1 | cut -d[ -f 2-)
	for jobid in ${jobidlist}
	do
		#echo ${jobid}
		kill -9 %${jobid}
		wait %${jobid} 2>/dev/null
	done
} # }}}
function killprocessall() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : killprocessall <keyword>"
		return 1
	fi
	keyword=$1
	pidlist=$(ps a -u ${USER} | grep -- "${keyword}" | grep -v "grep " | sed 's/^[ \t]*//' | cut -d" " -f 1)
	for pid in ${pidlist}
	do
		echo ${pid}
		kill -9 ${pid}
		wait ${pid} 2>/dev/null
	done
} # }}}
function cpd() { # {{{
	if [ $# -ne 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : cpd <src> <dst>"
		return 1
	fi
	srcpathraw=${1}
	dstpathraw=${2}
	srcpath=`_remove_tail_slash ${srcpathraw}`
	dstpath=`_remove_tail_slash ${dstpathraw}`
	dstpardirpath=${dstpath%/*}
	echo ${dstpardirpath}
	if [ -f ${srcpath} ] || [ -d ${srcpath} ]; then
		if [ ! -d ${dstpardirpath} ]; then
			echo "mkdir -p ${dstpardirpath}"
			mkdir -p ${dstpardirpath}
		fi
		echo "\cp -rf ${srcpath} ${dstpath}"
		\cp -rf ${srcpath} ${dstpath}
		return 0
	else
		echo "[error] specified path does not exists."
		return 1
	fi
}
	function _test_cpd() { # {{{
		srcpath="${HOME}/.bashrc"
		dstpath="${HOME}/test/aaa/.bashrc"
	#	echo ${srcpath}
	#	echo ${dstpath}
		cpd ${srcpath} ${dstpath}
		ll ${dstpath}
	} # }}}
# }}}
function catrange() { # {{{
	if [ $# -ne 3 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : catrange <file> <lineno_head> <lineno_tail>"
		return 1
	fi
	file=${1}
	lineno_head=${2}
	lineno_tail=${3}
	if [ ! -f ${file} ]; then
		echo "[error] ${file} does not exist."
		return 1
	fi
	if [ ${lineno_tail} -lt ${lineno_head} ]; then
		echo "[error] arguments error. tail(${lineno_tail}) < head(${lineno_head})"
		return 1
	fi
	if [ ${lineno_head} -lt 1 ]; then
		echo "[error] arguments error. head(${lineno_head}) < 1)"
		return 1
	fi
	file_line_num=`cat ${file} | wc -l`
	if [ ${file_line_num} -lt ${lineno_tail} ]; then
		echo "[error] arguments error. tail(${lineno_tail}) > file_line_num(${file_line_num})"
		return 1
	fi
	linenum=`expr ${lineno_tail} - ${lineno_head} + 1`
	#echo ${lineno_head}
	#echo ${lineno_tail}
	#echo ${linenum}
	cat ${file} | head -n ${lineno_tail} | tail -n ${linenum}
	return 0
}
	function _test_catrange() { # {{{
		inputfile=test.txt
		echo 1 >  ${inputfile}
		echo 2 >> ${inputfile}
		echo 3 >> ${inputfile}
		echo 4 >> ${inputfile}
		echo 5 >> ${inputfile}
		
		echo "=== test start ==="
		cmd="catrange tst2.txt 2 3";     echo "# ${cmd}"; ${cmd}; echo ""
		cmd="catrange ${inputfile} 2";   echo "# ${cmd}"; ${cmd}; echo ""
		cmd="catrange ${inputfile} 0 3"; echo "# ${cmd}"; ${cmd}; echo ""
		cmd="catrange ${inputfile} 4 6"; echo "# ${cmd}"; ${cmd}; echo ""
		
		cmd="catrange ${inputfile} 2 3"; echo "# ${cmd}"; ${cmd}; echo ""
		cmd="catrange ${inputfile} 1 5"; echo "# ${cmd}"; ${cmd}; echo ""
		cmd="catrange ${inputfile} 1 3"; echo "# ${cmd}"; ${cmd}; echo ""
		cmd="catrange ${inputfile} 4 5"; echo "# ${cmd}"; ${cmd}; echo ""
		echo "=== test finished ==="
		
		rm -f ${inputfile}
	} # }}}
# }}}
function ff() { # {{{
	if [ $# -eq 0 ]; then
		find . -type f 2> /dev/null
		find . -type l 2> /dev/null
	else
		find . -type f -name $1 2> /dev/null
		find . -type l -name $1 2> /dev/null
	fi
} # }}}
function fd() { # {{{
	if [ $# -eq 0 ]; then
		find . -type d 2> /dev/null
	else
		find . -type d -name $1 2> /dev/null
	fi
} # }}}
function outputleafdirlist() { # {{{
	dirlist=$(find . -type d)
	for dir in $dirlist
	do
		#echo $dir
		directory_num=$(find "${dir}" -maxdepth 1 -type d | wc -l)
		if [ ${directory_num} -eq 1 ]; then
			echo "${dir}"
		fi
	done
} # }}}
function path() { # {{{
	enable_abs=0
	args=()
	while (( $# > 0 ))
	do
		case $1 in
			-h | --help)
				echo "[summary]"
				echo "  Displays file/directory path. Copy to clipboard if possible."
				echo ""
				echo "[usage]"
				echo "  path [-a] [<file_dir_path>...]"
				echo ""
				echo "[option]"
				echo "  -a, --absolute: replace \${USER} to home directory full path."
				return 0
				;;
			-a | --absolute)
				enable_abs=1
				;;
			-*)
				echo "[error] invalid option: \"$1\". see options by --help."
				return 1
				;;
			*)
				args+=("$1")
				;;
		esac
		shift
	done
	curdir=${PWD}
	if [ ${enable_abs} -eq 0 ]; then
		curdir=${curdir//\/home\/${USER}/'${HOME}'}
	fi
	if [ ${#args[@]} -ge 1 ]; then
		path=${curdir}/${args[0]}
	else
		path=${curdir}
	fi
	set_clipboard "${path}"
	echo ${path}
}
	complete -F _complete_path path # {{{
	function _complete_path() { local cur prev; _get_comp_words_by_ref -n : cur prev; COMPREPLY=( $(compgen -f -- "${cur}") );} # }}}
# }}}
function outputencodesall() # {{{
{
	# This function requires "nkf" command.
	filelist=$(ff)
	for file in $filelist
	do
		encode=$(nkf --guess ${file})
		echo "${file} : ${encode}"
	done
} # }}}
function lsscpdata() { # {{{
	trgtdir=~/_scp_to_xxx
	echo "$ ll ${trgtdir}"; ll ${trgtdir};
	trgtdir=~/_scp_from_xxx
	echo "$ ll ${trgtdir}"; ll ${trgtdir};
} # }}}
function storescpsenddata() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : storescpsenddata <file/dir>"
		return 1
	fi
	trgtdir=~/_scp_to_xxx
	mkdir -p ${trgtdir}
	\cp -rf $1 ${trgtdir}/.
	lsscpdata
} # }}}
function clearscpsenddata() { # {{{
	trgtdir=~/_scp_to_xxx
	rm -rf ${trgtdir}/*
	trgtdir=~/_scp_from_xxx
	rm -rf ${trgtdir}/*
	lsscpdata
} # }}}
function sendscp() { # {{{
	if [ $# -lt 5 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : sendscp <host> <user> <password> <partnerdir> <myobj> [<myobj>...]"
		echo "    <partnerdir> partner directory path (absolute path)"
		echo "    <myobj>      my object path (absolute/relative path)"
		return 1
	fi
	argv=("$@")
	host=${argv[0]}
	user=${argv[1]}
	password=${argv[2]}
	partnerdir=${argv[3]}
	for i in $(seq 4 $(($# - 1)))
	do
		myobj=${argv[$i]}
		expect -c "spawn scp -r ${myobj} ${user}@${host}:${partnerdir} ; expect password: ; send ${password}\r ; expect $ ; interact"
	done
} # }}}
function fetchscp() { # {{{
	if [ $# -lt 5 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : fetchscp <host> <user> <password> <mydir> <partnerobj> [<partnerobj>...]"
		echo "    <mydir>       my directory path (absolute/relative path)"
		echo "    <partnerobj>  partner object path (absolute path)"
		return 1
	fi
	argv=("$@")
	host=${argv[0]}
	user=${argv[1]}
	password=${argv[2]}
	mydir=${argv[3]}
	for i in $(seq 4 $(($# - 1)))
	do
		partnerobj=${argv[$i]}
		expect -c "spawn scp -r ${user}@${host}:${partnerobj} ${mydir} ; expect password: ; send ${password}\r ; expect $ ; interact"
	done
} # }}}
function syncscp() { # {{{
	if [ $# -ne 5 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : syncscp <host> <user> <password> <partnerfile> <myfile>"
		echo "    <partnerfile>  partner file path (relative path from home dir)"
		echo "    <myfile>       my file path (relative path from home dir)"
		return 1
	fi
	mytmpdir=~/_scp_from_xxx
	host=$1
	user=$2
	password=$3
	partnerfile=$4
	myfile=$5
	fetchscp ${host} ${user} ${password} ${mytmpdir} /home/${user}/${partnerfile}
	diff ${mytmpdir}/${partnerfile} ~/${myfile} &> /dev/null
	if [ $? -eq 1 ]; then
		vimdiff ${mytmpdir}/${partnerfile} ~/${myfile}
	fi
#	sendscp ${host} ${user} ${password} /home/${user} ~/${myfile}
	rm -rf ${mytmpdir}/${partnerfile}
} # }}}
function sendscpto() { # {{{
	if [ $# -lt 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : sendscpto <partnername> <myobj> [<myobj>...]"
		echo "    <partnername> partner name"
		echo "    <myobj>       my object path (absolute/relative path)"
		return 1
	fi
	argv=("$@")
	partnername=${argv[0]}
	_get_scp_config ${partnername}
	if [ $? -eq 1 ]; then
		return 1
	fi
	partnerdir=/home/${user}/_scp_from_xxx
	for i in $(seq 1 $(($# - 1)))
	do
		myobj=${argv[$i]}
		sendscp ${host} ${user} ${password} ${partnerdir} ${myobj}
	done
} # }}}
function fetchscpfrom() { # {{{
	if [ $# -lt 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : fetchscpfrom <partnername> <partnerobj> [<partnerobj>...]"
		echo "    <partnername> partner name"
		echo "    <partnerobj>  partner object path (absolute path)"
		return 1
	fi
	argv=("$@")
	partnername=${argv[0]}
	_get_scp_config ${partnername}
	if [ $? -eq 1 ]; then
		return 1
	fi
	mydir=~/_scp_from_xxx
	for i in $(seq 1 $(($# - 1)))
	do
		partnerobj=${argv[$i]}
		fetchscp ${host} ${user} ${password} ${mydir} ${partnerobj}
	done
	echo "$ ll ${mydir}"; ll ${mydir};
} # }}}
function syncscpto() { # {{{
	if [ $# -ne 3 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : syncscpto <partnername> <partnerfile> <myfile>"
		echo "    <partnername> partner name"
		echo "    <partnerfile> partner file path (relative path from home dir)"
		echo "    <myfile>      my file path (relative path from home dir)"
		return 1
	fi
	partnername=$1
	partnerfile=$2
	myfile=$3
	_get_scp_config ${partnername}
	if [ $? -eq 1 ]; then
		return 1
	fi
	syncscp ${host} ${user} ${password} ${partnerfile} ${myfile}
} # }}}
function syncdotfiles() { # {{{
	file=".bashrc";					syncscpto a ${file} ${file}
	file=".gdbinit";				syncscpto a ${file} ${file}
	file=".inputrc";				syncscpto a ${file} ${file}
	file=".tigrc";					syncscpto a ${file} ${file}
	file=".tmux.conf";				syncscpto a ${file} ${file}
	file=".tmux.conf.mac.conf";		syncscpto a ${file} ${file}
	file=".tmux.conf.ubuntu.conf";	syncscpto a ${file} ${file}
	file=".vimrc";					syncscpto a ${file} ${file}
} # }}}
function convunixtimetodate() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : convunixtimetodate <unixtime>"
		return 1
	fi
	unixtime=$1
#	echo ${unixtime} | awk '{print strftime("%c",$1)}'
	date -d @${unixtime} +"%Y-%m-%d %H:%M:%S"
} # }}}
function aggregate() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : aggregate <file>"
		return 1
	fi
	filename=$1
	tmpfile=aggregated.tmp
	cat ${filename} | sort -n > ${tmpfile}
	
	min=$(cat ${tmpfile} | head -1)
	max=$(cat ${tmpfile} | tail -1)
	mid=$(cat ${tmpfile} | awk '{v[i++]=$1;}END {x=int((i+1)/2); if(x<(i+1)/2) print (v[x-1]+v[x])/2; else print v[x-1];}')
	avg=$(cat ${tmpfile} | awk '{x++;sum+=$1}END {print sum/x}')
	stddev=$(awk '{ x[NR] = $1 } END{ if(NR == 0) exit; for(i in x){ sum_x += x[i]; } m_x = sum_x / NR; for(i in x){ sum_dx2 += ((x[i] - m_x) ^ 2); } print sqrt(sum_dx2 / NR); }' ${tmpfile})
	
#	echo "[filename] [min] [max] [mid] [avg] [stddev]"
	echo "${filename} ${min} ${max} ${mid} ${avg} ${stddev}"
	
	rm -f ${tmpfile}
} # }}}
function outputhwinfo() { # {{{
	if [ "$(uname)" == 'Darwin' ]; then
		:
	else
		logfile=~/_hwinfo.log
		rm -f ${logfile}
		cmd="cat /etc/lsb-release"														; echo "### ${cmd}" 1>> ${logfile} 2>> ${logfile}; ${cmd} 1>> ${logfile} 2>> ${logfile}
		cmd="cat /proc/cpuinfo"															; echo "### ${cmd}" 1>> ${logfile} 2>> ${logfile}; ${cmd} 1>> ${logfile} 2>> ${logfile}
		cmd="cat /proc/meminfo"															; echo "### ${cmd}" 1>> ${logfile} 2>> ${logfile}; ${cmd} 1>> ${logfile} 2>> ${logfile}
		cmd="df -h"																		; echo "### ${cmd}" 1>> ${logfile} 2>> ${logfile}; ${cmd} 1>> ${logfile} 2>> ${logfile}
		cmd="lspci | grep VGA"															; echo "### ${cmd}" 1>> ${logfile} 2>> ${logfile}; ${cmd} 1>> ${logfile} 2>> ${logfile}
		cmd="nvidia-smi"																; echo "### ${cmd}" 1>> ${logfile} 2>> ${logfile}; ${cmd} 1>> ${logfile} 2>> ${logfile}
		cmd="nvidia-smi --query-gpu=name --format=csv,noheader"							; echo "### ${cmd}" 1>> ${logfile} 2>> ${logfile}; ${cmd} 1>> ${logfile} 2>> ${logfile}
		cmd="clinfo"																	; echo "### ${cmd}" 1>> ${logfile} 2>> ${logfile}; ${cmd} 1>> ${logfile} 2>> ${logfile}
		cmd="nvcc -V # CUDA version"													; echo "### ${cmd}" 1>> ${logfile} 2>> ${logfile}; ${cmd} 1>> ${logfile} 2>> ${logfile}
		cmd="cat /usr/include/cudnn_version.h | grep CUDNN_MAJOR -A 2 # cudnn version"	; echo "### ${cmd}" 1>> ${logfile} 2>> ${logfile}; ${cmd} 1>> ${logfile} 2>> ${logfile}
		cmd="sudo lshw"																	; echo "### ${cmd}" 1>> ${logfile} 2>> ${logfile}; ${cmd} 1>> ${logfile} 2>> ${logfile}
		cmd="uname -m"																	; echo "### ${cmd}" 1>> ${logfile} 2>> ${logfile}; ${cmd} 1>> ${logfile} 2>> ${logfile}
		cmd="arch"																		; echo "### ${cmd}" 1>> ${logfile} 2>> ${logfile}; ${cmd} 1>> ${logfile} 2>> ${logfile}
		cmd="gcc -v"																	; echo "### ${cmd}" 1>> ${logfile} 2>> ${logfile}; ${cmd} 1>> ${logfile} 2>> ${logfile}
		vim ${logfile}
	fi
} # }}}
function greprep() { # {{{
	if [ $# -ne 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : greprep <source_keyword> <destination_keyword>"
		return 1
	fi
	src=$1
	dst=$2
	
	echo "### before ###"
	echo "# source"
	gr "${src}"
	echo ""
	echo "# destination"
	gr "${dst}"
	echo ""
	
	grep -lr "${src}" | xargs sed -i -e "s/${src}/${dst}/g"
	
	echo "### after ###"
	echo "# source"
	gr "${src}"
	echo ""
	echo "# destination"
	gr "${dst}"
	echo ""
} # }}}
function convertimg() { # {{{
	if [ $# -ne 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : convertimg <size or ratio> <file_path>"
		return 1
	fi
	size=$1
	file=$2
	filebase=${file%%.*}
	fileext=${file#*.}
	bakfilebase=${filebase}_
	while [ -f ${bakfilebase}.${fileext} ]
	do
		bakfilebase=${bakfilebase}_
	done
	bakfile=${bakfilebase}.${fileext}
	\cp -f ${file} ${bakfile}
	dst=${file}
	convert -geometry "${size}" ${bakfile} ${file}
} # }}}
function renamedirfiles() { # {{{
	# Rename all files and directories under the current directory.
	if [ $# -ne 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : renamedirfiles <source_keyword> <destination_keyword>"
		return 1
	fi
	srckeyword=$1
	dstkeyword=$2
	
	# output before
	echo "### before ###"
	find . -type f
	echo ""
	
	# calc max layer number
	maxlayernum=0
	dirlist=$(find . -type d)
	for dir in ${dirlist}
	do
		dir_nodelim=${dir//\//}
		#echo ${dir}
		#echo ${dir_nodelim}
		layernum=$((${#dir}-${#dir_nodelim}))
		#echo ${layernum}
		if [ ${layernum} -gt ${maxlayernum} ]; then
			maxlayernum=${layernum}
		fi
	done
	#echo ${maxlayernum}
	
	# rename directorys
	for layernum in $(seq 1 ${maxlayernum})
	do
		dirlist=$(find . -mindepth ${layernum} -maxdepth ${layernum} -type d)
		#echo ${dirlist}
		for dir in ${dirlist}
		do
			srcdir=${dir}
			dstdir=${dir//${srckeyword}/${dstkeyword}}
			#echo "mv -f ${srcdir} ${dstdir} 2> /dev/null"
			mv -f ${srcdir} ${dstdir} 2> /dev/null
		done
		#echo ""
	done
	
	# rename files and symboliclinks
	typelist=("f" "l")
	for type in ${typelist[@]}
	do
		#echo ${type}
		filelist=$(find . -type ${type})
		#echo ${filelist}
		for file in ${filelist}
		do
			srcfile=${file}
			dstfile=${file//${srckeyword}/${dstkeyword}}
			#echo "mv -f ${srcfile} ${dstfile} 2> /dev/null"
			mv -f ${srcfile} ${dstfile} 2> /dev/null
		done
	done
	
	# output after
	echo "### after ###"
	find . -type f
	echo ""
} # }}}
function setenv() { # {{{
	# Set environment variable without duplication
	if [ $# -ne 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : setenv <env_var_name> <env_value>"
		return 1
	fi
	env_var_name=$1
	env_value_target=$2
	env_value_base=$(printenv ${env_var_name})
	if [ -z ${env_value_base} ]; then
		export ${env_var_name}="${env_value_target}"
	else
		_true=0
		_false=1
		_is_str_pos_head "${env_value_base}" "${env_value_target}:"
		is_exist_head="$?"
		_is_str_pos_tail "${env_value_base}" ":${env_value_target}"
		is_exist_tail="$?"
		_is_exist_str "${env_value_base}" ":${env_value_target}:"
		is_exist_mid="$?"
		
		if [ "${env_value_target}" = "${env_value_base}" ]; then
			:
		else
			if [ "${is_exist_head}" = "${_true}" ]; then
				:
			else
				if [ "${is_exist_mid}" = "${_true}" ]; then
					:
				else
					if [ "${is_exist_tail}" = "${_true}" ]; then
						:
					else
						export ${env_var_name}="${env_value_target}:${env_value_base}"
					fi
				fi
			fi
		fi
	fi
#	echo "${env_var_name}=$(printenv ${env_var_name})"
}
	function _test_setenv() { # {{{
		echo ""
		export TESTENV=
		setenv TESTENV aaa
		echo ""
		setenv TESTENV aaa
		echo ""
		setenv TESTENV bbb
		echo ""
		setenv TESTENV bbb
		echo ""
		setenv TESTENV ccc
		echo ""
		setenv TESTENV ccc
		
		echo ""
		export TESTENV=aaa:bbb:ccc
		echo ""
		setenv TESTENV aaa
		echo ""
		setenv TESTENV bbb
		echo ""
		setenv TESTENV ccc
		
	} # }}}
# }}}
function unsetenv() { # {{{
	# Unset environment variable
	if [ $# -ne 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : unsetenv <env_var_name> <env_value>"
		return 1
	fi
	env_var_name=$1
	env_value_target=$2
	env_value_base=$(printenv ${env_var_name})
	if [ -z ${env_value_base} ]; then
		:
	else
		_true=0
		_false=1
		_is_str_pos_head "${env_value_base}" "${env_value_target}:"
		is_exist_head="$?"
		_is_str_pos_tail "${env_value_base}" ":${env_value_target}"
		is_exist_tail="$?"
		_is_exist_str "${env_value_base}" ":${env_value_target}:"
		is_exist_mid="$?"
		
		if [ "${env_value_target}" = "${env_value_base}" ]; then
			export ${env_var_name}=""
		else
			if [ "${is_exist_head}" = "${_true}" ]; then
				export ${env_var_name}="${env_value_base//${env_value_target}:/}"
			else
				if [ "${is_exist_mid}" = "${_true}" ]; then
					export ${env_var_name}="${env_value_base//:${env_value_target}:/:}"
				else
					if [ "${is_exist_tail}" = "${_true}" ]; then
						export ${env_var_name}="${env_value_base//:${env_value_target}/}"
					else
						:
					fi
				fi
			fi
		fi
	fi
#	echo "${env_var_name}=$(printenv ${env_var_name})"
}
	function _test_unsetenv() { # {{{
		echo ""
		export TESTENV=aaa:bbb:ccc
		unsetenv TESTENV aaa
		
		echo ""
		export TESTENV=aaa:bbb:ccc
		unsetenv TESTENV bbb
		
		echo ""
		export TESTENV=aaa:bbb:ccc
		unsetenv TESTENV ccc
		
		
		echo ""
		echo ""
		export TESTENV=
		unsetenv TESTENV aaa
		
		echo ""
		export TESTENV=aaa:bbb:ccc
		unsetenv TESTENV ddd
		
		echo ""
		export TESTENV=aaa:bbb:ccc
		unsetenv TESTENV aa
		
		echo ""
		export TESTENV=aaa:bbb:ccc
		unsetenv TESTENV bb
		
		echo ""
		export TESTENV=aaa:bbb:ccc
		unsetenv TESTENV cc
		
		echo ""
		export TESTENV=aaa:bbb:ccc
		unsetenv TESTENV aaaa
	} # }}}
# }}}
function echopath() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : echoenv <env_var_name>"
		return 1
	fi
	env_var_name=$1
	env_value=$(printenv ${env_var_name})
	echo ${env_value} | sed "s/:/\n/g"
} # }}}
function avim() { # {{{
	# Colorize ANSI color code on vim
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : avim <file>"
		return 1
	fi
	file=$1
	while true
	do
		vim -c ":term ++hidden ++curwin ++open cat ${file}"
		echo "To exit, press ctrl-c."
		sleep 1
	done
}
	complete -F _complete_avim avim # {{{
	function _complete_avim() { local cur prev; _get_comp_words_by_ref -n : cur prev; COMPREPLY=( $(compgen -f -- "${cur}") );} # }}}
# }}}
function vimw() { # {{{
	# Open file on vim with no syntax (=White VIM)
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : vimw <file>"
		return 1
	fi
	file=$1
	vim "+syntax off" ${file}
}
	complete -F _complete_vimw vimw # {{{
	function _complete_vimw() { local cur prev; _get_comp_words_by_ref -n : cur prev; COMPREPLY=( $(compgen -f -- "${cur}") );} # }}}
# }}}
function viml() { # {{{
	# Launch vim with a file path that includes line numbers.
	#   e.g. viml file.txt:115: -> vim file.txt -c 115
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : viml <file:line>"
		return 1
	fi
	file_line=$1
	file=${file_line%%:*}
	line_tmp=${file_line#*:}
	line=${line_tmp%%:*}
	#echo ${file} ${line_tmp} ${line}
	vim ${file} -c ${line}
} # }}}
function set_clipboard() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : set_clipboard <string>"
		return 1
	fi
	string=$1
	
	xclip --version &> /dev/null
	if [ $? -eq 0 ]; then
		echo -n "${string}" | xclip
		return 0
	fi
	
	xsel --version &> /dev/null
	if [ $? -eq 0 ]; then
		echo -n "${string}" | xsel --clipboard --input
		return 0
	fi
}
# }}}
function extractdefine() { # {{{
	# [usage] 
	# usage : extract_define_range <infile> <outfile> <define_keyword> <remain_target_side>
	#    <remain_target_side>
	#       true  : remain "true" side
	#       false : remain "false" side
	#         e.g. specified "true" side
	#           #if AAA           #del
	#              true side      #remain
	#           #else /* AAA */   #del
	#              false side     #del
	#           #endif /* AAA */  #del
	#
	#           #ifdef AAA        #del
	#              true side      #remain
	#           #else /* AAA */   #del
	#              false side     #del
	#           #endif /* AAA */  #del
	#
	#           #ifndef AAA       #del
	#              true side      #del
	#           #else /* !AAA */  #del
	#              false side     #remain
	#           #endif /* !AAA */ #del
	if [ -z ${DEVDIR} ]; then
		echo "[error] \${DEVDIR} must be set."
		return 1
	fi
	scriptpath=${DEVDIR}/_script/python/extract_define_range.py
	python3 ${scriptpath} "$@"
} # }}}
function inode() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : inode <file>"
		return 1
	fi
	file=$1
	inode=$(ls -lai ${file} | cut -d" " -f 1)
	echo ${inode}
} # }}}
function diffinode() { # {{{
	if [ $# -ne 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : diffinode <file1> <file2>"
		return 1
	fi
	file1=$1
	file2=$2
	inode1=$(inode ${file1})
	inode2=$(inode ${file2})
	if [ "${inode1}" == "${inode2}" ]; then
		echo 0
	else
		echo 1
	fi
} # }}}
function camel2snake() { # {{{
	# https://genzouw.com/entry/2019/04/10/080016/1330/
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : camel2snake <word>"
		return 1
	fi
	inword=$1
	echo ${inword} | sed -r 's/^./\L\0/; s/([A-Z])/_\1/g; s/.*/\L\0/g;'
}
	function _test_camel2snake() { # {{{
		camel2snake aaaBbbCcc		# aaa_bbb_ccc
		camel2snake Aaa				# aaa
		camel2snake aaa				# aaa
		camel2snake AAA				# a_a_a
		camel2snake aaa_bbb_ccc		# aaa_bbb_ccc
	} # }}}
# }}}
function pascal2snake() { # {{{
	# https://genzouw.com/entry/2019/04/10/080016/1330/
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : pascal2snake <word>"
		return 1
	fi
	inword=$1
	camel2snake ${inword}
}
	function _test_pascal2snake() { # {{{
		pascal2snake aaaBbbCcc		# aaa_bbb_ccc
		pascal2snake Aaa			# aaa
		pascal2snake aaa			# aaa
		pascal2snake AAA			# a_a_a
		pascal2snake aaa_bbb_ccc	# aaa_bbb_ccc
	} # }}}
# }}}
function snake2camel() { # {{{
	# https://genzouw.com/entry/2019/04/10/080016/1330/
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : snake2camel <word>"
		return 1
	fi
	inword=$1
	echo ${inword} | sed -r 's/.*/\L\0/g; s/_([a-z0-9])/\U\1/g;'
}
	function _test_snake2camel() { # {{{
		snake2camel aaa_bbb_ccc		# aaaBbbCcc
		snake2camel _aaa			# Aaa
		snake2camel aaa_			# aaa_
		snake2camel AAA_BBB_CCC		# aaaBbbCcc
	} # }}}
# }}}
function snake2pascal() { # {{{
	# https://genzouw.com/entry/2019/04/10/080016/1330/
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : snake2pascal <word>"
		return 1
	fi
	inword=$1
	headchar=${inword:0:1}
	#echo ${headchar}
	if [ "${headchar}" == "_" ]; then
		snake2camel ${inword}
	else
		snake2camel _${inword}
	fi
}
	function _test_snake2pascal() { # {{{
		snake2pascal aaa_bbb_ccc	# AaaBbbCcc
		snake2pascal _aaa			# Aaa
		snake2pascal aaa_			# Aaa_
		snake2pascal AAA_BBB_CCC	# AaaBbbCcc
	} # }}}
# }}}
function camel2pascal() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : camel2pascal <word>"
		return 1
	fi
	inword=$1
	echo ${inword^}
}
	function _test_camel2pascal() { # {{{
		camel2pascal aaaBbbCcc	# AaaBbbCcc
		camel2pascal AaaBbbCcc	# AaaBbbCcc
		camel2pascal aaa		# Aaa
		camel2pascal Aaa		# Aaa
	} # }}}
# }}}
function pascal2camel() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : pascal2camel <word>"
		return 1
	fi
	inword=$1
	echo ${inword,}
}
	function _test_pascal2camel() { # {{{
		pascal2camel AaaBbbCcc	# aaaBbbCcc
		pascal2camel aaaBbbCcc	# aaaBbbCcc
		pascal2camel Aaa		# aaa
		pascal2camel aaa		# aaa
	} # }}}
# }}}
function searchpath2top() { # {{{
	if [ $# -ne 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : searchpath2top <start_dir_path> <key_file_name>"
		return 1
	fi
	start_dir_path=$1
	key_file_name=$2
	
	target_dir_path=${start_dir_path}
	top_dir_path=""
	while true
	do
		file_path=${target_dir_path}/${key_file_name}
		if [ -f ${file_path} ]; then
			top_dir_path=${target_dir_path}
			break
		fi
		if [ -z ${target_dir_path} ]; then
			break
		fi
		target_dir_path=${target_dir_path%/*}
	done
	echo ${top_dir_path}
} # }}}
function grepjapanese() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : grepjapanese <target_path>"
		return 1
	fi
	target_path=$1
	grep -nrP "[\x{3040}-\x{30FF}\x{4E00}-\x{9FFF}]" ${target_path}
} # }}}

### Git
alias gitlo="\
	git log \
	--all \
	--graph \
	--date=format:'%Y-%m-%d %H:%M' \
	--date-order \
	--decorate=full \
	--pretty=format:\" ::: %C(red)%ad%Creset ::: %C(blue)%h%Creset ::: %C(magenta)%d%Creset ::: %C(green)%an%Creset ::: %C(yellow)%s\""
#	--date=short \
alias gitstat="git status --ignored"
alias gitco="git checkout"
alias githeadmsg="git log --all --pretty=format:"%s" | head -1"
function gitdifftooldir() { # {{{
	if [ $# -eq 1 ]; then
		if [ "$1" == "-c" ]; then
			diffopt="--cached"
		else
			diffopt="$1"
		#	echo "[error] unsupported arguments: $1"
		#	return 1
		fi
	else
		diffopt=""
	fi
	_true=0
	_false=1
	
	echo "### git status ###"
	git status -s
	echo ""
	
	echo "### vimdiff modified files ###"
	if [ "${diffopt}" == "--cached" ]; then
		filelist=$(git status -s | grep "^M  " | sed "s/^M  //g")
	else
		filelist=$(git status -s | grep "^ M " | sed "s/^ M //g")
	fi
	for file in $filelist
	do
		is_binary=$(file --mime ${file} | grep "charset=binary" &> /dev/null; echo $?)
		#echo ${is_binary}
		if [ ${is_binary} -eq ${_false} ]; then
			echo "==> git difftool ${file} <=="
			git difftool ${diffopt} ${file}
			sleep 1
		fi
	done
} # }}}
function gitaddcmd() { # {{{
	echo "### git status -s"
	git status -s
	echo ""
	
	echo "### echo git add list"
	add_file_list=$(git status -s | grep "^.M " | cut -c 4- | cut -d" " -f 1)
	for add_file in ${add_file_list}
	do
		#echo ${add_file}
		echo "git add ${add_file}"
	done
} # }}}
function gitcommit() { # {{{
	commit_msg=$(githeadmsg)
	git commit -m "${commit_msg}"
} # }}}

### Tmux
alias tmrunsplit='tmux new-session \; source-file ~/.tmux.runsplit.conf'
alias tgr='vim ~/.tigrc'
alias tmc='vim ~/.tmux.conf'
alias tmcm='vim ~/.tmux.conf.mac.conf'
alias tml='tmux list-sessions'
alias tmb="export TMUX="

function tma() { # {{{
	# TMux Attach
	if [ ! -z "$TMUX" ]; then
		echo "[error] cannot be run on tmux."
		return 1
	fi
	config_path=~/.tmux.conf
	tmux source-file ${config_path}
	if [ $# -eq 1 ]; then
		session_name=${1}
		tmux attach-session -t ${session_name} || tmux new-session -s ${session_name}
	else
		tmux attach-session || tmux new-session
	fi
}
	complete -F _complete_tma tma # {{{
	function _complete_tma() { local cur; _get_comp_words_by_ref -n : cur; COMPREPLY=( $(compgen -W "${cmpllist_tma}" -- "${cur}") ); } # }}}
# }}}
function tmk() { # {{{
	# TMux Kill
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : tmk <session_name>"
		return 1
	fi
	session_name=${1}
	tmux kill-session -t ${session_name}
}
	complete -F _complete_tmk tmk # {{{
	function _complete_tmk() { local cur; _get_comp_words_by_ref -n : cur; COMPREPLY=( $(compgen -W "${cmpllist_tmk}" -- "${cur}") ); } # }}}
# }}}
function tmr() { # {{{
	# TMux Restart
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : tmr <session_name>"
		return 1
	fi
	session_name=${1}
	tmk ${session_name}
	tma ${session_name}
}
	complete -F _complete_tmr tmr # {{{
	function _complete_tmr() { local cur; _get_comp_words_by_ref -n : cur; COMPREPLY=( $(compgen -W "${cmpllist_tmr}" -- "${cur}") ); } # }}}
# }}}
function clear_session_name_to_cmplist() { # {{{
	cmpllist_tma=""
	cmpllist_tmk=""
	cmpllist_tmr=""
} # }}}
function add_session_name_to_cmplist() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : add_session_name_to_cmplist <session_name>"
		return 1
	fi
	session_name=$1
	cmpllist_tma="${cmpllist_tma} ${session_name}"
	cmpllist_tmk="${cmpllist_tmk} ${session_name}"
	cmpllist_tmr="${cmpllist_tmr} ${session_name}"
} # }}}
function add_session_list_to_cmplist() { # {{{
	session_list=$(tmux list-sessions | cut -d: -f 1)
	for session_name in "${session_list}"
	do
		add_session_name_to_cmplist "${session_name}"
	done
} # }}}
function _tmuxexecall() { # {{{
	if [ -f /.dockerenv ]; then
		echo "[error] can not be on docker container."
		return 1
	fi
	if [ -z "$TMUX" ]; then
		echo "[error] can only be run on tmux."
		return 1
	fi
	command_idx=0
	while (( $# > 0 ))
	do
		command_array[${command_idx}]="$1"
		command_idx=`expr ${command_idx} + 1`
		shift
	done
	
	winnum=$(tmux list-windows | tail -n 1 | cut -d: -f 1)
	activewinidx=$(tmux list-windows | grep "(active)" | cut -d: -f 1)
	winidx=1
	while [ ${winidx} -le ${winnum} ]
	do
		tmux select-window -t:${winidx}
		tmux set-window-option synchronize-panes on
		tmux send-keys ":qa!" C-m	# quit vim
		for cmd in "${command_array[@]}"
		do
			#echo "${cmd}"
			tmux send-keys "${cmd}" C-m
		done
		tmux set-window-option synchronize-panes off
		winidx=`expr ${winidx} + 1`
	done
	tmux select-window -t:${activewinidx}
} # }}}
function tmuxexecall_bre() { # {{{
	_tmuxexecall \
		"bre" \
		""
} # }}}
function _tmuxexec() { # {{{
	if [ $# -lt 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : _tmuxexec [<arguments>...]"
		return 1
	fi
	if [ -f /.dockerenv ]; then
		echo "[error] can not be on docker container."
		return 1
	fi
	if [ -z "$TMUX" ]; then
		echo "[error] can only be run on tmux."
		return 1
	fi
	command_idx=0
	while (( $# > 0 ))
	do
		command_array[${command_idx}]="$1"
		command_idx=`expr ${command_idx} + 1`
		shift
	done
	tmux set-window-option synchronize-panes on
	tmux send-keys ":qa!" C-m	# quit vim
	for cmd in "${command_array[@]}"
	do
		#echo "${cmd}"
		tmux send-keys "${cmd}" C-m
	done
	tmux set-window-option synchronize-panes off
} # }}}
function tmuxexec_bashtest() { # {{{
	_tmuxexec \
		"${HOME}/_repo/pj1tool-ros2dev/docker/bash_test.sh" \
		"cd workspace" \
		"lsetup" \
		""
} # }}}
function tmuxexec_bashtest2() { # {{{
	_tmuxexec \
		"cd ${HOME}/_repo/pj1tool-ros2dev/docker" \
		"./bash_test2.sh" \
		"cd workspace" \
		"lsetup" \
		""
} # }}}
if [ ! -f /.dockerenv ]; then
	clear_session_name_to_cmplist
	add_session_list_to_cmplist
	add_session_name_to_cmplist temp
fi
#export PROMPT_COMMAND="add_session_list_to_cmplist; ${PROMPT_COMMAND}"
#setenv PROMPT_COMMAND add_session_list_to_cmplist

### Docker
if [ -f /.dockerenv ]; then
	export TERM=screen-256color
fi

### WSL
unixname=$(uname -r)
unixname=$(tr '[:upper:]' '[:lower:]' <<< $unixname)
if [[ "${unixname}" == *microsoft* ]]; then
	alias cdw='cd /mnt/c/'
fi

### ROS2
alias gsetup="source /opt/ros/humble/setup.bash"
alias lsetup="source install/setup.bash"
#export RCUTILS_CONSOLE_OUTPUT_FORMAT="[{severity}] [{time}] [{name}]: {message} ({function_name}() at {file_name}:{line_number})"
export RCUTILS_CONSOLE_OUTPUT_FORMAT="[{severity}] [{time}] [{name}]: {message}"
#export ROS_LOCALHOST_ONLY=1
#export RCUTILS_LOGGING_USE_STDOUT=1		# The output from all debug levels goes to stderr by default. If 1, It is possible to force all output to go to stdout.
#export RCUTILS_COLORIZED_OUTPUT=1			# By default, the output is colorized when it's targeting a terminal. If 0 force disabling colorized. If 1 force enabling colorized.
export RCUTILS_LOGGING_BUFFERED_STREAM=1	# By default, all logging output is unbuffered. If 1, force buffer.
alias plotjuggler="ros2 run plotjuggler plotjuggler &> /dev/null &"
function killros2() { # {{{
	killprocessall ros2 &> /dev/null
	killprocessall "--ros-args" &> /dev/null
	killprocessall "/opt/ros/humble" &> /dev/null
	killprocessall "ros2cli.daemon" &> /dev/null
	killprocessall ign &> /dev/null
	ps a -u ${USER} | grep -v " 0:00 bash" | grep -v " 0:00 ps a -u "
} # }}}
function cbuild() { # {{{
	# colcon build --continue-on-error --executor sequential --symlink-install --packages-select <pkg_name>
	if [ $# -eq 0 ]; then
		pkg_sel_opt=""
	elif [ $# -eq 1 ]; then
		pkg_sel_opt="--packages-select ${1}"
	else
		echo "[error] wrong number of arguments."
		echo "  usage : cbuild [<package_name>]"
		return 1
	fi
	gsetup
	#lsetup || return 1
	colcon build \
		--continue-on-error \
		--executor sequential \
		--symlink-install \
		${pkg_sel_opt}
} # }}}
function cbuildc() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : cbuildc <package_name>"
		return 1
	fi
	pkg=${1}
	rm -rf install/${pkg} build/${pkg}
	cbuild ${pkg}
} # }}}
function ctest() { # {{{
	# colcon test --packages-select <pkg_name>
	if [ $# -eq 0 ]; then
		pkg_sel_opt=""
	elif [ $# -eq 1 ]; then
		pkg_sel_opt="--packages-select ${1}"
	else
		echo "[error] wrong number of arguments."
		echo "  usage : ctest [<package_name>]"
		return 1
	fi
	gsetup
	#lsetup || return 1
	# colcon test ${pkg_sel_opt} && colcon test-result --verbose
	colcon test ${pkg_sel_opt} && colcon test-result
} # }}}
function renamerospkg() { # {{{
	if [ $# -ne 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : renamerospkg <source> <destination>"
		echo "    <source>/<destination> must be small snake case. (e.g. aaa_bbb_ccc)"
		return 1
	fi
	
	src_snake_small=${1}
	dst_snake_small=${2}
	src_snake_large=${src_snake_small^^}
	dst_snake_large=${dst_snake_small^^}
	src_nodelim_large=${src_snake_large//_/}
	dst_nodelim_large=${dst_snake_large//_/}
	src_pascal=$(snake2pascal ${src_snake_small})
	dst_pascal=$(snake2pascal ${dst_snake_small})
	
	echo "Rename ROS2 package names as follows:"
	echo "  - ${src_snake_small} -> ${dst_snake_small}"
	echo "  - ${src_snake_large} -> ${dst_snake_large}"
	echo "  - ${src_nodelim_large} -> ${dst_nodelim_large}"
	echo "  - ${src_pascal} -> ${dst_pascal}"
	read -p "Do you want to continue? [y/n] : " answer
	if [ ! ${answer} == "y" ]; then
		echo "Canceled, process will be aborted."
		return 1
	fi
	echo ""
	
	echo "#### Rename directorys (${src_snake_small} -> ${dst_snake_small}) ####"
	renamedirfiles "${src_snake_small}" "${dst_snake_small}"
	echo ""
	
	echo "#### Grep replace with small snake case (${src_snake_small} -> ${dst_snake_small}) ####"
	greprep "${src_snake_small}" "${dst_snake_small}"
	echo ""
	
	echo "#### Grep replace with large snake case (${src_snake_large} -> ${dst_snake_large}) ####"
	greprep "${src_snake_large}" "${dst_snake_large}"
	echo ""
	
	echo "#### Grep replace with large no delimiter case (${src_nodelim_large} -> ${dst_nodelim_large}) ####"
	greprep "${src_nodelim_large}" "${dst_nodelim_large}"
	echo ""
	
	echo "#### Grep replace with pascal case (${src_pascal} -> ${dst_pascal}) ####"
	greprep "${src_pascal}" "${dst_pascal}"
	echo ""
} # }}}
function alignsdfxml() { # {{{
	if [ $# -ne 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : alignsdfxml <filepath>"
		return 1
	fi
	infile=${1}
	if [ ! -f ${infile} ]; then
		echo "[error] ${infile} does not exist."
		return 1
	fi
	
	bakfile=${infile}.bak
	while [ -f ${bakfile} ]
	do
		bakfile=${bakfile}_
	done
	\cp -f ${infile} ${bakfile}
	
	sed -i 's/\\n/\n/g' ${infile}
	sed -i "s/\\\\'/'/g" ${infile}
	sed -i 's/\\\"/\"/g' ${infile}
	sed -i 's/^data: "//g' ${infile}
	sed -i 's/^"$//g' ${infile}
	sed -i 's/^\n//g' ${infile}
#	sed -i '1i <?xml version="1.0" ?>' ${infile}
} # }}}
function outputros2nodesinfo() { # {{{
	if [ $# -ge 1 ]; then
		suffix=_${1}
	else
		suffix=
	fi
	
	# output node list
	listlog=_list_nodes${suffix}.log
	rm -rf ${listlog}
	echo "$ ros2 node list " &>> ${listlog}
	ros2 node list &>> ${listlog}
	
	# output node infos
	nodelist=$(ros2 node list)
	infolog=_info_nodes${suffix}.log
	rm -rf ${infolog}
	for node in ${nodelist}
	do
		echo "$ ros2 node info ${node}" &>> ${infolog}
		ros2 node info ${node} &>> ${infolog}
	done
	
	# output component list
	complistlog=_list_comps${suffix}.log
	rm -rf ${complistlog}
	echo "$ ros2 component list" &>> ${complistlog}
	ros2 component list &>> ${complistlog}
} # }}}
function outputros2topicsinfo() { # {{{
	if [ $# -ge 1 ]; then
		suffix=_${1}
	else
		suffix=
	fi
	
	# output topic list
	listlog=_list_topics${suffix}.log
	rm -rf ${listlog}
	echo "$ ros2 topic list " &>> ${listlog}
	ros2 topic list &>> ${listlog}
	
	# output topic infos
	topiclist=$(ros2 topic list)
	infolog=_info_topics${suffix}.log
	rm -rf ${infolog}
	for topic in ${topiclist}
	do
		echo "$ ros2 topic info ${topic}" &>> ${infolog}
		ros2 topic info ${topic} &>> ${infolog}
	done
} # }}}
function outputros2paramlist() { # {{{
	if [ $# -ge 1 ]; then
		suffix=_${1}
	else
		suffix=
	fi
	
	# output topic list
	listlog=_list_params${suffix}.log
	rm -rf ${listlog}
	echo "$ ros2 param list " &>> ${listlog}
	ros2 param list &>> ${listlog}
} # }}}
function formatros2nodesinfo() { # {{{
	if [ $# -ne 2 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : formatnodesinfo <infile> <outfile>"
		return 1
	fi
	if [ -z ${DEVDIR} ]; then
		echo "[error] \${DEVDIR} must be set."
		return 1
	fi
	scriptpath=${DEVDIR}/_script/python/format_nodes_info.py
	infile=${1}
	outfile=${2}
	python3 ${scriptpath} ${infile} ${outfile}
} # }}}
function latestloglaunch() { # {{{
	if [ -z ${DEVDIR} ]; then
		echo "[error] \${DEVDIR} must be set."
		return 1
	fi
	logdirpath=${HOME}/.ros/log
	if [ ! -d "${logdirpath}" ]; then
		echo "[error] ${logdirpath} does not exist."
		return 1
	fi
	
	# extract log file paths with the latest date
	curyear=$(date "+%Y")
	idx=5
	logfiledir=$(find ${logdirpath} -maxdepth 1 -type d | sort | tail -1)
	logfile=${logfiledir}/launch.log
	
	# output log file date
	logdirname=${logfiledir##*/}
	#echo ${logdirname}
	logfile_date=${logdirname:0:10}
	#logfile_date=${logfile_date//-/\/}
	logfile_time=${logdirname:11:8}
	logfile_time=${logfile_time//-/:}
#	echo ${logfile_date} ${logfile_time}
	
	# output log file path
	set_clipboard "${logfile}"
	echo ${logfile}
} # }}}
function latestlognode() { # {{{
	if [ $# -lt 1 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : latestlognode <node_name>"
		return 1
	fi
	logdirpath=${HOME}/.ros/log
	if [ ! -d "${logdirpath}" ]; then
		echo "[error] ${logdirpath} does not exist."
		return 1
	fi
	nodename=${1}
	
	# check for the existence of log files on the target node
	grepresult=$(find ${logdirpath} -maxdepth 1 -type f | grep -E "${nodename}_[0-9]+_[0-9]+\.log" &> /dev/null; echo $?)
	#echo ${grepresult}
	if [ ${grepresult} -ne 0 ]; then
		echo "[error] log file of \"${nodename}\" does not exist."
		return 1
	fi
	
	# get unixtime string position(index) when split by "_"
	filename=$(find ${logdirpath} -maxdepth 1 -type f | grep -E "${nodename}_[0-9]+_[0-9]+\.log" | tail -1)
	filename_org=${filename}
	filename_new=${filename//_/}
	#echo ${#filename_org} ${#filename_new}
	idx=$(expr ${#filename_org} - ${#filename_new} + 1)
	#echo ${idx}
	
	# extract log file paths with the latest unixtime
	logpath=$(find ${logdirpath} -maxdepth 1 -type f | grep -E "${nodename}_[0-9]+_[0-9]+\.log" | sort -t"_" -k ${idx},${idx} -n | tail -1)
	
	# output log file date
	logfile=${logpath##*/}
	unixtime_tmp=$(echo ${logfile} | cut -d"_" -f ${idx})
	unixtime_tmp=${unixtime_tmp%%.*}
	unixtime=${unixtime_tmp:0:10}
#	convunixtimetodate ${unixtime}
	
	# output log file path
	set_clipboard "${logpath}"
	echo ${logpath}
} # }}}
function cpk() { # {{{
	pkg_file_name=package.xml
	pkg_root_dir=$(searchpath2top ${PWD} ${pkg_file_name})
	#echo ${pkg_root_dir}
	if [ -z ${pkg_root_dir} ]; then
		echo "[error] ${pkg_file_name} does not exist in the upper level directory."
	else
		pkg_file_path=${pkg_root_dir}/${pkg_file_name}
		grep "<name>.*</name>" ${pkg_file_path} | sed "s/.*<name>//g" | sed "s/<\/name>//g"
	fi
} # }}}
function e2q() { # {{{
	# euler to quaternion
	order="wxyz"
	delimiter=" "
	if [ $# -ne 3 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : e2q <roll> <pitch> <yaw>   (unit:[rad])"
		return 1
	fi
	roll=$1
	pitch=$2
	yaw=$3
	values_str=$(quaternion_from_euler ${roll} ${pitch} ${yaw} | grep -P "^ [WXYZ] " | grep -oP "[-\d.]+" | tr "\n" " ")
	values_array=($values_str)
	if [ ${order} == "wxyz" ]; then
		echo "w${delimiter}x${delimiter}y${delimiter}z"
		echo "${values_array[0]}${delimiter}${values_array[1]}${delimiter}${values_array[2]}${delimiter}${values_array[3]}"
	else
		echo "x${delimiter}y${delimiter}z${delimiter}w"
		echo "${values_array[1]}${delimiter}${values_array[2]}${delimiter}${values_array[3]}${delimiter}${values_array[0]}"
	fi
} # }}}
function q2e() { # {{{
	# quaternion to euler
	delimiter=" "
	if [ $# -ne 4 ]; then
		echo "[error] wrong number of arguments."
		echo "  usage : q2e <w> <x> <y> <z>"
		return 1
	fi
	w=$1
	x=$2
	y=$3
	z=$4
	values_str=$(quaternion_to_euler ${w} ${x} ${y} ${z} | grep -P "^ (roll|pitch|yaw) .* degrees" | grep -oP "[-\d.]+" | tr "\n" " ")
	values_array=($values_str)
	echo "roll${delimiter}pitch${delimiter}yaw [degrees]"
	echo "${values_array[0]}${delimiter}${values_array[1]}${delimiter}${values_array[2]}"
} # }}}

### Ignition Gazebo
function addenv_ignresource() { # {{{
	path=""
	if [ $# -ge 1 ]; then
		path="$1"
	else
		path="${PWD}"
	fi
	envname=IGN_GAZEBO_RESOURCE_PATH
	setenv "${envname}" "${path}"
	echo "${envname} = $(printenv ${envname})"
} # }}}

#########################################################
# Environment dependent settings
#########################################################
setenv PATH "${HOME}/_work/gz-usd/build/bin"			# for sdf2usd, usd2sdf
setenv PATH "${HOME}/_prg/USD/bin"						# for USD

