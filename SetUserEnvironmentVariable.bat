:: *******************************************************************
:: * �ړI�F�X�N���v�g���J�����g�f�B���N�g�����擾����ہA���[�J������
:: *       �ɂ�������s�ƊǗ��Ҍ����ɂ�������s�Ƃł͕ԋp�l���قȂ�
:: *       �ꍇ������B���̏ꍇ�A���O�Ƀ��[�U�[���ϐ��ɐݒ肵�Ă���
:: *       ���ƂŁA�������̎��s���ʂ����킹�邱�Ƃ��ł���B
:: *       
:: *       �{�o�b�`�t�@�C���ł́A���[�U�[���ϐ��̐ݒ�ƍ폜��������
:: *       ���邱�Ƃ�ړI�Ƃ���B
:: *******************************************************************
@echo off

echo add or delete user environment variable?
set /p ANS="  add=>y, delete=>n : "
if %ANS% == y (
	setx MYPATH_CODE_BAT	"C:\codes\bat"
	setx MYPATH_CODE_C		"C:\codes\c"
	setx MYPATH_CODE_HTTP	"C:\codes\http"
	setx MYPATH_CODE_JAVA	"C:\codes\java"
	setx MYPATH_CODE_PYTHON	"C:\codes\python"
	setx MYPATH_CODE_RUBY	"C:\codes\ruby"
	setx MYPATH_CODE_SH		"C:\codes\sh"
	setx MYPATH_CODE_VBA	"C:\codes\vba"
	setx MYPATH_CODE_VBS	"C:\codes\vbs"
	setx MYPATH_CODE_VDM	"C:\codes\vdm++"
	echo In order to reflect this setting, please restart the windows!
) else if %ANS% == n (
	reg delete HKCU\Environment /v MYPATH_CODE_BAT
	reg delete HKCU\Environment /v MYPATH_CODE_C
	reg delete HKCU\Environment /v MYPATH_CODE_HTTP
	reg delete HKCU\Environment /v MYPATH_CODE_JAVA
	reg delete HKCU\Environment /v MYPATH_CODE_PYTHON
	reg delete HKCU\Environment /v MYPATH_CODE_RUBY
	reg delete HKCU\Environment /v MYPATH_CODE_SH
	reg delete HKCU\Environment /v MYPATH_CODE_VBA
	reg delete HKCU\Environment /v MYPATH_CODE_VBS
	reg delete HKCU\Environment /v MYPATH_CODE_VDM
	echo In order to reflect this setting, please restart the windows!
) else (
	echo [error] illegal answer!!!
)

pause
