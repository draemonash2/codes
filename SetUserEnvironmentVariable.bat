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
	setx MYPATH_CODES	"C:\codes"
	echo In order to reflect this setting, please restart the windows!
) else if %ANS% == n (
	reg delete HKCU\Environment /v MYPATH_CODES
	echo In order to reflect this setting, please restart the windows!
) else (
	echo [error] illegal answer!!!
)

pause
