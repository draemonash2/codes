@echo off
::	<<�T�v>>
::	  TARGET_DIR_PATH �z���i�T�u�t�H���_�܂ށj�ɂ���t�@�C��
::	  �u_add_to_path.bat�v��T���A���ꂪ�i�[���ꂽ�t�H���_�p�X��
::	   �V�X�e�����ϐ��́uPath�v�ɒǉ�����
::	
::	<<�g����>>
::	  �P�D�ȉ��̃t�@�C�����쐬
::			�t�@�C�����F_add_to_path.bat
::			�t�@�C���̒��g�Fcall %1 %~dp0
::	  �Q�D�P�ō쐬�����t�@�C���� Path �ɒǉ��������t�H���_���ɃR�s�[����
::	  �R�D�{�o�b�`�t�@�C�����u�Ǘ��҂Ƃ��Ď��s�v����

:: ### �ݒ��� ###
set ADD_TO_PATH_SCRIPT_PATH=C:\codes\vbs\700_AddToPathOfEnvVariable.vbs
set TARGET_DIR_PATH=C:\prg_exe\

:: ### ���� ###
FOR /R "%TARGET_DIR_PATH%" %%i IN (_add_to_path.bat) DO (
	if exist %%i (
		echo %%i
		call %%i %ADD_TO_PATH_SCRIPT_PATH%
	)
)
pause
