@echo off
::	<<�T�v>>
::	  TARGET_DIR_PATH �z���i�T�u�t�H���_�܂ށj�ɂ���t�@�C����
::	  KEY_FILE_NAME ��T���A���ꂪ�i�[���ꂽ�t�H���_�p�X��
::	   �V�X�e�����ϐ��́uPath�v�ɒǉ�����
::	
::	<<�g����>>
::	  �P�DKEY_FILE_NAME �Ɏw�肵�����O�̃t�@�C�����A�V�X�e�����ϐ���
::		  �uPath�v�ɒǉ��������t�H���_���ɃR�s�[����
::	  �Q�D�{�o�b�`�t�@�C�����u�Ǘ��҂Ƃ��Ď��s�v����

:: ### �ݒ��� ###
set ADD_TO_PATH_SCRIPT_PATH=C:\codes\vbs\700_AddToPathOfEnvVariable.vbs
set TARGET_DIR_PATH=C:\prg_exe\
set KEY_FILE_NAME=_add_to_sys_env_directory

:: ### ���� ###
FOR /R "%TARGET_DIR_PATH%" %%i IN (%KEY_FILE_NAME%) DO (
	if exist %%i (
		echo %%~dpi
		call %ADD_TO_PATH_SCRIPT_PATH% %%~dpi
	)
)
pause
