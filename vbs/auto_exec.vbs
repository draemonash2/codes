Option Explicit

' �{�X�N���v�g���������s�����邽�߂ɂ́A���O�Ɉȉ��̑�����s���Ă������ƁI
' 
' 1. �C�x���g�r���[�A�[���J���B
' 	�R���g���[���p�l�� => �Ǘ��c�[�� => �C�x���g�r���[�A�[
' 2. ��ʍ����u���O�v�y�C���ɂāA�ȉ���I���B
' 	�A�v���P�[�V�����ƃT�[�r�X���O => Microsoft => Windows => DriverFrameworks-UserMode => Operational
' 3. ��ʉE���u����v�y�C���ɂāu���O��L�����v���N���b�N�B

'�����L������
'   ����Avbs ����}�N�������s������ƁA�Ȃ����ȉ��̌��ۂ���������B
'   �E�}�N���������s�����
'   �E�}�N�����s���߂��Ⴍ����x��
'   ���̂��߁A����� Excel �t�@�C�����J���݂̂ɗ��߂Ă����B

Const MACRO_BOOK_PATH = "C:\Users\draem_000\Documents\Dropbox\100_Documents\143_�y�����z���ߐH�Z���H���^�g��\320_�y�g�́z�̏d�Ǘ�.xlsm"
Const MACRO_NAME = "���f�[�^����()"

Dim objExcel
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Open MACRO_BOOK_PATH
