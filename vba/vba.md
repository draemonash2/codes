# �A�h�C��
- [VbePlus](http://www.vector.co.jp/soft/dl/win95/prog/se176543.html)
	- �C���X�g�[����AVBE�̃c�[���o�[���A�h�C�����A�h�C���}�l�[�W��
	- VbePlus �̃��[�h���@���u�N����/���[�h�v�ɕύX
- [MZTools](http://zabi.0am.jp/?p=2676)
	- �C���X�g�[�����Ă��\������Ȃ��ꍇ�AHKEY\_CURRENT\_USER �ɕK�v�ȃ��W�X�g���L�[���o�^����Ă��Ȃ��\��������B
	- Administrator �� HKEY\_CURRENT\_USER ����R�s�[���邱�ƁB

# Tips
- ���W���[�����Ɗ֐����𓯂��ɂ��邱�Ƃ͂ł��Ȃ��B
- �񎟌��z��̍Ē�`�ɂ���
	- �ꎟ���ڂ̗v�f���͕ύX�ł��Ȃ��B�񎟌��ڂ̃T�C�Y�̂݊g���ł���B
	- �񎟌��ڂ̗v�f����ύX�����ꍇ�A�ꎟ���ڂɂ܂ŉe������B&br()�iex. (0, 1)�A(1, 1) �̗v�f�������z��ɑ΂��� (1, 1) �� (1, 3) �Ɨv�f����ύX�����ꍇ�A(0, 3)�A(1, 3) �̔z��ƂȂ�j
- �G���[�u�萔�����K�v�ł��v�ɂ���
	- �f�o�b�O��r���Œ�~����ƁA�ݒ肵�Ă���͂��̒l������`�����ƂȂ�R���p�C���G���[���������邱�Ƃ�����B
	- �΍�� ENUM ��`����ҏW���čēx�߂��B
- Sheets �� Worksheets �̈Ⴂ
	- Excel �̃V�[�g�ɂ͕����̎�ނ�����B�iex.�O���t�V�[�g�A���W���[���V�[�g�A���[�N�V�[�g�c�j
	- Sheets �́u��L�̂��ׂĂ��܂����v�I�u�W�F�N�g
	- Worksheets �́u���[�N�V�[�g�̂݁v�̃I�u�W�F�N�g
- Integer��Long�ǂ��炪�ǂ��H
	- �^�̈Ⴂ�ɂ�鑬�x��r
		- �T�疜��l�̑�����s�������ʁAInteger�^�F515ms�ALong�^�F465ms ���������B�i10%��Long�^�̕��������j
			- ���FWindows 7 64bit
			- �I�t�B�X�FOffice 2013 32bit
		- ���R�͓������ 32bit Office �̂��߁A 4Byte �^�ϐ��̕����A�N�Z�X���������߂Ǝv����B
	- �^�̈Ⴂ�ɂ��T�C�Y��r
		- Integer�^�ϐ��i2Byte�j�Ȃ�PGB�������܂łT���ϐ��錾�ł���
		- Long�^�ϐ��i4Byte�j�Ȃ�PGB�������܂�2.5���ϐ��錾�ł���i�e�ʂ̈����͋C�ɂȂ�Ȃ��j
	- ��L�𓥂܂���� Long �^���g�p����̂��ǂ��I�i�A�N�Z�X���������A�e�ʂ��C�ɂ���قǑ傫���Ȃ����߁j
- Const�z��̍���
	- Const�z��� VBA �ł͒�`�ł��Ȃ����AConst ������� Split ���邱�ƂŎ����\�B�i�����������Ă��܂����B�B�j
	- Const C\_AAA = "Hello!,World!" : asArray = Split(C\_AAA, ",")
- CreateObject �֐��ɂ���
	- �쐬�����I�u�W�F�N�g�́A�g�p��ɕK�� Nothing ��ݒ肷�邱�ƁI
	- �I�u�W�F�N�g���������Ȃ��Ȃ胁�����Ɏc�葱���邱�ƂɂȂ�I 
- �A�������󔒂���������� Split �֐��ŕ��������ꍇ�A�󕶎���̗v�f���ł��Ă��܂����
	- ���O�ɘA�������󔒂���ɂ܂Ƃ߂Ă��� Split ����
		```vba
		'�A�������󔒂���ɂ܂Ƃ߂�
		Do While InStr(sTrgtLine, "  ")
			sTrgtLine = Replace$(sTrgtLine, "  ", " ")
		Loop
		sTrgtLine = Trim$(sTrgtLine)
		asSplitLine = Split(sTrgtLine, " ")
		```
- �N���X���W���[���̃����b�g�E�f�����b�g
	- �f�����b�g
		- �C���X�^���X�������I�u�W�F�N�g�𑼃N���X�ŎQ�Ƃł��Ȃ��B
			- �˃I�u�W�F�N�g�𐶐������N���X���O���[�o���ɂ���Ηǂ�
		- Property ���g�p����ꍇ�A�R�[�h�ʂ�����
		- �\���̂��g���Ȃ�
			- �N���X�͓���q�ɂ��邱�Ƃ��ł���B
				```vba
				'Class���W���[�� ClassString
				Public Function GetDup( _
					ByVal sVal As String
				) As String
					GetDup = sVal & sVal
				End Function
				 
				'Class���W���[�� ClassCom
				Public STR As New ClassString
				 
				'�W�����W���[�� Module1
				Public C As New ClassCom
				Sub Test02()
					Debug.Print C.STR.GetDup("�͂�")
				End Sub
				```
		- �񋓑̂��g���Ȃ�
	- �����b�g
		- �O���[�o���ϐ����B���ł���B
			- �O���[�o���ϐ���ύX����ꍇ�̉e���͈͂��N���X���ɐ����ł���
		- �ʃN���X�Ȃ瓯���֐������g����B
			- �C���^�[�t�F�[�X
				```vba
				Private Sub Class_Initialize()
				End Sub
				 
				Public Function Execute() As String
				End Function
				```
			- �N���X�P
				```vba
				Implements ClassCmn
				 
				Dim a As String
				 
				Private Sub Class_Initialize()
					a = "A"
				End Sub
				 
				Public Function ClassCmn_Execute() As String
					ClassCmn_Execute = a
				End Function
				```
			- �N���X�Q
				```vba
				Implements ClassCmn
				 
				Dim b As String
				 
				Private Sub Class_Initialize()
					b = "B"
				End Sub
				 
				Public Function ClassCmn_Execute() As String
					ClassCmn_Execute = b
				End Function
				```
			- �g�p�ӏ�
				```vba
				Sub test()
					Dim clChkExecTable(1) As ClassCmn
				 
					Set clChkExecTable(0) = New Class1
					Set clChkExecTable(1) = New Class2
				 
					Debug.Print clChkExecTable(0).Execute
					Debug.Print clChkExecTable(1).Execute
				 
				End Sub
				```
		- �����������Ăяo���s�v�B
- ���������@�܂Ƃ�
	- csv �t�@�C����荞�ݕ��@
	- �ꗗ�ւ̃A�N�Z�X�� Dictionary ���g���B
	- �Z���ւ̃A�N�Z�X�� variant �^�ϐ��֑�����Ă��珑���߂��B
	- �O���t�����͋ɗ̓I�v�V���������炷�B
	- ���K�\�������͋ɗ͎g��Ȃ��BInStr �ő�p�i�擪�ɂ��邩�ǂ����� InStr() = 0 �Ŕ��肷��j
	- �󔒖��߂͈��Z���ɑ�������A�󔒃Z���I������̈ꊇ�l����B
	- With ���g���B
	- Integer �ł͂Ȃ� Long ���g���B
- �z��ECollection�EDictionary �̑��x
	- �ڍׂȑ��茋�ʂ͓Y�t�t�@�C���uArray�ECollection�EDictionary���x��r.xlsm�v�Q��
		- �z��
			- �S�̓I�ɔ����I�C���f�b�N�X�A�N�Z�X�����Ȃ�z����g�����ƁI
		- Dictionary
			- �C���f�b�N�X�A�N�Z�X�͒x������̂ŁA��΃C���f�b�N�X�A�N�Z�X����ȁI�K�� For Each �ŃA�N�Z�X���邱�ƁI
		- Collection
			- �傫�ȗv�f�ԍ��ɃA�N�Z�X����ۂ͒x���Ȃ�I�̂ŁA�v�f�����\�z�ł��Ȃ��ꍇ�͊�{�g��Ȃ��B�������AKey �A�N�Z�X�̏ꍇ�͖��Ȃ��I
- CPU �g�p���}������@
	- ���������̊Ԃ� sleep 1 ��ǉ�����B
	- �����͒����Ȃ邪�A�o�b�N�O���E���h�ł̏������͂��ǂ�B
- �t�@�C���I�[�v���u�o�C�i�����[�h�v�ɂ���
	- ���s�R�[�h�Ȃǂ���؂蕶���Ƃ��Ĉ��킸�A�P�o�C�g�P�ʂŃt�@�C���̐擪���璀���ǂݏ�������B
	- �o�C�i���t�@�C���̓ǂݏ���������ۂɎg�p����B
- Like ���Z�q�ɂ���
	- ���K�\���́u�悤�Ɂv��r�ł���B
	- �g�p��FIf sAddress Like "[!����,���l,��t]\*" Then �`
		- �u�����A���l�A��t�ł͂Ȃ��Z���v�̏ꍇ�`
			
| �L��         | �Ӗ�                                        | �g�p��      | �}�b�`���镶����          |
|:---|:---|:---|:---|
| ?            | �C�ӂ�1����                                 | ����?       | ���Ȃ��A���ȂׁA���Ȃ�(�Ȃ�)       |
| \*           | 0�ȏ�̔C�ӂ̕���                         | ����\*      | �������A�����Ȃ��A������Ȃ�(�Ȃ�) |
| #            | 1�����̐��l(0�`9)                           | ##          | 01�A26�A95(�Ȃ�) |
| [charlist]   | charlist�Ɏw�肵�������̒���1����           | [A-F]       | A�AB�AC�AD�AE�AF |
| [!charlist]  | charlist�Ɏw�肵�������̒��Ɋ܂܂�Ȃ�1���� | [!A-F]      | G �AH�AI(�Ȃ�) |

- On Error �̓���q�ɂ���
	- On Error �͊֐����ׂ�������q�����Ă����Ȃ��B
	- �������A�֐����̓���q�͓K�p����Ȃ��B
		```vb
		Sub main()
			On Error Resume Next
			On Error Resume Next
			Call subfunction
			Debug.Print 5 / 0 '�ˎ��s���G���[����
			On Error GoTo 0
		   'Debug.Print 5 / 0 '�ˎ��s���G���[���������ăv���O��������~����B
			On Error GoTo 0
		End Sub
		Sub subfunction()
			On Error Resume Next
			Debug.Print 5 / 0 '�ˎ��s���G���[����
			On Error GoTo 0
			Debug.Print 5 / 0 '�ˎ��s���G���[����
		End Sub
		```
- �傫���̒P�ʁu�|�C���g�v�Ɓu�s�N�Z���v�ɂ���
	- �|�C���g�FMicrosoft����`����P�ʁi�����炭�j�B8.38 �|�C���g��MS�S�V�b�N11�|�C���g(���p) ��8�����Ə����\���ł��镝�B
	- �s�N�Z���F�f�B�X�v���C�̕\����v�����^�̏o�͂��\������ŏ��P�ʁi�����ȓ_�j

# VBE �ݒ�
- Excel ���{���Ɂu�J���v��ǉ�
	- �c�[�� �� �I�v�V���� �� ���{�����[�U�[�ݒ�ɂ� [\*] �J��
- �ϐ��錾������ (�ϐ��������ł����s���܂ŃG���[���������Ȃ����ߕύX)
	- �c�[�� �� �I�v�V���� �� �ҏW �� [\*] �ϐ��̐錾����������
- �����\���`�F�b�N�𖳎� (���s�̓x�Ɍx���E�B���h�E���������邽��)
	- �c�[�� �� �I�v�V���� �� �ҏW �� [\_] �����\���`�F�b�N
- �G�f�B�^�����E�w�i�F�̕ύX
	- �c�[�� �� �I�v�V���� �� �G�f�B�^ �� �R�[�h�̕\���F ��ǂ��ȂɁc
- �����s�R�����g�A�E�g�{�^���ݒu
	- �\�� �� �c�[���o�[ �� [\*] �ҏW �� �I�v�V����

# VBE �V���[�g�J�b�g�L�[

| ���� | �L�[�z�u |
|:---|:---|
| IntelliSense ���s | Ctrl+Space |
| IntelliSense �I�� | Tab        |
| �}�N�������I��    | Ctrl+Break |
| ��`�ֈړ��i���^�O�W�����v�j  | Shift+F2 |
| ���̏ꏊ�ֈړ��i���^�O�W�����v�j  | Ctrl+Shift+F2 |

# �\��
## �u�`�v�͉��s�������B
- �y�ϐ�������`�zOption Explicit
- �y�ϐ�/�z���`�zDim aVal(5) As Integer '�v�f���͂O�I���W���B���̗�ł͗v�f���U�̔z�񂪍쐬�����
- �y�萔��`�zConst NUM As Integer = 1
- �y�\���̒�`�zType T\_XXX �` iVal1 As Integer �` iVal2 As Integer �` End Type
- �y�֐���`�zPrivate Function FuncA ( sVal1 As String, sVal2 As Integer ) �` End Function
- �y�֐��ďo�zCall Func()
- �y�֐��G���[�l�ԋp�zDim vRetVal As Variant �` vRetVal = CVErr(xlErrRef) '#VALUE��ԋp
- �y�}�N����`�zPublic Sub SubA ( sVal1 As String, sVal2 As Integer ) �` End Sub
- �y�񋓌^��`�zEnum E\_XXX �` NUM1 �` NUM2 �` End Enum
- �y�u���b�N�E�o�iSub/Function/For/Do�j�zExit (Sub|Function|For|Do)
- �y�A���R�}���h���s�zDim sStr As String : sStr = "abc"
- �y�ꎞ��~�zStop
- �y�N���X�C���X�^���X�����zDim cPrfrmMes As New PerformanceMeasurement
- �y�N���X�C���X�^���X�j���zSet cPrfrmMes = Nothing
- �y�v���O�����I���zEnd
- �y�֐��G���[�l�ԋp�zDim vRetVal As Variant �` vRetVal = CVErr(xlErrRef) '#VALUE��ԋp�i�G���[�l�̏ڍׂ� [������](https://msdn.microsoft.com/ja-jp/library/office/ff839168.aspx) �Q�Ɓj

- �yif�zIf iVal = 1 Or iVal = 2 Then �` ElseIf iVal = 3 Then �` Else �` End If
- �yif�i��I�u�W�F�N�g�m�F�j�zIf objTest Is Nothing Then �` Else �` End If
- �yswitch�zSelect Case iVal �` Case 1 �` Case Else �` End Select
- �yfor�zFor iVal1 = 1 To 3 [Step 1] �` Next Val
- �yfor each�zFor Each Value in Values �`�����` Next
- �ywhile�zDo �`(���������^)�` Loop While ������
- �ydo while�zDo While ������ �`(���������^)�` Loop
- �ydo until�zDo Until ������ �`(���������U)�` Loop
- �ywith�zWith �I�u�W�F�N�g�� �` End With
- �y�R�����g�z'�R�����g
- �y���́zsStr = InputBox( "�e�L�X�g����͂��Ă�������", "title", "default value" )
- �y�o�͂P�zMsgBox "Hello world", vbExclamation '���������͎��Ɍ�₪�\�������
- �y�o�͂Q�zDebug.Print "Hello world"
- �y�`�F�b�N�����zDebug.Assert ������ 'False�̏ꍇ�A�������ꎞ��~
- �y�m�F�����zDim vAnswer As Variant �` vAnswer = MsgBox("�������p�����܂����H", vbOKCancel, "�^�C�g��") '�������͕\���{�^���̎�ނ��w��B�{�^���̎�ނ� [������](http://www.kanaya440.com/contents/script/vbs/function/others/msgbox.html) �Q�ƁB

- �y���K�\���z�T���v���R�[�h�Q��

- �y�R���N�V���� ��`�zDim cTrgtPaths As Variant �` Set cTrgtPaths = CreateObject("System.Collections.ArrayList")
	- �y�R���N�V���� �ǉ��zcTrgtPaths.Add "c:\test\a.txt"
	- �y�R���N�V���� �l���o���i�P��j�zcTrgtPaths.Item(0) 'c:\test\a.txt�i�O�I���W���j
	- �y�R���N�V���� �l���o���i���[�v�j�zDim vTrgtPath As Variant �` For Each vTrgtPath In cTrgtPaths �` MsgBox vTrgtPath �` Next
	- �y�R���N�V���� �v�f���擾�zcTrgtPaths.Count '�v�f���i�����̗v�f�ԍ��ł͂Ȃ��j
	- �y�R���N�V���� �폜�zcTrgtPaths.Remove "c:\test\b.xlsx" '�v�f�̒l���w�肷��B�v�f�ԍ��ł͍폜�ł��Ȃ��B
	- �y�R���N�V���� �}���zcTrgtPaths.Insert 2, "c:\test\e.ppt" '�v�f�ԍ�2�֑}�������i���v�f�ԍ�2�ȍ~����v�f�����j
	- �y�R���N�V���� �\�[�g�zcTrgtPaths.Sort
	- �y�R���N�V���� �z��ϊ��zDim avTrgtPaths As Variant �` avTrgtPaths = cTrgtPaths.ToArray() 'Variant�^�z��ɕϊ�
	- �y�R���N�V���� �S�v�f�폜�zcTrgtPaths.Clear

- �y�A�z�z�� ��`�zDim oPriceOfFruit As Object �` Set oPriceOfFruit = CreateObject("Scripting.Dictionary")&br()�iCreateObject �֐��ŃC���X�^���X�������ꍇ�AVBA �G�f�B�^�[�ł̎����⊮�i�C���e���Z���X�j�������Ȃ��I�j
	- �y�A�z�z�� �L�[/���ڒǉ��zoPriceOfFruit.Add("�����S", "100�~")
	- �y�A�z�z�� ���݊m�F�zoPriceOfFruit.Exists("�����S")
	- �y�A�z�z�� �L�[�擾�iFor Each�j�zFor Each vKey In oPriceOfFruit �` Debug.print vKey �` Next 'vKey �� variant �^
	- �y�A�z�z�� ���ڎ擾�i�L�[�j�zoPriceOfFruit.Item("�����S")
	- �y�A�z�z�� �L�[�擾�i�C���f�b�N�X�j�zoPriceOfFruit.Keys()(0) '�O�I���W���i�A�N�Z�X���x������̂Œ��ӁI�j
	- �y�A�z�z�� ���ڎ擾�i�C���f�b�N�X�j�zoPriceOfFruit.Items()(0) '�O�I���W���i�A�N�Z�X���x������̂Œ��ӁI�j
	- �y�A�z�z�� �L�[�u���zoPriceOfFruit.Key("�����S") = "���"
	- �y�A�z�z�� �L�[�֘A�t���zoPriceOfFruit.Item("�����S") = "200�~"
	- �y�A�z�z�� �L�[/���ڐ��擾�zoPriceOfFruit.Count
	- �y�A�z�z�� �L�[/���ڍ폜�zoPriceOfFruit.Remove("�����S") '�w�肳�ꂽ�L�[�����݂��Ȃ��ꍇ�̓G���[
	- �y�A�z�z�� �L�[/���ڑS�폜�zoPriceOfFruit.RemoveAll
	- �y�A�z�z�� �z��ϊ��i���ځj�zasFruitPrice = oPriceOfFruit.Items 'Variant�^�z��A�O�I���W��
	- �y�A�z�z�� �z��ϊ��i�L�[�j�zasFruitName = oPriceOfFruit.Keys 'Variant�^�z��A�O�I���W��
	- �y�A�z�z�� �ݒ�ύX�zoPriceOfFruit.CompareMode = vbBinaryCompare '��/��������� /vbTextCompare�i��/��������ʂ��Ȃ�

- �y�G���[�ݒ�zOn Error Resume Next
- �y�G���[�����zOn Error Goto 0
- �y�G���[�ԍ��zErr.Number
- �y�G���[���e�zErr.Description
- �y�G���[���x���zOn Error GoTo ErrorLabel 'ErrorLabel�̓��x����
- �y���x����`�zErrorLabel:

- �yWScriptShellObject �擾�zDim objWshShell �` Set objWshShell = CreateObject("WScript.Shell")
	- �y�o�b�`���s�@�zobjWshShell.Exec("C:\test.bat") 'Exec�͕W�����o�͂ł��邪�AWSH 5.6�ȍ~���炵���T�|�[�g����Ă��Ȃ��̂Œ���
	- �y�o�b�`���s�A�zobjWshShell.Run "C:\test.bat", 0, True '�������F�E�B���h�E�̕\���X�^�C���i�E�B���h�E���\���A�ʂ̃E�B���h�E���A�N�e�B�u�j�A��O�����F�v���O�����̎��s���I������܂ŃX�N���v�g��ҋ@�����邩�ǂ����i�ڍׂ�[������](https://msdn.microsoft.com/ja-jp/library/cc364421.aspx)�j
	- �y���W�X�g���Ǎ��zobjWshShell.RegRead("HKCU\WshTest\Test1")
	- �y���W�X�g�������zobjWshShell.RegWrite("HKCU\WshTest\Test1", "test", "REG\_SZ") '�L�[/�l,�ݒ�l,�f�[�^�^
	- �y���ϐ� �l�擾�zobjWshShell.ExpandEnvironmentStrings( "%MYPATH\_CODES%" )
	- �y����t�H���_�̃p�X�擾�zobjWshShell.SpecialFolders("Desktop") '�f�X�N�g�b�v�t�H���_
		- �擾�ł���t�H���_�́uAllUsersDesktop�v �uAllUsersStartMenu�v �uAllUsersPrograms�v �uAllUsersStartup�v �uDesktop�v �uFavorites�v �uFonts�v �uMyDocuments�v �uNetHood�v �uPrintHood�v �uPrograms�v �uRecent�v �uSendTo�v �uStartMenu�v �uStartup�v �uTemplates�v
	- �y�V���[�g�J�b�g�쐬�zWith objWshShell.CreateShortcut( "c:\test\src.txt.lnk" ) �` .TargetPath = "c:\test\dst.txt" �` .Save �` End With
	- �y�V���[�g�J�b�g �w����p�X�擾�zobjWshShell.CreateShortcut( "c:\test\src.txt.lnk" ).TargetPath '�Q�Ƃ����łȂ��ύX����
	- �y�V���[�g�J�b�g �w����p�X�X�V�zWith objWshShell.CreateShortcut( "c:\test\src.txt.lnk" ) �` .TargetPath = "c:\test\dst2.txt" �` .Save �` End With

- �yFileSystemObject �擾�zDim objFSO As Object �` Set objFSO = CreateObject("Scripting.FileSystemObject")
	- �y�t�@�C���R�s�[�i���u�b�N�j�zobjFSO.CopyFile ThisWorkbook.FullName, "c:\temp\test.xlsm"
	- �y�t�@�C�� �R�s�[�@�zobjFSO.CopyFile "C:\codes\a.txt", "C:\codes\test\" '<src> <dst> [<overwrite>] �A<dst>�̖����� "\" �����邱�ƁI
	- �y�t�@�C�� �R�s�[�A�zobjFSO.CopyFile "C:\codes\a.txt", "C:\codes\test\a.txt" '<src> <dst> [<overwrite>]
	- �y�t�@�C�� �폜�zobjFSO.DeleteFile "c:\test", True
	- �y�t�@�C�� �ړ�/���l�[���zobjFSO.MoveFile "C:\codes\src.txt", "C:\codes\dst.txt"
	- �y�t�@�C�� ���݊m�F�@�zIf Dir("C:\Book1.xlsx") <> "" Then �`(����)�` Else �`(�񑶍�)�` End If
	- �y�t�@�C�� ���݊m�F�A�zobjFSO.FileExists("c:\codes\a.txt") 'True
	- �y�t�@�C�� ���擾�zobjFSO.GetFile( "C:\codes\a.txt" ).Attributes '32 (��)�l�̈Ӗ��� [[�y�t�@�C���E�t�H���_���擾�z]] �Q��
	- �y�t�@�C�� �B���t�@�C�����zobjFSO.GetFile( "C:\codes\a.txt" ).Attributes = 2
	- �y�t�@�C�� ��΃p�X�擾�zobjFSO.GetAbsolutePathName( "C:\codes\a.txt" ) ' C:\codes\a.txt
	- �y�t�@�C�� �h���C�u���擾�zobjFSO.GetDriveName( "C:\codes\a.txt" ) ' C:
	- �y�t�@�C�� �t�@�C�����擾�zobjFSO.GetFileName( "C:\codes\a.txt" ) ' a.txt
	- �y�t�@�C�� �t�@�C���x�[�X���擾�zobjFSO.GetBaseName( "C:\codes\a.txt" ) ' a
	- �y�t�@�C�� �g���q�擾�zobjFSO.GetExtensionName( "C:\codes\a.txt" ) ' txt
	- �y�t�@�C�� �e�t�H���_�p�X�擾�zobjFSO.GetParentFolderName( "C:\codes\a.txt" ) ' C:\codes
	- �y�t�H���_ �R�s�[�zobjFSO.CopyFolder "C:\codes\src", "C:\codes\dst", True '�z���t�H���_/�t�@�C�����ۂ��ƃR�s�[
	- �y�t�H���_ �폜�zobjFSO.DeleteFolder "C:\codes\test", True '�z���t�H���_/�t�@�C�����ۂ��ƍ폜
	- �y�t�H���_ �쐬�zobjFSO.CreateFolder( "C:\codes\test" ) '�e�t�H���_���Ȃ��ꍇ�A�G���[�ɂȂ�
	- �y�t�H���_ �ړ�/���l�[���zobjFSO.MoveFolder "C:\codes\src", "C:\codes\dst" '�z���t�H���_/�t�@�C�����ۂ��ƈړ�/���l�[��
	- �y�t�H���_ ���擾�zobjFSO.GetFolder( "C:\codes" ).Attributes '32 (��)�l�̈Ӗ��� [[�y�t�@�C���E�t�H���_���擾�z]] �Q��
	- �y�t�H���_ ���݊m�F�zobjFSO.FolderExists( "C:\codes" ) 'True
	- �y�t�H���_ �e�t�H���_�p�X�擾�zobjFSO.GetParentFolderName( "C:\codes\src" ) ' C:\codes

- �y�s�w�s�t�@�C���I�[�v��/�N���[�Y�zOpen �t�@�C���� For [Input|Output|Append] As #1 �` Close #1
- �y�s�w�s�t�@�C���Ǎ��i��s���j�zDo Until EOF(1) �` Line Input #1, ������ϐ� �` Loop
- �y�s�w�s�t�@�C���Ǎ��i�ꊇ�j�zsTestFile = objFSO.GetFile(�t�@�C���p�X).OpenAsTextStream.ReadAll&br()'�ԋp�l�́u�z��^�v�łȂ����Ƃɒ��ӁI���s�������܂񂾁u������^�v�ŕԋp�����I
- �y�s�w�s�t�@�C�������zPrint #1, ������ϐ�
- �y�w�k�r�t�@�C���I�[�v��/�N���[�Y�zSet wTargetBook = Workbooks.Open(sTargetBookName) �` wTargetBook.Close SaveChanges:=True

- �y�u���zReplace(������ϐ�, "  ", "")
- �y�����񌟍��i�O���j�zInStr("abcabc", "bc") ' �擪����̈ʒu�A�P�I���W���i���Ȃ��ꍇ��0���Ԃ�
- �y�����񌟍��i����j�zInStrRev("abcabc", "bc")  '��������̈ʒu�A�P�I���W���i���Ȃ��ꍇ��0���Ԃ�
- �y������ �����i�������j�zLen("�����S") '3
- �y������ �����i�o�C�g���j�zLenB("�����S") '6
- �y�����񌋍��z"abcdef" & "gh"
- �y�����񒊏o ���zLeft$("abcd", 3) 'abc
- �y�����񒊏o ���zMid$("abcdefgh", 3, 2) 'cd�i�P�I���W���j
- �y�����񒊏o �E�zRight$("abcd", 2) 'cd
- �y�������ASCII �ϊ��zAsc(����)
- �y������̐��l����zIsNumeric( sStr ) '1��True�A"a"��False�A""��False
- �yASCII�˕����� �ϊ��zChr(ASCII�R�[�h) (ex. Chr(Asc("�@") + 1) �� �A )
- �y������J��Ԃ��z"a" & String("b", 4) 'abbbb
- �y�啶�����zUCase("aaa")
- �y���������zLCase("AAA")
- �y�z��Ē�`�zReDim Preserve �z��(5) '�v�f���͂O�I���W���B���̗�ł͗v�f���U�̔z�񂪍쐬�����
- �y�z��ő�v�f���zUBound(�z��) '�ԋp�l�͂O�I���W���B�ԋp�l���R�̏ꍇ�A0�`3 �̔z��ł��邱�Ƃ�����
- �y�v�f���O�i���������j/�v�f���P�z�񔻒�zIf Sgn(asStr) = 0 Then �`���������z��` Else �`�v�f���P�z��` End If
- �y�z�� �����zJoin(�z��, ",")
- �y�z�� �����z������z�� = Split("aaa,bbb,ccc", ",") 
- �y�^�擾�i������j�zTypeName("Test") 'String
- �y�^�擾�i�l�j�zVarType("Test") '8�i�l�̏ڍׂ� [������](http://www.kanaya440.com/contents/script/vbs/function/data/var_type.html) �Q�Ɓj
- �y10��16�i���ϊ��z������ϐ� = Hex(734)
- �y16��10�i���ϊ��P�zLong�ϐ� = CLng("&H" & "FA")
- �y16��10�i���ϊ��Q�zInt�ϐ� = CInt("&H" & "FA")
- �y������16�i���\���z&HFFF0
- �y�����Ȃ�16�i���\���z&HFFF0&
- �y������ː��l�ϊ��zVal(������)
- �y���l�˕�����ϊ��zStr(���l)
- �y���s�z( vbNewLine | vbCr | vbLf | vbCrLf )
- �y���� ���� �؂�̂ć@�zFix( 99.224 ) '99
- �y���� ���� �؂�̂ćA�zInt( 99.224 ) '99
- �y���� ���� �؂�̂ć@�zFix( -99.224 ) '-99 (��)�����̏ꍇ�͐؂�グ
- �y���� ���� �؂�̂ćA�zInt( -99.224 ) '-100 (��)�����̏ꍇ�͐؂艺��
- �y���� ���� �l�̌ܓ��i���ʁj�zRound( 99.555, 0 ) '100
- �y���� ���� �l�̌ܓ��i���ʁj�zRound( 99.555, 1 ) '99.6
- �y���� ���� �l�̌ܓ��i��O�ʁj�zRound( 99.555, 2 ) '99.56
- �y���� ���� �l�̌ܓ��i���ʁj�zRound( -99.555, 0 ) '-100
- �y���� ���� �l�̌ܓ��i���ʁj�zRound( -99.555, 1 ) '-99.6
- �y���� ���� �l�̌ܓ��i��O�ʁj�zRound( -99.555, 2 ) '-99.56
- �y���� ���� �؂�グ�i���ʁj�zRound( 99.224 + 0.5, 0 ) '100
- �y���� ���� �؂�グ�i���ʁj�zRound( 99.224 + 0.05, 1 ) '99.3
- �y���� ���� �؂�グ�i���ʁj�zRound( -99.224 - 0.5, 0 ) '-100
- �y���� ���� �؂�グ�i���ʁj�zRound( -99.224 - 0.05, 1 ) '-99.3
- �y���� ���� �؂艺���i���ʁj�zRound( 99.224 - 0.5, 0 ) '99 �i���؂�̂āj
- �y���� ���� �؂艺���i���ʁj�zRound( 99.224 - 0.05, 1 ) '99.2
- �y���� ���� �؂艺���i���ʁj�zRound( -99.224 + 0.5, 0 ) '-99 �i���؂�̂āj
- �y���� ���� �؂艺���i���ʁj�zRound( -99.224 + 0.05, 1 ) '-99.2

- �y���ݎ����擾�zNow 'YYYY/DD/MM HH:MM:SS
- �y���ݔN�����擾�zDate 'YYYY/MM/DD
- �y0:00���猻�݂܂ł̌o�ߎ��ԁi�b���j�zTimer '49229.781�i13:40:29 .781�j
- �y���t��r�zIf DateDiff("s", sCmpBaseTime, sModDate ) > 0 Then �`�i sModDate ���V�����j�` ElseIf DateDiff("s", sCmpBaseTime, sModDate ) < 0 Then �`�i sModDate ���Â��j�` Else �`�i sModDate = sCmpBaseTime�j�` End If '���t�� "YYYY/MM/DD" �̕�����Ƃ��Ďw�肷��
- �yWait�����z    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) �` Sleep 1000 'ms �P��

- �y�u�b�N�쐬�zDim wTrgtBook As Workbook �` Application.SheetsInNewWorkbook = 1 �` Set wTrgtBook = Workbooks.Add 'SheetsInNewWorkbook �͍쐬���̃V�[�g�����w��
- �y�u�b�N���ɊJ���Ă��邩�m�F�zIf wCsvBook.Name <> Dir("C:\Book1.xlsx") Then �`(����)�` Else �`(�G���[)�` End If
- �y�u�b�N�ǉ��zDim bAddBook As Workbook �` Set bAddBook = Workbooks.Add
- �y�u�b�N�쐬���V�[�g�R�s�[�zThisWorkbook.Sheets(�V�[�g��).Copy: Set wTrgtBook = ActiveWorkbook
- �y�u�b�N�V�K�ۑ��zwTrgtBook.SaveAs Filename:=sFilePath
- �y�V�[�g���擾�z.Sheets.Count
- �y�V�[�g�ǉ��zDim shAddSht As Worksheet �` Set shAddSht = ThisWorkbook.Sheets.Add
- �y�V�[�g�ړ��i�����j�zThisWorkbook.Sheets(�V�[�g��).Move After:=ThisWorkbook.Sheets( ThisWorkbook.Sheets.Count )
- �y�V�[�g�폜�zApplication.DisplayAlerts = False �` .Sheets(�V�[�g��).Delete �` Application.DisplayAlerts = True
- �y�V�[�g�\��/��\���z.Sheets(�V�[�g��).Visible = (True|False)
- �y�V�[�g���בւ��z.Sheets(�V�[�g��).Move Before:=Sheets(1)
- �y��ʕ\�� ON�zApplication.ScreenUpdating = True
- �y��ʕ\�� OFF�zApplication.ScreenUpdating = False
- �y�v�Z���@�ؑ� �����zApplication.Calculation = xlCalculationAutomatic
- �y�v�Z���@�ؑ� �蓮�zApplication.Calculation = xlCalculationManual
- �y�u�b�N�Čv�Z�zApplication.Calculate
- �y�u�b�N�����Čv�Z�zApplication.CalculateFull
- �y�m�F���b�Z�[�W�}��/�\���zApplication.DisplayAlerts = (False|True)
- �y�Z�������i�s�ԍ��擾�j�z.Cells.Find("���", LookAt:=xlWhole).Row
- �y�Z�������i�s�ԍ��擾�j�z.Cells.Find("���", LookAt:=xlWhole).Column
- �y�Z���Q�ƕ��@�z.Cells(X, Y).Value 'I,X,Y�͂P�I���W��
- �y�Z���ʒu�擾�i�x���j�z.Cells(1, 1).Top '�P�ʁF�|�C���g
- �y�Z���ʒu�擾�i�w���j�z.Cells(1, 1).Left '�P�ʁF�|�C���g&br()�i�Z���̍���/���͎擾�ł��Ȃ��̂ŁA���Z���� Top/Left �Ƃ̍����Ŕ��f����K�v������j
- �y�s�폜�z.Cells(1, 1).EntireRow.Delete Shift:=xlShiftUp
- �y�s�ǉ��zApplication.CutCopyMode = False �` .Range("2:4").Insert
- �y������擾�z.Cells(�s, ��).Font.Strikethrough
- �y�g����\���zwTrgtBook.Sheets(�V�[�g��).Activate �` ActiveWindow.DisplayGridlines = False
- �y�t�H���g���擾�z������ϐ� = .Range("A1").Font.Name
- �y��/�s���`�F�b�N�i�s�j�z.Range("A1").EntireRow.Hidden
- �y��/�s���`�F�b�N�i��j�z.Range("A1").EntireColumn.Hidden
- �y�t�H���g�T�C�Y�ύX�z.Range("A1").Font.Size = 14
- �y�t�H���g�J���[�ύX�z.Range("A1").Font.Color = RGB(0, 255, 0)
- �y�t�H���g�����ύX�z.Range("A1").Font.Bold = True
- �y�t�H���g�����ύX�z.Range("A1").Font.Underline = True
- �y�w�i�F�ύX�z.Range("A1").Interior.Color = RGB(255, 255, 0)
- �y�r���i�i�q�j�ݒ�z.Range("A1:C3").Borders.LineStyle = xlContinuous
- �y�Z�������z.Range("A1:C3").MergeCells = True
- �y�Z���I��͈͓������z.Range("A1:C3").HorizontalAlignment = ( xlCenterAcrossSelection | xlGeneral )
- �y�ŏI�s�擾�z���l�ϐ� = .Cells(.Rows.Count, ��).End(xlUp).Row
- �y�ŏI��擾�z���l�ϐ� = .Cells(�s, .Columns.Count).End(xlToLeft).Column
- �y�ŏI�s�擾�i�S��̒��ōő�j�z���l�ϐ� = .Sheets(�V�[�g��).UsedRange.Rows.Count + 1
- �y�ŏI��擾�i�S�s�̒��ōő�j�z���l�ϐ� = .Sheets(�V�[�g��).UsedRange.Columns.Count + 1
- �y�I��͈͈ʒu�擾�i�擪�s�j�zSelection(1).Row
- �y�I��͈͈ʒu�擾�i�����s�j�zSelection(Selection.Count).Row
- �y�I��͈͈ʒu�擾�i�擪��j�zSelection(1).Column
- �y�I��͈͈ʒu�擾�i������j�zSelection(Selection.Count).Column
- �yRange�I�u�W�F�N�g�̍s/��ԍ��w��z.Range(.Cells(1, 1), .Cells(6, 3))
- �y�Z���R�s�[�i�����ێ��j�z.Range("A1:B9").Copy Destination:=ThisWorkBook.Sheets(�V�[�g���Q).Range("B1")
- �y�Z���\�[�g�z.Range(.Cells(1, 1), .Cells(.Rows.Count, 2)).Sort Key1:=.Cells(1, 2) ,order1:=xlAscending
- �y�͈̓Z�� �l�N���A�i�������̂܂܁j�z.Range("A1:A2").ClearContents
- �y�͈̓Z�� �����N���A�z.Range("A1:A2").ClearFormats
- �y�͈̓Z�� �����\��t���z.Range("A1:A2").PasteSpecial (xlPasteFormats) '�����FxlPasteFormulas�A�l�FxlPasteValues�A�����FxlPasteFormats�A�R�����g�FxlPasteComments�A���͋K���FxlPasteValidation
- �y�󔒃Z���I���z.Range("A1:CV100").SpecialCells(xlCellTypeBlanks).Select
- �y�񕝕ύX�z.Range("A1").ColumnWidth = 5 '�s��RowHeight
- �y�����񕝒����z.Range(.Cells(4, 2), .Cells(9, 2)).Columns.AutoFit '�s��Rows
- �y�����񕝒����i�S�̈�j�z.Sheets(�V�[�g��).UsedRange.Columns.AutoFit '�s��Rows
- �y�I���Z���ɑ΂��āu�I��͈͓��Œ����v�zSelection.HorizontalAlignment = xlCenterAcrossSelection
- �y�R�s�[/�؂��胂�[�h�����zApplication.CutCopyMode = False
- �y�O���[�v���i�s�j�z.Range( .Rows( lStrtRow ), .Rows( lLastRow ) ).Group '������ Group �� Ungroup
- �y�O���[�v���i�s�j�z.Range( .Columns( lStrtRow ), .Columns( lLastRow ) ).Group '������ Group �� Ungroup
- �y�A�E�g���C���ݒ�ύX�i�㉺�j�z.Sheets(�V�[�g��).Outline.SummaryRow = ( xlBelow | xlAbove )
- �y�A�E�g���C���ݒ�ύX�i���E�j�z.Sheets(�V�[�g��).Outline.SummaryColumn = ( xlRight | xlLeft )
- �y�A�E�g���C���ݒ�ύX�i�����j�z.Sheets(�V�[�g��).Outline.AutomaticStyles = ( True | False )

- �yChartObject��`�zDim oChartObj As ChartObject �` Set oChartObj = ThisWorkbook.Sheets(�V�[�g��).ChartObjects(1)
	- �y�O���t �ǉ��zSet oChartObj = .Sheets(�V�[�g��).ChartObjects.Add( XPOS, YPOS, WIDTH, HEIGHT ) 'XPOS, YPOS, WIDTH, HEIGHT�̒P�ʂ̓|�C���g
	- �y�O���t �폜�zoChartObj.Delete
	- �y�O���t �R�s�[�zoChartObj.Chart.ChartArea.Copy
	- �y�O���t �ړ��i�x���j�zoChartObj.Top = 10
	- �y�O���t �ړ��i�w���j�zoChartObj.Left = 20
	- �y�O���t �T�C�Y�ύX�i���j�zoChartObj.Width = 200
	- �y�O���t �T�C�Y�ύX�i�����j�zoChartObj.Height = 300
	- �y�O���t ��ʁzoChartObj.Chart.ChartType = xlXYScatterLines 'xlXYScatterLines:�܂���t���U�z�}�AxlLine:�܂���A...
	- �y�O���t �f�[�^�͈͕ύX�zoChartObj.Chart.SetSourceData Source:=Union(rXAxsRng, rDataRng) '�f�[�^�͈͎w��
	- �y�O���t �w�� �^�C�g�� �L���zoChartObj.Chart.Axes(xlCategory).HasTitle = True
	- �y�O���t �w�� �^�C�g�� �ύX�zoChartObj.Chart.Axes(xlCategory).AxisTitle.Text = "Test Axis X"
	- �y�O���t �w�� �ڐ��� �L���zoChartObj.Chart.Axes(xlCategory).HasMajorGridlines = True
	- �y�O���t �w�� �ڐ��� �F�zoChartObj.Chart.Axes(xlCategory).MajorGridlines.Border.Color = RGB(217, 217, 217)
	- �y�O���t �w�� �ڐ��� �����zoChartObj.Chart.Axes(xlCategory).MajorGridlines.Border.Weight = 2
	- �y�O���t �w�� �ڐ��� �X�^�C���zoChartObj.Chart.Axes(xlCategory).MajorGridlines.Border.LineStyle = (xlContinuous|xlDot|xlDouble|xlLineStyleNone|...)
	- �y�O���t �w�� �⏕�ڐ��� �V�z��L�y�O���t �w�� �ڐ��� �`�z�́uMajorGridlines�v���uMinorGridlines�v�ɕύX
	- �y�O���t �w�� �ŏ��l �����zoChartObj.Chart.Axes(xlCategory).MinimumScaleIsAuto = False
	- �y�O���t �w�� �ő�l �����zoChartObj.Chart.Axes(xlCategory).MaximumScaleIsAuto = False
	- �y�O���t �w�� �ŏ��l �ݒ�zoChartObj.Chart.Axes(xlCategory).MinimumScale = 0
	- �y�O���t �w�� �ő�l �ݒ�zoChartObj.Chart.Axes(xlCategory).MaximumScale = 100
	- �y�O���t �w�� �c���Ƃ̌�_�zoChartObj.Chart.Axes(xlCategory).Crosses = (xlMinimum|xlMaximum|xlAutomatic)
	- �y�O���t �x�� �V�z��L�y�O���t �w�� �`�z�́uxlCategory�v���uxlValue�v�ɕύX
	- �y�O���t �^�C�g�� �L���zoChartObj.Chart.HasTitle = True
	- �y�O���t �^�C�g�� �ύX�zoChartObj.Chart.ChartTitle.Text = "Test Title"
	- �y�O���t �^�C�g�� �O���t�ɏd�˂�zoChartObj.Chart.ChartTitle.IncludeInLayout = False 'False:�d�˂�ATrue:�d�˂Ȃ�
	- �y�O���t �}�� �L���zoChartObj.Chart.HasLegend = True
	- �y�O���t �}�� �ʒu�zoChartObj.Chart.Legend.Position = (xlLegendPositionTop|xlLegendPositionBottom|xlLegendPositionLeft|xlLegendPositionRight|...)
	- �y�O���t �}�� �O���t�ɏd�˂�zoChartObj.Chart.Legend.IncludeInLayout = False 'False:�d�˂�ATrue:�d�˂Ȃ�
- �y�O���t �摜�Ƃ��ē\��t���z.Sheets(�V�[�g��).PasteSpecial Format:="�} (JPEG)", Link:=False, DisplayAsIcon:=False

- �y���[�N�V�[�g�֐��zApplication.WorksheetFunction.VLookup(.Range("C1"), .Range("A1:B7"), 2, False)

- �y�t�H�[�� ���[�h�zDim goPrgrsBar As New ProgressBar �` goPrgrsBar.Show vbModeless
- �y�t�H�[�� �A�����[�h�zgoPrgrsBar.Hide �` Unload goPrgrsBar �` Set goPrgrsBar = Nothing

- �yCollection ���ڎ擾�i�L�[�j�zcCollection.Item("�����S")
- �yCollection ���ڎ擾�i�C���f�b�N�X�j�zcCollection(0)
- �yCollection ���ڎ擾�iFor Each�j�zFor Each vItem In cCollection �` Debug.print vItem �` Next

- �y�`�F�b�N�{�b�N�X�l�擾�i�t�H�[���R���g���[���j�zlChk = ThisWorkbook.Sheets(�V�[�g��).CheckBoxes(1).Value 'On:1 Off:-4146
- �y���[�U�t�H�[���\�����̃L�[����zPrivate Sub xxx\_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) �` End Sub
	- xxx �̓t�H�[�J�X���̃t�H�[�����B�t�H�[�� xxx �Ƀt�H�[�J�X������ꍇ�����AKeyUp �C�x���g����������B
	- �ǂ̃t�H�[���Ƀt�H�[�J�X�������Ă����� KeyUp �C�x���g����肽���ꍇ�A�S�t�H�[���ɑ΂��ď�L�C�x���g�����I

# ���C�u����
- [�y�w�k�r�t�@�C�����݊m�F�`�I�[�v���`�N���[�Y�z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/ExcelFile.bas)
- [�y�s�w�s�t�@�C�����݊m�F�`�I�[�v���`�N���[�Y�z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/FileSys.bas)
- [�y�s�w�s�t�@�C�����݊m�F�`�I�[�v���`�N���[�Y�i�L�����N�^�Z�b�g�w��j�z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/FileSys.bas)
- [�y�t�H���_�쐬�i�ċA�����j�z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/FileSys.bas)
- [�y�t�@�C�����ꗗ�擾�z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/FileSys.bas)
- [�y�t�@�C���E�f�B���N�g�� ���ʁz](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/FileSys.bas)
- [�y�t�@�C���E�f�B���N�g�� �I���_�C�A���O�z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/FileSys.bas)
- [�y�t�@�C���E�t�H���_���擾�z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/FileInfo.bas)
- [�y�R�}���h���s�z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/SysCmd.bas)
- [�y�v���O���X�o�[�\���z]
- [�y���\����z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/StopWatch.cls)
- [�y�G���[�����z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/Error.bas)
- [�y�z��� Push, Pop �֐��z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/ArrayMng.bas)
- [�y�C�~�f�B�G�C�g�E�B���h�E�N���A�z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/VbaMng.bas)
- [�y�V�[�g�ꗗ�z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/ExcelOpe.bas)
- [�y�c���[�}�I�u�W�F�N�g�����z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/ExcelOpe.bas)
- [�y�t�@�C�����E�t�H���_�����o�֐��z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/StringMng.bas)
- [�y�Z���͈͕�������/�����֐��z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/ExcelOpe.bas)
- [�y�Z�������擾�֐��z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/ExcelOpe.bas)
- [�y�r�b�g���Z�֐��z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/ExcelOpe.bas)
- [�y�L�[���M�z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/SendKeys.bas)
- [�y����\��t���z](https://github.com/draemonash2/codes/blob/master/vba/MacroBook/lib/SpecialPaste.bas)

# �T���v���R�[�h
- [[�y�v���O�����e���v���[�g�z]]
- �y�V�[�g���݊m�F�z
	```vba
	'�V�[�g���݊m�F
	Dim bIsSheetExist As Boolean
	bIsSheetExist = False
	For Each wSheet In ThisWorkbook.Worksheets
		If wSheet.Name = sTrgtSheetName Then
			bIsSheetExist = True
		Else
			'Do Nothing
		End If
	Next wSheet
	```
- �y�Z�������i���݂��Ȃ��ꍇ���l���j�z
	- �Z������
		```vba
		Dim rFindResult As Range
		Dim sSrchKeyword As String
		Dim lSrchCellRow As Long
		Dim lSrchCellClm As Long
		
		With ThisWorkbook.Sheets(INPUT_SHEET_NAME)
			sSrchKeyword = KYWD_INPUT_DIR_PATH
			Set rFindResult = .Cells.Find(sSrchKeyword, LookAt:=xlWhole)
			If rFindResult Is Nothing Then
				MsgBox sSrchKeyword & "��������܂���ł���"
			Else
				lSrchCellRow = rFindResult.Row
				lSrchCellClm = rFindResult.Column
			End If
		End With
		```
	- ���߃Z�����擾
		```vba
		Public Type T_EXCEL_NEAR_CELL_DATA
			bIsCellDataExist As Boolean
			lRow As Long
			lClm As Long
			sCellValue As String
		End Type
		
		Public Function GetNearCellData( _
			ByVal shTrgtSht As Worksheet, _
			ByVal sSrchKey As String, _
			ByVal lRowOffset As Long, _
			ByVal lClmOffset As Long _
		) As T_EXCEL_NEAR_CELL_DATA
			Dim rFindResult As Range
			Set rFindResult = shTrgtSht.Cells.Find(sSrchKey, LookAt:=xlWhole)
			If rFindResult Is Nothing Then
				GetNearCellData.bIsCellDataExist = False
			Else
				GetNearCellData.bIsCellDataExist = True
				GetNearCellData.lRow = rFindResult.Row + lRowOffset
				GetNearCellData.lClm = rFindResult.Column + lClmOffset
				GetNearCellData.sCellValue = shTrgtSht.Cells( _
													GetNearCellData.lRow, _
													GetNearCellData.lClm _
												 ).Value
			End If
		End Function
		```
- �y�Z���d���폜�z
	```vba
	For iRowIdx = .Cells(Rows.Count, 1).End(xlUp).Row To 2 Step -1
		For iChkTargetRowIdx = iRowIdx - 1 To 2 Step -1
			If .Cells(iRowIdx, 1) = .Cells(iChkTargetRowIdx, 1) Then
				.Cells(iChkTargetRowIdx, 1).EntireRow.Delete Shift:=xlShiftUp
			Else
				'Do Nothing
			End If
		Next iChkTargetRowIdx
	Next iRowIdx
	```
- [[�y�O���t�쐬���摜�ϊ��z]]
- �y�e�L�X�g�{�b�N�X�쐬�z
	```vba
	Dim spTxtBox As Shape
	Set spTxtBox = shTrgtSht.Shapes.AddTextbox( _
						Orientation:=msoTextOrientationHorizontal, _
						Left:=10, _
						Top:=10, _
						Width:=10, _
						Height:=10 _
				   )
	With spTxtBox
		.TextFrame.Characters.Text = "�G���[�t���[���Ȃ�"
		.TextFrame.AutoSize = True
	End With
	```
- �y���K�\���z
	- ���K�\���ƕ����񌟍� InStr �̑��x�ɂ��Ă̍l�@�c
		- �}�b�`���O�� 500000 ��J��Ԃ��ƁA���K�\���͖� 2800 ms�AInStr �ł͖� 50 ms�B
		- ���K�\���� 56 �{�x���I
	- �Q�lURL: http://officetanaka.net/excel/vba/tips/tips38.htm
		```vba
		Dim sSearchPattern As String
		Dim sTargetStr As String
		 
		Dim iMatchIdx As Integer
		Dim iSubMatchIdx As Integer
		 
		Dim oMatchResult As Object
		Dim oRegExp As Object
		Set oRegExp = CreateObject("VBScript.RegExp")
		 
		sSearchPattern = "(\w+)\((\w+) (\w+)\)"
		sTargetStr = "TestFunc01(char aaa) TestFunc02(int bbb)"
		 
		oRegExp.Pattern = sSearchPattern               '�����p�^�[����ݒ�
		oRegExp.IgnoreCase = True                      '�啶���Ə���������ʂ��Ȃ�
		oRegExp.Global = True                          '������S�̂�����
		Set oMatchResult = oRegExp.Execute(sTargetStr) '�p�^�[���}�b�`���s
		 
		Debug.Print oMatchResult.Count
		For iMatchIdx = 0 To oMatchResult.Count - 1
			Debug.Print oMatchResult(iMatchIdx).SubMatches.Count
			For iSubMatchIdx = 0 To oMatchResult(iMatchIdx).SubMatches.Count - 1
				Debug.Print oMatchResult(iMatchIdx).SubMatches(iSubMatchIdx)
			Next iSubMatchIdx
		Next iMatchIdx
		```
- �y�A�z�z��z
	- �A�z�z����g�p���邱�Ƃɂ��A�c��Ȍ���(��)�̌��������I�ɑ����Ȃ�B
		- �� �����Ώ� 15 ���ȏ�i���ꖢ�����Ɣz�񌟍��̂ق��������B.Exists API �̃I�[�o�[�w�b�h�̂����H�j
			```vba
			Dim oPriceOfFruit As Object
			Set oPriceOfFruit = CreateObject("Scripting.Dictionary")
			 
			'�o�^
			oPriceOfFruit.Add "�����S", "100�~"
			oPriceOfFruit.Add "�C�`�S", "400�~"
			oPriceOfFruit.Add "������", "1000�~"
			 
			'���݊m�F
			If oPriceOfFruit.Exists("�����S") Then
				Debug.Print "Exists!"
			Else
				Debug.Print "Not Exists!"
			End If
			 
			Set oPriceOfFruit = Nothing
			```

# ���̑�
- �y�^�ꗗ�z

| �f�[�^�^ | ���� | ������� | �i�[�ł���͈� |
|:---------|:-----|:-----------|:---------------|
| Integer | �����^ | 2�o�C�g | -32,768 �` 32,767 |
| Long | �������^ | 4�o�C�g | -2,147,483,648 �` 2,147,483,647 |
| Single | �P���x���������_���^ | 4�o�C�g | -3.402823E38 �` -1.401298E-45(���̒l) 1.401298E-45 �` 3.402823E38(���̒l) |
| Double | �{���x���������_���^ | 8�o�C�g | -1.79769313486232E308 �` -4.94065645841247E-324(���̒l) 4.94065645841247E-324 �` 1.79769313486232E308(���̒l) |
| Currency | �ʉ݌^ | 8 �o�C�g | -922,337,203,685,477.5808 �` 922,337,203,685,477.5807 |
| String | ������^ | 2�o�C�g | �ő��20�������܂� |
| Date | ���t�^ | 8 �o�C�g | ����100 �N1��1���`����9999�N12��31���܂ł̓��t�Ǝ��� |
| Object | �I�u�W�F�N�g�^ | 4 �o�C | �I�u�W�F�N�g���Q�Ƃ���f�[�^�^ |
| Variant | �o���A���g�^ | 16�o�C�g | �ϒ��̕�����^�͈̔͂Ɠ����B |
| Boolean | �u�[���^ | 2 �o�C�g | �^ (True) �܂��͋U (False) |

- �y���Z�q�z

| ���Z�q | �Ӗ� | �ϐ��ւ̑���� |
|:-------|:-----|:---------------|
| �{ | ���Z���� | i = 15 + 5 ( �� �̒l��20) |
| �| | ���Z���� | i = 15 - 5 ( �� �̒l��10) |
| �� | ��Z���� | i = 5 * 4 ( �� �̒l��20) |
| �^ | ���Z���� | i = 15 / 5 ( �� �̒l��3) |
| �� | ���Z�̏� | i = 15 \ 2 ( �� �̒l��7) |
| Mod | ���Z�̗]�� | i = 15 Mod 2 ( �� �̒l��1) |
| �O | �ׂ��悷�� | i = 2 ^ 5 ( �� �̒l��32) |

- �y�G���[��ʁz
	- ex) CVErr(xlErrNum)
	
| �萔       | �G���[�ԍ� | �Z���̃G���[�l |
|:-----------|:-----------|:---------------|
| xlErrDiv0  | 2007       | #DIV/0!        |
| XlErrNA    | 2042       | #N/A           |
| xlErrName  | 2029       | #NAME?         |
| XlErrNull  | 2000       | #NULL!         |
| XlErrNum   | 2036       | #NUM!          |
| XlErrRef   | 2023       | #REF!          |
| XlErrValue | 2015       | #VALUE!        |
