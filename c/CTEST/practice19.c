/*======================================================================
	�v���W�F�N�g��	�FC�����b
	�t�@�C����		�Fpractice19.c
	�@�\			�F�|�C���^�A�h���X�\���֐�
	�C������		�F1.00	2010/2/4	���C ���Y	�쐬
	Copyright(c) 2010 eSOL emBex inc. All Rights Reserved.
======================================================================*/

/* �w�b�_�t�@�C���ǂݍ��� */
#include <stdio.h>

/*======================================================================
	�֐���			�Fmain
	�@�\			�F���C������
	���͈�������	�FNone
	�o�͈�������	�FNone
	�߂�l			�F�I�����i��ɂO�j
	���͏��		�FNone
	�o�͏��		�FNone
	���L����		�FNone
	�C������		�F1.00	2010/2/4	���C ���Y	�쐬
======================================================================*/

int main(void)
{
	/* �ϐ��錾 */
	int iData=12;		/* int�f�[�^ 		*/
	char cData[10]={0};	/* �z��char�f�[�^ 	*/
//	char cData[10]	;	/* �z��char�f�[�^ 	*/
	long lData=13;		/* long�f�[�^		*/
	/* ���[�U���� */
	printf("------------------------------\n");
	printf("iData�̃A�h���X=0x%p\n",&iData);
	printf("iData�̒l=0x%x\n",iData);
	printf("------------------------------\n");
	printf("cData�̃A�h���X=0x%p\n",cData);
	printf("cData[0]�`[9]�̒l=[0x%x][0x%x][0x%x][0x%x][0x%x][0x%x][0x%x][0x%x][0x%x][0x%x]\n",
							cData[0],cData[1],cData[2],cData[3],cData[4],cData[5],cData[6],cData[7],cData[8],cData[9]);
	printf("------------------------------\n");
	printf("lData�̃A�h���X=0x%p\n",&lData);
	printf("lData�̒l=0x%x\n",lData);
	printf("------------------------------\n");
	
	/* ����return */
	return 0;
	
}
