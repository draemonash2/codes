/*======================================================================
	�v���W�F�N�g��	�FC����v���O���~���O
	�t�@�C����		�Fpractice24.c
	�@�\			�F�L���X�g���Z�q(Sample15)
	�C������		�F1.00	2010/3/14	���q�@�N�Y	�쐬
	Copyright(c) 2010 eSOL emBex inc. All Rights Reserved.
======================================================================*/

/* �w�b�_�t�@�C���ǂݍ��� */
#include <stdio.h>

/* ��` */
#define DEF_A
#define DEF_B
#define BOOL 0

/*======================================================================
	�֐���			�Fmain
	�@�\			�F���C������
	���͈�������	�FNone
	�o�͈�������	�FNone
	�߂�l			�F�I�����i��ɂO�j
	���͏��		�FNone
	�o�͏��		�FNone
	���L����		�FNone
	�C������		�F1.00	2010/2/18	���C ���Y	�쐬
======================================================================*/

int main(void)
{
	/* �ϐ���` */
	long a = 0xABCD;//long�ϐ�a

//	printf("���̂܂ܕ\��  :0x%x\n",a);//a�����̂܂܃w�L�T�ŕ\��
	printf("char  :%2d\n",(char)a);//a��char�^�ŃL���X�g�ϊ����ăf�V�}���ŕ\��
	printf("short :%4d\n",(short)a);//a��short�^�ŃL���X�g�ϊ����ăf�V�}���ŕ\��
//	printf("char  :0x%2x\n",(char)a);//a��char�^�ŃL���X�g�ϊ����ăw�L�T�ŕ\��
//	printf("short :0x%4x\n",(short)a);//a��short�^�ŃL���X�g�ϊ����ăw�L�T�ŕ\��
	
	/* ����return */
	return 0;
}
