/*======================================================================
	�v���W�F�N�g��	�FC����v���O���~���O
	�t�@�C����		�Fpractice23.c
	�@�\			�F�v���v���Z�b�T����
	�C������		�F1.00	2010/2/18	���C ���Y	�쐬
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
	/* #ifdef�`#endif */
#ifdef DEF_A
	printf("�@IFDEF\n");
#else
	printf("�AIFDEF_ELSE\n");
#endif
	
	/* #ifndef�`#endif */
#ifndef DEF_B
	printf("�BIFNDEF\n");
#else
	printf("�CIFNDEF_ELSE\n");
#endif
	
	/* #if�`#endif */
#if BOOL
	printf("�DIF\n");
#else
	printf("�EIF_ELSE\n");
#endif
	
	/* ����return */
	return 0;
}
