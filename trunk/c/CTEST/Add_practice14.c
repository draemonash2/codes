/*======================================================================
	�v���W�F�N�g��	�FC�����b
	�t�@�C����		�FAdd_practice14.c
	�@�\			�F���R���ߖ��
	�C������		�F1.00	2010/3/14	���q�@�N�Y	�쐬
	Copyright(c) 2010 eSOL emBex inc. All Rights Reserved.
======================================================================*/

/* �w�b�_�t�@�C���ǂݍ��� */
#include <stdio.h>

/* �萔��` */
#define NUM_ARR	300 //����v�f�̏���l

/*======================================================================
	�֐���			�Fmain
	�@�\			�F���C������
	���͈�������	�FNone
	�o�͈�������	�FNone
	�߂�l			�F�I�����i��ɂO�j
	���͏��		�FNone
	�o�͏��		�FNone
	���L����		�FNone
	�C������		�F1.00	2010/3/14	���q�@�N�Y	�쐬
======================================================================*/

int main(void)
{
	/* �ϐ��錾 */
	long idata1=1;	/* ���[�U�f�[�^1 	*/
	long idata2=1;	/* ���[�U�f�[�^2 	*/
	long Addition=0;/* ���Z����			*/
	long Answer=0;	/* ���v����			*/
	
	/* ���Z���[�v */
	printf("%d,%d",idata1,idata2) ;	//�ŏ��̂Q�̐�������ʕ\��
	Answer = idata1 +idata2 ;		//�ŏ��̂Q�̐��������Z���č��v�����ɑ��
	while (1) {/* ���[�v	*/
		Addition = idata1+idata2;//1�ڂƂQ�ڂ̐��l�����Z���ĉ��Z���ʂɑ��
		/* �I������ */
		if(NUM_ARR < Addition){//���Z����=����v�f�̏��+1���H
			printf("\n");//���s����
			/* ���ʕ\�� */
			printf("���v�F%d\n",Answer);
			/* ����return */
			return 0;
		}
		printf(",%d",Addition);
		Answer = Answer + Addition;	//���Z���ʂ����߂�B
		idata1 = idata2 ;			//�l�V�t�g
		idata2 = Addition;			//�l�V�t�g
	}
}
