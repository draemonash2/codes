/*======================================================================
	�v���W�F�N�g��	�FC�����b
	�t�@�C����		�FAdd_practice18.c
	�@�\			�F�ŏ����{�������߂�v���O����
	�C������		�F1.00	2011/3/16	���q�@�N�Y	�쐬
	Copyright(c) 2010 eSOL emBex inc. All Rights Reserved.
======================================================================*/

/* �w�b�_�t�@�C���ǂݍ��� */
#include <stdio.h>

/* �萔��` */


/*======================================================================
	�֐���			�Fmain
	�@�\			�F���C������
	���͈�������	�FNone
	�o�͈�������	�FNone
	�߂�l			�F�I�����i��ɂO�j
	���͏��		�FNone
	�o�͏��		�FNone
	���L����		�FNone
	�C������		�F1.00	2010/02/04	���C ���Y	�쐬
					�F1.01	2010/02/15	���C ���Y	���t�@�N�^�����O
======================================================================*/

int main(void)
{
	/* �ϐ���` */
	int in1;					/* ���͐���1 	    */
	int in2;					/* ���͐���2 	    */
	int multiplier1 = 1;		/* �搔1			*/
	int multiplier2 = 1;		/* �搔2			*/
	int least_common_multiple;	/* �ŏ����{��		*/

	/* �Q�̐�������� */
	/* ���[�U���� */
	printf("��ڂ̐�������͂��ĉ������B>");
	scanf("%d", &in1);
	printf("��ڂ̐�������͂��ĉ������B>");
	scanf("%d", &in2);

	while( 1 ){//���[�v

		if(in1 * multiplier1 == in2 * multiplier2){//�ŏ����{�������H
			least_common_multiple = in1 * multiplier1;//�ŏ����{���l���i�[
//			printf("[2]multiplier1=%d multiplier2=%d \n",multiplier1,multiplier2);
			break;//���[�v�𔲂���

		}else if(in1 * multiplier1 < in2 * multiplier2){//���͐����Q�̏�Z���ʂ��傫���ꍇ

//			printf("[1]multiplier1=%d multiplier2=%d \n",multiplier1,multiplier2);

			multiplier1++;	//�搔1���C���N�������g
			multiplier2 = 0;//�搔2��������
		}
		multiplier2++;
	}

	printf("%d��%d�̍ŏ����{��:%d\n",in1,in2,least_common_multiple);
	return(0);
}
