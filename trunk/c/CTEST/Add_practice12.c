/*======================================================================
	�v���W�F�N�g��	�FC�����b
	�t�@�C����		�Fpracticee15-1.c
	�@�\			�F�\����
	�C������		�F1.00	2010/2/4	���C ���Y	�쐬
	Copyright(c) 2010 eSOL emBex inc. All Rights Reserved.
======================================================================*/

/* �w�b�_�t�@�C���ǂݍ��� */
#include <stdio.h>
#include <string.h>

/* �萔��` */
#define NUM_EMP	2	/* �l�� */

/* �\���̐錾 */
struct ST_EMPLOYEE {
	int id;			/* �Ј�id�ԍ� 	*/
	char name[64];	/* ���O			*/
	int age;		/* �N��			*/
	int length;		/* �g��			*/
	int weight;		/* �̏d			*/
};

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
					�F1.00	2010/04/16	���C ���Y	#define��`
======================================================================*/

int main(void)
{
	/* �ϐ��錾 */
	struct ST_EMPLOYEE st_member[NUM_EMP];	/* �\���̃����o�f�[�^ */
	int i;								/* ���[�v�p�ϐ� */
	
	/* ���[�U���� */
	for (i = 0; i < NUM_EMP; i++) {
		printf("id�ԍ�����͂��Ă��������F");
		scanf("%d", &st_member[i].id);
		printf("���O����͂��Ă��������F");
		scanf("%s", &st_member[i].name);
		printf("�N�����͂��Ă��������F");
		scanf("%d", &st_member[i].age);
		printf("�g������͂��Ă��������F");
		scanf("%d", &st_member[i].length);
		printf("�̏d����͂��Ă��������F");
		scanf("%d", &st_member[i].weight);
		printf("\n");
	}
	printf("\n");
	printf("�f�[�^�͍\���̂֊i�[����܂����I\n");
	printf("\n");
	
	/* ���ʏo�� */
	for(i = 0; i < NUM_EMP; i++) {
		printf("id:%d\n", st_member[i].id);
		printf("���O:%s\t", st_member[i].name);
		printf("�N��:%d\t", st_member[i].age);
		printf("�g��:%d\t", st_member[i].length);
		printf("�̏d:%d", st_member[i].weight);
		printf(" \n");
	}
	
	/* ����return */
	return 0;
	
}
