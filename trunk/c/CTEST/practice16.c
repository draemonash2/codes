/*======================================================================
	�v���W�F�N�g��	�FC�����b
	�t�@�C����		�Fpractice16.c
	�@�\			�F���t�\��
	�C������		�F1.00	2010/2/4	���C ���Y	�쐬
	Copyright(c) 2010 eSOL emBex inc. All Rights Reserved.
======================================================================*/

/* �w�b�_�t�@�C���ǂݍ��� */
#include <stdio.h>
#include <string.h>

/* �\���̐錾 */
struct ST_EMPLOYEE {
	int id;
	char name[64];
	int age;
	int length;
	int weight;
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
	�C������		�F1.00	2010/2/4	���C ���Y	�쐬
======================================================================*/

int main(void)
{
	/* �ϐ��錾 */
	struct ST_EMPLOYEE st_member;	/* �����o�[�f�[�^ */
	
	/* ���[�U���� */
	printf("ID�ԍ�����͂��Ă��������F");
	scanf("%d", &st_member.id);
	printf("���O����͂��Ă��������F");
	scanf("%s", &st_member.name);
	printf("�N�����͂��Ă��������F");
	scanf("%d", &st_member.age);
	printf("�g������͂��Ă��������F");
	scanf("%d", &st_member.length);
	printf("�̏d����͂��Ă��������F");
	scanf("%d", &st_member.weight);
	printf("\n");
	printf("�f�[�^�͍\���̂֊i�[����܂����I\n");
	printf("\n");
	
	/* ���ʏo�� */
	printf("ID  �F%d\n", st_member.id);
	printf("���O�F%s\n", st_member.name);
	printf("�N��F%d\n", st_member.age);
	printf("�g���F%d\n", st_member.length);
	printf("�̏d�F%d\n", st_member.weight);
	
	/* ����return */
	return 0;
	
}
