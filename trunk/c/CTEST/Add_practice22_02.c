/*======================================================================
	�v���W�F�N�g��	�FC�����b
	�t�@�C����		�FAdd_practice22_02.c
	�@�\			�F���H����Ă݂�H���̂Q
	�C������		�F1.00	2011/3/14	���q�@�N�Y??	�쐬
	Copyright(c) 2010 eSOL emBex inc. All Rights Reserved.
======================================================================*/
//�����G�F��������
//2009/06/09
//EXILIAS
//
//�������������������Ɏ����������H���쐬����v���O�����B
//�Ȃ��A���H�����̃A���S���Y���Ƃ��Ĉȉ��̃T�C�g�ŏЉ��Ă���
//�u�_�|���@�v��p�����B
//http://www5d.biglobe.ne.jp/~stssk/maze/make.html

#include <stdio.h>
#define N 1000
#define X 31 //���H�̑傫���i��j
#define Y 31


/* ���l����� */
int scanInt(void){
	int a;
	scanf("%d", &a);
	return a;
}

/*======================================================================
	�֐���			�Fmain
	�@�\			�F���C������
	���͈�������	�FNone
	�o�͈�������	�FNone
	�߂�l			�F�I�����i��ɂO�j
	���͏��		�FNone
	�o�͏��		�FNone
	���L����		�FNone
	�C������		�F1.00	2011/3/14	���q�@�N�Y??	�쐬
======================================================================*/
/* main()�֐� */
int main(void){
	int a, b, m, n, i, j;
	int x[N], maze[Y+1][X+1];

	//�����������j�b�g
	a = 7;
	b = 3;
	m = 733; //���̑g�ݍ��킹��732�̎����̗����𓾂���
	n = m-1; //�ł���\���̂���ő�l
	x[0] = b;
	  //�����𐶐�
	for(i=0; i<n; i++){
		x[i+1] = (a*x[i])%m;
		if(x[i+1] == b){ n=i; }
	}

	//�O�g�����
	for(i=1; i<=X; i++){
		for(j=1; j<=Y; j++){
			if(i==1 || i==X){
				maze[i][j] = 1;
			}else if(j==1 || j==Y){
				maze[i][j] = 1; 
			}else if(i%2!=0 && j%2!=0){
				maze[i][j] = 1; 
			}else{
				maze[i][j] = 0;
			}
		}
	}

	//(i, 3)�̃��[���̖_��|��
	printf("100�ȉ��̍D���Ȑ�����͂��Ă��������F");
	a = 0;
	while(a=scanInt(), a==0 || a>=100);
	i = 3;
	for(j=3; j<=X-2; j+=2, a++){
		switch(x[a]%4){ //1/4�̊m���œ|�����������߂�
			case 0: maze[i-1][j] = 1;			//��֓|��
				break;
			case 1: 
				if(maze[i][j-1] !=1 ){
					maze[i][j-1] = 1;		//���֓|��
				}else{ //���łɓ|��Ă��āA�|����Ȃ��ꍇ
					if(x[a+1]%2 == 0){
						maze[i-1][j] = 1;	//��֓|��
					}else{
						maze[i+1][j] = 1;	//���֓|��
					}
				}
				break;
			case 2: maze[i+1][j] = 1;			//���֓|��
				break;
			case 3: maze[i][j+1] = 1;			//�E�֓|��
				break;
		}
	}

	//(j, i)�̃��[���̖_��|��
	for(i=5; i<=Y-2; i+=2,a++){
		for(j=3; j<=X-2; j+=2,a++){
			switch(x[a]%3){
				//case 0: break;
				case 0: 
					if(maze[i][j-1] !=1 ){
						maze[i][j-1] = 1;		//���֓|��
					}else{ //���łɓ|��Ă��āA�|����Ȃ��ꍇ
						if(x[a+1]%2 == 0){
							maze[i][j+1] = 1;	//�E�֓|��
						}else{
							maze[i+1][j] = 1;	//���֓|��
						}
					}
					break;
				case 1: maze[i+1][j] = 1;			//���֓|��
					break;
				case 2: maze[i][j+1] = 1;			//�E�֓|��
					break;
			}
		}
	}


	for(i=1; i<=Y; i++){
		for(j=1; j<=X; j++){
			if(i == 1 && j == 2 ){
				printf("SS");//����
			}else if(i == 31 && j == 30){
				printf("EE");//�o��
			}else{
				if(maze[i][j] == 1){
					printf("��");
				}else{
					printf("�@");
				}
				if(j == X){
					printf("\n");
				}
			}
		}
	}

	return 0;
}

