#include<stdio.h>
typedef struct CD{
	char title[10];
	char artist[10];
	char type[10];
	long sales;
	int favorite;
}CD;

int main()
{
	CD favoriteCD[3];
	int i;
	
	printf("���Ȃ��̍D����CD�������Ă��������B\n");
	for (i=0; i<3; i++) {
		printf("%d���ڂ̃^�C�g���F",i+1);
		scanf("%s",favoriteCD[i].title);
	}
	
	printf("���Ȃ��̍D����CD�́c\n");
	
	for (i=0; i<3; i++) {
		printf("%d���ڂ̃^�C�g����%s�ł�\n",i+1,favoriteCD[i].title);
	}
		
	return 0;
}


	