#include<stdio.h>
#include<string.h>
#include<stdlib.h>

typedef struct {
	char name[10];
	int point;
}ST_TEST;

int main()
{
	ST_TEST stTestData[5];
	int i;
	int input;
	
	strcpy ( stTestData[0].name , "endo" );
	stTestData[0].point = 100;
	strcpy ( stTestData[1].name , "matsuda" );
	stTestData[1].point = 10;
	strcpy ( stTestData[2].name , "horibe" );
	stTestData[2].point = 80;
	strcpy ( stTestData[3].name , "tamura" );
	stTestData[3].point = 20;
	strcpy ( stTestData[4].name , "okamoto" );
	stTestData[4].point = 40;
	
	//qsort( stTestData, 5, sizeof(stTestData), comp );
	
	input = getch();
	putch(input);
	
	printf("\n0x%x\n",input);
	
	/*
	for(i=0; i<=4; i++){
		printf("name  : ");
		scanf("%s",stTestData[i].name);
		printf("point : ");
		scanf("%d",&stTestData[i].point);
	}
	*/
	
	for(i=0; i<=4; i++){
		printf("name  : %s\n",stTestData[i].name);
		printf("point : %d\n\n",stTestData[i].point);
	}
	
	
	return 0;
}
					