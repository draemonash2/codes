#include<stdio.h>
int main()
{
	int in;
	int i;
	
	printf("整数を入力してください：");
	scanf("%d",&in);
	
	for (i = 1; i<=in; i++) {
		printf("%d",i%10);
	}
	return 0;
}
