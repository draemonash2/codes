#include<stdio.h>
int main()
{
	int in;
	int i;
	
	printf("®”‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢F");
	scanf("%d",&in);
	
	for (i = 1; i<=in; i++) {
		printf("%d",i%10);
	}
	return 0;
}
