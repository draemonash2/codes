#include<stdio.h>
int main()
{
	const char str[] = "abcdefcfc";
	
	printf("ŒÂ”F%d",str_chnum(str,'f'));
	
	return 0;
}

int str_chnum(const char *str, int c)
{
	int cnt = 0;
	
	do {
		if ( *str == c) {
			cnt++;
		}
	}while( *str++ != '\0');
	
	return cnt;
}

	