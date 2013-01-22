#include<stdio.h>

unsigned rrotate(unsigned x, int n);
unsigned lrotate(unsigned x, int n);

int main()
{
	int x = 0x88;
	int n = 3;
	
	printf("%x %x %x %x\n", x, n, rrotate(x,n), lrotate(x,n));
	
	return 0;
}

unsigned rrotate(unsigned x, int n)
{
	int x1;
	int x2;
	
	x1 = (x << n) & 0xff;
	x2 = (x >> (8-n)) & 0xff;
	
	return (x1 | x2);
}


unsigned lrotate(unsigned x, int n)
{
	int x1;
	int x2;
	
	x1 = (x << (8-n)) & 0xff;
	x2 = (x >> n) & 0xff;
	
	return (x1 | x2);
}