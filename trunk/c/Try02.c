#include <stdio.h>

typedef struct _T_STRUCT {
	short	a;
	short	b;
	int		c;
} T_STRUCT;

int main( void )
{
	char		a[10]	= {1, 2, 3, 4, 5, 6, 7, 8, 9, 10};
	T_STRUCT	*pkInfo;
	char		*p;
	
	p		=			  &a[0] + 1;
	pkInfo	= (T_STRUCT *)&a[0];
	
	printf("&a[0]		= %p\n", &a[0]);
	printf("pkInfo		= %p\n", pkInfo);
	printf("&pkInfo->a	= %p\n", &pkInfo->a);
	printf("&pkInfo->b	= %p\n", &pkInfo->b);
	printf("&pkInfo->c	= %p\n", &pkInfo->c);
	printf("p			= %p\n", p);
	printf("a[0]		= %d\n", a[0]);
	printf("pkInfo->a	= %d\n", pkInfo->a);
	printf("pkInfo->b	= %d\n", pkInfo->b);
	printf("pkInfo->c	= %d\n", pkInfo->c);
	printf("p*			= %d\n", *p);
	
	return(0);
}
