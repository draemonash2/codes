#include <stdio.h>

#define test_stub(arg1, arg2)	\
		{						\
			(*arg1) = 0xBEEF;	\
			(*arg2) = 0xDEAD;	\
		}

int main(void)
{
	int	intVal1	= 0x1;
	int	intVal2	= 0x2;
	
	test_stub(&intVal1, &intVal2);

	printf(" My favorite food is %x !\n",	intVal1);
	printf(" My hate is %x !\n",			intVal2);

	return (0);
}

