
#include <stdio.h>
#include <stdlib.h>

/* �v���g�^�C�v�錾 */
void Factorization (int value, int div);




int main (void)
{
	int input;//���͐���
	

	printf("\n��������͂��ĉ������B�f�����������܂��B> ");
	scanf("%d",&input);
	Factorization (input, 2) ;// * 1000 �̑f�������� *
	return 0;
}

void Factorization (int value, int div)
{
	if (value==0 || value==1) {
		printf ("%d\n", value);
	} else if (value % div ==0) { 
 		printf ("%d�~", div);
 		Factorization (value/div, div); 
	} else {
		Factorization (value, div+1);
	}

}
