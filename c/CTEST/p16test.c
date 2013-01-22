
#include <stdio.h>
#include <stdlib.h>

/* プロトタイプ宣言 */
void Factorization (int value, int div);




int main (void)
{
	int input;//入力数字
	

	printf("\n数字を入力して下さい。素因数分解します。> ");
	scanf("%d",&input);
	Factorization (input, 2) ;// * 1000 の素因数分解 *
	return 0;
}

void Factorization (int value, int div)
{
	if (value==0 || value==1) {
		printf ("%d\n", value);
	} else if (value % div ==0) { 
 		printf ("%d×", div);
 		Factorization (value/div, div); 
	} else {
		Factorization (value, div+1);
	}

}
