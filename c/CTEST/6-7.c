#include<stdio.h>
int main()
{
	const int vc[] = {-88,30,3,1,-9,6,6,6,6,6,6,6,6,6,6,6,6};
	int n = sizeof(vc);
	int ret;
	
	ret = min_of(vc, n);
	
	printf("Å¬’lF%d",ret);
	
	return 0;
}

int min_of(const int *vc,int no)
{
	int min = vc[0];
	int i;
	
	for (i=0; i<no; i++){
		if ( vc[i] < min ){
			min = vc[i];
		}
	}
	
	return min;
}
	
