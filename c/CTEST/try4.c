#include<stdio.h>

int getNumber(void);

int main()
{
	int idx = 0;
	int i;
	int j;
	int ret;
	int arr[10] = {0,0,0,0,0,0,0,0,0,0};
	char ast[12][13] = {{'1','0','|',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '},
						{' ',' ','|',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '},
						{' ',' ','|',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '},
						{' ',' ','|',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '},
						{' ',' ','|',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '},
						{' ','5','|',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '},
						{' ',' ','|',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '},
						{' ',' ','|',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '},
						{' ',' ','|',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '},
						{' ','1','|',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '},
						{'-','-','+','-','-','-','-','-','-','-','-','-','-'},
						{' ',' ','|',' ',' ',' ',' ',' ',' ',' ',' ',' ',' '}};
	
	i = 0;
	do{
		ret = getNumber();
		if( ret == -1 ) {
			break;
		} else {
			arr[idx] = ret;
			idx++;
		}
	} while ( idx < 10 );
	
	for(i=0; i < 10; i++) {
		printf("%d\n",arr[i]);
	}
	
	idx = 0;
	j = 3;
	while(j<13){
		i = 9;
		while (arr[idx] != 0){
			ast[i][j] = '*';
			arr[idx]--;
			i--;
		}
		j++;
		idx++;
	}
	
	for( i=0; i<12; i++){
		for (j=0; j<13; j++){
			printf("%c",ast[i][j]);
		}
		printf("\n");
	}
	
	return 0;
}

int getNumber()
{
	int strS;
	int strM;
	int num = 0;
	int kCnt = 0;
	int fin_flag = 0;
	int ret_flag = 0;
	int num_flag = 0;
	
	printf("input = ");
	do{
		strM = getch();
		putch(strM);
		
		if(strM == 0x0d ){
			printf("\n");
			if( kCnt == 0 ){
				ret_flag = 1;
				fin_flag = 1;
			}else{
				if(num_flag == 1 ){
					printf("<<ERROR!>>\n");
					printf("input =");
					fin_flag = 0;
					num_flag = 0;
					num = 0;
					kCnt = 0;
				}else{
					if( num >= 0 && num <= 10 ){
						ret_flag = 0;
						fin_flag = 1;
					}else{
						printf("<<ERROR!>>\n");
						printf("input = ");
						fin_flag = 0;
						num_flag = 0;
						num = 0;
						kCnt = 0;
					}
				}
			}
		}else{
			if( strM >=0x30 && strM <=0x39 ){
				strS = strM - 0x30;
				num = num * 10 + strS;
				kCnt++;
			}else{
				num_flag = 1;
			}
		}
	}while( fin_flag == 0 );
	
	if ( ret_flag == 0 ) {
		return num;
	} else {
		return -1;
	}
	
	return 0;
}