#include <stdio.h>

void printarray(int *,int);
void sort(int *,int);
void swap(int *,int *);
int test01;
int test02;

/* program entry */
int main(int argc, char **argv) {
	test01 = 1;
	int array[4] = {4, 1, 3, 2};
	printarray(&array[0], 4);
	sort(array, 4);
	printarray(&array[0], 4);
	test01 = 0;
	return 0;
}

/* printout */
void printarray(int *array, int length) {
	int i;
	for (i = 0; i < length; i++) {
		printf("%d ", array[i]);
	}
	printf("\n");
}

/* bubble sort */
void sort(int *array, int length) {
	int i, j;
	test02 = 1;
	for (i = 0; i < length - 1; i++) {
		for (j = 0; j < length - i - 1; j++) {
			if (array[j] > array[j+1]) {
				swap(&array[j], &array[j+1]);
			}
		}
	}
	test02 = 0;
}

/* swap a <=> b */
void swap(int *a, int *b) {
	int temp;
	temp = *a;
	*a = *b;
	*b = temp;
}

