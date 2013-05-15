#include <stdio.h>

extern void test_stub(void);

#define DEBUG 1

#if DEBUG
        int global01    = 100;
extern  int global02    = 100;
#else
        int global01;
extern  int global02;
#endif

int main(void)
{
    
    test_stub();

    return (0);
}

void test_stub()
{
#if DEBUG
    /* None */
#else
    global01    = 0;
    global02    = 0;
#endif
    
    printf("global01 = %d\n", global01);
    printf("global02 = %d\n", global02);
    return;
}
