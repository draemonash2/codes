#include <stdio.h>
#include "./Try03Dir/Try03.h"

int main(void)
{
    int     intInputVal     = 10;
    int     intOutputVal    = 0;
    
    behavior(
                ( FUNC_EXE_DOINCREMENT | FUNC_EXE_DODECREMENT | FUNC_EXE_DODOUBLE | FUNC_EXE_DOTRIPLE),
                intInputVal,
                &intOutputVal
            );

    printf("OutputVal is %d !", intOutputVal);

    return (0);
}

void behavior( int intFuncFlag, int intInputVal, int *intOutputVal)
{
    int     intLoopCnt  =   0;
    int     intBitMsk   =   0x01;
    
    debugPrint(intFuncFlag);
    
    for (intLoopCnt = 0; intLoopCnt < FUNC_NUM_MAX; intLoopCnt++) {
        if ((intFuncFlag & intBitMsk) == 1) {
            *intOutputVal   += fp[intLoopCnt](intInputVal);
        }
        intFuncFlag = intFuncFlag >> 1;
    }
    
    return;
}

int doIncrement(int arg01)
{
    return (arg01 + 1);
}

int doDecrement(int arg01)
{
    return (arg01 - 1);
}

int doDouble(int arg01)
{
    return (arg01 * 2);
}

int doTriple(int arg01)
{
    return (arg01 * 3);
}

void debugPrint(int arg01)
{
    printf("debug print \"%d\"\n", arg01);
    return;
}
