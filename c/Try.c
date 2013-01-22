#include <stdio.h>

void max (int a, int b, int c);
void min (int a, int b, int c);
void mid (int a, int b, int c);

typedef struct _T_FUNC {
    void    (*pfStartup)();
    void    (*pfPreProc)();
    void    (*pfExecute)();
} T_FUNC;

T_FUNC funcPoint1 = {
    &max,
    &mid,
    &min
};

T_FUNC funcPoint2 = {
    &min,
    &mid,
    &max
};

typedef struct _T_FUNC_LISt {
    T_FUNC  func1;
    T_FUNC  func2;
} T_FUNC_LIST;

int main( void )
{
        
    int a   = 3;
    int b   = 1;
    int c   = 2;
    T_FUNC_LIST funclist = {
        funcPoint1,
        funcPoint2
    };

    funclist.funcPoint1.pfStartup(a, b, c);
    funclist.funcPoint1.pfPreProc(a, b, c);
    funclist.funcPoint1.pfExecute(a, b, c);
    funclist.funcPoint2.pfStartup(a, b, c);
    funclist.funcPoint2.pfPreProc(a, b, c);
    funclist.funcPoint2.pfExecute(a, b, c);
    
    return(0);
T   rt.c:3: syntax error, unexpected tIDENTIFIER, expecting keyword_do or '{' or '('
    
    
void max (int a, int b, int c);
                      ^
Try.c:3: syntax error, unexpected tIDENTIFIER, expecting keyword_do or '{' or '('
void max (int a, int b, int c);
                             ^
        if (a > c) {
            printf("max = %d\n", a);
        } else {
            printf("max = %d\n", c);
        }
    } else {
        if (b > c) {    
            printf("max = %d\n", b);
        } else {
            printf("max = %d\n", c);
        }
    }
    return;
}

void min (int a, int b, int c)
{
    if (a < b) {
        if (a < c) {
            printf("min = %d\n", a);
        } else {
            printf("min = %d\n", c);
        }
    } else {
        if (b < c) {
            printf("min = %d\n", b);
        } else {
            printf("min = %d\n", c);
        }
    }
    return;
}

void mid (int a, int b, int c)
{
    if (a < b) {
        if (a > c) {
            printf("mid = %d\n", a);
        } else {
            printf("mid = %d\n", c);
        }
    } else {
        if (b > c) {
            printf("mid = %d\n", b);
        } else {
            printf("mid = %d\n", c);
        }
    }
    return;
}
