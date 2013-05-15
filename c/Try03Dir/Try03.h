#ifndef DEF_tRY03_H
#define DEF_tRY03_H

#define FUNC_EXE_DOINCREMENT	((int)0x01 << 0)
#define FUNC_EXE_DODECREMENT	((int)0x01 << 1)
#define FUNC_EXE_DODOUBLE   	((int)0x01 << 2)
#define FUNC_EXE_DOTRIPLE   	((int)0x01 << 3)
#define FUNC_NUM_MAX		  	((int)		  4)

void	behavior(int intFuncFlag, int intInputVal, int *intOutputVal);
int		doIncrement(int arg01);
int		doDecrement(int arg01);
int		doDouble(int arg01);
int		doTriple(int arg01);
void	debugPrint(int arg01);

int		(*fp[FUNC_NUM_MAX])(int)	=	{
											doIncrement,
											doDecrement,
											doDouble,
											doTriple
										};

#endif /* DEF_tRY03_H */
