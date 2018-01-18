#include <stdio.h>

typedef struct {
	unsigned char	a;
	unsigned char	b;
	unsigned char	c;
	unsigned char	d;
} ST_CMD_INFO;

typedef struct {
	ST_CMD_INFO	cmd[2];
} ST_PWON_SUB_SCHDL;

typedef struct {
	ST_PWON_SUB_SCHDL	sub[3];
} ST_PWON_MAIN_SCHDL;

typedef struct {
	ST_CMD_INFO			cmd[5];
} ST_PELI_SUB2_SCHDL;

typedef struct {
	ST_PELI_SUB2_SCHDL	sub2[2];
} ST_PELI_SUB_SCHDL;

typedef struct {
	ST_PELI_SUB_SCHDL	sub[4];
} ST_PELI_MAIN_SCHDL;

typedef struct {
	ST_PWON_MAIN_SCHDL	pwon;
	ST_PELI_MAIN_SCHDL	peri;
} ST_MAIN_SCHDL;

ST_MAIN_SCHDL	main_schdl = 
{
		/* pwon */
		{
			/* sub */
			{
				{
					/* cmd */
					{
						/*	a,	b,	c,	d */
						{	0,	1,	2,	3	},
						{	4,	5,	6,	7	}
					}
				},
				{
					/* cmd */
					{
						/*	a,	b,	c,	d */
						{	8,	9,	10,	11	},
						{	12,	13,	14,	15	}
					}
				},
				{
					/* cmd */
					{
						/*	a,	b,	c,	d */
						{	16,	17,	18,	19	},
						{	20,	21,	22,	23	}
					}
				}
			}
		},
#if 0
		/* peri */
		{
			/* sub */
			{
				{
					/* sub2 */
					{
						{
							/* cmd */
							{
								/*	a,	b,	c,	d */
								{	24,	25,	26,	27	},
								{	28,	29,	30,	31	},
								{	32,	33,	34,	35	},
								{	36,	37,	38,	39	},
								{	40,	41,	42,	43	}
							}
						}
					},
					{
						{
							/* cmd */
							{
								{	44,	45,	46,	47	},
								{	48,	49,	50,	51	},
								{	52,	53,	54,	55	}
							}
						}
					}
				}
			},
			{
				{
					/* sub2 */
					{
						{
							/* cmd */
							{
								/*	a,	b,	c,	d */
								{	60,	61,	62,	63	},
								{	64,	65,	66,	67	}
							}
						}
					}
				}
			},
			{
				{
					/* sub2 */
					{
						{
							/* cmd */
							{
								/*	a,	b,	c,	d */
								{	68,	69,	70,	71	},
								{	72,	73,	74,	75	}
							}
						}
					}
				}
			}
		}
#endif
};

int main (void) {
//	int subidx;
//	int cmdidx;
//	subidx = 0;
//	cmdidx = 1;
//	printf( "%d\n", main_schdl.pwon.sub[ subidx ].cmd[ cmdidx ].a );
//	printf( "%d\n", main_schdl.pwon.sub[ subidx ].cmd[ cmdidx ].b );
//	printf( "%d\n", main_schdl.pwon.sub[ subidx ].cmd[ cmdidx ].c );
//	printf( "%d\n", main_schdl.pwon.sub[ subidx ].cmd[ cmdidx ].d );
	return 0;
}
