#include<stdio.h>
#include<time.h>
#include<stdlib.h>

//道か壁かの判定用
#define ROAD    0      //道
#define WALL    1      //壁

#define FIELD   13

//穴を掘るための判定用:ランダムでどこに進むか決める
#define UP  0  //上
#define RIGHT   1    //右
#define UNDER   2    //下
#define LEFT    3      //左

int main(void)
{
    int i, j, k;
    int field[FIELD][FIELD];
        
    
    for(i = 0; i < FIELD; i++)
    {
        for(j = 0; j < FIELD; j++)
        {
            //初期化。
            field[i][j] = ROAD; //全部を道にします
            
            //最外周を壁にします
            field[0][j] = WALL;   //最左
            field[FIELD-1][j] = WALL;   //最右
            field[i][0] = WALL;   //最上
            field[i][FIELD-1] = WALL;   //最下
            
            //（偶数,偶数）の座標を壁にします
            if(i % 2 == 0 && j % 2 == 0){
                field[i][j] = WALL;
            }
        }
    }
    
//------------------------------------------------棒倒し処理
    srand((unsigned)time(NULL));
    for(i = 2; i < FIELD-2; i+=2)
    {
        for(j = 2; j < FIELD-2; j+=2)
        {
            if(i == 2)
            {
                k = rand() % 4;
                switch(k)
                {
                    case UP:
                        field[i - 1][j] = WALL;
                        break;
                    case RIGHT:
                        field[i][j + 1] = WALL;
                        break;
                    case UNDER:
                        field[i + 1][j] = WALL;
                        break;
                    default:
                        field[i][j - 1] = WALL;
                        break;
                }
            }
            else{
                k = rand() % 3;
                switch(k)
                {
                    case UP:
                        field[i][j + 1] = WALL;
                        break;
                    case RIGHT:
                        field[i + 1][j] = WALL;
                        break;
                    default:
                        field[i][j - 1] = WALL;
                        break;
                }
            }
        }
    }
    
//---------------------------------------------------
    
    //表示
    for(i = 0; i < FIELD; i++)
    {
        for(j = 0; j < FIELD; j++)
        {
            if(field[i][j] == WALL)
            {
                printf("■");
            }
            if(field[i][j] == ROAD)
            {
                printf("  ");
            }
        }
        printf("\n");
    }
    return 0;
}


