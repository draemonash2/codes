#include<stdio.h>
#include<time.h>
#include<stdlib.h>

//�����ǂ��̔���p
#define ROAD    0      //��
#define WALL    1      //��

#define FIELD   13

//�����@�邽�߂̔���p:�����_���łǂ��ɐi�ނ����߂�
#define UP  0  //��
#define RIGHT   1    //�E
#define UNDER   2    //��
#define LEFT    3      //��

int main(void)
{
    int i, j, k;
    int field[FIELD][FIELD];
        
    
    for(i = 0; i < FIELD; i++)
    {
        for(j = 0; j < FIELD; j++)
        {
            //�������B
            field[i][j] = ROAD; //�S���𓹂ɂ��܂�
            
            //�ŊO����ǂɂ��܂�
            field[0][j] = WALL;   //�ō�
            field[FIELD-1][j] = WALL;   //�ŉE
            field[i][0] = WALL;   //�ŏ�
            field[i][FIELD-1] = WALL;   //�ŉ�
            
            //�i����,�����j�̍��W��ǂɂ��܂�
            if(i % 2 == 0 && j % 2 == 0){
                field[i][j] = WALL;
            }
        }
    }
    
//------------------------------------------------�_�|������
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
    
    //�\��
    for(i = 0; i < FIELD; i++)
    {
        for(j = 0; j < FIELD; j++)
        {
            if(field[i][j] == WALL)
            {
                printf("��");
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


