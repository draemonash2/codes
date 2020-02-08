Attribute VB_Name = "Mng_Math"
Option Explicit

' math manage library v1.00

' ==================================================================
' = �T�v    �]�艉�Z
' =         Mod ���Z�q�� 2,147,483,647 ���傫�������̓I�[�o�[�t���[����B
' =         �{�֐��͏�L�ȏ�̐��l���������Ƃ��ł���B
' = ����    cNum1   Currency    [in]    �l1
' = ����    cNum2   Currency    [in]    �l2
' = �ߒl            Currency            ���Z����
' = �o��    �Ȃ�
' = �ˑ�    �Ȃ�
' = ����    Mng_Math.bas
' ==================================================================
Public Function ModEx( _
    ByVal cNum1 As Currency, _
    ByVal cNum2 As Currency _
) As Currency
    ModEx = CDec(cNum1) - Fix(CDec(cNum1) / CDec(cNum2)) * CDec(cNum2)
End Function
    Private Sub Test_ModEx()
        Debug.Print "*** test start! ***"
        Debug.Print ModEx(12, 2)             '0
        Debug.Print ModEx(12, 3)             '0
        Debug.Print ModEx(12, 5)             '2
        Debug.Print ModEx(2147483647, 5)     '2
        Debug.Print ModEx(2147483648#, 5)    '3
        Debug.Print ModEx(2147483649#, 5)    '4
        Debug.Print ModEx(-2147483647, 5)    '-2
        Debug.Print ModEx(5, 2147483648#)    '5
        Debug.Print ModEx(5, 2147483649#)    '5
        Debug.Print ModEx(0, 5)              '0
       'Debug.Print ModEx(5, 0)              '�v���O������~
        Debug.Print "*** test finished! ***"
    End Sub

