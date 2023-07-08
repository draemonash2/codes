// 独習C++ 2.3　参照より引用

int main()
{
    const int constant = 42;
    const int& ref_constant = constant;
    int& ref1 = constant; // エラー。変数自体がconstなので参照もconstが必要
    int& ref2 = ref_constant; // エラー。const参照からもconstを外すことはできない

    int value = 42;
    const int& creference = value;
    // エラー。たとえ元の変数にconstがなくてもconst参照からはconstを外せない
    int& ref3 = creference;

    int& ref4 = const_cast<int&>(ref_constant); // const_castすればconstを外せる
}
