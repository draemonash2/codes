---
name: func-tree
description: Generate a list of functions and a function call tree for the source code within the directory.
---

ディレクトリ配下のソースコードの関数一覧と関数コールツリーを作成する。

作成する際、以下の制約事項を遵守すること。

- 関数コールツリーで表現される関数は、ディレクトリ配下のソースコード内で定義されている関数のみとする
- 関数コールツリーの根元は呼び元が存在しない関数とすること
- 一度登場した関数配下のツリーは、関数名末尾に「...」を付与して省略すること
- 関数コールツリーは呼び出しレベルに応じたインデント（空白4文字）で表現すること
- 関数一覧に出てきた関数は、全て関数コールツリーに登場させること

作成する関数コールツリーの例は以下の通り。

- [入力] 関数定義

    ```cpp
    void funcA(void) {
        funcB();
        funcDummy();
        funcC();
    }
    void funcB(void) {
        funcC();
        funcD();
    }
    void funcC(void) {
        class01::funcE();
    }
    void funcD(void) {
        class01::funcE();
        funcDummy();
    }
    void class01::funcE(void) {
        funcF();
    }
    void funcF(void) {
        // Do nothing
    }
    ```

- [出力]
    - 関数一覧

        ```txt
        funcA
        funcB
        funcC
        funcD
        class01::funcE
        funcF
        ```

    - 関数コールツリー

        ```txt
        funcA
            funcB
                funcC
                    class01::funcE
                        funcF
                funcD
                    class01::funcE...
            funcC...
        ```

