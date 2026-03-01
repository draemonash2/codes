---
name: func-header-comment
description: Describe the function header comment.
---

ヘッダファイル内の関数プロトタイプ宣言上部に「`// TODO: prototype`」がある関数に対して、関数ヘッダコメントを記載する。

記載する際、以下の制約事項を遵守すること。

- ヘッダファイル(例: `*.hpp`)内に関数ヘッダコメントを付与すること
- 関数ヘッダコメントは英語で記載すること
- Doxygen形式で記載し、タグは`@brief`, `@param`, `@return`, `@details` を記載すること
- 要素が存在しない場合は`None`と記載すること（例: `@return None`）
- 仮引数の入出力方向に合わせて`[in]`か`[out]`をつけること（例: `@param[in] (left_top) coordinates of the upper left corner`）
    - 仮引数のデータ型にconstがついている場合は`[in]`、それ以外は`[out]`として問題ない。
- 仮引数名には()をつけること（例: `@param[in] (left_top) coordinates of the upper left corner`）
- 1行は最大90文字とし、それを超える文になる場合は改行して折り返すこと
- 空白2文字のインデントを設けること
- `@details`は簡潔に記載すること
- 記載例は以下の通り。

    ```cpp
      /**
       * @brief Converts all models in the SDF world file to static models, excluding specified models.
       * @param[in] (in_world_file_path) The input path of the SDF world file to be modified.
       * @param[in] (out_world_file_path) The output path where the modified SDF world file
       *  will be saved.
       * @param[in] (exclude_model_id_list) A list of model IDs to be excluded from the
       *  static conversion.
       * @param[in] (logger) The logger used for error reporting.
       * @return True if the operation is successful, false otherwise.
       * @details
       *  This function modifies the SDF world file, converting all models to static except
       *  those specified in the `exclude_model_id_list`. The function processes both <include>
       *  and <model> elements, adding or modifying the <static> tag to ensure models are
       *  static where applicable.
       */
    ```

