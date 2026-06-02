# SoundSwitcher — プロジェクトメモ

Windows 11 の再生／録音オーディオデバイスをトレイ常駐ウィンドウから切り替える WPF アプリ。

## 出力先（重要）

- **実行ファイルの出力先: `C:\prg_exe\SoundSwitcher`**
- 配布物のビルド: リポジトリの `SoundSwitcher` ディレクトリで

  ```powershell
  powershell -ExecutionPolicy Bypass -File build-dist.ps1
  ```

  既定で `C:\prg_exe\SoundSwitcher` に出力する。別の場所に出したい場合は
  `-OutDir <path>` を指定する。
- 出力構成:

  ```
  C:\prg_exe\SoundSwitcher\
    SoundSwitcher.exe          起動ランチャー（ランタイム確認・導入 → 本体起動）
    app\SoundSwitcherApp.exe   本体アプリ（フレームワーク依存）
    README.md                  利用者向けドキュメント
  ```
  ※ ランチャーと本体でファイル名を分けている（混同防止）。本体は
     csproj の `AssemblyName=SoundSwitcherApp`、ランチャーは `AssemblyName=SoundSwitcher`。

## 配布方式

- 本体は **フレームワーク依存 (FD)**。`net8.0-windows`（WinRT 投影は未使用なので
  バージョン無し TFM にしてサイズ削減）。.NET 8 デスクトップ ランタイムが必要。
- 入口の **ランチャー** (`Launcher/`) は自己完結＋トリムで、ランタイム非依存。
  起動時に `%ProgramFiles%\dotnet\shared\Microsoft.WindowsDesktop.App\8.*` を確認し、
  無ければダイアログ表示 → `winget install Microsoft.DotNet.DesktopRuntime.8`
  （winget 不可なら公式 DL ページ）→ 本体 `app\SoundSwitcherApp.exe` を起動して終了。
- 自己完結・単一ファイル方式（旧）は約 179MB・初回起動が遅かったため廃止。

## 主要構成

- `MainWindow.xaml(.cs)` — UI、ウィンドウのドラッグ／リサイズ／位置記憶、設定ボタン、
  ワーキングセット圧縮。
- `Services/AudioDeviceService.cs` — NAudio + Core Audio (MMDevice/IPolicyConfig) で
  デバイス列挙・既定切り替え・フォームファクター取得。
- `Services/DeviceBattery.cs` — SetupAPI + `DEVPKEY_Bluetooth_Battery` でバッテリー取得。
- `Services/TrayIconService.cs` — 通知領域アイコン（WinForms NotifyIcon）。終了時に
  確実に削除（冪等 Dispose ＋ シャットダウン前削除 ＋ Application.Exit フォールバック）。
- `Models/AudioDevice.cs` — デバイスモデル。`FormFactor` → `IconGlyph`（Segoe MDL2 Assets）。
- `app.ico` — アプリ／トレイ／ランチャーのアイコン。

## ビルド時の注意

- `UseWPF` と `UseWindowsForms` 併用のため、`System.Windows.Forms` と `System.Drawing`
  の暗黙 using を csproj で除去している（`Application`/`Color`/`Point` 等の曖昧回避）。
  WinForms/Drawing は `TrayIconService` で明示 using＋`WinForms.` エイリアスのみ使用。
- ランチャーのソース (`Launcher/**`) は本体 csproj の glob から除外している。
- メモリ対策で `RenderOptions.ProcessRenderMode = SoftwareOnly`（GPU ドライバを
  読み込ませない）。変更時はメモリ・スレッド数が増えないか確認すること。

## デバッグ実行

```powershell
dotnet run -c Debug         # 本体を直接起動（ランチャーを通さない）
```
