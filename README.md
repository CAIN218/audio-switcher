# AudioSwitcher

Windows用の常駐型オーディオ出力切り替えツールです。
あらかじめ登録した2つの再生デバイス（例: スピーカーとヘッドホン）を、グローバルホットキー1つで即座に切り替えられます。

- 切り替え時にはトースト通知で現在の出力先を表示
- Windows起動時の自動起動に対応
- 外部exe不要（Windows Core Audio APIを直接操作）

## ダウンロード

[Releases](../../releases) から最新の `AudioSwitcher.exe` をダウンロードしてください。インストール不要、exeを実行するだけで使えます。

## 使い方

### 1. 初回起動・設定

`AudioSwitcher.exe` をダブルクリックすると設定画面が開きます。

1. **DEVICE LAYER A** / **DEVICE LAYER B** で切り替え対象にしたい再生デバイスを2つ選択
2. **GLOBAL TRIGGER HOTKEY** の欄をクリックし、使いたいキーの組み合わせを押して切り替えショートカットを設定（初期値: `ctrl+alt+a`）
3. 「今すぐテスト切り替え」で動作確認
4. Windows起動時に自動で立ち上げたい場合は **Start on boot** をON
5. **Save & Run** で設定を保存してそのままバックグラウンド常駐開始（**Save only** は保存のみで起動しない）

### 2. 通常使用

常駐中はウィンドウは表示されません。設定したホットキーを押すたびに、デバイスA・Bの既定の再生デバイスが交互に切り替わり、画面右下にトースト通知が表示されます。

### 3. 設定を変更したい場合

常駐中のAudioSwitcherを一度終了し、`AudioSwitcher.exe` を再度起動すると設定画面が開きます（多重起動防止のため、古いプロセスは自動的に終了します）。

### 4. 終了方法

タスクマネージャーから `AudioSwitcher.exe` を終了してください。

## 設定ファイル

exeと同じフォルダに以下のファイルが作成されます。

- `app_config.txt` — デバイスA/B、ホットキー、自動起動設定
- `error.log` — エラー発生時のログ（トラブル時の調査用）

## 開発者向け: ソースからのビルド

```
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconfirm AudioSwitcher.spec
```

`dist/AudioSwitcher.exe` が生成されます。

## 動作環境

- Windows 10 / 11
