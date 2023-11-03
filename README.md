# Ditel Easy derivation Phase and Amplification
## 概要
`Ditel Easy derivation Phase and Amplification` (以降`DEPA`)は`Tektronix`社正のオシロスコープ`TBS 1072B-EDU`から出力されるCSVファイルをもとに, 位相差及び増幅率(ゲイン)を自動で導出するソフトウェアです.
## 導入方法
`DEPA`を導入するためには以下の手順を踏む必要があります.
1. Python3のインストール  
   [こちらのサイト](https://www.python.jp/install/windows/install.html)にアクセスし最新バージョンのPythonをインストールしてください.
2. pip3のインストール  
   下のリンクを右クリック→"名前を付けて保存を選択する"を選択し, 任意のフォルダに保存してください.
   https://bootstrap.pypa.io/get-pip.py  
   保存したターミナルに移動し, 右クリック→"ターミナルで開く"を選択してください.  
   その後次のコマンド  
   `python get-pip.py`  
   を実行し, 処理が終わったらターミナルを閉じてください.
3. 必要なライブラリの導入  
   Windowsキー+Rを押し, `cmd`と入力してOKを選択してください.  
   その後, 以下のコマンド  
   `pip3 install openpyxl`  
   `pip3 install pandas`  
   `pip3 install xlwings`  
   を実行し, 必要なライブラリをインストールしてください.
4. DEPA本体の準備  
   [DEPAの配布場所](https://github.com/Ditel252/Ditel_DEPA)に移動して, `Code`→`Download ZIP`を選択し, zipファイルをダウンロードし, ダウンロードが終わったらzipファイルを展開してください.

## 参考文献
- https://www.tek.com/ja/datasheet/digital-storage-oscilloscope-0
- https://www.python.jp/install/windows/install.html
- https://qiita.com/suzuki_y/items/3261ffa9b67410803443