# LLC-4000 Congestion Tool
# 輻輳ログ変換ツール仕様 ver: 0.1

## Made with
[![Electron Logo](https://electronjs.org/images/electron-logo.svg)](https://electronjs.org)

LLC-4000 Enc輻輳制御時に、UDPパケット出力するログデータをWiresharkでキャプチャする。
Wiresharkでキャプチャしたデータを一旦 テキストファイル変換し、このテキストファイルを本変換ツールに与え、見やすい形式に変換する。

Capture the log data output by UDP packets using Wireshark during LLC-4000 Enc congestion control.
The data captured with Wireshark is first converted into a text file, and this text file is fed to this conversion tool to convert it into an easy-to-read format.

![image](https://github.com/kimkimhun3/LLC-4000_Congestion_tool/assets/47348954/ad26eb8c-fe10-4f86-b7bb-0d45ab568c28)

Windows PC上で動作するツールとして作成する。(青網掛：ボタン、黄網掛：コメント)
開発言語は、問わない。

![image](https://github.com/kimkimhun3/LLC-4000_Congestion_tool/assets/47348954/84704cd4-86b0-4807-9042-0f8949ea8594)

## 入力ログファイル仕様![image](https://github.com/kimkimhun3/LLC-4000_Congestion_tool/assets/47348954/cb96e640-0d48-4932-b93f-46ed0d46721c)
Input log file specifications

![image](https://github.com/kimkimhun3/LLC-4000_Congestion_tool/assets/47348954/824578c5-9531-4b16-a3cb-0ac973e0464a)

## Installation
We will need to install:

- **Node.js**
  
### Node.js

To install **Node.js**, simply open up your **terminal** and run the following command

```sh
npm install nodejs
```

To ensure that Node.js is installed correctly on the server, run the following command

```sh
node -v
```

## Quick start & Electron Fiddle

To start this project:

Alternatively, clone and run the
repository to see a minimal Electron app in action:

```sh
git clone https://github.com/kimkimhun3/LLC-4000-Tool.git
cd LLC-4000_Congestion_tool
npm install
npm start
```
- **Note point**


## Release History

- 0.1
```sh
F: Few, M:Many
evt1 to PLOST-F
evt2 to JT-F
evt4 to RTT-SDA-F
evt8 to RTT-STRD-F
evt10 to RTT-LTRD-F
evt20 to PLOST-M
evt40 to JT-M
evt80 to RTT-SDA-M
evt100 to RTT-STRD-M
evt200 to RTT-LTRD-M

検出した輻輳イベントをbit単位で示す。(16進数)
b0:パケットロスト閾値オーバー 小
b1:ジッタ閾値オーバー 小
b2:RTTバラつき閾値オーバー 小
b3:RTT短期上昇閾値オーバー 小
b4:RTT長期上昇閾値オーバー 小
b5:パケットロスト閾値オーバー 大
b6:ジッタ閾値オーバー 大
b7:RTTバラつき閾値オーバー 大
b8:RTT短期上昇閾値オーバー 大
b9:RTT長期上昇閾値オーバー 大
```
LANに配信される輻輳データで、イベント発生時にマークする輻輳グラフデータの関連を示す。
Congestion data delivered to the LAN, showing the association of congestion graph data to be marked when an event occurs.

![image](https://github.com/kimkimhun3/LLC-4000-Tool/assets/47348954/8131a7eb-539d-4731-aedf-08a3f9f5d218)





