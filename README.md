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
git clone https://github.com/kimkimhun3/LLC-4000_Congestion_tool.git
cd LLC-4000_Congestion_tool
npm install
npm start
```


## Release History

- 0.1







