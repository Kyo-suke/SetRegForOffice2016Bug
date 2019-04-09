# SetRegForOffice2016Bug
## 概要
Microsoft Office 2016のAccessにおいて、ネットワーク上に配置したデータベースが頻繁に破損してしまう問題における一時的な対処を行います。  

Accessを利用する多数の人間は必ずしも技術的なスキルを十分に持ち合わせている訳では無い為、  
レジストリ処理を簡便化するバッチを作成しました。  

レジストリ変更を行う為、マシン動作に影響がある恐れがあります。  
自己責任で利用して下さい。

この問題については現在も以下のコミュニティスレッドで議論が行われている模様で、Microsoftの修正は行われていない様です。  
このレジストリ修正は確実に効果があるとは断言されている訳では無く、本質的な対処にならない可能性があります。  

https://answers.microsoft.com/en-us/msoffice/forum/all/access-database-is-getting-corrupt-again-and-again/d3fcc0a2-7d35-4a09-9269-c5d93ad0031d?page=15

### 2019/04/09
以下の記事に記述されたレジストリ操作も実施するように追加。

https://support.office.com/ja-jp/article/access-%E3%81%A7%E3%83%87%E3%83%BC%E3%82%BF%E3%83%99%E3%83%BC%E3%82%B9%E3%81%8C-%E7%9F%9B%E7%9B%BE%E3%81%8C%E3%81%82%E3%82%8B%E7%8A%B6%E6%85%8B-%E3%81%AB%E3%81%82%E3%82%8B%E3%81%A8%E5%A0%B1%E5%91%8A%E3%81%95%E3%82%8C%E3%82%8B-7ec975da-f7a9-4414-a306-d3a7c422dc1d

## 使い方
管理者権限で起動したコマンドプロンプトからbatを実行してください。  

```console
$ SetRegForOffice2016Bug.bat
```

その他詳しい使い方は「/?」オプションで確認できます。  

```console
$ SetRegForOffice2016Bug.bat /?
```
