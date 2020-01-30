# VBAtools

VBAを便利に使うためのライブラリをいくつか提供します。

* Iconv.cls
   * Iconvのような文字コード変換機能を提供するクラスです。
     内部的にはADODB.Streamを使って文字コードを変換しますのでMicrosoft ActiveX Data Objectsを参照設定してお使いください。
   * ADODB.Streamの文字コード指定だけでは若干使いにくいため、次のような挙動をするようにしています。
      * Shift_JISの別名として「CP932」「MS932」を使用可能にしています。
      * ADODB.Streamの挙動としてUTF-8、UTF-16LEではBOM (Byte Order Mark) が挿入されてしまうため、「UTF-8N」の場合はBOMなしUTF-8を、「UTF-16LE」の場合はBOMなしUTF-16LEを出力します。
         * そもそもUTF-16BE/LEではBOMを出力するのは[UTF-16の仕様に反する挙動](https://en.wikipedia.org/wiki/UTF-16#Byte_order_encoding_schemes)のため、ActiveXの挙動がおかしいです。
         * UTF-8でも何故かBOMが付く時と付かない時があるため、先頭2バイトないし3バイトを確認してBOMがある場合のみ除去します。
