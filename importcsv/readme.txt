今更になりますが、vbs で CSV ファイルを読み込むクラスになります。
クオート処理は excel に準じています。

ODBC の text driver を利用すれば、従来からクオート処理に対応した CSV ファイルの読み込みはできましたが、
schema.ini を用意しないと不本意な型認識を自動でされてしまったり、256 カラム以上のカラムを処理できなかったりと、
制限が許容できない場合があったかと思います。

処理速度は powershell レベルを目指しました。
※手元の実測では、クオート処理が多数入らなければ、powershell より速かったです。

ソースを公開しますので、亜流 CSV ファイルへの対応も容易に行えるかと思います。
powershell の登場により、時代遅れ感が漂ってきた VBS ですが、これで少しでも延命ができれば幸いです。

ご意見/ご感想、御座いましたらお送りください。
他の言語での高速CSVパーサーなど、ご要望ありましたら検討してみます。

---

This is a class to read the CSV file with vbs.
Quart processing is based on excel.

If you use ODBC's text driver, you can read CSV files that correspond to quote processing,
If you do not prepare schema.ini unintentional type recognition is done automatically or you can not process 256 or more columns,
I think whether there was a case where the limitation was not acceptable.

Processing speed aimed at powershell level.
* In the actual measurement at hand, it was faster than powershell unless a lot of quart processing was entered.

Since I will publish the source, I think whether it is easy to correspond to files derived from CSV format.
The appearance of powershell brings that VBS has gone with an out-of-date feeling, but I am pleased if we can extend the life even a little.

Please send us your comments / impressions.
Such as high-speed CSV parser in other languages, I will consider it.

2018/05/01 biwaz@outlook.jp
