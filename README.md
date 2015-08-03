# excel_test

## 準備

```
$ bundle install
$ bundle exec rake db:create
$ bundle exec rake db:migrate
$ bundle exec rake db:seed
$ bundle exec rails s
```

```
localhost:3000/fruits
```

にアクセス。

## ダウンロード

四季に紐づく果物をそれぞれ、`春`、`夏`、`秋`、`冬` のシートに羅列してダウンロードします。

## アップロード

`春`、`夏`、`秋`、`冬`のシートを持つブックを読み込んで各シートのA列に羅列された果物を 季節に紐付けてDBに登録します。

なお、Spreadsheetは `.xls`、rubyXLは `.xlsx` のみ扱えます。
docs下にサンプルあり。
