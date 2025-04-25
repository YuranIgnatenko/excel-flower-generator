# excel-flower-generator

> clone repo

```bash
git clone https://github.com/YuranIgnatenko/excel-flower-generator
```

> pips

```bash
cd excel-flower-generator
pip install -r requirements.txt
```

---

> launch

```bash
python main.py
```

> using config file

```bash
# task.ini

# range title and insertinting to every row
[TaskInsertHeaders]
titles = Букет роз / свежие цветы
	Розы свежие букеты по 25 51 151 201
	Букет роз 25 31 51 101 с доставкой
	Свежие розы букеты 31 51 71 101 151 201 с доставкой
	Розы букеты / Розы всех цветов с доставкой
	Тюльпаны в букетах / цветы
	Цветы тюльпаны / букеты тюльпаны
	Свежие тюльпаны / цветы с доставкой
	Тюльпаны / букет из тюльпанов / доставка
	Тюльпаны всех цветов / доставка / цветы

# paste this text to every rows
# morph only one params 
# 'ARTICLE_CODE' - replaced random (111111..999999)
# result view: ARTICLE982671
[TaskInsertText]
Text = Только СВЕЖИЕ цветы!
	Оплата ПРИ ПРОЛУЧЕНИИ!
	БЕСПЛАТНАЯ доставка !
	Вежливые и пунктуальные курьеры!
	Ваше внимание любимому человеку - ЛУЧШИЙ ПОДАРОК! Мы поможем!
	Почему выбирают нас:
	БЕСПЛАТНОЕ ОФОРМЛЕНИЕ БУКЕТА НА ВАШ ВКУС
	ОТКРЫТКА В ПОДАРОК
	ФОТО БУКЕТА ПЕРЕД ОТПРАВКОЙ
	ГАРАНТИЯ КАЧЕСТВА: ЕСЛИ ЦВЕТЫ НЕ ПРОСТОЯТ МИНИМУМ 5 ДНЕЙ -
	ПОМЕНЯЕМ БУКЕТ!!!
	СРАВНИВАЕМ ЦЕНЫ С ДРУГИМИ ПРОДАВЦАМИ (у нас дешевле всех )
	Так же наши флористы могут собрать букет любой сложности, в любом
	оформлении - просто сообщите о своих пожеланиях при заказе менеджеру, и МЫ
	СДЕЛАЕМ КРАСИВО!
	Артикул позиции - ARTICLE_CODE
	Возможно, Вы искали: свежие цветы оптом, свежие тюльпаны, желтые, красные,
	розовые, фиолетовые
	☎️
# paste this text to every rows
[TaskInsertCompanyInfo]
email = 123456@mail.ru
phone = 80123456789
id_package = 2

# paste this text to every rows
# 'now' or "12/12/12 12:12" 
[TaskInsertDateTime]
value = now

# paste this text to every rows
[TaskInsertPrice]
value = от 100 рублей

# paste this text to every rows
[TaskInsertPrecet]
name = цветы
city = Рязань
```

> getting xlsx-table output

```bash
output.xlsx
```
