import datetime as dt
from docxtpl import DocxTemplate

number = 0

while True:
    template = DocxTemplate("template_2.docx")

    number += 1

    seller = input("Введите имя продавца: ")

    goods = dict()
    while True:
        title = input("Введите название товара: ")
        price = int(input("Введите цену: "))
        goods[title] = price

        one_more_item = input("Добавить ещё один товар в чек? y/n: ")
        if one_more_item == "n":
            break

    is_discount = bool(int(input("Есть ли скидка? 1/0: ")))

    total = sum(goods.values()) * 0.95 if is_discount else sum(goods.values())

    context = {
        "number": number,
        "seller": seller,
        "date": dt.date.today(),
        "time": dt.datetime.now().strftime("%H:%M"),
        "goods": goods,
        "discount": is_discount,
        "total": total

    }
    template.render(context)
    template.save(str(number) + ".docx")

    one_more_receipt = input("Оформить ещё один чек? y/n: ")
    if one_more_receipt == "n":
        break
