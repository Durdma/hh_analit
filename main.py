from collections import defaultdict

import openpyxl as xl
import matplotlib.pyplot as plt


def task_one(sheet):
    """
    Вычислите общую выручку за июль 2021 по тем сделкам, приход денежных
    средств которых не просрочен.
    """

    l_bound: int = None
    r_bound: int = None

    for i in range(3, 732):
        if sheet[f"C{i}"].value == "Июль 2021":
            l_bound = i+1
        elif sheet[f"C{i}"].value == "Август 2021":
            r_bound = i-1

    value_status_dict = defaultdict(float)

    for cells_val, cells_status in zip(sheet[f"B{l_bound}:B{r_bound}"], sheet[f"C{l_bound}:C{r_bound}"]):
        for val, status in zip(cells_val, cells_status):
            value_status_dict[status.value] += float(val.value)

    print("Задание 1")
    print(f"Общая выручка за июль 2021 года по сделкам, приход денежных средств которых не просрочен, равен {value_status_dict.get('ОПЛАЧЕНО'):.2f} у.е")
    print("")

def task_two(sheet):
    """
    Как изменялась выручка компании за рассматриваемый период?
    Проиллюстрируйте графиком.
    Здесь подразумеваются сделки с любым статусом оплаты, так как в задании на этот счет нет никакой информации.
    """

    months = ("Май 2021", "Июнь 2021", "Июль 2021", "Август 2021", "Сентябрь 2021", "Октябрь 2021")

    month_profit = defaultdict(float)
    
    i = 2

    month = None

    while i < 732:
        if sheet[f"C{i}"].value in months:
            month = sheet[f"C{i}"].value
            i += 1
            continue
        
        month_profit[month] += float(sheet[f"B{i}"].value)
        
        i += 1

    print("Задание 2")
    print("Выручка кампании за каждый месяц отчетного периода")
    for key in list(month_profit.keys()):
        print(f"{key} {month_profit[key]:.2f} у.е")

    plt.plot(months, list(month_profit.values()))
    plt.title('График изменения выручки за отчетный период')
    plt.ylabel("Выручка у.е")
    plt.xlabel("Месяц")
    plt.ticklabel_format(style='plain', useOffset=False, axis='y')
    plt.grid(True)
    plt.xticks(rotation=90)
    plt.show()

    print("")
         

def task_three(sheet):
    """
    Кто из менеджеров привлек для компании больше всего денежных средств в
    сентябре 2021?
    Здесь подразумеваются сделки с любым статусом оплаты, так как в задании на этот счет нет никакой информации.
    """
    manager_profit = defaultdict(float)

    i = 3
    fl = False

    while sheet[f"C{i}"].value != "Октябрь 2021":
        if fl:
            manager_profit[sheet[f"D{i}"].value] += float(sheet[f"B{i}"].value)
        elif sheet[f"C{i}"].value == "Сентябрь 2021":
            fl = True

        i += 1

    sorted_managers = sorted(manager_profit.items(), key=lambda x: x[1], reverse=True)

    print("Задание 3")
    print(f"За сентярь 2021 года больше всех денежных средств привлек {sorted_managers[0][0]}. \nОн привлек {sorted_managers[0][1]:.2f} y.e")
    print("")


def task_four(sheet):
    """
    Какой тип сделок (новая/текущая) был преобладающим в октябре 2021?
    """
    statuses = defaultdict(int)

    i = 3
    fl = False
    
    while sheet[f"C{i}"].value is not None:
        if fl:
            if sheet[f"E{i}"].value is None:
                i += 1
                continue
            statuses[sheet[f"E{i}"].value] += 1
        elif sheet[f"C{i}"].value == "Октябрь 2021":
            fl = True

        i += 1

    sorted_statuses = sorted(statuses.items(), key=lambda x: x[1], reverse=True)

    print("Задание 4")
    print(f"В октябре 2021 года преобладающим типом сделок был {sorted_statuses[0][0]}.\nИх было {sorted_statuses[0][1]}")
    print("")


def task_five(sheet):
    """
    Сколько оригиналов договора по майским сделкам было получено в июне 2021?
    """
    count_origs = 0

    i = 2
    fl = False

    while sheet[f"C{i}"].value != "Июнь 2021":
        if fl:
            if (sheet[f"H{i}"].value is None) or isinstance(sheet[f"H{i}"].value, str):
                i += 1
                continue
            if sheet[f"H{i}"].value.year == 2021 and sheet[f"H{i}"].value.month == 6:
                count_origs += 1 
        elif sheet[f"C{i}"].value == "Май 2021":
            fl = True

        i += 1

    print("Задание 5")
    print(f"{count_origs} оригиналов договоров по майским сделкам было получено в июне 2021 года.")
    print("")


def last_task(sheet):
    """
        За новые сделки менеджер получает 7 % от суммы, при условии, что статус
    оплаты «ОПЛАЧЕНО», а также имеется оригинал подписанного договора с
    клиентом (в рассматриваемом месяце).
        За текущие сделки менеджер получает 5 % от суммы, если она больше 10 тыс.,
    и 3 % от суммы, если меньше. При этом статус оплаты может быть любым,
    кроме «ПРОСРОЧЕНО», а также необходимо наличие оригинала подписанного
    договора с клиентом (в рассматриваемом месяце).
        Бонусы по сделкам, оригиналы для которых приходят позже рассматриваемого
    месяца, считаются остатком на следующий период, который выплачивается по мере
    прихода оригиналов. Вычислите остаток каждого из менеджеров на 01.07.2021.
    """
    managers_remainder = defaultdict(float)

    i = 2

    while sheet[f"C{i}"].value != "Июль 2021":
        if sheet[f"H{i}"].value is None:
            i += 1
            continue

        if sheet[f"C{i}"].value == "ВНУТРЕННИЙ":
            i += 1 
            continue

        if sheet[f"H{i}"].value.month > 7:
            if sheet[f"E{i}"].value == "новая" and sheet[f"C{i}"].value == "ОПЛАЧЕНО":
                managers_remainder[sheet[f"D{i}"].value] += float(sheet[f"B{i}"].value) * 0.07
            elif sheet[f"E{i}"].value == "текущая" and sheet[f"C{i}"].value != "ПРОСРОЧЕНО":
                if float(sheet[f"B{i}"].value) > 10000:
                     managers_remainder[sheet[f"D{i}"].value] += float(sheet[f"B{i}"].value) * 0.05
                else:
                    managers_remainder[sheet[f"D{i}"].value] += float(sheet[f"B{i}"].value) * 0.03
        
        i += 1

    print("Задание про остаток от сделок")
    print("Под остатком от сделок на 01.07.2021 имею в виду те сделки, оригиналы документов которых")
    print("придут после июля месяца, так как 01.07 уже считается, что начался рассчетный период")
    print("за июль.")
    print("")
    print("Остаток каждого менеджера на 01.07.2021")
    for val in list(managers_remainder.items()):
        print(f"Менеджер: {val[0]} Остаток: {val[1]:.2f} y.e")
    print("")
        



def main():
    address = "data.xlsx" # файл data.xlsx должен быть размещен в корневой директории программы, либо с помощью lib os указать путь

    table: xl.Workbook  = xl.load_workbook(address, data_only=True)
    sheet = table["Лист1"]
    
    task_one(sheet)
    task_two(sheet)
    task_three(sheet)
    task_four(sheet)
    task_five(sheet)

    last_task(sheet)


if __name__ == "__main__":
    main()