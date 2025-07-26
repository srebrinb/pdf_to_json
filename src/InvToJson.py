from datetime import datetime
from os import name
import re
import json

bg_months = {
    'Януари': 'January',
    'Февруари': 'February',
    'Март': 'March',
    'Април': 'April',
    'Май': 'May',
    'Юни': 'June',
    'Юли': 'July',
    'Август': 'August',
    'Септември': 'September',
    'Октомври': 'October',
    'Ноември': 'November',
    'Декември': 'December',
    'януари': 'january',
    'февруари': 'february',
    'март': 'march',
    'април': 'april',
    'май': 'may',
    'юни': 'june',
    'юли': 'july',
    'август': 'august',
    'септември': 'september',
    'октомври': 'october',
    'ноември': 'november',
    'декември': 'december'
}
# Регулярен израз за намиране на нужните полета
address_pattern = r"Адрес на обекта: (.*) Кодов номер:"
code_pattern = r"Кодов номер:(\s*)"
name_pattern = r"Наименование на обекта: (.*)"
quantity_price_pattern = r"(.*?)(\d[\d\s.]+) кВтч\s*(\d+[\d\s.,]+)\s*"
total_amount_pattern = r"Общо сума (.*)"
period_pattern = r"За (.*) подпериод "
month_pattern = r"Основание: Електрическа енергия за месец (.*)"
def bg_to_en_month(date_str):
    for bg, en in bg_months.items():
        if bg in date_str:
            return date_str.replace(bg, en)
    return date_str

def fixSum(sum):
    """
    Функция за поправка на суми, премахва ненужните нули и форматира числата.
    """
    if not sum:
        return 0.0
    sum = sum.replace(" ", "")
    sum = sum.replace("\n", "")
    sum = sum.replace(",", "")
    sum = sum[::-1]
    sum = sum.lstrip("0")
    try:
        return float(sum)
    except ValueError:
        return 0.0042
def parse_detail_rows(block):
    """
    Функция за извличане на детайлни редове от блок.
    """
    rows = []
    quantities = re.findall(quantity_price_pattern, block)

    for textDetail, quantity, tmppricetotal in quantities:
        # Обработка на стойностите
        quantity_value = fixSum(quantity)
        
        match = re.search(r"(\d{5}\.\d)\s*([0-9\s.,]+)", tmppricetotal.replace(" ", "").replace(",", ""))
        if match:
            price = match.group(1)
            total = match.group(2)

        price_value = fixSum(price) 
        total_value = fixSum(total)
        
        # Изчисляване на стойностите
        total_calculated = quantity_value * price_value
        
        # Добавяне на данни за отчетния период
        row_data = {
            'text': textDetail.strip(),
            'quantity': quantity_value,
            'unit_price': price_value,
            'total': total_value,
            'calculated_total': round(total_calculated, 2),
            'match': abs(total_value - total_calculated) < 0.01
        }
        #Calculatedtotal += total_value
        rows.append(row_data)
    netService = re.search(r"Мрежови услуги ([0-9\s.,]+)", block)    
    if netService:
        netService_value = fixSum(netService.group(1))
     #   Calculatedtotal += netService_value
        rows.append({
            'text': 'Мрежови услуги',
            'quantity': 1,
            'unit_price': netService_value,
            'total': netService_value,
            'calculated_total': netService_value,
            'match': True
        })  

    return rows
# Функция за обработка на данните от фактурата
def parse_factura(content, invoices_data=[],filename=""):

    
    invoice_data =  {'details': {}, 'objects': []} # вземи последния добавен речник
    # Разделяне на съдържанието на блоков
    blocks = content.split("- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")
    if filename:
        invoice_data['filename'] = filename
    factura_data = []

    I=0
    # Парсване на всеки блок
    month_value = ""
    for block in blocks:
        block_data = {}
        I+=1
        if I == 1:
            month = re.search(month_pattern, block)
            month_value= month.group(1).strip()

            inv_rows= parse_detail_rows(block)
            #СОФИЯ 1612 Дебитно известие
            doc_type=block.split("\n")[3].replace("СОФИЯ 1612", "").strip()
            doc_number = block.split("\n")[4]
            inv_number = block.split("\n")[5]
            total_gose_amount = fixSum(re.search(r"Общо ([0-9\s.,]+)", block).group(1).strip())
            invoice_data['details']={"month": month_value,  'rows': inv_rows , 'doc_type': doc_type,
                                      'doc_number': doc_number, 'inv_number': inv_number,
                                      'total_amount': total_gose_amount}

            continue
        
        
        # Извличане на адрес, код и наименование
        address = re.search(address_pattern, block)
        
        name = re.search(name_pattern, block)

        if address and   name:
            block_data['address'] = address.group(1).strip()
            parts = name.group(1).strip().rsplit(" ", 1)
            if len(parts) == 2:
                block_data['code'] = parts[1]
                block_data['name'] = parts[0]
            else:
                block_data['code'] = name.group(1).strip()
                block_data['name'] = ""
            
        block_data['month']=month_value
        current_month = bg_to_en_month(month_value)
        current_month = datetime.strptime(current_month, "%B %Y") 
        block_data['std_month']= current_month.strftime("%Y-%m-%d")  # и
        # Извличане на отчетни периоди и стойности
        block_data['periods']= []
        # Разделяне на съдържанието на блокове
        blocks_period = block.split("Наименование Количество Ед. цена(лв.) Стойност(лв.)")    

        tmpblock_data = {}
        y= 0
        for block_period in blocks_period:
            tmpblock_data = {}
            y+=1
            if y == 1:continue
            
           
            total_amount = re.search(total_amount_pattern, block_period)
            if total_amount:
                total_amount_value = fixSum(total_amount.group(1).strip())
            else:
                total_amount_value = -0.042
            # print("Total amount:", total_amount_value, "Block:", I*1000+y)

            period_data = re.search(period_pattern, block_period)
            if period_data:
                period_data_value = period_data.group(1).strip()
            else:
                period_data_value =  "Не е посочен период"
            tmpblock_data["name"] =  period_data_value
            tmpblock_data["total_amount"] = total_amount_value
            # print("Total amount:", total_amount_value,"period_data_value",period_data_value, "Block:", I*1000+y)
            t= [1]
            rows = parse_detail_rows(block_period)

            tmpblock_data["rows"] = rows
            block_data['periods'].append(tmpblock_data)
        # Добавяне на текущия блок към основния списък
        factura_data.append(block_data)
    invoice_data['objects']=factura_data
    invoices_data["invoices"].append(invoice_data)
    
    return invoices_data

# Записване на резултата в JSON файл
def save_to_json(data, output_file):
    with open(output_file, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=4)

def main():
    input_file = "0417257708_6000020124_signed.pdf.txt"
    invoices_data = {"invoices": []}
    with open(input_file, "r", encoding="utf-8") as file:
        content = file.read()
    # Парсване на фактурата
    invoices_data = parse_factura(content, invoices_data, filename=input_file)
    
    # Записване на резултата в JSON файл
    output_file = "invoices_data.json"
    save_to_json(invoices_data, output_file)
    
    print(f"Данните са записани в {output_file}")

if __name__ == "__main__":
    main()