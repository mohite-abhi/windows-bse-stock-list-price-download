import sys
import xlsxwriter
from bsedata.bse import BSE
from time import sleep
b = BSE()

try:
    with open('bse-quotation-number.txt') as f:
        codelist = f.read().splitlines()
except:
    print("\nis folder me bse-quotation-number.txt file me problem hai")
    print("bse-quotation-number file me kuch is prakar ka data hona chahiye:")
    print("""500325
532281
500209""")
    sleep(15)
    sys.exit()

quotes = []

print("\ninternet se connect rahe aur kuch second wait kare")
print("bse-quotation-number file me die gae stock ka price download ho raha hai")
print("please wait...")
sleep(1)

try:
    b.getQuote("532978")
except:
    print("\nplease connect internet to download stock price")
    sleep(15)
    sys.exit()
    
i = 1
for code in codelist:
    try:
        quotes.append(b.getQuote(code))
    except:
        print("\nbse-quotation-number file me quotation ",code," me problem hai jo line ",i," me likha hai")
        print("please wait...")        
    i+=1
quotes = sorted(quotes, key = lambda i: i['pChange'], reverse=True)

kk = [["companyName",'currentValue', 'change', '% Change', 'time']] + [[i["companyName"], i['currentValue'], float(i['change']), float(i['pChange']), i['updatedOn'][12:17] ]for i in quotes]

with xlsxwriter.Workbook('stock-today.xlsx') as workbook:
    worksheet = workbook.add_worksheet()
    red = workbook.add_format()
    red.set_font_color('red')
    green = workbook.add_format()
    green.set_font_color('green')
    font_color = workbook.add_format()
    
    for row_num, data in enumerate(kk):
        worksheet.write_row(row_num, 0, data)

    ln = str(len(kk))
    worksheet.conditional_format('C2:D'+ln, {
            'type': 'cell',
            'criteria' : '>=',
            'value':0,
            'format':green})

    worksheet.conditional_format('C2:D'+ln, {
            'type': 'cell',
            'criteria' : '<',
            'value':0,
            'format':red})

    worksheet.set_column('A:A', 55, None)

print("\nAaj ke stock price stock-today excel file me ja chuke hain, please use copy kar le")
print("Next time program chalane par us file ka data delete ho kar naya data aa jaega")
sleep(20)
sys.exit()
