import sys
import xlsxwriter
from bsedata.bse import BSE
from time import sleep, time
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

set1 = set()
for i in codelist:
    set1.add(i)
codelist = set1

def waitFor(codelist):
    
    
    waittime = 0

    try:
        start = time()
        b.getQuote("532978")
        end = time()
        waittime = end - start
    except:
        print("\n----------Please connect internet !----------")
        sleep(15)
        sys.exit()


    waittime *= len(codelist)
    waittime = int(waittime)
    waittime = [waittime//60, waittime%60]

    print("\n----------Keep stock-today excel file closed----------")
    print("----------Stay connected to internet and wait "+str(waittime[0])+" min "+str(waittime[1])+" sec----------")
  
def mainProgram(codelist):
    quotes = []


    i = 1
    for code in codelist:
        try:
            quotes.append(b.getQuote(code))
        except:
            print("\nbse-quotation-number file me quotation ", code,
                  " me problem hai jo line ", i, " me likha hai")
            print("please wait...")
        i += 1
    quotes = sorted(quotes, key=lambda i: float(i['pChange']), reverse=True)

    kk = [["No","companyName", 'currVal', '    dayHigh', '    dayLow', '    % Change', '      change',  'time']] + [[i["companyName"], i['currentValue'],
                                            float(i['dayHigh']), float(i['dayLow']), float(i['pChange']), float(i['change']), i['updatedOn'][12:17]]for i in quotes]


    for i in range(1,len(kk)):
        kk[i] = [i]+kk[i]

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
        worksheet.conditional_format('F2:G'+ln, {
            'type': 'cell',
            'criteria': '>=',
            'value': 0,
            'format': green})

        worksheet.conditional_format('F2:G'+ln, {
            'type': 'cell',
            'criteria': '<',
            'value': 0,
            'format': red})

        worksheet.set_column('A:A', 2, None)
        worksheet.set_column('B:B', 28, None)
        worksheet.set_column('C:G', 11, None)

    print("\n----------Done !----------")
    sleep(20)
    sys.exit()


waitFor(codelist)
mainProgram(codelist)

    
