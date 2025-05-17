    # coding: utf-8
import os
dirpath = os.getcwd()

def pdfgenerate(path):
    import pandas as pd 
    import numpy as np
    #https://okanemamire.net/check-stocks-dividend/
    #https://mujinzou.com/2025_day_calendar.htm
    renzoguhaitou=0
    higenpai=10
    leastincomeband=500
    leastpercentban=4
    topn=1000
    maxincomband=10000000
    junban="配当利回り"#配当利回り#時価総額#非減配年数
    from datetime import datetime
    today=datetime.today()
    date=250516#today.strftime('%Y')[2:]+today.strftime('%m')+str(int(today.strftime('%d'))-5)
    print(date)
    #ここから株価データの結合
    df1=pd.read_csv(dirpath+"fy-balance-sheet.csv")
    csvData = pd.read_csv(path,encoding="cp932", dtype=object,names=['A', 'コード', 'C', '名前', 'E', 'F', 'G', '株価','I', 'J', 'K']) 
    result = pd.merge(df1, csvData, on="コード", how='left').fillna(0)
    #ここから増配データの結合
    titletext ="addedcsvDataplus"
    ruishinData =pd.read_csv(dirpath+titletext+".csv", encoding="cp932") 
    csvData = result
    ruishinData['コード']=ruishinData['コード'].astype(str)
    csvData['コード']=csvData['コード'].astype(str)
    csvData = ruishinData.merge(csvData,how='left', on='コード')
    #ここから純利益データの結合
    csvData2 = pd.read_csv(dirpath+"fy-profit-and-loss.csv", encoding="utf-8")
    csvData['コード']=csvData['コード'].astype(str)
    csvData2['コード']=csvData2['コード'].astype(str)
    csvData = csvData.merge(csvData2,how='left', on='コード')
    #ここから優待データの結合
    yutaidir=dirpath+"fixedcsvData.csv"
    yutai = pd.read_csv(yutaidir,  encoding="cp932")
    csvData['コード']=csvData['コード'].astype(str)
    yutai['コード']=yutai['コード'].astype(str)
    csvData = csvData.merge(yutai,how='left', on='コード')



    csvData = csvData[~(csvData['非減配年数'] < higenpai)] 
    set=csvData["純利益"].astype(float)*csvData["株価"].astype(float)/csvData["EPS"].astype(float)
    set = set.fillna(0)
    csvData["時価総額"]=[int(i/100000000) for i in set]

    cleaned=[]
    for x in csvData["連続増配"].fillna(0):
        if isinstance(x, int) or isinstance(x, float):cleaned.append(float(x))
        elif isinstance(x, str) and x.isdigit():cleaned.append(float(x))
        else:cleaned.append(0)
    csvData["増配年数"]=[i for i in cleaned]
    csvData = csvData[~(csvData['増配年数'] < renzoguhaitou)] 
    csvData["配当"]=[i for i in csvData["2024"].fillna(0)]
    cleaned=[]
    for x in csvData["配当"]:
        if isinstance(x, int) or isinstance(x, float):cleaned.append(float(x))
        elif isinstance(x, str) and x.isdigit():cleaned.append(float(x))
        else:cleaned.append(0)
    set2=cleaned/csvData["株価"].astype(float)
    set2.replace([np.inf, -np.inf], 0, inplace=True)
    set2 = set2.fillna(0)
    csvData["配当利回り"]=[int(i*10000)/100 for i in set2]
    csvData[junban].to_list = pd.to_numeric(csvData[junban].to_list, errors='coerce')
    csvData.sort_values([junban],  
                        axis=0, 
                        ascending=[False],  
                        inplace=True
                        ) 

    csvData["増配年数"]=[str(int(i))+"年" for i in csvData["増配年数"]]
    csvData["連続非減配"]=[str(int(i))+"年" for i in csvData["非減配年数"].fillna(0)]

    set=csvData["株価"].astype(float)/csvData["BPS"].astype(float)
    set = set.fillna(0)
    csvData["PBR"]=[int(i*100)/100 for i in set]
    set=csvData["株価"].astype(float)/csvData["EPS"].astype(float)
    set = set.fillna(0)
    csvData["PER"]=[int(i*100)/100 for i in set]
    print(csvData.columns.tolist()) 
    print(len(csvData.columns.tolist()))
    csvData = csvData[~(csvData['時価総額'] < leastincomeband)] 
    csvData = csvData[~(csvData['時価総額'] > maxincomband)] 
    csvData = csvData[~(csvData['配当利回り'] < leastpercentban)] 
    csvData["時価総額"]=[str(int(i))+"億円" for i in csvData["時価総額"]]
    csvData["配当利回り"]=[str(i)+"％" for i in csvData["配当利回り"]]
    csvData["銘柄名"]=[str(i)[0:7]  for i in csvData["銘柄名"]]
    csvData=csvData.head(topn)
    csvData["優待"]=[str(i)[0:12]  for i in csvData["優待情報"].fillna("X")]#
    csvData["非減配"]=[str(i) for i in csvData["連続非減配"]]
    csvData["増配"]=[str(i) for i in csvData["増配年数"]]
    csvData = csvData.loc[:, ['コード','銘柄名', '株価',"配当",'時価総額','配当利回り','増配','非減配',"PBR","PER","決算月","優待"]]

    name=str(date)+junban+"順＿時価総額"+str(int(leastincomeband))+"億円以上"+str(int(maxincomband))+"億円以下"+"＿配当利回り"+str(leastpercentban)+"%以上"+"＿連続増配年数"+str(renzoguhaitou)+"年以上"+"＿連続非減配"+str(higenpai)+"年以上"+titletext+"DataList"
    name2=str(date)+junban+"順＿時価総額"+str(int(leastincomeband))+"億円以上"+str(int(maxincomband))+"億円以下"+"＿配当利回り"+str(leastpercentban)+"%以上"+"＿連続増配年数"+str(renzoguhaitou)+"年以上"+"＿連続非減配"+str(higenpai)+"年以上"
    print(name)

    csvData.to_csv(name+".csv", sep=',', encoding="cp932", index=False, header=True)

    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))
    inputdatafile=name+".csv"
    data = pd.read_csv(inputdatafile, encoding="cp932") 
    df = pd.DataFrame(data)

    data_list =[df.columns.values.tolist()] + df.values.tolist()
    #print(data_list)
    pdf_path = name+".pdf"
    pdf = SimpleDocTemplate(pdf_path, pagesize=A4)

    table = Table(data_list, colWidths=[40,80,35,35,55,55,35,35,25,30,40,130])
    #print(table) 
    style = TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'HeiseiMin-W3'),
    ])

    import re
    for row, values, in enumerate(data_list):
        for column, value in enumerate(values):
            if "％" in str(value):
                if float(re.search(r'\d+', str(value)).group())>=3:
                        style.add('BACKGROUND', (column, row), (column, row), colors.pink)
                        if float(re.search(r'\d+', str(value)).group())>=4:
                            style .add('BACKGROUND', (column, row), (column, row), colors.yellow)
                            if float(re.search(r'\d+', str(value)).group())>=5:
                                style.add('TEXTCOLOR', (column, row), (column, row), colors.white)
                                style.add('BACKGROUND', (column, row), (column, row), colors.blue)

            elif "年" in str(value):
                if re.search(r'\d+', str(value)):
                    if float(re.search(r'\d+', str(value)).group())>=10:
                            style.add('TEXTCOLOR', (column, row), (column, row), colors.white)
                            style.add('BACKGROUND', (column, row), (column, row), colors.purple)
                    if float(re.search(r'\d+', str(value)).group())>=20:
                            style.add('TEXTCOLOR', (column, row), (column, row), colors.white)
                            style.add('BACKGROUND', (column, row), (column, row), colors.black)


    table.setStyle(style)
    elements = [table]
    #pdf.build(elements)

    from reportlab.platypus import Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
    from reportlab.lib.units import cm

    styles = getSampleStyleSheet()
    yourStyle = ParagraphStyle('Title',
                            fontName="HeiseiMin-W3",
                            fontSize=12,
                            alignment=1,
                            spaceAfter=14)

    flowables = [
        Paragraph(name2, yourStyle),
        table,
        Spacer(1 * cm, 1 * cm),
        Paragraph('last page')
    ]

    def onFirstPage(canvas, document):
        canvas.drawCentredString(100, 100, '')

    pdf.build(flowables, onFirstPage=onFirstPage)

    print("DataFrame exported to PDF successfully.")





import TkEasyGUI as eg
from PIL import Image 
from pytesseract import pytesseract 

layout = [
    [eg.Multiline("https://mujinzou.com/2025_day_calendar.htm")],
    [eg.Text("ファイルを選択してください")],
    [eg.Input("C:\\Users\\Owner\\Desktop\\csvfilelist\\T250516.csv", key="-name-")],
    [eg.Button("OK"), eg.Button("キャンセル")]
]
#[eg.Multiline("ここに出力", key="-DISPLAY")],
window = eg.Window("テスト", layout=layout)
# イベントループ --- (*2)
while window.is_alive():
    # イベントと値を取得
    event, values = window.read()
    # OKボタンを押した時
    if event == "OK":
        name = values["-name-"]
        pdfgenerate(name)
        #window["-DISPLAY"].update(value = imagetotext(name))
#window.close()