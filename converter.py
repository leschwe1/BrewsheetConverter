##############PYTHON 3.13.2
import json
import io
import csv
from openpyxl import load_workbook

def converter(batchno, brewdate,in_path, out_path, excel_layout = 'brewsheet_empty.xlsx'):
    #change this file / filepath to the Brewfather json output
    with open(in_path, 'r') as file:
        data = json.load(file)

    #############################helper functions for conversions
    def conv_plato(sg):
        return(round((-1 * 616.868) + (1111.14 * sg) - (630.272 * sg**2) + (135.997 * sg**3), 1))

    def conv_ebc(srm):
        return round(float(srm*1.97), 1)

    #############################variable collection

    ##line 1 title
    beerName = data["name"]

    #line 2 facts
    style = data['style']['name']
    og = conv_plato(data["og"])
    abv = round(data['abv'], 1)
    fg = conv_plato(data["fg"])
    color = conv_ebc(data["color"])
    ibu = data["ibu"]

    #line 5-11 ingredients list
    malts = [fermentable['name'] for fermentable in data['data']['mashFermentables']]
    while len(malts) < 7:
        malts.append(None)


    hoplist = []
    alphalist = []
    for hop in data['hops']:
        if hop['name'] not in hoplist:
            hoplist.append(hop['name'])
            alphalist.append(hop['alpha'])
    while len(hoplist) < 7:
        hoplist.append(None)
        alphalist.append(None)


    yeasts = [y['productId'] for y in data['yeasts']]
    while len(yeasts) < 7:
        yeasts.append(None)

    #line 13-20 water profile


    mashWaterAmount = data["water"]["mashWaterAmount"]
    spargeWaterAmount = data["water"]["spargeWaterAmount"]
    corrWaterAmount = 10

    #mash
    mashDilution = round(float(data['water']['dilutionAmount']/data['water']['totalAdjustments']['volume']),2)*100
    mashCaCl2 = round(float(data["water"]['mashAdjustments']["calciumChloride"]),1)
    mashCaSO4 = round(float(data["water"]['mashAdjustments']["calciumSulfate"]),1)
    mashMgSO4 = round(float(data["water"]['mashAdjustments']["magnesiumSulfate"]),1)
    mashNaCl = round(float(data["water"]['mashAdjustments']["sodiumChloride"]),1)
    mashNaHCO3 = round(float(data["water"]['mashAdjustments']['sodiumBicarbonate']),1)
    mashLA = round(float(data["water"]['mashAdjustments']['acids'][0]['amount']),1)

    #sparge
    spargeDilution = mashDilution
    spargeCaCl2 = round(float(data["water"]['spargeAdjustments']["calciumChloride"]),2)
    spargeCaSO4 = round(float(data["water"]['spargeAdjustments']["calciumSulfate"]),1)
    spargeMgSO4 = round(float(data["water"]['spargeAdjustments']["magnesiumSulfate"]),1)
    spargeNaCl = round(float(data["water"]['spargeAdjustments']["sodiumChloride"]),1)
    spargeNaHCO3 = round(float(data["water"]['spargeAdjustments']['sodiumBicarbonate']),1)
    spargeLA = round(float(data["water"]['spargeAdjustments']['acids'][0]['amount']),1)

    #corr
    corrDilution = mashDilution
    corrCaCl2 = round((spargeCaCl2/spargeWaterAmount)*corrWaterAmount,1)
    corrCaSO4 = round((spargeCaSO4/spargeWaterAmount)*corrWaterAmount,1)
    corrMgSO4 = round((spargeMgSO4/spargeWaterAmount)*corrWaterAmount,1)
    corrNaCl = round((spargeNaCl/spargeWaterAmount)*corrWaterAmount,1)
    corrNaHCO3 = round((spargeNaHCO3/spargeWaterAmount)*corrWaterAmount,1)
    corrLA = round((spargeLA/spargeWaterAmount)*corrWaterAmount,1)


    #line 22-27 mash
    mashpH = round(float(data['water']['mashPh']),2)
    mashMaltAmount = round(float(data['data']["mashFermentablesAmount"]),1)
    Gussführung = f"1:{round(mashWaterAmount/mashMaltAmount, 1)}"

    mashSteps = [{'dur': step['stepTime'], 'temp': step['stepTemp']} for step in data['mash']['steps']]
    while len(mashSteps) < 6:
        mashSteps.append({'dur': None, 'temp': None})


    #line 29-35 lauter 
    """none"""

    #line 37-46 boil
    preBoilGravity = conv_plato(data["preBoilGravity"])
    preBoilVolume = round(data['equipment']['boilSize'],1)

    postBoilGravity = og
    postBoilVolume = round(data['equipment']["postBoilKettleVol"],1)

    hopdoses = [{'name': h['name'], 'time': h['time'], 'amount': h['amount']} for h in data['hops']]
    while len(hopdoses) < 10:
        hopdoses.append({'name': None, 'time': None, 'amount': None})



    #line 48-52 cool
    whirlpoolTime = data["equipment"]["whirlpoolTime"]
    wortTemp = data['fermentation']['steps'][0]['stepTemp']


    #fermsheet
    fermYeastAmount = data["yeasts"][0]["amount"]

    fermSteps = []
    for step in data['fermentation']['steps']:
        pressureBar = round(float(step['pressure'])*0.0689476, 1) if step['pressure'] is not None else 0
        fermSteps.append({
            'name': step['name'],
            'temp': step['stepTemp'],
            'pressure': pressureBar
        })
    # Pad to 5 ferm steps
    while len(fermSteps) < 7:
        fermSteps.append({'name': None, 'temp': None, 'pressure': None})


    #############################csv generation
    #############################brewsheet
    df = [
        [None,None,None,None,None,None,None,None,None,None,None,None,None], #ok
        [beerName ,None,None,None,None,None,None,batchno,None,brewdate,None,None,"Brewer:" ,None,None], #ok
        [style,None,None,None,None,og,"°P",abv,"%",fg,"°P",color,"EBC",ibu,"IBU"], #ok
        ["ZUTATEN",None,"Prod","No",None,None,None,"alpha","Prod","No",None,None,None,"Prod","No"], #ok
        [malts[0], None, None, None,None, hoplist[0], None, alphalist[0], None, None, None, yeasts[0], None],
        [malts[1], None, None, None,None, hoplist[1], None, alphalist[1], None, None, None, yeasts[1], None],
        [malts[2], None, None, None,None, hoplist[2], None, alphalist[2], None, None, None, yeasts[2], None],
        [malts[3], None, None, None,None, hoplist[3], None, alphalist[3], None, None, None, yeasts[3], None],
        [malts[4], None, None, None,None, hoplist[4], None, alphalist[4], None, None, None, yeasts[4], None],
        [malts[5], None, None, None,None, hoplist[5], None, alphalist[5], None, None, None, yeasts[5], None],
        [malts[6], None, None, None,None, hoplist[6], None, alphalist[6], None, None, None, yeasts[6], None],
        ["WASSER",None,None,None,None,None,None,None,None,None,None,None,None,"Resp: ",None], #ok
        ["Maische [l]",None,mashWaterAmount,None,None,"Nachguss [l]",None,spargeWaterAmount,None,None,None,"Korrektur [l]",None,corrWaterAmount,None],
        ["Verdünnung [%]",None,mashDilution,None,None,"Verdünnung [%]",None,spargeDilution,None,None,None,"Verdünnung [%]",None,corrDilution,None],
        ["CaCl2 (33%) [g]",None,mashCaCl2,None,None,"CaCl2 (33%) [g]",None,spargeCaCl2,None,None,None,"CaCl2 (33%) [g]",None,corrCaCl2,None],
        ["CaSO4 [g]",None,mashCaSO4,None,None,"CaSO4 [g]",None,spargeCaSO4,None,None,None,"CaSO4 [g]",None,corrCaSO4,None],
        ["MgSO4 [g]",None,mashMgSO4,None,None,"MgSO4 [g]",None,spargeMgSO4,None,None,None,"MgSO4 [g]",None,corrMgSO4,None],
        ["NaCl [g]",None,mashNaCl,None,None,"NaCl [g]",None,spargeNaCl,None,None,None,"NaCl [g]",None,corrNaCl,None],
        ["NaHCO3 [g]",None,mashNaHCO3,None,None,"NaHCO3 [g]",None,spargeNaHCO3,None,None,None,"NaHCO3 [g]",None,corrNaHCO3,None],
        ["Lactic Acid 80% [ml]",None,mashLA,None,None,"Lactic Acid 80% [ml]",None,spargeLA,None,None,None,"Lactic Acid 80% [ml]",None,corrLA,None, None],
        ["MAISCHE",None,"act","min","soll","max",None,"°C","min","Notes",None,None,None,"Resp:" ,None],
        ["pH (10mins) [pH]",None,None,mashpH - 0.1,mashpH,mashpH+0.1,"Step1",mashSteps[0]['temp'], mashSteps[0]['dur'],None,None,None,None,None,None],
        ["Korrektur LA [ml]",None,None,0,0,50,"Step2",mashSteps[1]['temp'], mashSteps[1]['dur'],None,None,None,None,None,None],
        [None,None,None,None,None,None,"Step3",mashSteps[2]['temp'], mashSteps[2]['dur'],None,None,None,None,None,None],
        ["Wasser [l]",None,None,mashWaterAmount - 10,mashWaterAmount,160,"Step4",mashSteps[3]['temp'], mashSteps[3]['dur'],None,None,None,None,None,None],
        ["Malz [kg]",None,None,None,mashMaltAmount,None,"Step5",mashSteps[4]['temp'], mashSteps[4]['dur'],None,None,None,None,None,None],
        ["Gussführung",None,None,None,Gussführung,None,"Step6",mashSteps[5]['temp'], mashSteps[5]['dur'],None,None,None,None,None,None],
        ["LÄUTERN",None,"act","min","soll","max",None,None,None,"Notes",None,None,None,"Resp:" ,None],
        ["Vorderwürze pH",None,None,mashpH - 0.1,mashpH,mashpH + 0.1,None,None,None,None,None,None,None,None,None],
        ["Vorderwürze °P",None,None,(og * 2)-3,og * 2,(og * 2)+3,None,None,None,None,None,None,None,None],
        ["Milchsäure",None,None,0,0,50,None,None,None,None,None,None,None,None,None],
        ["Nachgusswasser",None,None,spargeWaterAmount-20,spargeWaterAmount,spargeWaterAmount+30,None,None,None,None,None,None,None,None,None],
        ["Glattwasser pH",None,None,mashpH,mashpH +0.3,mashpH +0.5,None,None,None,None,None,None,None,None,None],
        ["Glattwasser °P",None,None,3.0,5.0,7.0,None,None,None,None,None,None,None,None,None],
        [None,None,None,None,None,None,None,None,None,None,None,None,None], #ok
        ["KOCHEN",None,"act","min","soll","max","Sorte","Zeit","Menge","Notes",None,None,None,"Resp:" ,None],
        ["preBoil pH",None,None,mashpH-0.2,mashpH,mashpH+0.2,hopdoses[0]['name'],hopdoses[0]['time'],hopdoses[0]['amount'],None,None,None,None,None,None],
        ["preBoil °P",None,None,preBoilGravity-0.1,preBoilGravity,preBoilGravity+0.1,hopdoses[1]['name'],hopdoses[1]['time'],hopdoses[1]['amount'],None,None,None,None,None,None],
        ["preBoil Volume",None,None,preBoilVolume-20,preBoilVolume,preBoilVolume+20,hopdoses[2]['name'],hopdoses[2]['time'],hopdoses[2]['amount'],None,None,None,None,None,None],
        ["preBoil LA [ml]",None,None,0,0,50,hopdoses[3]['name'],hopdoses[3]['time'],hopdoses[3]['amount'],None,None,None,None,None,None],
        [None,None,None,None,None,None,hopdoses[4]['name'],hopdoses[4]['time'],hopdoses[4]['amount'],None,None,None,None,None,None],
        ["postBoil pH",None,None,mashpH-0.5,mashpH-0.4,mashpH-0.3,hopdoses[5]['name'],hopdoses[5]['time'],hopdoses[5]['amount'],None,None,None,None,None,None],
        ["postBoil °P",None,None,og-0.1,og,og+0.1,hopdoses[6]['name'],hopdoses[6]['time'],hopdoses[6]['amount'],None,None,None,None,None,None],
        ["postBoil Volume",None,None,postBoilVolume - 10,postBoilVolume,postBoilVolume + 10,hopdoses[7]['name'],hopdoses[7]['time'],hopdoses[7]['amount'],None,None,None,None,None,None],
        ["postBoil LA [ml]",None,None,0,0,50,hopdoses[8]['name'],hopdoses[8]['time'],hopdoses[8]['amount'],None,None,None,None,None,None],
        [None,None,None,None,None,None,hopdoses[9]['name'],hopdoses[9]['time'],hopdoses[9]['amount'],None,None,None,None,None,None],
        ["KÜHLEN",None,"act","min","soll","max",None,None,None,"Notes",None,None,None,"Resp:" ,None],
        ["Whirlpool [mins]",None,None,whirlpoolTime-5,whirlpoolTime,whirlpoolTime+5,None,None,None,None,None,None,None,None,None],
        ["Kühlen [mins]",None,None,15,20,25,None,None,None,None,None,None,None,None,None],
        ["Würzetemp [°C]",None,None,wortTemp-1,wortTemp,wortTemp+1,None,None,None,None,None,None,None,None,None],
        ["Stw [°P]",None,None,og-0.1,og,og+0.1,None,None,None,None,None,None,None,None,None],
        ["pH" ,None,None,mashpH-0.5,mashpH-0.4,mashpH-0.3,None,None,None,None,None,None,None,None,None]

    ]
    #############################ferm sheet
    df2 = [
        [None,None,None,None,None,None,None,None,None,None,None,None,None], #ok
        [beerName ,None,None,None,None,None,None,batchno,None,brewdate,None,None,"Brewer:" ,None,None], #ok
        [style,None,None,None,None,og,"°P",abv,"%",fg,"°P",color,"EBC",ibu,"IBU"], #ok
        ["FERMENTATION",None,"ist","min","soll","max","Schritt","Beding." ,None,"Temp","Druck","ABFÜLLUNG",None,"Resp:" ,None,None],
        ["Hefemenge",None,None,None,fermYeastAmount,None,1,fermSteps[0]['name'],None,fermSteps[0]['temp'],fermSteps[0]['pressure'],"Datum",None,None,None,None],
        ["Gen",None,None,None,None,None,2,fermSteps[1]['name'],None,fermSteps[1]['temp'],fermSteps[1]['pressure'],"Kegs 20l",None,None,None,None],
        ["Viability",None,None,None,None,None,3,fermSteps[2]['name'],None,fermSteps[2]['temp'],fermSteps[2]['pressure'],"Flaschen 0.5",None,None,None,None],
        ["Stammwürze",None,None,og-0.1,og,og+0.1,4,fermSteps[3]['name'],None,fermSteps[3]['temp'],fermSteps[3]['pressure'],"Flaschen 0.3",None,None,None,None],
        ["Restextrakt",None,None,round((fg-0.1),1),fg,round((fg+0.1),1),5,fermSteps[4]['name'],None,fermSteps[4]['temp'],fermSteps[4]['pressure'],"Direktausschank",None,None,None,None],
        ["pH" ,None,None,round((mashpH-1.5),2),round((mashpH-1.4),2),round((mashpH-1.3),2),6,fermSteps[5]['name'],None,fermSteps[5]['temp'],fermSteps[5]['pressure'],None,None,None,None,None],
        ["Alkoholgehalt",None,None,round((abv-0.1),1),abv,round((abv+0.1),1),7,fermSteps[6]['name'],None,fermSteps[6]['temp'],fermSteps[6]['pressure'],'TOTAL [l]',None,None,None,None],
        [None,"Datum","Zeit","°P","pH","Temp","soll","Druck","set","Truboff","Bemerkungen",None,None,None,None,"Initialen"], #ok
        [1 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [2 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [3 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [4 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [5 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [6 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [7 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [8 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [9 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [10 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [11 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [12 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [13 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [14 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [15 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [16 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [17 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [18 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [19 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [20 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [21 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [22 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [23 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [24 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [25 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [26 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [27 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [28 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [29 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [30 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [31 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [32 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [33 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [34 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [35 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [36 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [37 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [38 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [39 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None],
        [40 ,None,None,None,None,None,None,None,None,None,None,None,None,None,None]

    ]

    #############################csv export
    with open(f'{beerName}_brew.csv', mode='w', newline='') as file:
        writer = csv.writer(file, delimiter=';')
        for row in df:
            # Replace decimal point with comma for floats
            row = [str(value).replace('.', ',') if isinstance(value, float) else value for value in row]
            writer.writerow(row)

    with open(f'{beerName}_ferm.csv', mode='w', newline='') as file:
        writer = csv.writer(file, delimiter=';')
        for row in df2:
            # Replace decimal point with comma for floats
            row = [str(value).replace('.', ',') if isinstance(value, float) else value for value in row]
            writer.writerow(row)


    #############################xslx fusion
    wb = load_workbook(excel_layout)

    ws1 = wb['brew']
    ws2 = wb['ferm']

    start_row = 1
    start_column = 1
    for row in df:
        for col_num, value in enumerate(row, start=start_column):
            ws1.cell(row=start_row, column=col_num, value=value)
        start_row += 1  # Move to the next row

    start_row = 1
    start_column = 1
    for row in df2:
        for col_num, value in enumerate(row, start=start_column):
            ws2.cell(row=start_row, column=col_num, value=value)
        start_row += 1  # Move to the next row

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    # Return buffer so Flask can serve it
    wb.save(out_path)
    
    
    """
    wb.save(f'{beerName}_brewsheet.xlsx')
    """