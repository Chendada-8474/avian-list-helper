import pandas as pd
import numpy as np
import datetime as dt

fileName = input("請輸入xlsx檔案名稱(不用輸入副檔名)：")
dataPath = './%s.xlsx' % (fileName)
data = pd.read_excel(dataPath)
listPath = './2020_list.xlsx'
birdList = pd.read_excel(listPath)


data = data.merge(birdList, on='ch_name', how='left')
data = data.sort_values(by=['id_num'], ascending=True)
data = data.replace(np.nan, '', regex=True)

# print(dataOutput)


colNameCh = {
    "1": "中文俗名",
    "2": "學名",
    "3": "亞種分化",
    "4": "英文俗名",
    "5": "科名",
    "6": "台灣留棲狀態",
    "7": "馬祖留棲狀態",
    "8": "金門留棲狀態",
    "9": "東沙留棲狀態",
    "10": "特有性",
    "11": "保育等級",
    "12": "台灣鳥類紅皮書等級",
}

colNameEn = {
    "1": "ch_name",
    "2": "s_name",
    "3": "sub_sp",
    "4": "en_name",
    "5": "family",
    "6": "taiwan",
    "7": "matzu",
    "8": "kinmen",
    "9": "tonsa",
    "10": "endemic",
    "11": "con_level",
    "12": "red_list",
}

def myDiv():
    print("-" * 50)

def colSelctor():
    colList = []
    for i in colNameCh:
        print(i, ":",colNameCh[i])
    myDiv()

    while True:
        x = str(input('請輸入想要的欄位代號：')) 
        print("目前選擇的欄位有：\n選完請直接按Enter繼續")
        for y in colList:
            print(colNameCh[y])
        myDiv()

        if x == "":
            break

        elif x in colList:
            print("欄位重複，請重新輸入")
            myDiv()
        
        elif x in colNameCh.keys():
            colList.append(x)
            print("目前選擇的欄位：\n選完請直接按Enter")
            for y in colList:
                print(colNameCh[y])
            myDiv()

        else:
            print("請輸入正確的代號\n目前選擇的欄位有：")
            for y in colList:
                print(colNameCh[y])
            myDiv()

    return colList


def finalOutput(aList):
    colList = aList
    opDict = {}
    for i in colList:
        opDict[colNameCh[i]] = data[colNameEn[i]].tolist()
    opFinal = pd.DataFrame(opDict)
    return opFinal



yourBirdList = finalOutput(colSelctor())
print(yourBirdList)

timeNow = dt.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
yourBirdList.to_excel(r'./output_%s.xlsx' % (timeNow), index = False)


# ch_name = birdOutput['ch_name'].tolist()
# s_name = birdOutput['s_name'].tolist()
# family = birdOutput['family'].tolist()
# con_level = birdOutput['con_level'].tolist()
# red_list = birdOutput['red_list'].tolist()
# count = birdOutput['count'].tolist()
# precent = birdOutput['precent'].tolist()

# # print(birdOutput)



# outputDict = {
#     "中文名": ch_name,
#     "學名": s_name,
#     "科名": family,
#     "保育等級": con_level,
#     "紅皮書受脅等級": red_list,
#     "累計調查隻次": count,
#     "有記錄之樣區數": precent,
# }

# finalOutput = pd.DataFrame(outputDict)

# # finalOutput.to_excel(r'./output.xlsx', index = False)


# print(finalOutput)