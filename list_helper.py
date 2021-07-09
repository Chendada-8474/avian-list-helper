# pyinstaller -F list_helper.py

try:
    import pandas as pd
    import numpy as np
    import datetime as dt

    def myDiv(n):
        print("-" * n)

    def intro():
        print("")
        print("1. 準備你自己的鳥類紀錄檔案，可以是鳥類調查的row data 只要裡面有學名即可。\n2. 把檔案存在跟程式同一個資料夾內。\n3. 2020_list.xlsx 這一個檔案也要放在同一個資料夾內。這裡面的內容可以直接在裡面更改，例如鳥種有誤之類的。\n4. 啟動 list_helper.exe。程式啟動後輸入鳥種清單檔名，不用打副檔名。如入完檔名後輸入你的資料中，中文鳥名的欄為(例如demo中的 學名)\n5. 選擇想要的欄位，一次選一個。\n6. 式就會產生名錄excel檔，檔名為：output_目前日期時間.xlsx。")
        myDiv(50)

    intro()

    fileName = input("請輸入xlsx檔案名稱(不用輸入副檔名)：")
    colName = input("請輸入你的excel中，學名欄為名稱：")
    dataPath = './%s.xlsx' % (fileName)
    data = pd.read_excel(dataPath)
    listPath = './2020_list.xlsx'
    birdList = pd.read_excel(listPath)

    data = data.rename(columns= {"%s" % colName: "s_name"})

    sName = data["s_name"].tolist()
    sName = list(set(sName))

    dataDict = {
        "s_name": sName
    }

    data = pd.DataFrame(data = dataDict)

    data = data.merge(birdList, on='s_name', how='left')
    data = data.sort_values(by=['id_num'], ascending=True)
    data = data.replace(np.nan, '', regex=True)

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


    def colSelctor():
        colList = []
        for i in colNameCh:
            print(i, ":", colNameCh[i])
        myDiv(50)

        while True:
            x = str(input('請輸入想要的欄位代號：'))
            print("目前選擇的欄位有：\n選完請直接按Enter繼續")
            for y in colList:
                print(colNameCh[y])
            myDiv(50)

            if x == "":
                break

            elif x in colList:
                print("欄位重複，請重新輸入")
                myDiv(50)

            elif x in colNameCh.keys():
                colList.append(x)
                for i in colNameCh:
                    print(i, ":", colNameCh[i])
                myDiv(25)

                print("目前選擇的欄位：\n選完請直接按Enter")
                for y in colList:
                    print(colNameCh[y])
                myDiv(50)

            else:
                for i in colNameCh:
                    print(i, ":", colNameCh[i])
                myDiv(25)

                print("請輸入正確的代號\n目前選擇的欄位有：")
                for y in colList:
                    print(colNameCh[y])
                myDiv(50)

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
    myDiv(50)

    timeNow = dt.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    yourBirdList.to_excel(r'./output_%s.xlsx' % (timeNow), index=False)

except Exception as e:
    print(e)

input("名錄已經輸出至相同資料夾下，\n檔名為：output_%s.xlsx。\n感謝您的使用。\n請輸入任意鍵退出程式或直接按Enter結束程式：" %timeNow)