from flask import Flask,jsonify,request,send_file
import json
import urllib.request
import datetime
import pandas as pd

app = Flask(__name__)

@app.route('/total', methods=['GET'])
def get_tasks():
    if (request.method=="GET"):
        day=request.args.get("day")
        #print(day)
    data=urllib.request.urlopen("https://assignment-machstatz.herokuapp.com/excel").read()
    jdata=json.loads(data)
    retdict = {"totalWeight": 0, "totalLength": 0, "totalQuantity": 0}
    for d in jdata:
        dateTimeobj=datetime.datetime.strptime(d["DateTime"][0:10],'%Y-%m-%d')
        if(dateTimeobj.strftime("%d-%m-%Y")==day):
            retdict["totalWeight"]=retdict["totalWeight"]+d["Weight"]
            retdict["totalLength"]=retdict["totalLength"]+d["Length"]
            retdict["totalQuantity"]=retdict["totalQuantity"]+d["Quantity"]
    #print(day)
    return jsonify(retdict)
@app.route('/excelreport')
def download_excel():
    data = urllib.request.urlopen("https://assignment-machstatz.herokuapp.com/excel").read()
    jdata = json.loads(data)
    currentDateTime=jdata[0]["DateTime"][0:10]
    length=[]
    quantity=[]
    weight=[]
    dates=[]
    writer = pd.ExcelWriter('data_Final.xlsx', engine='xlsxwriter')
    for dat in jdata:

        if(dat["DateTime"][0:10]!=currentDateTime):

                #currentDateTime=dat["DateTime"][0:10]
                #print(currentDateTime)
                #print(length)
                d = {'DateTime':dates,'Length':length,'Quantity': quantity, 'Weight': weight}
                length=[]
                quantity=[]
                weight=[]
                dates=[]
                df = pd.DataFrame(d)

    # Convert DF


                # Write each dataframe to a different worksheet.
                df.to_excel(writer, sheet_name=currentDateTime,index=None)
                currentDateTime=dat["DateTime"][0:10]
                #print(currentDateTime)
        length.append(dat["Length"])
        quantity.append(dat["Quantity"])
        weight.append(dat["Weight"])
        dates.append(dat["DateTime"])
    d = {'DateTime':dates,'Length': length, 'Quantity': quantity, 'Weight': weight}
    df = pd.DataFrame(d)
    df.to_excel(writer, sheet_name=currentDateTime,index=None)
    writer.save()
    return send_file('data_Final.xlsx',
                     attachment_filename='finalfile.xlsx',
                     as_attachment=True)


