from flask import Flask,jsonify,request,send_file
import json
import urllib.request
import datetime
import pandas as pd

app = Flask(__name__)

@app.route('/total', methods=['GET'])
def get_tasks():
    if (request.method=="GET"):                                                      #Using HTTP Get method to fetch the date from url
        day=request.args.get("day")                                                  #Assigning date to "day" variable
        #print(day)
    data=urllib.request.urlopen("https://assignment-machstatz.herokuapp.com/excel").read()   #fetching json data from the given URL
    jdata=json.loads(data)
    retdict = {"totalWeight": 0, "totalLength": 0, "totalQuantity": 0}              #Initialising Dictionary with totalLength=0,totalWeight=0,totalQuantity=0
    for d in jdata:                                                                 #Generator for json data
        dateTimeobj=datetime.datetime.strptime(d["DateTime"][0:10],'%Y-%m-%d')      #Converting Date format from "YYYY-mm-ddTHH-MM-SS" to "YYYY-mm-dd"
        if(dateTimeobj.strftime("%d-%m-%Y")==day):                                  #Comparing dates of each item in dictionary with the day variable
            retdict["totalWeight"]=retdict["totalWeight"]+d["Weight"]               #Adding Weight of the item to the total Weight
            retdict["totalLength"]=retdict["totalLength"]+d["Length"]               #Adding Length of the item to the total Length
            retdict["totalQuantity"]=retdict["totalQuantity"]+d["Quantity"]         #Adding Quantity of the item to the total Quantity
    #print(day)
    return jsonify(retdict)                                                         #Returning Retdict(dictionary) in json format
@app.route('/excelreport')
def download_excel():
    data = urllib.request.urlopen("https://assignment-machstatz.herokuapp.com/excel").read() #Using HTTP Get method to fetch the date from URL
    jdata = json.loads(data)                                                            #Loading JSON data
    currentDateTime=jdata[0]["DateTime"][0:10]                   #Converting Date format from "YYYY-mm-ddTHH-MM-SS" to "YYYY-mm-dd and assign 1st itemdate to current date varible"
    length=[]                                                                           #List for Length
    quantity=[]                                                                         #List for Quantity
    weight=[]                                                                           #List for Weight
    dates=[]                                                                            #List for Dates
    writer = pd.ExcelWriter('data_Final.xlsx', engine='xlsxwriter')                     #Creation of Excel File
    for dat in jdata:                                                   #Iterating each and every Item in data and if data was found with new date
######################################################################### Write data into excel sheet with the name of current date
        if(dat["DateTime"][0:10]!=currentDateTime):

                #currentDateTime=dat["DateTime"][0:10]
                #print(currentDateTime)
                #print(length)
                d = {'DateTime':dates,'Length':length,'Quantity': quantity, 'Weight': weight}  #Creating dataframe for writing into excel Column name: DateTime,Length,Quantity
                length=[]
                quantity=[]
                weight=[]
                dates=[]
                df = pd.DataFrame(d)

    # Convert DF


                # Write each dataframe to a different worksheet.
                df.to_excel(writer, sheet_name=currentDateTime,index=None)          #Writing data into excel sheet for each and every date
                currentDateTime=dat["DateTime"][0:10]
                #print(currentDateTime)
        length.append(dat["Length"])
        quantity.append(dat["Quantity"])
        weight.append(dat["Weight"])
        dates.append(dat["DateTime"])
    d = {'DateTime':dates,'Length': length, 'Quantity': quantity, 'Weight': weight}   #Writing data into last sheet for last date
    df = pd.DataFrame(d)
    df.to_excel(writer, sheet_name=currentDateTime,index=None)
    writer.save()
    return send_file('data_Final.xlsx',
                     attachment_filename='finalfile.xlsx',
                     as_attachment=True)                                                #Returning XlSX file to the browser


