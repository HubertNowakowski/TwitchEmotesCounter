import requests
import json
import datetime
import sys
import xlwt

VERSION = "1.0"
api_url="https://twitchemotes.com/api_cache/v3/"
graph_url = "https://twitchemotes.com/api/stats/total/graph"


class Emotes:
    ID, CODE, NUM = range(3)



def getchannelId(data, name):
    try:
        id =  [ x for x in data if name in data[x]["channel_name"] ]
        return id[0]
    except IndexError:
        return None
    except:
        print("Error occured while searching for channel name")
        return None


def inputDatetime():
    while True:
        try:
            strin = input("Enter a date in YYYY-MM-DD format: ")
            date = datetime.datetime.strptime(strin,"%Y-%m-%d")
            result = True
        except ValueError as e:
            print(e)
        else:
            return date


def confirmInput(text):
    confirm = False
    while not confirm:
        ans = input(text+" (Y/n): ")
        if ans is "n":
            return False
        elif ans is "Y":
            return True
        else:
            print("Unknown answer.")


def createXLSFile(emotes, channelName, start_date, end_date):
    filename = "emotes_count_{}{}.xls".format(channelName,datetime.datetime.now().strftime("%Y-%m-%d"))

    sorted_emotes = sorted(emotes, key=lambda emotes: emotes[2], reverse=True)
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Emotes sum")

    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'D-MMM-YY'

    sheet1.write(0,0,"start date: ")
    sheet1.write(0,1,start_date, date_format)
    sheet1.write(1,0,"end date: ")
    sheet1.write(1,1,end_date, date_format)
    sheet1.write(2,0,"ID")
    sheet1.write(2,1,"CODE")
    sheet1.write(2,2,"COUNT")

    y=3
    for list in sorted_emotes:
        sheet1.write(y,0,list[Emotes.ID])
        sheet1.write(y,1,list[Emotes.CODE])
        sheet1.write(y,2,list[Emotes.NUM])
        y+=1
    try:
        book.save(filename)
        print("Data saved as {} in program directory.\n".format(filename))
    except:
        print("Error saving file.")


def printEmoteTable(channelName, channelId, emotes):
    print( "\nChannel name: {}\nChannel ID: {}".format(channelName, channelId) )
    sorted_emotes = sorted(emotes, key=lambda emotes: emotes[2], reverse=True)
    print("\n{:^16} | {:^16} | {:^16}".format("Id","Code","Count"))
    line = "------------------"
    print("{:16}|{:16}|{:16}".format(line, line, line))
    print("\n".join([" | ".join(["{:^16}".format(item) for item in row])
          for row in sorted_emotes]))


print("Twitch emotes counter {}\n".format(VERSION))

while True:
    channelName = input("Please enter the channel name: ")
    print("Connecting to API")
    response = requests.get(api_url+"subscriber.json")
    if response.status_code == 200:
        print("Succes!\nSearching for channel {}".format(channelName))
        data = response.json()
        channelId = getchannelId( data, channelName)
        if channelId is not None:
            print("Found channel: {}".format(data[channelId]["channel_name"]) )

            if confirmInput("Continue with that channel?"):
                channelName = data[channelId]["channel_name"]

                print("\nWhen do you want to start counting?")
                start_date = inputDatetime()
                print("\nWhen do you want to end counting?")
                end_date = inputDatetime()
                start_unix = start_date.timestamp()
                end_unix   = end_date.timestamp()
                print( "\nChecking data form: {} to {} ".format(start_date, end_date) )
                nr_emotes = len(data[channelId]["emotes"])
                emotes = [[0 for x in range(3)] for x in range(nr_emotes)]
                print("\nCounting emotes.")

                for ii in range ( 0, len(data[channelId]["emotes"]) ):
                    emotes[ii][Emotes.ID] = data[channelId]["emotes"][ii]["id"]
                    emotes[ii][Emotes.CODE] = data[channelId]["emotes"][ii]["code"]
                    data2=requests.get(graph_url,params="id={}".format(emotes[ii][Emotes.ID]) ).json()
                    sum = 0
                    for x in range (0, len(data2[0]["data"]) ):
                        day = data2[0]["data"][x][0]/1000
                        if day >= start_unix and day <= end_unix :
                            sum+=data2[0]["data"][x][1]
                    emotes[ii][Emotes.NUM] = sum

                printEmoteTable(channelName, channelId, emotes)

                if confirmInput("\nDo you want to create xls file?"):
                        createXLSFile(emotes, channelName, start_date, end_date)
        else:
            print("No data for channel {}.\n".format(channelName))
    else:
        print("Error connecting to API. STATUS CODE {}".format(response.status_code))

    if not confirmInput("Do you wish to search for another channel?"):
        sys.exit()
