import requests
import json
import datetime
import sys
import xlwt

__author__ = "Hubert Nowakowski"
__license__ = "GPL-3.0"
__version__ = "1.2"


api_url='https://twitchemotes.com/api_cache/v3/'
graph_url = 'https://twitchemotes.com/api/stats/total/graph'


class Emote:

    def __init__(self, id, code, set):
        self.id = id
        self.code = code
        self.set = set
        self.sum = 0
        self.dailyCount={}

    def __gt__(self, emote2):
        return self.sum > emote2.sum

    def calcSum(self, start_date, end_date):
        self.sum = 0
        for key, value in self.dailyCount.items():
            newkey = key/1000                                                   #because API sets the date with 000 at the end
            if newkey >= start_date and newkey <= end_date:
                self.sum += value


def getchannelId(data, name):
    try:
        id =  [ x for x in data if name in data[x]['channel_name'] ]
        return id[0]
    except IndexError:
        return None
    except:
        print('Error occured while searching for channel name')
        return None


def inputDatetime():
    while True:
        try:
            strin = input('Enter a date in YYYY-MM-DD format: ')
            date = datetime.datetime.strptime(strin,'%Y-%m-%d')
            result = True
        except ValueError as e:
            print(e)
        else:
            return date


def convertUnix(date):
    return int(date.timestamp())


def confirmInput(text):
    confirm = False
    while not confirm:
        ans = input(text+' (Y/n): ')
        if ans is 'n':
            return False
        elif ans is 'Y':
            return True
        else:
            print('Unknown answer.')


def printEmoteTable(channelName, channelId, emotes):
    print( '\nChannel name: {}\nChannel ID: {}'.format(channelName, channelId) )
    sorted_emotes = sorted(emotes, reverse=True)
    print('\n{:^10}|{:^10}|{:^10}'.format('Id','Code','Count'))
    line = '----------'
    print('{:^10}|{:^10}|{:^10}'.format(line, line, line))
    for emote in sorted_emotes:
        print('{:^10}|{:^10}|{:^10}'.format(emote.id, emote.code, emote.sum))


def createXLSFile(emotes, channelName, start_date, end_date):
    filename = 'emotes_counter_{}{}.xls'.format(channelName,datetime.datetime.now().strftime('%Y-%m-%d'))
    sorted_emotes = sorted(emotes, reverse=True)
    date_format = xlwt.XFStyle()
    date_format.num_format_str = 'D-MMM-YY'

    book = xlwt.Workbook(encoding='utf-8')
    sheet1 = book.add_sheet('Emotes sum')

    sheet1.write(0,0,'start date: ')
    sheet1.write(0,1,start_date, date_format)
    sheet1.write(1,0,'end date: ')
    sheet1.write(1,1,end_date, date_format)
    start_date = convertUnix(start_date)
    end_date = convertUnix(end_date)

    sheet1.write(2,0,'ID')
    sheet1.write(2,1,'CODE')
    sheet1.write(2,2,'SUM')
    y=3
    for emote in sorted_emotes:
        sheet1.write(y,0,emote.id)
        sheet1.write(y,1,emote.code)
        sheet1.write(y,2,emote.sum)
        y+=1

    sheet2 = book.add_sheet('Daily data')
    y=0
    for emote in sorted_emotes:
        sheet2.write(y,0, emote.code)
        y+=1
        for day, count in emote.dailyCount.items():
            newday = day/1000                                                   #because API sets the date with 000 at the end, probably in ms instead of s
            if newday >= start_date and newday <= end_date:
                sheet2.write(y,0, newday)
                sheet2.write(y,1, count)
                sheet2.write(y,2, datetime.datetime.utcfromtimestamp(newday).strftime('%Y-%m-%d'), date_format)
                y+=1
        y+=2

    try:
        book.save(filename)
        print('Data saved as {} in program directory.\n'.format(filename))
    except:
        print('Error saving file.')



print('Twitch emotes counter {}\n'.format(__version__))

while True:
    channelName = input('Please enter the channel name: ')
    print('Connecting to TwitchEmotes API')
    response = requests.get(api_url+'subscriber.json')
    if response.status_code == 200:
        print('Succes!\nSearching for channel {}'.format(channelName))
        data = response.json()
        channelId = getchannelId( data, channelName)
        if channelId is not None:
            print('Found channel: {}'.format(data[channelId]['channel_name']) )
            if confirmInput('Continue with that channel?'):
                channelName = data[channelId]['channel_name']

                print('\nWhen do you want to start counting?')
                start_date = inputDatetime()
                print('\nWhen do you want to end counting?')
                end_date = inputDatetime()
                print( '\nChecking data form: {} to {} '.format(start_date, end_date) )

                print("I'm getting the emote data from Graph API.")
                emoteData = data[channelId]['emotes']
                emotes = [ Emote( row['id'], row['code'], row['emoticon_set'] )
                           for row in emoteData]
                for emote in emotes:
                    dailyData = requests.get(graph_url,params='id={}'.format(emote.id)).json()
                    for day in dailyData[0]['data']:
                        if day[0]/1000 >= convertUnix(start_date) and day[0]/1000 <= convertUnix(end_date):
                            emote.dailyCount[day[0]] = day[1]
                    emote.calcSum(convertUnix(start_date), convertUnix(end_date))

                printEmoteTable(channelName, channelId, emotes)

                if confirmInput('\nDo you want to create xls file?'):
                        createXLSFile(emotes,channelName,start_date,end_date)
        else:
            print('No data for channel {}.\n'.format(channelName))
    else:
        print('Error connecting to API. STATUS CODE {}'.format(response.status_code))

    if not confirmInput('Do you wish to search for another channel?'):
        sys.exit()
