'''General Information
    Access Token: AlQG8tOS5MfwKtcIxhQZtgAkYvw1Ag6knsO6rTD5
    Link to Documentation: https://pypi.org/project/GroupyAPI/
'''
from groupy.client import Client

TOKEN = "AlQG8tOS5MfwKtcIxhQZtgAkYvw1Ag6knsO6rTD5"

client = Client.from_token(TOKEN)

def message(channel,content):
    if channel == None:
        channel = "Bot Testing"
    for group in client.groups.list_all():
        if str(group) == f"<Group(name='{channel}')>":
            message = group.post(text = content)
        
def getChannels():
    for group in client.groups.list_all():
        print(str(group.name))
        


