####################################################
#
#   GForms Bot: An easily adaptable Discord bot
#   that interacts with the Google Sheets API,
#   intended to report on Google Forms responses,
#   but probably useful for other stuff too :)
#
#   Made by Fireblend (https://www.fireblend.com)
#
####################################################

from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import discord
import asyncio
import random

################################
#            SETUP
################################

# Paste spreadsheet ID here
SPREADSHEET_ID = '1vq_HP94RIWGL5zWIIVyGMAv1U6nOmkPPpyg-7mds-I0'
# The id of the channel
CHANNEL_ID = '829011607513989170'
# Paste range here (can be sheet name)
RANGE = 'A:D'
# Discord Token
DISCORD_TOKEN = 'ODI4Njc5Nzc1NjMwNjU1NTA4.YGtGLg.lQ8xCwBlokwiUaJ9EzBJ8ER_kSc'
# The first response row on the sheet
STARTING_ROW = 2
# Update time (5 minutes default)
WAIT_TIME = 60*1
# User roles that can interact with this bot:
# Use all-lowercase regardless of role capitalization
# Leave empty if any
ROLES = ['admin', 'mods', 'zucced']
# Channels where interaction with this bot are allowed:
# Use all-lowercase regardless of role capitalization
# Leave empty if any
CHANNELS = ['anomologita']

#########################################################

# You can enable a command that lets anyone request a random
# response with a different formatting style. Set to True if you'd
# like to enable this command.
RANDOM_ENABLED = False
RANDOM_ROLES = []
RANDOM_CHANNELS = []

#################################
#       FORMATTING FUNCTION
#################################

# Modify this function to determine how each row will be formatted when posted to Discord.
def format_row(row,hashtag):
    date = row[0]
    text = row[1]
    if len(row)==3:
        if len(row[1])+len(row[2])>2000:
            return ('>>> ```diff\n- Η καταχώρηση έχει παραπάνω χαρακτήρες από το μέγιστο όριο του Discord και θα παραλειφθεί\n```')
        dept = row[2]
        return ('>>> __*#Ανομολόγητο%s*__\n*%s*\n__#%s__\nΣτείλε το δικό σου εσώψυχο! <https://forms.gle/8AuLBYHS9mKamJD4A>' % (hashtag, text, dept))
    elif len(row)==4:
        if len(row[1])+len(row[2])+len(row[3]) >2000:
            return ('>>> ```diff\n- Η καταχώρηση έχει παραπάνω χαρακτήρες από το μέγιστο όριο του Discord και θα παραλειφθεί\n```')
        dept = row[2]
        etos = row[3]
        return ('>>> __*#Ανομολόγητο%s*__\n*%s*\n__#%s__ #%so\nΣτείλε το δικό σου εσώψυχο! <https://forms.gle/8AuLBYHS9mKamJD4A>' % (hashtag, text, dept, etos))

# Modify this function to determine how an individual row will be formatted when using the random command.
# Only used if RANDOM_ENABLED is set to True above.
def format_random(row):
    username = row[1]
    stage_name = row[2]
    code = row[3]
    style = row[4]
    tags = row[5]
    return ('**%s** by **%s**\n**Code**: %s\n**Style**: %s\n**Tags**: %s' % (stage_name, username, code, style, tags))

################################
#      SHEETS API PORTION
################################

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/drive']

# This function contacts the sheets API, retrieves values from the specified sheet,
# and prepares the message to be posted by the bot.
def makePost(last_row=0, pick_random=False):
    try:
        # Google Sheets Authentication flow
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                # creds = flow.run_console() #Alternative authorization flow for console-based systems
                creds = flow.run_local_server()
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
            
        service = build('sheets', 'v4', credentials=creds)
        # Call the Sheets API
        sheet = service.spreadsheets()
        sortData("ASCENDING", service);
        # Retrieve values
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                    range=RANGE).execute()
        values = result.get('values', [])
        sortData("DESCENDING", service);
        # If we're picking a random value, just format a random row:
        if pick_random:
            random_row = random.sample(range(STARTING_ROW, len(values)), 1)
            return format_random(values[random_row[0]])

        # If it's an update, start with an empty string:
        post = []
        if not values:
            return post
        else:
            # Append a line to the result post for every row in the sheet:
            for row in values[last_row:]:
                post.append(format_row(row,last_row+1))
                last_row = last_row+1
            
            # Save the last row number
            with open('last_row.pickle', 'wb') as last_row_p:
                pickle.dump(last_row, last_row_p)
            return post
    except:
        pass

def sortData(order,service):
    body = {
    'requests': [
        {
            'sortRange': {
                'range': {
                    'sheetId': 563107007,
                    'startRowIndex': 1,
                    'startColumnIndex': 0,
                    'endColumnIndex': 4
                },
                'sortSpecs': [
                    {
                        'sortOrder': order,
                        'dimensionIndex': 0
                    }
                ]
            }
        }
    ]
    }
    try:
        service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
    except:
        print("Ignoring timeout exception")

############################
#        DISCORD BOT
############################

# Start the client
client = discord.Client()

# Controls update loop
stop = True
task = None

async def is_channel_allowed(message, channels=CHANNELS):
    if(len(channels) == 0):
        return True
    channel = message.channel
    if channel.name.lower() not in channels:
        #await message.channel.send("You can't do that here! Are you in the right channel?")
        return False
    return True

async def is_user_allowed(message, roles=ROLES):
    if(len(roles) == 0):
        return True
    user = message.author
    if next((x for x in user.roles if x.name.lower() in roles), None) is None:
        #await message.channel.send("You can't do that! Do you have the right role?")
        return False
    return True

@client.event
async def on_message(message):
    global task
    global stop

    # Prevent the vote replying to itself (unlikely, but still a good precation)
    if message.author == client.user:
        return

    # Picks and posts a random response. Only used if RANDOM_ENABLED is set to True.
    if message.content.startswith('!random'):
        if not RANDOM_ENABLED:
            return

        # Check channel:
        if not await is_channel_allowed(message, channels=RANDOM_CHANNELS):
            return

        # Check user role:
        if not await is_user_allowed(message, roles=RANDOM_ROLES):
            return

        # Retrieve last row number from cache:
        startFrom = STARTING_ROW
        if os.path.exists('last_row.pickle'):
            with open('last_row.pickle', 'rb') as last_row_p:
                startFrom = pickle.load(last_row_p)

        # Post a random response
        await message.channel.send("Random level for **%s**:\n\n%s" % (message.author.display_name, makePost(last_row=startFrom, pick_random=True)))

    # Loop interrupt logic
    if message.content.startswith('!stop'):
        # Check channel:
        if not await is_channel_allowed(message):
            return

        # Check user role:
        if not await is_user_allowed(message):
            return

        # If already stopped, show a message
        if(stop):
            #await message.channel.send("Already stopped!")
            return

        # Stop the update loop
        stop = True
        task.cancel()
        await message.channel.send("Updates stopped. Start me up again with the **!start** command!")

    # Loop start logic
    elif message.content.startswith('!start'):
        # Check channel:
        if not await is_channel_allowed(message):
            return

        # Check user role:
        if not await is_user_allowed(message):
            return

        # If already started, show a message
        if(not stop):
            #await message.channel.send("Already started!")
            return

        stop = False

        # Determine the row to start from
        startFrom = STARTING_ROW

        # If hard-coded in the start command, use that.
        if(len(message.content.strip().split(" ")) > 1):
            startFrom= int(message.content.strip().split(" ")[1])
            with open('last_row.pickle', 'wb') as last_row_p:
                pickle.dump(startFrom, last_row_p)
        # If not, check if we have a row number cache'd from a previous run and use that.
        else:
            if os.path.exists('last_row.pickle'):
                with open('last_row.pickle', 'rb') as last_row_p:
                    startFrom = pickle.load(last_row_p)

        # Startup message
        mins = int(WAIT_TIME/60)
        #await message.channel.send("Beep boop! Starting from row %s!\nUpdates every %s minutes!\nStop me with **!stop**\n---" % (startFrom, mins))
        await asyncio.sleep(3)
        print("Started")
        await message.delete()
        while not stop:
            # Retrieve last row number from cache:
            if os.path.exists('last_row.pickle'):
                with open('last_row.pickle', 'rb') as last_row_p:
                    startFrom = pickle.load(last_row_p)

            # Build the message using the spreadsheets API
            posts = makePost(startFrom)

            # If non-empty, post the message
            if(posts != None and len(posts) > 0):
                for msg in posts:
                    await message.channel.send(msg)

            # Sleep until it's time for the next update
            coro = asyncio.sleep(WAIT_TIME)
            task = asyncio.ensure_future(coro)

            # Try block in case the update loop is stopped by the user
            try:
                await task
            except asyncio.CancelledError:
                print("Stopped")
    #kai apo dw thes
    else:
        # Check channel:
        if not await is_channel_allowed(message):
            return

        # Check user role:
        if not await is_user_allowed(message):
            return

        # If already started, show a message
        if(not stop):
            #await message.channel.send("Already started!")
            return

        stop = False

        # Determine the row to start from
        startFrom = STARTING_ROW

        # If hard-coded in the start command, use that.
        if(len(message.content.strip().split(" ")) > 1):
            startFrom= int(message.content.strip().split(" ")[1])
            with open('last_row.pickle', 'wb') as last_row_p:
                pickle.dump(startFrom, last_row_p)
        # If not, check if we have a row number cache'd from a previous run and use that.
        else:
            if os.path.exists('last_row.pickle'):
                with open('last_row.pickle', 'rb') as last_row_p:
                    startFrom = pickle.load(last_row_p)

        # Startup message
        mins = int(WAIT_TIME/60)
        #await message.channel.send("Beep boop! Starting from row %s!\nUpdates every %s minutes!\nStop me with **!stop**\n---" % (startFrom, mins))
        await asyncio.sleep(3)
        print("Started")
        while not stop:
            # Retrieve last row number from cache:
            if os.path.exists('last_row.pickle'):
                with open('last_row.pickle', 'rb') as last_row_p:
                    startFrom = pickle.load(last_row_p)

            # Build the message using the spreadsheets API
            posts = makePost(startFrom)

            # If non-empty, post the message
            if(posts != None and len(posts) > 0):
                for msg in posts:
                    await message.channel.send(msg)

            # Sleep until it's time for the next update
            coro = asyncio.sleep(WAIT_TIME)
            task = asyncio.ensure_future(coro)

            # Try block in case the update loop is stopped by the user
            try:
                await task
            except asyncio.CancelledError:
                print("Stopped")

@client.event
async def on_ready():
    print('Logged in as %s' % client.user.name)
    print('Invite to your server: https://discordapp.com/oauth2/authorize?client_id=%s&scope=bot' % client.user.id)
    print('------')
    

# Run the client
client.run(DISCORD_TOKEN)

@client.event
async def on_ready():
        # Loop start logic

        stop = False

        # Determine the row to start from
        startFrom = STARTING_ROW
        #Check if we have a row number cache'd from a previous run and use that.
        if os.path.exists('last_row.pickle'):
            with open('last_row.pickle', 'rb') as last_row_p:
                startFrom = pickle.load(last_row_p)
        # Startup message
        mins = int(WAIT_TIME/60)
        await asyncio.sleep(3)
        print("Started")
        while not stop:
            # Retrieve last row number from cache:
            if os.path.exists('last_row.pickle'):
                with open('last_row.pickle', 'rb') as last_row_p:
                    startFrom = pickle.load(last_row_p)

            # Build the message using the spreadsheets API
            posts = makePost(startFrom)

            # If non-empty, post the message
            if(posts != None and len(posts) > 0):
                for msg in posts:
                    await client.channels.cache.get(CHANNEL_ID).send('Hello here!')

            # Sleep until it's time for the next update
            coro = asyncio.sleep(WAIT_TIME)
            task = asyncio.ensure_future(coro)

            # Try block in case the update loop is stopped by the user
            try:
                await task
            except asyncio.CancelledError:
                print("Stopped")