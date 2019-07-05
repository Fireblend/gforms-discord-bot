#####################################################
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

################################
#            SETUP
################################

# Paste spreadsheet ID here
SPREADSHEET_ID = '<INSERT SPREADSHEET ID HERE>'
# Paste range here (can be sheet name)
RANGE = '<INSERT SPREADSHEET NAME HERE>'
# Discord Token
DISCORD_TOKEN = '<INSERT DISCORD BOT TOKEN HERE>'
# Update time (10 minutes default)
WAIT_TIME = 60*10

#################################
#       FORMATTING FUNCTION
#################################

# Modify this function to determine how each row will be formatted when posted to Discord.
# The decode calls are there to sanitize the input.
def format_row(row):
    username = row[1].decode("utf8","ignore")
    stage_name = row[2].decode("utf8","ignore")
    code = row[3].decode("utf8","ignore")
    style = row[4].decode("utf8","ignore")
    tags = row[5].decode("utf8","ignore")
    return ('%s: \"**%s**\" [**%s**] (%s; %s)' % (username, stage_name, code, style, tags))

################################
#      SHEETS API PORTION
################################

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# This function contacts the sheets API, retrieves values from the specified sheet,
# and prepares the message to be posted by the bot.
def makePost(last_row=0):
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
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()

    # Retrieve values
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE).execute()
    values = result.get('values', [])

    # Start with an empty string
    post = ""
    if not values:
        return post
    else:
        # Append a line to the result post for every row in the sheet:
        for row in values[last_row:]:
            last_row = last_row+1
            post = post+format_row(row)+'\n'

        # Save the last row number
        with open('last_row.pickle', 'wb') as last_row_p:
            pickle.dump(last_row, last_row_p)

        return post

############################
#        DISCORD BOT
############################

# Start the client
client = discord.Client()

# Controls update loop
stop = True
task = None

@client.event
async def on_message(message):
    global task
    global stop

    # Prevent the vote replying to itself (unlikely, but still a good precation)
    if message.author == client.user:
        return

    # Loop interrupt logic
    if message.content.startswith('!stop'):
        # If already stopped, show a message
        if(stop):
            await message.channel.send("Already stopped!")
            return
        # Stop the update loop
        stop = True
        task.cancel()
        await message.channel.send("Updates stopped. Start me up again with the **!start** command!")

    # Loop start logic
    elif message.content.startswith('!start'):
        # If already started, show a message
        if(not stop):
            await message.channel.send("Already started!")
            return

        stop = False

        # Determine the row to start from
        startFrom = 0

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
        await message.channel.send("Beep boop! Starting from row %s!\nUpdates every 10 minutes!\nStop me with **!stop**\n---" % startFrom)
        await asyncio.sleep(3)
        print("Started")
        while not stop:
            # Retrieve last row number from cache:
            if os.path.exists('last_row.pickle'):
                with open('last_row.pickle', 'rb') as last_row_p:
                    startFrom = pickle.load(last_row_p)

            # Build the message using the spreadsheets API
            msg = makePost(startFrom)

            # If non-empty, post the message
            if(msg != None and msg != ""):
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
    print('Logged in as')
    print(client.user.name)
    print(client.user.id)
    print('------')

# Run the client
client.run(DISCORD_TOKEN)
