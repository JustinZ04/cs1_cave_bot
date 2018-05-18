#!/usr/bin/env python
# Written by Matthew Villegas & Justin Zabel
# Script to run the discord bot itself.
# Handles logging in and receiving the correct command to run the parse script. Outputs the current TA's holding office
# hours.

import discord
import config

from datetime import datetime
from parse_excel import parse

client = discord.Client()


@client.event
async def on_ready():
    # Logs the bot in so it goes online.
    print('Logged in as')
    print(client.user.name)
    print(client.user.id)
    print('-----')


@client.event
async def on_message(message):
    # Receives the '!cave' command and returns the current TA's holding office hours.
    # Make sure to parse the right spreadsheet, so use the channel name.
    # Need to make sure names match channel name of server.

    if message.content.startswith('!cave1'):
        ta_list = parse(1)

        if ta_list is None:
            await client.send_message(message.channel, "No TA's are in the cave right now!")
        else:
            s = '\n\n\n'.join(ta_list)
            await client.send_message(message.channel, s)

    if message.content.startswith('!cave2'):
        ta_list = parse(2)

        if ta_list is None:
            await client.send_message(message.channel, "No TA's are in the cave right now!")
        else:
            s = '\n\n\n'.join(ta_list)
            await client.send_message(message.channel, s)

        # Log every time a '!cave' command is received to a log file.
        cur_time = datetime.now().time()
        file = open("logs/commands.log", "a")
        file.write(str(cur_time) + " Cave command typed\n")
        file.close()

    if message.content.startswith('!help'):
        await client.send_message(message.channel, "This is the UCF CS Cave Bot! Available commands are:\n"
                                                   "!help: Displays this message.\n"
                                                   "!cave1: Displays the office hours for CS1 TAs.\n"
                                                   "!cave2: Displays the office hours for CS2 TAs.\n"
                                                   "A link to source code can be found in the references channel.")

# Holds the authorization token for the bot.
client.run(config.bot_token)
