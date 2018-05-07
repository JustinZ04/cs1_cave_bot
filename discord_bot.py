import discord
import asyncio
import config
from datetime import datetime

from parse_excel import parse

client = discord.Client()


@client.event
async def on_ready():
    print('Logged in as')
    print(client.user.name)
    print(client.user.id)
    print('-----')


@client.event
async def on_message(message):
    if message.content.startswith('!cave'):
        ta_list = parse()
        print(len(ta_list))
        if ta_list is None:
            await client.send_message(message.channel, "No TA's are in the cave right now!")

        else:
            # await client.send_message(message.channel, "Command received")

            s = '\n\n\n'.join(ta_list)


            await client.send_message(message.channel, s)

            # if i != len(ta_list)-1:
            #    await client.send_message(message.channel, '----------')

    cur_time = datetime.now().time()
    file = open("logs/commands.log", "a")
    file.write(str(cur_time) + " Cave command typed\n")
    file.close()


client.run(config.bot_token)
