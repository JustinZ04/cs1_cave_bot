import discord
import asyncio
import config
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
        await client.send_message(message.channel, "Command received")



client.run(config.token)
