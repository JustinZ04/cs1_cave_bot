#!/usr/bin/env python
"""
Check to see if an process is running. If not, restart.
Run this in a cron job
"""
import os
import config

from datetime import datetime
from subprocess import call
from twilio.rest import Client
process_name= "discord_bot.py" # change this to the name of your process

tmp = os.popen("ps -Af").read()

if process_name not in tmp[:]:
   
   call("python3 discord_bot.py", shell = True)
   cur_time = datetime.now().time()
   c = Client(config.sms_sid, config.sms_token)
   c.messages.create(body='The discord bot crashed and was restarted at ' + cur_time, from_=str(+14078900127),
                          to=config.m_phone)
   c.messages.create(body='The discord bot crashed and was restarted at ' + cur_time, from_=str(+14078900127),
                          to=config.j_phone)
