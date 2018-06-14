#!/usr/bin/env python

# code snippet taken from http://www.videntity.com/2010/05/check-to-make-sure-a-process-is-running-and-restart-it-if-its-not-a-recipe-in-python/
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
   
   cur_time = str(datetime.now().time())
   c = Client(config.sms_sid, config.sms_token)
   c.messages.create(body='The discord bot crashed and was restarted at ' + cur_time, from_=str(+14078900127),
                          to=config.m_phone)
   c.messages.create(body='The discord bot crashed and was restarted at ' + cur_time, from_=str(+14078900127),
                          to=config.j_phone)
   call("python3 discord_bot.py", shell = True)
   
