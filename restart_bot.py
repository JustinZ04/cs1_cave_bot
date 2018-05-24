#!/usr/bin/env python
"""
Check to see if an process is running. If not, restart.
Run this in a cron job
"""
import os
from subprocess import call
process_name= "discord_bot.py" # change this to the name of your process

tmp = os.popen("ps -Af").read()

if process_name not in tmp[:]:
   
   call("python3 discord_bot.py", shell = True) 

