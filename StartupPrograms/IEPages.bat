@echo off
start "" "https://tmna.service-now.com/navpage.do"
ping -n 2 localhost > nul

start "" "https://tmna.service-now.com/navpage.do"
ping -n 2 localhost > nul

start "" "https://tmna.service-now.com/$knowledge.do"
ping -n 2 localhost > nul

start "" "https://tmna.service-now.com/1ts"
ping -n 2 localhost > nul