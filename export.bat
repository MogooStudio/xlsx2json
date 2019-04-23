@echo off
title [convert excel to json]
echo press any button to start.
@pause > nul
echo start converting ....
node index.js --export StringData_CN C:\Users\pc7\Desktop\execl2Json\app\pages\xlsx2json-master\excel\String.xlsx C:\Users\pc7\Desktop\execl2Json\app\pages\xlsx2json-master\excel\json\
echo convert over!
@pause