Here's what I used to use for that scenario, obviously you'll have to tweak it for your drive letters, and it could be done a little bit better but will give you the general idea:

@Echo off SET /P variable=[Enter the Username Now]

@echo Prepping USMT

xcopy F:\USMT C:\USMT /s /y /i

set USMT_WORKING_DIR=c:\USMT

set MIG_OFFLINE_PLATFORM_ARCH=32

@echo Copying UserData of C:\ to F:\backups\

c:

cd USMT

scanstate.exe /UEL:30 F:\backups\%variable% /offline:c:\USMT\offline.xml /i:migapp.xml /i:miguser.xml /o /config:config.xml /v:5

And here's the restore script:

@Echo off SET /P variable=[Enter the Username Now]

e:\usmt\loadstate e:\backups\%variable% /config:e:\usmt\config.xml /i:e:\usmt\miguser.xml /i:e:\usmt\migapp.xml /v:5 /l:loadstate.log /lac

