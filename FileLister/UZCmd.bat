copy %1 %TEMP%\application.zip
..\7-Zip\7z x %TEMP%\application.zip -o%TEMP%\%2\
del %TEMP%\application.zip
