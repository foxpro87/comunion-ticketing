copy /b SendEmail.dll "C:\Program Files\IT Group, Inc\ComUnion\Plugin\SendEmail.dll" 

path = C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727
regasm "C:\Program Files\IT Group, Inc\ComUnion\Plugin\SendEmail.dll" 
regasm "C:\Program Files\IT Group, Inc\ComUnion\Plugin\SendEmail.dll" /codebase
regasm "C:\Program Files\IT Group, Inc\ComUnion\Plugin\SendEmail.dll" /tlb

@echo off

echo DLL registration Successful...
echo

pause