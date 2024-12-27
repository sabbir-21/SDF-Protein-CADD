@echo off 

::: 
:::    _____       __    __    _         ___    __                       __
:::   / ___/____ _/ /_  / /_  (_)____   /   |  / /_  ____ ___  ___  ____/ /
:::   \__ \/ __ `/ __ \/ __ \/ / ___/  / /| | / __ \/ __ `__ \/ _ \/ __  / 
:::  ___/ / /_/ / /_/ / /_/ / / /     / ___ |/ / / / / / / / /  __/ /_/ /  
::: /____/\__,_/_.___/_.___/_/_/     /_/  |_/_/ /_/_/ /_/ /_/\___/\__,_/   
:::                                                                       
:::
for /f "delims=: tokens=*" %%A in ('findstr /b ::: "%~f0"') do @echo(%%A

echo Downloading CADD_Helper version 1.11

curl -OL https://github.com/sabbir-21/SDF-Protein-CADD/releases/download/v1.11/CADD_Helper_v1.11_Windows.exe & CADD_Helper_v1.11_Windows.exe

pause