@echo off
:: Stap 1: Wissel expliciet naar de S-schijf
S:

:: Stap 2: Ga naar de exacte projectmap
cd "S:\Projects\Handy's\Mail scraper"

:: Stap 3: Start het script
python mailscraper.py
pause