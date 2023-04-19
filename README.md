# CADD-Helper
The CADD-Helper is designed to facilitate the download of 3D SDF drug files for computer-aided drug design. Also it can accelerate ADMET analysis process of SwissADME and pkcsm. The app is user-friendly and enables efficient & accurate downloads for drug design research.

## Features

- SDF Download (IMPPAT, SwissSimilarity, Pubchem)
- ADMET accelerator (SwissADME, pkcsm) [pubchem]
- 2D Structure png download [pubchem]
- Retrieve Chemical name [pubchem]

![main_v1.5](v1.5.png)

## Installation
Download via cmd. Copy the code, paste it on CMD and hit enter.

```
cd Desktop & curl -OL https://github.com/sabbir-21/SDF-Protein-CADD/releases/download/v1.5/CADD_Helper_v1.5.exe
```
Or

Download from [Releases](https://github.com/sabbir-21/SDF-Protein-CADD/releases/latest)

[![Windows](https://img.shields.io/badge/-Windows_x64-blue.svg?style=for-the-badge&logo=windows)](https://github.com/sabbir-21/SDF-Protein-CADD/releases/latest)

- Its a portable Application. No need to install it. Just double click to run.
- This application is not free for all. You need to purchase it from me. Purchase licence key from [Buy Now](https://sabbir-21.github.io/portfolio/buy.html)


## Usage
Application startup time is around 5-15 seconds varying on your computer.
- Enter licence key to activate it lifetime.
- Choose `Option` from left side.
- Select excel sheet by pressing `Select Excel Sheet ` button. By default it will load `list.xlsx` if file exists. Excel sheet should contain `IMPPAT ID` or `Drugbank ID` or `Pubchem CID`.
- You may need to enter `Sheet Name`. By default sheet name will be provided as the option.
- You may need to enter `Column Name` which contains the ID. By default required value will be provided.
- Press the `Download SDF` or `Write SMILES` button. Download progress will be shown to the top of the button. After completing the task, `Completed` text will be shown.

## Demo Screenshot of Excel

![ss_imppat](ss_imppat.png)

![ss_swiss](ss_swiss.png)

![ss_pubchem](ss_pubchem.png)

## Problems
~~1. two word can't be used in xlsx file name. ex: "curcuma longa.xlsx" can't be used; instead use one word name like "curcuma.xlsx".~~

~~2. if xlsx files first cell is empty, will download an empty sdf file. You need to delete that empty file later.~~
1. App can't be opened offline. App is fully based on Online. Internet connection problem may `crash` the app.
2. Windows defender detects the app as malware. So you have to turn off windows defender while downloading the app. See details [here ](https://stackoverflow.com/questions/43777106/program-made-with-pyinstaller-now-seen-as-a-trojan-horse-by-avg) why this happens.

## Language
Built with  [![Windows](https://www.python.org/static/favicon.ico
)](https://www.python.org/)

## License
MIT

**Paid Software, Yeah!**
