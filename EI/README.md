# LabCIF - MSTeams

## Getting Started

Teams artifacts extraction providing results easy to be analyzed. This executable cam be executed as a standalone, as an **Autopsy Module** - [GitHub](https://github.com/labcif/MSTeams) or it can be integrated with other apps.
## Functionalities

* Extract messages that user could interact
* Extract files that user could interact
* Extract reacts from users that our subject could interact
* Extract group calls where user could interact
* Extract private calls where user could interact
* Extract contacts that user could interact
* Extract teams formation where user could interact

### Results Screenshots

![Report Index](https://i.imgur.com/LCXpt6N.png)

![Report Sample](https://i.imgur.com/TZlQbiV.png)

## Prerequisites

* [Python](https://www.python.org/downloads/) (3.7+)

## How to use

The script can be used directly in terminal or it can be integrated with other apps.

### Running from Terminal

```bash
usage: EI.exe [-h] [-u,--users PathToUsersFolder] or EI.exe [-u,--users PathToUsersFolder] or EI.exe [--pathToEI PathToEiFolder] [-a folderNameToAnalyze]

MSTeams

optional arguments:
  -h, --help                                     show this help message and exit
  -u, --users                                    Indicates the path to users folder (eg: C:\Users)
  -a                                             Indicates the folder name located at %APPDATA%/Roaming/ProjetoEI. This folder is a copy of MSTeams Log LevelDB folder
  --pathToEI                                     Indicates EI folder location
```

### Integrating with other apps

When integrating with other apps, you only should use the last combination os parameters (EI.exe [--pathToEI PathToEiFolder] [-a folderNameToAnalyze]). 
You must create a copy of MSTeams Log LevelDB folder to %APPDATA%/Roaming/ProjetoEI.
Once you done that, you need to execute a subprocess where you call EI.exe, indicating the pathToEI (where you intended it's location) and the folder to analyze.
## Authors

* **Guilherme Beco** - [GitHub](https://github.com/GuilhermeBeco)
* **Tiago Garcia** - [GitHub](https://github.com/tiagohgarcia)

## Mentors

* **Miguel Frade** - [GitHub](https://github.com/mfrade)
* **Patrício Domingues** - [GitHub](https://github.com/PatricioDomingues)

Project developed as final project for Computer Engineering course in Escola Superior de Tecnologia e Gestão de Leiria.

## Environments Tested

* Windows

## License

This project is licensed under the terms of the GNU GPL v3 License.

* [golang/leveldb](https://github.com/golang/leveldb) - BSD 3-Clause "New" or "Revised" License
