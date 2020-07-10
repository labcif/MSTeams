# LabCIF - MSTeams

## Getting Started

Teams artifacts extraction providing results easy to be analyzed. This Autopsy Module has a **standalone version** - [GitHub](https://github.com/GuilhermeBeco/ProjetoEI) which can also be integrated with other apps.
## Functionalities

* Extract messages that user could interact
* Extract files that user could interact
* Extract reacts from users that our subject could interact
* Extract group calls where user could interact
* Extract private calls where user could interact
* Extract contacts that user could interact
* Extract teams formation where user could interact
* Extracted results can be analyzed on the blackboard
* Extracted results can be analyzed using HTML and/or CSV

### Results Screenshots

![Report Result Sample](https://i.imgur.com/R7Pmtf7.png)

![Blackboard Sample](https://i.imgur.com/mCWseew.png)

## Prerequisites

* [Autopsy](https://www.autopsy.com/download/) (4.15)

## How to use

This module can only be used with Autopsy

### Running this module
1. Download repository contents.
2. Open Autopsy -> Tools -> Python Plugins
3. Unzip previously downloaded zip in `python_modules` folder.
4. Restart Autopsy, create a case and select the module.
5. Select your module options in the Ingest Module window selector.
6. Analyze results at blackboard
7. In intended, Click "Generate Report" to generate a report containing HTML and CSV so you can analyze.

### Important note
* This Data Source Ingest Module only suports 3 types os data source:
1. Local Disk
2. Disk Image
3. Local Folder*
* *The local folder data source is recomended to be the location of Users folder, so you can get better results. If you dont want, your limit, on the directory tree must be %APPDATA%/Roaming
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
