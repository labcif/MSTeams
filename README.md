# LabCIF - MSTeams

## Getting Started

Teams artifacts extraction providing results easy to be analyzed. This Autopsy Module also contains a **Standalone Version** - [GitHub](https://github.com/GuilhermeBeco/ProjetoEI).
## Functionalities

* Extract messages that user could interact
* Extract files that user could interact
* Extract reacts from users that our subject could interact
* Extract group calls where user could interact
* Extract private calls where user could interact
* Extract contacts that user could interact
* Extract teams formation where user could interact
* Creation of blackboard entries to analyze artifacts
* Creation of a report containing the results files so they can be analyzed as a CSV file or an HTML file

### Results Screenshots

![BlackBoard](https://imgur.com/mCWseew)

![Report(https://imgur.com/R7Pmtf7)

## Prerequisites

* [Autopsy](https://www.autopsy.com/download/) (4.15)

## How to use
1. Download repository contents.
2. Open Autopsy -> Tools -> Python Plugins
3. Unzip previously downloaded zip in `python_modules` folder.
4. Restart Autopsy, create a case and select the module.
5. Select your module options in the Ingest Module window selector.
6. Analyze your results that are on the blackboard
6. If you intend to analyze the results using HTML or CSV, click "Generate Report", select 'MSTeams Report Module' and your report will be generated.

* This Data Source Ingest Module can only be execute using three types of data source. Disk image, local disk or local folders. One thing to keep in mind is that when using local folders it's recomended to use the Users folder to use this module to it's full potencial. If thats not your intention, your limit must be %APPDATA%\Roaming.


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
