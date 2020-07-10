import getopt
import csv
import math
import os
import sys
import time
from datetime import datetime
from pathlib import Path
import Messages
import Contacts
import Teams
import LogBuilder
import Calls
import ProduceHTML
import WriteCSV

appdata = os.getenv('APPDATA')
home = str(Path.home())
levelDBPath = appdata + "\\Microsoft\\Teams\\IndexedDB\\https_teams.microsoft.com_0.indexeddb.leveldb"
projetoEIAppDataPath = appdata + "\\ProjetoEI\\"
user = ""
arrayMensagens = []
arrayContactos = {}
arrayCallOneToOne = []
dictionaryConversationDetails = {}
arrayUsers = []
dictFiles = {}
idM = 1
pathModule = ""
tmCSV = ""

if __name__ == "__main__":
    args = sys.argv[1:]
    pathUsers = ""
    pathAppdata = ""
    arrayReturn = {}
    current_milli_time = lambda: datetime.fromtimestamp(math.floor(time.time())).strftime("%m-%d-%Y, %Hh_%Mm_%Ss")
    dupped = False
    try:
        opts, args = getopt.getopt(args, "hpua:", ["users=", "autopsy=", "pathToEI="])
    except getopt.GetoptError:
        print('ei.py -u <pathToUsers> ')
        print("or")
        print('ei.py --pathToEI <pathToEIFolder> -a <LDBFolderNameInProjetoEIAppData>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == "--pathToEI":
            ok = False
            for o, v in opts:
                if o == "-a":
                    ok = True
            if not ok:
                sys.exit(
                    'ei.py --pathToEI <pathToEIFolder> -a <LDBFolderNameInProjetoEIAppData> or ei.py -u <pathToUsers> ')
            pathModule = arg
            pathModule = pathModule.replace("\"", "\\")
            pathModule = pathModule.rstrip()
            pathModule = pathModule.lstrip()
            print(pathModule)
        if opt == '-h':
            print('ei.py -u <pathToUsers> ')
            print("or")
            print('ei.py --pathToEI <pathToEIFolder> -a <LDBFolderNameInProjetoEIAppData>')
            sys.exit()
        elif opt in ("-u", "--users"):
            pathModule = os.path.realpath(__file__)
            indexCutPath = pathModule.rfind("\\")
            pathModule = pathModule[0:indexCutPath + 1]
            print(arg)
            count = 0
            pathMulti = ""
            pathUsers = arg
            for root, dirs, files in os.walk(pathUsers):
                for name in dirs:
                    if "https_teams.microsoft.com" in name:
                        for user in arrayUsers:
                            if name == user:
                                dupped = True
                        if not dupped:
                            pathSplited = os.path.join(root, name).split("\\")
                            user = pathSplited[2]
                            arrayUsers.append(os.path.join(root, name))
                            levelDBPath = os.path.join(root, name)
                            dupped = False
                            if not os.path.exists(projetoEIAppDataPath):
                                try:
                                    os.mkdir(projetoEIAppDataPath)
                                except OSError:
                                    print("Creation of the directory %s failed" % projetoEIAppDataPath)
                                else:
                                    print("Successfully created the directory %s " % projetoEIAppDataPath)
                            try:
                                pathMulti = projetoEIAppDataPath + "Analise standalone {} {}".format(user,
                                                                                                     current_milli_time()) + "\\"
                                os.mkdir(pathMulti)
                            except OSError:
                                print("Creation of the directory %s failed" % pathMulti)
                            else:
                                print("Successfully created the directory %s " % pathMulti)
                            tmCSV = current_milli_time()
                            with open(os.path.join(pathMulti, 'Reacts_{}_{}.csv'.format(user, tmCSV)), 'a+', newline='',
                                      encoding="utf-8") as csvfile:
                                fieldnames = ['messageID', 'reacted_with', 'reacted_by', 'react_time', 'user']
                                messagewriter = csv.writer(csvfile, delimiter=';', quotechar='|',
                                                           quoting=csv.QUOTE_MINIMAL)
                                messagewriter.writerow(fieldnames)
                                csvfile.close()

                            LogBuilder.crialogtotal(pathMulti, pathModule, levelDBPath)

                            arrayContactos = Contacts.geraContactos(pathMulti)

                            Teams.criacaoDeEquipas(pathMulti)
                            arrayReturn = Teams.criarObjetosDeCriacaoDeEquipas(pathMulti, arrayContactos)
                            arrayContactos = arrayReturn["contacts"]
                            dictionaryConversationDetails = arrayReturn["teams"]
                            arrayReturn.clear()

                            Calls.extrairEventCallsToFile(pathMulti)
                            arrayEventCall = Calls.criarObjetosDeEventCalls(pathMulti, arrayContactos)

                            arrayReturn = Messages.findpadrao(pathMulti, arrayContactos, user, tmCSV)
                            arrayMensagens = arrayReturn["mensagem"]
                            arrayCallOneToOne = arrayReturn["callOneToOne"]

                            WriteCSV.writeFiles(user, pathMulti, arrayContactos, arrayMensagens, arrayEventCall,
                                                dictionaryConversationDetails, arrayCallOneToOne, tmCSV)
                            ProduceHTML.createhtmltables(pathMulti, user, arrayContactos, arrayEventCall,
                                                         arrayCallOneToOne,
                                                         dictionaryConversationDetails, current_milli_time(), tmCSV)

                            arrayContactos.clear()
                            dictFiles.clear()
                            arrayMensagens.clear()
                            arrayEventCall.clear()
                            dictionaryConversationDetails.clear()
                            arrayCallOneToOne.clear()
                            idM = 1

        elif opt in ("-a", "--autopsy"):
            if pathModule == "":
                sys.exit('ei.py --pathToEI <pathToEIFolder> -a <LDBFolderNameInProjetoEIAppData>')
            idMessage = 1
            print(arg)
            levelDBPath = projetoEIAppDataPath + "\\" + arg
            pathSplited = arg.split("_")
            user = pathSplited[3]
            pathToAutopsy = projetoEIAppDataPath + "\\Analise Autopsy {} {}".format(user, str(current_milli_time()))
            if not os.path.exists(projetoEIAppDataPath):
                try:
                    os.mkdir(projetoEIAppDataPath)
                except OSError:
                    print("Creation of the directory %s failed" % projetoEIAppDataPath)
                else:
                    print("Successfully created the directory %s " % projetoEIAppDataPath)
            try:
                os.mkdir(pathToAutopsy)
            except OSError:
                print("Creation of the directory %s failed" % pathToAutopsy)
            else:
                print("Successfully created the directory %s " % pathToAutopsy)
            tmCSV = current_milli_time()
            with open(os.path.join(pathToAutopsy, 'Reacts_{}_{}.csv'.format(user, tmCSV)), 'a+', newline='',
                      encoding="utf-8") as csvfile:
                fieldnames = ['messageID', 'reacted_with', 'reacted_by', 'react_time', 'user']
                messagewriter = csv.writer(csvfile, delimiter=';', quotechar='|',
                                           quoting=csv.QUOTE_MINIMAL)
                messagewriter.writerow(fieldnames)
                csvfile.close()

            LogBuilder.crialogtotal(pathToAutopsy, pathModule, levelDBPath)
            arrayContactos = Contacts.geraContactos(pathToAutopsy)

            Teams.criacaoDeEquipas(pathToAutopsy)
            arrayReturn = Teams.criarObjetosDeCriacaoDeEquipas(pathToAutopsy, arrayContactos)
            arrayContactos = arrayReturn["contacts"]
            dictionaryConversationDetails = arrayReturn["teams"]
            arrayReturn.clear()

            Calls.extrairEventCallsToFile(pathToAutopsy)
            arrayEventCall = Calls.criarObjetosDeEventCalls(pathToAutopsy, arrayContactos)

            arrayReturn = Messages.findpadrao(pathToAutopsy, arrayContactos, user, tmCSV)
            arrayMensagens = arrayReturn["mensagem"]
            arrayCallOneToOne = arrayReturn["callOneToOne"]

            WriteCSV.writeFiles(user, pathToAutopsy, arrayContactos, arrayMensagens, arrayEventCall,
                                dictionaryConversationDetails, arrayCallOneToOne, tmCSV)

            ProduceHTML.createhtmltables(pathToAutopsy, user, arrayContactos, arrayEventCall, arrayCallOneToOne,
                                         dictionaryConversationDetails, current_milli_time(), tmCSV)

