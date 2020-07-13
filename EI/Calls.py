import os
import Contacts
from dateutil.tz import tz
from datetime import datetime
import re


class Chamada:
    def __init__(self, originator, timestart, timefinish, target, state):
        self.criador = originator
        self.timestart = timestart
        self.timefinish = timefinish
        self.presente = target
        self.state = state

    def toString(self):
        return "Start {0} || End: {1} || State: {2}|| Originator: {3} ||Target: {4}".format(self.timestart,
                                                                                            self.timefinish, self.state,
                                                                                            self.criador.toString(),
                                                                                            self.presente.toString())


class EventCall:
    def __init__(self, calldate, creator, count, duration, participants, orgids):
        self.calldate = calldate
        self.creator = creator
        self.count = count
        self.duration = duration
        self.participants = participants
        self.orgids = orgids

    def toString(self):
        return "Data: " + self.calldate + " (UTC) || Criador da chamada: " + self.creator.toString() + "|| Numero de " \
                                                                                                       "participantes: " + self.count + " || Duração da chamada (mins): " + self.duration + "|| Nomes dos " \
                                                                                                                                                                                            "participantes: " + ', '.join(
            self.participants) + " || Orgids: " + ', '.join(self.orgids)


def extrairEventCallsToFile(pathArmazenamento):
    # Ficheiro onde vão ser colocadas as informações das chamadas
    logEventCalls = open(os.path.join(pathArmazenamento, "eventCalls.txt"), "a+", encoding="utf-8")

    # abrir ficheiro logTotal para extrair informações de chamadas
    logFinalRead = open(os.path.join(pathArmazenamento, "logTotal.txt"), "r", encoding="utf-8")

    flagWrite = 0  # flag que indica quando escrever para o ficheiro
    buffer = ""
    lastline = ""
    # ler cada linha do ficheiro logTotal e procurar "Event/Call"
    while line := logFinalRead.readline():

        if flagWrite == 1:
            buffer = buffer + line.rstrip()
        if "Event/Call" in line and "count=" in line and "count=\"0\"" not in line:  # inicio de nova informação de
            # chamada
            flagWrite = 1
            logEventCalls.write(lastline)
            buffer = line.rstrip()
        if "</partlist>" in line and flagWrite == 1:  # fim da informação
            logEventCalls.write(buffer)
            logEventCalls.write("\n-----------------------------------------------"
                                "------------------------------------------------------------------\n")
            logEventCalls.write("----------------------------------------------------------------"
                                "-------------------------------------------------\n")
            flagWrite = 0
            buffer = ""

        lastline = line

    # fechar ficheiros
    logFinalRead.close()
    logEventCalls.close()


# funcao que cria objetos com a info das chamadas obtida na funcao 'extrairEventCallsToFile()'
def criarObjetosDeEventCalls(pathArmazenamento, arrayContactos):
    # abrir ficheiro eventCalls para extrair informações de chamadas
    logEventCalls = open(os.path.join(pathArmazenamento, "eventCalls.txt"), "r", encoding="utf-8")
    calldate = ""
    callcreator = ""
    c = Contacts.Contacto()
    arrayEventCall = []
    while line := logEventCalls.readline():
        marcador = 0

        if "orgid:" in line and calldate == "":

            # get the call end date
            indexcalldatefinal = line.find(",8:orgid")
            indexcalldate = indexcalldatefinal - 13

            lista = list(line)
            for i in range(indexcalldate, indexcalldatefinal - 3):
                calldate = calldate + lista[i]

            # tentar converter timestamp para timezone UTC
            try:
                calldate = int(calldate[:10])
                calldate = datetime.utcfromtimestamp(calldate)
                calldate = calldate.astimezone(tz=tz.tzlocal())
            except:
                calldate = calldate

            # get the user who started the call
            indexcallcreator = line.find(",8:orgid:")
            indexcallcreatorfinal = indexcallcreator + 9 + 36  # 36 is the number of chars in orgid
            for i in range(indexcallcreator + 9, indexcallcreatorfinal):
                callcreator = callcreator + lista[i]

            # check if orgid exists in the contacts list
            if callcreator in arrayContactos:
                c = arrayContactos.get(callcreator)
            else:
                c = Contacts.Contacto('Unknown', 'Unknown', callcreator)

        if "Event/Call" in line:
            # ir buscar numero total de participantes na chamada
            participantscount = ""

            indexparticipantscount = line.find("count=\"")
            indexparticipantscountfinal = line.find("\"><part")

            lista = list(line)
            for i in range(indexparticipantscount, indexparticipantscountfinal):
                if "\"" in lista[i]:
                    marcador = 1
                    continue
                if marcador:
                    participantscount = participantscount + lista[i]

            # ir buscar duração da chamada
            callduration = ""

            indexcallduration = line.find("<duration>")
            indexcalldurationfinal = line.find("</duration>", indexcallduration)

            lista = list(line)
            for i in range(indexcallduration + 10, indexcalldurationfinal):
                callduration = callduration + lista[i]

            callduration = str("{:.1f}".format(int(callduration) / 60))
            # testar se está certo pois houve erro e tive de acrescentar o indexcallduration no find
            # ir buscar todos os nomes dos participantes
            participantstag = [m.start() for m in re.finditer('<displayName>', line)]
            participantsendtag = [m.start() for m in re.finditer('</displayName>', line)]

            participants = []

            for begin, end in zip(participantstag, participantsendtag):
                participantname = ""
                lista = list(line)
                for i in range(begin + 13, end):
                    participantname = participantname + lista[i]

                participants.append(participantname)

            # ir buscar o orgid dos participantes
            orgidtag = [m.start() for m in re.finditer('<name>', line)]
            orgidendtag = [m.start() for m in re.finditer('</name>', line)]

            orgids = []

            for begin, end in zip(orgidtag, orgidendtag):
                orgid = ""
                lista = list(line)
                for i in range(begin + 14, end):
                    orgid = orgid + lista[i]

                orgids.append(orgid)

            # criar novo objeto EventCall
            eventcall = EventCall(str(calldate), c, participantscount, callduration, participants, orgids)
            arrayEventCall.append(eventcall)

            calldate = ""
            callcreator = ""

    # fechar ficheiros
    logEventCalls.close()
    return arrayEventCall
