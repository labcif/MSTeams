import codecs
import getopt
import csv
import math
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
import pandas as pd
import zulu as zulu
from dateutil.tz import tz

from models import Contacto, MensagemCompleta, Reaction, File, Chamada, ConversationCreationDetails, EventCall

appdata = os.getenv('APPDATA')
home = str(Path.home())
levelDBPath = appdata + "\\Microsoft\\Teams\\IndexedDB\\https_teams.microsoft.com_0.indexeddb.leveldb"
levelDBPathLost = appdata + "\\Microsoft\\Teams\\IndexedDB\\https_teams.microsoft.com_0.indexeddb.leveldb\\lost"
projetoEIAppDataPath = appdata + "\\ProjetoEI\\"

wroteReacts = False
linkEmoji = "https://statics.teams.cdn.office.net/evergreen-assets/skype/v2/{0}/50.png"
arrayRich = []
arrayMensagens = []
arrayContactos = {}
bufferBuffer = []
arrayReacoes = []
arrayCallOneToOne = []
arrayTeams = []
arrayEventCall = []
dictionaryConversationDetails = {}
arrayUsers = []
dictFiles = {}
dictReacts = {}
idM = 1
pathModule = ""


def find(pt, path):
    fls = []
    for root, dirs, files in os.walk(path):
        for name in files:
            if pt in name:
                fls.append(os.path.join(root, name))
    return fls


def acentuar(acentos, st):
    for hexa, acento in acentos.items():
        st = st.replace(hexa, acento)
    return st


def decoder(string, pattern):
    s = string.encode("utf-8").find(pattern.encode("utf-16le"))
    l = string[s - 1:]
    t = l.encode("utf-16le")
    st = str(t)
    st = st.replace(r"\x00", "")
    st = st.replace(r"\x", "feff00")
    st = st.replace(r"b'", "")
    return st


def multiFind(text):
    f = ""
    index = 0
    a = ""
    l = list(text)
    acentos = {}
    while index < len(text):
        index = text.find('feff00', index)
        if index == -1:
            break
        for x in range(index, index + 8):
            f = f + l[x]
        a = bytes.fromhex(f).decode("utf-16be")
        acentos[f] = a
        index += 8  # +6 because len('feff00') == 2
        f = ""
    return acentos


def utf16customdecoder(string, pattern):
    st = decoder(string, pattern)
    acentos = multiFind(st)
    st = acentuar(acentos, st)
    # print(st)
    return st


def testeldb(path, pathArmazenamento):
    logFinal = open(os.path.join(pathArmazenamento, "logTotal.txt"), "a+", encoding="utf-8")
    from subprocess import Popen, PIPE
    import ast
    for f in os.listdir(path):
        if f.endswith(".ldb") and os.path.getsize(path + "\\" + f) > 0:
            process = Popen(["{}ldbdump".format(pathModule), os.path.join(path, f)], stdout=PIPE)
            (output, err) = process.communicate()
            # try:
            result = output.decode("utf-8")
            for line in (result.split("\n")[1:]):
                if line.strip() == "": continue
                parsed = ast.literal_eval("{" + line + "}")
                key = list(parsed.keys())[0]
                logFinal.write(parsed[key])
            # except:
            #     print("falhou")


# --pathToEI C:\Users\Beco\PycharmProjects\EI\ -a Analysis_Autopsy_LDB_06-15-2020_09h-58m-39s

def writelog(path, pathArmazenamento):
    logFinal = open(os.path.join(pathArmazenamento, "logTotal.txt"), "a+", encoding="utf-8")
    log = open(path, "r", encoding="mbcs")
    while line := log.readline():
        logFinal.write(line)
    log.close()


def crialogtotal(pathArmazenamento):
    # os.system(r'cmd /c "LDBReader\PrjTeam.exe"')
    logFinal = open(os.path.join(pathArmazenamento, "logTotal.txt"), "a+", encoding="utf-8")
    os.system('cmd /c "{}LevelDBReader\\eiTeste.exe "{}" "{}"'.format(pathModule, levelDBPath, pathArmazenamento))
    testeldb(levelDBPath + "\\lost", pathArmazenamento)
    logLost = find(".log", levelDBPath + "\\lost")
    writelog(os.path.join(pathArmazenamento, "logIndexedDB.txt"), pathArmazenamento)
    logsIndex = find(".log", levelDBPath)
    for path in logLost:
        writelog(path, pathArmazenamento)
    for path in logsIndex:
        writelog(path, pathArmazenamento)
    logFinal.close()


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
        if "Event/Call" in line and "count=" in line and "count=\"0\"" not in line:  # inicio de nova informação de chamada
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
def criarObjetosDeEventCalls(pathArmazenamento):
    # abrir ficheiro eventCalls para extrair informações de chamadas
    logEventCalls = open(os.path.join(pathArmazenamento, "eventCalls.txt"), "r", encoding="utf-8")
    calldate = ""
    callcreator = ""
    c = Contacto()

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
                c = Contacto('Unknown', 'Unknown', callcreator)

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


def criacaoDeEquipas(pathArmazenamento):
    # ficheiro onde vão ser colocadas as informações da criação de equipas
    logTeamsCreation = open(os.path.join(pathArmazenamento, "teamsCreation.txt"), "a+", encoding="utf-8")

    # abrir ficheiro logTotal para extrair informações de criacao das equipas
    logFinalRead = open(os.path.join(pathArmazenamento, "logTotal.txt"), "r", encoding="utf-8")

    flagWrite = 0  # flag que indica quando escrever para o ficheiro
    count = 1
    # ler cada linha do ficheiro logTotal e procurar "<addmember>"
    while line := logFinalRead.readline():

        if "messageso" in line and "<addmember>" in line and "<eventtime>" in line and "<initiator>" in line and count:  # inicio de nova informação de criacao de equipa # ThreadActivity/AddMember
            flagWrite = 1
            count = 0
            logTeamsCreation.write(line)
            continue
        if "_emailDetails_" in line and flagWrite == 1 or "messageso" in line and flagWrite == 1:  # fim da informação
            # logTeamsCreation.write(line)
            logTeamsCreation.write("\n-----------------------------------------------"
                                   "------------------------------------------------------------------\n")
            logTeamsCreation.write("----------------------------------------------------------------"
                                   "-------------------------------------------------\n")
            flagWrite = 0
            count = 1
            continue
        if flagWrite == 1:
            logTeamsCreation.write(line)

    # fechar ficheiros
    logFinalRead.close()
    logTeamsCreation.close()


def cleanMessage(message):
    p = ""
    span = ""
    r = ""
    while True:
        m = list(message)
        indexP = message.find("<p")
        indexPFinal = message.find(">", indexP)
        indexSpan = message.find("<span")
        indexspanFinal = message.find(">", indexSpan)

        if indexP == -1 and indexSpan == -1:
            break
        if indexP != -1:
            for x in range(indexP, indexPFinal + 1):
                p += m[x]

        if indexSpan != -1:
            for x in range(indexSpan, indexspanFinal + 1):
                span += m[x]
        # print(span)
        message = message.replace(p, "")
        message = message.replace(span, "")
        p = ""
        span = ""

    indexRich = message.find("RichText")
    if indexRich != -1:
        indexRichFinal = message.find("contentc", indexRich)
        for x in range(indexRich, indexRichFinal + 9):
            r += m[x]
        message = message.replace(r, "")
    message = message.replace("<﻿Ø﻿û﻿ß", "")
    message = message.replace(r"/n", "")
    message = message.replace("<strong>", "")
    message = message.replace("</strong>", "")
    message = message.replace("<em>", "")
    message = message.replace("</em>", "")
    message = message.replace("</span>", "")
    message = message.replace("</pre>", "")
    message = message.replace("<code>", "")
    message = message.replace("</code>", "")

    return message


def criarObjetosDeCriacaoDeEquipas(pathArmazenamento):
    # ficheiro para guardar objetos de criacao de conversacoes/Equipas

    # abrir ficheiro com a informacao das equipas para criar objetos
    logTeamsCreation = open(os.path.join(pathArmazenamento, "teamsCreation.txt"), "r", encoding="utf-8")

    conversationid = ""
    date = ""
    initiatororgid = ""
    targets = []

    flag1 = 0
    flag2 = 0
    first = 1

    while line := logTeamsCreation.readline():

        # ir buscar id da conversacao e data de criacao
        if "messageso\"" in line and first:
            flag1 = 1
            first = 0

            # ir buscar id da conversacao
            conversationid = ""

            indexconverstionid = line.find("messageso\"")
            indexconverstionidfinal = line.find("@")

            lista = list(line)
            for i in range(indexconverstionid + 25, indexconverstionidfinal):
                conversationid = conversationid + lista[i]

            # ir buscar data de criacao da conversacao
            date = ""

            timestamp = line[:10]
            # tentar converter timestamp para timezone UTC
            try:
                date = datetime.utcfromtimestamp(int(timestamp))
                date = date.astimezone(tz=tz.tzlocal())
            except:
                date = timestamp

        # ir buscar membros adicionados a conversacao e quem criou a conversacao
        if "renderContent" in line:
            flag2 = 1

            # ir buscar orgid de membro(s) adicionado(s)
            targettag = [m.start() for m in re.finditer('<target>', line)]
            targettagend = [m.start() for m in re.finditer('</target>', line)]

            targets = []

            for begin, end in zip(targettag, targettagend):
                targetorgid = ""

                lista = list(line)
                for i in range(begin + 16, end):
                    targetorgid = targetorgid + lista[i]

                targets.append(targetorgid)

            # ir buscar orgid de utilizador que criou a conversa/Team
            initiatororgid = ""

            indexinitiator = line.find("<initiator>")
            indexinitiatorfinal = line.find("</initiator>")

            for i in range(indexinitiator + 19, indexinitiatorfinal):
                initiatororgid = initiatororgid + lista[i]

        # if "department" in line:
        #     print("+++++++++++++++++++++")#########
        #     department = ""
        #
        #     indexdepartment = line.find("department")
        #     indexdepartmentfinal = line.find("accountEnabled")
        #
        #     lista = list(line)
        #     for i in range(indexdepartment+12, indexdepartmentfinal-2):
        #         if "\"" in lista[i]:
        #             break;
        #
        #         department = department + lista[i]
        #
        #
        #     if "physicalDeliveryOfficeName" in line:
        #         office = ""
        #
        #         indexoffice = line.find("physicalDeliveryOfficeName")
        #         indexofficefinal = line.find("accountEnabled")
        #
        #         for i in range(indexoffice + 28, indexofficefinal - 2):
        #             if "\"" in lista[i]:
        #                 break
        #             office = office + lista[i]
        #
        #
        #     if "userLocation" in line:
        #         location = ""
        #
        #         indexlocation  = line.find("userLocation")
        #         indexlocationfinal = line.find("accountEnabled")
        #
        #         for i in range(indexlocation + 14, indexlocationfinal - 2):
        #             location = location + lista[i]
        #

        # Utilizado para atualizar lista de contactos (caso contacto ainda não exista)
        # informacao de utilizadores (criador da conversa ou membro adicionado)
        if "isSipDisabled" in line:
            # if "jobTitle\"" in line:
            #     job = ""
            #
            #     indexjob = line.find("jobTitle\"")
            #     indexjobfinal = line.find("email\"")
            #
            #     lista = list(line)
            #     for i in range(indexjob + 10, indexjobfinal - 2):
            #         job = job + lista[i]

            if "email\"" in line:
                email = ""

                indexemail = line.find("email\"")
                indexemailfinal = line.find("userType\"")

                lista = list(line)
                for i in range(indexemail + 7, indexemailfinal - 2):
                    email = email + lista[i]

            # if "userType\"" in line:
            #     usertype = ""
            #
            #     indexusertype = line.find("userType\"")
            #     indexusertypefinal = line.find("displayName\"")
            #
            #     for i in range(indexusertype + 10, indexusertypefinal - 2):
            #         usertype = usertype + lista[i]

            if "displayName\"" in line:
                username = ""

                indexusername = line.find("displayName\"")
                subline = line[indexusername:]

                indexusername = subline.find("displayName\"")
                indexusernamefinal = subline.find("type\"")

                lista = list(subline)
                for i in range(indexusername + 13, indexusernamefinal - 2):
                    username = username + lista[i]

            if ":orgid:" in line:
                orgid = ""

                indexorgid = line.find(":orgid:")
                indexorgidfinal = line.find("objectId\"")

                lista = list(line)
                for i in range(indexorgid + 7, indexorgidfinal - 2):
                    orgid = orgid + lista[i]

                # ATUALIZAR LISTA DE CONTACTOS
                # verificar se contacto ja existe no array de contactos
                # se nao existir, entao adiciona
                if orgid not in arrayContactos:
                    contacto = Contacto(username, email, orgid)
                    arrayContactos[orgid] = contacto

        if flag1 and flag2:
            flag1 = 0
            flag2 = 0
            first = 1
            # criar objeto de ConversationCreationDetails()
            conversation_details = ConversationCreationDetails(conversationid, str(date), initiatororgid, targets)

            if str(date) not in dictionaryConversationDetails:
                dictionaryConversationDetails[str(date)] = conversation_details

    # ir buscar contacto ao criador e membros das conversacao
    for conversa in dictionaryConversationDetails.values():
        # ir buscar contacto do criador da conversacao
        if conversa.creator in arrayContactos:
            creator = arrayContactos[conversa.creator]
            conversa.creator = creator
        if isinstance(conversa.creator, str):
            oID = conversa.creator
            conversa.creator = Contacto('Unknown', 'Unknown', oID)

        # ir buscar contacto do(s) membro(s) a ser(em) adicionado(s)
        members = []
        for user_orgid in conversa.members:
            if user_orgid in arrayContactos:
                member = arrayContactos[user_orgid]
                members.append(member)
            else:
                us = Contacto("Desc.", "Desc.", user_orgid)
                members.append(us)
        conversa.members = members.copy()

    logTeamsCreation.close()


def filtro(buffer):
    filtros = {
        "messageStorageState": 0,
        "parentMessageId": 0,
        "composetime": 0
    }
    global idM
    sender = ""
    time = ""
    indexTexto = 0
    indexTextoFinal = 0
    name = ""
    isRichHtml = False
    isRichHtmlDone = False
    richHtml = ""
    reacaoChunk = ""
    isReacaoChunk = False
    hasFiles = False
    emojiReaction = ""
    orgidReaction = ""
    timeReaction = ""
    cvId = ""
    message = ""
    global wroteReacts
    elogio = ""
    adptiveCard = False
    reactions = []
    reacaoChunkReady = False
    files = []
    ok = 1
    link = ""
    hrefLink = ""
    nAspas = False
    pathReacts = ""

    for filter in filtros:
        for line in buffer:
            if filter in line:
                filtros[filter] = 1
    for y, x in filtros.items():
        if x == 0:
            ok = 0
    if ok:

        isMention = False
        mention = ""
        orgid = ""
        nome = ""

        existe = False
        for line in buffer:
            name = ""
            if "conversationId\"[" in line or "@thread.tacv2" in line:
                indexCvId = line.find("conversationId\"")
                indexCvIdFinal = line.find(".spaces", indexCvId)
                if indexCvIdFinal == -1:
                    indexCvIdFinal = line.find("d.tacv2", indexCvId)
                l = list(line)
                for x in range(indexCvId + 16, indexCvIdFinal + 7):
                    cvId += l[x]
            if "imdisplayname" in line:
                indexTexto = line.find("imdisplayname")
                indexTextoFinal = line.find("\"", indexTexto + 16)
                l = list(line)
                for x in range(indexTexto + 15, indexTextoFinal):
                    name += l[x]
                sender = name
            name = ""

            if "creator" in line:
                indexOrgid = line.find("creator\",8:orgid:")
                indexOrgid += 17
                indexOrgidFinal = line.find("\"", indexOrgid)
                l = list(line)
                org = ""
                for x in range(indexOrgid, indexOrgidFinal):
                    org += l[x]
                orgid = org

            if "composetime" in line:
                indexTempo = line.find("originalarrivaltime")
                indexTempoFinal = line.find("clientArrivalTime", indexTempo)

                l = list(line)
                for x in range(indexTempo + 21, indexTempoFinal - 2):
                    name = name + l[x]
                name = name.replace("\"", "")

                try:
                    dtStart = zulu.parse(name)
                    time = datetime.utcfromtimestamp(dtStart.timestamp())
                    time = time.astimezone(tz=tz.tzlocal())
                    time = name
                except:
                    time = "sem tempo"

                nAspas = False

            if "RichText/Html" in line:
                isRichHtml = True
                richHtml += line

            if "renderContent" in line:
                isRichHtml = False
                isRichHtmlDone = True

            if isRichHtml:
                richHtml += line

            if isRichHtmlDone:
                if "<div>".encode("utf-16le") in line.encode("utf-8"):
                    st = utf16customdecoder(richHtml, "<div>")
                    richHtml = st
                    if "https://statics.teams.cdn.office.net/evergreen-assets/skype/v2/" in st:
                        indexMultiplosEmojis = 0
                        indexSpan = 0
                        while indexSpan != -1:
                            indexSpan = st.find("<span class=\"animated-emoticon", indexMultiplosEmojis)
                            indexMultiplosEmojis = indexSpan
                            indexSpanfinal = st.find("</span>", indexSpan)
                            indexSpanfinal += 7
                            l = list(st)
                            span = ""
                            emoji = ""
                            for x in range(indexSpan, indexSpanfinal):
                                span = span + l[x]

                            indexEmoji = span.find("itemid=\"")
                            indexEmojiFinal = span.find(" itemscope")
                            indexEmoji += 7
                            spanList = list(span)
                            for x in range(indexEmoji + 1, indexEmojiFinal - 1):
                                emoji += spanList[x]
                            st = st.replace(span, " " + linkEmoji.format(emoji) + " ")
                            st = st.replace(r"\n", "")
                            richHtml = st
                if "https://statics.teams.cdn.office.net/evergreen-assets/emojioneassets/" in richHtml:
                    indexEmojione = richHtml.find("<img")
                    indexEmojioneFinal = richHtml.find("/>", indexEmojione)
                    l = list(richHtml)
                    emojiReplace = ""
                    for x in range(indexEmojione, indexEmojioneFinal + 2):
                        emojiReplace += l[x]
                    indexLinkEmoji = emojiReplace.find("src")
                    indexLinkEmojiFinal = emojiReplace.find("width", indexLinkEmoji)
                    emojiOne = ""
                    le = list(emojiReplace)
                    for x in range(indexLinkEmoji + 5, indexLinkEmojiFinal - 3):
                        emojiOne += le[x]
                    richHtml = richHtml.replace(emojiReplace, emojiOne)
                    richHtml = richHtml[richHtml.find(r"\n"):richHtml.__len__()]
                if "http://schema.skype.com/Mention" in richHtml:
                    indexMencionado = 0
                    span = ""
                    while indexMencionado != -1:
                        # obter pelo orgid
                        isMention = True
                        indexMencionado = richHtml.find("<span itemscope itemtype=\"http://schema.skype.com/Mention\"",
                                                        indexMencionado)
                        indexMencionadoFinal = richHtml.find("</span>", indexMencionado)
                        l = list(richHtml)
                        for x in range(indexMencionado, indexMencionadoFinal):
                            span += l[x]
                        indexSpanMention = span.find("itemid") + 11
                        s = list(span)
                        for x in range(indexSpanMention, len(s)):
                            name += s[x]
                        richHtml = richHtml.replace(span + "</span>", "Mention: {0}".format(name))
                        span = ""
                        mention = name
                        name = ""
                if "http://schema.skype.com/AMSImage" in richHtml:
                    l = list(richHtml)
                    src = " "
                    img = ""
                    indexSrc = richHtml.find("src=\"")
                    indexSrcFinal = richHtml.find("\" width")
                    indexImg = richHtml.find("<img")
                    indexImgFinal = richHtml.find("/>", indexImg)
                    for x in range(indexSrc + 5, indexSrcFinal):
                        src += l[x]
                    for x in range(indexImg, indexImgFinal + 2):
                        img += l[x]
                    src += " "
                    indexRH = richHtml.find(r"RichText/Html")
                    replaceRH = ""
                    for x in range(0, indexRH):
                        replaceRH += l[x]
                    richHtml = richHtml.replace(replaceRH, "", 1)
                    richHtml = richHtml.replace(img, src)
                    richHtml = richHtml.replace(r"RichText/Html", "")
                    richHtml = richHtml.replace(r"﻿contenttype", "")
                    richHtml = richHtml.replace(r"﻿contentc﻿", "")
                    richHtml = richHtml.replace(r"﻿text", "")
                    richHtml = richHtml.replace("<span>", "")
                    richHtml = richHtml.replace("<div>", "")
                    richHtml = richHtml.replace("</span>", "")
                    richHtml = richHtml.replace("</div>", "")
                    richHtml = richHtml.replace(r"\n", "")
                    richHtml = richHtml.replace("\"", "")
                    richHtml = richHtml[3::]

                    richHtml = "RichText<div>" + richHtml
                    richHtml += "</div>"

                if "http://schema.skype.com/InputExtension\"><span itemprop=\"cardId\"" in richHtml:
                    indexAdaptiveCard = richHtml.find("<span itemid")
                    indexAdaptiveCardFinal = richHtml.find("</div></div></div></div></div></div></div>",
                                                           indexAdaptiveCard)
                    spanCard = ""
                    l = list(richHtml)
                    for x in range(indexAdaptiveCard, indexAdaptiveCardFinal + 35):
                        spanCard += l[x]
                    richHtml = richHtml.replace(spanCard, "A praise was given to ")
                    adptiveCard = True
                    richHtml = richHtml.replace("<div>", "")
                    richHtml = richHtml.replace("</div>", "")
                    richHtml = "RichText<div>" + richHtml
                    richHtml += "</div>"

                if richHtml.find("RichText") != -1:
                    rHtml = ""
                    richHtml = richHtml.replace("<div style=\"font-size:14px;\"", "")
                    l = list(richHtml)
                    indexTexto = richHtml.find("<div")
                    for x in range(richHtml.find(r"RichText"), indexTexto):
                        rHtml += l[x]
                    indexTexto += 5
                    indexTextoFinal = richHtml.find("</div>")
                    indexHref = 0
                    for x in range(indexTexto, indexTextoFinal):
                        name = name + l[x]
                    name = name.replace("<div>", "")
                    name = name.replace(rHtml, "")
                    name = name.replace("\n", "")
                    name = name.strip()
                    # if name.find(r"\n",0,3)!=-1:
                    #     name = name.replace(r"r\")
                    if name.find("<a href") != -1:
                        n = list(name)
                        while True:
                            br = False
                            if "externalid" in name:
                                break
                            indexHref = name.find("<a href", indexHref)
                            if indexHref == -1:
                                break
                            indexHrefFinal = name.find("</a>", indexHref) + 4
                            indexTitle = name.find(" title=", indexHref)

                            for x in range(indexHref, indexHrefFinal):
                                try:
                                    hrefLink += n[x]
                                except:
                                    br = True
                                    break
                            if br:
                                break
                            if name.find("\" rel", indexHref, indexTitle) != -1:
                                indexTitle = name.find("\" rel", indexHref, indexTitle)
                            for x in range(indexHref + 9, indexTitle):
                                link += n[x]
                            name = name.replace(hrefLink, link)
                            indexHref += 20
                            hrefLink = ""
                            link = ""
                    if name.find("<div itemprop=\"copy-paste-block\"") != -1:
                        name = name.replace("<div itemprop=\"copy-paste-block\"", "")
                    nAspas = False
                    name = name.replace(r"&quot;", "\"")
                    name = name.replace("<u>", "")
                    if name.find(">", 2):
                        name = name.replace(">", "", 1)
                    message = name
                    message = message.replace("Mention:  ", "")
                    name = ""
                    isRichHtmlDone = False
            if "http://schema.skype.com/File" in line:
                indexFileUrl = 0
                indexNomeFile = 0
                while True:
                    localizacao = ""
                    indexFileUrl = line.find("fileUrl", indexFileUrl)
                    if indexFileUrl == -1:
                        break
                    indexFileUrl += 10
                    indexFileUrlFinal = line.find("siteUrl", indexFileUrl) - 3
                    if line.find("itemId", indexFileUrl) < indexFileUrlFinal and line.find("itemId",
                                                                                           indexFileUrl) != -1:
                        indexFileUrlFinal = line.find("itemId", indexFileUrl) - 3
                    indexNomeFile = line.find("fileName", indexNomeFile) + 11
                    indexNomeFileFinal = line.find("fileType", indexNomeFile) - 3
                    if line.find("filePreview", indexNomeFile) < indexNomeFileFinal and line.find("filePreview",
                                                                                                  indexNomeFile) != -1:
                        indexNomeFileFinal = line.find("filePreview", indexNomeFile) - 3
                    indexCleanTexto = richHtml.find("<div><span><img")
                    indexCleanTextoFinal = richHtml.find("</div>", indexCleanTexto) + 6
                    if indexCleanTexto != -1 and indexCleanTextoFinal != -1 and indexCleanTexto < indexCleanTextoFinal:
                        r = list(richHtml)
                        spanDelete = ""
                        for x in range(indexCleanTexto, indexCleanTextoFinal):
                            spanDelete += r[x]
                        richHtml = richHtml.replace(spanDelete, "")
                    l = list(line)
                    url = ""
                    nomeFile = ""
                    for x in range(indexFileUrl, indexFileUrlFinal):
                        url += l[x]
                    for x in range(indexNomeFile, indexNomeFileFinal):
                        nomeFile += l[x]

                    indexFileUrl += 50
                    indexNomeFile += 20
                    fich = File(url, nomeFile)
                    files.append(fich)
                    if files.__len__() != 0:
                        hasFiles = True
            if adptiveCard and "\"type\":\"AdaptiveCard\"" in line:
                indexAltText = line.find("altText")
                indexAltTextFinal = line.find("horizontalAlignment", indexAltText)
                l = list(line)
                for x in range(indexAltText + 10, indexAltTextFinal - 3):
                    elogio += l[x]

            if "call-log" in line:

                start = ""
                end = ""
                state = ""
                originator = ""
                target = ""
                indexStart = line.find("startTime")
                indexStartFinal = line.find("connectTime", indexStart)
                indexEnd = line.find("endTime", indexStartFinal)
                indexEndFinal = line.find("callDirection", indexEnd)
                indexCallState = line.find("callState", indexEndFinal)
                indexOriginator = line.find("originator", indexCallState)
                indexTarget = line.find("target", indexOriginator)
                indexTargetFinal = line.find("originatorParticipant", indexTarget)
                l = list(line)
                for x in range(indexStart + 12, indexStartFinal - 3):
                    start += l[x]
                for x in range(indexEnd + 10, indexEndFinal - 3):
                    end += l[x]
                for x in range(indexCallState + 12, indexOriginator - 3):
                    state += l[x]
                for x in range(indexOriginator + 21, indexTarget - 3):
                    originator += l[x]
                for x in range(indexTarget + 17, indexTargetFinal - 3):
                    target += l[x]
                if originator in arrayContactos:
                    contactoOriginator = arrayContactos.get(originator)
                else:
                    ind = line.find("originatorParticipant")
                    displayOrig = ""
                    for x in range(line.find("displayName", ind) + 14, line.find("\"}", ind)):
                        displayOrig += l[x]
                    contactoOriginator = Contacto(displayOrig, "sem email", originator)
                    arrayContactos[originator] = contactoOriginator
                if target in arrayContactos:
                    contactoTarget = arrayContactos.get(target)
                else:
                    displayTarget = ""
                    indT = line.find("targetParticipant")
                    for x in range(line.find("displayName", indT) + 14, line.find("\"}", indT)):
                        displayTarget += l[x]
                    contactoTarget = Contacto(displayTarget, "sem email", target)
                    arrayContactos[target] = contactoTarget
                dtStart = zulu.parse(start)
                callstart = datetime.utcfromtimestamp(dtStart.timestamp())
                callstart = callstart.astimezone(tz=tz.tzlocal())
                dtEnd = zulu.parse(end)
                callEnd = datetime.utcfromtimestamp(dtEnd.timestamp())
                callEnd = callEnd.astimezone(tz=tz.tzlocal())
                chamada = Chamada(contactoOriginator, str(callstart), str(callEnd), contactoTarget, state)
                arrayCallOneToOne.append(chamada)

            if isReacaoChunk and "ams_references" in line:
                isReacaoChunk = False
                reacaoChunkReady = True

            if isReacaoChunk:
                reacaoChunk += line
                reacaoChunk += ";"
            if "deltaEmotions" in line:
                isReacaoChunk = True

            if reacaoChunkReady:
                emojiReaction = ""
                timeReaction = ""
                arrayReacoes.clear()
                reacoes = reacaoChunk.split(";")
                reacoes = list(dict.fromkeys(reacoes))
                orgidReact = ""
                for r in reacoes:
                    if r != "":
                        reaction = Reaction()
                        rl = list(r)
                        indexKey = r.find("key")
                        indexKeyFinal = r.find("user", indexKey)
                        if indexKey != -1:
                            emojiReaction = ""
                            for x in range(indexKey + 5, indexKeyFinal - 2):
                                emojiReaction += rl[x]
                        indexOrgidReacao = r.find("orgid")
                        indexOrgidReacaoFinal = r.find("timeN", indexOrgidReacao)
                        for x in range(indexOrgidReacao + 6, indexOrgidReacaoFinal - 2):
                            orgidReact += rl[x]
                        try:
                            for x in range(0, 13):
                                timeReaction += rl[x]
                        except:
                            print("")
                        reaction.orgid = orgidReact
                        reaction.emoji = linkEmoji.format(emojiReaction)
                        try:
                            tr = datetime.utcfromtimestamp(float(timeReaction) / 1000.0)
                            tr = tr.astimezone(tz=tz.tzlocal())
                            reaction.time = str(tr)
                        except Exception:
                            reaction.time = str(timeReaction)

                        arrayReacoes.append(reaction)

                        orgidReact = ""
                        timeReaction = ""
                        reacaoChunkReady = False

        reacaoChunk = ""
        if message != "" or hasFiles:
            if message.count("A praise was given to") < 2:
                if "Call Log for Call" not in message:
                    if adptiveCard:
                        message += "Praise: {}".format(elogio)
                        message = message.replace(r"\n'", "")
                        message = message.replace("\"", "")
                        elogio = ""
                        adptiveCard = False
                    mensagem = MensagemCompleta()
                    mensagem.mention = mention
                    mensagem.message = cleanMessage(message)
                    mensagem.time = time
                    mensagem.sender = sender
                    mensagem.cvID = cvId
                    if arrayReacoes.__len__() > 0:
                        mensagem.reactions = arrayReacoes
                        # print(pathMulti)
                        try:
                            pathToAutopsy
                        except NameError:
                            pathReacts = pathMulti
                        else:
                            pathReacts = pathToAutopsy

                        with open(os.path.join(pathReacts, 'Reacts.csv'), 'a+', newline='',
                                  encoding="utf-8") as csvfile:
                            messagewriter = csv.writer(csvfile, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
                            for r in arrayReacoes:
                                if r.orgid in arrayContactos:
                                    c = arrayContactos.get(r.orgid)
                                else:
                                    c = Contacto(orgid=r.orgid)
                                messagewriter.writerow([str(idM), r.emoji, c.nome, r.time, user])
                            csvfile.close()
                    mensagem.files = files

                    arrayMensagens.append(mensagem)
                    hasFiles = False
                    idM += 1
                    arrayReacoes.clear()


def findpadrao(pathArmazenanto):
    logFinalRead = open(os.path.join(pathArmazenanto, "logTotal.txt"), "r", encoding="utf-8")
    flagMgs = 0
    stringBuffer = ""
    toCompare = False
    dupped = False
    buffer = []
    ready = False
    while line := logFinalRead.readline():
        # print(ready)
        if flagMgs == 1:
            buffer.append(line)
        if "RichText/Html" in line:
            flagMgs = 1
            toCompare = True
            buffer.append(line)
        if "renderContent" in line:
            if flagMgs != 1:
                flagMgs = 1
                toCompare = True
                buffer.append(line)
        if "parentMessageId" in line:
            flagMgs = 0
            ready = True
        if "amsreferences" in line:
            toCompare = False
        if toCompare:
            stringBuffer += line

        if ready:
            if len(bufferBuffer) != 0:
                for b in bufferBuffer:
                    if b == stringBuffer:
                        dupped = True
            bufferBuffer.append(stringBuffer)
            if not dupped:
                filtro(buffer)
            buffer.clear()
            dupped = False
            stringBuffer = ""
        ready = False
    logFinalRead.close()


def geraContactos(pathArmazenamento):
    logFinalRead = open(os.path.join(pathArmazenamento, "logTotal.txt"), "r", encoding="utf-8")
    orgid = ""
    name = ""
    mail = ""
    nAspas = False
    while line := logFinalRead.readline():
        if "itemRank" in line:
            indexNome = line.find("displayName")
            indexNomeFinal = line.find("userPrincipalName")
            indexMail = line.find("email")
            indexMailFinal = line.find("description")
            indexObject = line.find("objectId")
            indexObjectEnd = line.find("$$lastname_lowercase")
            lista = list(line)

            for a in range(indexNome + 13, indexNomeFinal - 2):
                name = name + lista[a]

            nameContacto = name
            nAspas = False

            for x in range(indexMail, indexMailFinal):
                if "\"" in lista[x]:
                    nAspas = not nAspas
                    continue
                if nAspas:
                    mail = mail + lista[x + 1]
            mail = mail.replace("\"", "")
            emailContacto = mail
            nAspas = False

            for z in range(indexObject, indexObjectEnd):
                if "\"" in lista[z]:
                    nAspas = not nAspas
                    continue
                if nAspas:
                    orgid = orgid + lista[z + 1]

            orgid = orgid.replace("\"", "")
            orgidContacto = orgid

            orgid = ""
            name = ""
            mail = ""
            nAspas = False

            contacto = Contacto(nameContacto, emailContacto, orgidContacto)
            if orgidContacto not in arrayContactos and orgidContacto != "":
                arrayContactos[orgidContacto] = contacto


def createhtmltables(pathToFolder, user):
    # FORMAT -------------------------------------------------

    css_string = '''
        .mystyle {
            font-size: 11pt; 
            font-family: Arial;
            border-collapse: collapse; 
            border: 1px solid silver;
            width: 100%;
        }
        .mystyle th {
            background-color: royalblue;
            color: white;
        }

        .mystyle td, th {
            padding: 5px;
        }

        .mystyle tr:nth-child(even) {
            background: #f2f2f2;;
        }

        .mystyle tr:hover {
            background: silver;
            cursor: pointer;
        }
        '''
    with open(os.path.join(pathToFolder, "df_style.css"), 'w') as f:
        f.write(css_string)

    html_string = '''
    <html>
      <head><title>{title}</title></head>
      <link rel="stylesheet" type="text/css" href="df_style.css"/>
      <body>
        <h1>{title}</h1>
        {table}
      </body>
    </html>.
    '''

    pd.set_option('display.width', 1000)
    pd.set_option('colheader_justify', 'center')

    # -------------------------------------------------

    def first_element(elem):
        return elem[0]

    # Create html table to present Contacts
    contacts_dict = {}
    for k, v in arrayContactos.items():
        contacts_dict[k] = [v.nome, v.email, user]

    sorted_by_name_contacts = {}
    iterator = 1
    for k, v in sorted(contacts_dict.items(), key=lambda item: item[1][0].upper()):
        sorted_by_name_contacts[iterator] = v
        iterator = iterator + 1

    frame_contacts = pd.DataFrame.from_dict(sorted_by_name_contacts, orient='index', columns=['Name', 'Email', 'User'])
    with open(os.path.join(pathToFolder, "User_Contacts.html"), 'w', encoding="utf-8") as file:
        file.write(html_string.format(title="CONTACTS", table=frame_contacts.to_html(classes='mystyle')))

    # Create html table to present Calls in Teams
    teams_calls = []
    for call_team in arrayEventCall:
        creator = call_team.creator.nome + ", " + call_team.creator.email
        new = [call_team.calldate, call_team.duration, call_team.count, creator, call_team.participants, user]
        teams_calls.append(new)

    teams_calls.sort(key=first_element)

    frame_team_calls = pd.DataFrame(teams_calls, columns=["End date", "Duration (minutes)", "Number of participants",
                                                          "Call Creator", "Call Participants", "user"])
    with open(os.path.join(pathToFolder, "Calls_Teams.html"), 'w', encoding="utf-8") as file:
        file.write(html_string.format(title="TEAM CALLS", table=frame_team_calls.to_html(classes='mystyle')))

    # Create html table to present Calls in private conversations
    private_calls = []
    for call_private in arrayCallOneToOne:
        new = [call_private.timestart, call_private.timefinish, call_private.criador.nome, call_private.criador.email,
               call_private.presente.nome, call_private.presente.email, user]
        private_calls.append(new)

    private_calls.sort(key=first_element)

    frame_private_calls = pd.DataFrame(private_calls, columns=["Date Start Call", "Date End Call", "Call Creator Name",
                                                               "Call Creator Email", "Call Participant Name",
                                                               "Call Participant Email", "User"])
    with open(os.path.join(pathToFolder, "Calls_Private_Conversations.html"), 'w', encoding="utf-8") as file:
        file.write(html_string.format(title="PRIVATE CALLS", table=frame_private_calls.to_html(classes='mystyle')))

    # Create html table to present formation of new Teams
    teams_formation = []
    for k, v in dictionaryConversationDetails.items():
        for member in v.members:
            new = [v.conversation_id, v.date, v.creator.nome, v.creator.email, member.nome, member.email, user]
            teams_formation.append(new)

    teams_formation.sort(key=first_element)

    frame_teams_formation = pd.DataFrame(teams_formation,
                                         columns=["Conversation", "Date", "Name of User who added member",
                                                  "Email of User who added member", "Member Name", "Member Email",
                                                  "User"])

    with open(os.path.join(pathToFolder, "Teams_Formation.html"), 'w', encoding="utf-8") as file:
        file.write(html_string.format(title="TEAMS FORMATION", table=frame_teams_formation.to_html(classes='mystyle')))

    # Create html table to present messages files
    frame_files = pd.read_csv(os.path.join(pathToFolder, "Files.csv"), delimiter=";", header=0,
                              names=["Message ID", "File name", "File link", 'User'])
    with open(os.path.join(pathToFolder, "User_Messages_Files.html"), 'w', encoding="utf-8") as file:
        file.write(html_string.format(title="MESSAGES FILES", table=frame_files.to_html(classes='mystyle')))

    # Create html table to present messages reactions
    frame_reacts = pd.read_csv(os.path.join(pathToFolder, "Reacts.csv"), delimiter=";", header=0,
                               names=["Message ID", "Reaction", "Reacted by", 'Date'])
    with open(os.path.join(pathToFolder, "User_Messages_Reacts.html"), 'w', encoding="utf-8") as file:
        file.write(html_string.format(title="MESSAGES REACTIONS", table=frame_reacts.to_html(classes='mystyle')))

    # Create html table to present messages
    frame_messages = pd.read_csv(os.path.join(pathToFolder, "Mensagens.csv"), delimiter=";", header=0,
                                 names=["Message ID", "Message", "Date", 'Sender', 'Conversation ID', 'User'],
                                 encoding="utf-8")
    with open(os.path.join(pathToFolder, "User_Messages.html"), 'w', encoding="utf-8") as file:
        file.write(html_string.format(title="MESSAGES", table=frame_messages.to_html(classes='mystyle')))

    # generate index html page
    html_string_index = '''
                <!DOCTYPE html>
            <html>
                <head>
                    <title>MS Teams Data</title>
                    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
                    <link rel="stylesheet" type="text/css" href="style.css"/
                </head>
                <body>
                    <h1 style="text-align: center; padding: 10% 0;">Microsoft Teams extracted user data</h1>
                    <h1 style="text-align: center; padding: 10% 0;">User: {user}</h1>
                    <ul>
                        <li>
                            <a href="User_Contacts.html">
                                <div class="icon">
                                    <i class="fa fa-address-book"></i>
                                    <i class="fa fa-address-book"></i>
                                </div>
                                <div class="name">Contacts</div>
                                <div class="number">{contacts_number}</div>
                            </a>
                        </li>
                        <li>
                            <a href="User_Messages.html">
                                <div class="icon">
                                    <i class="fa fa-envelope"></i>
                                    <i class="fa fa-envelope"></i>
                                </div>
                                <div class="name">Messages</div>
                                <div class="number">{messages_number}</div>
                            </a>
                        </li>
                        <li>
                            <a href="User_Messages_Files.html">
                                <div class="icon">
                                    <i class="fa fa-file"></i>
                                    <i class="fa fa-file"></i>
                                </div>
                                <div class="name">Messages Files</div>
                                <div class="number">{messages_files_number}</div>
                            </a>
                        </li>
                        <li>
                            <a href="User_Messages_Reacts.html">
                                <div class="icon">
                                    <i class="fa fa-thumbs-up"></i>
                                    <i class="fa fa-thumbs-up"></i>
                                </div>
                                <div class="name">Messages Reactions</div>
                                <div class="number">{messages_reacts_number}</div>
                            </a>
                        </li>
                        <li>
                            <a href="Calls_Private_Conversations.html">
                                <div class="icon">
                                    <i class="fa fa-phone"></i>
                                    <i class="fa fa-phone"></i>
                                </div>
                                <div class="name">Private Calls</div>
                                <div class="number">{private_calls_number}</div>
                            </a>
                        </li>
                        <li>
                            <a href="Teams_Formation.html">
                                <div class="icon">
                                    <i class="fa fa-users"></i>
                                    <i class="fa fa-users"></i>
                                </div>
                                <div class="name">Teams Formations</div>
                                <div class="number">{teams_formation_number}</div>
                            </a>
                        </li>
                        <li>
                            <a href="Calls_Teams.html">
                                <div class="icon">
                                    <i class="fa fa-phone"></i>
                                    <i class="fa fa-phone"></i>
                                </div>
                                <div class="name">Teams Calls</div>
                                <div class="number">{team_calls_number}</div>
                            </a>
                        </li>
                    </ul>
                </body>
            </html>
            '''
    with open(os.path.join(pathToFolder, "index.html"), 'w', encoding="utf-8") as file:
        file.write(html_string_index.format(contacts_number=str(frame_contacts.__len__()),
                                            messages_number=str(frame_messages.__len__()),
                                            messages_files_number=str(frame_files.__len__()),
                                            messages_reacts_number=str(frame_reacts.__len__()),
                                            private_calls_number=str(frame_private_calls.__len__()),
                                            teams_formation_number=str(frame_teams_formation.__len__()),
                                            team_calls_number=str(frame_team_calls.__len__()),
                                            user=user))

    css_string = '''
                    body 
                    {
                        margin: 0;
                        padding: 0;
                        font-family: sans-serif;
                    }

                    ul 
                    {
                        position: absolute;
                        top: 50%;
                        left: 50%;
                        transform: translate(-50%, -50%);
                        margin: 0;
                        padding: 20px 0;
                        display: flex;
                    }

                    ul li 
                    {
                        list-style: none;
                        text-align: center;
                        display: block;
                    }

                    ul li:last-child 
                    {
                        border-right: none;
                    }

                    ul li a 
                    {
                        text-decoration: none;
                        padding: 0 20px;
                        display: block;
                    }

                    ul li a .icon
                    {
                        width: 150px;
                        height: 40px;
                        text-align: center;
                        overflow: hidden;
                        margin: 0 auto 10px;
                    }

                    ul li a .icon .fa
                    {
                        width: 100%;
                        height: 100%;
                        line-height: 40px;
                        font-size: 34px;
                        transition: 0.3s;
                        color: #000;
                    }

                    ul li a:hover .icon .fa:last-child
                    {
                        color: royalblue;
                    }

                    ul li a:hover .icon .fa
                    {
                        transform: translateY(-100%);
                    }

                    .name
                    {
                        color: #000;
                    }

                    .number
                    {
                        color: red;
                    }

            '''
    with open(os.path.join(pathToFolder, "style.css"), 'w') as f:
        f.write(css_string)

    html = open(os.path.join(pathToFolder, "User_Messages.html"), 'r', encoding="utf-8")
    source_code = html.read()
    tables = pd.read_html(source_code)
    for i, table in enumerate(tables):
        table.to_csv(os.path.join(pathToFolder, 'messages_from_html.csv'), ';')


# ----------------------------------------------------------------------------------------------------------------


if __name__ == "__main__":
    args = sys.argv[1:]
    pathUsers = ""
    pathAppdata = ""
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
                sys.exit('ei.py --pathToEI <pathToEIFolder> -a <LDBFolderNameInProjetoEIAppData> or ei.py -u <pathToUsers> ')
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

                            with open(os.path.join(pathMulti, 'Reacts.csv'), 'a+', newline='',
                                      encoding="utf-8") as csvfile:
                                fieldnames = ['messageID', 'reacted_with', 'reacted_by', 'react_time', 'user']
                                messagewriter = csv.writer(csvfile, delimiter=';', quotechar='|',
                                                           quoting=csv.QUOTE_MINIMAL)
                                messagewriter.writerow(fieldnames)
                                csvfile.close()

                            crialogtotal(pathMulti)

                            geraContactos(pathMulti)

                            criacaoDeEquipas(pathMulti)
                            criarObjetosDeCriacaoDeEquipas(pathMulti)

                            extrairEventCallsToFile(pathMulti)
                            criarObjetosDeEventCalls(pathMulti)
                            idMessage = 1
                            findpadrao(pathMulti)

                            with open(os.path.join(pathMulti, 'Contactos.csv'), 'a+', newline='',
                                      encoding="utf-8") as csvfile:
                                fieldnames = ['nome', 'email', 'orgid''user']
                                messagewriter = csv.writer(csvfile, delimiter=';',
                                                           quotechar='|', quoting=csv.QUOTE_MINIMAL)
                                messagewriter.writerow(fieldnames)
                                for key, value in arrayContactos.items():
                                    messagewriter.writerow([value.nome, value.email, value.orgid, user])
                                csvfile.close()

                            with open(os.path.join(pathMulti, 'Mensagens.csv'), 'a+', newline='',
                                      encoding="utf-8") as csvfile:
                                fieldnames = ['messageID', 'message', 'time', 'sender', 'conversation_id', 'user']
                                messagewriter = csv.writer(csvfile, delimiter=';', quotechar='|',
                                                           quoting=csv.QUOTE_MINIMAL)
                                messagewriter.writerow(fieldnames)
                                for m in arrayMensagens:
                                    m.message = m.message.replace(";", "(semicolon)")
                                    if len(m.files) > 0:
                                        dictFiles[idMessage] = m.files
                                    if len(m.reactions) > 0:
                                        dictReacts[idMessage] = m.reactions
                                    if m.message == "":
                                        m.message = "Empty Text"
                                    messagewriter.writerow([str(idMessage), m.message, m.time, m.sender, m.cvID, user])
                                    idMessage += 1
                                csvfile.close()

                            with open(os.path.join(pathMulti, 'Files.csv'), 'a+', newline='',
                                      encoding="utf-8") as csvfile:
                                fieldnames = ['messageID', 'file_name', 'file_url', 'user']
                                messagewriter = csv.writer(csvfile, delimiter=';', quotechar='|',
                                                           quoting=csv.QUOTE_MINIMAL)
                                messagewriter.writerow(fieldnames)
                                for key, value in dictFiles.items():
                                    for f in value:
                                        print(f.toString())
                                        print(str([str(key), f.nome, f.local, 'user']))
                                        messagewriter.writerow([str(key), f.nome, f.local, 'user'])
                                csvfile.close()

                            with open(os.path.join(pathMulti, 'EventCall.csv'), 'a+', newline='',
                                      encoding="utf-8") as csvfile:
                                fieldnames = ['calldate', 'creator_name', 'creator_email', 'count', 'duration',
                                              'participant_name', 'participant_email', 'user']
                                callwriter = csv.writer(csvfile, delimiter=';',
                                                        quotechar='|', quoting=csv.QUOTE_MINIMAL)
                                callwriter.writerow(fieldnames)
                                for call in arrayEventCall:
                                    for oID in call.orgids:
                                        if oID in arrayContactos:
                                            c = arrayContactos.get(oID)
                                        else:
                                            c = Contacto('Desc....', 'Desc.', oID)

                                        callwriter.writerow(
                                            [call.calldate, call.creator.nome, call.creator.email, call.count, c.nome,
                                             c.email, user])
                                csvfile.close()

                            with open(os.path.join(pathMulti, 'Conversations.csv'), 'a+', newline='',
                                      encoding="utf-8") as csvfile:
                                fieldnames = ['conversation_ID', 'date', 'nome_creator', 'email_creator', 'member_name',
                                              'member_email', 'user']
                                callwriter = csv.writer(csvfile, delimiter=';',
                                                        quotechar='|', quoting=csv.QUOTE_MINIMAL)
                                callwriter.writerow(fieldnames)
                                for key, value in dictionaryConversationDetails.items():
                                    for member in value.members:
                                        callwriter.writerow(
                                            [value.conversation_id, value.date, value.creator.nome, value.creator.email,
                                             member.nome, member.email, user])
                                csvfile.close()

                            with open(os.path.join(pathMulti, 'CallOneToOne.csv'), 'a+', newline='',
                                      encoding="utf-8") as csvfile:
                                fieldnames = ['originator_name', 'originator_email', 'time_start', 'time_finish',
                                              'target_nome', 'target_email', 'state', 'user']
                                callwriter = csv.writer(csvfile, delimiter=';', quotechar='|',
                                                        quoting=csv.QUOTE_MINIMAL)
                                callwriter.writerow(fieldnames)
                                for ch in arrayCallOneToOne:
                                    callwriter.writerow(
                                        [ch.criador.nome, ch.criador.email, ch.timestart, ch.timefinish,
                                         ch.presente.nome, ch.presente.email, ch.state, user])
                                csvfile.close()

                                createhtmltables(pathMulti, user)

                                arrayContactos.clear()
                                dictFiles.clear()
                                arrayMensagens.clear()
                                arrayEventCall.clear()
                                dictionaryConversationDetails.clear()
                                arrayCallOneToOne.clear()


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

            with open(os.path.join(pathToAutopsy, 'Reacts.csv'), 'a+', newline='',
                      encoding="utf-8") as csvfile:
                fieldnames = ['messageID', 'reacted_with', 'reacted_by', 'react_time', 'user']
                messagewriter = csv.writer(csvfile, delimiter=';', quotechar='|',
                                           quoting=csv.QUOTE_MINIMAL)
                messagewriter.writerow(fieldnames)
                csvfile.close()

            crialogtotal(pathToAutopsy)
            geraContactos(pathToAutopsy)

            criacaoDeEquipas(pathToAutopsy)
            criarObjetosDeCriacaoDeEquipas(pathToAutopsy)

            extrairEventCallsToFile(pathToAutopsy)
            criarObjetosDeEventCalls(pathToAutopsy)

            findpadrao(pathToAutopsy)

            with open(os.path.join(pathToAutopsy, 'Contactos.csv'), 'a+', newline='', encoding="utf-8") as csvfile:
                fieldnames = ['nome', 'email', 'orgid', 'user']
                messagewriter = csv.writer(csvfile, delimiter=';',
                                           quotechar='|', quoting=csv.QUOTE_MINIMAL)
                messagewriter.writerow(fieldnames)
                for key, value in arrayContactos.items():
                    messagewriter.writerow([value.nome, value.email, value.orgid, user])
                csvfile.close()

            with open(os.path.join(pathToAutopsy, 'Mensagens.csv'), 'a+', newline='', encoding="utf-8") as csvfile:
                fieldnames = ['messageID', 'message', 'time', 'sender', 'conversation_id', 'user']
                messagewriter = csv.writer(csvfile, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
                messagewriter.writerow(fieldnames)
                for m in arrayMensagens:
                    m.message = m.message.replace(";", "(semicolon)")
                    if len(m.files) > 0:
                        dictFiles[idMessage] = m.files
                    if m.message == "":
                        m.message = "Empty Text"
                    messagewriter.writerow([str(idMessage), m.message, m.time, m.sender, m.cvID, user])
                    idMessage += 1
                csvfile.close()

            with open(os.path.join(pathToAutopsy, 'Files.csv'), 'a+', newline='', encoding="utf-8") as csvfile:
                fieldnames = ['messageID', 'file_name', 'file_url', 'user']
                messagewriter = csv.writer(csvfile, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
                messagewriter.writerow(fieldnames)
                for key, value in dictFiles.items():
                    for f in value:
                        print(f.toString())
                        print(str([str(key), f.nome, f.local, 'user']))
                        messagewriter.writerow([str(key), f.nome, f.local, 'user'])
                csvfile.close()

            with open(os.path.join(pathToAutopsy, 'EventCall.csv'), 'a+', newline='',
                      encoding="utf-8") as csvfile:
                fieldnames = ['calldate', 'creator_name', 'creator_email', 'count', 'duration', 'participant_name',
                              'participant_email', 'user']
                callwriter = csv.writer(csvfile, delimiter=';',
                                        quotechar='|', quoting=csv.QUOTE_MINIMAL)
                callwriter.writerow(fieldnames)
                for call in arrayEventCall:
                    for oID in call.orgids:
                        if oID in arrayContactos:
                            c = arrayContactos.get(oID)
                        else:
                            c = Contacto('Desc.', 'Desc.', oID)
                        callwriter.writerow(
                            [call.calldate, call.creator.nome, call.creator.email, call.count, str(call.duration),
                             c.nome, c.email, user])
                csvfile.close()

            with open(os.path.join(pathToAutopsy, 'Conversations.csv'), 'a+', newline='',
                      encoding="utf-8") as csvfile:
                fieldnames = ['conversation_ID', 'date', 'nome_creator', 'email_creator', 'member_name',
                              'member_email', 'user']
                callwriter = csv.writer(csvfile, delimiter=';',
                                        quotechar='|', quoting=csv.QUOTE_MINIMAL)
                callwriter.writerow(fieldnames)
                for key, value in dictionaryConversationDetails.items():
                    for member in value.members:
                        callwriter.writerow(
                            [value.conversation_id, value.date, value.creator.nome, value.creator.email, member.nome,
                             member.email, user])
                csvfile.close()
            with open(os.path.join(pathToAutopsy, 'CallOneToOne.csv'), 'a+', newline='',
                      encoding="utf-8") as csvfile:
                fieldnames = ['originator_name', 'originator_email', 'time_start', 'time_finish', 'target_nome',
                              'target_email', 'state', 'user']
                callwriter = csv.writer(csvfile, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
                callwriter.writerow(fieldnames)
                for ch in arrayCallOneToOne:
                    callwriter.writerow(
                        [ch.criador.nome, ch.criador.email, ch.timestart, ch.timefinish, ch.presente.nome,
                         ch.presente.email, ch.state, user])
                csvfile.close()

            createhtmltables(pathToAutopsy, user)
