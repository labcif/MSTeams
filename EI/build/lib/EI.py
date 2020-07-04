import getopt
import glob
import os
import pathlib
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import zulu as zulu
from dateutil.tz import tz

from models import Contacto, MensagemCompleta, Reaction, File, Chamada, ConversationCreationDetails, EventCall

appdata = os.getenv('APPDATA')
home = str(Path.home())
levelDBPath = appdata + "\\Microsoft\\Teams\\IndexedDB\\https_teams.microsoft.com_0.indexeddb.leveldb"
levelDBPathLost = appdata + "\\Microsoft\\Teams\\IndexedDB\\https_teams.microsoft.com_0.indexeddb.leveldb\\lost"
projetoEIAppDataPath = appdata + "\\ProjetoEI\\"

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
dictArrays = {}


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
    return st


def testeldb(path, pathArmazenamento):
    logFinal = open(os.path.join(pathArmazenamento, "logTotal.txt"), "a+", encoding="utf-8")
    from subprocess import Popen, PIPE
    import ast
    for f in os.listdir(path):
        if f.endswith(".ldb") and os.path.getsize(path + "\\" + f) > 0:
            process = Popen([os.path.join(home, r"PycharmProjects\EI\ldbdump"), os.path.join(path, f)], stdout=PIPE)
            (output, err) = process.communicate()
            try:
                result = output.decode("utf-8")
                for line in (result.split("\n")[1:]):
                    if line.strip() == "": continue
                    parsed = ast.literal_eval("{" + line + "}")
                    key = list(parsed.keys())[0]
                    logFinal.write(parsed[key])
            except:
                print("falhou")


def writelog(path, pathArmazenamento):
    logFinal = open(os.path.join(pathArmazenamento, "logTotal.txt"), "a+", encoding="utf-8")
    log = open(path, "r", encoding="mbcs")
    while line := log.readline():
        logFinal.write(line)
    log.close()


def crialogtotal(pathArmazenamento):
    # os.system(r'cmd /c "LDBReader\PrjTeam.exe"')
    logFinal = open(os.path.join(pathArmazenamento, "logTotal.txt"), "a+", encoding="utf-8")
    os.system('cmd /c "LevelDBReader\eiTeste.exe "{0}" "{1}"'.format(levelDBPath, pathArmazenamento))
    testeldb(levelDBPathLost, pathArmazenamento)
    logLost = find(".log", levelDBPathLost)
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

    while line := logEventCalls.readline():
        marcador = 0

        if "pinnedTime_" in line:
            # ir buscar data da chamada
            calldate = ""

            indexcalldate = line.find("policyViolation")
            indexcalldatefinal = line.find(",8:orgid")

            lista = list(line)
            for i in range(indexcalldate + 20, indexcalldatefinal):
                calldate = calldate + lista[i]

            # tentar converter timestamp para timezone UTC
            try:
                calldate = int(calldate[:10])
                calldate = datetime.utcfromtimestamp(calldate)
            except:
                calldate = calldate

            calldate = calldate.astimezone(tz=tz.tzlocal())

            # ir buscar utilizador que iniciou a chamada
            callcreator = ""

            indexcallcreator = line.find("8:orgid:")
            indexcallcreatorfinal = line.find("ackrequired_")

            for i in range(indexcallcreator + 8, indexcallcreatorfinal - 3):
                callcreator = callcreator + lista[i]

            # verficar se orgid existe na lista de contactos do utilizador
            if callcreator in arrayContactos:
                callcreator = arrayContactos[callcreator].nome

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
            eventcall = EventCall(str(calldate), callcreator, participantscount, callduration, participants, orgids)
            arrayEventCall.append(eventcall)
            # escrever para log

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

        if flagWrite == 1:
            logTeamsCreation.write(line)
        if "<addmember>" in line and count:  # inicio de nova informação de criacao de equipa # ThreadActivity/AddMember
            flagWrite = 1
            count = 0
            logTeamsCreation.write(line)
        if "_emailDetails_" in line and flagWrite == 1:  # fim da informação
            logTeamsCreation.write(line)
            logTeamsCreation.write("\n-----------------------------------------------"
                                   "------------------------------------------------------------------\n")
            logTeamsCreation.write("----------------------------------------------------------------"
                                   "-------------------------------------------------\n")
            flagWrite = 0
            count = 1

    # fechar ficheiros
    logFinalRead.close()
    logTeamsCreation.close()


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

    while line := logTeamsCreation.readline():

        # ir buscar id da conversacao e data de criacao
        if "messageso\"" in line:
            flag1 = 1

            # ir buscar id da conversacao
            conversationid = ""

            indexconverstionid = line.find("messageso\"")
            indexconverstionidfinal = line.find("@")

            lista = list(line)
            for i in range(indexconverstionid + 25, indexconverstionidfinal):
                conversationid = conversationid + lista[i]

            # ir buscar data de criacao da conversacao
            date = ""

            timestamp = int(line[:10])
            # tentar converter timestamp para timezone UTC
            try:
                date = datetime.utcfromtimestamp(timestamp)
            except:
                date = timestamp

            date = date.astimezone(tz=tz.tzlocal())

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
            # criar objeto de ConversationCreationDetails()
            conversation_details = ConversationCreationDetails(conversationid, str(date), initiatororgid, targets)

            if str(date) not in dictionaryConversationDetails:
                dictionaryConversationDetails[str(date)] = conversation_details

    # ir buscar contacto ao criador e membros das conversacao
    for conversa in dictionaryConversationDetails.values():
        # ir buscar contacto do criador da conversacao
        if conversa.creator in arrayContactos:
            creator = arrayContactos[conversa.creator]
            conversa.creator = ""
            conversa.creator = creator.toString()

        # ir buscar contacto do(s) membro(s) a ser(em) adicionado(s)
        members = []
        for user_orgid in conversa.members:
            if user_orgid in arrayContactos:
                member = arrayContactos[user_orgid]
                members.append(member.toString())
            else:
                members.append(user_orgid)
        conversa.members = members.copy()

    logTeamsCreation.close()


def filtro(buffer):
    filtros = {
        "trimmedMessageContent": 0,
        "renderContent": 0,
        "skypeguid": 0
    }
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
    reactions = []
    reacaoChunkReady = False
    files = []
    ok = 1
    link = ""
    hrefLink = ""
    nAspas = False

    for filter in filtros:
        for line in buffer:
            if filter in line:
                filtros[filter] = 1
    for x in filtros.values():
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
                indexTextoFinal = line.find("skypeguid")
                l = list(line)
                for x in range(indexTexto, indexTextoFinal):
                    if "\"" in l[x]:
                        nAspas = not nAspas
                        continue
                    if nAspas:
                        name = name + l[x + 1]
                name = name.replace("\"", "")
                sender = name
                nAspas = False
                nome = sender
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
                            if "externalid" in name:
                                break
                            indexHref = name.find("<a href", indexHref)
                            if indexHref == -1:
                                break
                            indexHrefFinal = name.find("</a>", indexHref) + 4
                            indexTitle = name.find("\" title=", indexHref)
                            for x in range(indexHref, indexHrefFinal):
                                hrefLink += n[x]
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
                    indexNomeFile = line.find("fileName", indexNomeFile) + 11
                    indexNomeFileFinal = line.find("fileType", indexNomeFile) - 3
                    indexCleanTexto = richHtml.find("<div><span><img")
                    indexCleanTextoFinal = richHtml.find("</div>", indexCleanTexto) + 6
                    if indexCleanTexto != -1 and indexCleanTextoFinal != -1 and indexCleanTexto < indexCleanTextoFinal:
                        r = list(richHtml)
                        spanDelete = ""
                        for x in range(indexCleanTexto, indexCleanTextoFinal):
                            spanDelete += r[x]
                        richHtml = richHtml.replace(spanDelete, "")
                        print(richHtml)
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
            if "call-log" in line:
                # print(line)
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
                dtEnd = zulu.parse(start)
                callEnd = datetime.utcfromtimestamp(dtEnd.timestamp())
                callEnd = callstart.astimezone(tz=tz.tzlocal())

                chamada = Chamada(contactoOriginator, str(callstart), str(callEnd), contactoTarget, state)
                arrayCallOneToOne.append(chamada)

            if isReacaoChunk and "ams_referencesa" in line:
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
                reacoes = reacaoChunk.split(";")
                reacoes = list(dict.fromkeys(reacoes))
                orgidReact = ""
                for r in reacoes:

                    if r != "":
                        reaction = Reaction()
                        rl = list(r)
                        indexKey = r.find("key")
                        indexKeyFinal = r.find("usersa", indexKey)
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
                        reaction.emoji = emojiReaction
                        reaction.time = timeReaction
                        arrayReacoes.append(reaction)
                        orgidReact = ""
                        timeReaction = ""
                        reacaoChunkReady = False

        reacaoChunk = ""

        if message != "" or hasFiles:
            if "Call Log for Call" not in message:
                mensagem = MensagemCompleta()
                if isMention:
                    mensagem.isMention = True
                mensagem.mention = mention
                mensagem.message = message
                mensagem.time = time
                mensagem.sender = sender

                mensagem.cvID = cvId
                mensagem.reactions = arrayReacoes
                arrayReacoes.clear()
                mensagem.files = files
                arrayMensagens.append(mensagem)
                hasFiles = False


def findpadrao(pathArmazenanto):
    logFinalRead = open(os.path.join(pathArmazenanto, "logTotal.txt"), "r", encoding="utf-8")
    flagMgs = 0
    stringBuffer = ""
    toCompare = False
    dupped = False
    buffer = []
    ready = False
    while line := logFinalRead.readline():
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
        if "pinnedTime_" in line:
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
    print(arrayMensagens.__len__())
    for m in arrayMensagens:
        print(m.toString())

    # logCall.close()


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
            # print(contacto.toString())

            if orgidContacto not in arrayContactos and orgidContacto != "":
                arrayContactos[orgidContacto] = contacto


if __name__ == "__main__":
    args = sys.argv[1:]
    pathUsers = ""
    pathAppdata = ""
    dupped = False
    try:
        opts, args = getopt.getopt(args, "hua:", ["users=", "autopsy="])
    except getopt.GetoptError:
        print('ei.py -u <pathToUsers> ')
        print("or")
        print('ei.py -a <pathToAppdataProjetoEI>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print('ei.py -u <pathToUsers> ')
            print("or")
            print('ei.py -a <pathToAppdataProjetoEI>')
            sys.exit()
        elif opt in ("-u", "--users"):
            print(arg)
            count = 0
            pathMulti = ""
            # fileTest = open(projetoEIAppDataPath + "teste.txt", "a+", encoding="UTF-8")
            pathUsers = arg
            for root, dirs, files in os.walk(pathUsers):
                for name in dirs:
                    if "https_teams.microsoft.com" in name:
                        for user in arrayUsers:
                            if name == user:
                                dupped = True
                        if not dupped:
                            arrayUsers.append(os.path.join(root, name))
                            levelDBPath = os.path.join(root, name)
                            current_milli_time = lambda: int(round(time.time() * 1000))
                            dupped = False
                            try:
                                pathMulti = projetoEIAppDataPath + str(current_milli_time()) + "\\"
                                os.mkdir(pathMulti)
                            except OSError:
                                print("Creation of the directory %s failed" % pathMulti)
                            else:
                                print("Successfully created the directory %s " % pathMulti)

                            crialogtotal(pathMulti)

                            geraContactos(pathMulti)

                            criacaoDeEquipas(pathMulti)
                            criarObjetosDeCriacaoDeEquipas(pathMulti)

                            extrairEventCallsToFile(pathMulti)
                            criarObjetosDeEventCalls(pathMulti)

                            findpadrao(pathMulti)
                            dictArrays[pathMulti + " - Contactos"] = arrayContactos
                            for key, value in arrayContactos.items():
                                # fileTest.write(value.toString())
                                print(key, value.toString())
                            arrayContactos.clear()
                            dictArrays[pathMulti + " - Mensagens"] = arrayMensagens
                            for m in arrayMensagens:
                                print(m.toString())
                                # fileTest.write(m.toString())
                            arrayMensagens.clear()
                            dictArrays[pathMulti + " - EventCall"] = arrayEventCall
                            for call in arrayEventCall:
                                print(call.toString())
                                # fileTest.write(call.toString())
                            arrayEventCall.clear()
                            dictArrays[pathMulti + " - Conversation Details"] = dictionaryConversationDetails
                            for key, value in dictionaryConversationDetails.items():
                                print(key, value.toString())
                                # fileTest.write(value.toString())
                            dictionaryConversationDetails.clear()
                            dictArrays[pathMulti + " - ChamadasOneToOne"] = arrayCallOneToOne
                            for ch in arrayCallOneToOne:
                                print(ch.toString())
                                # fileTest.write(ch.toString())
                            arrayCallOneToOne.clear()
            #                 fileTest.write("-------------------------------------")
            #                 fileTest.write("-------------------------------------")
            #                 fileTest.write("-------------------------------------")
            # fileTest.close()
        elif opt in ("-a", "--autopsy"):
            # -a https_teams.microsoft.com_0.indexeddb.leveldb
            levelDBPath = projetoEIAppDataPath + "\\" + arg

            crialogtotal(projetoEIAppDataPath)
            geraContactos(projetoEIAppDataPath)

            criacaoDeEquipas(projetoEIAppDataPath)
            criarObjetosDeCriacaoDeEquipas(projetoEIAppDataPath)

            extrairEventCallsToFile(projetoEIAppDataPath)
            criarObjetosDeEventCalls(projetoEIAppDataPath)

            findpadrao(projetoEIAppDataPath)

            for key, value in arrayContactos.items():
                print(key, value.toString())
            for m in arrayMensagens:
                print(m.toString())
            for call in arrayEventCall:
                print(call.toString())
            for key, value in dictionaryConversationDetails.items():
                print(key, value.toString())
            for ch in arrayCallOneToOne:
                print(ch.toString())
