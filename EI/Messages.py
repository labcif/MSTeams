import csv
import os
import zulu as zulu
import Contacts
import MSTeamsDecoder
from Calls import Chamada
from dateutil.tz import tz
from datetime import datetime

linkEmoji = "https://statics.teams.cdn.office.net/evergreen-assets/skype/v2/{0}/50.png"
idM = 1
arrayMensagens = []
bufferBuffer = []
arrayCallOneToOne = []
tmCSV=""


class File:
    def __init__(self, local, nome):
        self.local = local
        self.nome = nome

    def toString(self):
        return "Local:{0} Nome {1}".format(self.local, self.nome)


class Reaction:
    def __init__(self):
        self.emoji = ""
        self.time = ""
        self.orgid = ""

    def toString(self):
        return "emoji: {0} || Time: {1} || Orgid: {2}".format(self.emoji, self.time, self.orgid)


class MensagemCompleta:
    def __init__(self):
        self.message = ""
        self.sender = ""
        self.time = ""
        self.files = []
        self.hasEmoji = False
        self.isMention = False
        self.mention = ""
        self.reactions = []
        self.cvID = ""

    def toString(self):
        return "Message: {0} || Time: {1} || Sender: {2} || IsMention: {3} || Mention: {4} || HasEmoji: {5} || " \
               "hasFiles:{6} || hasReactions:{7} || cvID:{8}".format(self.message, self.time, self.sender,
                                                                     str(self.isMention),
                                                                     self.mention, str(self.hasEmoji),
                                                                     str(len(self.files)), str(len(self.reactions)),
                                                                     self.cvID)


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


def filtro(buffer, arrayContactos, pathArmazenanto, user):
    filtros = {
        "messageStorageState": 0,
        "parentMessageId": 0,
        "composetime": 0
    }
    global idM
    sender = ""
    time = ""
    arrayReacoes = []
    indexTexto = 0
    indexTextoFinal = 0
    name = ""
    isRichHtml = False
    isRichHtmlDone = False
    richHtml = ""
    reacaoChunk = ""
    isReacaoChunk = False
    hasFiles = False
    cvId = ""
    message = ""
    elogio = ""
    adptiveCard = False
    reacaoChunkReady = False
    files = []
    ok = 1
    link = ""
    hrefLink = ""

    for filter in filtros:
        for line in buffer:
            if filter in line:
                filtros[filter] = 1
    for y, x in filtros.items():
        if x == 0:
            ok = 0
    if ok:
        mention = ""
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
                    st = MSTeamsDecoder.utf16customdecoder(richHtml, "<div>")
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
                    contactoOriginator = Contacts.Contacto("Desc.", "sem email", originator)
                    arrayContactos[originator] = contactoOriginator
                if target in arrayContactos:
                    contactoTarget = arrayContactos.get(target)
                else:
                    displayTarget = ""
                    contactoTarget = Contacts.Contacto("desc.", "sem email", target)
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
                        if reaction.time.find("value")==-1:
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
                        with open(os.path.join(pathArmazenanto, 'Reacts_{}_{}.csv'.format(user,tmCSV)), 'a+', newline='',
                                  encoding="utf-8") as csvfile:
                            messagewriter = csv.writer(csvfile, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL)
                            for r in arrayReacoes:
                                if r.orgid in arrayContactos:
                                    c = arrayContactos.get(r.orgid)
                                else:
                                    c = Contacts.Contacto(orgid=r.orgid)
                                messagewriter.writerow([str(idM), r.emoji, c.nome, r.time, user])
                            csvfile.close()
                    mensagem.files = files
                    arrayMensagens.append(mensagem)
                    hasFiles = False
                    idM += 1
                    arrayReacoes.clear()


def findpadrao(pathArmazenanto, arrayContactos, user,tm):
    logFinalRead = open(os.path.join(pathArmazenanto, "logTotal.txt"), "r", encoding="utf-8")
    flagMgs = 0
    stringBuffer = ""
    toCompare = False
    dupped = False
    buffer = []
    arrayReturn = {}
    ready = False
    global tmCSV
    tmCSV=tm
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
                filtro(buffer, arrayContactos, pathArmazenanto, user)
            buffer.clear()
            dupped = False
            stringBuffer = ""
        ready = False
    logFinalRead.close()
    arrayReturn["mensagem"] = arrayMensagens
    arrayReturn["callOneToOne"] = arrayCallOneToOne
    return arrayReturn
