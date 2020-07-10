import os
from dateutil.tz import tz
from datetime import datetime
import re
import Contacts


class ConversationCreationDetails:
    def __init__(self, conversation_id, date, creator, members):
        self.conversation_id = conversation_id
        self.date = date
        self.creator = creator
        self.members = members

    def toString(self):
        return "CONVERSATION: {0} ||| DATE: {1} ||| CREATOR: {2} ||| MEMBERS: {3}".format(self.conversation_id,
                                                                                          self.date, self.creator,
                                                                                          self.members)


def criacaoDeEquipas(pathArmazenamento):
    # ficheiro onde vão ser colocadas as informações da criação de equipas
    logTeamsCreation = open(os.path.join(pathArmazenamento, "teamsCreation.txt"), "a+", encoding="utf-8")

    # abrir ficheiro logTotal para extrair informações de criacao das equipas
    logFinalRead = open(os.path.join(pathArmazenamento, "logTotal.txt"), "r", encoding="utf-8")

    flagWrite = 0  # flag que indica quando escrever para o ficheiro
    count = 1
    # ler cada linha do ficheiro logTotal e procurar "<addmember>"
    while line := logFinalRead.readline():

        if "messageso" in line and "<addmember>" in line and "<eventtime>" in line and "<initiator>" in line and count:
            # inicio de nova informação de criacao de equipa # ThreadActivity/AddMember
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


def criarObjetosDeCriacaoDeEquipas(pathArmazenamento, arrayContactos):
    # ficheiro para guardar objetos de criacao de conversacoes/Equipas

    # abrir ficheiro com a informacao das equipas para criar objetos
    logTeamsCreation = open(os.path.join(pathArmazenamento, "teamsCreation.txt"), "r", encoding="utf-8")
    dictionaryConversationDetails = {}
    conversationid = ""
    date = ""
    initiatororgid = ""
    targets = []
    arrayReturn = {}
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

            if conversationid.find(",") != -1:
                arr = conversationid.split(",")
                conversationid = arr[1]
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
                    contacto = Contacts.Contacto(username, email, orgid)
                    arrayContactos[orgid] = contacto
        #     todo fix
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
            conversa.creator = Contacts.Contacto('Unknown', 'Unknown', oID)

        # ir buscar contacto do(s) membro(s) a ser(em) adicionado(s)
        members = []
        for user_orgid in conversa.members:
            if user_orgid in arrayContactos:
                member = arrayContactos[user_orgid]
                members.append(member)
            else:
                us = Contacts.Contacto("Desc.", "Desc.", user_orgid)
                members.append(us)
        conversa.members = members.copy()

    logTeamsCreation.close()
    arrayReturn["contacts"] = arrayContactos
    arrayReturn["teams"] = dictionaryConversationDetails
    return arrayReturn
