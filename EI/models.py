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


class Contacto:
    def __init__(self, nome="Desc.", email="Desc.", orgid="Desc."):
        self.nome = nome
        self.email = email
        self.orgid = orgid

    def toString(self):
        return "Nome: {0} || Email: {1} || Orgid {2}".format(self.nome, self.email, self.orgid)


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


class File:
    def __init__(self, local, nome):
        self.local = local
        self.nome = nome

    def toString(self):
        return "Local:{0} Nome {1}".format(self.local, self.nome)


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
