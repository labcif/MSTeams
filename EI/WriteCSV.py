import csv
import os

import Contacts


def writeFiles(user, pathArmazenamento, arrayContactos, arrayMensagens, arrayEventCall, dictionaryConversationDetails,
               arrayCallOneToOne, tmCSV):
    idMessage = 1
    dictFiles = {}
    with open(os.path.join(pathArmazenamento, 'Contactos_{}_{}.csv'.format(user, tmCSV)), 'a+', newline='',
              encoding="utf-8") as csvfile:
        fieldnames = ['nome', 'email', 'orgid', 'user']
        messagewriter = csv.writer(csvfile, delimiter=';',
                                   quotechar='|', quoting=csv.QUOTE_MINIMAL)
        messagewriter.writerow(fieldnames)
        for key, value in arrayContactos.items():
            messagewriter.writerow([value.nome, value.email, value.orgid, user])
        csvfile.close()

    with open(os.path.join(pathArmazenamento, 'Mensagens_{}_{}.csv'.format(user, tmCSV)), 'a+', newline='',
              encoding="utf-8") as csvfile:
        fieldnames = ['messageID', 'message', 'time', 'sender', 'conversation_id', 'user']
        messagewriter = csv.writer(csvfile, delimiter=';', quotechar='|',
                                   quoting=csv.QUOTE_MINIMAL)
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

    with open(os.path.join(pathArmazenamento, 'Files_{}_{}.csv'.format(user, tmCSV)), 'a+', newline='',
              encoding="utf-8") as csvfile:
        fieldnames = ['messageID', 'file_name', 'file_url', 'user']
        messagewriter = csv.writer(csvfile, delimiter=';', quotechar='|',
                                   quoting=csv.QUOTE_MINIMAL)
        messagewriter.writerow(fieldnames)
        for key, value in dictFiles.items():
            for f in value:
                print(f.toString())
                print(str([str(key), f.nome, f.local, user]))
                messagewriter.writerow([str(key), f.nome, f.local, user])
        csvfile.close()

    with open(os.path.join(pathArmazenamento, 'EventCall_{}_{}.csv'.format(user, tmCSV)), 'a+', newline='',
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
                    c = Contacts.Contacto('Desc....', 'Desc.', oID)

                callwriter.writerow(
                    [call.calldate, call.creator.nome, call.creator.email, call.count, call.duration, c.nome,
                     c.email, user])
        csvfile.close()

    with open(os.path.join(pathArmazenamento, 'Conversations_{}_{}.csv'.format(user, tmCSV)), 'a+', newline='',
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

    with open(os.path.join(pathArmazenamento, 'CallOneToOne_{}_{}.csv'.format(user, tmCSV)), 'a+', newline='',
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
