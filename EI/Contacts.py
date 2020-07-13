import os

class Contacto:
    def __init__(self, nome="Desc.", email="Desc.", orgid="Desc."):
        self.nome = nome
        self.email = email
        self.orgid = orgid

    def toString(self):
        return "Nome: {0} || Email: {1} || Orgid {2}".format(self.nome, self.email, self.orgid)



def geraContactos(pathArmazenamento):
    arrayContactos = {}
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
    return arrayContactos
