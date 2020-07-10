import os


def find(pt, path):
    fls = []
    for root, dirs, files in os.walk(path):
        for name in files:
            if pt in name:
                fls.append(os.path.join(root, name))
    return fls


def testeldb(path, pathArmazenamento,pathModule):
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


def writelog(path, pathArmazenamento):
    logFinal = open(os.path.join(pathArmazenamento, "logTotal.txt"), "a+", encoding="utf-8")
    log = open(path, "r", encoding="mbcs")
    while line := log.readline():
        logFinal.write(line)
    log.close()


def crialogtotal(pathArmazenamento, pathModule, levelDBPath):
    # os.system(r'cmd /c "LDBReader\PrjTeam.exe"')
    logFinal = open(os.path.join(pathArmazenamento, "logTotal.txt"), "a+", encoding="utf-8")
    os.system('cmd /c "{}LevelDBReader\\eiTeste.exe "{}" "{}"'.format(pathModule, levelDBPath, pathArmazenamento))
    testeldb(levelDBPath + "\\lost", pathArmazenamento, pathModule)
    logLost = find(".log", levelDBPath + "\\lost")
    writelog(os.path.join(pathArmazenamento, "logIndexedDB.txt"), pathArmazenamento)
    logsIndex = find(".log", levelDBPath)
    for path in logLost:
        writelog(path, pathArmazenamento)
    for path in logsIndex:
        writelog(path, pathArmazenamento)
    logFinal.close()
