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