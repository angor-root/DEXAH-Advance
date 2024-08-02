from pptx import Presentation
import os
import json
import time
from pptx.util import Pt
import string

# ----------------- Configuración -----------------
texto_mayuscula = True # True para mayúscula, False para minúscula (default: False)
font_size = 24 # Tamaño de fuente (default: 24)
# ----------------- Configuración -----------------


# set yout template here, slide[0] for textbox(title), textbox(slide number), slide[1] for content
hymns = []
hymnBooks = []
# template = "a.pptx"
def main():
    print("""
░       ░░░        ░░  ░░░░  ░░░      ░░
▒  ▒▒▒▒  ▒▒  ▒▒▒▒▒▒▒▒▒  ▒▒  ▒▒▒  ▒▒▒▒  ▒
▓  ▓▓▓▓  ▓▓      ▓▓▓▓▓▓    ▓▓▓▓  ▓▓▓▓  ▓
█  ████  ██  █████████  ██  ███        █
█       ███        ██  ████  ██  ████  █
Welcome to DEXA Sophos by angor.root                                  
""")
    files = os.listdir("DEXHymnBooks")
    print("[ ] Available hymn books:")
    i = 1
    for file in files:
        with open('DEXHymnBooks/'+file, 'r', encoding='utf-8') as file:
            data = json.load(file)
            hymnBooks.append(data)
            print(f"     [{i}] {data['title']} [{data['tag']}]")
            i += 1
    letra = input("letra: ").split()
    return letra
letras = main()

final_PPTX = Presentation("base.pptx")
start_time = time.time()
i = 0
himanrios = input("Himnarios [tag]: ").split()
for letra in letras:
    for hymnBook in hymnBooks:
        if hymnBook['tag'].lower() not in himanrios:
            continue
        print(f"    [H] {hymnBook['title']} Letra: {letra.capitalize()}")
        for hymn in hymnBook['hymns']:
            cleaned_title = hymn['title'].lstrip(string.punctuation)
            if letra.lower() in cleaned_title.lower() and cleaned_title.lower().startswith(letra.lower()):
                print(f"{hymn['number']} {hymn['title']}")
                i += 1
                print("Coro:")
                print(hymn['chorus'])
                print("Versos:")
                for verso in hymn['verses']:
                    print(verso)
                print()


print(f"Total: {i}")


# tiempo de ejecución
print("--- %s seconds ---" % (time.time() - start_time))