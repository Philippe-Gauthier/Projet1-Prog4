import mistletoe
import re
from pathlib import Path
from textwrap import dedent
from mistletoe.block_token import Heading
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.text.run import Run
from docx.oxml.ns import qn
import ast

# Raccourci utilisé pour déclencher le style personnalisé dans le markdown
SHORTCUT = "!!"
DOSSIER = Path(__file__).resolve().parent
FILE = DOSSIER / "MD_integration.md"
OUTPUT_FILE = DOSSIER / "HTML_integration.html"

# Antoine

# Liste des couleurs acceptées
COULEURS_VALIDES = {
    "red", "blue", "green", "yellow", "orange",
    "purple", "pink", "black", "white", "gray"
}

def couleur_valide(couleur):
    # Accepte une couleur présente dans la liste
    if couleur.lower() in COULEURS_VALIDES:
        return True

    # Accepte un code hexadécimal comme #FF0000
    if re.fullmatch(r"#[0-9a-fA-F]{6}", couleur):
        return True

    return False

def ajouter_style(match):
    # Récupère le groupe correspondant au style (ex: couleur texte / fond)
    style = match.group(1)
    # Récupère le texte auquel le style doit être appliqué
    texte = match.group(2)

    # Initialisation des couleurs (texte et fond) à vide par défaut
    couleur_texte = ""
    couleur_fond = ""

    # Cas où le style définit à la fois une couleur de texte et une couleur de fond
    # Format attendu : "couleur_texte,=couleur_fond"
    if ",=" in style:
        couleur_texte, couleur_fond = style.split(",=", 1)
        couleur_texte = couleur_texte.strip()
        couleur_fond = couleur_fond.strip()

    elif style.strip().startswith("="):
        couleur_fond = style.strip()[1:].strip()

    else:
        couleur_texte = style.strip()
    # Vérifie les couleurs après avoir retiré les espaces
    if couleur_texte and not couleur_valide(couleur_texte):
        print(f"Erreur : couleur de texte invalide : {couleur_texte}")
        return texte

    if couleur_fond and not couleur_valide(couleur_fond):
        print(f"Erreur : couleur de fond invalide : {couleur_fond}")
        return texte
    # Construction de la chaîne de style CSS inline
    style_html = ""

    # Ajoute la couleur du texte si elle est définie
    if couleur_texte:
        style_html += f"color: {couleur_texte}; "

    # Ajoute la couleur de fond si elle est définie
    if couleur_fond:
        style_html += f"background-color: {couleur_fond}; "

    # Retourne le texte enveloppé dans un span avec le style appliqué
    return f'<span style="{style_html}">{texte}</span>'

# Amé
def checklistMD(texte):

    # Liste qui contiendra toutes les lignes (modifiées ou non) du fichier
    modifiedLines = []

    # Séparer les lignes du texte
    lignes = texte.splitlines(keepends=True)

    for line in lignes:
        # Liste temporaire des caractères/éléments pour construire la ligne modifiée
        modifiedLine = []

        # Vérifie si la ligne est une ligne de checklist (contient '+' ou '&')
        if  (line.find('+') != -1) or (line.find('&') != -1):
            # Parcourt chaque caractère de la ligne
            for letter in line:
                if letter == "+":
                    # Insère la balise HTML de la checkbox et le label au début de la ligne
                    modifiedLine.append('<input type="checkbox"> <label>')
                elif letter == "&":
                    # Insère la balise HTML de la checkbox et le label au début de la ligne
                    modifiedLine.append('<input type="checkbox" checked> <label>')
                else:
                    modifiedLine.append(letter)

            # Si la ligne se termine par un saut de ligne, remplace le dernier élément
            # par la fermeture du label suivie du saut de ligne
            if line.endswith('\n'):
                modifiedLine[-1] = '</label><br>\n'
            else:
                # Sinon, ajoute simplement la fermeture du label à la fin
                modifiedLine.append('</label><br>\n')

            # Cette ligne doit être ici
            # Reconstitue la ligne complète et l'ajoute à la liste des lignes modifiées
            modifiedLines.append("".join(modifiedLine))

        else:
            # Ligne normale (pas de checklist) : ajoutée telle quelle
            modifiedLines.append(line)

    # Retourne le contenu complet reconstitué sous forme de chaîne unique
    return "".join(modifiedLines)

# Bruno
def convertir_diapositive(texte, fichier_html):

    # Découpe le texte en diapositives à partir du séparateur "Slide::"
    slides = texte.split("Slide::")
    # Chaîne qui accumulera le HTML final de toutes les diapositives
    resultat = ""
    #Nombre de lignes MAX
    MAX_LIGNES = 26

    # Feuille de style CSS appliquée aux diapositives et à l'arborescence de fichiers

    # Parcourt chaque diapositive extraite du texte
    for numero_slide, slide in enumerate(slides[0:], start=0):
          # Compte le nombre de lignes dans la diapositive
        nombre_lignes = len(slide.strip().splitlines())

        #Gestion des avertissements 
        if nombre_lignes > MAX_LIGNES:
            print(f"AVERTISSEMENT : La diapositive {numero_slide} dépasse la limite de lignes.")
            print(f"Nombre de lignes : {nombre_lignes}")
            print(f"Limite permise : {MAX_LIGNES}")
            print("Une scrollbar sera créée pour permettre de faire défiler le contenu.")

        # Convertit le contenu markdown de la diapositive en HTML
        rendu = mistletoe.markdown(slide)

        # Ajoute les id aux titres pour les liens de la table des matières
        rendu = ajouter_id_titres(rendu)

        # Antoine : couleurs
        # Applique le style personnalisé (couleurs) via la fonction ajouter_style
        rendu = re.sub(r"\{\{([^|]+)\|(.+?)\}\}", ajouter_style, rendu)
        if "???" in rendu:
            slide_style = f'style="{rendu.strip().split("???")[1].split("???")[0].strip()}"'
            print(f"Slide style detected: {slide_style}")
        else:
            slide_style = 'style="background-color: white;"'

        ## we split the rendu from the shortcut ??slidebg: {color}?? and then we use strip to remove whitespace to get consistent results. if the shortcut is not found, we set it to white

        rendu = re.sub(r"\?\?\?.+?\?\?\?", "", rendu).strip()
        ## we take out the shortcut from the final rendu so that it doesnt show up as text
        ## need to do \? because ? is a special char

        # Enveloppe le rendu HTML de la diapositive dans une div avec la classe "slide"
        slide_html = f'<div class="slide" {slide_style}>' + rendu + '</div>'

        # Ajoute la diapositive au résultat final
        resultat += slide_html


    css = f"""<!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8"/>
        <style>
        .file-tree {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            font-size: 14px;
            line-height: 1.8;
        }}

        .file-tree.centered-tree {{
            width: fit-content;
            margin: 0 auto;
        }}

        /* Reset and indent nested folders */
        .file-tree ul {{
            list-style-type: none;
            padding-left: 20px;
            margin: 0;
            position: relative;
            overflow: auto;
        }}

        /* Vertical branch line for connecting sub-folders/files */
        .file-tree ul::before {{
            content: "";
            position: absolute;
            top: 0;
            left: 7px;
            bottom: 12px;
            border-left: 2px solid #b8b8b8;
        }}

        /* Individual list items */
        .file-tree li {{
            margin: 0;
            padding: 3px 0 3px 15px;
            position: relative;
        }}

        /* Horizontal branch lines pointing to folders/files */
        .file-tree li::before {{
            content: "";
            position: absolute;
            top: 13px;
            left: -8px;
            width: 15px;
            height: 1px;
            border-top: 2px solid #b8b8b8;
        }}

        /* Stops the vertical line at the last item of a directory level */
        .file-tree li:last-child::before {{
            background: transparent;
            height: 1px;
        }}

        /* Folder styling */
        .folder {{
            font-weight: 600;
        }}

        .folder::before {{
            content: "📁 ";
        }}

        /* File styling */
        .file::before {{
            content: "📄 ";
        }}

        .slide {{
            background-color: rgb(100, 100, 100);
            width: 100vw;
            height: 70vh;
            margin-bottom: 30px;
            overflow: auto;
            }}
        </style>
    </head>
    <body>
        {resultat}
    </body>
    </html>
    """
    with open(fichier_html, 'w', encoding='utf-8') as fout:
            fout.write(css)




#jay
def CenterText(File):
    found = False
    ERROR = False
    modified_lines = []

    for line in File.splitlines(keepends=True):

        if line.lstrip().startswith("(MD)") and not found: # Vérifie si c'est le début d'un bloc centré et qu'on n'est pas déjà dans un bloc centré
            modified_lines.append('<div align="center">\n') # Met la ligne modifié avec la balise de fermeture <div align="center"> dans modified_lines 
            found = True

        elif line.lstrip().startswith("Slide::") and found: # Vérification si l'utilisateur ne tente de centrer une diapositive
            found = False # Reset de la détection
            ERROR = True # Mise en erreur du code et demande de correction
            break # Si on rencontre une nouvelle diapositive avant de fermer le bloc centré, on sort de la boucle

        elif line.lstrip().startswith("(MD)") and found: # Vérifie si nous somme déja dans un bloc centré à fermer
            modified_lines.append('</div>\n') # Met la ligne modifié avec la balise de fermeture </div> dans modified_lines
            found = False

        else:
            modified_lines.append(line) # Met la ligne non modifiée dans modified_lines

    if found:
        raise ValueError("(MD) manquant dans la dernière slide")
    
    elif ERROR:
        raise ValueError("Vérifiez la syntaxe de vos blocs centrés, un (MD) a été ouvert mais pas fermé avant une diapositive. \n\nL'erreur ressemble probablement à ceci dans le fichier MD_integration.md :\n\n(MD) \nSlide::\n\n")

    return ''.join(modified_lines)


#Nico

def creer_lien(titre): 
    # Mettre le titre en minuscules 
    lien = titre.lower() 
    
    # Remplacer les espaces par des tirets 
    lien = lien.replace(" ", "-") 
    
    # Enlever les caractères spéciaux 
    lien = re.sub(r"[^a-z0-9\-]", "", lien)
    
    return lien

def creer_table_matiere(texte):

    # Vérifier si le marqueur existe
    if "**contenu:**" not in texte.lower():
        return texte

    print("Table des matières créée.")

    # Initialise la table des matières avec son titre en markdown
    tableMatieres = "## Table des matières\n\n"

    # Parser le document
    document = mistletoe.Document(texte)

    # Parcourt les éléments de premier niveau du document markdown
    for titreMarkdown in document.children:
        # Ne traite que les éléments qui sont des titres (Heading)
        if isinstance(titreMarkdown, Heading):

            # Les titres ## à ###### sont ajoutés à la table
            if titreMarkdown.level >= 2:

                # Chaîne qui accumulera le texte complet du titre
                texteTitre = ""

                # Reconstruit le texte du titre à partir de ses sous-éléments
                for partieTitre in titreMarkdown.children:
                    if hasattr(partieTitre, "content"):
                        texteTitre += partieTitre.content

                # Crée le lien d'ancre en remplaçant les espaces par des tirets
                lienTitre = creer_lien(texteTitre)

                # Ajouter le titre dans la table
                # L'indentation dépend du niveau du titre (##, ###, ####, etc.)
                if titreMarkdown.level == 2:
                    tableMatieres += f"- [{texteTitre}](#{lienTitre})\n"

                elif titreMarkdown.level == 3:
                    tableMatieres += f"  - [{texteTitre}](#{lienTitre})\n"

                elif titreMarkdown.level == 4:
                    tableMatieres += f"    - [{texteTitre}](#{lienTitre})\n"

                elif titreMarkdown.level == 5:
                    tableMatieres += f"      - [{texteTitre}](#{lienTitre})\n"

                elif titreMarkdown.level == 6:
                    tableMatieres += f"        - [{texteTitre}](#{lienTitre})\n"

    # Remplacer **contenu:** par la table des matières
    New_texte = texte.replace("**contenu:**", tableMatieres, 1)

    return New_texte

def ajouter_id_titres(html): 
    """ Ajoute un id aux titres h2 à h6
    pour permettre aux liens de la table des matières 
    de fonctionner. 
    """ 
    
    def remplacer_titre(match): 
        
        niveau = match.group(1)
        titre = match.group(2)
        
        # Créer le même lien que dans la table des matières 
        lien = creer_lien(titre)
        
        return f'<h{niveau} id="{lien}">{titre}</h{niveau}>' 
    
    # Chercher les titres h2 à h6 
    html = re.sub( r'<h([2-6])>(.*?)</h\1>', remplacer_titre, html ) 
    
    return html

# Zach
def build_tree(path: str, max_depth,blacklist, current_depth=1, ) -> dict | list | str:
    """
    Reads a folder and builds a tree consisting of all the files at a certain depth.
    """

    # S'assure que max_depth est bien un entier
    max_depth = int(max_depth)

    # Cas particulier : si le chemin correspond au motif générique, on le remplace par un chemin vide
    if path in [r'**\\**', r'**\\**'[0]]:
        path = r''

    # Convertit la chaîne de chemin en objet Path
    path = Path(path)

    # Si le chemin pointe directement vers un fichier, retourne simplement son nom
    if path.is_file():
        return path.name

    # Si on a atteint la profondeur maximale, retourne uniquement la liste des noms
    # des fichiers/dossiers à ce niveau (en ignorant les fichiers cachés)
    if current_depth == max_depth:
        try : 
            list = []
            for item in path.iterdir():
                for name in blacklist:
                    if name.endswith("*"):
                        if item.name.startswith(name[:-1]):
                            break
                    elif item.name == name:
                        break
                else:
                    list.append(item.name)
                    
            return list
        except PermissionError:
            return ""
    # Dictionnaire qui représentera l'arborescence à ce niveau
    tree = {}

    # Parcourt chaque élément (fichier ou dossier) du chemin courant
    try:
        for item in path.iterdir():
            for name in blacklist:
                if name.endswith("*"):
                    if item.name.startswith(name[:-1]):
                        break
                elif item.name == name:
                    break
            else:
                
                if item.is_dir():
                    tree[item.name] = build_tree(item, max_depth, blacklist, current_depth + 1)
                else:
                    tree[item.name] = item.name
    except PermissionError:
        return tree
    return tree   


def build_html(tree: dict | list | str) -> str:
    """
    Builds the html tree fragments
    """

    # Chaîne HTML accumulée qui sera retournée
    html = ''

    # Cas où l'arbre est une liste de fichiers (niveau de profondeur maximale atteint)
    if isinstance(tree, list):
        for item in tree:
            html += f'<li><span class="file">{item}</span></li>'

    # Cas où l'arbre est directement une chaîne (un seul fichier)
    elif isinstance(tree, str):
        html += f'<li><span class="file">{tree}</span></li>'

    # Cas où l'arbre est un dictionnaire (dossier contenant fichiers et/ou sous-dossiers)
    elif isinstance(tree, dict):
        for folder_name, content in tree.items():

            # Contenu = liste de fichiers -> sous-dossier avec ses fichiers
            if isinstance(content, list):
                html += f'<li><span class="folder">{folder_name}</span><ul>'
                html += build_html(content)
                html += '</ul></li>'

            # Contenu = chaîne -> simple fichier
            elif isinstance(content, str):
                html += f'<li><span class="file">{content}</span></li>'

            # Contenu = dictionnaire -> sous-dossier avec sa propre arborescence
            elif isinstance(content, dict):
                html += f'<li><span class="folder">{folder_name}</span><ul>'
                html += build_html(content)
                html += '</ul></li>'

    return html


def render_tree_block(tree_dict):
    """
    Génère uniquement le bloc conteneur de l'arbre
    """

    # Construit le fragment HTML représentant l'arborescence
    corps_arbre = build_html(tree_dict)

    # Enveloppe le tout dans une div conteneur avec une liste HTML
    return f'<div class="file-tree"><ul>{corps_arbre}</ul></div>'





#Will
# Creation d'un gabarit HTML pour le rendu final
PAGE_TEMPLATE = """<!DOCTYPE html>
<html>
<head>
<style>
</style>
</head>
<body>
{body}
</body>
</html>
"""

# Liste des propriétés CSS valides reconnues
# Pour faire un check des propriétés CSS, on peut utiliser cette liste pour valider les entrées de l'utilisateur
VALID_CSS_PROPERTIES = {
    "align-content", "align-items", "align-self", "all", "animation",
    "animation-delay", "animation-direction", "animation-duration",
    "animation-fill-mode", "animation-iteration-count", "animation-name",
    "animation-play-state", "animation-timing-function", "backdrop-filter",
    "background", "background-attachment", "background-blend-mode",
    "background-clip", "background-color", "background-image",
    "background-origin", "background-position", "background-repeat",
    "background-size", "border", "border-bottom", "border-bottom-color",
    "border-bottom-left-radius", "border-bottom-right-radius",
    "border-bottom-style", "border-bottom-width", "border-collapse",
    "border-color", "border-image", "border-left", "border-left-color",
    "border-left-style", "border-left-width", "border-radius",
    "border-right", "border-right-color", "border-right-style",
    "border-right-width", "border-spacing", "border-style", "border-top",
    "border-top-color", "border-top-left-radius", "border-top-right-radius",
    "border-top-style", "border-top-width", "border-width", "bottom",
    "box-shadow", "box-sizing", "caption-side", "caret-color", "clear",
    "clip", "clip-path", "color", "column-count", "column-gap",
    "column-rule", "column-width", "columns", "content", "cursor",
    "direction", "display", "filter", "flex", "flex-basis",
    "flex-direction", "flex-flow", "flex-grow", "flex-shrink", "flex-wrap",
    "float", "font", "font-family", "font-size", "font-style",
    "font-variant", "font-weight", "gap", "grid", "grid-area",
    "grid-column", "grid-column-end", "grid-column-start", "grid-gap",
    "grid-row", "grid-row-end", "grid-row-start", "grid-template",
    "grid-template-areas", "grid-template-columns", "grid-template-rows",
    "height", "inset", "justify-content", "justify-items", "justify-self",
    "left", "letter-spacing", "line-height", "list-style",
    "list-style-image", "list-style-position", "list-style-type", "margin",
    "margin-bottom", "margin-left", "margin-right", "margin-top",
    "max-height", "max-width", "min-height", "min-width", "object-fit",
    "object-position", "opacity", "order", "outline", "outline-color",
    "outline-offset", "outline-style", "outline-width", "overflow",
    "overflow-x", "overflow-y", "padding", "padding-bottom",
    "padding-left", "padding-right", "padding-top", "perspective",
    "pointer-events", "position", "resize", "right", "rotate", "scale",
    "table-layout", "text-align", "text-decoration", "text-indent",
    "text-overflow", "text-shadow", "text-transform", "top", "transform",
    "transform-origin", "transition", "transition-delay",
    "transition-duration", "transition-property",
    "transition-timing-function", "translate", "user-select",
    "vertical-align", "visibility", "white-space", "width", "word-break",
    "word-spacing", "word-wrap", "z-index", "src", "alt", "title",
    "poster", "preload", "controls", "autoplay", "loop", "muted", "playsinline"
}

def preprocess(text):
    """
    Trouve les blocs d'image @@@ et les remplace par du HTML.

    Syntaxe:

        @@@
        src: ocean.jpg
        width: 300
        height: 300
        margin-left: 50px
        margin-top: 20px
        border-radius: 10px
        @@@

    Chaque ligne est une propriété sous la forme:

        propriété: valeur
    """

    # Analyse le texte ligne par ligne
    lines = text.splitlines()
    output = []

    i = 0

    while i < len(lines):

#######################################PHOTO#####################################################################################################################
        # Cherche le début d'un bloc @@@
        # Une fois trouvé, on cherche le prochain @@@ pour marquer la fin du bloc 
        if lines[i].strip() in ("@@@", ">>>"):

            # Détermine le type selon le délimiteur
            delimiter = lines[i].strip()
            media_type = "image" if delimiter == "@@@" else "video"

            # Recherche le prochain @@@ qui marque la fin du bloc
            j = i + 1

            while j < len(lines) and lines[j].strip() != delimiter:
                j += 1

            # Aucun @@@ de fermeture trouvé
            if j >= len(lines):
                raise ValueError(
                    f"@@@ ou >>> Invalide a la ligne {i + 1}: "
                    "@@@ ou >>> de fermeture attendue."
                )

            # Dictionnaire contenant les propriétés de l'image
            properties = {}

            # Analyse chaque ligne du bloc
            for line in lines[i + 1:j]:

                # Ignore les lignes vides
                line = line.strip()

                if not line:
                    continue

                # Vérifie que la ligne contient ":"
                if ":" not in line:
                    raise ValueError(
                        f"Propriétés invalide dans le bloc @@@/>>>: {line}"
                    )

                # Sépare la propriété et sa valeur
                # On détermine la clé et la valeur en utilisant le premier ":" trouvé comme séparateur
                key, value = line.split(":", 1)

                # Nettoie les espaces
                key = key.strip()
                value = value.strip()

                if key == "alt":
                    alt_description = value
                    continue

                if key == "title":
                    title_description = value
                    print("title detected")
                    continue

                # Vérifie que la clé est une propriété CSS valide
                if key not in VALID_CSS_PROPERTIES:
                    raise ValueError(
                        f"Propriété CSS inconnu dans le bloc @@@/>>> "
                        f"a la ligne {i + 1}: '{key}'"
                    )

                # Ajoute la propriété au dictionnaire
                """
                Cela ressemble a:
                {
                    "width": "300px"
                }
                """
                properties[key] = value

            # L'attribut src est obligatoire (ofc lol)
            if "src" not in properties:
                raise ValueError(
                    f"@@@ ou >>> Invalide a la ligne {i + 1}: "
                    "'src' manquant."
                )

            # Récupère le chemin de l'image
            src = properties.pop("src")

            # Construction des attributs CSS
            css_properties = []

            # Transforme les propriétés en attributs CSS
            for key, value in properties.items():
                css_properties.append(f"{key}: {value}")

            # Transforme les propriétés CSS en une seule chaîne
            style = "; ".join(css_properties)

            if media_type == "image":
                print("Image Detected")
                html = f'<img src="{src}"'
            else:
                print("Video Detected")
                html = f'<video src="{src}" controls'

            if style:
                html += f' style="{style};"'

            if alt_description:
                html += f' alt="{alt_description}"'

            alt_description = ""  # Reset pour pas leak
            title_description = ""  # Reset pour pas leak

            # Pour fermer la balise img, on ajoute le ">" à la fin
            html += '>'

            # Ajoute le HTML à la sortie
            output.append(html)

            # Saute tout le bloc, y compris les deux @@@
            i = j + 1

        else:
            # Ligne normale de Markdown
            output.append(lines[i])
            i += 1

    # Reconstitue le texte final
    return "\n".join(output)

# SAID 
def preparer_markdown_depuis_word(chemin_word):

    fichier_md = Path(FILE)
    dossier_images = fichier_md.parent / "images"
    dossier_images.mkdir(parents=True, exist_ok=True)
    document = Document(chemin_word)

    def convertir_run(element, paragraphe):
        run = Run(element, paragraphe)
        texte = run.text

        if run.bold and run.italic:
            texte = f"***{texte}***"
        elif run.bold:
            texte = f"**{texte}**"
        elif run.italic:
            texte = f"*{texte}*"

        for image_xml in element.iter(qn("a:blip")):
            relation_id = image_xml.get(qn("r:embed"))

            if relation_id:
                image = paragraphe.part.rels[relation_id].target_part
                nom = Path(image.partname).name
                donnees = image.blob

                chemin = dossier_images / nom
                compteur = 1

                # Chercher un nom libre si une image différente existe déjà
                while chemin.exists() and chemin.read_bytes() != donnees:
                    chemin = dossier_images / f"{Path(nom).stem}_{compteur}{Path(nom).suffix}"
                    compteur += 1

                # Sauvegarder seulement si l'image n'existe pas
                if not chemin.exists():
                    chemin.write_bytes(donnees)

                # Utiliser le nom réel dans le Markdown
                texte += f"\n\n![{chemin.name}](images/{chemin.name})\n\n"

        return texte

    def convertir_paragraphe(paragraphe):
        resultat = ""

        for element in paragraphe._p:
            if element.tag == qn("w:r"):
                resultat += convertir_run(element, paragraphe)

            elif element.tag == qn("w:hyperlink"):
                texte = "".join(
                    convertir_run(run, paragraphe)
                    for run in element.findall(qn("w:r"))
                )

                relation_id = element.get(qn("r:id"))

                if relation_id:
                    url = paragraphe.part.rels[relation_id].target_ref
                    resultat += f"[{texte}]({url})"
                else:
                    resultat += texte

        return resultat

    def convertir_tableau(tableau):
        lignes = []

        for numero, ligne in enumerate(tableau.rows):
            cellules = []

            for cellule in ligne.cells:
                texte = " ".join(
                    convertir_paragraphe(p).strip()
                    for p in cellule.paragraphs
                )

                texte = texte.replace("|", "\\|")
                texte = texte.replace("\n", "<br>")
                cellules.append(texte)

            lignes.append("| " + " | ".join(cellules) + " |")

            if numero == 0:
                lignes.append(
                    "| " + " | ".join("---" for _ in cellules) + " |"
                )

        return "\n".join(lignes) + "\n\n"

    markdown = ""
    liste_en_cours = False

    for element in document.element.body.iterchildren():

        if element.tag == qn("w:p"):
            paragraphe = Paragraph(element, document)
            texte = convertir_paragraphe(paragraphe).strip()

            if not texte:
                continue

            style = paragraphe.style.name

            if style.startswith("Heading ") and style[8:].isdigit():
                niveau = int(style[8:])

                if 1 <= niveau <= 6:
                    if liste_en_cours:
                        markdown += "\n"

                    markdown += "#" * niveau + " " + texte + "\n\n"
                    liste_en_cours = False
                    continue

            if style.startswith("List Bullet"):
                if not liste_en_cours and markdown and not markdown.endswith("\n\n"):
                    markdown += "\n"

                markdown += "- " + texte + "\n"
                liste_en_cours = True

            elif style.startswith("List Number"):
                if not liste_en_cours and markdown and not markdown.endswith("\n\n"):
                    markdown += "\n"

                markdown += "1. " + texte + "\n"
                liste_en_cours = True

            else:
                if liste_en_cours:
                    markdown += "\n"

                markdown += texte + "\n\n"
                liste_en_cours = False

        elif element.tag == qn("w:tbl"):
            tableau = Table(element, document)
            if liste_en_cours:
                markdown += "\n"

            markdown += convertir_tableau(tableau)
            liste_en_cours = False

    # Intégration sans effacer le Markdown des autres
    debut = "<!-- SAID_START -->"
    fin = "<!-- SAID_END -->"

    bloc_said = (
        f"{debut}\n\n"
        f"Slide::\n\n"
        f"## 7. Test Said\n\n"
        f"{markdown.strip()}\n\n"
        f"{fin}"
    )

    if fichier_md.exists():
        contenu = fichier_md.read_text(encoding="utf-8")
    else:
        contenu = ""

    if debut in contenu and fin in contenu:
        position_debut = contenu.index(debut)
        position_fin = contenu.index(fin) + len(fin)

        contenu = (
            contenu[:position_debut]
            + bloc_said
            + contenu[position_fin:]
        )
    else:
        contenu = contenu.rstrip() + "\n\n" + bloc_said + "\n"

    fichier_md.write_text(contenu, encoding="utf-8")

    print("conversion Word vers Markdown terminée.")

def generer_html_depuis_markdown():
    """
    Fonction principale qui exécute tout le pipeline de conversion :
    lecture du fichier markdown, application de toutes les transformations
    (table des matières, arbre de fichiers, checklist, texte centré),
    puis génération du fichier HTML final.
    """

    # Lecture du fichier Markdown
    # Ouvre le fichier source(.md) (FILE) en lecture et récupère tout son contenu
    with open(FILE, "r", encoding="utf-8") as fichier:
        texte = fichier.read()

    if not texte.strip():
        print("Erreur : le fichier Markdown est vide.")
        return


    # Nico : table des matières
    # Cherche le marqueur "**contenu:**" dans le texte et le remplace par
    # une table des matières générée à partir des titres markdown (## à ######)
       
    texte = creer_table_matiere(texte)

    # Zach : arbre de fichiers
    # Découpe le texte en lignes (en conservant les sauts de ligne)
    lignes = texte.splitlines(keepends=True)
    lignesModifiees = []

    # Indique si le prochain arbre doit être centré
    centrer_arbre = False

    # Parcourt chaque ligne du texte
    for ligne in lignes:


        # Détection de la commande de centrage d'arbre (
        # Détection du raccourci !! pour créer un arbre
        if ligne.startswith(SHORTCUT) and ligne.endswith(SHORTCUT+"\n"):
            # Extrait les paramètres après le raccourci (ex: chemin et profondeur)
            parameters = ligne[len(SHORTCUT):len(ligne)-len(SHORTCUT)-1].strip().split(';')
            real_param={}
            for param in parameters:
                x = param.strip().split('=')
                print(x)
                real_param[x[0]] = x[1]
            
            # Le premier paramètre est le chemin du dossier à explorer
            path = parameters[0]

            # Tente de récupérer la profondeur maximale
            # Si absente ou invalide, utilise 1
            try:
                path = real_param["path"]
            except:
                path = "."

            try:
                depth = int(real_param["depth"])
            except (IndexError, ValueError):
                depth = 1

            try:
                blacklist = ast.literal_eval((real_param["blacklist"]))
            except:
                print("Erreur de lecture de la blacklist d'un arbre")
                blacklist = []


            centrer_arbre =True if real_param.get("centered",False) in ["True","true","1"] else False

            tree_data = build_tree(path, depth,blacklist)
            # Convertit cette structure en bloc HTML (div + liste)
            html_tree = render_tree_block(tree_data)


            # Si (TREE) est détecté avant le shortcut !!, on centre l'arbre dans le HTML final
            if centrer_arbre:
                html_tree = html_tree.replace('<div class="file-tree">','<div class="file-tree centered-tree">')

                # Reset : le prochain arbre ne sera pas centré par défaut
                centrer_arbre = False

            # Remplace la ligne du raccourci par le HTML généré
            lignesModifiees.append(html_tree + "\n")

        else:

            # Ligne normale : conservée telle quelle
            lignesModifiees.append(ligne)

    # Reconstitue le texte complet avec les arbres de fichiers insérés
    texte = "".join(lignesModifiees)

    # Reconstitue le texte complet avec les checklists converties en HTML
    texte = checklistMD(texte)

    # Jay : texte centré
    # Recherche les paires de "()" dans le texte pour ouvrir/fermer
    # des blocs <div align="center"> autour du contenu concerné
    texte = CenterText(texte)

    # Will : image

    texte = preprocess(texte)
    
   


    # Création du HTML final
    # Découpe le texte en diapositives ("Slide::"), convertit chacune en HTML,
    # applique les styles de couleur (Antoine), puis écrit le résultat
    # (CSS + diapositives) dans le fichier de sortie OUTPUT_FILE
    convertir_diapositive(texte, OUTPUT_FILE)
    
    


if __name__ == "__main__":

    fichier_word = DOSSIER / "test.docx"

    if fichier_word.is_file():
        preparer_markdown_depuis_word(fichier_word)
    else:
        print("Aucun document Word trouvé. Utilisation du Markdown existant.")
    
    if FILE.is_file():
        generer_html_depuis_markdown()
    else:
        print("Erreur : MD_integration.md est introuvable.")
