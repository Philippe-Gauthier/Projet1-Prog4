# Extension de Mistletoe pour Markdown

Cette extension ajoute une variétée de fonctions au language Markdown.

## Bibliothèques nécessaires

Pour utiliser le programme, il faut installer :

- **mistletoe** : convertit le Markdown en HTML;
- **python-docx** : permet de lire les documents Word. Dans le code, cette bibliothèque est importée sous le nom `docx`.

Les modules `re`, `pathlib` et `textwrap` sont inclus avec Python et n'ont pas besoin d'être installés.

## Installation

1. Ouvrir PowerShell ou le terminal de VS Code.
2. Installer les bibliothèques avec cette commande :

   ```powershell
   py -m pip install mistletoe python-docx
   ```

3. Vérifier que l'installation a fonctionné :

   ```powershell
   py -c "import mistletoe, docx; print('Installation réussie')"
   ```

Si la commande `py` n'est pas reconnue, la remplacer par `python` dans les deux commandes :

```powershell
python -m pip install mistletoe python-docx
python -c "import mistletoe, docx; print('Installation réussie')"
```

## Utilisation

```
python script.py
```
Vous avez besoin des fichiers suivants: 

1. "MD_integration.md", ce fichier contient le contenu que vous avez écrit pour cette extension.
2. Fichier word "Test.docx"
3. Les images que vous importer (au besoin)
4. Le fichier css "style.css"

## Fonctionnalité

Voici une liste des différentes fonctionnalitées ajoutées par cette extension.

1. Diapositive
2. Table des matières
3. Arborescence du dossier
4. Liste à cocher
5. Texte de couleur + surlignage
6. Centrage de contenu
7. Gestion d'image
8. Conversion de fichier word


## Macros

|Fonctions| Macro                       | Effet                                                                 |
|------------------------------| --------------------------- | --------------------------------------------------------------------- |
|Diapositive| `Slide::`                 | Début d'une nouvelle diapositive                                     |
|Personalisation de diapositive| `??? css ???`             | Style CSS de la diapositive (ex.`??? background-color: black; ???`) |
|Table des matières| `**contenu:**`            | Remplacé par une table des matières (titres`##` à `######`)    |
|Arborescence de dossier| `!!{chemin} {profondeur}`     | Remplacé par l'arborescence du dossier (profondeur par défaut : 1)  |
|Liste à cocher| `/// texte`               | Case à cocher                                                        |
|Texte coloré| `{{couleur|texte}}`       | Texte en couleur                                                      |
|Texte surligné| `{{=couleur|texte}}`      | Texte surligné                                                       |
|Texte coloré et surligné| `{{couleur,=fond|texte}}` | Texte en couleur + surlignage                                         |
|Centrage de contenu| `(MD)` … `(MD)`            | Centre le bloc de texte entre les deux`(MD)`                          |
|Gestion d'image| `@@@` … `@@@`          | Insère une image (une propriété`clé: valeur` par ligne)         |

### Slides

- Chaque diapositive peut afficher 26 lignes en même temps.
- Le style (Couleur de fond, couleur de texte, style de font, taille du font etc..) de chaque diapositive est personalisable à l'aide de la syntaxe suivante:
```
???{contenu style HTML}???
```
- Il faut ouvrir en utilisant "???", mettre le contenu du style de la diapositive suivant la syntax et les propriétés disponibles de HTML; Reference: https://www.w3schools.com/html/html_styles.asp
- Exemple:
```
???background-color: yellow; color: white???
```

### Tableau de matières

Ajouter la macro `**contenu:**` au début du fichier markdown et la table des matières se créée elle même.

### Arborescence de dossier

Ajouter la macro `!!path={chemin} ; depth={profondeur} ; blacklist=["{objet1}", "{objet2}", "{objet3}"]!!`

- Le chemin est le chemin vers le dossier voulue. | EX: C:\Users\mammouth\Documents\Github
- La profondeur de l'arborescence (en partant du dossier voulue). | EX: Profondeur = 3, C:\Users\mammouth\Documents\Github\ Dossier 1 \ Dossier 2 \ Dossier 3
- La blacklist qui permet d'ignorer certain types de dossier ou fichier.

### Liste à cocher

Ajouter la macro: /// et insérer ensuite un titre. Cela crée une boite à cocher avec un titre associé | EX: /// {Titre}

### Texte en Couleur

- Suivre la syntaxe: `{{couleur|texte}}`, vous donne un texte dans la couleur de votre choix.
- La syntaxe: `{{=couleur|texte}}` affiche du texte avec un surlignage dans la couleur voulue.
- Vous pouvez aussi combiné ces deux fonctions en utilisant la syntaxe: `{{couleur,=fond|texte}}`, cela vous donne un texte en couleur de votre choix et surligné avec la couleur préscrit 
- Noms de couleur acceptés : `red, blue, green, yellow, orange, purple, pink, black, white, gray`, ou un code hex `#RRGGBB`.

### Centrage de contenu

- L'utilisation de la macro `(MD) *Saut de ligne* {Contenu}  *Saut de ligne* (MD)` permet de centré le contenu trouvé à l'intérieur des deux symboles de macro.

Exemple: 

(MD)

{Contenu}

(MD)

- Si vous voulez centré une arborescence de dossier il faut utiliser la macro (centered_Tree). Il faut seulement appeler la macro une fois par arborescence.

Exemple: 

(centered_Tree)

{Contenu}

### Images

```
@@@
src: images/mammouth.jpg
width: 300px
border-radius: 10px
alt: Une description
@@@
```

- La propriété: `src` (le chemin vers la photo que vous voulez importer, ou le lien web) est obligatoire. Les autres clés doivent être des propriétés CSS valides (sauf `alt`).
- Reference: https://www.w3schools.com/css/css3_images.asp

## Conversion de fichier word

- Si le fichier word `test.docx` existe dans le même dossier, il est d'abord converti en Markdown et ajouté à la fin de `MD_integration.md` sous forme de diapositive (entre `<!-- SAID_START -->` et `<!-- SAID_END -->`). Les images sont extraites dans `images/`.
- Ensuite `MD_integration.md` est converti en `HTML_integration.html`.

## Limites et avertissements

- Chaque diapositive ne peut afficher que 26 lignes en même temps. Si la diapositive contient plus de 26 lignes, une scrollbar seras créée 
- Ne jamais entourer `Slide::` avec `(MD)`, cela cause une erreur.
- Un `(MD)` non fermé constitue une erreur.
- Une couleur invalide ou une propriété CSS inconnue lève une erreur. Vérifier la syntaxe.

