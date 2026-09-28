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