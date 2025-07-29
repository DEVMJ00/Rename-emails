# Rename-emails
## 📨 Renommage automatique de fichiers `.msg` via une interface graphique (PowerShell)

## 📋 Description

Ce script PowerShell permet de renommer facilement des fichiers email au format `.msg` à l’aide d’une interface graphique simple.  
Il est conçu pour automatiser le classement de courriels exportés, en leur appliquant un préfixe descriptif et une date, selon un format prédéfini.

---

## 🎯 Fonctionnalités

- Interface graphique Windows Forms
- Sélection automatique des fichiers `.msg` présents dans le dossier courant
- Liste déroulante de catégories prédéfinies (ex. : relances, devis, factures, etc.)
- Renommage du fichier au format : `YYYYMMDD_Prefixe.extension`
- Retour d’état visuel : succès ou erreur
- Rafraîchissement automatique de la liste après chaque renommage

---

## 🛠️ Pré-requis

- **Système** : Windows
- **Logiciel** : PowerShell v5.1 ou supérieur
- **Environnement** : Exécution locale avec interface graphique

---

## ▶️ Utilisation

1. **Copiez le script dans un fichier** nommé par exemple `renommage_msg.ps1`
2. Placez-le **dans le dossier contenant vos fichiers `.msg`**
3. Exécutez le script via PowerShell :
   ` .\renommage_msg.ps1`
4. Dans la fenêtre :

Choisissez le fichier à renommer
Sélectionnez le type de document (1ere relance, devis, facture, etc.)
Cliquez sur `Renommer`
Le fichier sera renommé avec la date du jour et le préfixe choisi (ex : 20250729_Devis.msg)

---
## 🚫 Limitations
Ne prend en charge que les fichiers .msg (format Outlook)
Le script doit être lancé depuis le dossier contenant les fichiers
La fenêtre est en taille fixe, non redimensionnable

---
## 👨‍💻 Auteur
DEVMJ – [En reconversion vers le développement logiciel]
Script PowerShell réalisé dans un but d’automatisation personnelle et professionnelle.


---
##📜 Licence
Ce projet est distribué sans licence – libre d'utilisation et de modification.
