# Add-in OVH Exchange Helper

Complément pour serveur Exchange OVH (ex5.mail.ovh.net)

## Installation

1. Installer Node.js si pas déjà fait
2. Dans ce dossier, lancer: `npm install http-server`
3. Démarrer le serveur: `npm start`
4. Ouvrir Outlook Web ou Desktop
5. Aller dans Paramètres > Compléments > Mes compléments
6. Cliquer sur "Ajouter un complément personnalisé"
7. Choisir "Ajouter depuis un fichier" et sélectionner manifest.xml

## Développement

- `npm run dev` : Lance le serveur et ouvre le navigateur
- Le complément sera accessible sur https://localhost:3000

## Fonctionnalités

- ✅ Analyse des emails
- ✅ Export vers systèmes OVH  
- ✅ Informations serveur Exchange
- ✅ Interface utilisateur moderne

## Configuration

Modifiez les URLs dans manifest.xml si nécessaire.
