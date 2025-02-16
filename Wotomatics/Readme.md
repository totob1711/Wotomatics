# Wotomatics - Outils d'automatisation de publipostage et d'envoie d'email
<br>

## Prérequis nécéssaires
- Aucun prérequis n'est nécéssaire (aucun logiciel à installer)  

- Pensez à changer les chemin d'accès (PATH) au fichier et remplacez les par le votre.  
4 PATH à modifier dans le fichier `createFiles.ps1`.  
3 PATH à modifier dans le fichier `sendMails.ps1`.  

- Penez à mofifier les patterns du mail d'envoie (celui actuel est *@mail.fr*)
<br>

## Présentation de l'outil
### Publipostage : 
- Vous avez la possibilité de générer un publiopstage à partir d'un modèle (fichier word) et d'une base de données (fichier excel ou csv).  
-> voir l'exemple fourni

- Les fichiers du publipostage sont enregistrés sous le nom de *Prénom_Nom.pdf* de la personne concernée.

- A l'avenir il sera surement possible de rajouter des champs selont vos besoins de manières plus simple qu'en modifiant le code.

### Envoie d'e-mails : 
- Vous avez la possibilité de réaliser un envoie de mail une fois les fichiers générés par le publipostage.

- Le serveur utilisé par défaut est le serveur outlook.  
A l'avenir, il sera suremement possible de choisir le serveur SMTP en fonction de la boite mail que vous avez.

- Les mails sont envoyés via un pattern *prenom.nom@email.fr*

### Logs : 
- Un fichier de logs est mis à disposition si jamais des difficultés sont rencontrées dans l'utilisation de l'outil (dans le répertoire `\logs`)  
<br>

## Utilisation
- Pour utiliser l'outil, il vous suffit de vous rendre dans le dossier `Wotomatics` et d'executer la commande `.\createFiles.ps1`  
<br>


## Notes
- Je rappelle que ceci est une première version qui a été développée dans un cadre de stage et que des contraintes étaient imposées. A l'avenir une version plus simple d'utilisation pour tous sera développée.  


<br>
<br>

---
```
 _        _        _     _ _____ _ _ 
| |_ ___ | |_ ___ | |__ / |___  / / |
| __/ _ \| __/ _ \| '_ \| |  / /| | |
| || (_) | || (_) | |_) | | / / | | |
 \__\___/ \__\___/|_.__/|_|/_/  |_|_|
```
> **totob1711** - Architecte du monde numérique, créateur de réalités virtuelles
>> "Un bug n'est pas un échec, mais une invitation à grandir et à apprendre."  
Chaque erreur, une leçon. Chaque solution, un pas de plus dans l’infini.  
> GitHub : [https://github.com/H4kT1m3](https://github.com/totob1711)
---

