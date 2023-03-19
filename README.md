###Informations générales####
Ce programme récupère des informations sur une entreprise en utilisant l'API Pappers et les stocke dans un fichier Excel. Les informations récupérées comprennent le nom de l'entreprise, l'adresse, la ville, le code postal, le pays, le domaine d'activité, le dirigeant, la date de naissance du dirigeant, le SIRET, la forme juridique, le numéro de TVA intracommunautaire, le numéro RCS, le capital social, le code APE, la date de clôture comptable et le bénéficiaire.

###Installation et exécution###
1- Installer Python 3.x et pip sur votre ordinateur si ce n'est pas déjà fait.
2- Télécharger ou cloner le dépôt GitHub contenant le code source.
3- Dans un terminal, naviguer jusqu'au dossier contenant le code source.
4- Installer les dépendances du programme en exécutant la commande suivante : pip install -r requirements.txt
5- Renommer le fichier .env.sample en .env et ajouter votre clé API dans la variable API_KEY.
6- Exécuter le programme en tapant python entreprise_info.py dans le terminal.

###Utilisation###
1- Le programme demande d'entrer le numéro SIREN de l'entreprise à rechercher.
2- Si le fichier Excel spécifié n'existe pas, un nouveau fichier sera créé avec les en-têtes de colonnes appropriés. Si le fichier existe, le programme ajoutera les informations récupérées à la fin du fichier.
3- Une fois que les informations ont été ajoutées au fichier Excel, le programme affichera un message de confirmation.

###Contribuer###
Si vous souhaitez contribuer à ce projet, veuillez cloner le dépôt GitHub et créer une branche pour vos modifications. Après avoir effectué les modifications, veuillez créer une demande d'extraction (pull request) pour que vos modifications soient examinées et intégrées au projet principal.

###Auteur###
Ce programme a été écrit par Samuel Shemtov.
