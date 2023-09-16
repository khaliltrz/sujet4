# sujet4
Mon application se compose de deux parties bien distinctes, chacune ayant un rôle essentiel dans la gestion des offres de stage via votre boîte de réception e-mail. 

La première partie qui ce trouve dans les fichiers "extracting data from internships offer in emails.py"et/où dans le notebook "extracting data from internships offer in emails.IPYNB" 
est conçue pour surveiller régulièrement votre boîte de réception, offrant la flexibilité de régler automatiquement la fréquence de cette surveillance grâce à une variable de temps ajustable.
Cette composante télécharge systématiquement tous les e-mails contenant des offres de stage, créant ainsi une base de données comprenant des informations cruciales telles que:
la 'Date', le 'Nom de l'expéditeur', l''Adresse e-mail de l'expéditeur', le 'Destinataire', l’objet, le 'Corps de l'e-mail', et même les pièces jointes, qu'elles soient au format docx, pdf, image ou Excel.
Ensuite, l'application effectue un traitement minutieux de ces données, extrayant des informations pertinentes telles que:
la langue, les adresses e-mail de contact, les URL, les numéros de téléphone et les compétences.

Dès qu'elle détecte un nouvel e-mail contenant une offre de stage, elle envoie automatiquement un e-mail comprenant cette base de données au format Excel. 
Notons que ce processus peut prendre jusqu'à 10 minutes pour télécharger les e-mails depuis 2019, mais il fonctionne en arrière-plan, minimisant ainsi toute interruption.

La deuxième partie de l'application, accessible via une interface web conviviale, permet aux utilisateurs de télécharger rapidement le dataset au format Excel.
Cela est rendu possible grâce à la synchronisation avec la partie de surveillance automatique semi-continue de la boîte e-mail.
De plus, vous avez la possibilité d'ajuster et de modifier la variable de temps en utilisant cette interface web, offrant ainsi un contrôle personnalisé sur la fréquence de la surveillance.
Cette combinaison de fonctionnalités garantit une expérience fluide et efficace pour la gestion des offres de stage directement depuis votre boîte de réception.
