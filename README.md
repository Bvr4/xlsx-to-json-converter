# XLSX to JSON converter
Ce programme est un exemple de convertisseur "Excel vers json" avec interface graphique.
Il permet de générer des json respectant un format attendu par une administration ou autre, en renseignant un fichier Excel en entrée.  
Il est inspiré d'un programme que j'ai réalisé pour un client, le programme ici est une version simplifié pour ne garder que l'essentiel et comprendre le fonctionnement général. Les champs et valeurs renseignées sont totalement fictifs et donnés à titre d'exemple.

## Fichier en entrée
Le fichier XLSX en entrée comprends deux feuilles.  
La première contient une liste de travaux, qui pourrait être extraite d'un outil de gestion interne. Je suis parti du cas où l'on traite des travaux réalisées uniquement pour des communes. On y retrouve des informations telles que l'identifiant, l'insee, le nom de la commune, la date de commande, la date de réception, le type de travaux et le montant HT.
La deuxième feuille contient la liste des financeurs. On y retrouve l'identifiant Travaux (pour faire le lien avec la première liste), l'identité du financeur, le type de financeur, le montant et taux de participation.

## Fichier en sortie
En sortie est généré un fichier json reprenant les informations contenues dans le fichier en entrée, en respectant un formalisme imposé.
