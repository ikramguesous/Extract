**Extraction des tableaux financiers**

**Contexte**

Le projet vise à automatiser l’extraction des tableaux financiers à partir des rapports PDF disponibles sur le site de l’AMMC. Ces rapports présentent des mises en page variées et contiennent plusieurs tableaux (Bilan Actif, Bilan Passif, Compte de Produits et Charges, États des Soldes et Gestion). Le défi est de concevoir une méthode capable de :

1.Scrapper automatiquement les liens des rapports.

2.Télécharger et traiter des fichiers PDF aux formats parfois non standards.

3.Identifier et extraire précisément les tableaux financiers malgré les variations de mise en forme.

**Dépendances et Environnement**

Pour garantir la reproductibilité du projet, nous utilisons :

Python : Version 3.9 (ou supérieure)

PyPDF2 : Pour l’extraction et la normalisation du texte des PDF

pandas : Pour la transformation des données en DataFrames

requests et BeautifulSoup : Pour le scrapping des pages d’index de l’AMMC

**Décision**

Nous avons décidé d’utiliser la bibliothèque PyPDF2 pour extraire le texte des PDF, combinée avec pandas pour transformer les données en DataFrames. La méthode consiste à :

1) Scrapper : Récupérer les URLs des rapports en parcourant les pages d’index du site de l’AMMC (en utilisant requests et BeautifulSoup).
   
2) Télécharger : Extraire les liens des fichiers PDF à l’aide d’expressions régulières.
   
3) Extraire : Utiliser PyPDF2 pour normaliser le texte et appliquer des expressions régulières afin d’identifier les débuts et fins des tableaux.
   
4) Transformer : Convertir les données extraites en DataFrames avec pandas, puis les exporter en format Excel pour une analyse ultérieure.
   
**Alternatives utilisées**

Tabula-py : Bien que cette bibliothèque permette l’extraction de tableaux, elle s’est révélée moins performante sur des PDF aux mises en page non uniformes.

Camelot : Une autre option pour l’extraction des tableaux, mais avec des difficultés similaires sur des rapports complexes.

pdfplumber : Le même probléme que les autres bibliothéques.

Développement sur mesure : Concevoir une solution entièrement personnalisée. Cette approche était trop coûteuse en termes de temps et de ressources .

**Conséquences**

 **Avantages**
 
+ Précision et fiabilité : PyPDF2 fournit une extraction de texte précise et adaptée aux documents PDF variés.
  
+ Flexibilité : La combinaison avec pandas offre des possibilités avancées de manipulation et de transformation des données.
  
+ Modularité : La séparation des étapes (scrapping, extraction, transformation) facilite la maintenance du projet.

  **Inconvénients**

- Cas particuliers : Certains rapports aux mises en page atypiques posent encore des difficultés pour une extraction complète et homogène.
  
- Maintenance continue : Les expressions régulières doivent être régulièrement ajustées pour rester efficaces face aux variations de format des rapports.
  
- Problème avec certaines lignes : Certaines lignes, qui étaient calculées et écrites manuellement dans les rapports, ne peuvent donc pas être extraites automatiquement du tableau.

**Gestion des Erreurs**

Gestion des Exceptions : Mise en place de mécanismes pour gérer les PDF mal formatés ou les échecs d’extraction (par exemple, en réessayant l’extraction ou en passant le document concerné pour un traitement manuel).

  **Suivi et prochaines étapes**
  
Choisir un modèle de LLM : Évaluer et intégrer un modèle de traitement du langage naturel (Large Language Model) afin d'améliorer l'extraction automatique.

Améliorations itératives : Affiner les algorithmes d’extraction pour mieux gérer les variations de mise en forme et réduire les erreurs d’association des colonnes.

Voir le Schéma (Schéma.png) 



