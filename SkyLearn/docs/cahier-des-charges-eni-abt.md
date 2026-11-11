# Synthèse du cahier des charges — SI intégré ENI-ABT

## 1. Décision de cadrage

Le livrable de référence est le [cahier des charges fonctionnel et technique,
version 1.3 (PDF)](./Cahier_des_charges_SI_ENI_ABT_2026-2027.pdf), accompagné de sa
[version Word modifiable](./Cahier_des_charges_SI_ENI_ABT_2026-2027.docx).
Cette note en résume les décisions; elle ne remplace pas ces documents.

Le projet vise un **système d'information destiné à un usage officiel de
l'ENI-ABT**. La première mise en service doit couvrir les démarches à distance
prioritaires, l'inscription/réinscription en ligne, les paiements traçables,
l'accès individuel aux résultats publiés, un premier fonds d'archives
vérifiées et un assistant public d'information. La migration des archives
historiques doit être progressive et contrôlée.

Le projet est piloté par un concepteur-développeur principal avec une équipe
associant les référents métier, les représentants des utilisateurs et les
contributeurs techniques.

## 2. Vision

Le futur SI propre à l'ENI-ABT doit réunir dans un même service la scolarité,
la gestion des cours et des évaluations, les démarches à distance et les
archives officielles. Son modèle de données et ses règles seront définis à
partir des processus approuvés par les services de l'ENI-ABT. Les solutions
étudiées servent à comparer des fonctions, pas à dicter le modèle de données,
les rôles ou les procédures de l'établissement.

Le SI doit réduire les doubles saisies, rendre les procédures traçables et
offrir un accès adapté à chaque profil :

- direction et administration ;
- agents de scolarité ;
- Directeur des études ;
- chefs de DER / département ;
- enseignants ;
- étudiants ;
- anciens étudiants et diplômés ;
- candidats et public ;
- autres acteurs autorisés de l'établissement.

La solution doit être bilingue français/anglais, accessible depuis un
navigateur et conçue pour un déploiement hybride. Aucune expérimentation sur
des données personnelles réelles ne doit être réalisée sans autorisation
formelle de l'ENI-ABT.

## 3. Périmètre du service pilote

### 3.1 Fonctionnalités incluses

1. **Référentiels**
   - années académiques ;
   - départements et filières ;
   - niveaux, programmes et groupes ;
   - statuts et motifs administratifs.
2. **Identités et accès**
   - comptes utilisateurs ;
   - rôles et permissions ;
   - activation, désactivation et récupération de compte ;
   - journal des opérations sensibles.
   - délégations éventuelles à un représentant, désactivées par défaut et limitées à des autorisations expresses.
3. **Dossiers étudiants**
   - identité et coordonnées ;
   - pièces justificatives ;
   - historique académique et archives de fin de parcours ;
   - statut du dossier ;
   - rattachement à une filière, un niveau et un département.
4. **Admissions et inscriptions**
   - candidature, dépôt de pièces et suivi en ligne ;
   - inscription et réinscription à distance ;
   - contrôle des pièces ;
   - génération d'un identifiant étudiant ;
   - paiement à distance par un prestataire retenu par l'ENI-ABT ;
   - reçus, rapprochements et attestations.
5. **Consultation et pilotage**
   - consultation des résultats propres à l'étudiant après publication officielle ;
   - consultation publique des seuls résultats dont la publication est autorisée ;
   - vérification limitée de l'authenticité d'une attestation ou d'un diplôme ;
   - recherche multicritère pour les agents habilités ;
   - tableaux de bord par rôle ;
   - statistiques de suivi ;
   - exports contrôlés ;
   - notifications liées aux actions et échéances.
6. **Anciens étudiants et archives**
   - conservation des inscriptions, résultats validés, décisions de jury, diplômes et attestations ;
   - recherche ou demande par un ancien étudiant après vérification d'identité ;
   - provenance, statut de vérification, intégrité, historique de rectification et calendrier de conservation ;
   - page de vérification de document ne révélant que les informations nécessaires.
7. **Information de l'établissement**
   - chatbot LLM public fondé sur les sources officielles approuvées et datées ;
   - réponses avec références, abstention si la source manque et transfert vers un service ;
   - aucun accès du chatbot aux notes, paiements ou dossiers personnels.
8. **Vie académique**
   - programmes, inscriptions pédagogiques, cours/UE et affectation des enseignants ;
   - ressources de cours, activités et quiz formatifs ;
   - emploi du temps, salles, présences, examens et rattrapages ;
   - demandes d'absence et réclamations en ligne ;
   - formulaires et rôles métier à confirmer avec les responsables de DER, le SEGAL et les enseignants.

Le Directeur des études coordonne le périmètre académique de l'établissement.
Le détail de ses droits d'approbation par rapport aux chefs de DER, aux jurys
et à la scolarité doit être validé par les textes et délégations de l'ENI-ABT.

### 3.2 Hors périmètre de la première version

Les ressources humaines, la bibliothèque, la gestion financière élargie et
l'analytique prédictive sont des lots ultérieurs. Le paiement des inscriptions,
l'accès individuel aux résultats publiés, la préservation des dossiers des
anciens et le chatbot d'information font partie des objectifs du service
public initial. Les listes nominatives de résultats ne sont accessibles au
public que si l'ENI-ABT en autorise la publication conformément aux règles
applicables.

## 4. Rôles et droits de haut niveau

| Rôle | Responsabilités principales |
| --- | --- |
| Direction | Consulter les indicateurs, valider les décisions et superviser l'activité |
| Administration | Gérer les référentiels, les dossiers, les inscriptions et les documents |
| Directeur des études | Coordonner calendrier, enseignements, examens et suivi académique à l'échelle de l'établissement, selon délégation |
| Chef de DER / département | Gérer le suivi académique de son périmètre, selon délégation |
| Enseignant | Consulter les informations pédagogiques autorisées |
| Étudiant | Consulter son dossier, ses démarches et ses documents |
| Ancien étudiant | Demander et consulter ses archives, relevés et attestations après vérification d'identité |
| Candidat / public | Consulter les informations officielles, procédures et résultats autorisés |
| Administrateur technique | Administrer la plateforme sans accéder aux données métier au-delà du nécessaire |

Toute permission doit être accordée explicitement. Les accès doivent être
limités au périmètre organisationnel utile au rôle.

## 5. Exigences fonctionnelles prioritaires

### EF-01 — Authentification

Le système doit permettre une authentification sécurisée, une déconnexion,
une récupération de compte et une gestion des sessions.

### EF-02 — Gestion des dossiers

Un agent autorisé doit pouvoir créer, modifier, rechercher, consulter et
archiver un dossier étudiant selon son périmètre.

### EF-03 — Workflow de validation

Une candidature ou une inscription doit suivre des statuts explicites :

`Brouillon` → `Soumis` → `À vérifier` → `Validé` ou `Rejeté`.

Chaque transition doit conserver l'auteur, la date et, si nécessaire, un motif.

### EF-04 — Pièces justificatives

Le système doit permettre de déposer, vérifier, remplacer et consulter les
pièces autorisées. Les types et tailles de fichiers doivent être contrôlés.

### EF-05 — Recherche et traçabilité

Les recherches doivent respecter les permissions. Les consultations,
modifications, validations, exports et suppressions logiques doivent être
journalisés.

### EF-06 — Documents

Le système doit produire des documents administratifs à partir de modèles
validés par l'ENI-ABT et conserver leur historique.

### EF-07 — Tableaux de bord

Les indicateurs doivent être calculés à partir des données validées et
présenter leur période, leur périmètre et leur date de mise à jour.

### EF-08 — Bilinguisme

Les libellés de l'interface et les documents prévus doivent pouvoir être
présentés en français et en anglais, sans dupliquer les données métier.

### EF-09 — Inscription et paiement à distance

Un candidat ou étudiant doit pouvoir effectuer en ligne les démarches,
téléverser les pièces, suivre le traitement et régler les frais applicables
par un moyen à distance retenu par l'ENI-ABT. Le parcours nominal ne doit pas
imposer un déplacement au secrétariat. Une inscription payante n'est confirmée
qu'après vérification serveur du paiement; le reçu est disponible dans le
compte. Le SI ne conserve pas les secrets carte ou code PIN du payeur.

### EF-10 — Archives et preuves pour les anciens étudiants

Le système doit préserver les dossiers, résultats validés, décisions de jury,
diplômes et attestations selon un calendrier de conservation approuvé. Un
ancien étudiant doit pouvoir demander et récupérer ses documents après
vérification de son identité. Toute archive importée conserve sa provenance
et son état de vérification; une donnée incertaine ne doit pas être présentée
comme une preuve officielle.

### EF-11 — Résultats étudiants et publication publique

Un étudiant consulte en ligne ses résultats et relevés uniquement après leur
publication officielle, via un compte sécurisé. Des résultats peuvent aussi
être consultés par le public lorsqu'une décision de l'ENI-ABT et les règles
applicables l'autorisent; les données affichées et la durée de publication
sont alors minimisées et validées. Les relevés personnels ne sont jamais
accessibles par recherche publique. La vérification d'un diplôme par un tiers
utilise un code non devinable et ne divulgue que le minimum nécessaire.

### EF-12 — Assistant LLM d'information

Un chatbot public répond sur les formations et les procédures à partir d'un
corpus officiel approuvé et actualisé. Il cite ses sources, s'abstient s'il
ne dispose pas d'une source suffisante et propose le contact du service
compétent. Il n'accède à aucun dossier privé, n'effectue pas de décision
administrative et ne réalise pas de paiement.

### EF-13 — Direction des études

Le Directeur des études doit disposer d'un espace de coordination du
calendrier, des enseignements, des examens et du suivi des DER. Les
approbations qu'il exerce sont configurées uniquement après validation des
délégations de l'ENI-ABT. Le rôle ne peut pas modifier directement une note
publiée ni contourner les workflows de validation et de jury.

## 6. Exigences non fonctionnelles

- **Sécurité :** contrôle d'accès par rôle, mots de passe hachés, protection
  CSRF, validation des entrées, chiffrement en transit et sauvegardes chiffrées.
- **Confidentialité :** minimisation des données, conservation maîtrisée,
  accès justifié et procédure de révocation.
- **Archives :** provenance, intégrité, formats exportables, métadonnées,
  conservation définie et restauration vérifiée.
- **Paiement :** confirmation serveur, opérations idempotentes, reçus,
  rapprochement, gestion des remboursements; aucune donnée de carte ou PIN
  stockée dans le SI.
- **IA :** sources versionnées, protection contre la fuite de données,
  abstention en cas d'incertitude, contrôle humain et désactivation possible.
- **Disponibilité :** sauvegarde automatisée, restauration testée et procédure
  de reprise documentée.
- **Qualité :** tests unitaires, tests d'intégration, revue de code et
  validation métier avant chaque livraison.
- **Interopérabilité :** API documentée, imports contrôlés et exports
  traçables.
- **Accessibilité :** interface responsive, navigation clavier et contraste
  conforme aux bonnes pratiques.
- **Observabilité :** journaux techniques séparés des journaux métier,
  alertes d'erreur et suivi des opérations critiques.

## 7. Architecture cible indicative

- interface web responsive ;
- API métier Django REST Framework ou équivalent ;
- PostgreSQL pour les données transactionnelles ;
- stockage séparé pour les documents ;
- service de notifications ;
- conteneurisation Docker ;
- CI/CD avec contrôles de tests et de sécurité ;
- infrastructure hybride : services institutionnels maîtrisés par l'ENI-ABT,
  sauvegardes et services complémentaires externalisés selon autorisation.

Les recherches documentaires sur Frappe Education, Ellucian Student et Open
edX ont servi à comparer les admissions, l'inscription aux programmes, les
portails étudiants, la pédagogie, les notes et les contrôles d'accès. Les
fonctions retenues sont réécrites comme exigences propres à l'ENI-ABT :
inscription pédagogique liée au programme et au semestre, cours obligatoires
ou optionnels, ressources et quiz, présence, évaluations paramétrables,
publication officielle, paiements traçables et conservation distincte des
archives. Un cours archivé ou un compte désactivé ne vaut pas archive
institutionnelle d'un dossier étudiant.

La base de données opérationnelle doit utiliser un modèle canonique approuvé
par l'ENI-ABT. Elle ne contient pas les noms, schémas ou données de
démonstration des solutions comparées. Le benchmark est un matériau de
conception, non une source de données.

### Repères documentaires du benchmark

- [Frappe Education](https://docs.frappe.io/education/student-portal) documente un portail étudiant avec emploi du temps, résultats, frais/paiements, présences et profil; sa documentation décrit aussi [la candidature](https://docs.frappe.io/education/student-applicant) et [l'inscription à un programme](https://docs.frappe.io/education/program-enrollment).
- [Ellucian Student](https://www.ellucian.com/products/student) décrit un SIS commercial couvrant le cycle étudiant, les admissions, les programmes, les portails et des fonctions IA annoncées par l'éditeur. Ces affirmations ne constituent pas une évaluation indépendante.
- [Open edX](https://docs.openedx.org/en/latest/educators/quickstarts/build_a_course.html) documente les cours, les rôles pédagogiques, les carnets de notes et un assistant IA ancré dans une leçon. C'est une plateforme d'apprentissage, pas un registre officiel de scolarité ni d'archives étudiantes.

Ces sources ne garantissent pas une correspondance intégrale aux besoins réels
de l'école. La cible de 100 % porte sur la traçabilité et la recette de toutes
les exigences prioritaires approuvées, après ateliers métier et validation
institutionnelle.

## 8. Critères d'acceptation du pilote

Le pilote sera considéré comme présentable lorsque :

1. les rôles et permissions sont validés par les référents métier ;
2. un candidat de test peut déposer un dossier, suivre son état et effectuer
   une inscription à distance avec paiement simulé, reçu et rapprochement ;
3. chaque changement sensible est retrouvé dans le journal d'audit ;
4. un utilisateur ne voit aucune donnée hors de son périmètre ;
5. un étudiant consulte seulement ses résultats publiés et un tiers ne peut
   pas lire son relevé privé ;
6. un ancien étudiant de test retrouve une archive vérifiée et un document
   partagé est vérifiable sans exposer de dossier complet ;
7. le Directeur des études voit les informations de coordination à l'échelle
   de l'établissement sans obtenir de droits de modification non délégués ;
8. le chatbot répond avec une source approuvée, s'abstient pour une procédure
   inconnue et n'accède pas aux données personnelles ;
9. les documents et exports produits sont vérifiés par la scolarité ;
10. les sauvegardes et une restauration de test sont démontrées ;
11. les tests automatisés et la recette métier sont passés ;
12. une documentation utilisateur et une procédure d'assistance sont livrées.
13. 100 % des exigences prioritaires approuvées sont reliées à un scénario de test et à une preuve de recette acceptée.

## 9. Gouvernance et prochaines étapes

Avant tout développement impliquant des données réelles, l'équipe doit :

1. obtenir le parrainage et l'autorisation officielle de l'ENI-ABT ;
2. nommer un référent métier de la scolarité et un référent sécurité ;
3. cartographier les outils, fichiers et responsabilités réellement utilisés par l'ENI-ABT ;
4. inventorier les archives papier et numériques, fixer le calendrier de
   conservation et les critères de vérification ;
5. valider les workflows, les statuts, les documents, les frais, exemptions,
   prestataire de paiement, remboursements et règles de publication des résultats ;
6. faire approuver le corpus, les sources et les règles d'escalade du chatbot ;
7. réaliser un prototype sur données fictives ;
8. organiser une recette avec des représentants de tous les profils ;
9. décider du passage en production sur la base de critères mesurables.

Le chatbot public d'information peut faire partie de la première mise en
service après approbation de son corpus et validation de la confidentialité,
des réponses et de l'escalade humaine. Les fonctions prédictives ou d'analyse
académique restent ultérieures et ne peuvent pas prendre de décision
définitive sur un étudiant.
