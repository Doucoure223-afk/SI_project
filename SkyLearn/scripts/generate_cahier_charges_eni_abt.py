from __future__ import annotations

from html import escape
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.platypus.tableofcontents import TableOfContents


ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / "docs"
PDF_PATH = DOCS / "Cahier_des_charges_SI_ENI_ABT_2026-2027.pdf"
DOCX_PATH = DOCS / "Cahier_des_charges_SI_ENI_ABT_2026-2027.docx"
AS_OF = "11 novembre 2026"
DOCUMENT_VERSION = "1.3"

BLACK = "#000000"
INK = "#202020"
MUTED = "#555555"
PALE = "#EEEEEE"
GRID = "#888888"
WHITE = "#FFFFFF"
DOCUMENT_REFERENCE = "ENI-ABT-SI-CDC-001"


def plain_text(text: str) -> str:
    return text.replace("—", ":").replace("–", "-").replace("•", "-")


def P(text: str) -> dict:
    return {"type": "p", "text": text}


def B(*items: str) -> dict:
    return {"type": "bullets", "items": list(items)}


def T(headers: list[str], rows: list[list[str]], widths=None) -> dict:
    return {"type": "table", "headers": headers, "rows": rows, "widths": widths}


def C(title: str, text: str) -> dict:
    return {"type": "callout", "title": title, "text": text}


SECTIONS = [
    {
        "title": "Synthèse exécutive",
        "blocks": [
            P(
                "Les éléments vérifiés concernant l’ENI-ABT comprennent son site institutionnel "
                "public et le portail de candidature eCandidat. Les outils internes, les bases "
                "de données métier et les procédures détaillées devront être recensés avec les "
                "services de l’établissement avant toute réalisation."
            ),
            P(
                "Le présent cahier des charges propose la conception d’un système d’information "
                "intégré, centré d’abord sur le socle académique : référentiels, identité, dossiers "
                "étudiants, admissions en interopérabilité avec eCandidat, inscriptions et "
                "paiements en ligne, cours, notes, résultats et portails par profil. Il prévoit "
                "aussi la conservation des dossiers et preuves de diplômes des anciens étudiants, "
                "ainsi qu’un assistant public fondé sur les informations officielles de l’école. "
                "Les services RH, bibliothèque et l’analytique avancée seront introduits par lots."
            ),
            C(
                "Décision structurante",
                "Concevoir un système propre à l’ENI-ABT, en réunissant dans un même service "
                "les fonctions de scolarité, de pédagogie et de gestion des services à distance. "
                "Chaque workflow, rôle, règle de notation et donnée doit être validé par les "
                "responsables de l’école. Définir les systèmes maîtres, vérifier les archives, "
                "valider le paiement et tester sur des données fictives ou autorisées.",
            ),
            T(
                ["Rubrique", "Proposition de cadrage"],
                [
                    ["Finalité", "SI officiel, intégré, évolutif et exploitable par l’ENI-ABT."],
                    ["Stratégie", "Conception par domaines; archives protégées; démarches et paiements à distance; publication contrôlée."],
                    ["Premier pilote", "Scolarité : dossier étudiant, première inscription/réinscription en ligne, paiement traçable et consultation des résultats publiés."],
                    ["Anciens étudiants", "Rechercher un dossier historique vérifié et demander/obtenir des attestations ou relevés selon les habilitations."],
                    ["Information publique", "Assistant conversationnel sur corpus approuvé; aucun accès aux dossiers étudiants privés."],
                    ["Socle pédagogique", "Cours, ressources, évaluations, quiz, présences et notes intégrés au suivi officiel des inscriptions et résultats."],
                    ["Déploiement cible", "Hybride, sous réserve d’une décision formelle de l’établissement."],
                    ["Langues", "Français en priorité; anglais selon les contenus et ressources validés."],
                    ["Statut du document", "Version de travail à valider par la direction et les services ENI-ABT."],
                ],
                [0.23, 0.77],
            ),
        ],
    },
    {
        "title": "A. Présentation du projet",
        "blocks": [
            P(
                "Nom de travail : Système d’information intégré de l’ENI-ABT. Le périmètre "
                "et le nom définitifs relèvent de l’établissement."
            ),
            {
                "subtitle": "A.1 Contexte et justification",
                "blocks": [
                    P(
                        "L’ENI-ABT est un établissement public d’enseignement supérieur au Mali. "
                        "Son site institutionnel présente huit domaines/DER, des programmes de "
                        "Licence et de Master, des appels à candidatures, des actualités et des "
                        "événements. La candidature en ligne est portée par eCandidat, un portail "
                        "distinct qui décrit la création de compte, la constitution du dossier, "
                        "le dépôt de pièces, la soumission et le suivi de décision."
                    ),
                    P(
                        "Un SI institutionnel doit relier ces processus au parcours académique "
                        "après admission, sans recopier inutilement les données et sans imposer "
                        "une rupture des outils qui fonctionnent déjà. Le besoin principal est "
                        "de clarifier les responsabilités des données, fiabiliser les référentiels, "
                        "sécuriser les dossiers et rendre les opérations traçables."
                    ),
                    P(
                        "Un besoin prioritaire complémentaire est la continuité de la preuve "
                        "académique : un ancien étudiant doit pouvoir faire vérifier son parcours "
                        "et demander une attestation ou un relevé même plusieurs années après "
                        "son départ. L’admission, l’inscription, le dépôt des pièces, le paiement "
                        "et le suivi doivent être réalisables à distance, sans déplacement "
                        "obligatoire au secrétariat lorsque les règles et moyens de paiement le permettent."
                    ),
                    C(
                        "Périmètre de l’étude",
                        "Les pages publiques sont une source d’orientation, non une preuve "
                        "complète des procédures internes. Les processus, règles de notes, frais, "
                        "calendriers, rôles et données réels restent à confirmer par les services "
                        "compétents de l’ENI-ABT.",
                    ),
                ],
            },
            {
                "subtitle": "A.2 Analyse de l’existant",
                "blocks": [
                    T(
                        ["Source étudiée", "Capacités documentées ou besoins exprimés", "Écart / conséquence"],
                        [
                            [
                                "Site public eni-abt.ml",
                                "Présentation, catalogue de formations, informations de contact, événements et actualités.",
                                "Catalogue à normaliser; pages publiques ne décrivent pas les workflows internes.",
                            ],
                            [
                                "eCandidat ENI-ABT",
                                "Compte candidat, formulaire sauvegardé, vœux, pièces, soumission, numéro de dossier et suivi.",
                                "Système distinct; contrat d’API, export, propriétaire des données et responsabilités à confirmer.",
                            ],
                            [
                                "Capture de besoins transmise par le demandeur",
                                "Accès étudiant aux notes, examens, relevés et diplômes; inscriptions; absences et demandes; salles, rattrapages et emploi du temps. Répartition indiquée : responsables de DER pour professeurs, emplois du temps et notes; SEGAL pour formulaires; enseignants pour saisie des notes.",
                                "Expression de besoins à confirmer en atelier; ne prouve pas l’existence ni la couverture de ces fonctions dans les systèmes actuels. Le développement doit distinguer information publique, dossier privé et archive historique.",
                            ],
                        ],
                        [0.18, 0.39, 0.43],
                    ),
                    P(
                        "La consultation des outils publics n’a pas permis de confirmer une "
                        "archive institutionnelle centralisée, un service de vérification des "
                        "diplômes ou un parcours de paiement en ligne de bout en bout. Le cahier "
                        "traite donc ces points comme des besoins prioritaires à instruire, et "
                        "non comme des fonctions déjà disponibles."
                    ),
                    P(
                        "Le SI ne peut pas recréer une preuve historique absente des registres, "
                        "fichiers ou archives conservés. Le projet devra inventorier les sources "
                        "retrouvées, faire contrôler les dossiers critiques par la scolarité, "
                        "signaler les éléments manquants ou incertains et ne délivrer un document "
                        "certifié qu’après validation selon une procédure institutionnelle."
                    ),
                ],
            },
            {
                "subtitle": "A.3 Objectifs et résultats attendus",
                "blocks": [
                    B(
                        "Créer un référentiel institutionnel cohérent des personnes, structures, programmes, cours, années et semestres.",
                        "Relier programmes, inscriptions pédagogiques, cours/UE, enseignants affectés, ressources, évaluations, quiz, présences et résultats dans un parcours académique cohérent.",
                        "Réduire les doubles saisies par l’intégration contrôlée avec eCandidat et les systèmes autorisés.",
                        "Donner à chaque acteur un espace de travail adapté et limité à ses responsabilités.",
                        "Sécuriser le cycle des notes et résultats, depuis la saisie jusqu’à la publication et à l’archivage.",
                        "Préserver le dossier académique sur la durée définie par les règles applicables et donner aux anciens étudiants un moyen sûr de prouver leur parcours.",
                        "Permettre les admissions, inscriptions, demandes de pièces et paiements à distance, avec confirmation et reçu consultable en ligne.",
                        "Rendre les résultats consultables en ligne par les étudiants après publication officielle et encadrer séparément les résultats destinés au public.",
                        "Répondre aux questions générales sur l’ENI-ABT avec un assistant conversationnel fondé sur les contenus institutionnels approuvés.",
                        "Produire des indicateurs de pilotage définis, explicables et réconciliables avec les données sources.",
                        "Garantir la continuité opérationnelle, la traçabilité, la qualité des données et la capacité de reprise.",
                        "Permettre une extension par lots vers finances, RH, bibliothèque, communication et services numériques.",
                    ),
                    T(
                        ["Indicateur proposé", "Cible initiale de cadrage", "Méthode de mesure"],
                        [
                            ["Traçabilité des actions sensibles", "100 % des validations, publications, corrections et exports journalisés.", "Tests d’acceptation et échantillonnage des journaux."],
                            ["Contrôle d’accès", "Aucun accès inter-département non autorisé dans les scénarios de recette.", "Tests négatifs par rôle et périmètre."],
                            ["Performance", "95e percentile < 2 s pour les consultations simples sous charge convenue.", "Test de charge avec nombre d’utilisateurs convenu par l’établissement."],
                            ["Reprise des données", "100 % des enregistrements importés réconciliés ou consignés en anomalie.", "Rapport source/cible, doublons et rejets."],
                            ["Disponibilité pilote", "Cible initiale 99,5 % mensuelle hors maintenance annoncée.", "Supervision; cible à valider selon connectivité et budget."],
                            ["Restauration", "RPO cible 24 h, RTO cible 8 h pour le pilote.", "Exercice de restauration chronométré; valeurs à valider par la direction."],
                            ["Accès au parcours", "Les étudiants retrouvent leurs résultats publiés et les anciens peuvent initier une demande de preuve sans déplacement.", "Tests d’accès étudiant, de récupération de compte et de demande d’archive."],
                            ["Transactions", "Chaque paiement confirmé possède une référence vérifiable et un reçu rapproché.", "Tests de paiement accepté, refusé, doublé, retardé, remboursé et rapproché."],
                        ],
                        [0.24, 0.42, 0.34],
                    ),
                    P(
                        "Ces cibles sont des propositions de cadrage, pas des engagements de "
                        "service. Elles devront être adaptées après mesure du réseau, du nombre "
                        "d’utilisateurs, du budget et des contraintes de l’établissement."
                    ),
                ],
            },
            {
                "subtitle": "A.4 Cibles et acteurs",
                "blocks": [
                    T(
                        ["Acteur", "Besoins principaux", "Accès de principe"],
                        [
                            ["Direction générale", "Indicateurs, arbitrages, rapports consolidés.", "Consultation agrégée; approbations expressément déléguées."],
                            ["Secrétariat général", "Coordination, référentiels, actes, calendrier institutionnel.", "Gestion transversale selon délégation."],
                            ["Scolarité / admissions", "Dossiers, inscriptions, pièces, attestations et corrections.", "Traitement des dossiers affectés; actions sensibles auditées."],
                            ["Directeur des études", "Coordination académique transversale : calendrier, enseignements, examens, suivi des résultats et coordination des DER.", "Périmètre institutionnel; validations et approbations uniquement si une délégation officielle le prévoit; actions auditées."],
                            ["Chefs de DER / département", "Organisation du département, suivi des enseignants, emplois du temps, cours et notes selon délégation.", "Périmètre de leur DER; pas d’accès aux autres DER hors fonction."],
                            ["Enseignants", "Cours, listes de classe, évaluations, saisie et soumission des notes.", "Cours et étudiants affectés."],
                            ["Étudiants", "Profil, inscription, emploi du temps, résultats publiés, documents et demandes.", "Données personnelles et parcours propres."],
                            ["Représentants légaux ou tiers délégués (option)", "Démarches expressément confiées par un étudiant lorsque le cadre applicable le permet.", "Fonction désactivée par défaut; délégation limitée, explicite, révocable et sans accès automatique aux notes."],
                            ["Anciens étudiants / diplômés", "Retrouver les preuves de scolarité, consulter ses documents et demander une vérification ou un duplicata.", "Accès à son propre dossier après vérification d’identité; partage externe limité et contrôlé."],
                            ["Candidats et public", "Comprendre l’offre de formation, les procédures, les calendriers et les résultats officiellement publics.", "Contenus publics uniquement; aucune recherche libre dans les dossiers privés."],
                            ["Finances", "Frais, échéances, rapprochements et reçus.", "Données financières selon rôle; paiement intégré après validation."],
                            ["RH et bibliothèque", "Personnel, services, prêts et ressources.", "Domaines fonctionnels séparés et autorisés."],
                            ["Administrateur technique / sécurité", "Comptes, supervision, sauvegarde, incidents.", "Privilèges administratifs contrôlés; accès métier réduit et audité."],
                        ],
                        [0.23, 0.42, 0.35],
                    ),
                ],
            },
            {
                "subtitle": "A.5 Description globale",
                "blocks": [
                    P(
                        "La cible est une plateforme web responsive, bilingue et accessible, avec "
                        "un noyau métier modulaire, une API documentée, une base transactionnelle "
                        "relationnelle, un stockage privé des pièces, une archive électronique "
                        "institutionnelle et des connecteurs d’échange. "
                        "Le noyau est conçu comme un monolithe modulaire au démarrage : il limite "
                        "la complexité d’exploitation tout en isolant les domaines pour permettre "
                        "une évolution future."
                    ),
                    T(
                        ["Couche", "Responsabilité"],
                        [
                            ["Portails web / mobile", "Espaces candidat, étudiant, ancien étudiant, enseignant, chef de DER, Directeur des études, administration, direction et public; interface responsive; français/anglais."],
                            ["API et identité", "API HTTPS documentée OpenAPI; authentification, autorisation, validation, limites et journalisation."],
                            ["Domaines métier", "Admissions, scolarité, archives, examens/résultats, paiements, emploi du temps et documents; RH et bibliothèque en phases futures."],
                            ["Données", "PostgreSQL; migrations versionnées; stockage documentaire privé et dépôt d’archives contrôlé; sauvegardes chiffrées."],
                            ["Services externes", "Prestataire de paiement intégré sans conservation de données de carte; modèle LLM avec récupération sur sources institutionnelles approuvées."],
                            ["Traitements différés", "Notifications, génération de documents, rapprochements de paiement et imports via tâches asynchrones si volumes le justifient."],
                            ["Exploitation", "Conteneurs, intégration continue, supervision, journaux structurés, alertes et environnement de reprise."],
                        ],
                        [0.24, 0.76],
                    ),
                    C(
                        "Principe de modernité",
                        "La modernité est définie par la sécurité maintenue, l’accessibilité, "
                        "l’interopérabilité, l’observabilité, la résilience et l’expérience "
                        "utilisateur mesurable, et non par l’adoption systématique de microservices "
                        "ou d’IA.",
                    ),
                ],
            },
            {
                "subtitle": "A.6 Description détaillée — périmètre et priorités",
                "blocks": [
                    T(
                        ["Lot", "Inclus", "Exclusions / prérequis"],
                        [
                            ["Lot 0 - Cadrage et autorisations", "Cartographie des processus/données; règles de conservation; benchmark des SIS/LMS; parcours candidats, étudiants et anciens; exigences de paiement et chatbot.", "Valider les règles métier, les fondements réglementaires, les systèmes maîtres, le prestataire, la répartition des responsabilités et les autorisations avant données réelles."],
                            ["Lot 1 - Socle et services à distance", "Comptes, rôles, référentiels, admission et inscription en ligne, pièces, paiement avec reçu, portail étudiant, archive pilote, consultation personnelle des résultats et chatbot d’information publique.", "Intégration eCandidat selon interface autorisée; paiement activé après choix et validation du prestataire."],
                            ["Lot 2 - Résultats et diplômes", "Notes, examens/rattrapages, jury, publication, relevés, diplômes et archives historiques; vérification externe limitée des titres.", "Règlements approuvés; données historiques réconciliées; publication publique conforme aux règles institutionnelles."],
                            ["Lot 3 - Vie académique", "Emplois du temps, salles, présence, demandes d’absence, réclamations, formulaires et notifications.", "Confirmer le périmètre de chaque acteur; données de santé particulièrement restreintes."],
                            ["Lot 4 - Services institutionnels", "RH, bibliothèque, finances élargies, rapprochements et pilotage.", "Cadrage par les services responsables et validation sécurité/juridique."],
                            ["Lot 5 - Analyses avancées", "Indicateurs avancés et éventuelles alertes d’aide.", "Sources documentées, contrôle, confidentialité, évaluation et supervision humaine; aucun automatisme décisionnel."],
                        ],
                        [0.2, 0.47, 0.33],
                    ),
                    B(
                        "Hors périmètre du premier pilote : remplacement intégral du portail eCandidat, décision automatique d’admission ou de réussite, mise en production avec données non réconciliées, stockage des secrets de paiement et toute IA prenant une décision sur un étudiant. Le paiement en ligne et le parcours d’inscription à distance sont des objectifs du premier service public, soumis au choix d’un prestataire et à la validation des règles financières.",
                        "Aucun produit externe ne définit les règles de l’ENI-ABT. Les solutions étudiées servent uniquement à comparer des fonctions; le modèle de données et les workflows finaux seront construits à partir des besoins approuvés par l’établissement.",
                    ),
                ],
            },
            {
                "subtitle": "A.7 Règles métier majeures à formaliser",
                "blocks": [
                    B(
                        "Chaque étudiant, candidat, enseignant et agent dispose d’un identifiant stable; les doublons potentiels sont rapprochés avec contrôle humain.",
                        "Les programmes et cours sont versionnés par année académique; une modification ne réécrit pas les résultats historiques.",
                        "Une note est rattachée à une évaluation, une inscription, une période académique et un enseignant habilité.",
                        "Les barèmes, coefficients, seuils et règles de calcul sont configurés et versionnés par règlement/période, jamais enfouis dans le code métier.",
                        "La note soumise n’est pas encore une note officielle. La publication intervient après les validations et délibérations habilitées.",
                        "Une correction après publication crée une nouvelle version avec auteur, motif, date, validation et référence à la version précédente.",
                        "Les décisions de jury et résultats publiés restent consultables selon les politiques de conservation autorisées.",
                        "Les étudiants consultent leurs résultats personnels après publication officielle via un compte sécurisé. Une liste nominative publique n’est diffusée que sur décision de l’ENI-ABT conforme aux règles applicables; la vérification d’un diplôme par un tiers ne révèle que le minimum nécessaire.",
                        "Le dossier d’un ancien étudiant et les pièces délivrées sont archivés avec provenance, période, référence, état de vérification et historique de correction; une archive importée non vérifiée est clairement signalée.",
                        "Une inscription payante à distance n’est définitive qu’après confirmation serveur du prestataire agréé; toute exemption, remise, annulation ou restitution suit une règle approuvée et un journal financier.",
                        "Les règles d’admission, pièces obligatoires, frais, échéances, rattrapages et recours sont confirmées par chaque service compétent.",
                        "Le Directeur des études coordonne le périmètre académique; les décisions qu’il valide et celles relevant des chefs de DER ou du jury sont définies par les textes et délégations de l’ENI-ABT, jamais supposées par le logiciel.",
                    ),
                ],
            },
            {
                "subtitle": "A.8 Ressources et responsabilités",
                "blocks": [
                    T(
                        ["Rôle projet", "Responsabilité"],
                        [
                            ["Sponsor institutionnel", "Mandat, arbitrage, budget, accès aux services et décision de mise en production."],
                            ["Product owner métier", "Priorisation du backlog et validation des critères d’acceptation."],
                            ["Référents scolarité, DER, enseignants, étudiants", "Règles métier, scénarios, recette et retours d’expérience."],
                            ["Directeur des études", "Référent métier de la coordination académique; valide les workflows et responsabilités qui relèvent de sa délégation."],
                            ["Architecte / développeur principal", "Architecture du SI cible; benchmark des solutions; décision documentée sur les composants et intégrations à retenir."],
                            ["Équipe produit", "Backend, interface, tests, données, intégration et documentation."],
                            ["Référent sécurité / données", "Classification, accès, risques, réponse aux incidents et obligations applicables."],
                            ["Équipe exploitation", "Environnements, supervision, sauvegardes, déploiement et reprise."],
                        ],
                        [0.32, 0.68],
                    ),
                ],
            },
        ],
    },
    {
        "title": "B. Benchmark et positionnement",
        "blocks": [
            P(
                "Benchmark documentaire réalisé le 11 novembre 2026 à partir de pages officielles "
                "de produits et de documentation. Il compare des fonctions publiquement décrites, "
                "pas la qualité réelle, le coût total, la sécurité, les contrats ou l’adaptation au "
                "Mali. Les descriptions fournisseurs ne constituent pas une évaluation indépendante."
            ),
            T(
                ["Système étudié", "Fonctions documentées", "Transposition retenue pour l’ENI-ABT"],
                [
                    ["Frappe Education (ERPNext)", "Candidatures avec approbation et inscription à un programme; année/terme, cours obligatoires ou optionnels; résultats d’évaluation; portail étudiant avec emploi du temps, notes, présence, historique et paiement des frais.", "Retenir un dossier longitudinal, le lien inscription-programme-cours, un portail étudiant unifié, les présences et des frais/reçus traçables. Le prestataire de paiement cité dans sa documentation n’est pas présumé convenir à l’ENI-ABT."],
                    ["Ellucian Student", "Système commercial d’enseignement supérieur présenté comme couvrant le cycle étudiant, les admissions, la gouvernance des programmes, les portails et les services de réussite étudiante. Le fournisseur décrit aussi des fonctions d’IA, notamment un conseiller virtuel contextualisé.", "Retenir le parcours complet candidat-diplômé, des catalogues de programmes gouvernés, des espaces par rôle et une aide contextualisée seulement si autorisée. Une IA liée au compte étudiant ne remplace pas le chatbot public de procédures; prix, API, résidence et intégration restent à étudier."],
                    ["Open edX", "Plateforme d’apprentissage : création de cours, inscription, carnet de notes et rapports, rôles d’équipe limités au cours; la version Verawood documente un assistant IA ancré dans le contenu d’une leçon.", "Retenir les cours, activités/quiz, ressources, évaluations formatives et délégations pédagogiques par cours. Open edX n’est pas un SIS : l’archivage d’un cours n’est pas la conservation officielle du dossier étudiant."],
                    ["eCandidat ENI-ABT", "Portail public de candidature et suivi des dossiers décrit par l’établissement.", "Préserver le système maître candidat tant que les responsables n’ont pas décidé autrement; confirmer interfaces, export, responsabilités et rapprochement avant toute intégration."],
                ],
                [0.18, 0.39, 0.43],
            ),
            P(
                "Le SI cible est un produit institutionnel cohérent, non une juxtaposition de modules "
                "copiés. Les pratiques utiles du benchmark sont transposées en exigences et en "
                "modèle de données ENI-ABT, sans importer les schémas, comptes ou données de "
                "démonstration des solutions étudiées. L’objectif mesurable est une traçabilité "
                "à 100 % des exigences prioritaires approuvées vers un scénario de recette réussi; "
                "une correspondance intégrale aux besoins réels ne peut être déclarée qu’après "
                "ateliers métier, validation officielle et recette par les utilisateurs."
            ),
        ],
    },
    {
        "title": "C. Développement du logiciel",
        "blocks": [
            {
                "subtitle": "C.1 Type et caractéristiques",
                "blocks": [
                    P(
                        "Application web transactionnelle propre à l’ENI-ABT, complétée "
                        "par des portails utilisateurs et une API HTTP versionnée. La première "
                        "phase doit privilégier un monolithe modulaire Django (domaines séparés, "
                        "contrats d’API et migrations propres) plutôt qu’une architecture "
                        "microservices prématurée. Les modèles de données, règles métier, rôles "
                        "et intégrations seront définis à partir des processus validés par "
                        "l’ENI-ABT; les produits du marché ne dictent pas les choix finaux."
                    ),
                    T(
                        ["Composant cible", "Proposition à valider"],
                        [
                            ["Backend", "Python + version Django maintenue; DRF/API ou mécanisme API équivalent à sélectionner après audit."],
                            ["Versionnement framework", "Choisir au démarrage une version maintenue de Django; vérifier le calendrier de support, les dépendances, les correctifs de sécurité et la compatibilité des composants retenus."],
                            ["Données relationnelles", "PostgreSQL, version majeure supportée, mises à jour mineures régulières et stratégie de sauvegarde/restauration."],
                            ["Interface", "Web responsive; conserver les gabarits Django si cela réduit le risque, ou ajouter un frontend React/PWA seulement si les besoins UX et compétences le justifient."],
                            ["Identité", "Django auth au pilote; SSO OIDC/MFA institutionnel à intégrer si un fournisseur d’identité existe et après décision d’architecture."],
                            ["Documents", "Stockage objet privé compatible S3 ou service institutionnel; URL signée/contrôle d’accès, antivirus et métadonnées."],
                            ["Asynchronisme", "Tâches de fond et cache (Redis/Celery ou équivalent) lorsque les volumes et l’exploitation justifient ces composants."],
                            ["Interface d’intégration", "API décrite avec OpenAPI; imports/exports idempotents, journalisés et versionnés."],
                            ["Déploiement", "Images conteneurisées; environnements séparés; pipeline CI/CD, secrets hors dépôt et infrastructure reproductible."],
                        ],
                        [0.24, 0.76],
                    ),
                    C(
                        "Décision de version",
                        "Les versions de framework changent au fil du calendrier 2026-2027. "
                        "Le marché et l’équipe doivent figer la version effectivement maintenue "
                        "au début de l’implémentation, avec compatibilité dépendances vérifiée et "
                        "fenêtre de mise à jour définie.",
                    ),
                ],
            },
            {
                "subtitle": "C.2 Structure fonctionnelle (menus cibles)",
                "blocks": [
                    T(
                        ["Espace", "Rubriques"],
                        [
                            ["Accueil", "Tableau de bord, tâches en attente, notifications, calendrier institutionnel."],
                            ["Référentiels", "Années, semestres, DER, filières, diplômes, niveaux, cours, groupes, salles."],
                            ["Admissions", "Campagnes, dossiers importés/liaison eCandidat, pièces, contrôles et décisions autorisées."],
                            ["Scolarité", "Dossiers, inscriptions/réinscriptions en ligne, pièces, attestations, réclamations et demandes d’absence."],
                            ["Anciens étudiants", "Archives, diplômes, attestations, relevés, vérification de parcours et demandes de duplicata."],
                            ["Direction des études", "Coordination académique, calendrier, suivi des DER, organisation et consultation des cours/examens selon délégations."],
                            ["Pédagogie", "Programmes, inscriptions pédagogiques, cours/UE, ressources, quiz, enseignants, affectations, emplois du temps, salles, présences et évaluations."],
                            ["Notes & examens", "Saisie, examens/rattrapages, barèmes, délibérations, résultats, recours et relevés."],
                            ["Finances et paiements", "Frais, échéances, paiement à distance, reçus, remboursements et rapprochements; accès séparé."],
                            ["Assistant d’information", "Chatbot LLM public sur le site et les procédures; sources citées et transfert vers un service."],
                            ["RH", "Personnel, affectations, contrats et absences, sous réserve du cadrage RH."],
                            ["Bibliothèque", "Catalogue, adhésions, prêts, retours et ressources numériques."],
                            ["Communication", "Annonces, événements, notifications et documents institutionnels."],
                            ["Pilotage", "Indicateurs validés, exports contrôlés, audit, administration et supervision."],
                        ],
                        [0.22, 0.78],
                    ),
                ],
            },
            {
                "subtitle": "C.3 Exigences fonctionnelles prioritaires",
                "blocks": [
                    P(
                        "Priorités : M = Must have pour le pilote ou le lot indiqué; S = Should have; "
                        "C = Could have. L’acceptation se fait sur scénarios de bout en bout, données "
                        "fictives ou autorisées et traces de recette."
                    ),
                    T(
                        ["ID", "Exigence testable", "Priorité"],
                        [
                            ["IAM-01", "L’administrateur habilité crée, active, désactive et rattache un compte à un ou plusieurs rôles autorisés.", "M"],
                            ["IAM-02", "Chaque demande est autorisée côté serveur selon rôle, action et périmètre organisationnel; masquer un bouton ne suffit pas.", "M"],
                            ["IAM-03", "Le système applique les politiques de mot de passe, limitation des tentatives, expiration de session et réinitialisation sécurisée.", "M"],
                            ["IAM-04", "Si l’ENI-ABT valide ce besoin et sa base réglementaire, un étudiant peut accorder à un tiers une délégation limitée, explicite, traçable et révocable; aucun accès aux notes ou pièces n’est automatique.", "C"],
                            ["REF-01", "Un administrateur métier gère années, semestres, DER, filières, diplômes, niveaux et cours avec historique.", "M"],
                            ["DB-01", "Le modèle relationnel suit un référentiel canonique approuvé par l’ENI-ABT; les données de démonstration et schémas propres aux produits comparés ne sont pas importés dans la base métier.", "M"],
                            ["EDU-01", "Le Directeur des études coordonne le calendrier, les enseignements, les examens et le suivi des DER à l’échelle de l’établissement, dans un périmètre distinct de celui des chefs de DER.", "M"],
                            ["EDU-02", "Toute validation ou approbation attribuée au Directeur des études est activée seulement si une délégation officielle la prévoit; le rôle ne peut pas écraser directement une note publiée.", "M"],
                            ["ADM-01", "Un agent autorisé recherche un dossier candidat/étudiant, consulte son état et voit uniquement les pièces autorisées.", "M"],
                            ["ADM-02", "Un candidat dépose en ligne sa candidature, ses pièces et ses vœux, peut sauvegarder un brouillon et suit l’état de son dossier sans passage physique obligatoire.", "M"],
                            ["ADM-03", "Le système importe ou reçoit un dossier eCandidat par interface autorisée; il détecte les doublons et journalise les rejets.", "M"],
                            ["ADM-04", "Si l’API eCandidat n’existe pas, un import contrôlé avec validation, rapport d’erreurs et reprise idempotente est fourni.", "M"],
                            ["SCO-01", "Un agent crée ou valide une inscription à une année, une filière, un niveau et un groupe après vérification des prérequis configurés.", "M"],
                            ["SCO-02", "Le dossier étudiant conserve les changements de programme, statut et coordonnées sans écraser les événements historiques.", "M"],
                            ["SCO-03", "Un étudiant consulte ses informations, son inscription et les documents publiés liés à son compte.", "M"],
                            ["SCO-04", "L’étudiant initie et termine en ligne son inscription/réinscription, soumet les pièces et suit les étapes; la présence au secrétariat n’est pas requise dans le parcours nominal.", "M"],
                            ["PAY-01", "Le système présente les frais, échéances et exonérations validés; il offre au moins un moyen de paiement à distance retenu par l’ENI-ABT (par exemple monnaie mobile ou banque) via prestataire contractuellement autorisé, sans collecter ni stocker les données carte/PIN.", "M"],
                            ["PAY-02", "Le paiement passe à confirmé uniquement après vérification serveur du prestataire; appels répétés, délais et notifications dupliquées ne créent ni double paiement ni double inscription.", "M"],
                            ["PAY-03", "Après confirmation, un reçu numéroté est disponible en ligne; transactions, remboursements et rejets sont traçables et rapprochables par les finances.", "M"],
                            ["HIST-01", "Le système conserve les résultats validés, décisions de jury, inscriptions et pièces de fin de parcours selon un calendrier de conservation approuvé.", "M"],
                            ["HIST-02", "Un ancien étudiant peut retrouver ou demander son dossier et ses attestations/relevés après vérification d’identité, même sans compte actif d’étudiant.", "M"],
                            ["HIST-03", "Un document délivré porte une référence vérifiable; un tiers vérifie son authenticité via un code non devinable et ne voit que le minimum nécessaire.", "M"],
                            ["HIST-04", "Un document d’archive importé conserve sa source, sa date, son statut vérifié/non vérifié et la trace des corrections; une donnée incertaine n’est jamais présentée comme certifiée.", "M"],
                            ["PED-01", "Un responsable affecte les cours à un enseignant et à une cohorte; l’accès de l’enseignant est limité à ses affectations.", "M"],
                            ["PED-02", "Chaque inscription pédagogique relie l’étudiant, le programme, l’année/le semestre, la cohorte et les cours obligatoires ou optionnels suivis.", "M"],
                            ["PED-03", "Un responsable publie des ressources de cours versionnées dans le périmètre de la cohorte; l’étudiant ne consulte que les contenus de ses cours.", "M"],
                            ["PED-04", "Un enseignant peut créer des quiz formatifs; leurs règles, tentatives et résultats sont séparés des notes officielles sauf validation explicite du règlement.", "S"],
                            ["PED-05", "L’enseignant ou l’agent habilité enregistre la présence par séance ou selon le mode approuvé; l’étudiant consulte son historique et peut déposer une demande de correction.", "S"],
                            ["GRD-01", "Un responsable configure évaluations, poids, échelle, arrondis et seuils par cours/période avec version de règlement.", "M"],
                            ["GRD-02", "Le système valide plage, type, complétude, doublon et inscription avant d’accepter une note.", "M"],
                            ["GRD-03", "L’enseignant enregistre brouillon, soumet un lot et reçoit un état de validation avec erreurs localisées.", "M"],
                            ["GRD-04", "La validation et la publication suivent les approbations définies; un étudiant ne voit un résultat qu’après publication.", "M"],
                            ["GRD-05", "Après publication, une modification est une correction tracée, justifiée, approuvée et référençant l’ancienne valeur.", "M"],
                            ["GRD-06", "Moyennes, crédits, décision et relevés sont recalculés avec les règles approuvées et reproductibles.", "M"],
                            ["GRD-07", "Un relevé PDF affiche période, échelle, date d’émission, statut officiel, numéro de vérification et version.", "S"],
                            ["GRD-08", "Après publication, l’étudiant consulte en ligne ses notes, résultats d’examen et relevés depuis son compte; chaque nouvelle version ou rectification est signalée.", "M"],
                            ["PUB-01", "Les résultats d’examen publiables sont consultables sans compte par le public uniquement après décision institutionnelle conforme aux règles applicables; les champs exposés et la durée sont limités et approuvés. Les relevés individuels restent privés.", "M"],
                            ["PUB-02", "Aucun tiers ne parcourt les dossiers ou relevés; la vérification publique d’un document utilise un jeton aléatoire et révoquable et une réponse minimale.", "M"],
                            ["EXM-01", "Un étudiant dépose une demande de vérification/recours dans une fenêtre et un workflow définis par le règlement.", "S"],
                            ["LIFE-01", "Un étudiant ou agent dépose une demande d’absence, de rattrapage ou une réclamation en ligne; l’accès aux justificatifs médicaux est restreint au strict besoin.", "S"],
                            ["LIFE-02", "Les étudiants consultent leur emploi du temps, salles et changements publiés; les responsables habilités gèrent les séances, salles et rattrapages.", "S"],
                            ["DOC-01", "Les pièces sont privées, limitées en type/taille, contrôlées contre les logiciels malveillants et accessibles avec autorisation.", "M"],
                            ["COM-01", "Les notifications n’incluent pas de note sensible ou pièce jointe dans un courriel non sécurisé; elles renvoient vers le portail.", "M"],
                            ["RPT-01", "Les tableaux de bord indiquent définition, période, filtre, source et date de calcul de chaque indicateur.", "M"],
                            ["RPT-02", "Un export de données sensibles exige une permission dédiée, un motif si applicable et une entrée d’audit.", "M"],
                            ["INT-01", "Chaque import/export est idempotent, horodaté, réconciliable et conserve l’identifiant externe sans exposer de secrets.", "M"],
                            ["AUD-01", "Les actions sensibles conservent acteur, action, objet, horodatage, contexte, résultat et référence de corrélation.", "M"],
                            ["AI-01", "Un assistant conversationnel LLM répond aux questions publiques sur l’ENI-ABT, les formations, inscriptions, calendriers, pièces et procédures à partir d’un corpus institutionnel approuvé, à jour et référencé.", "M"],
                            ["AI-02", "Le chatbot cite les pages/documents utilisés, affiche leur date de mise à jour si connue, refuse d’inventer une règle, signale l’incertitude et propose le contact du service compétent.", "M"],
                            ["AI-03", "Le chatbot public avertit de ne pas saisir de données personnelles; les requêtes ne sont pas transmises au modèle avec des identifiants, notes, dossiers, paiements ou pièces, et les journaux sont minimisés.", "M"],
                            ["AI-04", "L’administration peut publier, versionner, retirer et auditer les sources du chatbot; toute réponse sans source suffisante est transférée à un humain ou indiquée comme non disponible.", "M"],
                            ["AI-05", "Le modèle et l’hébergement sont choisis après examen de confidentialité, localisation/transfert, coûts, disponibilité, sécurité et réversibilité; aucune donnée personnelle ne sert à entraîner un modèle externe sans autorisation.", "M"],
                            ["AI-06", "Toute IA d’analyse ou prédiction académique reste distincte du chatbot public, facultative, évaluée et sous supervision humaine; elle ne décide jamais d’une admission, d’une note ou d’un diplôme.", "S"],
                        ],
                        [0.12, 0.76, 0.12],
                    ),
                ],
            },
            {
                "subtitle": "C.4 Workflows cibles",
                "blocks": [
                    T(
                        ["Workflow", "Étapes / états", "Contrôles"],
                        [
                            ["Admission à l’inscription", "Candidat → dossier en ligne ou eCandidat → contrôle → décision autorisée → paiement si dû → dossier étudiant et inscription confirmée.", "Référence eCandidat; doublons; pièces; confirmation du prestataire; reçu et journal des transitions."],
                            ["Inscription annuelle", "Brouillon en ligne → pièces → paiement ou exonération → vérification → à corriger ou validée → reçu et inscription active.", "Le paiement ne vaut pas validation académique; cas d’échec et recours documentés; aucun passage physique dans le parcours nominal."],
                            ["Saisie des notes", "Brouillon enseignant → soumission → vérification → correction demandée ou validée → jury → publication.", "Affectation enseignant; valeurs; verrouillage après soumission; approbation; audit."],
                            ["Préparation pédagogique", "Programme/version → année et semestre → inscription pédagogique → affectation de l’enseignant → publication des ressources et activités → séance et présence.", "Cours obligatoires/optionnels selon le programme; rôle limité à l’affectation; ressources et quiz séparés des actes officiels de notation."],
                            ["Consultation des résultats", "Résultat validé → publication institutionnelle → accès privé par compte étudiant; publication publique distincte si autorisée.", "Journaliser les publications/corrections; ne pas révéler le relevé privé à un tiers."],
                            ["Correction après publication", "Demande motivée → vérification → approbation habilitée → nouvelle version → notification.", "Pas d’écrasement silencieux; conservation valeur antérieure; justification; droits limités."],
                            ["Dossier d’ancien étudiant", "Recherche sécurisée ou demande → vérification d’identité → contrôle des archives → délivrance de document officiel ou réponse motivée.", "Provenance et état de vérification; authentification du demandeur; délais et historique d’accès."],
                            ["Vérification d’un document", "Le diplômé partage un code de vérification → un tiers consulte une page publique limitée → validité confirmée ou document révoqué.", "Jeton non devinable, révocable, sans recherche par nom ou numéro étudiant; données affichées minimisées."],
                            ["Production d’attestation", "Demande en ligne → contrôle d’éligibilité → génération → signature/validation selon autorité → mise à disposition.", "Numérotation, vérification, statut, trace d’émission et révocation si erreur."],
                            ["Paiement", "Frais présentés → redirection prestataire → authentification prestataire → confirmation serveur → reçu et rapprochement.", "Ne jamais se fier à la seule page de retour navigateur; idempotence; remboursement et contestation traçables."],
                        ],
                        [0.2, 0.45, 0.35],
                    ),
                ],
            },
            {
                "subtitle": "C.5 Modèle de données conceptuel",
                "blocks": [
                    T(
                        ["Domaine", "Entités principales", "Principes"],
                        [
                            ["Identité", "Personne, compte, rôle, affectation de rôle, relation responsable-étudiant.", "Une personne peut avoir plusieurs rôles; identifiants institutionnels uniques et gouvernés."],
                            ["Structure", "DER, programme/filière, diplôme, niveau, année, semestre, groupe, salle.", "Référentiels versionnés et hiérarchie conforme à l’organisation validée."],
                            ["Admissions", "Campagne, dossier, vœu, pièce, décision, référence eCandidat.", "Portail maître à confirmer; données minimisées; cycle de vie et rétention définis."],
                            ["Scolarité", "Étudiant, statut étudiant/ancien, inscription, changement de parcours, attestation.", "Historique conservé; pas de suppression destructive des actes validés."],
                            ["Pédagogie", "Programme, inscription pédagogique, cours/UE, cohorte, affectation, ressource, quiz, séance et présence.", "Affectations bornées par année et périmètre; contenus accessibles uniquement aux personnes inscrites ou habilitées."],
                            ["Résultats", "Évaluation, barème/version, note, soumission, jury, publication, recours.", "Calcul reproductible; décimaux; versionnage; correction traçable; règles versionnées."],
                            ["Archives et titres", "Dossier académique historique, pièce, diplôme, attestation, version, provenance, preuve d’émission, statut et jeton de vérification.", "Classement, intégrité, conservation, restriction d’accès, correction et sort final selon calendrier approuvé."],
                            ["Paiement", "Frais, échéance, tentative, référence prestataire, transaction, reçu, rapprochement, remboursement.", "Le SI ne stocke pas les secrets de paiement; statut financier fondé sur la confirmation vérifiée du prestataire."],
                            ["Assistant d’information", "Source officielle, version, date d’effet, index, réponse, référence citée, transfert à un agent.", "Séparer corpus public et données privées; retirer/reconstruire l’index quand une source change."],
                            ["Finances/RH/ressources", "Échéance, rapprochement, agent, contrat, prêt bibliothèque et autres entités métier à valider.", "Domaines séparés, accès basé sur le besoin et conformité à valider."],
                            ["Audit", "Événement d’audit, événement d’intégration, journal de traitement.", "Stockage protégé contre modification non autorisée; rétention définie."],
                        ],
                        [0.2, 0.4, 0.4],
                    ),
                    P(
                        "Le modèle de données est conçu à partir des référentiels et règles approuvés "
                        "par l’ENI-ABT. Les noms, schémas et données de démonstration des solutions "
                        "du benchmark ne font pas partie de la base métier institutionnelle. "
                        "Contraintes recommandées : clés étrangères et contraintes d’unicité "
                        "en base; valeurs financières en décimaux et non flottants; dates explicites "
                        "avec fuseau; index sur identifiants et filtres usuels; migrations de schéma "
                        "versionnées et testées sur copie anonymisée."
                    ),
                ],
            },
            {
                "subtitle": "C.6 Intelligence artificielle responsable",
                "blocks": [
                    P(
                        "L’IA est une capacité optionnelle d’aide, jamais une autorité de "
                        "décision. Aucun modèle ne sélectionne un candidat, ne modifie une note, "
                        "ne prononce une sanction ou ne décide seul d’une orientation. Le pilote "
                        "commence par des indicateurs descriptifs et n’active une fonction "
                        "prédictive qu’après autorisation, évaluation et validation institutionnelles."
                    ),
                    P(
                        "Le premier usage prévu est un chatbot public propulsé par un modèle "
                        "de langage (LLM) avec génération augmentée par recherche (RAG). Il "
                        "recherche uniquement dans un corpus d’informations publiques ou "
                        "explicitement approuvées par l’ENI-ABT : formations, calendriers, "
                        "frais, formulaires, contacts et procédures publiées. Il fournit les "
                        "références utilisées, rappelle que les décisions officielles relèvent "
                        "des services compétents et dirige vers un agent lorsque la source est "
                        "absente, ancienne ou ambiguë. Le chatbot ne consulte aucun dossier "
                        "individuel et ne réalise ni inscription ni paiement à la place de l’usager."
                    ),
                    T(
                        ["ID", "Exigence / garde-fou"],
                        [
                            ["AI-GOV-01", "Assistant de procédures limité à des documents approuvés; réponses sourcées, abstention si incertain et transfert vers un service compétent; pas d’accès aux dossiers individuels."],
                            ["AI-GOV-02", "Toute alerte de risque académique est explicable, réservée aux personnels habilités et suivie d’une vérification humaine; aucun effet automatique sur les droits ou résultats de l’étudiant."],
                            ["AI-GOV-03", "Avant chaque usage : finalité, responsable, base juridique, sources, conservation, destinataires et transferts sont documentés et approuvés; aucune donnée personnelle n’entraîne un modèle externe sans autorisation écrite."],
                            ["AI-GOV-04", "Évaluer qualité, erreurs, faux positifs, biais par groupe pertinent et dérive; tracer la version, permettre signalement et désactivation, et revoir périodiquement les résultats."],
                        ],
                        [0.14, 0.86],
                    ),
                ],
            },
            {
                "subtitle": "C.7 Exigences non fonctionnelles",
                "blocks": [
                    T(
                        ["ID", "Exigence / critère de vérification"],
                        [
                            ["NFR-SEC", "HTTPS/TLS en production; cookies sécurisés; CSRF; protections XSS/injection; secrets hors dépôt; MFA pour comptes privilégiés si identité disponible."],
                            ["NFR-PRIV", "Inventaire et classification des données, finalité, minimisation, durée de conservation, droits et règles de transfert validés avec l’établissement et selon le droit applicable au Mali."],
                            ["NFR-AVAIL", "Sauvegarde chiffrée, contrôle de réussite, restauration testée, objectifs RPO/RTO approuvés."],
                            ["NFR-PERF", "p95 consultation simple <2 s sous charge convenue; opérations lourdes asynchrones et progrès visible."],
                            ["NFR-LOWNET", "Pages légères, pagination, reprise de formulaires non sensibles, compression et tests sur réseau mobile à faible débit."],
                            ["NFR-A11Y", "Viser WCAG 2.2 niveau AA; navigation clavier, focus visible, contrastes, labels, erreurs compréhensibles et tests lecteurs d’écran."],
                            ["NFR-I18N", "Français par défaut; anglais activable; formats locaux date/nombre/devise; textes d’interface séparés des données métier."],
                            ["NFR-OBS", "Logs structurés sans mot de passe ni contenu de note; métriques, traces corrélées, alertes sur erreurs et traitements en échec."],
                            ["NFR-MAINT", "Tests automatisés, migrations reproductibles, documentation, dépendances maintenues et revue des changements."],
                            ["NFR-PORT", "API versionnée et schéma OpenAPI; export des données institutionnelles et procédure de réversibilité documentés."],
                            ["NFR-ARCH", "Intégrité des archives contrôlée, métadonnées et provenance conservées, formats exportables et lisibles à long terme; durée et sort final définis par un calendrier de conservation validé."],
                            ["NFR-PAY", "Aucune donnée carte/PIN/CVV stockée dans le SI; notifications du prestataire signées ou authentifiées; transactions rapprochées selon une périodicité approuvée."],
                            ["NFR-LLM", "Corpus chatbot versionné et sourcé; tests de réponses non fondées, injections de consignes et fuites; requêtes/journaux minimisés; latence, disponibilité et coût suivis; désactivation et liens de secours."],
                        ],
                        [0.17, 0.83],
                    ),
                ],
            },
        ],
    },
    {
        "title": "D. Charte graphique et expérience utilisateur",
        "blocks": [
            P(
                "L’identité visuelle finale doit être fournie ou validée par l’ENI-ABT. "
                "Les couleurs et le logo du site public ne sont pas repris automatiquement "
                "comme charte approuvée."
            ),
            T(
                ["Sujet", "Exigence de conception"],
                [
                    ["Identité", "Logo officiel, couleurs et règles d’usage validés par la communication institutionnelle."],
                    ["Navigation", "Menus cohérents par profil, repères de fil d’Ariane, actions fréquentes visibles et tableaux filtrables."],
                    ["Responsive", "Conception mobile-first et fonctionnement sur navigateurs modernes; parcours simples pour smartphone."],
                    ["Accessibilité", "Cible WCAG 2.2 AA; clavier, lecteur d’écran, contrastes et erreurs de formulaire testés."],
                    ["Langue", "Interface française en premier; anglais si traduction relue et financée."],
                    ["Typographie", "Famille sans-serif lisible, tailles hiérarchisées, paragraphes courts et bonne lisibilité à faible résolution."],
                    ["États", "Brouillon, soumis, à corriger, validé, publié, archivé clairement distingués par texte et pas uniquement par couleur."],
                    ["Données sensibles", "Masquage par défaut, absence de note dans notification externe, confirmation avant export et avertissements de confidentialité."],
                ],
                [0.2, 0.8],
            ),
            C(
                "Expérience adaptée au contexte",
                "Prévoir une interface rapide sur connexion mobile, une saisie tolérante aux "
                "interruptions et des reçus d’opération. Le cache hors ligne ne doit jamais "
                "rendre accessibles des notes ou pièces sensibles sur un appareil partagé.",
            ),
        ],
    },
    {
        "title": "E. Spécificités, sécurité, livrables et planning",
        "blocks": [
            {
                "subtitle": "E.1 Sécurité, données et conformité",
                "blocks": [
                    B(
                        "Réaliser une analyse de risques et une classification des données avant les imports; déterminer avec l’ENI-ABT les obligations légales locales et contractuelles applicables.",
                        "Faire valider par les responsables compétents et consulter l’Autorité de Protection des Données à caractère Personnel du Mali (APDP) si nécessaire, notamment sur finalités, notices, droits d’accès, délais, archives, transferts et publications.",
                        "Séparer les environnements développement, test, préproduction et production; ne jamais copier une base réelle dans un environnement de développement sans autorisation et anonymisation.",
                        "Utiliser un contrôle d’accès par rôle et attribut (fonction, DER, cohorte); revue périodique des privilèges et retrait rapide des accès.",
                        "Protéger les notes et pièces avec accès restreint, TLS, chiffrement des sauvegardes et journalisation des consultations sensibles.",
                        "Protéger les archives par contrôle d’intégrité, provenance, métadonnées, versions, gestion des corrections et conservation/élimination selon un calendrier autorisé.",
                        "Séparer les résultats privés visibles par l’étudiant des listes de résultats publiées au public; définir et approuver au préalable les champs et la durée de toute publication nominative.",
                        "Utiliser un prestataire de paiement sélectionné après vérification contractuelle, technique et réglementaire; ne jamais stocker de données d’authentification de paiement dans l’application.",
                        "Traiter le chatbot public comme un service non authentifié exposé à des entrées adverses; isoler son corpus, limiter les requêtes, ne pas exposer d’outils d’accès aux données privées et journaliser sans conserver inutilement les conversations personnelles.",
                        "Gérer les corrections par workflows; préserver les décisions officielles et empêcher la modification silencieuse des résultats publiés.",
                        "Appliquer gestion de secrets, MFA/SSO si possible, mises à jour, scan de dépendances, tests de sécurité et procédure de divulgation/incidents.",
                        "Établir un plan de réponse aux incidents, contacts d’escalade, sauvegarde hors site autorisée, exercice de restauration et plan de continuité.",
                    ),
                    P(
                        "Référentiel de vérification : OWASP ASVS 5.0.0 pour les exigences "
                        "applicatives, NIST CSF 2.0 comme cadre de gestion des risques, "
                        "WCAG 2.2 pour l’accessibilité. Ces références guident les contrôles; "
                        "elles ne constituent pas une certification automatique."
                    ),
                ],
            },
            {
                "subtitle": "E.2 Intégration, migration et interopérabilité",
                "blocks": [
                    T(
                        ["Source / cible", "Méthode proposée", "Contrôles de sortie"],
                        [
                            ["eCandidat → SI", "API officielle privilégiée; à défaut, import/export autorisé, documenté et signé.", "Comptages, clés externes, rejets, doublons, idempotence et rapprochement métier."],
                            ["Benchmark fonctionnel → SI ENI-ABT", "Comparer les fonctions et interfaces documentées des solutions SIS/LMS avec les parcours et rôles validés par l’école.", "Matrice de correspondance approuvée; aucun modèle fournisseur ne remplace le modèle de données ni les règles de l’ENI-ABT."],
                            ["Archives papier et historiques", "Inventaire des registres, archives physiques, bases ou fichiers existants; numérisation/indexation selon procédures approuvées; import par lots.", "Provenance, qualité d’image, double contrôle des données critiques, statut vérifié/non vérifié, rapprochement, conservation des originaux et procès-verbal."],
                            ["Prestataire de paiement", "API/webhook serveur documenté; environnement de test; notification signée; référence stable pour l’opération; procédure de remboursement et contact d’incident.", "Confirmation côté serveur, idempotence, montant/devise, rapprochement, indisponibilité, délai de règlement et responsabilités contractuelles."],
                            ["Corpus du chatbot", "Sources officielles approuvées, version, date d’effet, propriétaire de contenu, indexation et retrait contrôlé.", "Réponse rattachée aux passages cités; tests d’obsolescence et de retrait; aucune donnée étudiante indexée."],
                            ["Fichiers bureautiques", "Gabarit d’import, prévalidation, rapport ligne par ligne et approbation.", "Journal d’import, source, opérateur, date et total accepté/rejeté."],
                            ["Services externes", "Contrat d’API, authentification service-à-service, quotas et délais de reprise.", "Tests contrat, monitoring, rotation secrets, procédure de panne."],
                        ],
                        [0.21, 0.42, 0.37],
                    ),
                    C(
                        "Prérequis de migration",
                        "La qualité, les volumes, les formats et les droits d’accès aux bases sources "
                        "n’ont pas encore été vérifiés. Aucune migration de données institutionnelles "
                        "ne peut être chiffrée ou autorisée avant l’inventaire et l’accord des "
                        "propriétaires de données.",
                    ),
                ],
            },
            {
                "subtitle": "E.3 Méthode, tests et critères de recette",
                "blocks": [
                    T(
                        ["Niveau", "Vérifications minimales"],
                        [
                            ["Tests unitaires", "Règles d’inscription, barèmes, calculs, permissions et validations de données."],
                            ["Tests d’intégration", "Base PostgreSQL, fichiers, notifications et connecteurs avec environnements de test."],
                            ["Tests de sécurité", "Accès interdit, IDOR, CSRF, XSS, injection, brute force, gestion de session, fuite par export et upload."],
                            ["Tests de migration", "Chargement répété, doublons, valeurs limites, rollback et rapprochement source/cible."],
                            ["Tests d’archives", "Retrouver un ancien dossier, vérifier la provenance d’une pièce, délivrer un document et valider sa révocation sans altérer l’original."],
                            ["Tests de paiement", "Succès, refus, timeout, retour navigateur falsifié, notification dupliquée, rapprochement, annulation et remboursement."],
                            ["Tests de publication", "Accès étudiant à son propre résultat, refus d’accès croisé et vérification publique sans exposition du relevé."],
                            ["Tests du chatbot", "Réponses avec sources approuvées, demande hors corpus, source périmée, tentative d’obtenir des données privées et escalade humaine."],
                            ["Tests accessibilité", "Audit clavier/lecteur d’écran, contraste et scénarios représentatifs."],
                            ["Recette métier", "Parcours candidat/étudiant, agent, enseignant, chef de DER, Directeur des études et direction sur jeux de données fictifs."],
                            ["Résilience", "Restauration de sauvegarde, indisponibilité d’un connecteur, reprise d’une tâche et rotation de secret."],
                        ],
                        [0.2, 0.8],
                    ),
                    B(
                        "100 % des exigences prioritaires Must approuvées sont reliées à un identifiant, un scénario de test, une preuve et un résultat de recette signé avant le passage en production.",
                        "Les notes publiées et le calcul des moyennes sont validés sur des cas de référence approuvés par les responsables pédagogiques.",
                        "Les défauts critiques et élevés de sécurité ou d’intégrité des résultats bloquent la mise en production.",
                        "La bascule n’est autorisée qu’après sauvegarde, répétition de migration, plan de retour arrière et accord des responsables.",
                    ),
                ],
            },
            {
                "subtitle": "E.4 Planning indicatif",
                "blocks": [
                    P(
                        "Les durées ci-dessous sont des ordres de grandeur de planification, non "
                        "un engagement de délai ou de coût. Elles supposent un sponsor actif, des "
                        "référents disponibles, des règles métier accessibles et une équipe "
                        "multidisciplinaire. Le calendrier se recalcule après la phase Discovery."
                    ),
                    T(
                        ["Phase", "Durée indicative", "Jalon / sortie"],
                        [
                            ["0. Autorisation et cadrage", "3–5 semaines", "Sponsor, équipe, cartographie, besoins confirmés, risques et décision systèmes maîtres."],
                            ["1. Architecture et prototype UX", "3–4 semaines", "Architecture validée, maquettes testées, modèle initial et backlog estimé."],
                            ["2. Socle technique et référentiels", "4–6 semaines", "Environnements, identité, rôles, audit, DER, programmes et calendrier."],
                            ["3. Pilote scolarité", "6–10 semaines", "Dossiers, inscriptions, intégration/import contrôlé et portails sur données de test."],
                            ["4. Paiements et mise en service des démarches", "3–6 semaines", "Prestataire validé, paiements testés, reçus, rapprochements et parcours d’inscription à distance."],
                            ["5. Archives et résultats", "6–10 semaines", "Pilote d’archives vérifiées, consultation personnelle, workflow de notes, publication et vérification de diplômes."],
                            ["6. Recette et déploiement contrôlé", "3–5 semaines", "Tests complets, formation, restauration testée, validation métier et mise en production autorisée."],
                            ["7. Extensions", "Lots ultérieurs", "Emploi du temps, salles, absences, réclamations, RH, bibliothèque et analytique avancée."],
                        ],
                        [0.25, 0.18, 0.57],
                    ),
                ],
            },
            {
                "subtitle": "E.5 Livrables attendus",
                "blocks": [
                    B(
                        "Étude de l’existant et cartographie des données/systèmes.",
                        "Backlog et cahier des règles métier signés par les propriétaires concernés.",
                        "Architecture logique et physique, modèle de données et contrats d’API.",
                        "Plan de classement, calendrier de conservation, procédures de numérisation et de migration des archives, règles de vérification des titres.",
                        "Parcours de paiement en ligne, contrat et configuration du prestataire, reçus, rapprochement et procédures de remboursement.",
                        "Corpus institutionnel approuvé, gestion de sources et procédure d’exploitation du chatbot LLM avec escalade humaine.",
                        "Maquettes des espaces, guide d’accessibilité et charte UX validée.",
                        "Code source, tests automatisés, migrations et documentation d’exploitation.",
                        "Scripts de migration, dictionnaire de correspondance et rapports de réconciliation.",
                        "Dossier sécurité, analyse de risques, gestion des secrets, politique de sauvegarde et plan d’incident.",
                        "Guide utilisateur par rôle, supports de formation et procédure d’assistance.",
                        "Rapports de test/recette, plan de déploiement, retour arrière et procès-verbal de validation.",
                    ),
                ],
            },
            {
                "subtitle": "E.6 Exploitation et maintenance",
                "blocks": [
                    B(
                        "Désigner le propriétaire du service, l’équipe de support, les horaires d’assistance et la chaîne d’escalade.",
                        "Définir un calendrier de mises à jour de sécurité et de dépendances avec fenêtre de maintenance.",
                        "Suivre disponibilité, erreurs, tâches en attente, capacité stockage, sauvegardes et expirations de certificats.",
                        "Tester périodiquement restauration, comptes d’urgence et procédure de reprise.",
                        "Documenter installation, configuration, mise à niveau, diagnostic et réversibilité.",
                        "Prévoir formation initiale des administrateurs, agents de scolarité, enseignants et utilisateurs finaux.",
                    ),
                ],
            },
            {
                "subtitle": "E.7 Risques et mesures de mitigation",
                "blocks": [
                    T(
                        ["Risque", "Niveau initial", "Mesure"],
                        [
                            ["Absence de mandat/accès aux services et données", "Élevé", "Sponsor officiel, référents et autorisations écrites avant collecte."],
                            ["Duplication avec eCandidat ou autre SI", "Élevé", "Cartographier la source de vérité; intégrer plutôt que reconstruire."],
                            ["Confusion des responsabilités académiques", "Élevé", "Définir et faire approuver la matrice de rôles du Directeur des études, des chefs de DER, des jurys, du SEGAL et de la scolarité avant le paramétrage."],
                            ["Règles de notation mal spécifiées", "Élevé", "Ateliers règlement, exemples de calcul, versionnement et signature métier."],
                            ["Conception fondée sur un produit générique sans validation locale", "Élevé", "Valider chaque workflow, rôle, règle de notation et état de dossier avec les responsables métier de l’ENI-ABT."],
                            ["Barèmes ou rôles figés dans le code", "Élevé", "Référentiels versionnés et paramétrables; contrôle de délégation; tests de non-régression et audit des changements."],
                            ["Fuite de notes/pièces", "Élevé", "Permissions par périmètre, chiffrement, audit, tests de sécurité, minimisation."],
                            ["Données historiques incohérentes", "Élevé", "Profilage, seuils qualité, lots pilotes, rapport de rapprochement et retour arrière."],
                            ["Perte ou impossibilité de vérifier les dossiers anciens", "Élevé", "Inventaire et archivage contrôlé, provenance, formats durables, vérification métier et restauration testée."],
                            ["Double débit ou paiement non rapproché", "Élevé", "Prestataire autorisé, vérification serveur, idempotence, journal financier et tests des notifications en échec."],
                            ["Résultats exposés à des tiers", "Élevé", "Accès privé par défaut; listes publiques séparées et autorisées; vérification documentaire à divulgation minimale."],
                            ["Chatbot invente une procédure ou divulgue des données", "Élevé", "RAG sur corpus approuvé, citations, abstention, séparation des accès, tests adversariaux et escalade humaine."],
                            ["Connexion internet irrégulière", "Moyen/élevé", "UX légère, tests faible débit, processus dégradé et consigne de synchronisation."],
                            ["Dépendance à un seul développeur", "Moyen/élevé", "Revue de code, documentation, transfert de compétences, CI et propriété institutionnelle."],
                            ["Extension trop vaste et retards", "Élevé", "MVP à périmètre fermé, priorisation Must/Should/Could et jalons de sortie."],
                            ["IA biaisée ou utilisée comme décision", "Élevé", "Exclure du pilote; supervision humaine; évaluation et gouvernance avant expérimentation."],
                        ],
                        [0.33, 0.15, 0.52],
                    ),
                ],
            },
            {
                "subtitle": "E.8 Décisions à obtenir avant lancement",
                "blocks": [
                    B(
                        "Nommer sponsor, propriétaire produit, responsables métier et référent sécurité/données.",
                        "Confirmer le périmètre pilote et le ou les départements participants.",
                        "Décider si eCandidat reste le système maître des dossiers candidats et quelles interfaces sont autorisées.",
                        "Fournir règlements de scolarité, d’évaluation, de délibération, de conservation et modèles officiels.",
                        "Valider les responsabilités du Directeur des études, des chefs de DER, du SEGAL, de la scolarité et des jurys, ainsi que les délégations de validation.",
                        "Décider de la politique des archives historiques, des délais de conservation, des originaux, des signatures et du processus de vérification des diplômes.",
                        "Valider les frais, exemptions, remboursements, modes de paiement, prestataire, contrat, rapprochement et gestion des réclamations.",
                        "Désigner les propriétaires de contenu du chatbot et approuver les réponses institutionnelles, sources, contacts et fréquence de mise à jour.",
                        "Décider si les listes de résultats sont publiques; définir la base, les champs autorisés, la durée de mise en ligne et la procédure de rectification.",
                        "Définir hébergement, résidence des données, sauvegarde, accès réseau, identité et budget d’exploitation.",
                        "Confirmer nombre d’utilisateurs, pics de charge, postes partagés, appareils et connectivité réelle.",
                        "Valider les obligations nationales de protection des données et la politique d’accès aux dossiers.",
                        "Valider la matrice des rôles et délégations, les barèmes, les cycles académiques et les règles de publication des résultats.",
                    ),
                ],
            },
        ],
    },
    {
        "title": "Références et traçabilité des sources",
        "blocks": [
            P(
                "Le cahier distingue les faits observés des propositions. Les pages institutionnelles "
                "et documentations officielles de produits citées ci-dessous ont été consultées le "
                "11 novembre 2026. Elles constituent des références de cadrage, pas une preuve de "
                "l’existant interne ni une validation des choix par l’ENI-ABT."
            ),
            P(
                "La structure s’appuie sur les pratiques d’ingénierie des exigences décrites "
                "dans ISO/IEC/IEEE 29148:2018, notamment le contenu, le format et la "
                "traçabilité des éléments d’information. La norme donne des recommandations "
                "de format, sans imposer une charte graphique universelle. Son édition 2018 "
                "est indiquée en cours de révision par l’ISO. Aucune conformité intégrale "
                "n’est revendiquée pour ce document."
            ),
            T(
                ["Source", "Usage dans le cahier"],
                [
                    ["Cahier des Charges Complement-Projet tutoré.docx fourni par l’utilisateur", "Structure de référence A–E; le document source est un formulaire/canevas non renseigné avec champs à compléter."],
                    ["https://eni-abt.ml/", "Présentation publique, concours/résultats publiés, événements et liens vers les formations."],
                    ["https://eni-abt.ml/presentation/", "Historique public de l’établissement et passage au système LMD."],
                    ["https://eni-abt.ml/formations/", "Domaines/DER et catalogue public des formations."],
                    ["https://ecandidat.eni-abt.ml/", "Parcours public de candidature en ligne et suivi de dossier."],
                    ["https://eni-abt.ml/contact/", "Coordonnées publiques et horaires affichés."],
                    ["https://eni-abt.ml/wp-sitemap.xml", "Inventaire public des pages, articles et événements."],
                    ["https://docs.frappe.io/education/student-applicant", "Documentation officielle Frappe Education : candidature en ligne, approbation et passage à l’inscription."],
                    ["https://docs.frappe.io/education/program-enrollment", "Documentation officielle : inscription d’un étudiant à un programme, une année/terme et des cours obligatoires ou optionnels."],
                    ["https://docs.frappe.io/education/assessment_result", "Documentation officielle : journal des notes obtenues par étudiant et évaluation."],
                    ["https://docs.frappe.io/education/student-portal", "Documentation officielle : portail étudiant, calendrier, résultats, frais/paiement, présences et profil. Le prestataire de paiement décrit est un exemple fournisseur, non une recommandation locale."],
                    ["https://www.ellucian.com/products/student", "Présentation fournisseur d’Ellucian Student : cycle étudiant, admissions, programmes, portails et fonctions IA décrites par l’éditeur."],
                    ["https://www.ellucian.com/products/student/student-information-systems", "Présentation fournisseur du SIS et des tableaux de bord par rôle; affirmations commerciales, non évaluation indépendante."],
                    ["https://docs.openedx.org/en/latest/educators/quickstarts/build_a_course.html", "Documentation officielle Open edX : création, programmation et organisation des cours."],
                    ["https://docs.openedx.org/en/latest/educators/how-tos/data/view_learner_grades.html", "Documentation officielle : carnet de notes et rapports par cours; les exports contiennent des données personnelles."],
                    ["https://docs.openedx.org/en/latest/community/release_notes/verawood/verawood_rp.html", "Notes officielles Open edX Verawood (juillet 2026) : rôles d’équipe et périmètre d’accès par cours/organisation."],
                    ["https://docs.openedx.org/en/latest/community/release_notes/verawood/verawood_marketing_notes.html", "Notes officielles Open edX Verawood : assistant IA apprenant ancré dans une leçon; fonction d’aide pédagogique, pas assistant de procédure publique."],
                    ["https://docs.openedx.org/en/latest/learners/SFD_content_availability.html", "Documentation officielle Open edX : contenu et cours archivés; ne décrit pas l’archivage des dossiers académiques institutionnels."],
                    ["https://www.postgresql.org/support/versioning/", "Politique officielle de support de PostgreSQL."],
                    ["https://asvs.owasp.org/", "OWASP ASVS 5.0.0, exigences de vérification sécurité applicative."],
                    ["https://www.nist.gov/cyberframework", "NIST Cybersecurity Framework 2.0."],
                    ["https://www.nist.gov/itl/ai-risk-management-framework", "NIST AI RMF 1.0, cadre volontaire de gestion des risques IA (statut à reconfirmer au lancement)."],
                    ["https://apdp.ml/", "Autorité de Protection des Données à caractère Personnel du Mali; point de contact et cadre local à confirmer avec l’ENI-ABT."],
                    ["https://www.iso.org/standard/62542.html", "ISO 15489-1:2016, principes de création, capture, métadonnées, responsabilités et gestion des documents d’activité dans le temps; référence de conception, pas une certification."],
                    ["Capture de besoins fournie dans la conversation", "Liste préliminaire des besoins sur notes, examens, relevés, diplômes, inscriptions, absence, emploi du temps, salles, rattrapages et formulaires; à valider avec les services."],
                    ["https://www.w3.org/TR/WCAG22/", "Recommandation W3C WCAG 2.2."],
                    ["https://spec.openapis.org/oas/latest.html", "Spécification de description d’API OpenAPI."],
                    ["https://www.iso.org/standard/72089.html", "ISO/IEC/IEEE 29148:2018, ingénierie des exigences : processus, contenus des éléments d’information et recommandations de format. L’ISO indique que l’édition 2018 est en cours de révision."],
                ],
                [0.48, 0.52],
            ),
            C(
                "Mention de responsabilité",
                "Document de cadrage préparé pour discussion. Il ne vaut ni validation "
                "institutionnelle, ni autorisation de traitement des données, ni engagement "
                "budgétaire, contractuel ou de mise en production.",
            ),
        ],
    },
]


def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill.replace("#", ""))
    tc_pr.append(shd)


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def add_page_field(paragraph) -> None:
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for element in (begin, instr, separate, text, end):
        run._r.append(element)


def add_docx_table(doc, headers, rows):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.autofit = True
    header = table.rows[0]
    set_repeat_table_header(header)
    for i, value in enumerate(headers):
        cell = header.cells[i]
        cell.text = plain_text(value)
        set_cell_shading(cell, PALE)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)
                run.font.size = Pt(8)
    for row_index, values in enumerate(rows):
        cells = table.add_row().cells
        for i, value in enumerate(values):
            cells[i].text = plain_text(value)
            cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if row_index % 2 == 1:
                set_cell_shading(cells[i], "F7F7F7")
            for paragraph in cells[i].paragraphs:
                paragraph.paragraph_format.space_after = Pt(2)
                for run in paragraph.runs:
                    run.font.size = Pt(8)
                    run.font.name = "Arial"
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return table


def add_docx_blocks(doc, blocks):
    for block in blocks:
        if "subtitle" in block:
            doc.add_heading(plain_text(block["subtitle"]), level=2)
            add_docx_blocks(doc, block["blocks"])
            continue
        kind = block["type"]
        if kind == "p":
            paragraph = doc.add_paragraph(plain_text(block["text"]))
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        elif kind == "bullets":
            for item in block["items"]:
                paragraph = doc.add_paragraph(style="List Bullet")
                paragraph.add_run(plain_text(item))
                paragraph.paragraph_format.space_after = Pt(3)
        elif kind == "table":
            add_docx_table(doc, block["headers"], block["rows"])
        elif kind == "callout":
            para = doc.add_paragraph()
            run = para.add_run(plain_text(block["title"]) + ": ")
            run.bold = True
            para.add_run(plain_text(block["text"]))


def build_docx():
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.9)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.95)
    section.right_margin = Inches(0.85)
    section.different_first_page_header_footer = True
    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10)
    styles["Normal"].font.color.rgb = RGBColor(0, 0, 0)
    styles["Normal"].paragraph_format.space_after = Pt(7)
    for name, size in [
        ("Title", 22),
        ("Heading 1", 15),
        ("Heading 2", 12),
        ("Heading 3", 11),
    ]:
        style = styles[name]
        style.font.name = "Arial"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.keep_with_next = True

    header = section.header.paragraphs[0]
    header.text = "ENI-ABT   Cahier des charges"
    header.style = styles["Caption"]
    header.runs[0].font.color.rgb = RGBColor(0, 0, 0)
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer.add_run(f"{DOCUMENT_REFERENCE}   Version {DOCUMENT_VERSION}   Page ")
    add_page_field(footer)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(72)
    run = p.add_run("ÉCOLE NATIONALE D’INGÉNIEURS\n")
    run.bold = True
    run.font.size = Pt(14)
    title = p.add_run("ABDERHAMANE BABA TOURÉ\n")
    title.bold = True
    title.font.size = Pt(14)
    sub = doc.add_paragraph("CAHIER DES CHARGES")
    sub.paragraph_format.space_before = Pt(38)
    sub.runs[0].bold = True
    sub.runs[0].font.size = Pt(20)
    doc.add_paragraph("Système d’information institutionnel intégré")
    doc.add_paragraph()
    cover = doc.add_table(rows=7, cols=2)
    cover.style = "Table Grid"
    metadata = [
        ("Institution", "École Nationale d’Ingénieurs Abderhamane Baba Touré (ENI-ABT)"),
        ("Référence", DOCUMENT_REFERENCE),
        ("Version / date", f"{DOCUMENT_VERSION} - {AS_OF}"),
        ("Révision", "Benchmark de SIS existants et intégration des fonctions pédagogiques; ajout du Directeur des études."),
        ("Statut", "Projet de document, pour revue et validation institutionnelle"),
        ("Rédaction", "Équipe projet, à compléter"),
        ("Approbation", "Nom, fonction, signature et date, à renseigner"),
    ]
    for i, (left, right) in enumerate(metadata):
        cover.cell(i, 0).text = plain_text(left)
        cover.cell(i, 1).text = plain_text(right)
        set_cell_shading(cover.cell(i, 0), PALE)
        cover.cell(i, 0).paragraphs[0].runs[0].bold = True
    note = doc.add_paragraph(
        "Cadrage préparé à partir du canevas fourni, d’un benchmark de systèmes d’information "
        "d’enseignement supérieur et de sources publiques consultées. Ce document ne constitue pas une approbation "
        "officielle ni une autorisation de traiter des données réelles."
    )
    note.paragraph_format.space_before = Pt(24)
    note.runs[0].font.size = Pt(9)

    doc.add_page_break()
    doc.add_heading("Sommaire", level=1)
    for sec in SECTIONS:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.add_run(plain_text(sec["title"])).bold = True
        for block in sec["blocks"]:
            if "subtitle" in block:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.25)
                p.paragraph_format.space_after = Pt(2)
                p.add_run(plain_text(block["subtitle"]))

    for sec in SECTIONS:
        doc.add_page_break()
        doc.add_heading(plain_text(sec["title"]), level=1)
        add_docx_blocks(doc, sec["blocks"])

    doc.core_properties.title = "Cahier des charges du système d’information de l’ENI-ABT"
    doc.core_properties.subject = "Cahier des charges fonctionnel et technique"
    doc.core_properties.author = "Équipe projet, à valider par l’ENI-ABT"
    doc.core_properties.keywords = "ENI-ABT, système d'information, scolarité, notes, architecture"
    doc.save(DOCX_PATH)


def xml(text: str) -> str:
    return escape(plain_text(text), quote=False).replace("\n", "<br/>")


def make_reportlab_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="BodyFR", parent=styles["BodyText"], fontName="Arial",
        fontSize=9.5, leading=13.2, textColor=colors.HexColor(INK),
        alignment=TA_JUSTIFY, spaceAfter=7,
    ))
    styles.add(ParagraphStyle(
        name="BulletFR", parent=styles["BodyFR"], leftIndent=13, firstLineIndent=-8,
        bulletIndent=1, spaceAfter=3.5,
    ))
    styles.add(ParagraphStyle(
        name="H1FR", parent=styles["Heading1"], fontName="Arial-Bold",
        fontSize=15, leading=19, textColor=colors.HexColor(BLACK),
        spaceBefore=8, spaceAfter=8, keepWithNext=True,
    ))
    styles.add(ParagraphStyle(
        name="H2FR", parent=styles["Heading2"], fontName="Arial-Bold",
        fontSize=12, leading=15, textColor=colors.HexColor(BLACK),
        spaceBefore=10, spaceAfter=5, keepWithNext=True,
    ))
    styles.add(ParagraphStyle(
        name="TableFR", fontName="Arial",         fontSize=7.7, leading=10,
        textColor=colors.HexColor(INK), spaceAfter=0, alignment=TA_LEFT,
    ))
    styles.add(ParagraphStyle(
        name="TableHeadFR", fontName="Arial-Bold", fontSize=7.8, leading=10,
        textColor=colors.black, spaceAfter=0, alignment=TA_LEFT,
    ))
    styles.add(ParagraphStyle(
        name="SmallFR", fontName="Arial", fontSize=7.2, leading=9.2,
        textColor=colors.HexColor(MUTED), spaceAfter=5,
    ))
    styles.add(ParagraphStyle(
        name="CoverInstitution", fontName="Arial-Bold", fontSize=12,
        leading=16, textColor=colors.black,
    ))
    styles.add(ParagraphStyle(
        name="CoverTitle", fontName="Arial-Bold", fontSize=22,
        leading=27, textColor=colors.black,
    ))
    styles.add(ParagraphStyle(
        name="CoverSub", fontName="Arial", fontSize=12,
        leading=16, textColor=colors.black,
    ))
    styles.add(ParagraphStyle(
        name="TOC1FR", fontName="Arial-Bold", fontSize=9.5, leading=14,
        leftIndent=0, firstLineIndent=0, textColor=colors.black,
    ))
    styles.add(ParagraphStyle(
        name="TOCTitleFR", parent=styles["H1FR"],
    ))
    styles.add(ParagraphStyle(
        name="TOC2FR", fontName="Arial", fontSize=8, leading=11,
        leftIndent=14, firstLineIndent=0, textColor=colors.HexColor(MUTED),
    ))
    return styles


def draw_page(canvas, doc):
    canvas.saveState()
    w, h = A4
    if doc.page > 1:
        canvas.setStrokeColor(colors.HexColor(GRID))
        canvas.setLineWidth(0.4)
        canvas.line(25 * mm, h - 17 * mm, w - 25 * mm, h - 17 * mm)
        canvas.setFont("Arial-Bold", 8)
        canvas.setFillColor(colors.black)
        canvas.drawString(25 * mm, h - 13 * mm, "ENI-ABT")
        canvas.setFont("Arial", 8)
        canvas.setFillColor(colors.HexColor(MUTED))
        canvas.drawRightString(w - 25 * mm, h - 13 * mm, "CAHIER DES CHARGES")
        canvas.line(25 * mm, 16 * mm, w - 25 * mm, 16 * mm)
        canvas.setFont("Arial", 8)
        canvas.setFillColor(colors.black)
        canvas.drawString(25 * mm, 10.5 * mm, f"{DOCUMENT_REFERENCE}   Version {DOCUMENT_VERSION}")
        canvas.drawRightString(w - 25 * mm, 10.5 * mm, f"Page {doc.page}")
    canvas.restoreState()


class SpecDocTemplate(BaseDocTemplate):
    def __init__(self, filename, styles, **kwargs):
        super().__init__(filename, **kwargs)
        self.styles = styles
        frame = Frame(
            19 * mm, 19 * mm, A4[0] - 38 * mm, A4[1] - 39 * mm,
            leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
            id="normal",
        )
        self.addPageTemplates([
            PageTemplate(id="cover", frames=frame, onPage=draw_page),
            PageTemplate(id="normal", frames=frame, onPage=draw_page),
        ])
        self.toc = TableOfContents()
        self.toc.levelStyles = [styles["TOC1FR"], styles["TOC2FR"]]
        self.toc.dotsMinLevel = 0
        self.heading_count = 0

    def beforeDocument(self):
        self.heading_count = 0

    def afterFlowable(self, flowable):
        if not isinstance(flowable, Paragraph):
            return
        style = flowable.style.name
        if style not in ("H1FR", "H2FR"):
            return
        text = flowable.getPlainText()
        if text == "Sommaire":
            return
        level = 0 if style == "H1FR" else 1
        self.heading_count += 1
        key = f"heading-{level}-{self.page}-{self.heading_count}"
        self.canv.bookmarkPage(key)
        self.canv.addOutlineEntry(text, key, level=level, closed=False)
        self.notify("TOCEntry", (level, text, self.page, key))


def make_table(block, styles, available_width):
    headers = block["headers"]
    rows = block["rows"]
    widths = block.get("widths")
    if widths is None:
        widths = [1 / len(headers)] * len(headers)
    data = [[Paragraph(xml(h), styles["TableHeadFR"]) for h in headers]]
    for row in rows:
        data.append([Paragraph(xml(str(cell)), styles["TableFR"]) for cell in row])
    table = Table(
        data,
        colWidths=[available_width * ratio for ratio in widths],
        repeatRows=1,
        hAlign="LEFT",
        splitByRow=1,
    )
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(PALE)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F7F7")]),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor(GRID)),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
    ]))
    return table


def add_pdf_blocks(story, blocks, styles, available_width):
    for block in blocks:
        if "subtitle" in block:
            story.append(Paragraph(xml(block["subtitle"]), styles["H2FR"]))
            add_pdf_blocks(story, block["blocks"], styles, available_width)
            continue
        kind = block["type"]
        if kind == "p":
            story.append(Paragraph(xml(block["text"]), styles["BodyFR"]))
        elif kind == "bullets":
            for item in block["items"]:
                story.append(Paragraph("- " + xml(item), styles["BulletFR"]))
        elif kind == "table":
            story.append(make_table(block, styles, available_width))
            story.append(Spacer(1, 7))
        elif kind == "callout":
            story.append(Paragraph(
                f"<b>{xml(block['title'])}:</b> {xml(block['text'])}",
                styles["BodyFR"],
            ))


def build_pdf():
    font_dir = Path(r"C:\Windows\Fonts")
    pdfmetrics.registerFont(TTFont("Arial", str(font_dir / "arial.ttf")))
    pdfmetrics.registerFont(TTFont("Arial-Bold", str(font_dir / "arialbd.ttf")))
    pdfmetrics.registerFont(TTFont("Arial-Italic", str(font_dir / "ariali.ttf")))
    pdfmetrics.registerFontFamily(
        "Arial", normal="Arial", bold="Arial-Bold",
        italic="Arial-Italic", boldItalic="Arial-Bold",
    )

    styles = make_reportlab_styles()
    doc = SpecDocTemplate(
        str(PDF_PATH),
        styles,
        pagesize=A4,
        title="Cahier des charges du système d’information de l’ENI-ABT",
        author="Équipe projet, à valider par l’ENI-ABT",
        subject="Cahier des charges fonctionnel et technique",
        leftMargin=25 * mm,
        rightMargin=25 * mm,
        topMargin=24 * mm,
        bottomMargin=22 * mm,
    )
    available_width = A4[0] - 50 * mm
    story = []
    story.append(Spacer(1, 29 * mm))
    story.append(Paragraph(
        "ÉCOLE NATIONALE D’INGÉNIEURS ABDERHAMANE BABA TOURÉ",
        styles["CoverInstitution"],
    ))
    story.append(Spacer(1, 32 * mm))
    story.append(Paragraph("CAHIER DES CHARGES", styles["CoverTitle"]))
    story.append(Spacer(1, 3 * mm))
    story.append(Paragraph(
        "Système d’information institutionnel intégré",
        styles["CoverSub"],
    ))
    story.append(Spacer(1, 14 * mm))
    cover_rows = [
        [Paragraph("<b>Référence</b>", styles["TableFR"]), Paragraph(DOCUMENT_REFERENCE, styles["TableFR"])],
        [Paragraph("<b>Version et date</b>", styles["TableFR"]), Paragraph(f"{DOCUMENT_VERSION} - {AS_OF}", styles["TableFR"])],
        [Paragraph("<b>Révision</b>", styles["TableFR"]), Paragraph("Benchmark de SIS existants, fonctions pédagogiques intégrées et ajout du Directeur des études", styles["TableFR"])],
        [Paragraph("<b>Statut</b>", styles["TableFR"]), Paragraph("Projet de document, pour revue et validation", styles["TableFR"])],
        [Paragraph("<b>Rédaction</b>", styles["TableFR"]), Paragraph("Équipe projet, à compléter", styles["TableFR"])],
        [Paragraph("<b>Approbation</b>", styles["TableFR"]), Paragraph("Nom, fonction, signature et date, à renseigner", styles["TableFR"])],
    ]
    cover_table = Table(cover_rows, colWidths=[39 * mm, available_width - 39 * mm])
    cover_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor(PALE)),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor(GRID)),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(cover_table)
    story.append(Spacer(1, 10 * mm))
    story.append(Paragraph(
        "Cadrage préparé à partir du canevas fourni, d’un benchmark de systèmes d’information "
        "d’enseignement supérieur et de sources publiques consultées. Ce document ne constitue pas une approbation "
        "officielle ni une autorisation de traiter des données réelles.",
        styles["SmallFR"],
    ))
    story.append(NextPageTemplate("normal"))
    story.append(PageBreak())
    story.append(Paragraph("Sommaire", styles["TOCTitleFR"]))
    story.append(Spacer(1, 4 * mm))
    story.append(doc.toc)
    story.append(PageBreak())

    for sec in SECTIONS:
        story.append(Paragraph(xml(sec["title"]), styles["H1FR"]))
        add_pdf_blocks(story, sec["blocks"], styles, available_width)

    DOCS.mkdir(parents=True, exist_ok=True)
    doc.multiBuild(story)


def main():
    DOCS.mkdir(parents=True, exist_ok=True)
    build_docx()
    build_pdf()
    print(f"DOCX: {DOCX_PATH}")
    print(f"PDF:  {PDF_PATH}")


if __name__ == "__main__":
    main()
