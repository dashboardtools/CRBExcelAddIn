# ORBIT Data Access Plugin — v2.0
## Guide d'architecture et d'installation

---

## Architecture globale

```
┌─────────────────────────────────────────────────────────────────────┐
│  Excel (XLAM)                                                       │
│                                                                     │
│  Ruban ORBIT                                                        │
│    ├── Refresh All ──────────────────────────────────────────────┐  │
│    ├── Requête…                                                   │  │
│    └── Signaler une dysqualité                                    │  │
│                                                                   │  │
│  modMain ──────────────────────────────────────────────────────┐  │  │
│  modContext  → CaptureContext() : user, fichier, onglet        │  │  │
│  modParameters → ConsolidateParameters() : dialog unique       │  │  │
│  modSQL      → ParseParameters(), ExecuteQuery()               │  │  │
│  modAPI      → CallApi() : REST JSON/CSV                       │  │  │
│  modExcel    → InjectDataTable(), UpdateTraceSheet()           │  │  │
│  modDysqualite → SignalerDysqualite(), PrepareDQEmail()        │  │  │
│  modSharePoint → Catalogue, Audit, DQ Issues, Contacts         │  │  │
│  modSettings → Registre Windows, Custom XML, Logging           │  │  │
│  modConnections → DSN ODBC MySQL/PG/SQL Server                 │  │  │
│                                                                   │  │
│  frmConsolidatedParams → Dialog unique multi-requêtes          │  │  │
│  frmDysqualite         → Formulaire signalement DQ             │  │  │
└───────────────────────────────────────────────────────────────────┘  │
                            │                                          │
              ┌─────────────┴──────────────┐                          │
              ▼                            ▼                           │
   SharePoint (REST API SSO)      Bases de données (ODBC DSN)        │
     ├── OrbitQueries               MySQL, PostgreSQL,                │
     ├── OrbitAuditLog              SQL Server                        │
     ├── DataQualityIssues                                            │
     ├── DataQualityContacts        APIs REST externes                │
     └── OrbitParameterValues       (Bearer token / SSO)             │
                                                                       │
                            └─────── Onglet _ORBIT_TRACE ────────────┘
                                     (traçabilité locale classeur)
```

---

## Structure du projet (14 fichiers)

```
OrbitVBA_v2/
│
├── Modules VBA (.bas)
│   ├── modMain.bas          ← Point d'entrée, callbacks ruban, =ORBIT()
│   ├── modContext.bas       ← Capture contexte : user, fichier, onglet, table
│   ├── modParameters.bas   ← Consolidation multi-requêtes, validation
│   ├── modSQL.bas           ← Parser {{Param:type}}, exécution ADO
│   ├── modAPI.bas           ← Appels REST JSON/CSV, UrlEncode
│   ├── modExcel.bas         ← Injection ListObject, Refresh All, _ORBIT_TRACE
│   ├── modDysqualite.bas   ← Signalement DQ, email Outlook, SP historisation
│   ├── modSharePoint.bas   ← Client REST SP : catalogue, audit, DQ, contacts
│   ├── modSettings.bas     ← Registre Windows, Custom XML, logging fichier
│   └── modConnections.bas  ← Chaînes connexion DSN ODBC
│
├── UserForms (.frm)
│   ├── frmConsolidatedParams.frm  ← Dialog unique consolidé multi-requêtes
│   └── frmDysqualite.frm         ← Formulaire signalement dysqualité
│
└── Config
    ├── customUI.xml         ← Définition du ruban personnalisé
    └── README_v2.md         ← Ce fichier
```

---

## Listes SharePoint à créer

### 1. OrbitQueries — Catalogue de requêtes

| Colonne interne | Type | Obligatoire | Description |
|---|---|---|---|
| Title | Ligne de texte | ✓ | Nom lisible |
| QueryId | Ligne de texte | ✓ | Identifiant machine |
| SourceType | Choix | ✓ | SQL / API |
| SQLQuery | Plusieurs lignes | | Requête SQL avec {{Params}} |
| ApiEndpoint | Ligne de texte | | URL API avec {{Params}} |
| TargetSheet | Ligne de texte | ✓ | Nom de la feuille cible |
| TargetTable | Ligne de texte | ✓ | Nom du tableau ListObject cible |
| DataSource | Ligne de texte | ✓ | Clé DSN ODBC |
| Description | Plusieurs lignes | | Description fonctionnelle |
| AllowedValues | Plusieurs lignes | | Valeurs autorisées (une/ligne) |
| IsActive | Oui/Non | ✓ | Activer/désactiver la requête |

### 2. OrbitAuditLog — Journal d'audit centralisé

| Colonne interne | Type | Description |
|---|---|---|
| Title | Ligne de texte | QueryId |
| UserLogin | Ligne de texte | Login Windows |
| UserFullName | Ligne de texte | Nom complet |
| WorkbookName | Ligne de texte | Nom du fichier Excel |
| WorkbookPath | Ligne de texte | Chemin complet |
| ActiveSheet | Ligne de texte | Onglet actif |
| ActiveTable | Ligne de texte | Tableau actif |
| QueryId | Ligne de texte | Requête exécutée |
| SourceType | Ligne de texte | SQL / API |
| ApiEndpoint | Ligne de texte | URL appelée |
| ParametersJson | Plusieurs lignes | Paramètres JSON |
| RowCount | Nombre | Lignes retournées |
| DurationMs | Nombre | Durée en millisecondes |
| ExecutionStatus | Ligne de texte | OK / ERREUR |
| ErrorMessage | Plusieurs lignes | Message d'erreur |
| OrbitVersion | Ligne de texte | Version du plugin |
| MachineName | Ligne de texte | Nom du poste |
| ExecutedAt | Date/Heure | Horodatage |

### 3. DataQualityIssues — Incidents de qualité

| Colonne interne | Type | Description |
|---|---|---|
| Title | Ligne de texte | IncidentID (DQ-YYYYMMDD-...) |
| DeclarationDate | Date/Heure | Horodatage déclaration |
| UserLogin | Ligne de texte | Déclarant |
| WorkbookName | Ligne de texte | Fichier source |
| WorkbookPath | Ligne de texte | Chemin |
| ActiveSheet | Ligne de texte | Onglet concerné |
| ActiveTable | Ligne de texte | Tableau concerné |
| QueryId | Ligne de texte | Requête/API source |
| ParametersJson | Plusieurs lignes | Filtres appliqués |
| Description | Plusieurs lignes | Commentaire utilisateur |
| Criticality | Choix | Faible / Moyen / Critique |
| Status | Choix | Nouveau / En cours / Résolu |
| ResponsibleTeam | Ligne de texte | Équipe en charge |

### 4. DataQualityContacts — Contacts DQ

| Colonne interne | Type | Description |
|---|---|---|
| Title | Ligne de texte | Nom contact / équipe |
| Team | Ligne de texte | Nom de l'équipe |
| DataDomain | Ligne de texte | Domaine de données |
| Email | Ligne de texte | Adresse email |
| Criticality | Ligne de texte | Niveau concerné |
| Scope | Ligne de texte | Périmètre |

### 5. OrbitParameterValues — Valeurs autorisées

| Colonne interne | Type | Description |
|---|---|---|
| Title | Ligne de texte | Nom du paramètre (ex: portfolio) |
| AllowedValues | Plusieurs lignes | Une valeur par ligne |
| QueryId | Ligne de texte | Optionnel — restreindre à une requête |

---

## Installation — checklist

```
□ 1. Installer MySQL Connector/ODBC 8.x
□ 2. Configurer le DSN dans Windows (odbcad32.exe)
□ 3. Créer les 5 listes SharePoint (colonnes ci-dessus)
□ 4. Alt+F11 → copier les 10 modules .bas
□ 5. Créer les 2 UserForms (frmConsolidatedParams, frmDysqualite)
□ 6. Outils > Références > ADO 6.1
□ 7. Ajouter le ruban via Custom UI Editor (customUI.xml)
□ 8. Enregistrer en .xlam
□ 9. Fichier > Options > Compléments > Activer ORBIT
□ 10. ORBIT > Paramètres > saisir l'URL SharePoint
□ 11. TestConnection "MYSQL1" dans la console VBA
□ 12. ORBIT > Recharger → vérifier que le catalogue se charge
```

---

## Flux Refresh All (EDB §4 + §6 + §7 + §8 + §9)

```
Clic "Refresh All"
    │
    ├─ CaptureContext()          → user, fichier, onglet, timestamp
    ├─ GetAllQueries()           → catalogue SharePoint (cache 5 min)
    ├─ GetAllTablesInWorkbook()  → scan des ListObjects du classeur
    ├─ Matching TargetSheet/Table
    │
    ├─ ConsolidateParameters()   → merge de tous les {{Params}}, dédoublonnage
    ├─ LoadLastConsolidatedParams() → recharger dernières valeurs
    ├─ frmConsolidatedParams     → UN SEUL dialog pour tout le classeur
    ├─ ValidateConsolidatedParams() → validation type + listes autorisées
    ├─ SaveConsolidatedParams()  → sauvegarder pour prochain refresh
    │
    └─ Pour chaque requête matchée :
        ├─ InjectConsolidatedParams() → affecter valeurs à chaque requête
        ├─ Si SQL  → SubstituteParameters() + ExecuteQuery() + RecordsetToArray()
        ├─ Si API  → CallApi() → ParseJsonResponse() / ParseCsvResponse()
        ├─ InjectDataTable()    → injection préservant formules et formats
        ├─ LogAuditEntry()      → SharePoint OrbitAuditLog
        └─ UpdateTraceSheet()   → onglet _ORBIT_TRACE du classeur
```

---

## Flux Signalement Dysqualité (EDB §15)

```
Clic "Signaler une dysqualité"
    │
    ├─ CaptureContext()     → contexte automatique complet
    ├─ FindBinding()        → requête liée au tableau actif
    ├─ LoadLastParams()     → paramètres utilisés lors du dernier refresh
    │
    ├─ frmDysqualite.Show   → criticité + description (seuls champs manuels)
    │
    ├─ CreateDQIssue()      → POST SharePoint DataQualityIssues
    │    └─ IncidentID généré : DQ-YYYYMMDD-HHMMSS-USER
    │
    └─ PrepareDQEmail()
         ├─ GetDQContacts() → destinataires depuis SP DataQualityContacts
         ├─ Corps prérempli : contexte + description + incidentId
         ├─ Pièce jointe optionnelle (configurable)
         └─ olMail.Display  → email ouvert, non envoyé (utilisateur finalise)
```
