# RFID-Tags

Application Windows avec interface graphique permettant d'utiliser le lecteur "ACR122U" pour lire ou écrire via des listes Excel ou CSV sur des puces RFID.

## Matériel

- **OS cible** : Windows (utilisation native de l'API Win32 pour l'interface graphique).
- **Lecteur RFID** : ACR122U.
- **Tags supportés** : Puces RFID (norme NTAG).
- **API Matérielle** : Protocole PC/SC via la librairie `winscard`.
- **Dépendances externes** : Microsoft Excel (requis en arrière-plan pour l'extraction via script VBScript).

## Build

Le projet est écrit en C natif. La compilation nécessite un compilateur supportant l'API Windows (comme GCC/MinGW ou MSVC) en incluant les bibliothèques système nécessaires.

```bash
# Exemple de compilation avec GCC (MinGW)
gcc rfid_donnees_gui.c -o rfid_tags.exe -lwinscard -lcomdlg32 -mwindows

# Lancer l'application
./rfid_tags.exe
```

## Flux de données
```bash
Fichier (CSV ou Excel) → RAM (CsvData)
                              ↓
                      Choix du Mode (Auto/Manuel/Lecture)
                              ↓
        WorkerThreadProc ↔ winscard (PC/SC) ↔ Lecteur ACR122U
                              ↓
                     Puce RFID (Requêtes APDU)

 ```
## Architecture Globale du Code (à lire avant de plonger dans le code) :

- Traitement de Données (CSV/Excel) : Le programme gère l'import natif de fichiers texte et fait appel à un script VBScript généré à la volée pour piloter de manière invisible (OLE Automation) l'extraction d'un fichier Excel en CSV.

- Protocole Matériel (PC/SC) : Toute la communication bas niveau vers le lecteur RFID (type ACR122U) transite via la librairie winscard. Le code envoie des requêtes APDU brutes pour lire et écrire par blocs de 4 octets (norme NTAG).

- Encodage NDEF : Il respecte la norme "NFC Forum Type 2" pour formater et décoder la charge utile (payload) du tag sous forme d'URI (lien compressé).

- Multithreading & Interface (Win32 API) : Pour empêcher l'interface graphique de figer pendant qu'on attend la présentation d'un badge, une fonction WorkerThreadProc tourne en arrière-plan et dialogue avec la fenêtre principale (WndProc) à l'aide de signaux de messagerie Windows (PostMessage).
