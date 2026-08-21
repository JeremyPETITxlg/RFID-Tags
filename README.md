# RFID-Tags
Utilisation du lecteur "ACR122U" pour lire ou écrire via des listes Excel ou CSV sur des puces RFID.

Architecture Globale du Code (à lire avant de plonger dans le code) :

- Traitement de Données (CSV/Excel) : Le programme gère l'import natif de fichiers texte et fait appel à un script VBScript généré à la volée pour piloter de manière invisible (OLE Automation) l'extraction d'un fichier Excel en CSV.

- Protocole Matériel (PC/SC) : Toute la communication bas niveau vers le lecteur RFID (type ACR122U) transite via la librairie winscard. Le code envoie des requêtes APDU brutes pour lire et écrire par blocs de 4 octets (norme NTAG).

- Encodage NDEF : Il respecte la norme "NFC Forum Type 2" pour formater et décoder la charge utile (payload) du tag sous forme d'URI (lien compressé).

- Multithreading & Interface (Win32 API) : Pour empêcher l'interface graphique de figer pendant qu'on attend la présentation d'un badge, une fonction WorkerThreadProc tourne en arrière-plan et dialogue avec la fenêtre principale (WndProc) à l'aide de signaux de messagerie Windows (PostMessage).