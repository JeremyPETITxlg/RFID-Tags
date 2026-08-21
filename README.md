# RFID-NFC-QR-Toolkit

Application Windows native (C / Win32) dédiée à l'automatisation de la traçabilité matérielle. L'outil permet d'encoder par lots des puces RFID/NFC (norme NTAG) et de générer simultanément des étiquettes QR Code au format PNG à partir de listes Excel ou CSV.

## Matériel & Compatibilité

- **OS cible** : Windows 10 / 11 (GUI native Win32 sans framework lourd).
- **Lecteurs RFID/NFC** : Tout lecteur compatible PC/SC (testé et optimisé sur ACR122U).
- **Puces supportées** : Famille NTAG (NTAG213, NTAG215, NTAG216, etc.) via commandes APDU standardisées.
- **Dépendances bureautiques** : Microsoft Excel (requis uniquement pour l'extraction automatique via VBScript).
## Fonctionnalités Clés
  - Mode Automatique : Déroulement séquentiel du fichier source avec reprise sur incident (persistance dans state.txt).

  - Mode Hybride : Possibilité d'encoder la puce RFID, de générer le QR Code, ou de réaliser les deux opérations simultanément.

  - Mode Clonage / Conversion : Lecture d'une puce RFID physique et génération instantanée du QR Code équivalent.

  - Étiquetage Visuel : Incrustation automatique de l'URL lisible par un opérateur sous l'image du QR Code généré.
## Bibliothèques & Dépendances

Le projet a été conçu pour être **100 % autonome et portable**, sans installateur lourd ni gestionnaire de paquets complexe :

- **`windows.h` & `commdlg.h`** : Gestion native de l'interface graphique Win32, du multithreading et des boîtes de dialogue de sélection de fichiers.
- **`winscard.h` (`winscard.lib`)** : API standard Windows PC/SC pour la communication bas niveau et l'envoi des trames APDU aux puces sans contact.
- **`qrcodegen.c` / `qrcodegen.h` (Projet Nayuki)** : Moteur mathématique C pur et léger pour le calcul de la matrice QR Code (ECC Low, masque automatique).
- **`stb_image_write.h` (Sean Barrett - Header-only)** : Compression et écriture directe du tableau de pixels en fichier `.png` standard sur le disque.
- **Police Bitmap 8x8 embarquée** : Matrice typographique intégrée directement dans le code pour dessiner l'étiquette texte sous le QR Code sans dépendre de polices système externes.

## Build

La compilation s'effectue directement via GCC (MinGW / w64devkit) :

```bash
# Compilation
gcc RFID-NFC-Tags.c qrcodegen.c -o RFID-NFC-Tags.exe -lwinscard -lcomdlg32 -mwindows

# Exécution
./RFID-NFC-Tags.exe
```

## Flux de données
```bash
Fichier Source (CSV / Excel via VBS)  ──┐
                                       ├─► RAM (CsvData) ──► Validation & Sélections
Saisie Manuelle / Scan Tag RFID       ──┘
                                                                │
                                   ┌────────────────────────────┴────────────────────────────┐
                                   ▼                                                         ▼
                      [ Export QR Code ]                                        [ Export RFID / NFC ]
                                   │                                                         │
                        qrcodegen + stb_image                                     Requêtes APDU (winscard)
                                   │                                                         │
                      Fichier PNG avec étiquette                                Puce physique NTAG 
```
