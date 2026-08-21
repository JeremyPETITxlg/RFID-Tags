/* ============================================================================
 * RFID_NFC_Tags.c
 *
 * Version "interface graphique" avec importation Excel directe
 * et option de limite de recupération de donnees.
 * ==========================================================================*/

/* --- INCLUSIONS DES BIBLIOTHÈQUES --- */
#include <windows.h>    // API Windows de base pour l'interface graphique (GUI)
#include <winscard.h>   // API PC/SC pour la communication avec le lecteur de carte à puce/RFID
#include <commdlg.h>    // Boîtes de dialogue communes (pour parcourir les fichiers)
#include <stdio.h>      // Entrées/sorties standards (fichiers, printf, etc.)
#include <stdlib.h>     // Fonctions utilitaires standards (allocations dynamiques, conversions)
#include <string.h>     // Manipulation des chaînes de caractères
#include <stdarg.h>     // Gestion des arguments variables (utilisé pour la fonction de log)

/* --- NOUVELLES BIBLIOTHÈQUES QRcode --- */
#include "qrcodegen.h"

#define STB_IMAGE_WRITE_IMPLEMENTATION
#include "stb_image_write.h"

/* --- DIRECTIVES DE COMPILATION --- */
// Indique au linker d'inclure automatiquement ces bibliothèques pour éviter des erreurs de compilation
#pragma comment(lib, "winscard.lib") 
#pragma comment(lib, "comdlg32.lib")

/* --- DÉFINITION DES CONSTANTES ET PARAMÈTRES --- */
#define STATE_FILE          "state.txt" // Fichier de sauvegarde de l'état (donnée courante, chemin CSV)
#define NTAG_PAGE_SIZE       4          // Taille d'une page NTAG (4 octets)
#define NTAG_USER_START      4          // Première page utilisateur de la mémoire NTAG
#define NTAG_USER_PAGES      36         // Nombre de pages utilisateur (ici adapté pour une puce type NTAG213)
#define POLL_MS              300        // Temps d'attente (en ms) entre chaque vérification du lecteur matériel

/* Identifiants (IDs) des contrôles de l'interface graphique (Boutons, Textes, etc.) */
#define IDC_BROWSE          101
#define IDC_CSV_LABEL       102
#define IDC_donnee_LABEL    103
#define IDC_RADIO_AUTO      104
#define IDC_RADIO_MANUAL    105
#define IDC_RADIO_READ      106
#define IDC_GOTO_EDIT       107
#define IDC_GOTO_BTN        108
#define IDC_RESET_BTN       109
#define IDC_MANUAL_EDIT     110
#define IDC_STARTSTOP_BTN   111
#define IDC_LOG_EDIT        112
#define IDC_IMPORT_EXCEL    113 
#define IDC_CHECK_RFID      114
#define IDC_CHECK_QR        115

/* Messages Windows personnalisés pour la communication inter-threads (Thread Travailleur -> Thread GUI) */
#define WM_APP_LOG   (WM_APP + 1)   // Message pour ajouter une ligne dans l'historique visuel (log)
#define WM_APP_donnee  (WM_APP + 2) // Message pour mettre à jour l'affichage de l'index de donnée courante
#define WM_APP_DONE  (WM_APP + 3)   // Message indiquant l'arrêt du processus de lecture/écriture

/* --- POLICE D'ÉCRITURE PIXEL ART 8x8 (ASCII 32 à 127) --- */
static const uint8_t font8x8[96][8] = {
    {0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00},{0x18,0x3C,0x3C,0x18,0x18,0x00,0x18,0x00},{0x66,0x66,0x66,0x00,0x00,0x00,0x00,0x00},{0x6C,0x6C,0xFE,0x6C,0xFE,0x6C,0x6C,0x00},
    {0x18,0x3E,0x60,0x3C,0x06,0x7C,0x18,0x00},{0x00,0xC6,0xCC,0x18,0x30,0x66,0xC6,0x00},{0x38,0x6C,0x6C,0x38,0x6D,0x66,0x3B,0x00},{0x18,0x18,0x30,0x00,0x00,0x00,0x00,0x00},
    {0x0C,0x18,0x30,0x30,0x30,0x18,0x0C,0x00},{0x30,0x18,0x0C,0x0C,0x0C,0x18,0x30,0x00},{0x00,0x66,0x3C,0xFF,0x3C,0x66,0x00,0x00},{0x00,0x18,0x18,0x7E,0x18,0x18,0x00,0x00},
    {0x00,0x00,0x00,0x00,0x00,0x18,0x18,0x30},{0x00,0x00,0x00,0x7E,0x00,0x00,0x00,0x00},{0x00,0x00,0x00,0x00,0x00,0x18,0x18,0x00},{0x06,0x0C,0x18,0x30,0x60,0xC0,0x80,0x00},
    {0x3C,0x66,0x66,0x6E,0x76,0x66,0x3C,0x00},{0x18,0x38,0x18,0x18,0x18,0x18,0x7E,0x00},{0x3C,0x66,0x06,0x0C,0x18,0x30,0x7E,0x00},{0x3C,0x66,0x06,0x1C,0x06,0x66,0x3C,0x00},
    {0x1C,0x3C,0x6C,0xCC,0xFE,0x0C,0x1E,0x00},{0x7E,0x60,0x7C,0x06,0x06,0x66,0x3C,0x00},{0x3C,0x66,0x60,0x7C,0x66,0x66,0x3C,0x00},{0x7E,0x06,0x0C,0x18,0x30,0x30,0x30,0x00},
    {0x3C,0x66,0x66,0x3C,0x66,0x66,0x3C,0x00},{0x3C,0x66,0x66,0x3E,0x06,0x66,0x3C,0x00},{0x00,0x18,0x18,0x00,0x00,0x18,0x18,0x00},{0x00,0x18,0x18,0x00,0x00,0x18,0x18,0x30},
    {0x06,0x0C,0x18,0x30,0x18,0x0C,0x06,0x00},{0x00,0x00,0x7E,0x00,0x7E,0x00,0x00,0x00},{0x60,0x30,0x18,0x0C,0x18,0x30,0x60,0x00},{0x3C,0x66,0x06,0x1C,0x18,0x00,0x18,0x00},
    {0x3C,0x66,0x6E,0x6E,0x60,0x66,0x3C,0x00},{0x18,0x3C,0x66,0x66,0x7E,0x66,0x66,0x00},{0x7C,0x66,0x66,0x7C,0x66,0x66,0x7C,0x00},{0x3C,0x66,0x60,0x60,0x60,0x66,0x3C,0x00},
    {0x78,0x6C,0x66,0x66,0x66,0x6C,0x78,0x00},{0x7E,0x60,0x60,0x78,0x60,0x60,0x7E,0x00},{0x7E,0x60,0x60,0x78,0x60,0x60,0x60,0x00},{0x3C,0x66,0x60,0x6E,0x66,0x66,0x3E,0x00},
    {0x66,0x66,0x66,0x7E,0x66,0x66,0x66,0x00},{0x3C,0x18,0x18,0x18,0x18,0x18,0x3C,0x00},{0x06,0x06,0x06,0x06,0x06,0x66,0x3C,0x00},{0x66,0x6C,0x78,0x70,0x78,0x6C,0x66,0x00},
    {0x60,0x60,0x60,0x60,0x60,0x60,0x7E,0x00},{0x63,0x77,0x7F,0x6B,0x63,0x63,0x63,0x00},{0x66,0x76,0x7E,0x7E,0x6E,0x66,0x66,0x00},{0x3C,0x66,0x66,0x66,0x66,0x66,0x3C,0x00},
    {0x7C,0x66,0x66,0x7C,0x60,0x60,0x60,0x00},{0x3C,0x66,0x66,0x66,0x6A,0x6C,0x36,0x00},{0x7C,0x66,0x66,0x7C,0x6C,0x66,0x66,0x00},{0x3C,0x66,0x60,0x3C,0x06,0x66,0x3C,0x00},
    {0x7E,0x18,0x18,0x18,0x18,0x18,0x18,0x00},{0x66,0x66,0x66,0x66,0x66,0x66,0x3C,0x00},{0x66,0x66,0x66,0x66,0x66,0x3C,0x18,0x00},{0x63,0x63,0x63,0x6B,0x7F,0x77,0x63,0x00},
    {0x66,0x66,0x3C,0x18,0x3C,0x66,0x66,0x00},{0x66,0x66,0x66,0x3C,0x18,0x18,0x18,0x00},{0x7E,0x06,0x0C,0x18,0x30,0x60,0x7E,0x00},{0x3C,0x30,0x30,0x30,0x30,0x30,0x3C,0x00},
    {0x80,0xC0,0x60,0x30,0x18,0x0C,0x06,0x00},{0x3C,0x0C,0x0C,0x0C,0x0C,0x0C,0x3C,0x00},{0x18,0x3C,0x66,0x00,0x00,0x00,0x00,0x00},{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xFF},
    {0x30,0x18,0x0C,0x00,0x00,0x00,0x00,0x00},{0x00,0x00,0x3C,0x06,0x3E,0x66,0x3E,0x00},{0x60,0x60,0x7C,0x66,0x66,0x66,0x7C,0x00},{0x00,0x00,0x3C,0x60,0x60,0x60,0x3C,0x00},
    {0x06,0x06,0x3E,0x66,0x66,0x66,0x3E,0x00},{0x00,0x00,0x3C,0x66,0x7E,0x60,0x3C,0x00},{0x1C,0x30,0x7C,0x30,0x30,0x30,0x30,0x00},{0x00,0x00,0x3E,0x66,0x66,0x3E,0x06,0x3C},
    {0x60,0x60,0x7C,0x66,0x66,0x66,0x66,0x00},{0x18,0x00,0x38,0x18,0x18,0x18,0x3C,0x00},{0x0C,0x00,0x1C,0x0C,0x0C,0x0C,0x0C,0x38},{0x60,0x60,0x66,0x6C,0x78,0x6C,0x66,0x00},
    {0x38,0x18,0x18,0x18,0x18,0x18,0x3C,0x00},{0x00,0x00,0x6C,0x7E,0x54,0x54,0x54,0x00},{0x00,0x00,0x7C,0x66,0x66,0x66,0x66,0x00},{0x00,0x00,0x3C,0x66,0x66,0x66,0x3C,0x00},
    {0x00,0x00,0x7C,0x66,0x66,0x7C,0x60,0x60},{0x00,0x00,0x3E,0x66,0x66,0x3E,0x06,0x06},{0x00,0x00,0x7C,0x60,0x60,0x60,0x60,0x00},{0x00,0x00,0x3E,0x60,0x3C,0x06,0x7C,0x00},
    {0x30,0x30,0x7C,0x30,0x30,0x34,0x18,0x00},{0x00,0x00,0x66,0x66,0x66,0x66,0x3E,0x00},{0x00,0x00,0x66,0x66,0x66,0x3C,0x18,0x00},{0x00,0x00,0x63,0x6B,0x7F,0x3E,0x36,0x00},
    {0x00,0x00,0x66,0x3C,0x18,0x3C,0x66,0x00},{0x00,0x00,0x66,0x66,0x66,0x3E,0x06,0x3C},{0x00,0x00,0x7E,0x0C,0x18,0x30,0x7E,0x00},{0x0E,0x18,0x18,0x70,0x18,0x18,0x0E,0x00},
    {0x18,0x18,0x18,0x00,0x18,0x18,0x18,0x00},{0x70,0x18,0x18,0x0E,0x18,0x18,0x70,0x00},{0x3B,0x6E,0x00,0x00,0x00,0x00,0x00,0x00}
};

/* Enumération définissant les 3 modes de fonctionnement de l'application */
typedef enum { MODE_AUTO, MODE_MANUAL, MODE_READ } AppMode;

/* ---------------------------------------------------------------------- */
/* Donnees CSV : Gestion du chargement et de la mémoire                   */
/* ---------------------------------------------------------------------- */

// Structure représentant les données du fichier CSV chargé en RAM
typedef struct {
    char **lines; // Tableau dynamique de chaînes de caractères (lignes)
    int count;    // Nombre total de lignes récupérées
} CsvData;

/**
 * Charge un fichier CSV en mémoire.
 * Gère l'encodage (retrait du BOM UTF-8) et stocke chaque ligne dynamiquement.
 */
static CsvData load_csv(const char *path) {
    CsvData csv = { NULL, 0 };
    FILE *f = fopen(path, "rb"); // Lecture binaire pour sécuriser la gestion des sauts de ligne
    if (!f) return csv;

    // Gestion du BOM (Byte Order Mark) UTF-8 éventuel présent en début de fichier
    unsigned char bom[3];
    size_t read = fread(bom, 1, 3, f);
    if (!(read == 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)) {
        // Si aucun BOM détecté, on remet le curseur de lecture au tout début du fichier
        fseek(f, 0, SEEK_SET);
    }

    int capacity = 1024; // Capacité initiale du tableau (évite les réallocations trop fréquentes)
    csv.lines = malloc(capacity * sizeof(char *));

    char buffer[1024];
    while (fgets(buffer, sizeof(buffer), f)) {
        // Nettoyage systématique des sauts de ligne (\r, \n) à la fin de la ligne lue
        size_t len = strlen(buffer);
        while (len > 0 && (buffer[len - 1] == '\n' || buffer[len - 1] == '\r')) {
            buffer[--len] = '\0';
        }
        if (len == 0) continue; // Ignore les lignes vides

        // Agrandissement du tableau si la limite est atteinte
        if (csv.count >= capacity) {
            capacity *= 2;
            csv.lines = realloc(csv.lines, capacity * sizeof(char *));
        }
        csv.lines[csv.count] = _strdup(buffer); // Duplique la ligne brute et l'ajoute
        csv.count++;
    }
    fclose(f);
    return csv;
}

/**
 * Libère proprement la mémoire allouée pour les données CSV afin d'éviter les fuites (Memory Leaks).
 */
static void free_csv(CsvData *csv) {
    for (int i = 0; i < csv->count; i++) free(csv->lines[i]);
    free(csv->lines);
    csv->lines = NULL;
    csv->count = 0;
}

/**
 * Récupère la ligne de donnée correspondant à l'ID demandé.
 * (Attention : donnee_id est indexé à partir de 1, alors que le code C (tableau) commence à 0).
 */
static const char *get_donnee_data(CsvData *csv, int donnee_id) {
    if (donnee_id < 1 || donnee_id > csv->count) return NULL;
    return csv->lines[donnee_id - 1];
}

/* ---------------------------------------------------------------------- */
/* Importation Excel via VBScript (avec limite optionnelle)               */
/* ---------------------------------------------------------------------- */
/**
 * Génère dynamiquement un script VBScript pour ouvrir Excel en arrière-plan,
 * extraire les données depuis une case donnée jusqu'à la première case vide,
 * et les sauvegarder silencieusement dans un fichier temporaire .csv.
 */
static int import_excel_via_vbs_gui(const char *excel_path, const char *csv_path) {
    char abs_excel[MAX_PATH];
    char abs_csv[MAX_PATH];
    // Convertit les chemins relatifs en chemins absolus pour éviter les erreurs d'accès de VBS
    GetFullPathNameA(excel_path, MAX_PATH, abs_excel, NULL);
    GetFullPathNameA(csv_path, MAX_PATH, abs_csv, NULL);

    FILE *vbs = fopen("extract.vbs", "w");
    if (!vbs) return 0;

    // Début de l'écriture du code source du script VBS
    fprintf(vbs, "start_cell = InputBox(\"Entrez la case de depart (ex: A1) :\", \"Importation Excel\", \"A1\")\n");
    fprintf(vbs, "If start_cell = \"\" Then WScript.Quit 2\n");
    
    /* Demande de limite avec gestion du vide / erreur */
    fprintf(vbs, "max_rows_str = InputBox(\"Combien de données maximum a récupérer ?\" & vbCrLf & \"(Laissez vide pour tout récupérer jusqu'à la première case vide)\", \"Limite (Optionnel)\", \"\")\n");
    fprintf(vbs, "max_rows = -1\n"); /* -1 signifie aucune limite */
    fprintf(vbs, "If max_rows_str <> \"\" Then\n");
    fprintf(vbs, "  If IsNumeric(max_rows_str) Then\n");
    fprintf(vbs, "    max_rows = CInt(max_rows_str)\n");
    fprintf(vbs, "    If max_rows <= 0 Then\n");
    fprintf(vbs, "      MsgBox \"La limite doit être supérieure a 0. Aucune limite ne sera appliquée.\", 48, \"Attention\"\n");
    fprintf(vbs, "      max_rows = -1\n");
    fprintf(vbs, "    End If\n");
    fprintf(vbs, "  Else\n");
    fprintf(vbs, "    MsgBox \"Valeur invalide. Aucune limite ne sera appliquée.\", 48, \"Attention\"\n");
    fprintf(vbs, "  End If\n");
    fprintf(vbs, "End If\n");

    /* Traitement Excel via OLE Automation (ComObject) */
    fprintf(vbs, "Set fso = CreateObject(\"Scripting.FileSystemObject\")\n");
    fprintf(vbs, "Set out = fso.CreateTextFile(\"%s\", True)\n", abs_csv);
    fprintf(vbs, "Set xl = CreateObject(\"Excel.Application\")\n");
    fprintf(vbs, "xl.Visible = False\n");      // Rend la fenêtre Excel invisible
    fprintf(vbs, "xl.DisplayAlerts = False\n"); // Désactive les alertes/pop-ups
    fprintf(vbs, "On Error Resume Next\n");
    fprintf(vbs, "Set wb = xl.Workbooks.Open(\"%s\")\n", abs_excel);
    fprintf(vbs, "If Err.Number <> 0 Then\n");
    fprintf(vbs, "  MsgBox \"Erreur : Impossible d'ouvrir le fichier Excel.\", 16, \"Erreur\"\n");
    fprintf(vbs, "  WScript.Quit 1\n");
    fprintf(vbs, "End If\n");
    fprintf(vbs, "Set ws = wb.Sheets(1)\n");
    fprintf(vbs, "Set cell = ws.Range(start_cell)\n");
    fprintf(vbs, "If Err.Number <> 0 Then\n");
    fprintf(vbs, "  MsgBox \"Erreur : Case de départ invalide.\", 16, \"Erreur\"\n");
    fprintf(vbs, "  wb.Close False\n");
    fprintf(vbs, "  xl.Quit\n");
    fprintf(vbs, "  WScript.Quit 1\n");
    fprintf(vbs, "End If\n");
    
    /* Boucle de récupération de l'Excel cellule par cellule vers le texte */
    fprintf(vbs, "row = cell.Row\n");
    fprintf(vbs, "col = cell.Column\n");
    fprintf(vbs, "count = 0\n");
    
    fprintf(vbs, "Do While ws.Cells(row, col).Value <> \"\"\n");
    fprintf(vbs, "  If max_rows <> -1 And count >= max_rows Then Exit Do\n"); /* Arrêt si limite atteinte */
    fprintf(vbs, "  out.WriteLine ws.Cells(row, col).Value\n");
    fprintf(vbs, "  row = row + 1\n");
    fprintf(vbs, "  count = count + 1\n");
    fprintf(vbs, "Loop\n");
    
    // Fermeture propre des objets COM pour éviter un processus Excel fantôme
    fprintf(vbs, "wb.Close False\n");
    fprintf(vbs, "xl.Quit\n");
    fprintf(vbs, "out.Close\n");
    fprintf(vbs, "MsgBox \"Extraction terminée : \" & count & \" lignes récupérées.\", 64, \"Succès\"\n");
    fprintf(vbs, "WScript.Quit 0\n");
    fclose(vbs);

    // Exécution du script VBS de manière synchrone via une commande terminal classique
    int ret = system("wscript.exe extract.vbs");
    remove("extract.vbs"); // Nettoyage du fichier script temporaire
    return (ret == 0); 
}


/* ---------------------------------------------------------------------- */
/* Etat persistant (Sauvegarde de l'avancement localement)                */
/* ---------------------------------------------------------------------- */
static char g_csvPath[MAX_PATH] = { 0 }; // Conserve en mémoire le chemin du fichier chargé

/**
 * Lit le fichier de configuration au lancement.
 * Permet de reprendre le travail exactement là où il a été arrêté.
 */
static int load_state(void) {
    FILE *f = fopen(STATE_FILE, "r");
    if (!f) return 1; // Retourne par défaut la ligne 1 si aucun historique n'existe

    char line[MAX_PATH + 32];
    int donnee = 1;
    while (fgets(line, sizeof(line), f)) {
        line[strcspn(line, "\r\n")] = '\0'; // Supprime le saut de ligne à la fin
        // Extrait les valeurs basées sur le préfixe identifié
        if (strncmp(line, "current_donnee=", 15) == 0) {
            donnee = atoi(line + 15);
        } else if (strncmp(line, "csv_path=", 9) == 0) {
            strncpy(g_csvPath, line + 9, sizeof(g_csvPath) - 1);
        }
    }
    fclose(f);
    if (donnee < 1) donnee = 1; // Mesure de sécurité
    return donnee;
}

/**
 * Écrit l'index de progression courante et le chemin du fichier dans le TXT.
 */
static void save_state(int donnee) {
    FILE *f = fopen(STATE_FILE, "w");
    if (!f) return;
    fprintf(f, "current_donnee=%d\n", donnee);
    if (g_csvPath[0] != '\0') fprintf(f, "csv_path=%s\n", g_csvPath);
    fclose(f);
}

/* ---------------------------------------------------------------------- */
/* PC/SC bas niveau : Gestion de la communication Lecteur <-> Puce RFID   */
/* ---------------------------------------------------------------------- */
static SCARDCONTEXT g_hContext; // Handle vers le gestionnaire de ressources de carte à puce
static char g_readerName[256];  // Nom du lecteur détecté par Windows (ex: "ACS ACR122 0")
static BOOL g_readerReady = FALSE;

/**
 * Initialise le contexte API et recherche le premier lecteur disponible branché au PC.
 */
static BOOL init_pcsc(void) {
    LONG rv = SCardEstablishContext(SCARD_SCOPE_SYSTEM, NULL, NULL, &g_hContext);
    if (rv != SCARD_S_SUCCESS) return FALSE;

    char readers[1024];
    DWORD readersLen = sizeof(readers);
    rv = SCardListReaders(g_hContext, NULL, readers, &readersLen);
    if (rv != SCARD_S_SUCCESS) return FALSE;

    strncpy(g_readerName, readers, sizeof(g_readerName) - 1); // Sélectionne le lecteur primaire
    return TRUE;
}

/**
 * Écrit physiquement exactement 4 octets sur une page mémoire précise du tag.
 * Utilise la commande APDU universelle de modification : 0xFF 0xD6 0x00 <Page> 0x04 <Datas>.
 */
static int write_page(SCARDHANDLE hCard, BYTE page, const BYTE data4[4]) {
    BYTE apdu[9] = { 0xFF, 0xD6, 0x00, page, 0x04, 0, 0, 0, 0 };
    memcpy(apdu + 5, data4, 4);

    BYTE response[16];
    DWORD responseLen = sizeof(response);
    LONG rv = SCardTransmit(hCard, SCARD_PCI_T1, apdu, sizeof(apdu),
                             NULL, response, &responseLen);
    if (rv != SCARD_S_SUCCESS) return 0;
    // La réponse d'une puce saine doit se terminer par 0x90 0x00
    return (responseLen >= 2 && response[responseLen - 2] == 0x90 &&
            response[responseLen - 1] == 0x00);
}

/**
 * Lit exactement 4 octets (une page) depuis le tag.
 * Commande APDU de lecture brute : 0xFF 0xB0 0x00 <Page> 0x04.
 */
static int read_page(SCARDHANDLE hCard, BYTE page, BYTE out4[4]) {
    BYTE apdu[5] = { 0xFF, 0xB0, 0x00, page, 0x04 };
    BYTE response[16];
    DWORD responseLen = sizeof(response);
    LONG rv = SCardTransmit(hCard, SCARD_PCI_T1, apdu, sizeof(apdu),
                             NULL, response, &responseLen);
    if (rv != SCARD_S_SUCCESS) return 0;
    if (responseLen < 6) return 0; // Sécurité de longueur (les 4 octets de data + 2 octets de statut)
    if (response[responseLen - 2] != 0x90 || response[responseLen - 1] != 0x00) return 0;
    memcpy(out4, response, 4); // Capture l'information retournée
    return 1;
}

/**
 * Fonction de blocage : Polle (interroge) en boucle le lecteur jusqu'à l'arrivée d'une puce.
 * Permet l'interruption logicielle externe via l'écoute de "stopFlag".
 */
static SCARDHANDLE wait_for_card_i(volatile BOOL *stopFlag) {
    SCARD_READERSTATE state;
    memset(&state, 0, sizeof(state));
    state.szReader = g_readerName;
    state.dwCurrentState = SCARD_STATE_EMPTY;

    // Boucle d'attente
    while (!*stopFlag) {
        LONG rv = SCardGetStatusChange(g_hContext, POLL_MS, &state, 1);
        if (rv == (LONG)SCARD_E_TIMEOUT) continue; // Si timeout normal atteint, on reboucle
        if (rv != SCARD_S_SUCCESS) return 0;
        if (state.dwEventState & SCARD_STATE_PRESENT) break; // Un tag a été présenté !
        state.dwCurrentState = state.dwEventState;
    }
    if (*stopFlag) return 0;

    // Négocie la connexion avec le protocole de la carte (T=0 ou T=1)
    SCARDHANDLE hCard;
    DWORD activeProtocol;
    LONG rv = SCardConnect(g_hContext, g_readerName, SCARD_SHARE_SHARED,
                            SCARD_PROTOCOL_T0 | SCARD_PROTOCOL_T1,
                            &hCard, &activeProtocol);
    if (rv != SCARD_S_SUCCESS) return 0;
    return hCard;
}

/**
 * Fonction de blocage inversée : Attend que l'utilisateur dépose la puce physiquement en-dehors de la portée du lecteur.
 */
static void wait_for_card_removal_i(volatile BOOL *stopFlag) {
    SCARD_READERSTATE state;
    memset(&state, 0, sizeof(state));
    state.szReader = g_readerName;
    state.dwCurrentState = SCARD_STATE_PRESENT;

    while (!*stopFlag) {
        LONG rv = SCardGetStatusChange(g_hContext, POLL_MS, &state, 1);
        if (rv == (LONG)SCARD_E_TIMEOUT) continue;
        if (rv != SCARD_S_SUCCESS) return;
        if (state.dwEventState & SCARD_STATE_EMPTY) return; // Tag retiré
        state.dwCurrentState = state.dwEventState;
    }
}

/* ---------------------------------------------------------------------- */
/* Encodage / decodage NDEF URI (Protocole NFC Forum)                     */
/* ---------------------------------------------------------------------- */
/**
 * Traduit le code de compression URI (0x01, 0x02, etc.) en sa véritable racine texte.
 */
static const char *uri_prefix_for_code(BYTE code) {
    switch (code) {
        case 0x01: return "http://www.";
        case 0x02: return "https://www.";
        case 0x03: return "http://";
        case 0x04: return "https://";
        default:   return ""; // Cas de chaîne brute ou protocole non couvert
    }
}

/**
 * Analyse une URL texte pour retirer son préfixe et le remplacer par le code URI
 * afin de minimiser drastiquement l'espace requis sur la mémoire très faible du Tag.
 */
static BYTE uri_code_for_url(const char *url, const char **payloadStart) {
    if (strncmp(url, "https://www.", 12) == 0) { *payloadStart = url + 12; return 0x02; }
    if (strncmp(url, "http://www.", 11) == 0)  { *payloadStart = url + 11; return 0x01; }
    if (strncmp(url, "https://", 8) == 0)      { *payloadStart = url + 8;  return 0x04; }
    if (strncmp(url, "http://", 7) == 0)       { *payloadStart = url + 7;  return 0x03; }
    *payloadStart = url;
    return 0x00;
}

/**
 * Compile une trame de type NDEF (NFC Data Exchange Format).
 * Structure TLV complète requise par les smartphones pour lire la puce nativement.
 */
static size_t build_ndef_uri(const char *url, BYTE *out, size_t outCapacity) {
    const char *payloadStr;
    BYTE code = uri_code_for_url(url, &payloadStr);
    size_t strLen = strlen(payloadStr);
    size_t payloadLen = 1 + strLen; // 1 octet pour le code d'URI + texte
    if (payloadLen > 255) return 0; // Restriction : ne traite que le Short Record (SR=1)

    size_t recordLen = 4 + payloadLen; // Header Record (4) + Payload Data
    size_t needed = 2 + recordLen + 1; // TLV NDEF (T=0x03, L=...) + Record + Terminator(0xFE)
    if (needed > outCapacity) return 0;

    size_t pos = 0;
    // --- Phase 1: TLV Block
    out[pos++] = 0x03;               // T = Message NDEF
    out[pos++] = (BYTE)recordLen;    // L = Longueur
    // --- Phase 2: NDEF Record Header
    out[pos++] = 0xD1;               // MB=1, ME=1, CF=0, SR=1, IL=0, TNF=0x01 (Type Well Known)
    out[pos++] = 0x01;               // Type Length = 1
    out[pos++] = (BYTE)payloadLen;   // Payload Length
    out[pos++] = 'U';                // Record Type = 'U' (Indique que c'est une URI)
    out[pos++] = code;               // Préfixe URI Hexa
    // --- Phase 3: Copie Payload
    memcpy(out + pos, payloadStr, strLen);
    pos += strLen;
    // --- Phase 4: Termination du message
    out[pos++] = 0xFE;               // Fin stricte de l'information (TLV Terminator)
    return pos;
}

/**
 * Parcourt un dump mémoire brut de puce et l'inspecte pour localiser un message de type "URI NDEF" valide.
 */
static int decode_ndef_uri(const BYTE *buffer, size_t total, char *outUrl, size_t outUrlCap) {
    if (total < 3 || buffer[0] != 0x03) return 0; // Condition NDEF non remplie
    size_t recordLen = buffer[1];
    size_t p = 2;
    if (recordLen < 3 || p + recordLen > total) return 0; // Longueur anormale
    if (buffer[p + 3] != 'U') return 0; // Ce n'est pas un message Web/URI (Type U)
    size_t payloadLen = buffer[p + 2];
    BYTE code = buffer[p + 4];
    size_t strLen = payloadLen - 1;
    if (p + 4 + 1 + strLen > total) return 0;

    const char *prefix = uri_prefix_for_code(code);
    snprintf(outUrl, outUrlCap, "%s%.*s", prefix, (int)strLen, (const char *)(buffer + p + 5)); // Reconstruction
    return 1;
}

/**
 * Relit physiquement toutes les pages qui viennent d'être encodées pour s'assurer que la gravure a fonctionné.
 */
static int read_back_url(SCARDHANDLE hCard, int pageCount, char *outUrl, size_t outUrlCap) {
    BYTE buffer[1024];
    size_t total = 0;
    for (int i = 0; i < pageCount; i++) {
        BYTE page4[4];
        if (!read_page(hCard, (BYTE)(NTAG_USER_START + i), page4)) return 0;
        memcpy(buffer + total, page4, 4);
        total += 4;
    }
    return decode_ndef_uri(buffer, total, outUrl, outUrlCap);
}

/**
 * Permet de scanner des tags (Mode de Lecture). 
 * La fonction aspire le contenu mémoire jusqu'à avoir reconstitué l'architecture complète du lien NDEF.
 */
static int read_ntag_auto(SCARDHANDLE hCard, char *outUrl, size_t outUrlCap) {
    BYTE buffer[600];
    size_t total = 0;
    int maxPages = NTAG_USER_PAGES + 4;

    for (int i = 0; i < maxPages && total + 4 <= sizeof(buffer); i++) {
        BYTE page4[4];
        if (!read_page(hCard, (BYTE)(NTAG_USER_START + i), page4)) break;
        memcpy(buffer + total, page4, 4);
        total += 4;

        if (total >= 4 && buffer[0] != 0x03) return 0; // Ce n'est pas un badge formaté NDEF
        if (total >= 6) {
            size_t recordLen = buffer[1];
            size_t needed = 2 + recordLen + 1; // Taille prévue selon le paramètre T L V
            if (total >= needed) {
                // Dès que les données lues couvrent l'intégralité du NDEF déclaré, on décode et on sort.
                return decode_ndef_uri(buffer, total, outUrl, outUrlCap);
            }
        }
    }
    return 0;
}

/**
 * Fonction maîtresse. Assemble l'ensemble des tâches d'un cycle matériel.
 * Formatage du message -> Alignement binaire -> Écriture page par page -> Vérification par relecture.
 */
static int write_ntag(SCARDHANDLE hCard, const char *url) {
    BYTE msg[512];
    size_t msgLen = build_ndef_uri(url, msg, sizeof(msg));
    if (msgLen == 0) return 0;
    if (msgLen > (size_t)(NTAG_USER_PAGES * NTAG_PAGE_SIZE)) return 0;

    // Padding (Remplissage) : On force la taille du tableau en tant que multiple de 4 octets.
    size_t paddedLen = ((msgLen + 3) / 4) * 4;
    BYTE *buffer = calloc(paddedLen, 1);
    memcpy(buffer, msg, msgLen);

    // Boucle de flashage
    int pageCount = (int)(paddedLen / NTAG_PAGE_SIZE);
    for (int i = 0; i < pageCount; i++) {
        if (!write_page(hCard, (BYTE)(NTAG_USER_START + i), buffer + (i * NTAG_PAGE_SIZE))) {
            free(buffer);
            return 0;
        }
    }
    free(buffer);

    // Contrôle qualité de fin d'opération
    char reread[300];
    if (!read_back_url(hCard, pageCount, reread, sizeof(reread))) return 0;
    if (strcmp(reread, url) != 0) return 0;
    return 1;
}

/* ---------------------------------------------------------------------- */
/* Variables globales de l'interface / du thread de travail               */
/* ---------------------------------------------------------------------- */
// Handles (pointeurs natifs Windows) des composants graphiques
static HINSTANCE g_hInst;
static HWND g_hWndMain;
static HWND g_hCsvLabel, g_hdonneeLabel;
static HWND g_hRadioAuto, g_hRadioManual, g_hRadioRead;
static HWND g_hCheckRfid, g_hCheckQr;
static HWND g_hGotoEdit, g_hGotoBtn, g_hResetBtn, g_hBrowseBtn, g_hImportExcelBtn;
static HWND g_hManualEdit, g_hStartStopBtn, g_hLogEdit;
static HFONT g_hFont;

static CsvData g_csv = { NULL, 0 }; // Base de données en mémoire vive
static int g_currentdonnee = 1;     // Index actif

// Drapeaux (Flags) asynchrones. 'volatile' oblige le compilateur à relire la variable car elle change hors du flux direct.
static volatile BOOL g_isRunning = FALSE;
static volatile BOOL g_stopRequested = FALSE;
static AppMode g_mode;
static char g_manualValueSnapshot[512];

/* ---------------------------------------------------------------------- */
/* Journal (Log Consolide)                                                */
/* ---------------------------------------------------------------------- */
/* ---------------------------------------------------------------------- */
/* Génération d'image QR Code (PNG)                                       */
/* ---------------------------------------------------------------------- */
/* Dessine un caractère unique sur l'image */
static void draw_char(uint8_t *img, int img_width, int x, int y, char c, int scale) {
    if (c < 32 || c > 127) c = '?'; // Si caractère inconnu, on affiche '?'
    int char_idx = c - 32;          // Calcule l'index dans notre tableau font8x8
    
    for (int row = 0; row < 8; row++) {
        uint8_t line = font8x8[char_idx][row];
        for (int col = 0; col < 8; col++) {
            if (line & (1 << (7 - col))) { // Opération bit-à-bit : Si le pixel doit être noir
                for (int dy = 0; dy < scale; dy++) {
                    for (int dx = 0; dx < scale; dx++) {
                        int px = x + col * scale + dx;
                        int py = y + row * scale + dy;
                        if (px < img_width) { // Sécurité pour ne pas déborder de l'image
                            int idx = (py * img_width + px) * 3;
                            img[idx + 0] = 0x00; // Rouge : 0
                            img[idx + 1] = 0x00; // Vert : 0
                            img[idx + 2] = 0x00; // Bleu : 0
                        }
                    }
                }
            }
        }
    }
}

/* Dessine une phrase complète avec retour à la ligne automatique */
static void draw_string(uint8_t *img, int img_width, int img_height, int start_x, int start_y, const char *text, int scale) {
    int x = start_x;
    int y = start_y;
    while (*text) {
        draw_char(img, img_width, x, y, *text, scale);
        x += 8 * scale; // Avance le curseur pour la prochaine lettre
        
        // Si la prochaine lettre dépasse la largeur du QR Code, on retourne à la ligne
        if (x + (8 * scale) > img_width) { 
            x = start_x;
            y += 10 * scale; // Descend d'une ligne (+ espacement)
            if (y >= img_height) break; // Arrêt d'urgence si le texte est vraiment trop long
        }
        text++;
    }
}
static int generer_qr_png(const char *texte, int id_donnee) {
    CreateDirectoryA("QRCodes_Export", NULL);

    char filepath[MAX_PATH];
    snprintf(filepath, sizeof(filepath), "QRCodes_Export/Donnee_%03d.png", id_donnee);

    uint8_t qr0[qrcodegen_BUFFER_LEN_MAX];
    uint8_t tempBuffer[qrcodegen_BUFFER_LEN_MAX];

    bool ok = qrcodegen_encodeText(texte, tempBuffer, qr0, qrcodegen_Ecc_LOW,
                                   qrcodegen_VERSION_MIN, qrcodegen_VERSION_MAX,
                                   qrcodegen_Mask_AUTO, true);
    if (!ok) return 0;

    int size = qrcodegen_getSize(qr0);
    int border = 4;
    int scale = 10;
    
    // NOUVEAU : On définit une zone spécifique pour l'image
    int img_width = (size + border * 2) * scale;
    int text_area_height = 60; // 60 pixels de haut pour écrire le texte
    int img_height = img_width + text_area_height; 

    uint8_t *img_data = (uint8_t *)malloc(img_width * img_height * 3);
    if (!img_data) return 0;

    // NOUVEAU : On remplit TOUTE l'image de blanc (Background)
    for (int i = 0; i < img_width * img_height * 3; i++) {
        img_data[i] = 0xFF; 
    }

    // Dessin du QR Code
    for (int y = 0; y < img_width; y++) {
        if (y >= img_width) break; // Sécurité : on s'arrête à la bordure du QR code
        for (int x = 0; x < img_width; x++) {
            int module_x = (x / scale) - border;
            int module_y = (y / scale) - border;

            if (qrcodegen_getModule(qr0, module_x, module_y)) {
                int pixel_idx = (y * img_width + x) * 3;
                img_data[pixel_idx + 0] = 0x00;
                img_data[pixel_idx + 1] = 0x00;
                img_data[pixel_idx + 2] = 0x00;
            }
        }
    }

    // NOUVEAU : On écrit le texte dans la bande blanche du bas
    // (x=10, y=img_width, texte, echelle=2 pour que ce soit bien lisible)
    draw_string(img_data, img_width, img_height, 10, img_width, texte, 2);

    int stride_in_bytes = img_width * 3;
    int result = stbi_write_png(filepath, img_width, img_height, 3, img_data, stride_in_bytes);

    free(img_data);
    return result;
}
/**
 * Concatène les messages, ajoute l'horodatage système, et transfère 
 * l'affichage à la boucle principale de manière Thread-Safe (évite les crashs graphiques croisés).
 */
static void post_log(const char *fmt, ...) {
    char msg[600];
    va_list args;
    va_start(args, fmt);
    vsnprintf(msg, sizeof(msg), fmt, args); // Écriture avec variables multiples comme un printf standard
    va_end(args);

    SYSTEMTIME st;
    GetLocalTime(&st);
    char *full = malloc(700); // L'allocation dynamique évite d'écraser le message si le thread en envoie plusieurs rapidement
    snprintf(full, 700, "[%02d:%02d:%02d] %s\r\n", st.wHour, st.wMinute, st.wSecond, msg);
    // Délégation asynchrone : "Message Queue", traité quand le Thread GUI est disponible
    PostMessage(g_hWndMain, WM_APP_LOG, 0, (LPARAM)full); 
}

/* ---------------------------------------------------------------------- */
/* Thread de travail en arrière-plan (Worker)                             */
/* ---------------------------------------------------------------------- */
static DWORD WINAPI WorkerThreadProc(LPVOID param) {
    (void)param;

    if (g_mode == MODE_AUTO) {
        while (!g_stopRequested) {
            const char *data = get_donnee_data(&g_csv, g_currentdonnee);
            if (!data) {
                post_log("Toutes les donnees du fichier ont ete attribuees.");
                break;
            }

            // Lecture de l'état des cases d'exportation
            BOOL isRfidChecked = (SendMessage(g_hCheckRfid, BM_GETCHECK, 0, 0) == BST_CHECKED);
            BOOL isQrChecked = (SendMessage(g_hCheckQr, BM_GETCHECK, 0, 0) == BST_CHECKED);

            // --- ÉTAPE 1 : GÉNÉRATION DU QR CODE ---
            if (isQrChecked) {
                if (generer_qr_png(data, g_currentdonnee)) {
                    post_log("QR Code généré (Donnee_%03d.png).", g_currentdonnee);
                } else {
                    post_log("Erreur lors de la création du QR Code %d.", g_currentdonnee);
                }
            }

            // --- ÉTAPE 2 : ENCODAGE DE LA PUCE RFID ---
            if (isRfidChecked) {
                post_log("donnée %d : presentez le tag...", g_currentdonnee);
                SCARDHANDLE hCard = wait_for_card_i(&g_stopRequested);
                if (!hCard) break;

                int ok = write_ntag(hCard, data);
                SCardDisconnect(hCard, SCARD_LEAVE_CARD);

                if (ok) {
                    post_log("donnée %d écrite et vérifiée : %s", g_currentdonnee, data);
                    g_currentdonnee++;
                    save_state(g_currentdonnee);
                    PostMessage(g_hWndMain, WM_APP_donnee, (WPARAM)g_currentdonnee, 0);
                } else {
                    post_log("ECHEC d'écriture pour la donnée %d. Reessayez.", g_currentdonnee);
                }
                wait_for_card_removal_i(&g_stopRequested);
            } 
            // --- ÉTAPE 3 : BYPASS MATÉRIEL ---
            else {
                // Si l'utilisateur n'a coché QUE la case QR Code, on incrémente 
                // directement pour passer à la ligne suivante sans attendre de badge.
                g_currentdonnee++;
                save_state(g_currentdonnee);
                PostMessage(g_hWndMain, WM_APP_donnee, (WPARAM)g_currentdonnee, 0);
                Sleep(20); // Petite pause pour ne pas figer l'interface si on génère 500 images d'un coup
            }
        }
    } else if (g_mode == MODE_MANUAL) {
        post_log("Mode manuel actif.");

        BOOL isRfidChecked = (SendMessage(g_hCheckRfid, BM_GETCHECK, 0, 0) == BST_CHECKED);
        BOOL isQrChecked = (SendMessage(g_hCheckQr, BM_GETCHECK, 0, 0) == BST_CHECKED);

        // --- ÉTAPE 1 : GÉNÉRATION DU QR CODE MANUEL ---
        if (isQrChecked) {
            // On utilise l'ID 0 pour signifier que c'est l'export manuel
            if (generer_qr_png(g_manualValueSnapshot, 0)) {
                post_log("QR Code manuel généré (Donnee_000.png).");
            } else {
                post_log("Erreur lors de la création du QR Code manuel.");
            }
        }

        // --- ÉTAPE 2 : ENCODAGE RFID EN BOUCLE ---
        if (isRfidChecked) {
            post_log("Chaque tag présenté recevra : \"%s\"", g_manualValueSnapshot);
            while (!g_stopRequested) {
                SCARDHANDLE hCard = wait_for_card_i(&g_stopRequested);
                if (!hCard) break;

                int ok = write_ntag(hCard, g_manualValueSnapshot);
                SCardDisconnect(hCard, SCARD_LEAVE_CARD);

                if (ok) post_log("Tag écrit et vérifié : %s", g_manualValueSnapshot);
                else post_log("ECHEC d'écriture sur ce tag. Reessayez.");

                wait_for_card_removal_i(&g_stopRequested);
            }
        } else {
            // Si seule la case QR Code était cochée, on a fini le travail.
            post_log("Export manuel terminé.");
        }
    } else if (g_mode == MODE_READ) {
        post_log("Mode lecture actif : présentez des tags pour afficher leur contenu.");
        
        BOOL isQrChecked = (SendMessage(g_hCheckQr, BM_GETCHECK, 0, 0) == BST_CHECKED);
        int read_counter = 900; // Compteur arbitraire pour numéroter les scans (Donnee_900.png, etc.)

        while (!g_stopRequested) {
            SCARDHANDLE hCard = wait_for_card_i(&g_stopRequested);
            if (!hCard) break;

            char url[300];
            int ok = read_ntag_auto(hCard, url, sizeof(url));
            SCardDisconnect(hCard, SCARD_LEAVE_CARD);

            if (ok) {
                post_log("Tag lu : %s", url);
                
                // --- CONVERSION À LA VOLÉE ---
                if (isQrChecked) {
                    if (generer_qr_png(url, read_counter)) {
                        post_log("=> QR Code du tag sauvegardé (Donnee_%03d.png)", read_counter);
                        read_counter++;
                    }
                }
            } else {
                post_log("Aucun enregistrement NDEF URI valide trouvé sur ce tag.");
            }

            wait_for_card_removal_i(&g_stopRequested);
        }
    }

    // Sortie propre du Thread, notifie le moteur d'Interface
    g_isRunning = FALSE;
    PostMessage(g_hWndMain, WM_APP_DONE, 0, 0);
    return 0;
}

/**
 * Permet de verrouiller (griser) ou libérer d'un coup l'ensemble des clics possibles de l'UI.
 */
static void SetControlsEnabled(BOOL enabled) {
    EnableWindow(g_hBrowseBtn, enabled);
    EnableWindow(g_hImportExcelBtn, enabled);
    EnableWindow(g_hGotoEdit, enabled);
    EnableWindow(g_hGotoBtn, enabled);
    EnableWindow(g_hResetBtn, enabled);
    EnableWindow(g_hManualEdit, enabled);
    EnableWindow(g_hRadioAuto, enabled);
    EnableWindow(g_hRadioManual, enabled);
    EnableWindow(g_hRadioRead, enabled);
}

/**
 * Actualise visuellement la fraction "N / X"
 */
static void UpdatedonneeLabel(void) {
    char buf[64];
    snprintf(buf, sizeof(buf), "%d / %d", g_currentdonnee, g_csv.count);
    SetWindowText(g_hdonneeLabel, buf);
}

/* ---------------------------------------------------------------------- */
/* Procedure de fenetre principale (Cœur du Système de l'Interface GUI)   */
/* ---------------------------------------------------------------------- */
static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_COMMAND: {
            int id = LOWORD(wParam); // Identifiant de l'élément cliqué

            if (id == IDC_BROWSE) {
                // Boîte de dialogue standard Windows "Fichier à ouvrir" pour CSV
                char path[MAX_PATH] = { 0 };
                OPENFILENAMEA ofn;
                ZeroMemory(&ofn, sizeof(ofn));
                ofn.lStructSize = sizeof(ofn);
                ofn.hwndOwner = hwnd;
                ofn.lpstrFile = path;
                ofn.nMaxFile = sizeof(path);
                ofn.lpstrFilter = "Fichiers CSV\0*.csv\0Tous les fichiers\0*.*\0";
                ofn.nFilterIndex = 1;
                ofn.lpstrTitle = "Selectionner le fichier CSV";
                ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_NOCHANGEDIR;

                if (GetOpenFileNameA(&ofn)) {
                    if (g_csv.lines) free_csv(&g_csv); // Décharge la table précédente
                    g_csv = load_csv(path);
                    strncpy(g_csvPath, path, sizeof(g_csvPath) - 1);
                    SetWindowText(g_hCsvLabel, path);
                    save_state(g_currentdonnee);
                    UpdatedonneeLabel();
                    post_log("Fichier CSV chargé : %d données.", g_csv.count);
                }
            } 
            else if (id == IDC_IMPORT_EXCEL) {
                // Boîte de dialogue spécialisée pour Excel XLSX et déclenchement VBScript
                char path[MAX_PATH] = { 0 };
                OPENFILENAMEA ofn;
                ZeroMemory(&ofn, sizeof(ofn));
                ofn.lStructSize = sizeof(ofn);
                ofn.hwndOwner = hwnd;
                ofn.lpstrFile = path;
                ofn.nMaxFile = sizeof(path);
                ofn.lpstrFilter = "Fichiers Excel\0*.xlsx;*.xls\0Tous les fichiers\0*.*\0";
                ofn.nFilterIndex = 1;
                ofn.lpstrTitle = "Selectionner le fichier Excel";
                ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_NOCHANGEDIR;

                if (GetOpenFileNameA(&ofn)) {
                    post_log("Lancement de l'importation Excel...");
                    if (import_excel_via_vbs_gui(path, "donnees.csv")) {
                        if (g_csv.lines) free_csv(&g_csv);
                        g_csv = load_csv("donnees.csv");
                        strncpy(g_csvPath, "donnees.csv", sizeof(g_csvPath) - 1);
                        SetWindowText(g_hCsvLabel, "donnees.csv (Import Excel)");
                        g_currentdonnee = 1; // Un nouvel import doit toujours forcer la remise à la ligne 1
                        save_state(g_currentdonnee);
                        UpdatedonneeLabel();
                        post_log("Excel converti avec succès : %d données chargées.", g_csv.count);
                    } else {
                        post_log("Importation Excel annulée ou échouée.");
                    }
                }
            }
            else if (id == IDC_GOTO_BTN) {
                // Saisie manuelle d'une ligne précise pour relancer une série spécifique (Correction/Oubli)
                char buf[32];
                GetWindowText(g_hGotoEdit, buf, sizeof(buf));
                int target = atoi(buf);
                if (target < 1) {
                    MessageBox(hwnd, "Numéro de donnée invalide.", "Erreur", MB_OK | MB_ICONWARNING);
                } else {
                    g_currentdonnee = target;
                    save_state(g_currentdonnee);
                    UpdatedonneeLabel();
                    post_log("Positionné sur la donnée %d.", g_currentdonnee);
                }
            } else if (id == IDC_RESET_BTN) {
                // Reboot du compteur d'avancement
                g_currentdonnee = 1;
                save_state(g_currentdonnee);
                UpdatedonneeLabel();
                post_log("Remise à zéro effectuée (donnée 1).");
            } else if (id == IDC_STARTSTOP_BTN) {
                // Séquence d'Engagement ou Interruption d'une action matérielle
                // 1. On lit l'état des cases d'exportation
                BOOL isRfidChecked = (SendMessage(g_hCheckRfid, BM_GETCHECK, 0, 0) == BST_CHECKED);
                BOOL isQrChecked = (SendMessage(g_hCheckQr, BM_GETCHECK, 0, 0) == BST_CHECKED);

                // 2. On vérifie si l'utilisateur est en mode "Lecture"
                BOOL isReadMode = (SendMessage(g_hRadioRead, BM_GETCHECK, 0, 0) == BST_CHECKED);

                // 3. Condition de blocage : Si on n'est PAS en lecture, et qu'aucun export n'est coché
                if (!isReadMode && !isRfidChecked && !isQrChecked) {
                    MessageBox(hwnd, "Veuillez cocher au moins une cible d'écriture.",
                          "Sélectionner un mode de fonctionnement", MB_OK | MB_ICONWARNING);
                    return 0; // On bloque le processus ici
                }
                if (!g_isRunning) {
                    // Analyse logique des radio-boutons pour dicter le comportement du Thread
                    if (SendMessage(g_hRadioManual, BM_GETCHECK, 0, 0) == BST_CHECKED) {
                        g_mode = MODE_MANUAL;
                        GetWindowText(g_hManualEdit, g_manualValueSnapshot, sizeof(g_manualValueSnapshot));
                        if (strlen(g_manualValueSnapshot) == 0) {
                            MessageBox(hwnd, "Entrez une valeur à écrire avant de démarrer.",
                                       "Valeur manquante", MB_OK | MB_ICONWARNING);
                            return 0;
                        }
                    } else if (SendMessage(g_hRadioRead, BM_GETCHECK, 0, 0) == BST_CHECKED) {
                        g_mode = MODE_READ;
                    } else {
                        g_mode = MODE_AUTO;
                        if (g_csv.count == 0) {
                            MessageBox(hwnd, "Chargez d'abord un fichier CSV ou Excel.", "Aucune donnée",
                                       MB_OK | MB_ICONWARNING);
                            return 0;
                        }
                    }

                    // Ne lance pas la mécanique si aucun périphérique compatible détecté
                    if (!g_readerReady) {
                        MessageBox(hwnd, "Aucun lecteur détecté. Branchez un lecteur RFID PC/SC compatible.",
                                   "Erreur", MB_OK | MB_ICONERROR);
                        return 0;
                    }

                    // Allumage du thread et verrouillage Interface
                    g_stopRequested = FALSE;
                    g_isRunning = TRUE;
                    SetWindowText(g_hStartStopBtn, "Arrêter");
                    SetControlsEnabled(FALSE);
                    CloseHandle(CreateThread(NULL, 0, WorkerThreadProc, NULL, 0, NULL)); // Déclenchement Asynchrone
                } else {
                    // Ordre de fin à l'intention du Thread (ne le kill pas violemment, laisse finir le cycle actuel)
                    g_stopRequested = TRUE;
                    EnableWindow(g_hStartStopBtn, FALSE);
                    SetWindowText(g_hStartStopBtn, "Arrêt en cours...");
                }
            }
            return 0;
        }

        case WM_APP_LOG: {
            // Unité de réception : Réception des paquets mémoires depuis "post_log"
            char *text = (char *)lParam;
            int len = GetWindowTextLength(g_hLogEdit);
            SendMessage(g_hLogEdit, EM_SETSEL, (WPARAM)len, (LPARAM)len); // Positionne le curseur tout en bas
            SendMessage(g_hLogEdit, EM_REPLACESEL, FALSE, (LPARAM)text);  // Insère le nouveau flux
            free(text); // TRÈS IMPORTANT : Détruit le pointeur alloué, sinon fuite mémoire géante
            return 0;
        }

        case WM_APP_donnee: {
            // Unité de réception : MàJ affichage ID listé
            g_currentdonnee = (int)wParam;
            UpdatedonneeLabel();
            return 0;
        }

        case WM_APP_DONE: {
            // Unité de réception : Déverrouillage visuel après fin d'une tâche asynchrone
            SetWindowText(g_hStartStopBtn, "Démarrer");
            EnableWindow(g_hStartStopBtn, TRUE);
            SetControlsEnabled(TRUE);
            return 0;
        }

        case WM_CLOSE:
            // Croix de fermeture de fenêtre en haut à droite
            if (g_isRunning) {
                // Interdit de crasher le programme si un encodage physique est en cours
                MessageBox(hwnd, "Arrêtez le mode en cours avant de fermer.", "Mode actif", MB_OK | MB_ICONWARNING);
                return 0;
            }
            DestroyWindow(hwnd);
            return 0;

        case WM_DESTROY:
            // Nettoyage final du processus
            if (g_hContext) SCardReleaseContext(g_hContext); // Restitue le controle du lecteur à Windows
            PostQuitMessage(0);
            return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/* ---------------------------------------------------------------------- */
/* Utilitaire C pour simplifier l'implémentation graphique (DRY principle)*/
/* ---------------------------------------------------------------------- */
static HWND CreateCtrl(LPCSTR className, LPCSTR text, DWORD style,
                        int x, int y, int w, int h, HWND parent, int id) {
    // Instancie un élément (bouton, texte, tableau) avec des coordonnées X/Y strictes
    HWND hwnd = CreateWindowEx(0, className, text, style, x, y, w, h,
                                parent, (HMENU)(INT_PTR)id, g_hInst, NULL);
    SendMessage(hwnd, WM_SETFONT, (WPARAM)g_hFont, TRUE); // Héritage de la police standard Microsoft
    return hwnd;
}

/* ---------------------------------------------------------------------- */
/* Point de départ officiel de l'application sous l'OS Windows.           */
/* (Remplace la classique fonction "main" utilisée en mode console)       */
/* ---------------------------------------------------------------------- */
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    (void)hPrevInstance; (void)lpCmdLine;
    g_hInst = hInstance;
    g_hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

    // Initialisation Classe de Fenêtre
    WNDCLASS wc = { 0 };
    wc.lpfnWndProc = WndProc; // Attachement du centre de contrôle principal
    wc.hInstance = hInstance;
    wc.lpszClassName = "RFIDdonneesGUIClass";
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    RegisterClass(&wc);

    // Forme finale de la fenêtre d'application (Taille Fixe : 660x620)
    g_hWndMain = CreateWindowEx(0, wc.lpszClassName, "Attribution de données RFID",
                                 WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
                                 CW_USEDEFAULT, CW_USEDEFAULT, 660, 620,
                                 NULL, NULL, hInstance, NULL);

    /* --- CREATION DES BLOCS VISUELS (Positionnés au pixel près) --- */
    /* --- Fichier --- */
    CreateCtrl("STATIC", "Fichier data :", WS_CHILD | WS_VISIBLE, 10, 14, 85, 20, g_hWndMain, 0);
    g_hCsvLabel = CreateCtrl("STATIC", "Aucun fichier chargé", WS_CHILD | WS_VISIBLE | SS_LEFTNOWORDWRAP,
                              95, 14, 310, 20, g_hWndMain, IDC_CSV_LABEL);
    
    g_hBrowseBtn = CreateCtrl("BUTTON", "Parcourir CSV...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                               415, 10, 105, 26, g_hWndMain, IDC_BROWSE);
    g_hImportExcelBtn = CreateCtrl("BUTTON", "Importer Excel...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                               525, 10, 115, 26, g_hWndMain, IDC_IMPORT_EXCEL);


    /* --- donnee courante --- */
    CreateCtrl("STATIC", "Donnée actuelle :", WS_CHILD | WS_VISIBLE, 10, 50, 100, 20, g_hWndMain, 0);
    g_hdonneeLabel = CreateCtrl("STATIC", "- / -", WS_CHILD | WS_VISIBLE, 115, 50, 200, 20, g_hWndMain, IDC_donnee_LABEL);

    /* --- Groupe de modes (Radio Buttons mutuellement exclusifs) --- */
    /* --- Groupe 1 : Méthode d'entrée (Côté Gauche) --- */
    CreateCtrl("BUTTON", "Méthode d'entrée", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 10, 80, 310, 90, g_hWndMain, 0);
    g_hRadioAuto = CreateCtrl("BUTTON", "Automatique (fichier CSV / Excel)", WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON | WS_GROUP,
                               25, 100, 250, 20, g_hWndMain, IDC_RADIO_AUTO);
    g_hRadioManual = CreateCtrl("BUTTON", "Ecriture manuelle", WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON,
                                 25, 124, 250, 20, g_hWndMain, IDC_RADIO_MANUAL);
    g_hRadioRead = CreateCtrl("BUTTON", "Lecture de tag", WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON,
                               25, 148, 250, 20, g_hWndMain, IDC_RADIO_READ);
    SendMessage(g_hRadioAuto, BM_SETCHECK, BST_CHECKED, 0);

    /* --- Groupe 2 : Mode d'export (Côté Droit) --- */
    CreateCtrl("BUTTON", "Mode d'export", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 330, 80, 310, 90, g_hWndMain, 0);
    
    g_hCheckRfid = CreateCtrl("BUTTON", "Encodage puce RFID", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                               345, 100, 250, 20, g_hWndMain, IDC_CHECK_RFID);
    g_hCheckQr = CreateCtrl("BUTTON", "Génération fichier QR Code", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                               345, 124, 250, 20, g_hWndMain, IDC_CHECK_QR);

    // Initialisation : les deux cases sont cochées par défaut au démarrage
    SendMessage(g_hCheckRfid, BM_SETCHECK, BST_CHECKED, 0);
    SendMessage(g_hCheckQr, BM_SETCHECK, BST_CHECKED, 0);


    /* --- Controles mode automatique --- */
    CreateCtrl("STATIC", "Aller à la donnée n° :", WS_CHILD | WS_VISIBLE, 10, 185, 120, 20, g_hWndMain, 0);
    g_hGotoEdit = CreateCtrl("EDIT", "", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL, 135, 183, 70, 24, g_hWndMain, IDC_GOTO_EDIT);
    g_hGotoBtn = CreateCtrl("BUTTON", "Aller", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 215, 182, 70, 26, g_hWndMain, IDC_GOTO_BTN);
    g_hResetBtn = CreateCtrl("BUTTON", "Remise à zéro (donnée 1)", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 300, 182, 180, 26, g_hWndMain, IDC_RESET_BTN);

    /* --- Controle mode manuel --- */
    CreateCtrl("STATIC", "Valeur à écrire :", WS_CHILD | WS_VISIBLE, 10, 222, 150, 20, g_hWndMain, 0);
    g_hManualEdit = CreateCtrl("EDIT", "", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL, 10, 244, 630, 24, g_hWndMain, IDC_MANUAL_EDIT);

    /* --- Demarrer / Arreter --- */
    g_hStartStopBtn = CreateCtrl("BUTTON", "Démarrer", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, 10, 280, 630, 34, g_hWndMain, IDC_STARTSTOP_BTN);

    /* --- Journal Console (Scrollable Text Area) --- */
    CreateCtrl("STATIC", "Journal :", WS_CHILD | WS_VISIBLE, 10, 326, 100, 20, g_hWndMain, 0);
    g_hLogEdit = CreateCtrl("EDIT", "", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
                             10, 348, 630, 220, g_hWndMain, IDC_LOG_EDIT);

    /* --- Procédure Initiale Silencieuse --- */
    g_readerReady = init_pcsc();
    g_currentdonnee = load_state();
    if (g_csvPath[0] != '\0') {
        g_csv = load_csv(g_csvPath);
        if (g_csv.count > 0) SetWindowText(g_hCsvLabel, g_csvPath);
    }
    UpdatedonneeLabel();

    // Affichage natif et forçage de rendu final
    ShowWindow(g_hWndMain, nCmdShow);
    UpdateWindow(g_hWndMain);

    // Pop-up d'alerte si l'utilisateur lance le soft sans avoir branché l'ACR122U
    if (!g_readerReady) {
        MessageBox(g_hWndMain, "Aucun lecteur détecte au démarrage.\nBranchez un lecteur RFID PC/SC compatible et redémarrez.",
                   "Lecteur introuvable", MB_OK | MB_ICONWARNING);
    } else {
        post_log("Lecteur détecté : %s", g_readerName);
    }
    if (g_csv.count > 0) post_log("Fichier rechargé : %d données.", g_csv.count);

    /* --- MESSAGE PUMP (Coeur de Réactivité Windows) --- */
    MSG msg;
    // Intercepte toute interaction utilisateur et la route de manière appropriée vers le "switch" de "WndProc"
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return 0;
}