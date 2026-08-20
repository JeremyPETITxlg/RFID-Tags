/* ============================================================================
 * rfid_zones_gui.c
 *
 * Version "interface graphique" du programme d'attribution de zones RFID
 * (NTAG213/215/216) via lecteur ACR122U (PC/SC), pour Windows.
 *
 * Reprend exactement la meme logique bas niveau que la version console
 * (rfid_zones.c) : lecture/ecriture PC/SC, encodage NDEF URI, relecture de
 * verification, gestion du BOM UTF-8, memorisation de la progression.
 *
 * Nouveautes :
 *   - Fenetre graphique (pas de terminal).
 *   - Bouton pour charger le fichier CSV depuis l'application (explorateur
 *     de fichiers Windows).
 *   - Mode "Lecture" : affiche le contenu d'un tag scanne sans le modifier.
 *   - Mode "ecriture manuelle" : une valeur saisie par l'utilisateur est
 *     ecrite sur chaque tag presente, en boucle, jusqu'a l'arret du mode.
 *   - Mode "Automatique" : reprend le comportement du programme console
 *     (avance case par case dans le CSV), en continu tant que le mode est
 *     actif.
 * ==========================================================================*/

#include <windows.h>
#include <winscard.h>
#include <commdlg.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdarg.h>

#pragma comment(lib, "winscard.lib")
#pragma comment(lib, "comdlg32.lib")

/* ---------------------------------------------------------------------- */
/* Constantes                                                              */
/* ---------------------------------------------------------------------- */

#define STATE_FILE          "state.txt"

#define NTAG_PAGE_SIZE       4
#define NTAG_USER_START      4
/* Valeur prudente par defaut (NTAG213). Augmentez si vous utilisez des
 * NTAG215 (jusqu'a ~126) ou NTAG216 (jusqu'a ~222) avec des URLs longues. */
#define NTAG_USER_PAGES      36

#define POLL_MS              300  /* granularite du sondage interruptible */

#define IDC_BROWSE          101
#define IDC_CSV_LABEL       102
#define IDC_ZONE_LABEL      103
#define IDC_RADIO_AUTO      104
#define IDC_RADIO_MANUAL    105
#define IDC_RADIO_READ      106
#define IDC_GOTO_EDIT       107
#define IDC_GOTO_BTN        108
#define IDC_RESET_BTN       109
#define IDC_MANUAL_EDIT     110
#define IDC_STARTSTOP_BTN   111
#define IDC_LOG_EDIT        112

#define WM_APP_LOG   (WM_APP + 1)
#define WM_APP_ZONE  (WM_APP + 2)
#define WM_APP_DONE  (WM_APP + 3)

typedef enum { MODE_AUTO, MODE_MANUAL, MODE_READ } AppMode;

/* ---------------------------------------------------------------------- */
/* Donnees CSV                                                             */
/* ---------------------------------------------------------------------- */

typedef struct {
    char **lines;
    int count;
} CsvData;

static CsvData load_csv(const char *path) {
    CsvData csv = { NULL, 0 };
    FILE *f = fopen(path, "rb");
    if (!f) return csv;

    /* Ignore un eventuel BOM UTF-8 (Excel "CSV UTF-8") */
    unsigned char bom[3];
    size_t read = fread(bom, 1, 3, f);
    if (!(read == 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)) {
        fseek(f, 0, SEEK_SET);
    }

    int capacity = 1024;
    csv.lines = malloc(capacity * sizeof(char *));

    char buffer[1024];
    while (fgets(buffer, sizeof(buffer), f)) {
        size_t len = strlen(buffer);
        while (len > 0 && (buffer[len - 1] == '\n' || buffer[len - 1] == '\r')) {
            buffer[--len] = '\0';
        }
        if (len == 0) continue; /* ignore les lignes vides */

        if (csv.count >= capacity) {
            capacity *= 2;
            csv.lines = realloc(csv.lines, capacity * sizeof(char *));
        }
        csv.lines[csv.count] = _strdup(buffer);
        csv.count++;
    }
    fclose(f);
    return csv;
}

static void free_csv(CsvData *csv) {
    for (int i = 0; i < csv->count; i++) free(csv->lines[i]);
    free(csv->lines);
    csv->lines = NULL;
    csv->count = 0;
}

static const char *get_zone_data(CsvData *csv, int zone_id) {
    if (zone_id < 1 || zone_id > csv->count) return NULL;
    return csv->lines[zone_id - 1];
}

/* ---------------------------------------------------------------------- */
/* etat persistant (zone courante + chemin du CSV)                        */
/* ---------------------------------------------------------------------- */

static char g_csvPath[MAX_PATH] = { 0 };

static int load_state(void) {
    FILE *f = fopen(STATE_FILE, "r");
    if (!f) return 1;

    char line[MAX_PATH + 32];
    int zone = 1;
    while (fgets(line, sizeof(line), f)) {
        line[strcspn(line, "\r\n")] = '\0';
        if (strncmp(line, "current_zone=", 13) == 0) {
            zone = atoi(line + 13);
        } else if (strncmp(line, "csv_path=", 9) == 0) {
            strncpy(g_csvPath, line + 9, sizeof(g_csvPath) - 1);
        }
    }
    fclose(f);
    if (zone < 1) zone = 1;
    return zone;
}

static void save_state(int zone) {
    FILE *f = fopen(STATE_FILE, "w");
    if (!f) return;
    fprintf(f, "current_zone=%d\n", zone);
    if (g_csvPath[0] != '\0') fprintf(f, "csv_path=%s\n", g_csvPath);
    fclose(f);
}

/* ---------------------------------------------------------------------- */
/* PC/SC bas niveau                                                        */
/* ---------------------------------------------------------------------- */

static SCARDCONTEXT g_hContext;
static char g_readerName[256];
static BOOL g_readerReady = FALSE;

static BOOL init_pcsc(void) {
    LONG rv = SCardEstablishContext(SCARD_SCOPE_SYSTEM, NULL, NULL, &g_hContext);
    if (rv != SCARD_S_SUCCESS) return FALSE;

    char readers[1024];
    DWORD readersLen = sizeof(readers);
    rv = SCardListReaders(g_hContext, NULL, readers, &readersLen);
    if (rv != SCARD_S_SUCCESS) return FALSE;

    strncpy(g_readerName, readers, sizeof(g_readerName) - 1);
    return TRUE;
}

static int write_page(SCARDHANDLE hCard, BYTE page, const BYTE data4[4]) {
    BYTE apdu[9] = { 0xFF, 0xD6, 0x00, page, 0x04, 0, 0, 0, 0 };
    memcpy(apdu + 5, data4, 4);

    BYTE response[16];
    DWORD responseLen = sizeof(response);
    LONG rv = SCardTransmit(hCard, SCARD_PCI_T1, apdu, sizeof(apdu),
                             NULL, response, &responseLen);
    if (rv != SCARD_S_SUCCESS) return 0;
    return (responseLen >= 2 && response[responseLen - 2] == 0x90 &&
            response[responseLen - 1] == 0x00);
}

static int read_page(SCARDHANDLE hCard, BYTE page, BYTE out4[4]) {
    BYTE apdu[5] = { 0xFF, 0xB0, 0x00, page, 0x04 };
    BYTE response[16];
    DWORD responseLen = sizeof(response);
    LONG rv = SCardTransmit(hCard, SCARD_PCI_T1, apdu, sizeof(apdu),
                             NULL, response, &responseLen);
    if (rv != SCARD_S_SUCCESS) return 0;
    if (responseLen < 6) return 0;
    if (response[responseLen - 2] != 0x90 || response[responseLen - 1] != 0x00) return 0;
    memcpy(out4, response, 4);
    return 1;
}

/* Attente interruptible (sondage toutes les POLL_MS ms) d'une presentation
 * de tag. Retourne 0 si l'arret a ete demande via *stopFlag. */
static SCARDHANDLE wait_for_card_i(volatile BOOL *stopFlag) {
    SCARD_READERSTATE state;
    memset(&state, 0, sizeof(state));
    state.szReader = g_readerName;
    state.dwCurrentState = SCARD_STATE_EMPTY;

    while (!*stopFlag) {
        LONG rv = SCardGetStatusChange(g_hContext, POLL_MS, &state, 1);
        if (rv == (LONG)SCARD_E_TIMEOUT) continue;
        if (rv != SCARD_S_SUCCESS) return 0;
        if (state.dwEventState & SCARD_STATE_PRESENT) break;
        state.dwCurrentState = state.dwEventState;
    }
    if (*stopFlag) return 0;

    SCARDHANDLE hCard;
    DWORD activeProtocol;
    LONG rv = SCardConnect(g_hContext, g_readerName, SCARD_SHARE_SHARED,
                            SCARD_PROTOCOL_T0 | SCARD_PROTOCOL_T1,
                            &hCard, &activeProtocol);
    if (rv != SCARD_S_SUCCESS) return 0;
    return hCard;
}

static void wait_for_card_removal_i(volatile BOOL *stopFlag) {
    SCARD_READERSTATE state;
    memset(&state, 0, sizeof(state));
    state.szReader = g_readerName;
    state.dwCurrentState = SCARD_STATE_PRESENT;

    while (!*stopFlag) {
        LONG rv = SCardGetStatusChange(g_hContext, POLL_MS, &state, 1);
        if (rv == (LONG)SCARD_E_TIMEOUT) continue;
        if (rv != SCARD_S_SUCCESS) return;
        if (state.dwEventState & SCARD_STATE_EMPTY) return;
        state.dwCurrentState = state.dwEventState;
    }
}

/* ---------------------------------------------------------------------- */
/* Encodage / decodage NDEF URI                                           */
/* ---------------------------------------------------------------------- */

static const char *uri_prefix_for_code(BYTE code) {
    switch (code) {
        case 0x01: return "http://www.";
        case 0x02: return "https://www.";
        case 0x03: return "http://";
        case 0x04: return "https://";
        default:   return "";
    }
}

static BYTE uri_code_for_url(const char *url, const char **payloadStart) {
    if (strncmp(url, "https://www.", 12) == 0) { *payloadStart = url + 12; return 0x02; }
    if (strncmp(url, "http://www.", 11) == 0)  { *payloadStart = url + 11; return 0x01; }
    if (strncmp(url, "https://", 8) == 0)      { *payloadStart = url + 8;  return 0x04; }
    if (strncmp(url, "http://", 7) == 0)       { *payloadStart = url + 7;  return 0x03; }
    *payloadStart = url;
    return 0x00;
}

static size_t build_ndef_uri(const char *url, BYTE *out, size_t outCapacity) {
    const char *payloadStr;
    BYTE code = uri_code_for_url(url, &payloadStr);
    size_t strLen = strlen(payloadStr);
    size_t payloadLen = 1 + strLen;
    if (payloadLen > 255) return 0;

    size_t recordLen = 4 + payloadLen;
    size_t needed = 2 + recordLen + 1;
    if (needed > outCapacity) return 0;

    size_t pos = 0;
    out[pos++] = 0x03;
    out[pos++] = (BYTE)recordLen;
    out[pos++] = 0xD1;
    out[pos++] = 0x01;
    out[pos++] = (BYTE)payloadLen;
    out[pos++] = 'U';
    out[pos++] = code;
    memcpy(out + pos, payloadStr, strLen);
    pos += strLen;
    out[pos++] = 0xFE;
    return pos;
}

static int decode_ndef_uri(const BYTE *buffer, size_t total, char *outUrl, size_t outUrlCap) {
    if (total < 3 || buffer[0] != 0x03) return 0;
    size_t recordLen = buffer[1];
    size_t p = 2;
    if (recordLen < 3 || p + recordLen > total) return 0;
    if (buffer[p + 3] != 'U') return 0;
    size_t payloadLen = buffer[p + 2];
    BYTE code = buffer[p + 4];
    size_t strLen = payloadLen - 1;
    if (p + 4 + 1 + strLen > total) return 0;

    const char *prefix = uri_prefix_for_code(code);
    snprintf(outUrl, outUrlCap, "%s%.*s", prefix, (int)strLen, (const char *)(buffer + p + 5));
    return 1;
}

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

/* Lecture "a l'aveugle" : on ne connaît pas la longueur a l'avance, on lit
 * page par page jusqu'a pouvoir decoder un message NDEF complet. */
static int read_ntag_auto(SCARDHANDLE hCard, char *outUrl, size_t outUrlCap) {
    BYTE buffer[600];
    size_t total = 0;
    int maxPages = NTAG_USER_PAGES + 4;

    for (int i = 0; i < maxPages && total + 4 <= sizeof(buffer); i++) {
        BYTE page4[4];
        if (!read_page(hCard, (BYTE)(NTAG_USER_START + i), page4)) break;
        memcpy(buffer + total, page4, 4);
        total += 4;

        if (total >= 4 && buffer[0] != 0x03) return 0; /* pas de NDEF ici */

        if (total >= 6) {
            size_t recordLen = buffer[1];
            size_t needed = 2 + recordLen + 1;
            if (total >= needed) {
                return decode_ndef_uri(buffer, total, outUrl, outUrlCap);
            }
        }
    }
    return 0;
}

/* ecrit l'URL sous forme de message NDEF URI, puis relit et verifie. */
static int write_ntag(SCARDHANDLE hCard, const char *url) {
    BYTE msg[512];
    size_t msgLen = build_ndef_uri(url, msg, sizeof(msg));
    if (msgLen == 0) return 0;
    if (msgLen > (size_t)(NTAG_USER_PAGES * NTAG_PAGE_SIZE)) return 0;

    size_t paddedLen = ((msgLen + 3) / 4) * 4;
    BYTE *buffer = calloc(paddedLen, 1);
    memcpy(buffer, msg, msgLen);

    int pageCount = (int)(paddedLen / NTAG_PAGE_SIZE);
    for (int i = 0; i < pageCount; i++) {
        if (!write_page(hCard, (BYTE)(NTAG_USER_START + i), buffer + (i * NTAG_PAGE_SIZE))) {
            free(buffer);
            return 0;
        }
    }
    free(buffer);

    char reread[300];
    if (!read_back_url(hCard, pageCount, reread, sizeof(reread))) return 0;
    if (strcmp(reread, url) != 0) return 0;
    return 1;
}

/* ---------------------------------------------------------------------- */
/* Variables globales de l'interface / du thread de travail                */
/* ---------------------------------------------------------------------- */

static HINSTANCE g_hInst;
static HWND g_hWndMain;
static HWND g_hCsvLabel, g_hZoneLabel;
static HWND g_hRadioAuto, g_hRadioManual, g_hRadioRead;
static HWND g_hGotoEdit, g_hGotoBtn, g_hResetBtn, g_hBrowseBtn;
static HWND g_hManualEdit, g_hStartStopBtn, g_hLogEdit;
static HFONT g_hFont;

static CsvData g_csv = { NULL, 0 };
static int g_currentZone = 1;

static volatile BOOL g_isRunning = FALSE;
static volatile BOOL g_stopRequested = FALSE;
static AppMode g_mode;
static char g_manualValueSnapshot[512];

/* ---------------------------------------------------------------------- */
/* Journal (log) : ajout thread-safe via PostMessage                       */
/* ---------------------------------------------------------------------- */

static void post_log(const char *fmt, ...) {
    char msg[600];
    va_list args;
    va_start(args, fmt);
    vsnprintf(msg, sizeof(msg), fmt, args);
    va_end(args);

    SYSTEMTIME st;
    GetLocalTime(&st);
    char *full = malloc(700);
    snprintf(full, 700, "[%02d:%02d:%02d] %s\r\n", st.wHour, st.wMinute, st.wSecond, msg);
    PostMessage(g_hWndMain, WM_APP_LOG, 0, (LPARAM)full);
}

/* ---------------------------------------------------------------------- */
/* Thread de travail (un seul, reutilise pour les 3 modes)                 */
/* ---------------------------------------------------------------------- */

static DWORD WINAPI WorkerThreadProc(LPVOID param) {
    (void)param;

    if (g_mode == MODE_AUTO) {
        while (!g_stopRequested) {
            const char *data = get_zone_data(&g_csv, g_currentZone);
            if (!data) {
                post_log("Toutes les zones du fichier ont ete attribuees.");
                break;
            }
            post_log("Zone %d : presentez le tag...", g_currentZone);
            SCARDHANDLE hCard = wait_for_card_i(&g_stopRequested);
            if (!hCard) break;

            int ok = write_ntag(hCard, data);
            SCardDisconnect(hCard, SCARD_LEAVE_CARD);

            if (ok) {
                post_log("Zone %d ecrite et verifiee : %s", g_currentZone, data);
                g_currentZone++;
                save_state(g_currentZone);
                PostMessage(g_hWndMain, WM_APP_ZONE, (WPARAM)g_currentZone, 0);
            } else {
                post_log("eCHEC d'ecriture pour la zone %d. Reessayez (retirez et repositionnez le tag).", g_currentZone);
            }
            wait_for_card_removal_i(&g_stopRequested);
        }
    } else if (g_mode == MODE_MANUAL) {
        post_log("Mode manuel actif : chaque tag presente recevra \"%s\"", g_manualValueSnapshot);
        while (!g_stopRequested) {
            SCARDHANDLE hCard = wait_for_card_i(&g_stopRequested);
            if (!hCard) break;

            int ok = write_ntag(hCard, g_manualValueSnapshot);
            SCardDisconnect(hCard, SCARD_LEAVE_CARD);

            if (ok) post_log("Tag ecrit et verifie : %s", g_manualValueSnapshot);
            else post_log("eCHEC d'ecriture sur ce tag. Reessayez.");

            wait_for_card_removal_i(&g_stopRequested);
        }
    } else if (g_mode == MODE_READ) {
        post_log("Mode lecture actif : presentez des tags pour afficher leur contenu.");
        while (!g_stopRequested) {
            SCARDHANDLE hCard = wait_for_card_i(&g_stopRequested);
            if (!hCard) break;

            char url[300];
            int ok = read_ntag_auto(hCard, url, sizeof(url));
            SCardDisconnect(hCard, SCARD_LEAVE_CARD);

            if (ok) post_log("Tag lu : %s", url);
            else post_log("Aucun enregistrement NDEF URI valide trouve sur ce tag.");

            wait_for_card_removal_i(&g_stopRequested);
        }
    }

    g_isRunning = FALSE;
    PostMessage(g_hWndMain, WM_APP_DONE, 0, 0);
    return 0;
}

/* ---------------------------------------------------------------------- */
/* Aide : activer/desactiver les contrôles pendant l'execution             */
/* ---------------------------------------------------------------------- */

static void SetControlsEnabled(BOOL enabled) {
    EnableWindow(g_hBrowseBtn, enabled);
    EnableWindow(g_hGotoEdit, enabled);
    EnableWindow(g_hGotoBtn, enabled);
    EnableWindow(g_hResetBtn, enabled);
    EnableWindow(g_hManualEdit, enabled);
    EnableWindow(g_hRadioAuto, enabled);
    EnableWindow(g_hRadioManual, enabled);
    EnableWindow(g_hRadioRead, enabled);
}

static void UpdateZoneLabel(void) {
    char buf[64];
    snprintf(buf, sizeof(buf), "%d / %d", g_currentZone, g_csv.count);
    SetWindowText(g_hZoneLabel, buf);
}

/* ---------------------------------------------------------------------- */
/* Procedure de fenetre principale                                         */
/* ---------------------------------------------------------------------- */

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_COMMAND: {
            int id = LOWORD(wParam);

            if (id == IDC_BROWSE) {
                char path[MAX_PATH] = { 0 };
                OPENFILENAMEA ofn;
                ZeroMemory(&ofn, sizeof(ofn));
                ofn.lStructSize = sizeof(ofn);
                ofn.hwndOwner = hwnd;
                ofn.lpstrFile = path;
                ofn.nMaxFile = sizeof(path);
                ofn.lpstrFilter = "Fichiers CSV\0*.csv\0Tous les fichiers\0*.*\0";
                ofn.nFilterIndex = 1;
                ofn.lpstrTitle = "Selectionner le fichier CSV des zones";
                ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

                if (GetOpenFileNameA(&ofn)) {
                    if (g_csv.lines) free_csv(&g_csv);
                    g_csv = load_csv(path);
                    strncpy(g_csvPath, path, sizeof(g_csvPath) - 1);
                    SetWindowText(g_hCsvLabel, path);
                    save_state(g_currentZone);
                    UpdateZoneLabel();
                    post_log("Fichier CSV charge : %d zones.", g_csv.count);
                }
            } else if (id == IDC_GOTO_BTN) {
                char buf[32];
                GetWindowText(g_hGotoEdit, buf, sizeof(buf));
                int target = atoi(buf);
                if (target < 1) {
                    MessageBox(hwnd, "Numero de zone invalide.", "Erreur", MB_OK | MB_ICONWARNING);
                } else {
                    g_currentZone = target;
                    save_state(g_currentZone);
                    UpdateZoneLabel();
                    post_log("Positionne sur la zone %d.", g_currentZone);
                }
            } else if (id == IDC_RESET_BTN) {
                g_currentZone = 1;
                save_state(g_currentZone);
                UpdateZoneLabel();
                post_log("Remise a zero effectuee (zone 1).");
            } else if (id == IDC_STARTSTOP_BTN) {
                if (!g_isRunning) {
                    if (SendMessage(g_hRadioManual, BM_GETCHECK, 0, 0) == BST_CHECKED) {
                        g_mode = MODE_MANUAL;
                        GetWindowText(g_hManualEdit, g_manualValueSnapshot, sizeof(g_manualValueSnapshot));
                        if (strlen(g_manualValueSnapshot) == 0) {
                            MessageBox(hwnd, "Entrez une valeur a ecrire avant de demarrer le mode manuel.",
                                       "Valeur manquante", MB_OK | MB_ICONWARNING);
                            return 0;
                        }
                    } else if (SendMessage(g_hRadioRead, BM_GETCHECK, 0, 0) == BST_CHECKED) {
                        g_mode = MODE_READ;
                    } else {
                        g_mode = MODE_AUTO;
                        if (g_csv.count == 0) {
                            MessageBox(hwnd, "Chargez d'abord un fichier CSV.", "Aucune donnee",
                                       MB_OK | MB_ICONWARNING);
                            return 0;
                        }
                    }

                    if (!g_readerReady) {
                        MessageBox(hwnd, "Aucun lecteur RFID detecte. Branchez le ACR122U et redemarrez l'application.",
                                   "Lecteur introuvable", MB_OK | MB_ICONERROR);
                        return 0;
                    }

                    g_stopRequested = FALSE;
                    g_isRunning = TRUE;
                    SetWindowText(g_hStartStopBtn, "Arreter");
                    SetControlsEnabled(FALSE);
                    CloseHandle(CreateThread(NULL, 0, WorkerThreadProc, NULL, 0, NULL));
                } else {
                    g_stopRequested = TRUE;
                    EnableWindow(g_hStartStopBtn, FALSE);
                    SetWindowText(g_hStartStopBtn, "Arret en cours...");
                }
            }
            return 0;
        }

        case WM_APP_LOG: {
            char *text = (char *)lParam;
            int len = GetWindowTextLength(g_hLogEdit);
            SendMessage(g_hLogEdit, EM_SETSEL, (WPARAM)len, (LPARAM)len);
            SendMessage(g_hLogEdit, EM_REPLACESEL, FALSE, (LPARAM)text);
            free(text);
            return 0;
        }

        case WM_APP_ZONE: {
            g_currentZone = (int)wParam;
            UpdateZoneLabel();
            return 0;
        }

        case WM_APP_DONE: {
            SetWindowText(g_hStartStopBtn, "Demarrer");
            EnableWindow(g_hStartStopBtn, TRUE);
            SetControlsEnabled(TRUE);
            return 0;
        }

        case WM_CLOSE:
            if (g_isRunning) {
                MessageBox(hwnd, "Arretez le mode en cours avant de fermer l'application.",
                           "Mode actif", MB_OK | MB_ICONWARNING);
                return 0;
            }
            DestroyWindow(hwnd);
            return 0;

        case WM_DESTROY:
            if (g_hContext) SCardReleaseContext(g_hContext);
            PostQuitMessage(0);
            return 0;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/* ---------------------------------------------------------------------- */
/* Creation des contrôles                                                  */
/* ---------------------------------------------------------------------- */

static HWND CreateCtrl(LPCSTR className, LPCSTR text, DWORD style,
                        int x, int y, int w, int h, HWND parent, int id) {
    HWND hwnd = CreateWindowEx(0, className, text, style, x, y, w, h,
                                parent, (HMENU)(INT_PTR)id, g_hInst, NULL);
    SendMessage(hwnd, WM_SETFONT, (WPARAM)g_hFont, TRUE);
    return hwnd;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    (void)hPrevInstance; (void)lpCmdLine;
    g_hInst = hInstance;
    g_hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

    WNDCLASS wc = { 0 };
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = "RFIDZonesGUIClass";
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    RegisterClass(&wc);

    g_hWndMain = CreateWindowEx(0, wc.lpszClassName, "Attribution de zones RFID",
                                 WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
                                 CW_USEDEFAULT, CW_USEDEFAULT, 660, 620,
                                 NULL, NULL, hInstance, NULL);

    /* --- Fichier CSV --- */
    CreateCtrl("STATIC", "Fichier CSV :", WS_CHILD | WS_VISIBLE,
               10, 14, 90, 20, g_hWndMain, 0);
    g_hCsvLabel = CreateCtrl("STATIC", "Aucun fichier charge", WS_CHILD | WS_VISIBLE | SS_LEFTNOWORDWRAP,
                              105, 14, 430, 20, g_hWndMain, IDC_CSV_LABEL);
    g_hBrowseBtn = CreateCtrl("BUTTON", "Parcourir...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                               545, 10, 95, 26, g_hWndMain, IDC_BROWSE);

    /* --- Zone courante --- */
    CreateCtrl("STATIC", "Zone actuelle :", WS_CHILD | WS_VISIBLE,
               10, 50, 100, 20, g_hWndMain, 0);
    g_hZoneLabel = CreateCtrl("STATIC", "- / -", WS_CHILD | WS_VISIBLE,
                               115, 50, 200, 20, g_hWndMain, IDC_ZONE_LABEL);

    /* --- Groupe de modes --- */
    CreateCtrl("BUTTON", "Mode", WS_CHILD | WS_VISIBLE | BS_GROUPBOX,
               10, 80, 630, 90, g_hWndMain, 0);
    g_hRadioAuto = CreateCtrl("BUTTON", "Automatique (fichier CSV)",
                               WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON | WS_GROUP,
                               25, 100, 250, 20, g_hWndMain, IDC_RADIO_AUTO);
    g_hRadioManual = CreateCtrl("BUTTON", "ecriture manuelle",
                                 WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON,
                                 25, 124, 250, 20, g_hWndMain, IDC_RADIO_MANUAL);
    g_hRadioRead = CreateCtrl("BUTTON", "Lecture de tag",
                               WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON,
                               25, 148, 250, 20, g_hWndMain, IDC_RADIO_READ);
    SendMessage(g_hRadioAuto, BM_SETCHECK, BST_CHECKED, 0);

    /* --- Contrôles mode automatique --- */
    CreateCtrl("STATIC", "Aller a la zone n\xB0 :", WS_CHILD | WS_VISIBLE,
               10, 185, 120, 20, g_hWndMain, 0);
    g_hGotoEdit = CreateCtrl("EDIT", "", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
                              135, 183, 70, 24, g_hWndMain, IDC_GOTO_EDIT);
    g_hGotoBtn = CreateCtrl("BUTTON", "Aller", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                             215, 182, 70, 26, g_hWndMain, IDC_GOTO_BTN);
    g_hResetBtn = CreateCtrl("BUTTON", "Remise a zero (zone 1)", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                              300, 182, 180, 26, g_hWndMain, IDC_RESET_BTN);

    /* --- Contrôle mode manuel --- */
    CreateCtrl("STATIC", "Valeur a ecrire (URL) :", WS_CHILD | WS_VISIBLE,
               10, 222, 150, 20, g_hWndMain, 0);
    g_hManualEdit = CreateCtrl("EDIT", "", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
                                10, 244, 630, 24, g_hWndMain, IDC_MANUAL_EDIT);

    /* --- Demarrer / Arreter --- */
    g_hStartStopBtn = CreateCtrl("BUTTON", "Demarrer", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                                  10, 280, 630, 34, g_hWndMain, IDC_STARTSTOP_BTN);

    /* --- Journal --- */
    CreateCtrl("STATIC", "Journal :", WS_CHILD | WS_VISIBLE,
               10, 326, 100, 20, g_hWndMain, 0);
    g_hLogEdit = CreateCtrl("EDIT", "",
                             WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL |
                             ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
                             10, 348, 630, 220, g_hWndMain, IDC_LOG_EDIT);

    /* --- Initialisation --- */
    g_readerReady = init_pcsc();
    g_currentZone = load_state();
    if (g_csvPath[0] != '\0') {
        g_csv = load_csv(g_csvPath);
        if (g_csv.count > 0) SetWindowText(g_hCsvLabel, g_csvPath);
    }
    UpdateZoneLabel();

    ShowWindow(g_hWndMain, nCmdShow);
    UpdateWindow(g_hWndMain);

    if (!g_readerReady) {
        MessageBox(g_hWndMain,
                   "Aucun lecteur RFID detecte au demarrage.\n"
                   "Branchez le ACR122U et redemarrez l'application.",
                   "Lecteur introuvable", MB_OK | MB_ICONWARNING);
    } else {
        post_log("Lecteur detecte : %s", g_readerName);
    }
    if (g_csv.count > 0) post_log("Fichier CSV recharge automatiquement : %d zones.", g_csv.count);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return 0;
}

/* ============================================================================
 * COMPILATION (Windows, avec w64devkit ou tout mingw-w64 complet)
 * ----------------------------------------------------------------------------
 *     gcc rfid_zones_gui.c -o rfid_zones_gui.exe -lwinscard -lcomdlg32 -mwindows
 *
 * -mwindows : lance l'application sans fenetre de console derriere.
 *
 * PRePARATION : identique a la version console — exportez la colonne I
 * (a partir de I2) dans un CSV, une valeur par ligne. Vous le chargez
 * ensuite depuis le bouton "Parcourir..." de l'application.
 *
 * state.txt (zone courante + chemin du CSV) est cree/mis a jour automatique-
 * ment dans le dossier où se trouve l'executable.
 * ==========================================================================*/
