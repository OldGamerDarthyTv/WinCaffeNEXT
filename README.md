# ☕ OGD WinCaffe NEXT v10.0.0

**Sistema Definitivo di Ottimizzazione Windows 11 per Gaming e Produttività**  
*Autore: OldGamerDarthy | #DarkPlayer84Tv Productions*  
*Community Discord: [Entra nel server ufficiale!](https://discord.gg/9Dvr7wkafC)*

---

## 🌟 Novità della versione 10.0.0 (The "Safe & Precise" Update)

Questa versione segna un salto generazionale per **OGD WinCaffe**, con un focus maniacale sulla stabilità, sulla pulizia del codice e sulla sicurezza. È stato fatto un refactoring profondo per garantire un'ottimizzazione del 200% superiore, senza rischi di ban dagli anti-cheat e con una compatibilità totale per l'ecosistema moderno (Windows 11 24H2 e 25H2+).

### 🚀 Miglioramenti Principali

*   🛡️ **100% Safe per Anti-Cheat (Network Zero-Touch)**
    Abbiamo rimosso completamente tutte le vecchie manipolazioni aggressive di rete e LAN. Questo garantisce che nessun anti-cheat rileverà anomalie legate ai pacchetti di rete, annullando il rischio di instabilità nei giochi competitivi.

*   💻 **Sostituzione WMI con CIM**
    L'intero core dello script è stato modernizzato. Abbiamo sostituito i vecchi e lenti comandi `Get-WmiObject` con le più veloci e sicure API `CIM`. Questo rende WinCaffe NEXT pienamente compatibile in modo nativo con **PowerShell 7+ (pwsh)**.

*   🛠️ **Integrazione ViVeTool Intelligente**
    L'integrazione di ViVeTool (per attivare feature nascoste o sperimentali di Windows) è ora gestita interamente **tramite chiavi di registro sicure (registry-only)**, aggirando del tutto l'uso di eseguibili esterni invasivi che potrebbero interferire con le sicurezze del sistema.

*   🌐 **.NET Auto-Updater Avanzato**
    Il gestore di dipendenze .NET è stato potenziato! Ora include la ricerca e l'aggiornamento automatico (tramite `winget`) di Runtime, Desktop, SDK e ASP.NET fino alla **versione 15**, includendo anche le Ultimissime Preview ufficiali di Microsoft.

*   🧠 **Profili CPU Dedicati (Intel vs AMD)**
    Nuovi menu dedicati per ottimizzazioni mirate al millimetro: un profilo prudente e intelligente per la gestione delle temperature su **AMD Ryzen / X3D**, e uno ad altissime prestazioni per le architetture **Intel Core / Core Ultra**.

*   🔧 **Fix Gestore OpenDyslexic**
    Risolto un fastidioso bug nella procedura d'installazione che impediva la corretta visualizzazione e interazione con il gestore del font OpenDyslexic per l'accessibilità opzionale.

*   🖥️ **Installatore Globale (`wincaffe`) Corretto**
    Risolti i fastidiosi crash istantanei della shell (la "finestra che si chiude sùbito") causati dai lanci tramite PowerShell 7 senza permessi di amministratore. Ora l'elevazione UAC funziona sempre in modo pulito e affidabile.

---

## ⚙️ Come Installare o Eseguire

OGD WinCaffe NEXT richiede i diritti di Amministratore per poter modificare i parametri di sistema in modo sicuro e per creare sempre i punti di ripristino.

1.  Apri **Windows Terminal** o **PowerShell** come Amministratore.
2.  Avvia lo script:
    ```powershell
    & "Percorso\Al\File\OGD_WinCaffe_10.0.0.ps1"
    ```
3.  Premi `1` per **Installare** lo script globalmente. 
4.  Da quel momento in poi, ti basterà digitare `wincaffe` in qualsiasi console o premere `Win+R` e digitare `wincaffe` per aprire il tuo hub di ottimizzazione.

## 🤝 Community & Supporto
Tutti i link sono stati aggiornati. Se hai dubbi, vuoi condividere il tuo preset, o hai bisogno di aiuto, entra nella nostra community Discord ufficiale: **[https://discord.gg/9Dvr7wkafC](https://discord.gg/9Dvr7wkafC)**

> *Attenzione: OGD WinCaffe crea sempre un punto di ripristino prima di applicare le modifiche. L'autore declina ogni responsabilità per eventuali malfunzionamenti. Usalo con coscienza e non esitare a usare la funzione di RESET inclusa.*
