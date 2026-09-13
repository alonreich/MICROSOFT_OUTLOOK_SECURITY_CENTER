# DeskGuard for Microsoft Outlook

> **Advanced Endpoint Email Security, Forensics & Redundancy Management Hub for Microsoft Outlook**

---

## 1. Overview

**DeskGuard for Microsoft Outlook** is a high-performance, client-side security sidecar engineered specifically for desktop Microsoft Outlook installations on Windows. Operating natively alongside Outlook, DeskGuard performs real-time forensic inspection, automated spear-phishing intercept, multi-stage heuristic analysis, and duplicate email identification—without introducing latency, deadlocks, or UI blocking into the user's interactive Outlook session.

---

## 2. Architecture & Design Principles

### A. Single-Writer Multi-Process Architecture
To eliminate race conditions, file tearing, and storage corruption across concurrent OS processes:
- **Background Service Process (`--service`):** Sole owner and authorized writer of application storage (`electron-store` on `config.json` and `data.json`). Hosts a dedicated Windows Named Pipe server (`\\.\pipe\mos_service_<hash>`) protected with per-session token authentication (`MOS_AUTH_<hash>`). Spawns, monitors, and orchestrates the underlying PowerShell analysis engines.
- **Client GUI Process:** Operates as a stateless client to the background service. Communicates exclusively over the authenticated Named Pipe IPC stream with correlated request-response IDs (`rid`), holding in-memory read-only caches to drive an instant, zero-latency interface. Never accesses storage files directly.

### B. Outlook STA COM Synchronization & Zero-Leak Lifecycle
Microsoft Outlook's COM server operates strictly within a Single-Threaded Apartment (STA) and cannot tolerate concurrent re-entrant method calls from multiple out-of-process clients:
- **Named Mutex Coordination (`Global\MOS_Outlook_COM_Lock`):** Serializes all COM calls across child worker processes and live monitors, eliminating `0x8001010A` (`RPC_E_SERVERCALL_RETRYLATER`) and `0x80010001` (`RPC_E_CALL_REJECTED`).
- **Progressive Exponential Backoff & Circuit Breaker:** Employs intervals `[50ms, 150ms, 500ms, 1000ms]` with randomized jitter `(10-50ms)`, automatically tripping a 30-second circuit breaker on continuous COM channel saturation to protect the host Outlook process.
- **Strict Zero-Leak RCW Hierarchy Traversal:** Hierarchical navigation helpers (`Get-TargetFolder-Safe`) explicitly release intermediate COM objects (`Parent`, `Store`, `Attachments`, `PropertyAccessor`, `AddressEntry`) via `[Runtime.InteropServices.Marshal]::FinalReleaseComObject()`, preventing MAPI RPC handle table saturation.
- **Sink Retention & Descriptor Traversal:** Event sink references (`$inbox`, `$items`) are anchored in `$Global:Watchers` to prevent garbage collection. Mailbox traversals use lightweight MAPI descriptors (`FolderId`, `StoreId`, `EntryID`) rather than holding open COM folder references.

### C. Bounded Process Lifecycle & Passive Standby Watcher
Eliminates runaway process respawns and CPU saturation when Outlook is closed or unresponsive:
- **Strict Pre-Spawn Verification:** `ensureOutlookRunning()` queries `tasklist` before spawning child processes. If Outlook is absent, transitions immediately to standby without process execution.
- **Bounded Retry Policy & Circuit Breaker:** Limits consecutive engine spawn failures to 3 (`MAX_ENGINE_FAILURES`) with exponential backoff (4s, 8s, 16s). Tripping the circuit breaker halts automatic respawns and notifies the UI.
- **Low-Frequency Passive Watcher:** Polls process readiness once every 15 seconds (`15000ms`) with ~0.0% CPU usage and zero child processes. Automatically re-launches the scanner and resets failure counters upon detecting `outlook.exe`.
- **Watchdog Circuit Breaker Integration:** 30-second watchdog timeouts route through the unified circuit breaker. All timers are deterministically disarmed on application shutdown.

### D. Thread-Synchronized Stdout Packet Dispatch & Safe IPC Framing
- **Process-Wide Monitor Serialization (`$Global:StdoutLock`):** Serializes all PowerShell standard output writes across workers, monitors, and concurrent `$RunspacePool` threads via `[System.Threading.Monitor]`.
- **Atomic Contiguous Binary Framing:** Packs length headers (4 bytes LE) and UTF-8 JSON payloads into single contiguous byte arrays before dispatch, preventing interleaved chunks or packet tearing under high concurrency.
- **Resilient SafeIPCParser:** Node.js stream parser enforces strict frame length validation (`0 < len <= 10MB`). Corrupt frames trigger immediate desync logging and buffer purges without stream desynchronization.

### E. In-Memory State Persistence & High-Throughput Write-Back Buffer
- **In-Memory Hash Set (`processedIdsCache`):** Tracks up to 100,000 processed message identifiers in an in-memory `Set<string>`. Lookups and insertions execute in O(1) time (< 0.001ms) on the hot scanning path with zero synchronous disk writes.
- **Unified Debounced Asynchronous Flush:** Coordinates incident statistics and processed IDs in a debounced background pass (every 5,000ms or 500 items). Serializes to `data.json.tmp` and performs an atomic swap via `fsPromises.rename()` with automated `.bak` snapshot rotation.
- **Zero Event-Loop Starvation:** Prevents main-thread blocking, preserving steady 5-second watchdog timer cadence without engine kill cascades. Synchronous flush fallbacks protect against data loss on shutdown.

### F. Canonical Forensic Snapshot Persistence & Non-Blocking Retrieval
- **Single-Phase Plain UTF-8 Persistence:** Email transport headers and message bodies are decoded once from Base64 upon ingestion and persisted as canonical plain UTF-8 JSON files in `logs\forensics\<sha256>.json`.
- **Asynchronous IPC Retrieval:** `ipcMain.handle('get-forensics')` utilizes `fsPromises.readFile` for sub-millisecond retrieval latency (< 1ms vs < 10ms ceiling), eliminating main-thread disk stalls.
- **Self-Healing Backward Compatibility:** Detects and decodes legacy Base64 snapshots on-the-fly, eliminating byte scrambling or character truncation while returning modern UTF-8 snapshots unaltered.
- **Error Containment:** Missing or malformed snapshots safely return `{ fullHeaders: 'Unavailable', body: 'Unavailable' }` without throwing unhandled exceptions.

### G. Virtualized DOM Rendering & Anti-XSS Sanitization
- **Viewport Virtualization:** Limits active DOM elements to a 50-row sliding viewport window in incident and duplicate tables, supporting up to 5,000 records per category without DOM thrashing or memory spikes.
- **Diff-Based State Reconciliation:** Selectively mutates modified table cells and status badges rather than destroying and recreating the DOM tree, maintaining smooth 60 FPS scrolling.
- **Strict Anti-XSS Sanitization:** Neutralizes adversarial payloads embedded in hostile email headers, subjects, and senders via centralized entity encoding (`window.escapeHtml`). Message bodies and headers are populated strictly via `element.textContent`.

### H. Integrated Action Dispatch & Remediation
- **Quarantine:** Relocates suspect items to **Junk Email** (`olFolderJunk` = 23) or **Deleted Items** (`olFolderDeletedItems` = 3) via `Robust-Move`.
- **Release:** Restores false positives from quarantine back to the primary **Inbox** (`olFolderInbox` = 6) and registers item fingerprints in `releasedFingerprints` to eliminate re-quarantine loops.
- **Delete:** Permanently purges targeted emails via MAPI `$item.Delete()` with immediate RCW cleanup.
- **Check-Existence:** Non-blocking item verification querying MAPI `GetItemFromID` to confirm record persistence without throwing COM exceptions.
- **DuplicateScan & ResetDuplicateStack:** Drives background mailbox redundancy scanning with pause, resume, and stack reset capabilities.
- **CloudVirusScan:** Deep threat intelligence auditing using DPAPI-decrypted VirusTotal credentials.

---

## 3. Defense Pipeline & Threat Intelligence

DeskGuard executes a 5-stage inspection pipeline on every processed email:

1. **Reputation Filtering:** Instant short-circuit verification against user-managed Whitelists and Blacklists (IPs, domains, sender addresses, and combo pairs).
2. **Transport & RFC Compliance:** Analyzes raw transport headers (`PR_TRANSPORT_MESSAGE_HEADERS`) for SPF alignment, DMARC evaluation, and suspicious relay indicators.
3. **MIME Structure & Forensic Analysis:** Inspects MIME parts, attachment extensions, executable signatures, and transport spam flags (`X-Spam-Flag`).
4. **Heuristic Engine:** Regex-based scanning against customizable corporate and threat keywords with weighted scoring.
5. **Advanced Threat Intelligence:**
   - **Level 0 (Off):** Local heuristics only. No external telemetry.
   - **Level 1 (Smart - Default):** Automatically generates SHA-256 hashes of body and attachments for suspicious items (Score < 65%) and audits against threat intelligence databases.
   - **Level 2 (Max Enforcement):** Mandatory cryptographic hashing and cloud intelligence verification for all incoming email components.

---

## 4. Verdict Routing & Actions

| Verdict | Score Threshold | Routing Action |
| :--- | :--- | :--- |
| **MALICIOUS** | Score ≤ 20% | Relocated to **Deleted Items / Quarantine** (`olFolderDeletedItems`) |
| **SPAM** | Score ≤ User Threshold | Relocated to **Junk Email** (`olFolderJunk`) |
| **SAFE** | Score > User Threshold | Retained in **Inbox** (`olFolderInbox`) |

---

## 5. Duplicate Email Discovery & Pre-Flight Survivor Validation

A native mailbox optimization engine designed to discover redundant items across all configured PST and OST stores:
- **Descriptor-Based Traversal:** Discovers and catalogues mail items using lightweight MAPI descriptors (`EntryID` and `StoreID`) rather than holding open COM folder references.
- **Multi-Factor DNA Matching:** Unique hash generation combining `Sender`, normalized `Subject`, ISO `Timestamp`, and a body sample fingerprint.
- **Survivor-Duplicate Coupling:** Duplicate records persistently store candidate metadata linked to definitive survivor identifiers (`entryId`, `survivorId`, `survivorFolder`, `survivorStore`, `dna`).
- **Pre-Flight Survivor Validation Protocol:**
  - Before moving or deleting any duplicate, verifies survivor existence in Outlook via `GetItemFromID`.
  - Confirms survivor is not in **Deleted Items** (`olFolderDeletedItems` = 3).
  - Recalculates survivor's DNA fingerprint to ensure it has not been altered; aborts cleanup if survivor is missing, deleted, or modified.
  - Verifies candidate duplicate is distinct from survivor (`$target.EntryID -ne $survivor.EntryID`) and matches cluster DNA.
- **Safe Isolation vs. Permanent Purge:**
  - **Move to Deleted Items (Safe - Default):** Relocates redundant copies to Deleted Items via `Robust-Move`, preserving read/unread flags.
  - **Permanent Purge:** Requires explicit keyword typing confirmation ("PURGE") before executing permanent deletion.
- **Non-Blocking Controls:** Interactive pause, resume, reset, and batch cleanup.

---

## 6. Execution & Deployment

### Quick Launch
Run the primary launcher batch script:
```cmd
.\DeskGuard_for_Microsoft_Outlook.bat
```

### Background Service Launch
Run the silent VBScript launcher for headless background service operation:
```cmd
wscript .\silent_launcher.vbs "node_modules\electron\dist\electron.exe"
```

### Packaging & Distribution
To package the application for standalone Windows x64 deployment:
```cmd
npm run package
```
Packaged binaries will be output to the `dist\` directory.

### Storage & Diagnostics
- **Configuration:** `%APPDATA%\DeskGuard for Microsoft Outlook\config.json` (with automatic `.bak` backups)
- **Audit Data & Telemetry:** `%APPDATA%\DeskGuard for Microsoft Outlook\data.json` (with `.bak` and `.tmp` atomic swaps)
- **Application Logs:** `logs\deskguard_outlook.log`
- **Forensic Snapshots:** `logs\forensics\<sha256>.json` (plain UTF-8 forensic headers and body snapshots)

### Developer Tools
- **`.\developer_tools\Pull_Code_files.py`**: Condenses and bundles the complete codebase and architecture manifest into <= 10 structured upload files for AI review.
- **`.\developer_tools\Find_and_Remove_Comments.py`**: Strips comments across `.js`, `.ps1`, `.css`, `.html` while preserving code, validates delimiter syntax, and detects duplicate declarations across CPU worker pools.
- **`.\developer_tools\native_mapi_poc\`**: Standalone native MAPI Table bulk-querying benchmark and proof-of-concept engine (>1,000 items/sec, <30MB RAM).
- **`.\docs\ARCHITECTURAL_MIGRATION_ROADMAP.md`**: Formal architectural design specification for migrating from Electron/PowerShell to native compiled sidecar (Rust/C++ + Tauri v2 / WebView2).

