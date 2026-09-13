import sys
import os

# ══════════════════════════════════════════════════════════════════════════════════════════
# NO CACHE. BOTH LINES, BEFORE EVERY OTHER IMPORT.
# Prevents __pycache__ generation when run or imported.
# ══════════════════════════════════════════════════════════════════════════════════════════
sys.dont_write_bytecode = True
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'

import shutil
import subprocess
from pathlib import Path

def get_downloads_directory():
    user_profile = os.environ.get('USERPROFILE')
    if user_profile:
        return Path(user_profile) / "Downloads"
    return Path.home() / "Downloads"

def is_binary(file_path: Path) -> bool:
    try:
        with open(file_path, 'rb') as f:
            chunk = f.read(2048)
            if b'\0' in chunk:
                return True
    except Exception:
        pass
    return False

def get_group_name(file_path: Path, project_root: Path) -> str:
    ext = file_path.suffix.lower()
    rel_path = str(file_path.relative_to(project_root)).lower()
    name = file_path.name.lower()

    # 00: Architecture, Specifications, and System Governance (Must be read first by AI)
    if rel_path.startswith("docs\\") or rel_path.startswith("docs/") or name in [
        'project_structure.txt', 'readme.md', 'license'
    ]:
        return "00_Specifications_and_Architecture"

    # 01: Core Orchestration, IPC & Service Backend
    if name in ['main.js', 'preload.js']:
        return "01_Core_Orchestration_and_Service"

    # 02: MAPI PowerShell Security Engine (COM STA automation, Heuristics, RunspacePool)
    if ext == '.ps1' or 'outlook-scanner' in name:
        return "02_MAPI_PowerShell_Security_Engine"

    # 03: UI Renderer & View Controllers (HTML, CSS, Vanilla JS modules)
    if ext in ['.html', '.htm', '.css'] or (ext == '.js' and name.startswith('ui_')):
        return "03_UI_Renderer_and_Views"

    # 04: Project & Build Configuration
    if ext in ['.json', '.json5', '.xml', '.config', '.ini', '.toml'] or name in [
        'package.json', 'package-lock.json', '.gitignore', '.gitattributes', '.editorconfig'
    ]:
        return "04_Project_and_Build_Configuration"

    # 05: Launchers & Developer Tools
    if 'developer_tools' in rel_path or ext in ['.bat', '.vbs', '.cmd'] or name in [
        'deskguard_for_microsoft_outlook.bat', 'build.bat', 'silent_launcher.vbs'
    ]:
        return "05_Launchers_and_Developer_Tools"

    # 06: Verification & Test Suites
    if 'test' in rel_path or 'tests' in rel_path or 'spec' in rel_path:
        return "06_Verification_and_Tests"

    # 07: Assets & Binary Stubs (Metadata and file size tracking)
    if ext in ['.ico', '.png', '.jpg', '.jpeg', '.gif', '.svg', '.dll', '.exe', '.bmp', '.webp']:
        return "07_Assets_and_Binaries"

    return "08_Misc"

MAX_CODE_FILES = 8  # 8 code bundles + 1 manifest + 1 tree file = 10 files max upload limit
MIN_PART_CHARS = 250_000  # Target balanced bundle sizes

def pack_into_capped_parts(items, max_parts=8):
    """Packs (rel_path, block) tuples across at most max_parts bundles."""
    if not items:
        return []

    total_chars = sum(len(b) for _, b in items)
    target_limit = max(MIN_PART_CHARS, (total_chars // max_parts) + 1)

    parts = []
    current_items = []
    used = 0

    for rel_path, block in items:
        if current_items and (used + len(block) > target_limit) and (len(parts) < max_parts - 1):
            parts.append(current_items)
            current_items = [(rel_path, block)]
            used = len(block)
        else:
            current_items.append((rel_path, block))
            used += len(block)

    if current_items:
        parts.append(current_items)

    return parts

def generate_project_manifest(project_root: Path, groups: dict, total_files: int) -> str:
    divider = "=" * 80
    lines = [
        divider,
        "DESKGUARD FOR MICROSOFT OUTLOOK - ARCHITECTURAL MANIFEST & SYSTEM CONTEXT",
        divider,
        f"Project Root:        {project_root}",
        "Target Platform:     Windows 10 / Windows 11 (x64 Mandate)",
        "Host Application:    Desktop Microsoft Outlook (MAPI COM STA Server)",
        "Runtime Stack:       Electron (Node.js) + PowerShell 5.1/7+ RunspacePool + Vanilla JS (DOM Virtualized)",
        "Architecture Style:  Single-Writer Background Service + IPC Named Pipe + STA COM Sidecar",
        "",
        "PRIMARY ARCHITECTURAL INVARIANTS (From project_structure.txt & README.md):",
        "  1. Single-Writer Multi-Process Architecture: Background Service (--service) is sole writer to storage.",
        "     Hosts authenticated Named Pipe (\\\\.\\pipe\\mos_service_<hash>). Client GUI holds in-memory read-only cache.",
        "  2. Outlook STA COM Synchronization: Microsoft Outlook runs as a Single-Threaded Apartment COM server.",
        "     All COM calls across processes serialize via Global\\MOS_Outlook_COM_Lock named mutex.",
        "     Deterministic RCW traversal (Release-Com / FinalReleaseComObject) prevents RPC handle table leaks.",
        "  3. Bounded Process Lifecycle & Standby Manager: Pre-spawn tasklist check guards against launching when Outlook",
        "     is closed. 3-failure circuit breaker with exponential backoff (4s/8s/16s). 15s passive standby polling at ~0.0% CPU.",
        "  4. Thread-Synchronized Stdout Packet Dispatch: All PowerShell stdout writes serialized via $Global:StdoutLock.",
        "     Emits atomic 4-byte LE binary frames. Node.js SafeIPCParser enforces strict length boundaries (<= 10MB).",
        "  5. In-Memory State Persistence Write-Back Buffer: Hot telemetry path uses in-memory Set<string> capped at 100k IDs",
        "     with O(1) ops. Debounced 5s atomic disk swap via .tmp and .bak snapshot rotation prevents event-loop starvation.",
        "  6. Canonical Forensic Snapshot Persistence: Message headers and body decoded once to plain UTF-8 on disk",
        "     (logs\\forensics\\<hash>.json). Async non-blocking retrieval (< 1ms latency) with self-healing legacy Base64 decoding.",
        "  7. Viewport DOM Virtualization & Strict Anti-XSS: Incident and duplicate lists cap active DOM elements to 50 rows.",
        "     Diff-based cell mutation sustains 60 FPS. All untrusted metadata strictly sanitized via escapeHtml and textContent.",
        "  8. Pre-Flight Survivor Validation Protocol: Before deleting duplicates, verifies survivor exists in Outlook,",
        "     is NOT in Deleted Items (Folder 3), and matches DNA fingerprint. Never deletes the last remaining copy of an email.",
        "",
        "SUBSYSTEM MAP & DIRECTORY HIERARCHY:",
        "  - main.js:               Single-Writer process coordinator, Named Pipe server/client, startup manager, tray lifecycle.",
        "  - preload.js:            Context bridge exposing window.securityApi securely to the renderer.",
        "  - outlook-scanner.ps1:   Dual-mode PowerShell engine (Worker & Scanner), Named Mutex STA sync, MAPI automation.",
        "  - index.html, style.css: Enterprise dark-mode application shell, status headers, navigation tabs, modal dialogs.",
        "  - ui_*.js:               Modular Vanilla JS controllers (ui_main, ui_list, ui_forensics, ui_duplicates, etc.).",
        "  - logs\\:                 Application logs (deskguard_outlook.log) and persistent forensics snapshots (logs\\forensics\\).",
        "  - developer_tools\\:      Automation, code extractors, bytecode sentinels, build tools.",
        "",
        "EXPORT GROUP METRICS (Ordered by AI ingestion priority):",
    ]
    for group_name in sorted(groups):
        count = len(groups[group_name])
        lines.append(f"  [{group_name}] -> {count:>3} file(s)")
    lines.extend([
        "",
        f"Total Tracked Items: {total_files}",
        "AI INSTRUCTION: Inspect groups in order (00 -> 01 -> 02 -> ...). Enforce invariants above during review and code generation.",
        divider,
        ""
    ])
    return "\n".join(lines)

def run_aggregator():
    project_root = Path(__file__).resolve().parent.parent
    download_dir = get_downloads_directory() / "deskguard_for_microsoft_outlook"

    if not download_dir.exists():
        download_dir.mkdir(parents=True, exist_ok=True)

    # Clean up previous text exports to avoid orphaned files
    for old_file in download_dir.glob("*.txt"):
        old_file.unlink(missing_ok=True)

    # developer_tools is deliberately excluded from ignored_dirs so AI has visibility into automation & scripts
    ignored_dirs = {
        '.git', 'bin', 'obj', '.vs', '.idea', 'node_modules', 'dist',
        'compile', 'compiled', 'old_code', 'artifacts', 'packages', 'testresults',
        'venv', '.venv', 'env', '.pytest_cache', '__pycache__', 'ran already', 'logs'
    }

    divider = "=" * 80

    # 1. Generate annotated directory tree with file sizes
    tree_file = download_dir / "00_file_structure.txt"
    with open(tree_file, "w", encoding="utf-8") as tf:
        tf.write(f"Directory Tree of: {project_root}\n")
        tf.write(f"{divider}\n")
        try:
            for root, dirs, files in os.walk(project_root):
                dirs[:] = sorted((d for d in dirs if d.lower() not in ignored_dirs), key=str.lower)
                rel = Path(root).relative_to(project_root)
                level = len(rel.parts) if rel.name else 0
                indent = ' ' * 4 * level
                folder_name = rel.name if rel.name else project_root.name
                tf.write(f"{indent}{folder_name}/\n")
                subindent = ' ' * 4 * (level + 1)
                for f in sorted(files, key=str.lower):
                    fp = Path(root) / f
                    try:
                        sz = fp.stat().st_size
                        sz_str = f"{sz / (1024 * 1024):.1f} MB" if sz >= 1024 * 1024 else f"{sz / 1024:.1f} KB"
                        tf.write(f"{subindent}{f:<50} [{sz_str:>9}]\n")
                    except Exception:
                        tf.write(f"{subindent}{f}\n")
        except Exception as e:
            tf.write(f"[ERROR GENERATING DIRECTORY TREE: {e}]")

    source_whitelist = {
        '.js', '.ps1', '.cmd', '.bat', '.vbs', '.py', '.json', '.json5', '.xml',
        '.txt', '.md', '.svg', '.manifest', '.config', '.editorconfig',
        '.gitignore', '.gitattributes', '.yaml', '.yml', '.css', '.html', '.htm',
        '.ini', '.toml'
    }

    binary_whitelist = {
        '.ico', '.png', '.jpg', '.jpeg', '.gif', '.dll', '.exe', '.bmp', '.webp'
    }

    print(f"Aggregating grouped code into {download_dir}...")

    processed_count = 0
    groups = {}

    for root, dirs, files in os.walk(project_root):
        dirs[:] = sorted((d for d in dirs if d.lower() not in ignored_dirs), key=str.lower)
        current_path = Path(root)

        for filename in sorted(files, key=str.lower):
            file_path = current_path / filename
            relative_path = file_path.relative_to(project_root)
            ext = file_path.suffix.lower()

            is_src = ext in source_whitelist or file_path.name.lower() in source_whitelist
            is_bin = ext in binary_whitelist

            if not is_src and not is_bin:
                continue

            if is_bin or is_binary(file_path):
                file_size = file_path.stat().st_size
                sz_str = f"{file_size / (1024 * 1024):.2f} MB" if file_size >= 1024 * 1024 else f"{file_size / 1024:.1f} KB"
                content_text = (
                    f"[BINARY / MEDIA ASSET OMITTED FROM TEXT EXPORT]\n"
                    f"Relative Path: {relative_path}\n"
                    f"File Size:     {sz_str} ({file_size:,} bytes)\n"
                    f"Type:          {ext.upper() if ext else 'Binary'} Asset\n"
                    f"Note:          Full path indexed in 00_file_structure.txt. Kept on disk.\n\n"
                )
                line_count = 0
                size_kb = file_size / 1024.0
            else:
                try:
                    with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                        content_text = f.read()
                    if not content_text.endswith("\n"):
                        content_text += "\n"
                    content_text += "\n"
                except Exception as e:
                    content_text = f"[ERROR READING FILE: {e}]\n\n"
                line_count = content_text.count("\n")
                size_kb = len(content_text) / 1024.0

            header = f"{divider}\nFILE: {relative_path}  ({line_count:,} lines, {size_kb:,.1f} KB)\n{divider}\n"
            block = header + content_text
            group_name = get_group_name(file_path, project_root)
            groups.setdefault(group_name, []).append((relative_path, block))
            processed_count += 1

    # 2. Generate and write high-level manifest
    manifest_file = download_dir / "00_PROJECT_MANIFEST.txt"
    manifest_content = generate_project_manifest(project_root, groups, processed_count)
    with open(manifest_file, "w", encoding="utf-8") as mf:
        mf.write(manifest_content)

    # 3. Assemble all blocks following the strict group priority order (00 -> 01 -> 02 -> ...)
    all_blocks = []
    for group in sorted(groups):
        for item in groups[group]:
            all_blocks.append(item)

    parts = pack_into_capped_parts(all_blocks, max_parts=MAX_CODE_FILES)
    count = len(parts)

    report_rows = []
    if manifest_file.exists():
        report_rows.append((manifest_file.name, manifest_file.stat().st_size))
    if tree_file.exists():
        report_rows.append((tree_file.name, tree_file.stat().st_size))

    for i, part_items in enumerate(parts, start=1):
        part_name = f"code_bundle_part{i:02d}of{count:02d}.txt" if count > 1 else "code_bundle.txt"
        toc = "\n".join(f"  - {rel_path}" for rel_path, _ in part_items)
        banner = (
            f"{divider}\n"
            f"CODE EXPORT BUNDLE - PART {i} OF {count}\n"
            f"Context Manifest: 00_PROJECT_MANIFEST.txt\n"
            f"Directory Tree:   00_file_structure.txt\n"
            f"Files in this bundle ({len(part_items)}):\n"
            f"{toc}\n"
            f"{divider}\n\n"
        )
        part_body = "".join(block for _, block in part_items)
        full_payload = banner + part_body
        with open(download_dir / part_name, "w", encoding="utf-8", errors="replace") as out:
            out.write(full_payload)
        report_rows.append((part_name, len(full_payload.encode('utf-8'))))

    print()
    print("UPLOAD REPORT - upload 00_PROJECT_MANIFEST.txt, 00_file_structure.txt, and bundle parts together")
    print("-" * 80)
    grand_total = 0
    for name, byte_size in report_rows:
        grand_total += byte_size
        print(f"  {name:<38} {byte_size / 1024.0:>10,.1f} KB   ~{byte_size // 4:>9,} tokens")
    print("-" * 80)
    print(f"  TOTAL: {len(report_rows)} files (<= 10 limit respected), {grand_total / 1024.0:,.1f} KB")
    print()
    print(f"Done! Aggregated {processed_count} files into {len(parts)} bundle(s) + 1 manifest + 1 structure file.")

def purge_bytecode_cache():
    tools_dir = Path(__file__).resolve().parent
    removed = []
    try:
        for cache_dir in tools_dir.rglob("__pycache__"):
            if cache_dir.is_dir():
                shutil.rmtree(cache_dir, ignore_errors=True)
                if not cache_dir.exists():
                    removed.append(cache_dir.name)
        for stray in list(tools_dir.rglob("*.pyc")) + list(tools_dir.rglob("*.pyo")):
            try:
                stray.unlink()
                removed.append(stray.name)
            except Exception:
                pass
    except Exception as exc:
        print(f"[!] Could not sweep bytecode cache: {exc}")
        return
    if removed:
        print(f"[*] Removed {len(removed)} bytecode cache item(s).")

if __name__ == '__main__':
    purge_bytecode_cache()
    run_aggregator()
