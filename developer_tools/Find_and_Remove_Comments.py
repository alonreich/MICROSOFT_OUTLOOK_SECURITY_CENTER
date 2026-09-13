import sys
import os

# ══════════════════════════════════════════════════════════════════════════════════════════
# NO CACHE. BOTH LINES, BEFORE EVERY OTHER IMPORT.
#
# ⚠️ THE ENV VAR IS NOT OPTIONAL IN THIS FILE — IT IS THE ONLY ONE THAT WORKS HERE.
# This script uses multiprocessing.Pool. On Windows that SPAWNS fresh Python processes which
# re-import this module from scratch, and a child process does NOT inherit the parent's
# `sys.dont_write_bytecode` — it reads PYTHONDONTWRITEBYTECODE from the environment. With only
# the sys flag set, every worker process was free to write a __pycache__.
# `dont_write_bytecode` also has to sit ABOVE the imports: it governs imports made after it, so
# below the import block it protected nothing that had already been loaded.
# ══════════════════════════════════════════════════════════════════════════════════════════
sys.dont_write_bytecode = True
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'

import re
import ctypes
import shutil
from pathlib import Path
from multiprocessing import Pool, cpu_count

def get_downloads_directory():
    user_profile = os.environ.get('USERPROFILE')
    if user_profile:
        return Path(user_profile) / "Downloads"
    return Path.home() / "Downloads"

WORKING_DIRECTORY = Path(__file__).resolve().parent.parent
TOOL_OUTPUT_DIRECTORY = get_downloads_directory() / "deskguard_for_microsoft_outlook" / "developer_tools"
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
os.environ['PYTHONPYCACHEPREFIX'] = str(TOOL_OUTPUT_DIRECTORY / "pycache")
sys.pycache_prefix = os.environ['PYTHONPYCACHEPREFIX']

try:
    os.chdir(WORKING_DIRECTORY)
except Exception as e:
    print(f"Failed to change working directory: {e}")
    sys.exit(1)

RED = '\033[48;5;52m\033[97;1m'
GREEN = '\033[48;5;22m\033[97m'
CYAN = '\033[96m'
YELLOW = '\033[93m'
RESET = '\033[0m'

def center_console():
    try:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
        user32 = ctypes.windll.user32
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if not hwnd:
            return
        rect = ctypes.Structure()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
    except Exception:
        pass

EXCLUDE_FOLDERS = [
    '.git', 'node_modules', 'dist', 'out', 'release-build', 'ran already',
    'logs', '__pycache__', 'developer_tools', 'bin', 'obj', '.vs'
]
EXCLUDE_FILES = ['package-lock.json']
TARGET_EXTS = ['.js', '.ps1', '.css', '.html']

JS_REGEX = re.compile(
    r'^\s*(?:(?:export|default|async|static)\s+)*(?:function\s+([a-zA-Z0-9_$]+)|class\s+([a-zA-Z0-9_$]+)|(?:const|let|var)\s+([a-zA-Z0-9_$]+)\s*=\s*(?:async\s*)?(?:\([^)]*\)|[a-zA-Z0-9_$]+)\s*=>|window\.([a-zA-Z0-9_$]+)\s*=)'
)
PS_REGEX = re.compile(r'^\s*(?:function|filter|workflow)\s+([a-zA-Z0-9_\-:]+)', re.IGNORECASE)

def get_target_files(root_dir):
    targets = []
    for root, dirs, files in os.walk(root_dir):
        dirs[:] = [d for d in dirs if d.lower() not in EXCLUDE_FOLDERS]
        for file in files:
            if file in EXCLUDE_FILES:
                continue
            _, ext = os.path.splitext(file)
            if ext.lower() in TARGET_EXTS:
                targets.append(os.path.join(root, file))
    return sorted(targets, key=lambda p: p.lower())

def display_path(filepath):
    try:
        return str(Path(filepath).resolve().relative_to(WORKING_DIRECTORY))
    except Exception:
        return str(filepath)

def analyze_js_css_comments(filepath, is_css=False):
    try:
        with open(filepath, 'r', encoding='utf-8-sig', errors='replace') as f:
            lines = f.readlines()
        
        actions = {}
        in_block_comment = False
        in_template = False
        
        for i, line in enumerate(lines):
            stripped = line.strip()
            nl = line[len(line.rstrip('\r\n')):] or '\n'
            
            if in_block_comment:
                end_idx = line.find('*/')
                if end_idx != -1:
                    in_block_comment = False
                    after = line[end_idx + 2:]
                    if after.strip():
                        actions[i] = {
                            'action': 'EDIT',
                            'type': 'BLOCK COMMENT',
                            'line': i + 1,
                            'content': line[:end_idx + 2].strip(),
                            'new_content': after.lstrip(' ')
                        }
                    else:
                        actions[i] = {
                            'action': 'DELETE',
                            'type': 'BLOCK COMMENT',
                            'line': i + 1,
                            'content': stripped
                        }
                else:
                    actions[i] = {
                        'action': 'DELETE',
                        'type': 'BLOCK COMMENT',
                        'line': i + 1,
                        'content': stripped
                    }
                continue

            out_chars = []
            j = 0
            n = len(line)
            line_comments = []
            
            while j < n:
                c = line[j]
                
                if in_template:
                    if c == '\\':
                        out_chars.append(c)
                        if j + 1 < n:
                            out_chars.append(line[j + 1])
                            j += 2
                        else:
                            j += 1
                        continue
                    if c == '`':
                        in_template = False
                        out_chars.append(c)
                        j += 1
                        continue
                    out_chars.append(c)
                    j += 1
                    continue
                
                if c == "'":
                    out_chars.append(c)
                    j += 1
                    while j < n:
                        sc = line[j]
                        if sc == '\\':
                            out_chars.append(sc)
                            if j + 1 < n:
                                out_chars.append(line[j + 1])
                                j += 2
                            else:
                                j += 1
                            continue
                        out_chars.append(sc)
                        j += 1
                        if sc == "'":
                            break
                    continue
                
                if c == '"':
                    out_chars.append(c)
                    j += 1
                    while j < n:
                        sc = line[j]
                        if sc == '\\':
                            out_chars.append(sc)
                            if j + 1 < n:
                                out_chars.append(line[j + 1])
                                j += 2
                            else:
                                j += 1
                            continue
                        out_chars.append(sc)
                        j += 1
                        if sc == '"':
                            break
                    continue
                
                if not is_css and c == '`':
                    in_template = True
                    out_chars.append(c)
                    j += 1
                    continue
                
                if line[j:j+2] == '/*':
                    end_idx = line.find('*/', j + 2)
                    if end_idx != -1:
                        line_comments.append(('BLOCK COMMENT', line[j:end_idx + 2]))
                        j = end_idx + 2
                        continue
                    else:
                        line_comments.append(('BLOCK COMMENT', line[j:]))
                        in_block_comment = True
                        break
                
                if not is_css and line[j:j+2] == '//':
                    is_inline = bool(out_chars and "".join(out_chars).strip())
                    line_comments.append(('INLINE COMMENT' if is_inline else 'COMMENT', line[j:].rstrip('\r\n')))
                    break
                
                out_chars.append(c)
                j += 1
                
            if line_comments:
                cleaned_code = "".join(out_chars)
                if not cleaned_code.strip():
                    actions[i] = {
                        'action': 'DELETE',
                        'type': line_comments[0][0],
                        'line': i + 1,
                        'content': stripped
                    }
                else:
                    new_line = cleaned_code.rstrip() + nl
                    actions[i] = {
                        'action': 'EDIT',
                        'type': line_comments[0][0],
                        'line': i + 1,
                        'content': line_comments[0][1].strip(),
                        'new_content': new_line
                    }

        empty_count = 0
        for i, line in enumerate(lines):
            if i in actions:
                empty_count = 0
                continue
            if not line.strip():
                empty_count += 1
                if empty_count >= 3:
                    actions[i] = {
                        'action': 'DELETE',
                        'type': 'EXCESSIVE EMPTY',
                        'line': i + 1,
                        'content': '<Excessive Empty>'
                    }
            else:
                empty_count = 0

        return [v for k, v in sorted(actions.items())]
    except Exception:
        return []

def analyze_powershell_comments(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8-sig', errors='replace') as f:
            lines = f.readlines()
        
        actions = {}
        in_block_comment = False
        
        for i, line in enumerate(lines):
            stripped = line.strip()
            nl = line[len(line.rstrip('\r\n')):] or '\n'
            
            if stripped.lower().startswith('#requires'):
                continue
                
            if in_block_comment:
                end_idx = line.find('#>')
                if end_idx != -1:
                    in_block_comment = False
                    after = line[end_idx + 2:]
                    if after.strip():
                        actions[i] = {
                            'action': 'EDIT',
                            'type': 'BLOCK COMMENT',
                            'line': i + 1,
                            'content': line[:end_idx + 2].strip(),
                            'new_content': after.lstrip(' ')
                        }
                    else:
                        actions[i] = {
                            'action': 'DELETE',
                            'type': 'BLOCK COMMENT',
                            'line': i + 1,
                            'content': stripped
                        }
                else:
                    actions[i] = {
                        'action': 'DELETE',
                        'type': 'BLOCK COMMENT',
                        'line': i + 1,
                        'content': stripped
                    }
                continue

            out_chars = []
            j = 0
            n = len(line)
            line_comments = []
            
            while j < n:
                c = line[j]
                
                if c == "'":
                    out_chars.append(c)
                    j += 1
                    while j < n:
                        sc = line[j]
                        if sc == "'":
                            out_chars.append(sc)
                            j += 1
                            if j < n and line[j] == "'":
                                out_chars.append(line[j])
                                j += 1
                                continue
                            break
                        out_chars.append(sc)
                        j += 1
                    continue
                
                if c == '"':
                    out_chars.append(c)
                    j += 1
                    while j < n:
                        sc = line[j]
                        if sc == '`':
                            out_chars.append(sc)
                            if j + 1 < n:
                                out_chars.append(line[j + 1])
                                j += 2
                            else:
                                j += 1
                            continue
                        if sc == '"':
                            out_chars.append(sc)
                            j += 1
                            if j < n and line[j] == '"':
                                out_chars.append(line[j])
                                j += 1
                                continue
                            break
                        out_chars.append(sc)
                        j += 1
                    continue
                
                if line[j:j+2] == '<#':
                    end_idx = line.find('#>', j + 2)
                    if end_idx != -1:
                        line_comments.append(('BLOCK COMMENT', line[j:end_idx + 2]))
                        j = end_idx + 2
                        continue
                    else:
                        line_comments.append(('BLOCK COMMENT', line[j:]))
                        in_block_comment = True
                        break
                
                if c == '#':
                    if line[j:].strip().lower().startswith('#requires'):
                        out_chars.append(c)
                        j += 1
                        continue
                    is_inline = bool(out_chars and "".join(out_chars).strip())
                    line_comments.append(('INLINE COMMENT' if is_inline else 'COMMENT', line[j:].rstrip('\r\n')))
                    break
                
                out_chars.append(c)
                j += 1
                
            if line_comments:
                cleaned_code = "".join(out_chars)
                if not cleaned_code.strip():
                    actions[i] = {
                        'action': 'DELETE',
                        'type': line_comments[0][0],
                        'line': i + 1,
                        'content': stripped
                    }
                else:
                    new_line = cleaned_code.rstrip() + nl
                    actions[i] = {
                        'action': 'EDIT',
                        'type': line_comments[0][0],
                        'line': i + 1,
                        'content': line_comments[0][1].strip(),
                        'new_content': new_line
                    }

        empty_count = 0
        for i, line in enumerate(lines):
            if i in actions:
                empty_count = 0
                continue
            if not line.strip():
                empty_count += 1
                if empty_count >= 3:
                    actions[i] = {
                        'action': 'DELETE',
                        'type': 'EXCESSIVE EMPTY',
                        'line': i + 1,
                        'content': '<Excessive Empty>'
                    }
            else:
                empty_count = 0

        return [v for k, v in sorted(actions.items())]
    except Exception:
        return []

def analyze_html_comments(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8-sig', errors='replace') as f:
            lines = f.readlines()
        
        actions = {}
        in_block_comment = False
        
        for i, line in enumerate(lines):
            stripped = line.strip()
            nl = line[len(line.rstrip('\r\n')):] or '\n'
            
            if in_block_comment:
                end_idx = line.find('-->')
                if end_idx != -1:
                    in_block_comment = False
                    after = line[end_idx + 3:]
                    if after.strip():
                        actions[i] = {
                            'action': 'EDIT',
                            'type': 'HTML COMMENT',
                            'line': i + 1,
                            'content': line[:end_idx + 3].strip(),
                            'new_content': after.lstrip(' ')
                        }
                    else:
                        actions[i] = {
                            'action': 'DELETE',
                            'type': 'HTML COMMENT',
                            'line': i + 1,
                            'content': stripped
                        }
                else:
                    actions[i] = {
                        'action': 'DELETE',
                        'type': 'HTML COMMENT',
                        'line': i + 1,
                        'content': stripped
                    }
                continue

            out_chars = []
            j = 0
            n = len(line)
            line_comments = []
            
            while j < n:
                if line[j:j+4] == '<!--':
                    end_idx = line.find('-->', j + 4)
                    if end_idx != -1:
                        line_comments.append(('HTML COMMENT', line[j:end_idx + 3]))
                        j = end_idx + 3
                        continue
                    else:
                        line_comments.append(('HTML COMMENT', line[j:]))
                        in_block_comment = True
                        break
                out_chars.append(line[j])
                j += 1
                
            if line_comments:
                cleaned_code = "".join(out_chars)
                if not cleaned_code.strip():
                    actions[i] = {
                        'action': 'DELETE',
                        'type': line_comments[0][0],
                        'line': i + 1,
                        'content': stripped
                    }
                else:
                    new_line = cleaned_code.rstrip() + nl
                    actions[i] = {
                        'action': 'EDIT',
                        'type': line_comments[0][0],
                        'line': i + 1,
                        'content': line_comments[0][1].strip(),
                        'new_content': new_line
                    }

        empty_count = 0
        for i, line in enumerate(lines):
            if i in actions:
                empty_count = 0
                continue
            if not line.strip():
                empty_count += 1
                if empty_count >= 3:
                    actions[i] = {
                        'action': 'DELETE',
                        'type': 'EXCESSIVE EMPTY',
                        'line': i + 1,
                        'content': '<Excessive Empty>'
                    }
            else:
                empty_count = 0

        return [v for k, v in sorted(actions.items())]
    except Exception:
        return []

def analyze_comments(filepath):
    ext = Path(filepath).suffix.lower()
    if ext in ['.js', '.css']:
        return analyze_js_css_comments(filepath, is_css=(ext == '.css'))
    elif ext == '.ps1':
        return analyze_powershell_comments(filepath)
    elif ext == '.html':
        return analyze_html_comments(filepath)
    return []

def nuke_comments(filepath, items):
    try:
        print(f"\n{CYAN}Executing Cleanup: {display_path(filepath)}{RESET}")
        with open(filepath, 'r', encoding='utf-8-sig', errors='replace') as f:
            lines = f.readlines()
        action_map = {item['line'] - 1: item for item in items}
        with open(filepath, 'w', encoding='utf-8-sig') as f:
            for i, line in enumerate(lines):
                if i in action_map:
                    act = action_map[i]
                    print("-" * 60)
                    print(f"Line {act['line']}: {act['type']}")
                    print(f"{RED}- {line.rstrip()}{RESET}")
                    if act['action'] == 'EDIT':
                        print(f"{GREEN}+ {act['new_content'].rstrip()}{RESET}")
                        f.write(act['new_content'])
                else:
                    f.write(line)
        return True
    except Exception as e:
        print(f"Error: {e}")
        return False

def check_syntax(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8-sig', errors='replace') as f:
            source = f.read()
        issues = []
        if '\t' in source:
            issues.append("Contains Tabs")
        if source.count('{') != source.count('}'):
            issues.append("Mismatched Braces")
        if source.count('(') != source.count(')'):
            issues.append("Mismatched Parentheses")
        return ", ".join(issues) if issues else None
    except Exception:
        return None

def analyze_duplicates(filepath):
    found = {}
    duplicates = []
    ext = Path(filepath).suffix.lower()
    base = os.path.basename(filepath)
    
    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            for i, line in enumerate(f, 1):
                stripped = line.strip()
                if not stripped:
                    continue
                if ext in ['.js', '.css']:
                    if stripped.startswith('//') or stripped.startswith('/*') or stripped.startswith('*'):
                        continue
                    m = JS_REGEX.match(line)
                    if m:
                        name = next((g for g in m.groups() if g), None)
                        if name:
                            key = (base, 'JS_SCOPE', 'function/class', name)
                            found.setdefault(key, []).append(i)
                elif ext == '.ps1':
                    if stripped.startswith('#'):
                        continue
                    m = PS_REGEX.match(line)
                    if m:
                        name = m.group(1)
                        key = (base, 'PS_SCOPE', 'function', name)
                        found.setdefault(key, []).append(i)
                        
        for (f_name, scope, kind, sig), lines in found.items():
            if len(lines) > 1:
                duplicates.append([f_name, scope, kind, sig, ", ".join(map(str, lines))])
    except Exception:
        pass
    return duplicates

def print_table(title, data, headers):
    if not data:
        print(f"\n{title}: No issues found.")
        return
    print(f"\n{title}")
    widths = [len(h) for h in headers]
    for row in data:
        for i, val in enumerate(row):
            widths[i] = max(widths[i], len(str(val)))
    widths = [w + 2 for w in widths]
    h_str = " | ".join(f"{h:^{w}}" for h, w in zip(headers, widths))
    print("-" * len(h_str))
    print(h_str)
    print("-" * len(h_str))
    for row in data:
        print(" | ".join(f"{str(val):<{w}}" for val, w in zip(row, widths)))
    print("-" * len(h_str))

def purge_bytecode_cache():
    """
    Removes any __pycache__ / .pyc this tool's folder has accumulated.
    Scoped to developer_tools ONLY.
    """
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

def main():
    purge_bytecode_cache()
    center_console()
    os.system('title DeskGuard for Microsoft Outlook Advanced Code Cleaner')
    print(f"{CYAN}--- DESKGUARD FOR MICROSOFT OUTLOOK ADVANCED CODE CLEANER ---{RESET}")
    print("Target Directory: .")
    
    files = get_target_files(WORKING_DIRECTORY)
    print(f"Analyzing {len(files)} files...")
    
    junk_data = []
    files_with_junk = {}
    
    for idx, f in enumerate(files):
        if idx % 20 == 0:
            print(f"  Scanning for junk... {idx}/{len(files)}", end='\r')
        items = analyze_comments(f)
        if items:
            files_with_junk[f] = items
            for item in items:
                content_display = item['content'][:50] + '..' if len(item['content']) > 50 else item['content']
                junk_data.append([os.path.basename(f), item['line'], item['type'], content_display])

    print("\n" + "=" * 80)
    print(f"{YELLOW}STEP 1: REVIEW COMMENTS & UNNECESSARY EMPTY LINES{RESET}")
    print("=" * 80)
    
    if junk_data:
        print_table("TABLE 1: IDENTIFIED JUNK (COMMENTS & OUTRAGEOUS EMPTY LINES)", junk_data, ["File", "Line", "Type", "Content Preview"])
        print(f"\n{YELLOW}WARNING: This action will permanently remove all items listed above.{RESET}")
        try:
            if '-y' in sys.argv or '--yes' in sys.argv:
                q = 'Y'
            elif '--dry-run' in sys.argv:
                q = 'N'
            else:
                q = input(">>> Do you approve the removal of these comments/empty lines? (Y/N): ").strip().upper()
        except EOFError:
            q = 'N'
            
        if q == 'Y':
            for f, items in files_with_junk.items():
                nuke_comments(f, items)
            print(f"\n{GREEN}Cleanup complete.{RESET}")
        else:
            print(f"\n{CYAN}Cleanup cancelled by user.{RESET}")
    else:
        print("No comments or unnecessary empty lines found.")

    print("\n" + "=" * 80)
    print(f"{YELLOW}STEP 2: SYSTEM ANALYSIS (SYNTAX & DUPLICATES){RESET}")
    print("=" * 80)
    
    syntax_data = []
    for f in files:
        err = check_syntax(f)
        if err:
            syntax_data.append([os.path.basename(f), err])
    print_table("TABLE 2: SYNTAX & INDENTATION WARNINGS", syntax_data, ["File", "Issue"])

    all_dupes = []
    with Pool(processes=cpu_count()) as pool:
        results = pool.map(analyze_duplicates, files)
    for res in results:
        all_dupes.extend(res)
    print_table("TABLE 3: SCOPE-AWARE DUPLICATES (REPORT ONLY)", all_dupes, ["File", "Scope", "Type", "Signature", "Lines"])

    purge_bytecode_cache()
    print(f"\n{CYAN}Done.{RESET}")

if __name__ == "__main__":
    main()
