#!/usr/bin/env python3
"""
终极修复 v3:
1. 将所有 ![[image]] 转换为 ![image](正确相对路径)
2. 对链接路径中的特殊字符（空格、中文等）做 URL 编码
3. 修正所有损坏的链接

使用方式：修改下方 CONFIG 区域的 MODE / TARGET / VAULT_ROOT，直接运行即可。
"""
import os, re
from urllib.parse import quote, unquote

# ============================================================
# CONFIG — modify these to switch mode
# ============================================================
MODE = "dir"                # "single" = single file  |  "dir" = entire directory
# TARGET_FILE not used in dir mode
VAULT_ROOT  = r"/mnt/c/Users/Administrator/My Documents/Obsidian Vault/dai"

# ============================================================
IMAGE_EXTS = {'.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp', '.svg', '.ico', '.tiff', '.tif'}


# ============================================================
# Image index builder (shared)
# ============================================================
def build_image_index(vault_root):
    """Walk vault and build filename -> info dict."""
    print("Building image index...")
    image_index = {}
    for root, dirs, files in os.walk(vault_root):
        dirs[:] = [d for d in dirs if not d.startswith('.')]
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            if ext in IMAGE_EXTS:
                abs_path = os.path.join(root, f)
                rel_unix = os.path.relpath(abs_path, vault_root).replace('\\', '/')
                key = f.lower()
                if key not in image_index:
                    image_index[key] = {'rel': rel_unix, 'abs': abs_path}
    print(f"  Found {len(image_index)} unique images\n")
    return image_index


def encode_link_path(path):
    """URL-encode each path segment, keep / as separator"""
    parts = path.split('/')
    return '/'.join(quote(part) for part in parts)


def find_image(filename_str, filedir, image_index):
    """Find image by filename in index, return (encoded_relative_path, found_bool)"""
    fname = unquote(os.path.basename(filename_str)).strip()
    key = fname.lower()
    if key in image_index:
        info = image_index[key]
        try:
            raw_rel = os.path.relpath(info['abs'], filedir).replace('\\', '/')
        except ValueError:
            raw_rel = info['rel']
        return encode_link_path(raw_rel), True
    return encode_link_path(filename_str), False


# Regex patterns
WIKI_IMAGE_RE = re.compile(r'!\[\[([^\]|]+)(?:\|[^\]]*)?\]\]')
MD_IMAGE_RE = re.compile(r'!\[([^\]]*)\]\((.+)\)')


def process_file(filepath, image_index, stats):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    original = content
    filedir = os.path.dirname(filepath)

    # --- Convert wiki image links ---
    def wiki_replacer(m):
        wiki_path = m.group(1).strip()
        ext = os.path.splitext(wiki_path)[1].lower()
        if ext not in IMAGE_EXTS:
            return m.group(0)
        encoded, found = find_image(wiki_path, filedir, image_index)
        stats['wiki'] += 1
        if not found:
            print(f"    [WARN] Not found: {wiki_path}")
        return f'![{os.path.basename(wiki_path)}]({encoded})'

    content = WIKI_IMAGE_RE.sub(wiki_replacer, content)

    # --- Fix/re-encode existing MD links ---
    def md_replacer(m):
        alt, link = m.group(1), m.group(2)
        if link.startswith(('http://', 'https://', 'data:', 'file://')):
            return m.group(0)
        decoded = unquote(link)
        link_abs = os.path.abspath(os.path.normpath(os.path.join(filedir, decoded)))
        if os.path.isfile(link_abs):
            new_enc = encode_link_path(decoded)
            if new_enc != link:
                stats['fixed'] += 1
                return f'![{alt}]({new_enc})'
            return m.group(0)
        encoded, found = find_image(decoded, filedir, image_index)
        if found:
            stats['fixed'] += 1
            return f'![{alt or os.path.basename(decoded)}]({encoded})'
        return m.group(0)

    content = MD_IMAGE_RE.sub(md_replacer, content)

    if content != original:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    return False


def fix_dir(vault_root):
    """Walk entire directory and fix all .md files."""
    image_index = build_image_index(vault_root)
    stats = {'wiki': 0, 'fixed': 0, 'files': 0}

    print("Processing .md files...\n")
    for root_dir, dirs, files in os.walk(vault_root):
        dirs[:] = [d for d in dirs if not d.startswith('.')]
        for fn in sorted(files):
            if not fn.endswith('.md'):
                continue
            fp = os.path.join(root_dir, fn)
            if process_file(fp, image_index, stats):
                stats['files'] += 1
                rel = os.path.relpath(fp, vault_root)
                print(f"  [OK] {rel}")

    print(f"\n{'='*60}")
    print(f"DONE!")
    print(f"  Files modified: {stats['files']}")
    print(f"  Wiki->MD converted: {stats['wiki']}")
    print(f"  Links re-encoded/fixed: {stats['fixed']}")
    print(f"{'='*60}")


def fix_single(filepath):
    """Fix a single .md file.  image_index is built from the file's parent directory tree."""
    vault_root = os.path.dirname(filepath)
    image_index = build_image_index(vault_root)
    stats = {'wiki': 0, 'fixed': 0, 'files': 0}

    print(f"Processing: {filepath}\n")
    if process_file(filepath, image_index, stats):
        stats['files'] += 1
        print(f"  [OK] {filepath}")
    else:
        print("  (no changes needed)")

    print(f"\n{'='*60}")
    print(f"DONE!")
    print(f"  Files modified: {stats['files']}")
    print(f"  Wiki->MD converted: {stats['wiki']}")
    print(f"  Links re-encoded/fixed: {stats['fixed']}")
    print(f"{'='*60}")


# ============================================================
# Main
# ============================================================
if __name__ == '__main__':
    if MODE == "single":
        fp = os.path.abspath(TARGET_FILE)
        if not os.path.isfile(fp):
            print(f"[ERROR] File not found: {fp}")
            exit(1)
        fix_single(fp)
    else:
        fix_dir(VAULT_ROOT)
