#!/usr/bin/env python3
"""Verify image links in Obsidian markdown — configure MODE / TARGET / VAULT_ROOT below"""
import os, re
from urllib.parse import unquote

# ============================================================
# CONFIG — modify these to switch mode
# ============================================================
MODE = "dir"                # "single" = single file  |  "dir" = entire directory
# TARGET_FILE not used in dir mode
VAULT_ROOT  = r"/mnt/c/Users/Administrator/My Documents/Obsidian Vault/dai"

# ============================================================
IMAGE_EXTS = {'.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp', '.svg'}
MD_IMAGE_RE = re.compile(r'!\[([^\]]*)\]\((.+)\)')
WIKI_RE = re.compile(r'!\[\[[^\]]+\.(png|jpg|jpeg|gif|webp|bmp|svg)\]\]', re.I)


def verify_file(fp, vault_root=None):
    """Verify image links in one .md file. Returns (total, ok, broken, external, wiki_left, broken_list)."""
    if vault_root is None:
        vault_root = os.path.dirname(fp)

    with open(fp, 'r', encoding='utf-8') as f:
        content = f.read()

    wiki_left = len(WIKI_RE.findall(content))
    total, ok, broken, external = 0, 0, 0, 0
    broken_list = []

    for m in MD_IMAGE_RE.finditer(content):
        total += 1
        link = m.group(2)
        if link.startswith(('http://', 'https://', 'data:', 'file://')):
            external += 1
            continue

        decoded = unquote(link)
        abs_path = os.path.abspath(os.path.normpath(
            os.path.join(os.path.dirname(fp), decoded)))
        if os.path.isfile(abs_path):
            ok += 1
        else:
            broken += 1
            broken_list.append((os.path.relpath(fp, vault_root), link, abs_path))

    return total, ok, broken, external, wiki_left, broken_list


def verify_dir(vault_root):
    """Walk entire directory and verify all .md files."""
    total, ok, broken, external, wiki_left = 0, 0, 0, 0, 0
    all_broken = []

    for root, dirs, files in os.walk(vault_root):
        dirs[:] = [d for d in dirs if not d.startswith('.')]
        for fn in files:
            if not fn.endswith('.md'):
                continue
            fp = os.path.join(root, fn)
            t, o, b, e, w, bl = verify_file(fp, vault_root)
            total += t; ok += o; broken += b
            external += e; wiki_left += w
            all_broken.extend(bl)

    return total, ok, broken, external, wiki_left, all_broken


# ============================================================
# Main
# ============================================================
if __name__ == '__main__':
    if MODE == "single":
        fp = os.path.abspath(TARGET_FILE)
        if not os.path.isfile(fp):
            print(f"[ERROR] File not found: {fp}")
            exit(1)
        total, ok, broken, external, wiki_left, broken_list = verify_file(fp)
        print(f"\nFile:     {fp}")
        print(f"  Total:    {total}")
        print(f"  Working:  {ok}")
        print(f"  Broken:   {broken}")
        print(f"  External: {external}")
        print(f"  Wiki:     {wiki_left}")
        if broken_list:
            print("\n  Broken links:")
            for _, link, abs_path in broken_list:
                print(f"    ![]({link})")
                print(f"      -> not found: {abs_path}")
    else:
        vault = VAULT_ROOT
        print(f"Scanning: {vault}\n")
        total, ok, broken, external, wiki_left, all_broken = verify_dir(vault)

        if all_broken:
            print("Broken links (max 10 shown):")
            for rel_fp, link, _ in all_broken[:10]:
                print(f"  BROKEN in {rel_fp}: ![]({link})")
            if len(all_broken) > 10:
                print(f"  ... and {len(all_broken) - 10} more")

        print(f"\nTotal links: {total}")
        print(f"Working:     {ok}")
        print(f"Broken:      {broken}")
        print(f"External:     {external}")
        print(f"Wiki left:    {wiki_left}")
