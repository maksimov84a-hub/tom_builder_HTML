# -*- coding: utf-8 -*-
"""
Bootstrap: download wheels directly from PyPI and unpack into ./vendor
without pip. Targets Python 3.12 (cp312) for binary wheels.

Usage: python bootstrap_deps.py [py312|py314]
"""
import json, os, sys, urllib.request, zipfile, shutil

PROJ = os.path.dirname(os.path.abspath(__file__))
WHEELS = os.path.join(PROJ, "wheels")
VENDOR = os.path.join(PROJ, "vendor")
TAG = "312" if (sys.argv[1] if len(sys.argv) > 1 else "312") == "312" else "314"
os.makedirs(WHEELS, exist_ok=True)
os.makedirs(VENDOR, exist_ok=True)

PACKAGES = ["PySide6", "python-docx", "PyMuPDF", "pywin32"]

def compatible(fn):
    low = fn.lower()
    if not low.endswith(".whl"):
        return False
    if "win_amd64" not in low:
        if "py3-none-any" in low:
            return True
        return False
    # win_amd64 binary wheel
    if TAG == "312":
        return "cp312" in low or "abi3" in low
    return True

def resolve(pkg):
    url = "https://pypi.org/pypi/%s/json" % pkg
    j = json.loads(urllib.request.urlopen(url, timeout=30).read())
    return j["info"], [u for u in j["urls"] if compatible(u["filename"])]

def download(u):
    fn = u["filename"]; dest = os.path.join(WHEELS, fn)
    if os.path.exists(dest): return dest
    data = urllib.request.urlopen(u["url"], timeout=180).read()
    open(dest,"wb").write(data); return dest

def unpack(wheel):
    with zipfile.ZipFile(wheel) as z:
        for n in z.namelist():
            if n.startswith("__pycache__"): continue
            t = os.path.join(VENDOR, n)
            os.makedirs(os.path.dirname(t), exist_ok=True)
            with z.open(n) as s, open(t,"wb") as d:
                shutil.copyfileobj(s, d)

def deps(info):
    out=[]
    for req in (info.get("requires_dist") or []):
        base = req.split(";")[0].strip()
        base = base.split(">=")[0].split("<=")[0].split("==")[0].split(">(")[0].strip()
        if not base: continue
        # unify names
        for cand in ["PySide6-Essentials","PySide6-Addons","shiboken6","python-docx","lxml",
                     "typing-extensions","lxml-html-clean","pywin32","PyMuPDF","packaging","Pillow"]:
            if base.lower().replace("_","-").startswith(cand.lower()):
                out.append(cand); break
    return dict.fromkeys(out)

def main():
    queue=list(PACKAGES); seen=set()
    while queue:
        pkg=queue.pop(0)
        if pkg in seen: continue
        seen.add(pkg)
        info, wheels = resolve(pkg)
        if not wheels:
            print("!! no compatible wheel for", pkg); continue
        w=None
        for c in wheels:
            if "win_amd64" in c["filename"].lower():
                w=c; break
        if w is None: w=wheels[0]
        print("download", w["filename"])
        unpack(download(w))
        for d in deps(info): queue.append(d)
    print("vendor files:", len(os.listdir(VENDOR)))

if __name__=="__main__": main()
