"""Preview by default. Copy only explicitly confirmed legacy day mappings."""
import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import sqlite3
import uuid
from src.attendance_dates import parse_attendance_date

def digest(path):
    h=hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda:f.read(1024*1024),b""): h.update(chunk)
    return h.hexdigest()

def migrate_project(project_dir,apply=False,identity_db=None):
    root=Path(project_dir).resolve(strict=True)
    manifest=root/"attendance-period.json"
    data=json.loads(manifest.read_text(encoding="utf-8")) if manifest.exists() else {"version":1,"mappings":{}}
    if data.get("version")!=1 or not isinstance(data.get("mappings"),dict):
        raise ValueError("Invalid mapping manifest")
    report={"project":str(root),"apply":apply,"files":[],"unmapped":[],"pending_identities":[]}
    processed={}
    for folder in sorted(root.iterdir()):
        if not folder.is_dir() or not re.fullmatch(r"\d{1,2}",folder.name): continue
        if folder.is_symlink(): raise ValueError("Legacy source cannot be a symlink")
        entry=data["mappings"].get(folder.name,{})
        if not isinstance(entry,dict) or not entry.get("confirmed_by") or not entry.get("confirmed_at"):
            report["unmapped"].append(folder.name);continue
        day=parse_attendance_date(entry.get("date"))
        if day.day!=int(folder.name): raise ValueError("Mapped day does not match source folder")
        target=root/day.isoformat()
        if target.is_symlink() or not target.resolve().is_relative_to(root):
            raise ValueError("Target escapes project")
        hashes={}
        for source in sorted(folder.rglob("*")):
            if source.is_symlink(): raise ValueError("Symlinks require manual review")
            if not source.is_file(): continue
            relative=source.relative_to(folder).as_posix()
            destination=target/relative
            if not destination.resolve().is_relative_to(root) or any(p.is_symlink() for p in destination.parents if p!=root.parent):
                raise ValueError("Unsafe destination")
            sha=digest(source)
            if destination.exists() and (not destination.is_file() or digest(destination)!=sha):
                raise ValueError("Destination collision: "+str(destination))
            hashes[relative]=sha
            report["files"].append({"source":str(source),"target":str(destination),"sha256":sha})
        processed[folder.name]=(target,hashes)
    if identity_db:
        db=Path(identity_db).resolve(strict=True)
        with sqlite3.connect(db.as_uri()+"?mode=ro",uri=True) as c:
            report["pending_identities"]=[{"project_id":r[0],"employee_id":r[1],"source":r[2]} for r in c.execute(
                "SELECT project_id,employee_id,relative_path FROM legacy_sources l WHERE NOT EXISTS "
                "(SELECT 1 FROM memberships m WHERE m.project_id=l.project_id AND m.employee_id=l.employee_id)")]
    if not apply or not processed: return report
    token=datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S")+"_"+uuid.uuid4().hex[:8]
    shutil.copy2(manifest,root/("attendance-period."+token+".backup.json"))
    if identity_db:
        with sqlite3.connect(db.as_uri()+"?mode=ro",uri=True) as source:
            with sqlite3.connect(root/("identity."+token+".backup.sqlite3")) as backup: source.backup(backup)
    for target,_ in processed.values(): target.mkdir(parents=True,exist_ok=True)
    for item in report["files"]:
        source=Path(item["source"]);target=Path(item["target"])
        if digest(source)!=item["sha256"]: raise ValueError("Source changed during migration")
        target.parent.mkdir(parents=True,exist_ok=True)
        if not target.exists():
            with source.open("rb") as src,target.open("xb") as dst: shutil.copyfileobj(src,dst)
        if digest(target)!=item["sha256"]: raise ValueError("Copied file checksum mismatch")
    for folder,(target,hashes) in processed.items():
        data["mappings"][folder].update(migrated_to=target.name,source_hashes=hashes,migrated_at=token)
    # Audit file is separate so reruns never remove earlier evidence.
    audit=root/("attendance-migration."+token+".json")
    audit.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
    temp=root/("attendance-period."+token+".tmp")
    temp.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding="utf-8")
    os.replace(temp,manifest)
    return report

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument("project_dir")
    parser.add_argument("--apply",action="store_true")
    parser.add_argument("--identity-db")
    args=parser.parse_args()
    print(json.dumps(migrate_project(args.project_dir,args.apply,args.identity_db),ensure_ascii=False,indent=2))

if __name__=="__main__": main()
