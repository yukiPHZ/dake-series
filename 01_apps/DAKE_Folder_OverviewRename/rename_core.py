"""Folder-only operations. No file mutation, recursion, replacement or rollback replay."""
from __future__ import annotations
import os
import stat
import json
import re
import uuid
from dataclasses import dataclass, replace
from pathlib import Path

class SafetyError(ValueError):
    pass

def plain_path(path):
    path = Path(os.path.abspath(path))
    for part in (path, *path.parents):
        s = part.lstat()
        if stat.S_ISLNK(s.st_mode) or getattr(s, 'st_file_attributes', 0) & 0x400:
            raise SafetyError('link: ' + str(part))
    return path

def identity(path):
    s = plain_path(path).stat()
    if not stat.S_ISDIR(s.st_mode):
        raise SafetyError('not_directory: ' + str(path))
    return (s.st_dev, s.st_ino)

def fingerprint(path):
    # Immediate children only; includes subfolder identity, never their contents.
    with os.scandir(path) as entries:
        return tuple(sorted((e.name, *(lambda s: (s.st_dev, s.st_ino, s.st_size, s.st_mtime_ns, s.st_mode))(Path(e.path).lstat())) for e in entries))

@dataclass(frozen=True)
class Snapshot:
    path: Path
    ident: tuple
    parent_ident: tuple
    contents: tuple

    @classmethod
    def capture(cls, path):
        path = plain_path(path)
        ident = identity(path)
        contents = fingerprint(path)
        if identity(path) != ident:
            raise SafetyError('external_change: ' + str(path))
        return cls(path, ident, identity(path.parent), contents)

    def verify(self):
        if identity(self.path) != self.ident or identity(self.path.parent) != self.parent_ident or fingerprint(self.path) != self.contents:
            raise SafetyError('external_change: ' + str(self.path))

@dataclass(frozen=True)
class Change:
    snapshot: Snapshot
    name: str

    @property
    def destination(self):
        return self.snapshot.path.with_name(self.name)

def validate_name(name):
    reserved = {'CON','PRN','AUX','NUL','CONIN$','CONOUT$'} | {p+str(i) for p in ('COM','LPT') for i in range(1,10)} | {p+i for p in ('COM','LPT') for i in '¹²³'}
    if not name or name in ('.','..') or name[-1:] in (' ','.') or re.search(r'[<>:"/\\|?*\x00-\x1f\x7f]', name) or name.split('.')[0].upper() in reserved or len(name.encode('utf-16-le'))//2 > 255:
        raise SafetyError('invalid_name: ' + name)

def validate(changes):
    if not changes:
        raise SafetyError('empty')
    sources, targets = set(), set()
    for change in changes:
        validate_name(change.name)
        snap = change.snapshot
        snap.verify()
        source = str(snap.path).casefold()
        target = str(change.destination).casefold()
        if source in sources or target in targets:
            raise SafetyError('duplicate: ' + str(change.destination))
        sources.add(source)
        targets.add(target)
        # Conservatively reject swaps and case-only names; no temporary rename.
        with os.scandir(snap.path.parent) as entries:
            if any(e.name.casefold() == change.name.casefold() for e in entries):
                raise SafetyError('collision: ' + str(change.destination))
    paths = [c.snapshot.path for c in changes]
    if any(a in b.parents for a in paths for b in paths if a != b):
        raise SafetyError('nested_selection')

class Journal:
    def __init__(self, directory):
        directory = Path(directory)
        directory.mkdir(parents=True, exist_ok=True)
        self.path = directory / (uuid.uuid4().hex + '.jsonl')
    def write(self, record):
        with self.path.open('a', encoding='utf-8') as stream:
            stream.write(json.dumps(record, ensure_ascii=False) + '\n')
            stream.flush()
            os.fsync(stream.fileno())

def execute(changes, journal, move=None):
    """Stop on first failure; keep confirmed successes available for checked Undo."""
    validate(changes)
    results, undo = [], []
    for change in changes:
        source, destination = change.snapshot.path, change.destination
        moved = False
        try:
            validate([change])
            journal.write({'state':'intent','source':str(source),'destination':str(destination),'identity':change.snapshot.ident})
            # Windows rename is no-replace. Refuse other platforms for actual mutation.
            if os.name != 'nt':
                raise SafetyError('windows_only')
            if move is None:
                from win_rename import rename_checked
                rename_checked(change, validate)
            else:
                move(source, destination)
            moved = True
            renamed = replace(change.snapshot, path=destination)
            undo.append(Change(renamed, source.name))
            journal.write({'state':'success','source':str(source),'destination':str(destination)})
            results.append((str(source), str(destination), 'success'))
        except Exception as exc:
            results.append((str(source), str(destination), ('renamed_log_failed: ' if moved else 'failed: ') + str(exc)))
            try:
                journal.write({'state':'failure','source':str(source),'destination':str(destination),'renamed':moved,'error':str(exc)})
            except OSError:
                pass
            break
    return results, list(reversed(undo))
