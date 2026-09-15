import os, sqlite3, uuid
from pathlib import Path
from datetime import datetime
from text_normalizer import NORMALIZATION_VERSION
SCHEMA_VERSION=1; BUILDER_VERSION="0.3.0"
class DatabaseManager:
    def __init__(self,base_dir):
        self.base_dir=Path(base_dir); self.db_path=self.base_dir/"mail.db"; self.build_path=self.base_dir/"mail_building.db"; self.old_path=self.base_dir/"mail_old.db"
    def _connect(self,path):
        c=sqlite3.connect(path,timeout=30); c.execute("PRAGMA foreign_keys=ON"); return c
    def _schema(self,c,fts=True):
        c.executescript('''
CREATE TABLE IF NOT EXISTS mails(id INTEGER PRIMARY KEY,store_id TEXT NOT NULL,entry_id TEXT NOT NULL,folder_entry_id TEXT NOT NULL DEFAULT '',folder_path TEXT NOT NULL DEFAULT '',message_class TEXT NOT NULL DEFAULT '',internet_message_id TEXT NOT NULL DEFAULT '',mail_datetime INTEGER NOT NULL DEFAULT 0,received_time INTEGER NOT NULL DEFAULT 0,sent_time INTEGER NOT NULL DEFAULT 0,last_modified INTEGER NOT NULL DEFAULT 0,subject TEXT NOT NULL DEFAULT '',body TEXT NOT NULL DEFAULT '',sender_name TEXT NOT NULL DEFAULT '',sender_address TEXT NOT NULL DEFAULT '',to_names TEXT NOT NULL DEFAULT '',to_addresses TEXT NOT NULL DEFAULT '',cc_names TEXT NOT NULL DEFAULT '',cc_addresses TEXT NOT NULL DEFAULT '',has_attachment INTEGER NOT NULL DEFAULT 0 CHECK(has_attachment IN(0,1)),UNIQUE(store_id,entry_id));
CREATE TABLE IF NOT EXISTS mail_search(mail_id INTEGER PRIMARY KEY,sender_text TEXT NOT NULL DEFAULT '',recipient_text TEXT NOT NULL DEFAULT '',subject TEXT NOT NULL DEFAULT '',body TEXT NOT NULL DEFAULT '');
CREATE TABLE IF NOT EXISTS sync_folders(store_id TEXT NOT NULL,folder_entry_id TEXT NOT NULL,folder_path TEXT NOT NULL DEFAULT '',is_active INTEGER NOT NULL DEFAULT 0,is_initialized INTEGER NOT NULL DEFAULT 0,last_incremental_sync INTEGER NOT NULL DEFAULT 0,last_full_sync INTEGER NOT NULL DEFAULT 0,last_sync_status TEXT NOT NULL DEFAULT '',last_error TEXT NOT NULL DEFAULT '',PRIMARY KEY(store_id,folder_entry_id));
CREATE TABLE IF NOT EXISTS db_info(key TEXT PRIMARY KEY,value TEXT NOT NULL DEFAULT '');
''')
        if fts: c.execute("CREATE VIRTUAL TABLE IF NOT EXISTS mail_fts USING fts5(sender_text,recipient_text,subject,body,tokenize='trigram')")
    def ensure_database(self):
        with self._connect(self.db_path) as c: self._schema(c)
    def create_build_database(self):
        self.cleanup_build_database(); c=self._connect(self.build_path); c.execute("PRAGMA journal_mode=OFF"); c.execute("PRAGMA synchronous=OFF"); self._schema(c,False)
        vals={"schema_version":SCHEMA_VERSION,"builder_version":BUILDER_VERSION,"db_generation":str(uuid.uuid4()),"revision":"1","normalization_version":NORMALIZATION_VERSION,"created_at":datetime.now().isoformat(timespec="seconds")}
        c.executemany("INSERT OR REPLACE INTO db_info(key,value) VALUES(?,?)",[(k,str(v)) for k,v in vals.items()]); c.commit(); c.close()
    def insert_build_batch(self,batch):
        ins=0
        with self._connect(self.build_path) as c:
            for m in batch:
                cur=c.execute('''INSERT OR IGNORE INTO mails(store_id,entry_id,folder_entry_id,folder_path,message_class,internet_message_id,mail_datetime,received_time,sent_time,last_modified,subject,body,sender_name,sender_address,to_names,to_addresses,cc_names,cc_addresses,has_attachment) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',(m["store_id"],m["entry_id"],m["folder_entry_id"],m["folder_path"],m["message_class"],m["internet_message_id"],m["mail_datetime"],m["received_time"],m["sent_time"],m["last_modified"],m["subject"],m["body"],m["sender_name"],m["sender_address"],m["to_names"],m["to_addresses"],m["cc_names"],m["cc_addresses"],m["has_attachment"]))
                if cur.rowcount:
                    mid=cur.lastrowid; c.execute("INSERT INTO mail_search VALUES(?,?,?,?,?)",(mid,m["search_sender"],m["search_recipient"],m["search_subject"],m["search_body"])); ins+=1
        return ins
    def mark_build_folder_success(self,store_id,folder_entry_id,folder_path,full_sync_time):
        with self._connect(self.build_path) as c: c.execute("INSERT OR REPLACE INTO sync_folders VALUES(?,?,?,?,?,?,?,?,?)",(store_id,folder_entry_id,folder_path,1,1,0,full_sync_time,"ok",""))
    def mark_build_folder_error(self,store_id,folder_entry_id,folder_path,error_text):
        with self._connect(self.build_path) as c: c.execute("INSERT OR REPLACE INTO sync_folders VALUES(?,?,?,?,?,?,?,?,?)",(store_id,folder_entry_id,folder_path,1,0,0,0,"error",str(error_text)[-4000:]))
    def finalize_build_database(self,mail_count=0):
        with self._connect(self.build_path) as c:
            c.executescript("CREATE INDEX IF NOT EXISTS idx_mails_datetime ON mails(mail_datetime DESC,id DESC); CREATE INDEX IF NOT EXISTS idx_mails_folder ON mails(store_id,folder_entry_id); CREATE INDEX IF NOT EXISTS idx_mails_modified ON mails(last_modified);")
            c.execute("CREATE VIRTUAL TABLE mail_fts USING fts5(sender_text,recipient_text,subject,body,tokenize='trigram')")
            c.execute("INSERT INTO mail_fts(rowid,sender_text,recipient_text,subject,body) SELECT mail_id,sender_text,recipient_text,subject,body FROM mail_search")
            c.execute("INSERT OR REPLACE INTO db_info VALUES('mail_count',?)",(str(mail_count),)); c.execute("INSERT OR REPLACE INTO db_info VALUES('last_updated_at',?)",(datetime.now().isoformat(timespec='seconds'),)); c.execute("ANALYZE")
    def validate_build_database(self): return self._validate(self.build_path)
    def _validate(self,path):
        with self._connect(path) as c:
            if c.execute("PRAGMA integrity_check").fetchone()[0] != "ok": raise RuntimeError("SQLite integrity_check failed")
            a=c.execute("SELECT count(*) FROM mails").fetchone()[0]; b=c.execute("SELECT count(*) FROM mail_search").fetchone()[0]; d=c.execute("SELECT count(*) FROM mail_fts").fetchone()[0]
            if not (a==b==d): raise RuntimeError(f"DB count mismatch mails={a} search={b} fts={d}")
        return True
    def swap_build_database(self):
        self.validate_build_database()
        if self.old_path.exists(): self.old_path.unlink()
        if self.db_path.exists(): os.replace(self.db_path,self.old_path)
        try: os.replace(self.build_path,self.db_path); self._validate(self.db_path)
        except Exception:
            if self.old_path.exists(): os.replace(self.old_path,self.db_path)
            raise
    def cleanup_build_database(self):
        for p in [self.build_path,Path(str(self.build_path)+"-wal"),Path(str(self.build_path)+"-shm")]:
            try:
                if p.exists(): p.unlink()
            except Exception: pass
    def _walk(self,nodes,parent=True):
        for n in nodes:
            own=bool(n.get("enabled",False)) and bool(n.get("available",True)) and not bool(n.get("hard_disabled",False)); eff=parent and own
            if n.get("node_type")=="folder" and eff: yield n
            yield from self._walk(n.get("children",[]),eff)
    def get_active_folders_from_settings(self,stores): return list(self._walk(stores,True))
    def sync_folder_settings(self,stores):
        if not self.db_path.exists(): return
        with self._connect(self.db_path) as c:
            def walk(nodes,parent=True):
                for n in nodes:
                    own=bool(n.get("enabled",False)) and bool(n.get("available",True)) and not bool(n.get("hard_disabled",False)); eff=parent and own
                    if n.get("node_type")=="folder": c.execute("INSERT INTO sync_folders(store_id,folder_entry_id,folder_path,is_active) VALUES(?,?,?,?) ON CONFLICT(store_id,folder_entry_id) DO UPDATE SET folder_path=excluded.folder_path,is_active=excluded.is_active",(n.get("store_id",""),n.get("folder_entry_id",""),n.get("folder_path",""),int(eff)))
                    walk(n.get("children",[]),eff)
            walk(stores)
    def get_status(self):
        if not self.db_path.exists(): return {"exists":False,"mail_count":0}
        try:
            with self._connect(self.db_path) as c: return {"exists":True,"mail_count":c.execute("SELECT count(*) FROM mails").fetchone()[0]}
        except Exception as e: return {"exists":True,"error":str(e),"mail_count":0}
