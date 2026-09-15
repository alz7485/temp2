import pythoncom
import win32com.client
import win32timezone  # noqa

class ScanCancelled(Exception): pass

DEFAULT_OFF_NAMES={x.casefold() for x in ["Yammer のルート","ExternalContacts","ファイル","連絡先","予定表","ニュース フィード","クイック操作設定","タスク","同期の失敗","迷惑メール","下書き","Conversation Action Settings","メモ","RSS フィード","ジャーナル"]}
DELETED_NAMES={"削除済みアイテム","deleted items"}
NO_RECURSE={"同期の失敗","sync issues","rss フィード","rss feeds","検索フォルダー","search folders"}

class OutlookScanner:
    def __init__(self,cancel_callback=None,progress_callback=None):
        self.cancel_callback=cancel_callback; self.progress_callback=progress_callback
    def _cancel(self):
        if self.cancel_callback and self.cancel_callback(): raise ScanCancelled()
    def _emit(self,msg):
        if self.progress_callback: self.progress_callback(msg)
    def scan(self):
        pythoncom.CoInitialize()
        try:
            ns=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            result=[]
            for i in range(1, int(ns.Stores.Count)+1):
                self._cancel(); store=ns.Stores.Item(i); sid=str(store.StoreID); name=str(store.DisplayName or "")
                root=store.GetRootFolder(); children=[]
                self._walk(root, sid, "", children)
                result.append({"node_type":"store","display_name":name,"store_id":sid,"folder_entry_id":"","folder_path":name,"parent_folder_entry_id":"","enabled":True,"hard_disabled":False,"available":True,"is_new":False,"effective_enabled":True,"children":children})
            return result
        finally: pythoncom.CoUninitialize()
    def _walk(self,parent,store_id,parent_id,out):
        try: count=int(parent.Folders.Count)
        except Exception: return
        for i in range(1,count+1):
            self._cancel()
            try:
                f=parent.Folders.Item(i); name=str(f.Name or ""); eid=str(f.EntryID or ""); path=str(f.FolderPath or name)
            except Exception: continue
            low=name.strip().casefold(); deleted=low in {x.casefold() for x in DELETED_NAMES}
            enabled=(low not in DEFAULT_OFF_NAMES) and not deleted
            node={"node_type":"folder","display_name":name,"store_id":store_id,"folder_entry_id":eid,"folder_path":path,"parent_folder_entry_id":parent_id,"enabled":enabled,"hard_disabled":deleted,"available":True,"is_new":False,"effective_enabled":enabled,"children":[]}
            out.append(node); self._emit(path)
            if low not in {x.casefold() for x in NO_RECURSE}: self._walk(f,store_id,eid,node["children"])
