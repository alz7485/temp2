from datetime import datetime
import pythoncom, win32com.client, win32timezone  # noqa
from text_normalizer import normalize_search_text, prepare_outlook_text
from logger import log_info, log_warning, log_exception
OL_MAIL=43; FULL_BUILD_BATCH_SIZE=500
class BuildCancelled(Exception): pass
class MailBuilder:
    def __init__(self,db_manager,folders,cancel_callback=None,progress_callback=None):
        self.db_manager=db_manager; self.folders=folders; self.cancel_callback=cancel_callback; self.progress_callback=progress_callback
        self.total_inserted=0; self.total_skipped=0; self.total_field_warnings=0; self.folder_success_count=0; self.folder_error_count=0; self.conversation_history_stats={}
    def _check(self):
        if self.cancel_callback and self.cancel_callback(): raise BuildCancelled()
    def _emit(self,d):
        if self.progress_callback: self.progress_callback(d)
    def _warn(self,field): self.total_field_warnings+=1; log_warning(f"Mail field read failed field={field} warning_count={self.total_field_warnings}")
    def _safe(self,fn,default="",field=None):
        try:
            v=fn(); return default if v is None else v
        except Exception:
            if field: self._warn(field)
            return default
    def _ts(self,v):
        try: return int(v.timestamp()) if v else 0
        except Exception: self._warn("datetime"); return 0
    def _smtp_sender(self,item):
        name=str(self._safe(lambda:item.SenderName,"","SenderName")); addr=str(self._safe(lambda:item.SenderEmailAddress,"","SenderEmailAddress")); typ=str(self._safe(lambda:item.SenderEmailType,""))
        if typ.upper()=="SMTP": return name,addr
        try: ae=item.Sender
        except Exception: ae=None
        if ae:
            for meth in ("GetExchangeUser","GetExchangeDistributionList"):
                try:
                    obj=getattr(ae,meth)(); smtp=obj.PrimarySmtpAddress if obj else ""
                    if smtp: return name,str(smtp)
                except Exception: pass
            try:
                smtp=ae.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
                if smtp: return name,str(smtp)
            except Exception: pass
        return name,addr
    def _recipient_addr(self,r):
        try: ae=r.AddressEntry
        except Exception: return ""
        if not ae:return ""
        for meth in ("GetExchangeUser","GetExchangeDistributionList"):
            try:
                obj=getattr(ae,meth)(); smtp=obj.PrimarySmtpAddress if obj else ""
                if smtp:return str(smtp)
            except Exception:pass
        try:
            smtp=ae.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
            if smtp:return str(smtp)
        except Exception:pass
        return str(self._safe(lambda:ae.Address,""))
    def _recipients(self,item):
        tn=[];ta=[];cn=[];ca=[]
        try: rs=item.Recipients; count=int(rs.Count)
        except Exception: self._warn("Recipients"); return "","","",""
        for i in range(1,count+1):
            try:r=rs.Item(i); typ=int(r.Type); name=str(self._safe(lambda:r.Name,"")); addr=self._recipient_addr(r)
            except Exception:continue
            if typ==1: tn += [name] if name else []; ta += [addr] if addr else []
            elif typ==2: cn += [name] if name else []; ca += [addr] if addr else []
        return "; ".join(tn),"; ".join(ta),"; ".join(cn),"; ".join(ca)
    def _extract(self,item,fi):
        eid=str(self._safe(lambda:item.EntryID,""));
        if not eid:return None
        folder_name=fi.get("display_name","")
        subject=prepare_outlook_text(str(self._safe(lambda:item.Subject,"","Subject")),folder_name); body=prepare_outlook_text(str(self._safe(lambda:item.Body,"","Body")),folder_name)
        rn=self._ts(self._safe(lambda:item.ReceivedTime,None,"ReceivedTime")); sn=self._ts(self._safe(lambda:item.SentOn,None,"SentOn")); lm=self._ts(self._safe(lambda:item.LastModificationTime,None,"LastModificationTime"))
        try:sname,saddr=self._smtp_sender(item)
        except Exception:self._warn("Sender");sname=saddr=""
        ton,toa,ccn,cca=self._recipients(item)
        try:attach=1 if int(item.Attachments.Count)>0 else 0
        except Exception:self._warn("Attachments");attach=0
        try:imid=str(item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F") or "")
        except Exception:imid=""
        mc=str(self._safe(lambda:item.MessageClass,"","MessageClass")); sender=(sname+" "+saddr).strip(); recip=" ".join([ton,toa,ccn,cca]).strip()
        return {"store_id":fi["store_id"],"entry_id":eid,"folder_entry_id":fi["folder_entry_id"],"folder_path":fi.get("folder_path",""),"message_class":mc,"internet_message_id":imid,"mail_datetime":rn or sn or lm,"received_time":rn,"sent_time":sn,"last_modified":lm,"subject":subject,"body":body,"sender_name":sname,"sender_address":saddr,"to_names":ton,"to_addresses":toa,"cc_names":ccn,"cc_addresses":cca,"has_attachment":attach,"search_sender":normalize_search_text(sender),"search_recipient":normalize_search_text(recip),"search_subject":normalize_search_text(subject),"search_body":normalize_search_text(body)}
    def _conversation(self,name): return str(name or "").strip().casefold()=="会話履歴".casefold()
    def _record_type(self,item):
        try:c=int(item.Class)
        except Exception:c=-1
        try:m=str(item.MessageClass or "")
        except Exception:m="<unavailable>"
        self.conversation_history_stats[(c,m)]=self.conversation_history_stats.get((c,m),0)+1
    def build(self):
        pythoncom.CoInitialize(); started=int(datetime.now().timestamp()); log_info(f"Full build started folder_count={len(self.folders)}")
        try:
            ns=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI"); self.db_manager.create_build_database(); total=len(self.folders)
            for fn,fi in enumerate(self.folders,1):
                self._check(); path=fi.get("folder_path",""); self._emit({"type":"folder_start","folder_number":fn,"folder_total":total,"folder_path":path,"inserted":self.total_inserted,"skipped":self.total_skipped,"warnings":self.total_field_warnings})
                try:
                    folder=ns.GetFolderFromID(fi["folder_entry_id"],fi["store_id"]); items=folder.Items; count=int(items.Count); batch=[]; fins=0; fskip=0
                    for idx in range(1,count+1):
                        self._check()
                        try:item=items.Item(idx)
                        except Exception:self.total_skipped+=1;fskip+=1;continue
                        if self._conversation(fi.get("display_name","")): self._record_type(item)
                        try:cls=int(item.Class)
                        except Exception:cls=0
                        if cls!=OL_MAIL:continue
                        try:m=self._extract(item,fi)
                        except Exception:self.total_skipped+=1;fskip+=1;continue
                        if not m:self.total_skipped+=1;fskip+=1;continue
                        batch.append(m)
                        if len(batch)>=FULL_BUILD_BATCH_SIZE:
                            n=self.db_manager.insert_build_batch(batch); self.total_inserted+=n; fins+=n; batch.clear()
                        if idx==1 or idx%100==0 or idx==count:self._emit({"type":"mail_progress","folder_number":fn,"folder_total":total,"folder_path":path,"item_number":idx,"item_total":count,"inserted":self.total_inserted,"skipped":self.total_skipped,"warnings":self.total_field_warnings})
                    if batch:
                        n=self.db_manager.insert_build_batch(batch); self.total_inserted+=n; fins+=n
                    self.db_manager.mark_build_folder_success(fi["store_id"],fi["folder_entry_id"],path,started); self.folder_success_count+=1
                    self._emit({"type":"folder_complete","folder_number":fn,"folder_total":total,"folder_path":path,"folder_inserted":fins,"folder_skipped":fskip,"inserted":self.total_inserted,"skipped":self.total_skipped,"warnings":self.total_field_warnings})
                except BuildCancelled: raise
                except Exception as e:
                    self.folder_error_count+=1; log_exception(f"Folder build failed folder={fn}/{total}"); self.db_manager.mark_build_folder_error(fi["store_id"],fi["folder_entry_id"],path,f"{type(e).__name__}: {e}"); self._emit({"type":"folder_error","folder_number":fn,"folder_total":total,"folder_path":path,"inserted":self.total_inserted,"skipped":self.total_skipped,"warnings":self.total_field_warnings})
            for (c,m),n in sorted(self.conversation_history_stats.items()): log_info(f"Conversation history item type class={c} message_class={m} count={n}")
            if total and not self.folder_success_count: raise RuntimeError("対象フォルダを1つも正常に読み込めませんでした。")
            self._emit({"type":"finalizing","inserted":self.total_inserted,"skipped":self.total_skipped,"warnings":self.total_field_warnings}); self.db_manager.finalize_build_database(self.total_inserted)
            self._emit({"type":"validating","inserted":self.total_inserted,"skipped":self.total_skipped,"warnings":self.total_field_warnings}); self.db_manager.validate_build_database(); self._check()
            self._emit({"type":"swapping","inserted":self.total_inserted,"skipped":self.total_skipped,"warnings":self.total_field_warnings}); self.db_manager.swap_build_database()
            return {"mail_count":self.total_inserted,"skipped_count":self.total_skipped,"warning_count":self.total_field_warnings,"folder_success_count":self.folder_success_count,"folder_error_count":self.folder_error_count}
        except BuildCancelled:self.db_manager.cleanup_build_database();raise
        except Exception:log_exception("Full build failed");self.db_manager.cleanup_build_database();raise
        finally:pythoncom.CoUninitialize()
