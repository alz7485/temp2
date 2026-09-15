from copy import deepcopy
from PySide6.QtCore import QObject,QThread,Qt,Signal
from PySide6.QtGui import QColor,QBrush,QCloseEvent
from PySide6.QtWidgets import QMainWindow,QWidget,QVBoxLayout,QHBoxLayout,QLabel,QPushButton,QTreeWidget,QTreeWidgetItem,QProgressBar,QMessageBox
from outlook_scanner import OutlookScanner,ScanCancelled
from db_manager import DatabaseManager
from mail_builder import MailBuilder,BuildCancelled

class ScanWorker(QObject):
    finished=Signal(object); failed=Signal(str); cancelled=Signal(); progress=Signal(str)
    def __init__(self):super().__init__();self.cancel=False
    def run(self):
        try:self.finished.emit(OutlookScanner(lambda:self.cancel,self.progress.emit).scan())
        except ScanCancelled:self.cancelled.emit()
        except Exception as e:self.failed.emit(f"{type(e).__name__}: {e}")
class BuildWorker(QObject):
    finished=Signal(object); failed=Signal(str); cancelled=Signal(); progress=Signal(object)
    def __init__(self,db,folders):super().__init__();self.cancel=False;self.db=db;self.folders=folders
    def run(self):
        try:self.finished.emit(MailBuilder(self.db,self.folders,lambda:self.cancel,self.progress.emit).build())
        except BuildCancelled:self.cancelled.emit()
        except Exception as e:self.failed.emit(f"{type(e).__name__}: {e}")

class MainWindow(QMainWindow):
    def __init__(self,settings_manager):
        super().__init__(); self.sm=settings_manager; self.db=DatabaseManager(settings_manager.base_dir); self.stores=[]; self.dirty=False; self.thread=None; self.worker=None
        self.setWindowTitle("Outlook DB Builder - 設定"); self.resize(900,700); self._ui(); self._load()
    def _ui(self):
        w=QWidget();v=QVBoxLayout(w);h=QHBoxLayout();h.addWidget(QLabel("Outlookフォルダ"));h.addStretch();self.rescan=QPushButton("Outlookを再解析");self.rescan.clicked.connect(self.start_scan);h.addWidget(self.rescan);v.addLayout(h)
        v.addWidget(QLabel("チェックしたフォルダを検索DBの対象にします。親OFF時は子のチェック状態を保持します。"))
        self.tree=QTreeWidget();self.tree.setHeaderLabels(["フォルダ"]);self.tree.itemChanged.connect(self._changed);v.addWidget(self.tree,1)
        self.db_label=QLabel();v.addWidget(self.db_label);self.status=QLabel("待機中");self.status.setMinimumHeight(90);v.addWidget(self.status);self.progress=QProgressBar();v.addWidget(self.progress)
        b=QHBoxLayout();self.cancel=QPushButton("キャンセル");self.cancel.setEnabled(False);self.cancel.clicked.connect(self._cancel);self.save=QPushButton("設定保存");self.save.clicked.connect(self._save);self.build=QPushButton("フルビルド");self.build.clicked.connect(self.start_build);close=QPushButton("閉じる");close.clicked.connect(self.close)
        b.addWidget(self.cancel);b.addStretch();b.addWidget(self.save);b.addWidget(self.build);b.addWidget(close);v.addLayout(b);self.setCentralWidget(w);self._dbstatus()
    def _load(self):
        try:self.stores=self.sm.load().get("stores",[])
        except Exception as e:QMessageBox.warning(self,"設定",str(e));self.stores=[]
        if self.stores:self._populate()
        else:self.start_scan()
    def _key(self,n):return (n.get("store_id",""),n.get("folder_entry_id",""),n.get("node_type",""))
    def _flatten(self,nodes):
        d={}
        for n in nodes:d[self._key(n)]=n;d.update(self._flatten(n.get("children",[])))
        return d
    def _merge(self,new):
        old=self._flatten(self.stores)
        def walk(nodes):
            for n in nodes:
                o=old.get(self._key(n))
                if o:n["enabled"]=o.get("enabled",n.get("enabled",False));n["is_new"]=False
                else:n["is_new"]=True;n["enabled"]=False
                walk(n.get("children",[]))
        walk(new);return new
    def _populate(self):
        self.tree.blockSignals(True);self.tree.clear()
        for n in self.stores:self._add(None,n,True)
        self.tree.blockSignals(False);self.tree.expandToDepth(1);self._refresh_effective()
    def _add(self,parent,n,parent_eff):
        item=QTreeWidgetItem([n.get("display_name","")+("  [NEW]" if n.get("is_new") else "")]);item.setData(0,Qt.ItemDataRole.UserRole,n);item.setFlags(item.flags()|Qt.ItemFlag.ItemIsUserCheckable);item.setCheckState(0,Qt.CheckState.Checked if n.get("enabled") else Qt.CheckState.Unchecked)
        (parent.addChild(item) if parent else self.tree.addTopLevelItem(item))
        for c in n.get("children",[]):self._add(item,c,parent_eff)
    def _changed(self,item,col):
        n=item.data(0,Qt.ItemDataRole.UserRole);n["enabled"]=item.checkState(0)==Qt.CheckState.Checked;self.dirty=True;self._title();self._refresh_effective()
    def _refresh_effective(self):
        self.tree.blockSignals(True)
        def rec(item,parent_eff=True):
            n=item.data(0,Qt.ItemDataRole.UserRole);own=n.get("enabled",False) and n.get("available",True) and not n.get("hard_disabled",False);eff=parent_eff and own;n["effective_enabled"]=eff
            flags=item.flags()
            if n.get("hard_disabled") or not parent_eff:flags &= ~Qt.ItemFlag.ItemIsEnabled;item.setForeground(0,QBrush(QColor("#888888")))
            else:flags |= Qt.ItemFlag.ItemIsEnabled;item.setForeground(0,QBrush())
            item.setFlags(flags)
            for i in range(item.childCount()):rec(item.child(i),eff)
        for i in range(self.tree.topLevelItemCount()):rec(self.tree.topLevelItem(i),True)
        self.tree.blockSignals(False)
    def _title(self):self.setWindowTitle(("*" if self.dirty else "")+"Outlook DB Builder - 設定")
    def _save(self):
        self.sm.save(self.stores);self.db.sync_folder_settings(self.stores)
        def clear(nodes):
            for n in nodes:n["is_new"]=False;clear(n.get("children",[]))
        clear(self.stores);self.dirty=False;self._title();self._populate();self.status.setText("設定を保存しました。")
    def _busy(self,on):
        for x in (self.rescan,self.save,self.build,self.tree):x.setEnabled(not on)
        self.cancel.setEnabled(on)
    def start_scan(self):
        if self.thread:return
        self._busy(True);self.status.setText("Outlookフォルダを解析しています...");self.progress.setRange(0,0);self.thread=QThread();self.worker=ScanWorker();self.worker.moveToThread(self.thread);self.thread.started.connect(self.worker.run);self.worker.progress.connect(self.status.setText);self.worker.finished.connect(self._scan_ok);self.worker.failed.connect(self._fail);self.worker.cancelled.connect(self._cancelled);self.worker.finished.connect(self.thread.quit);self.worker.failed.connect(self.thread.quit);self.worker.cancelled.connect(self.thread.quit);self.thread.finished.connect(self._thread_done);self.thread.start()
    def _scan_ok(self,data):self.stores=self._merge(data);self.dirty=True;self._title();self._populate();self.status.setText("解析が完了しました。内容を確認して設定保存してください。")
    def start_build(self):
        if self.thread:return
        if self.dirty:self._save()
        folders=self.db.get_active_folders_from_settings(self.stores)
        if not folders:QMessageBox.warning(self,"フルビルド","有効なフォルダがありません。");return
        if QMessageBox.question(self,"フルビルド",f"{len(folders)} フォルダから検索DBを作成します。実行しますか？")!=QMessageBox.StandardButton.Yes:return
        self._busy(True);self.progress.setRange(0,100);self.progress.setValue(0);self.thread=QThread();self.worker=BuildWorker(self.db,folders);self.worker.moveToThread(self.thread);self.thread.started.connect(self.worker.run);self.worker.progress.connect(self._build_progress);self.worker.finished.connect(self._build_ok);self.worker.failed.connect(self._fail);self.worker.cancelled.connect(self._cancelled);self.worker.finished.connect(self.thread.quit);self.worker.failed.connect(self.thread.quit);self.worker.cancelled.connect(self.thread.quit);self.thread.finished.connect(self._thread_done);self.thread.start()
    def _build_progress(self,d):
        t=d.get("type","");fn=int(d.get("folder_number",0));ft=int(d.get("folder_total",0));ins=int(d.get("inserted",0));sk=int(d.get("skipped",0));wa=int(d.get("warnings",0));path=d.get("folder_path","")
        if t=="mail_progress":
            i=int(d.get("item_number",0));it=int(d.get("item_total",0));frac=((fn-1)+(i/it if it else 0))/ft if ft else 0;self.progress.setValue(min(98,int(frac*98)));self.status.setText(f"フルビルド中\nフォルダ: {fn}/{ft}\n{path}\n処理: {i:,}/{it:,}\nDB登録: {ins:,}  スキップ: {sk:,}  取得警告: {wa:,}")
        elif t=="folder_start":self.status.setText(f"フォルダ: {fn}/{ft}\n{path}\nDB登録: {ins:,}")
        elif t=="finalizing":self.progress.setValue(98);self.status.setText("FTS検索インデックスを作成しています...")
        elif t=="validating":self.progress.setValue(99);self.status.setText("新しい検索DBを検証しています...")
        elif t=="swapping":self.progress.setValue(99);self.status.setText("新しい検索DBへ切り替えています...")
    def _build_ok(self,r):
        self.progress.setValue(100);self._dbstatus();txt=f"登録メール: {r.get('mail_count',0):,}\nスキップ: {r.get('skipped_count',0):,}\n取得警告: {r.get('warning_count',0):,}\n成功フォルダ: {r.get('folder_success_count',0):,}\nエラーフォルダ: {r.get('folder_error_count',0):,}";self.status.setText("フルビルド完了\n"+txt);QMessageBox.information(self,"フルビルド完了",txt)
    def _fail(self,e):self.status.setText("エラー: "+e);QMessageBox.critical(self,"エラー",e+"\n\nOutlookDBBuilder.log を確認してください。")
    def _cancelled(self):self.status.setText("処理をキャンセルしました。")
    def _cancel(self):
        if self.worker:self.worker.cancel=True;self.status.setText("キャンセルしています...")
    def _thread_done(self):self.worker=None;self.thread=None;self._busy(False);self.progress.setRange(0,100);self._dbstatus()
    def _dbstatus(self):
        s=self.db.get_status();self.db_label.setText("DB: 未作成" if not s.get("exists") else f"DB: mail.db / メール {s.get('mail_count',0):,} 件")
    def closeEvent(self,e:QCloseEvent):
        if self.thread:QMessageBox.information(self,"終了","処理中です。先にキャンセルしてください。");e.ignore();return
        if self.dirty and QMessageBox.question(self,"未保存","設定変更を保存せず終了しますか？")!=QMessageBox.StandardButton.Yes:e.ignore();return
        e.accept()
