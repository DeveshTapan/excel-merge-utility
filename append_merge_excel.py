#!/usr/bin/env python3
"""
append_merge_excel.py — Colorful GUI + STRICT merge (first sheet only) + Verify + Move Sources + Auto-Save Log

What’s included:
- Merge (first sheet only), strict schema (names+count+order), read as TEXT, .xlsx only
- Colorful GUI: file picker, reorder, progress bar, scrolling log, Save Log…, row count
- Auto output name: merged-YYYY-MM-DD_%I-%M-%S_%p-<6HEX>.xlsx (unique each run)
- Verify after merge (default ON): columns/row count/cell-by-cell + SHA-256 checksum, diff CSV on mismatch (no JSON files)
- Move source files after success (GUI checkbox / CLI --move-sources --move-dest)
- Auto-Save Log — a log file is always written at the end of GUI runs (success OR failure) under <outdir>/Log/
- NEW: Merged files are always saved under <outdir>/MergedFiles/ (folder auto-created)

NOTE: ASCII-safe strings in error paths (no f-strings).
"""
import argparse, glob, os, sys, secrets, threading, queue, traceback, hashlib, shutil
from datetime import datetime
from typing import List, Tuple
import pandas as pd
import numpy as np

# ---------------- Core helpers ----------------
def expand_patterns(patterns, sort_mode):
    out = []
    for pat in patterns:
        m = glob.glob(pat)
        if sort_mode == 'name':
            m.sort()
        out.extend(m)
    return out

def infer_engine(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xlsx', '.xlsm', '.xltx', '.xltm'):
        return 'openpyxl'
    if ext == '.xls':
        return 'xlrd'
    return 'openpyxl'

# Accept only real Excel workbooks we support
_EXCEL_EXTS = ('.xlsx', '.xlsm', '.xltx', '.xltm', '.xls')

def is_excel_file(path: str) -> bool:
    return os.path.splitext(path)[1].lower() in _EXCEL_EXTS

def read_first_sheet(path, force_text, verbose):
    engine = infer_engine(path)
    if verbose:
        print("Reading: {} (engine={}, sheet=0-only, force_text={})".format(path, engine, force_text))
    kwargs = {'engine': engine, 'sheet_name': 0}
    if force_text:
        kwargs['dtype'] = str
    df = pd.read_excel(path, **kwargs)
    df.columns = [str(c) for c in df.columns]
    return df

def _repr_list(values, max_items=50):
    vals = [repr(v) for v in values]
    if len(vals) > max_items:
        vals = vals[:max_items] + ['... ({} total)'.format(len(values))]
    return '[' + ', '.join(vals) + ']'

def print_schema_mismatch_error(baseline_cols, file_cols, file_path):
    base_set = set(baseline_cols)
    file_set = set(file_cols)
    extra_cols = sorted(list(file_set - base_set))
    missing_cols = sorted(list(base_set - file_set))
    lines = []
    lines.append("ERROR: Schema mismatch in later file: {}".format(os.path.basename(file_path)))
    if len(file_cols) != len(baseline_cols):
        lines.append("- Column count differs (first file: {}, later file: {})".format(len(baseline_cols), len(file_cols)))
    if extra_cols:
        lines.append("- Extra columns in later file: {}".format(extra_cols))
    if missing_cols:
        lines.append("- Missing columns in later file: {}".format(missing_cols))
    if file_cols != baseline_cols:
        order_diffs = []
        for i, (exp, got) in enumerate(zip(baseline_cols, file_cols)):
            if exp != got:
                order_diffs.append((i, exp, got))
        if order_diffs:
            preview = ', '.join(["pos {}: expected {} got {}".format(i, repr(exp), repr(got)) for i, exp, got in order_diffs[:10]])
            lines.append("- Column order differs (examples: {})".format(preview))
    lines.append("- Baseline columns: {}".format(_repr_list(baseline_cols)))
    lines.append("- Later-file columns: {}".format(_repr_list(file_cols)))
    lines.append("Aborting without writing output.")
    sys.stderr.write("\n".join(lines) + "\n")

# ---------- Output naming ----------
def _stamp():
    return datetime.now().strftime('%Y-%m-%d_%I-%M-%S_%p')

def build_output_path(outdir, base_name):
    ts = _stamp()
    uid = secrets.token_hex(3).upper()
    filename = '{}-{}-{}.xlsx'.format(base_name, ts, uid)
    return os.path.join(outdir, filename)

def build_folder_name(base='archived-sources'):
    ts = _stamp()
    uid = secrets.token_hex(3).upper()
    return '{}-{}-{}'.format(base, ts, uid)

# ---------- Canonical hashing for verify ----------
def _df_to_canonical_lines(df):
    df2 = df.fillna('')
    cols = [str(c) for c in df2.columns]
    sep = '\x1f'
    lines = [sep.join(cols)]
    for row in df2.itertuples(index=False, name=None):
        lines.append(sep.join([str(x) for x in row]))
    return lines

def _sha256_hex_from_lines(lines):
    h = hashlib.sha256()
    for ln in lines:
        h.update((ln + '\n').encode('utf-8', 'surrogatepass'))
    return h.hexdigest()

# ---------- Verification (NO JSON writes) ----------
def verify_merged_against_inputs(merged_path, input_paths, force_text=True, verbose=False, outdir=None, max_diffs=1000):
    m = pd.read_excel(merged_path, engine='openpyxl', sheet_name=0, dtype=str)
    m_cols = [str(c) for c in m.columns]

    frames = [read_first_sheet(p, True, verbose) for p in input_paths]
    exp = pd.concat(frames, axis=0, ignore_index=True)
    exp_cols = [str(c) for c in exp.columns]

    info = {
        'merged_rows': int(len(m)),
        'expected_rows': int(len(exp)),
        'columns_equal': (m_cols == exp_cols),
        'columns': m_cols,
        'inputs_count': int(len(input_paths)),
        'inputs': list(input_paths),
        'merged_file': merged_path,
        'diff_csv': None
    }

    if not info['columns_equal']:
        info['reason'] = 'columns_mismatch'
        return False, info

    if len(m) != len(exp):
        info['reason'] = 'row_count_mismatch'
        return False, info

    m_ck = _sha256_hex_from_lines(_df_to_canonical_lines(m))
    e_ck = _sha256_hex_from_lines(_df_to_canonical_lines(exp))
    info['checksum_merged'] = m_ck
    info['checksum_expected'] = e_ck
    info['checksum_equal'] = bool(m_ck == e_ck)

    m2 = m.fillna(''); e2 = exp.fillna('')
    same = (m2.values == e2.values)
    if same.all():
        info['reason'] = 'ok'
        return True, info

    diff_mask = ~same
    row_idx, col_idx = np.where(diff_mask)
    col_counts = {}
    for j, col in enumerate(m2.columns):
        col_counts[col] = int((~(m2.iloc[:, j] == e2.iloc[:, j])).sum())

    records = []
    for r, c in zip(row_idx[:max_diffs], col_idx[:max_diffs]):
        records.append({'row_index': int(r), 'column': m2.columns[c], 'expected': e2.iat[r, c], 'merged': m2.iat[r, c]})
    diff_df = pd.DataFrame.from_records(records, columns=['row_index','column','expected','merged'])

    diff_path = None
    if outdir:
        diff_path = os.path.splitext(build_output_path(outdir, 'verify-diff'))[0] + '.csv'
        diff_df.to_csv(diff_path, index=False, encoding='utf-8')

    info['reason'] = 'values_mismatch'
    info['diffs_found'] = int(len(records))
    info['per_column_mismatch_counts'] = col_counts
    info['diff_csv'] = diff_path
    return False, info

# ---------- Move sources ----------
def _safe_move(src_path, dest_dir):
    import shutil
    os.makedirs(dest_dir, exist_ok=True)
    base = os.path.basename(src_path)
    target = os.path.join(dest_dir, base)
    if not os.path.exists(target):
        shutil.move(src_path, target)
        return target
    name, ext = os.path.splitext(base)
    n = 1
    while True:
        candidate = os.path.join(dest_dir, "{}_copy{}{}".format(name, n, ext))
        if not os.path.exists(candidate):
            shutil.move(src_path, candidate)
            return candidate
        n += 1

def move_sources(files, archive_dir, log=None):
    moved = []
    for f in files:
        try:
            newp = _safe_move(f, archive_dir)
            if log: log("Moved: {} -> {}".format(f, newp))
            moved.append((f, newp))
        except Exception as e:
            if log: log("ERROR moving {}: {}".format(f, e))
    return moved

# ---------------- CLI parsing ----------------
def parse_args():
    p = argparse.ArgumentParser(description=(
        "Append multiple Excel files into one (first sheet only), preserving order; "
        "fail fast if any schema difference is detected. GUI available via --gui."))
    p.add_argument('files', nargs='*', help='Input Excel files in the exact order to append')
    p.add_argument('--pattern', action='append', default=[], help='Glob pattern(s) to include additional files')
    p.add_argument('--pattern-sort', choices=['name','none'], default='name', help='Sort order for pattern matches')
    p.add_argument('--outdir', default=None, help='Directory to write the outputs (MergedFiles & Log folders are created here)')
    p.add_argument('--output', default=None, help='Explicit output Excel filename (basename used under MergedFiles)')
    p.add_argument('--sheet-out', default='Sheet1', help='Output sheet name (default: Sheet1)')
    p.add_argument('--force-text', action='store_true', default=True, help='(Default ON) Read cells as text')
    p.add_argument('--no-force-text', action='store_false', dest='force_text', help='Disable reading-as-text')
    p.add_argument('--verbose', action='store_true', help='Print progress messages to console')
    p.add_argument('--gui', action='store_true', help='Open a colorful GUI to select files and output folder')
    # verification
    p.add_argument('--verify', action='store_true', default=True, help='Verify merged equals inputs (default ON)')
    p.add_argument('--no-verify', action='store_false', dest='verify', help='Disable verification')
    # move sources
    p.add_argument('--move-sources', action='store_true', help='Move source files after success (and verify pass if enabled)')
    p.add_argument('--move-dest', default=None, help='Destination folder for moved source files')
    return p.parse_args()

# ---------------- GUI (colorful + progress + save log + verify + move sources + auto-save log) ----------------
def try_import_tk():
    try:
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox
        return tk, ttk, filedialog, messagebox
    except Exception:
        print("ERROR: Tkinter GUI is not available. Install tkinter or run without --gui.", file=sys.stderr)
        raise

class MergeGUI:
    def __init__(self):
        tk, ttk, filedialog, messagebox = try_import_tk()
        self.tk = tk; self.ttk = ttk; self.filedialog = filedialog; self.messagebox = messagebox
        self.root = tk.Tk()
        self.root.title("Excel Merge - First Sheet Only")
        self.root.geometry("980x700")
        self.root.minsize(900, 620)
        self.root.configure(bg="#0f172a")

        self.files = []
        self.verbose = tk.BooleanVar(value=False)
        self.verify = tk.BooleanVar(value=True)
        self.move_sources_var = tk.BooleanVar(value=False)
        self.outdir = tk.StringVar(value=os.getcwd())
        self.archive_dir = tk.StringVar(value="")
        self.is_merging = False

        self.log_q = queue.Queue()
        self._build_ui()
        self._drain_log_queue()

    def _build_ui(self):
        tk = self.tk; ttk = self.ttk
        banner = tk.Canvas(self.root, height=120, bd=0, highlightthickness=0, relief='ridge'); banner.pack(fill='x', side='top')
        colors = ["#06b6d4","#22d3ee","#38bdf8","#6366f1","#a78bfa"]; width=240
        for i,c in enumerate(colors): banner.create_rectangle(i*width,0,(i+1)*width+60,120,fill=c,outline='')
        banner.create_text(28,34,anchor='w',text='Excel Merge',font=('Segoe UI',20,'bold'),fill='white')
        banner.create_text(28,70,anchor='w',text='First sheet only - Strict schema - Append in your order',font=('Segoe UI',11),fill='white')

        main = tk.Frame(self.root,bg="#0f172a"); main.pack(fill='both',expand=True,padx=16,pady=16)

        left = tk.Frame(main,bg="#0f172a"); left.pack(side='left',fill='both',expand=True)
        btns = tk.Frame(left,bg="#0f172a"); btns.pack(fill='x',pady=(0,8))
        self._mk_button(btns,"Select Files",self._pick).pack(side='left',padx=(0,8))
        self._mk_button(btns,"Clear",self._clear).pack(side='left',padx=(0,8))
        self._mk_button(btns,"Move Up",self._move_up).pack(side='left',padx=(0,8))
        self._mk_button(btns,"Move Down",self._move_down).pack(side='left')

        list_frame = tk.Frame(left,bg="#0f172a"); list_frame.pack(fill='both',expand=True)
        self.lst = tk.Listbox(list_frame,selectmode='extended',activestyle='dotbox',bg='#111827',fg='white',
                              highlightthickness=1,highlightcolor='#22d3ee',highlightbackground='#374151',font=('Consolas',10))
        self.lst.pack(side='left',fill='both',expand=True)
        vsb = tk.Scrollbar(list_frame,orient='vertical',command=self.lst.yview); vsb.pack(side='right',fill='y')
        self.lst.config(yscrollcommand=vsb.set)

        tk.Label(left,text='Progress',bg="#0f172a",fg='#cbd5e1',font=('Segoe UI',10,'bold')).pack(anchor='w',pady=(10,2))
        prog_frame = tk.Frame(left,bg="#0f172a"); prog_frame.pack(fill='both',expand=True)
        self.log_text = tk.Text(prog_frame,height=13,bg='#0b1220',fg='#e5e7eb',insertbackground='white',relief='flat',wrap='word',font=('Consolas',10))
        self.log_text.pack(side='left',fill='both',expand=True)
        vsb2 = tk.Scrollbar(prog_frame,orient='vertical',command=self.log_text.yview); vsb2.pack(side='right',fill='y')
        self.log_text.config(yscrollcommand=vsb2.set,state='disabled')

        right = tk.Frame(main,bg="#0f172a"); right.pack(side='left',fill='y',padx=(12,0))
        card = tk.Frame(right,bg="#111827",bd=0,highlightbackground='#374151',highlightcolor='#374151',highlightthickness=1)
        card.pack(fill='y',padx=4,pady=4)

        tk.Label(card,text='Output folder',bg="#111827",fg='white',font=('Segoe UI',10,'bold')).pack(anchor='w',padx=12,pady=(12,2))
        out_row = tk.Frame(card,bg="#111827"); out_row.pack(fill='x',padx=12)
        self.out_entry = tk.Entry(out_row,textvariable=self.outdir,bg='#0b1220',fg='white',insertbackground='white',relief='flat')
        self.out_entry.pack(side='left',fill='x',expand=True)
        self._mk_button(out_row,'Browse...',self._browse_out).pack(side='left',padx=(8,0))

        v_row = tk.Frame(card,bg="#111827"); v_row.pack(fill='x',padx=12,pady=(12,4))
        tk.Checkbutton(v_row,text='Verbose console output',variable=self.verbose,bg="#111827",fg='white',
                       activebackground="#111827",selectcolor="#111827").pack(side='left')
        tk.Checkbutton(card,text='Verify after merge (row-by-row)',variable=self.verify,bg="#111827",fg='white',
                       activebackground="#111827",selectcolor="#111827").pack(anchor='w',padx=12)
        tk.Checkbutton(card,text='Move source files after success',variable=self.move_sources_var,bg="#111827",fg='white',
                       activebackground="#111827",selectcolor="#111827").pack(anchor='w',padx=12,pady=(6,0))

        arch_row = tk.Frame(card,bg="#111827"); arch_row.pack(fill='x',padx=12,pady=(2,0))
        tk.Label(arch_row,text='Archive folder (optional):',bg="#111827",fg='#cbd5e1').pack(side='left')
        self.arch_entry = tk.Entry(card,textvariable=self.archive_dir,bg='#0b1220',fg='white',insertbackground='white',relief='flat')
        self.arch_entry.pack(fill='x',padx=12,pady=(2,10))
        self._mk_button(card,'Archive…',self._browse_arch).pack(anchor='e',padx=12)

        pbar_row = tk.Frame(card,bg="#111827"); pbar_row.pack(fill='x',padx=12,pady=(8,2))
        self.pbar = self.ttk.Progressbar(pbar_row,orient='horizontal',length=260,mode='determinate',maximum=100,value=0)
        self.pbar.pack(fill='x')

        self._mk_cta(card,'Merge Now',self._merge).pack(fill='x',padx=12,pady=(10,6))
        self._mk_button(card,'Save Log...',self._save_log).pack(fill='x',padx=12,pady=(0,12))

        self.status = tk.Label(self.root,text='Ready.',bg="#0f172a",fg='#cbd5e1',anchor='w')
        self.status.pack(fill='x',side='bottom',padx=12,pady=(0,8))

    def _mk_button(self,parent,text,cmd):
        return self.tk.Button(parent,text=text,command=cmd,bg='#1f2937',fg='white',
                              activebackground='#22d3ee',activeforeground='white',bd=0,padx=10,pady=6,cursor='hand2')

    def _mk_cta(self,parent,text,cmd):
        return self.tk.Button(parent,text=text,command=cmd,bg='#22c55e',fg='black',
                              activebackground='#16a34a',activeforeground='black',bd=0,padx=12,pady=10,cursor='hand2',font=('Segoe UI',11,'bold'))

    # --- logging ---
    def log(self,msg):
        t = datetime.now().strftime('%H:%M:%S'); self.log_q.put(('log','[{}] {}'.format(t,msg)))
    def set_status(self,text): self.log_q.put(('status',text))
    def set_prog(self,val): self.log_q.put(('prog',val))
    def _append_log_now(self,text):
        self.log_text.config(state='normal'); self.log_text.insert('end',text+'\n'); self.log_text.see('end'); self.log_text.config(state='disabled')

    def _drain_log_queue(self):
        try:
            while True:
                kind,*rest = self.log_q.get_nowait()
                if kind=='log': self._append_log_now(rest[0])
                elif kind=='status': self.status.config(text=rest[0])
                elif kind=='prog': self.pbar['value']=rest[0]
                elif kind=='finish': self.is_merging=False; self._enable_controls(True); self._auto_save_log_safe()
        except queue.Empty:
            pass
        self.root.after(120,self._drain_log_queue)

    # --- auto-save log (Log folder) ---
    def _auto_save_log_safe(self):
        """
        Always save the log at the end of a run (success OR failure),
        now under <outdir>/Log/.
        """
        try:
            outdir = self.outdir.get() or os.getcwd()
            log_dir = os.path.join(outdir, 'Log')
            os.makedirs(log_dir, exist_ok=True)
            path = os.path.splitext(build_output_path(log_dir,'merge-log'))[0]+'.txt'
            data = self.log_text.get('1.0','end').strip()+'\n'
            with open(path,'w',encoding='utf-8') as f: f.write(data)
            self.log("Auto-saved log: {}".format(path))
            self.status.config(text='Auto-saved log: {}'.format(path))
        except Exception as e:
            try:
                self.messagebox.showerror('Auto-Save Log Error', str(e))
            except Exception:
                pass

    # --- actions ---
    def _enable_controls(self,enable):
        state='normal' if enable else 'disabled'
        for w in (self.out_entry,self.arch_entry):
            try: w.config(state=state)
            except Exception: pass

    def _pick(self):
        paths = self.filedialog.askopenfilenames(
            title='Select Excel files in order to append',
            filetypes=[('Excel files','*.xlsx *.xlsm *.xltx *.xltm *.xls'),('All files','*.*')])
        if not paths: return
        kept = [p for p in paths if is_excel_file(p)]
        dropped = [p for p in paths if not is_excel_file(p)]
        for p in kept:
            self.files.append(p); self.lst.insert('end',p)
        if dropped:
            self.log('Ignored non-Excel files: {}'.format(', '.join([os.path.basename(x) for x in dropped])))
        self.set_status('Added {} file(s). Total: {}'.format(len(kept), len(self.files)))

    def _clear(self):
        self.files[:] = []; self.lst.delete(0,'end'); self.set_status('Cleared file list.')

    def _move_up(self):
        sel=list(self.lst.curselection());
        if not sel: return
        for i in sel:
            if i==0: continue
            self._swap(i,i-1)
        self.lst.select_clear(0,'end')
        for i in [i-1 for i in sel if i>0]: self.lst.select_set(i); self.lst.activate(i)

    def _move_down(self):
        sel=list(self.lst.curselection());
        if not sel: return
        for i in reversed(sel):
            if i==self.lst.size()-1: continue
            self._swap(i,i+1)
        self.lst.select_clear(0,'end')
        for i in [i+1 for i in sel if i<self.lst.size()]: self.lst.select_set(i); self.lst.activate(i)

    def _swap(self,i,j):
        li=self.lst.get(i); lj=self.lst.get(j)
        self.lst.delete(i); self.lst.insert(i,lj)
        self.lst.delete(j); self.lst.insert(j,li)
        self.files[i],self.files[j]=self.files[j],self.files[i]

    def _browse_out(self):
        d = self.filedialog.askdirectory(title='Choose output folder', mustexist=True, initialdir=self.outdir.get())
        if d: self.outdir.set(d); self.set_status('Output folder: {}'.format(d))

    def _browse_arch(self):
        d = self.filedialog.askdirectory(title='Choose archive (move to) folder', mustexist=True, initialdir=self.outdir.get())
        if d: self.archive_dir.set(d); self.set_status('Archive folder: {}'.format(d))

    def _save_log(self):
        # Manual save (still available) -> <outdir>/Log/
        try:
            outdir = self.outdir.get() or os.getcwd()
            log_dir = os.path.join(outdir, 'Log')
            os.makedirs(log_dir, exist_ok=True)
            path = os.path.splitext(build_output_path(log_dir,'merge-log'))[0]+'.txt'
            data = self.log_text.get('1.0','end').strip()+'\n'
            with open(path,'w',encoding='utf-8') as f: f.write(data)
            self.set_status('Log saved: {}'.format(path))
            self.log('Log saved: {}'.format(path))
        except Exception as e:
            self.messagebox.showerror('Error', str(e))

    def _merge(self):
        if self.is_merging: return
        if not self.files:
            self.messagebox.showerror('No files','Select at least one Excel file.'); return
        missing=[p for p in self.files if not os.path.exists(p)]
        if missing:
            self.messagebox.showerror('Missing files','These files do not exist:\n'+'\n'.join(missing)); return

        extra = 1 if self.verify.get() else 0
        extra += 1 if self.move_sources_var.get() else 0
        total_steps = len(self.files) + 1 + extra

        self.pbar['mode']='determinate'; self.pbar['maximum']=total_steps; self.pbar['value']=0
        self.is_merging=True; self._enable_controls(False)
        self.set_status('Merging...'); self.log('Starting merge job...')
        th = threading.Thread(target=self._merge_worker, daemon=True); th.start()

    def _merge_worker(self):
        try:
            # Filter only Excel files (defensive)
            files=[f for f in self.files if is_excel_file(f)]
            outdir = self.outdir.get() or os.getcwd()

            # Always write merged file under <outdir>/MergedFiles/
            merged_dir = os.path.join(outdir, 'MergedFiles')
            os.makedirs(merged_dir, exist_ok=True)
            out = build_output_path(merged_dir,'merged')

            self.log('Files to merge: {}'.format(len(files)))
            if self.verbose.get():
                for i,fpath in enumerate(files,1): self.log(' {:>3}. {}'.format(i,fpath))

            self.log('Reading baseline (first file)...')
            df0 = read_first_sheet(files[0], True, self.verbose.get()); base=list(df0.columns)
            frames=[df0]; per=[(files[0],len(df0))]; self.set_prog(1)

            for idx, path in enumerate(files[1:], start=2):
                self.log('Reading file {}/{}: {}'.format(idx,len(files),os.path.basename(path)))
                df = read_first_sheet(path, True, self.verbose.get())
                if (len(df.columns)!=len(base)) or (set(df.columns)!=set(base)) or (list(df.columns)!=base):
                    print_schema_mismatch_error(base, list(df.columns), path); raise RuntimeError('Schema mismatch. See console for details.')
                frames.append(df); per.append((path,len(df))); self.set_prog(idx)

            self.log('Concatenating frames...')
            merged = pd.concat(frames, axis=0, ignore_index=True)

            self.log('Writing Excel -> {}'.format(out))
            with pd.ExcelWriter(out, engine='openpyxl') as w: merged.to_excel(w, sheet_name='Sheet1', index=False)
            self.set_prog(len(files)+1)

            total_rows = len(merged)
            self.log('Merge complete. Summary:'); [self.log(' - {}: {} rows'.format(os.path.basename(p),n)) for p,n in per]
            self.log(' = TOTAL: {} rows'.format(total_rows))
            self.log('Output Excel: {}'.format(out))

            ok_verify = True
            if self.verify.get():
                self.log('Verifying merged workbook against originals...')
                ok_verify, info = verify_merged_against_inputs(out, files, force_text=True, verbose=self.verbose.get(), outdir=outdir)
                self.log('Verify columns equal: {}'.format(info.get('columns_equal')))
                self.log('Verify row counts: merged={}, expected={}'.format(info.get('merged_rows'), info.get('expected_rows')))
                self.log('Verify checksum equal: {}'.format(info.get('checksum_equal')))
                cm = (info.get('checksum_merged','')[:12] if info.get('checksum_merged') else '')
                ce = (info.get('checksum_expected','')[:12] if info.get('checksum_expected') else '')
                if cm or ce: self.log('Checksums (first 12): merged={}, expected={}'.format(cm, ce))
                if not ok_verify:
                    self.log('Verification FAILED: reason={}'.format(info.get('reason','unknown')))
                    if info.get('diff_csv'): self.log('Differences CSV: {}'.format(info['diff_csv']))
                    pcm = info.get('per_column_mismatch_counts')
                    if pcm:
                        self.log('Per-column mismatch counts:')
                        for k,v in pcm.items(): self.log('  {}: {}'.format(k,v))
            self.set_prog(len(files)+2)

            if self.move_sources_var.get():
                if ok_verify and self.verify.get() or (not self.verify.get()):
                    target_dir = self.archive_dir.get().strip()
                    if not target_dir:
                        target_dir = os.path.join(outdir, build_folder_name('archived-sources'))
                    self.log('Moving source files to: {}'.format(target_dir))
                    moved = move_sources(files, target_dir, log=self.log)
                    self.log('Moved {} file(s).'.format(len(moved)))
                else:
                    self.log('Skip moving sources because verification failed.')

            self.set_prog(self.pbar['maximum'])
            if not self.verify.get():
                self.set_status('Success: {} ({} rows)'.format(out, total_rows))
            else:
                if ok_verify: self.set_status('Success (verified): {} ({} rows)'.format(out, total_rows))
                else: self.set_status('Merged written; verification failed. See log.')
            self.messagebox.showinfo('Merge complete', '{}\nRows: {}'.format(out, total_rows))
        except Exception as e:
            self.log('ERROR during merge.')
            if self.verbose.get(): traceback.print_exc()
            self.set_status('Error: {}'.format(e)); self.messagebox.showerror('Error', str(e))
        finally:
            # signal finisher -> auto-save log will run from _drain_log_queue()
            self.log_q.put(('finish',))

    # expose a run() for convenience (not required)
    def run(self):
        self.root.mainloop()

# ---------------- Main (CLI) ----------------
def main():
    args = parse_args()

    if args.gui:
        print('[INFO] Launching GUI...')
        app = MergeGUI()
        try: app.run()
        except AttributeError: app.root.mainloop()
        return 0

    files = list(args.files) + expand_patterns(args.pattern, args.pattern_sort)
    # Keep only Excel files
    files = [f for f in files if is_excel_file(f)]

    if not files:
        sys.stderr.write('ERROR: No input Excel files supplied. Provide files/patterns or use --gui.\n'); return 2

    missing=[f for f in files if not os.path.exists(f)]
    if missing:
        sys.stderr.write('ERROR: These files do not exist:\n - ' + '\n - '.join(missing) + '\n'); return 2

    if args.verbose:
        print('Final input order:'); [print(' {:>3}. {}'.format(i,f)) for i,f in enumerate(files,1)]

    # Base directory user intended for outputs
    base_outdir = args.outdir if args.outdir else os.getcwd()

    # Ensure <base_outdir>/MergedFiles/ exists and place the final merged file there
    merged_dir = os.path.join(base_outdir, 'MergedFiles')
    os.makedirs(merged_dir, exist_ok=True)

    # If --output is given, use its basename under MergedFiles; else auto-name under MergedFiles
    if args.output:
        out = os.path.join(merged_dir, os.path.basename(args.output))
    else:
        out = build_output_path(merged_dir, 'merged')

    df0 = read_first_sheet(files[0], args.force_text, args.verbose); base=list(df0.columns)
    frames=[df0]; per=[(files[0],len(df0))]
    for p in files[1:]:
        df=read_first_sheet(p,args.force_text,args.verbose)
        if (len(df.columns)!=len(base)) or (set(df.columns)!=set(base)) or (list(df.columns)!=base):
            print_schema_mismatch_error(base, list(df.columns), p); return 3
        frames.append(df); per.append((p,len(df)))

    merged = pd.concat(frames, axis=0, ignore_index=True); total_rows=len(merged)
    if args.verbose: print('Writing Excel: {} (sheet={})'.format(out, args.sheet_out))
    with pd.ExcelWriter(out, engine='openpyxl') as w: merged.to_excel(w, sheet_name=args.sheet_out, index=False)
    print('Merge complete. Summary:'); [print(' - {}: {} rows'.format(os.path.basename(p),n)) for p,n in per]
    print(' = TOTAL: {} rows'.format(total_rows)); print('Output Excel: {}'.format(out))

    ok_verify = True
    if args.verify:
        print('Verifying merged workbook against originals...')
        ok_verify, info = verify_merged_against_inputs(out, files, force_text=args.force_text, verbose=args.verbose, outdir=base_outdir)
        print('Verify columns equal:', info.get('columns_equal'))
        print('Verify row counts: merged={}, expected={}'.format(info.get('merged_rows'), info.get('expected_rows')))
        print('Verify checksum equal:', info.get('checksum_equal'))
        cm = (info.get('checksum_merged','')[:12] if info.get('checksum_merged') else '')
        ce = (info.get('checksum_expected','')[:12] if info.get('checksum_expected') else '')
        if cm or ce: print('Checksums (first 12): merged={}, expected={}'.format(cm, ce))
        if not ok_verify:
            print('Verification FAILED: reason={}'.format(info.get('reason','unknown')))
            if info.get('diff_csv'): print('Differences CSV:', info['diff_csv'])

    if args.move_sources:
        if ok_verify and args.verify or (not args.verify):
            dest = args.move_dest if args.move_dest else os.path.join((base_outdir or os.getcwd()), build_folder_name('archived-sources'))
            print('Moving source files to:', dest)
            for f in files:
                try:
                    newp = _safe_move(f, dest)
                    if args.verbose: print('  Moved:', f, '->', newp)
                except Exception as e:
                    print('  ERROR moving {}: {}'.format(f, e))
        else:
            print('Skip moving sources because verification failed.')

    return 0

if __name__ == '__main__':
    print('[DEBUG] file: {}'.format(__file__)); print('[DEBUG] argv: {}'.format(sys.argv))
    if '--gui' in sys.argv[1:]:
        print('[INFO] Launching GUI (pre-argparse fallback)...')
        app = MergeGUI()
        try: app.run()
        except AttributeError: app.root.mainloop()
        sys.exit(0)
    sys.exit(main())