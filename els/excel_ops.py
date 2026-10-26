"""Bounded Excel operations and per-generation resources; no persistent cache."""
from contextlib import contextmanager,ExitStack
from contextvars import ContextVar
from functools import wraps
from pathlib import Path
import json,tempfile,time

_run=ContextVar('els_generation',default=None)
_sources=ContextVar('els_reference_pool',default=None)
_printing=ContextVar('els_print_batch_depth',default=0)


class GenerationRun:
    def __init__(self,callback=None):
        self.callback=callback;self.started=time.perf_counter();self.phase=None;self.records=[]
    def advance(self,name,message=''):
        now=time.perf_counter()
        if self.phase:
            key,start=self.phase;self.records.append(dict(stage=key,seconds=round(now-start,4)))
        self.phase=(name,now) if name else None
        if message and self.callback:self.callback(message)
    def finish(self,result):
        self.advance(None)
        result['performance']=dict(total_seconds=round(time.perf_counter()-self.started,3),stages=self.records)
        target=Path(result['workbook']).parent/'result.json'
        target.write_text(json.dumps(result,ensure_ascii=False,indent=2,allow_nan=False),encoding='utf8')
        return result


@contextmanager
def generation_run(callback=None):
    run=GenerationRun(callback);token=_run.set(run)
    try:
        run.advance('validate','检查输入及参考版本…');yield run
    finally:_run.reset(token)


def progress(stage,message):
    run=_run.get()
    if run:run.advance(stage,message)


@contextmanager
def print_batch(app):
    depth=_printing.get();token=_printing.set(depth+1)
    previous=None
    try:
        if not depth:
            previous=app.PrintCommunication;app.PrintCommunication=False
        yield
    finally:
        try:
            if not depth and previous is not None and app.PrintCommunication!=previous:app.PrintCommunication=previous
        finally:_printing.reset(token)


def batched_print(function):
    @wraps(function)
    def wrapped(ws,*args,**kwargs):
        with print_batch(ws.Application):
            return function(ws,*args,**kwargs)
    return wrapped


def commit_print(app):
    """Commit pending page settings before inspecting print areas or exporting."""
    if not app.PrintCommunication:app.PrintCommunication=True


def last_data_row(ws):
    found=ws.Cells.Find(What='*',LookIn=-4123,LookAt=2,SearchOrder=1,
                        SearchDirection=2,MatchCase=False,SearchFormat=False)
    return found.Row if found is not None else 0


def hide_after(ws,row):
    # A print area excludes blank tail rows. Formatting 65,536 rows expands
    # UsedRange and turns a small workbook into millions of marshalled cells.
    last=last_data_row(ws)
    if last>row:ws.Rows(f'{row+1}:{last}').Hidden=True


def sanitize_filename_cells(ws,output_name):
    cells=ws.Cells
    while True:
        found=cells.Find(What='CELL("FILENAME"',LookIn=-4123,LookAt=2,
                         SearchOrder=1,SearchDirection=1,MatchCase=False,SearchFormat=False)
        if found is None:return
        found.Value2=f'OUTPUT/[{output_name}]{ws.Name}'


def finish_headers(wb,output_name,failed):
    with print_batch(wb.Application):
        for ws in wb.Worksheets:
            sanitize_filename_cells(ws,output_name)
            # Historical sheets are retained very hidden and are never printed.
            # Their original English headers need no printer reconfiguration.
            if ws.Visible!=-1:continue
            page=ws.PageSetup
            for field in ('LeftHeader','RightHeader','LeftFooter','CenterFooter','RightFooter'):setattr(page,field,'')
            page.CenterHeader='REVIEW DRAFT - NOT PASS' if failed else 'REVIEW DRAFT'


class SourcePool:
    def __init__(self,app):self.app=app;self.books={};self.stack=ExitStack()
    def acquire(self,source):
        key=str(Path(source).resolve())
        if key not in self.books:
            from .connections import copy_source
            from .sheetpile import excel_call
            folder=self.stack.enter_context(tempfile.TemporaryDirectory(prefix='els-reference-'))
            snapshot=Path(folder)/Path(source).name;copy_source(source,snapshot)
            book=self.app.Workbooks.Open(str(snapshot),UpdateLinks=0,ReadOnly=True,AddToMru=False)
            close=book.Close
            self.stack.callback(lambda:excel_call(lambda:close(SaveChanges=False)))
            self.books[key]=book
        return self.books[key]


@contextmanager
def reference_pool(app):
    pool=SourcePool(app);token=_sources.set(pool)
    try:yield
    finally:
        try:pool.stack.close()
        finally:_sources.reset(token)


def pooled_source(source):
    pool=_sources.get()
    return pool.acquire(source) if pool else None
