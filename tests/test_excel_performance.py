import json,tempfile,unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock,patch
from els.excel_ops import (print_batch,hide_after,sanitize_filename_cells,
                           generation_run,progress,reference_pool,pooled_source,commit_print,finish_headers)


class ExcelPerformanceTests(unittest.TestCase):
 def test_print_batch_commits_once_and_restores_on_exception(self):
  app=SimpleNamespace(PrintCommunication=True)
  with self.assertRaises(ValueError):
   with print_batch(app):
    self.assertFalse(app.PrintCommunication)
    with print_batch(app):self.assertFalse(app.PrintCommunication)
    self.assertFalse(app.PrintCommunication);raise ValueError('interrupted')
  self.assertTrue(app.PrintCommunication)
  app.PrintCommunication=False
  with print_batch(app):pass
  self.assertFalse(app.PrintCommunication)

 def test_deferred_print_settings_are_committed_before_export(self):
  app=SimpleNamespace(PrintCommunication=True)
  with print_batch(app):
   with print_batch(app):self.assertFalse(app.PrintCommunication)
   commit_print(app);self.assertTrue(app.PrintCommunication)
  self.assertTrue(app.PrintCommunication)

 def test_hidden_sheets_are_sanitized_without_reconfiguring_the_printer(self):
  visible=Mock();visible.Visible=-1
  class HistoricalSheet:
   Visible=2
   @property
   def PageSetup(self):raise AssertionError('Hidden sheets are not printed')
  hidden=HistoricalSheet();wb=SimpleNamespace(Worksheets=[visible,hidden],Application=SimpleNamespace(PrintCommunication=True))
  with patch('els.excel_ops.sanitize_filename_cells') as sanitize:
   finish_headers(wb,'Review.xlsx',True)
  self.assertEqual(sanitize.call_count,2)
  self.assertEqual(visible.PageSetup.CenterHeader,'REVIEW DRAFT - NOT PASS')

 def test_only_nonempty_tail_rows_are_hidden(self):
  ws=Mock();ws.Cells.Find.return_value.Row=137
  hide_after(ws,53);ws.Rows.assert_called_once_with('54:137')
  ws.reset_mock();ws.Cells.Find.return_value.Row=50
  hide_after(ws,53);ws.Rows.assert_not_called()
  ws.Cells.Find.return_value=None;hide_after(ws,53);ws.Rows.assert_not_called()

 def test_filename_cleanup_does_not_read_usedrange(self):
  class Sheet:
   Name='Hidden original'
   Cells=Mock()
   @property
   def UsedRange(self):raise AssertionError('Unbounded read')
  ws=Sheet();one=SimpleNamespace(Value2=None);two=SimpleNamespace(Value2=None)
  ws.Cells.Find.side_effect=[one,two,None]
  sanitize_filename_cells(ws,'Review.xlsx')
  self.assertEqual(one.Value2,'OUTPUT/[Review.xlsx]Hidden original')
  self.assertEqual(two.Value2,one.Value2)
  self.assertEqual(ws.Cells.Find.call_args.kwargs['LookIn'],-4123)

 def test_source_snapshot_is_shared_only_within_one_generation(self):
  app=Mock();book=app.Workbooks.Open.return_value
  with patch('els.connections.copy_source') as copy:
   with reference_pool(app):
    self.assertIs(pooled_source('source.xlsx'),pooled_source('source.xlsx'))
    self.assertEqual(app.Workbooks.Open.call_count,1);book.Close.assert_not_called()
   book.Close.assert_called_once_with(SaveChanges=False)
   self.assertIsNone(pooled_source('source.xlsx'))
   with reference_pool(app):pooled_source('source.xlsx')
   self.assertEqual(copy.call_count,2)
  self.assertTrue(app.Workbooks.Open.call_args.kwargs['ReadOnly'])

 def test_source_book_closes_when_generation_fails(self):
  app=Mock()
  with patch('els.connections.copy_source'):
   with self.assertRaises(ValueError):
    with reference_pool(app):
     pooled_source('source.xlsx');raise ValueError('Formula error')
  app.Workbooks.Open.return_value.Close.assert_called_once_with(SaveChanges=False)
  self.assertIsNone(pooled_source('source.xlsx'))

 def test_progress_and_timings_are_written_after_completion(self):
  with tempfile.TemporaryDirectory() as folder:
   messages=[]
   with generation_run(messages.append) as run:
    progress('export_pdf','导出 PDF…')
    result=run.finish(dict(workbook=str(Path(folder)/'Review.xlsx')))
   self.assertEqual(messages,['检查输入及参考版本…','导出 PDF…'])
   saved=json.loads((Path(folder)/'result.json').read_text(encoding='utf8'))
   self.assertEqual(saved,result)
   self.assertEqual([s['stage'] for s in saved['performance']['stages']],['validate','export_pdf'])
   self.assertGreaterEqual(saved['performance']['total_seconds'],0)
