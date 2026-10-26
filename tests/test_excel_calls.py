import unittest
from unittest.mock import Mock,patch
try:
 import pywintypes
except ImportError:pywintypes=None
from els.sheetpile import excel_call


@unittest.skipIf(pywintypes is None,'Windows Excel automation required')
class ExcelCallTests(unittest.TestCase):
 def test_busy_excel_retries_then_returns_result(self):
  busy=pywintypes.com_error(-2147418111,'busy',None,None)
  action=Mock(side_effect=[busy,busy,'saved'])
  with patch('time.sleep'),patch('pythoncom.PumpWaitingMessages'):
   self.assertEqual(excel_call(action),'saved')
  self.assertEqual(action.call_count,3)
 def test_nonbusy_error_is_not_suppressed(self):
  error=pywintypes.com_error(-2147352567,'invalid workbook',None,None);action=Mock(side_effect=error)
  with self.assertRaises(pywintypes.com_error):excel_call(action)
  self.assertEqual(action.call_count,1)
