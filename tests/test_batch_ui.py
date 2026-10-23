from io import BytesIO
from pathlib import Path
import unittest
from unittest.mock import patch
from streamlit.testing.v1 import AppTest
import test_batch_quotes


class BatchUITest(unittest.TestCase):
    def test_upload_compute_and_margin(self):
        upload=BytesIO(test_batch_quotes.BatchTest().workbook())
        upload.name='test.xlsx'
        upload.size=len(upload.getvalue())
        def uploader(*args,**kwargs):
            return upload if kwargs.get('key')=='shipment_batch' else None
        with patch('streamlit.file_uploader',side_effect=uploader):
            app=AppTest.from_file(str(Path(__file__).resolve().parents[1]/'app.py'),default_timeout=60)
            app.secrets['APP_PASSWORD']='test'
            app.session_state['authenticated']=True
            app.run()
            self.assertFalse(app.exception)
            self.assertIn('Excel-zendingen',[t.label for t in app.tabs])
            self.assertEqual(app.selectbox(key='batch_country').value,'FR')
            app.button(key='batch_calculate').click().run()
            self.assertFalse(app.exception)
            self.assertIn('batch_result',app.session_state)
            first=app.session_state['batch_result'][1][1][0]
            self.assertFalse(any(k.endswith('Verkoop EUR') for k in first))
            app.checkbox(key='batch_margin').check().run()
            app.button(key='batch_calculate').click().run()
            self.assertFalse(app.exception)
            first=app.session_state['batch_result'][1][1][0]
            self.assertTrue(any(k.endswith('Verkoop EUR') for k in first))
            app.checkbox(key='batch_scale').check().run()
            self.assertTrue(app.button(key='batch_calculate').disabled)
            app.number_input(key='batch_scale_kg').set_value(300.0).run()
            app.button(key='batch_calculate').click().run()
            self.assertFalse(app.exception)
            self.assertEqual(app.session_state['batch_result'][1][2]['staffel']['bestemmingen'],2)
