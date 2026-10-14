"""Regressão do WinError 5 na publicação do relatório."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from saida_segura import criar_temporario, publicar_excel


class SaidaSeguraTest(unittest.TestCase):
    def test_normal_and_unique_temporary(self):
        with tempfile.TemporaryDirectory() as folder:
            dest = Path(folder) / 'report.xlsx'
            first, second = criar_temporario(dest), criar_temporario(dest)
            self.assertNotEqual(first, second)
            first.write_bytes(b'novo')
            self.assertEqual(publicar_excel(first, dest), dest)
            self.assertEqual(dest.read_bytes(), b'novo')
            self.assertFalse(first.exists())

    def test_locked_destination_preserves_existing_and_new(self):
        with tempfile.TemporaryDirectory() as folder:
            dest = Path(folder) / 'report.xlsx'
            dest.write_bytes(b'anterior')
            temp = criar_temporario(dest)
            temp.write_bytes(b'coleta nova')
            with patch('saida_segura.os.replace', side_effect=PermissionError('WinError 5')), \
                 patch('saida_segura.time.sleep'):
                actual = publicar_excel(temp, dest)
            self.assertNotEqual(actual, dest)
            self.assertEqual(actual.read_bytes(), b'coleta nova')
            self.assertEqual(dest.read_bytes(), b'anterior')

    def test_transient_lock(self):
        import os
        replace = os.replace
        with tempfile.TemporaryDirectory() as folder:
            dest = Path(folder) / 'report.xlsx'
            temp = criar_temporario(dest)
            temp.write_bytes(b'novo')
            calls = []
            def transient(src, dst):
                calls.append(dst)
                if len(calls) == 1:
                    raise PermissionError('scanner')
                return replace(src, dst)
            with patch('saida_segura.os.replace', side_effect=transient), patch('saida_segura.time.sleep'):
                self.assertEqual(publicar_excel(temp, dest), dest)
            self.assertEqual(len(calls), 2)

    def test_total_failure_keeps_recoverable_source(self):
        with tempfile.TemporaryDirectory() as folder:
            dest = Path(folder) / 'report.xlsx'
            temp = criar_temporario(dest)
            temp.write_bytes(b'coleta')
            with patch('saida_segura.os.replace', side_effect=PermissionError()), \
                 patch('saida_segura.time.sleep'), patch('pathlib.Path.open', side_effect=PermissionError()):
                with self.assertRaisesRegex(OSError, 'coleta permanece'):
                    publicar_excel(temp, dest)
            self.assertEqual(temp.read_bytes(), b'coleta')


if __name__ == '__main__':
    unittest.main()
