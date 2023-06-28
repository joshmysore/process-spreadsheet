import unittest
from unittest.mock import patch, MagicMock
import pandas as pd
import numpy as np
from main_script import process_files

class TestMainScript(unittest.TestCase):
    @patch('pandas.read_excel')
    def test_invalid_datatypes(self, mock_read_excel):
        # Mock the return value of pandas.read_excel to simulate an Excel file with invalid datatypes
        df = pd.DataFrame({
            'Estudio': ['Si', 'No'],
            'Habitaciones': ['one', 'two'],
            'Baños': [1.5, 2.5],
            'Latitud': [np.nan, np.nan]
        })
        mock_read_excel.return_value = {'Sheet1': df}

        # Test that process_files raises an error
        with self.assertRaises(ValueError):
            process_files(['dummy_file_path'])

    @patch('pandas.read_excel')
    def test_missing_columns(self, mock_read_excel):
        # Mock the return value of pandas.read_excel to simulate an Excel file with missing columns
        df = pd.DataFrame({
            'Estudio': ['Si', 'No'],
            'Habitaciones': [1, 2]
            # Missing 'Baños' and 'Latitud' columns
        })
        mock_read_excel.return_value = {'Sheet1': df}

        # Test that process_files raises an error
        with self.assertRaises(KeyError):
            process_files(['dummy_file_path'])

    # Add more test cases here...

if __name__ == '__main__':
    unittest.main()
