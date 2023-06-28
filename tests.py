import unittest
from unittest.mock import patch, mock_open
import pandas as pd
import numpy as np
import algo  # Import your main script

class TestAlgo(unittest.TestCase):

    @patch('pandas.read_excel')
    def test_process_files(self, mock_read_excel):
        # Mock data for the tests
        mock_data = {
            'Habitaciones': pd.Series([1, 2, 3], dtype=np.int64),
            'Baños': pd.Series([1, 2, 3], dtype=np.int64),
            'm2 totales': pd.Series([100, 200, 300]),
        }

        # Simulate the DataFrame that read_excel() would return
        mock_df = pd.DataFrame(mock_data)
        mock_read_excel.return_value = {'Sheet1': mock_df}

        # Call the process_files() function with a mock file path
        result = algo.process_files(['mock_file.xlsx'])

        # Check the result
        self.assertEqual(result, {'Sheet1': mock_df})

if __name__ == '__main__':
    unittest.main()
