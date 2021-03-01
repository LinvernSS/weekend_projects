import unittest
import NYL_data_analysis as nda
import pandas as pd
from pandas.testing import assert_frame_equal
from pandas.testing import assert_index_equal
import logging
import numpy as np

d1 = pd.read_csv('test_data/regular.csv')
d2 = pd.read_csv('test_data/bad.csv')


class MyTestCase(unittest.TestCase):
    def test_find_recent(self):
        self.assertEqual(nda.find_recent_file(''), False)
        self.assertEqual(nda.find_recent_file(np.array(['.csv', '.csv'])), False)
        self.assertEqual(nda.find_recent_file(np.array(['0.csv', '1.csv'])), '1.csv')
        self.assertEqual(nda.find_recent_file(np.array(['a_b_c_0.csv', 'x_y_z_1.csv'])), 'x_y_z_1.csv')

    def test_load(self):
        self.assertEqual(nda.load_data('fake file'), False)
        assert_frame_equal(nda.load_data('NYL_FieldAgent_20210129.csv'), d1)

    def test_validate(self):
        self.assertEqual(nda.validate_file_len(d1, d1), True)
        self.assertEqual(nda.validate_file_len(d2, d1), False)

    def test_replace(self):
        assert_index_equal(nda.replace_headers(d1).columns, d1.columns)
        assert_index_equal(nda.replace_headers(d2).columns, d1.columns)

    def test_invalid_pn(self):
        self.assertEqual(nda.is_invalid_pn(''), False)
        self.assertEqual(nda.is_invalid_pn(' '), False)
        self.assertEqual(nda.is_invalid_pn('pho.num.bers'), True)
        self.assertEqual(nda.is_invalid_pn('654.2181'), True)
        self.assertEqual(nda.is_invalid_pn('804,984,4561'), True)
        self.assertEqual(nda.is_invalid_pn('804.984.4561'), False)
        self.assertEqual(nda.is_invalid_pn('804-984-4561'), False)
        self.assertEqual(nda.is_invalid_pn('804 984 4561'), False)

    def test_numbers(self):
        self.assertEqual(nda.find_valid_pn(d1), True)
        self.assertEqual(nda.find_valid_pn(d2), False)

    def test_states(self):
        self.assertEqual(nda.find_valid_state(d1), False)
        self.assertEqual(nda.find_valid_state(d2), False)

    def test_email(self):
        self.assertEqual(nda.find_valid_email(d1), True)
        self.assertEqual(nda.find_valid_email(d2), False)


if __name__ == '__main__':
    logging.disable(logging.CRITICAL)
    unittest.main()
