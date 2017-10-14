import unittest
import datetime
from cleanup_script import correct_output_csv_file
from cleanup_script import process_datum
from cleanup_script import cleanup

#we need to create a class in order to test
class TestCleanUp(unittest.TestCase):


    def test_correct_output_csv_file_vI(self):
        """
        Test: A proper file name remains such and has .csv appended
        """
        self.assertEqual('test.csv',correct_output_csv_file('test'))

    def test_correct_output_csv_file_vII(self):
        """
        Test: A filename with spaces has the spaces substituted
        with underscores
        """
        self.assertEqual('test_space_file.csv',correct_output_csv_file('test space file'))

    def test_correct_output_csv_file_vIII(self):
        """
        Test: A filename with spaces has the spaces substituted
        with underscores but keeps its .csv extension
        """
        self.assertEqual('test_space_file.csv',correct_output_csv_file('test space file.csv'))

    def test_process_datum_vI(self):
        """
        Test: An empty datum column returns
        proper error msg
        """
        self.assertEqual('15',process_datum(None))

    def test_process_datum_vII(self):
        """
        Test: An empty datum column returns
        proper error msg
        """
        self.assertEqual('15',process_datum(""))

    def test_process_datum_vIII(self):
        """
        Test: Convert correcttly the given datetime object
        """
        date = datetime.datetime(2017,1,5)
        rslt = process_datum(date)
        self.assertEqual('2017-01-05',rslt[0])
        self.assertEqual('2016-2017',rslt[1])

    def test_process_datum_vIV(self):
        """
        Test: Convert correcttly the given datetime object
        """
        date = datetime.datetime(2017,10,5)
        rslt = process_datum(date)
        self.assertEqual('2017-2018',rslt[1])

    def test_clean_up_vI(self):
        """
        Test: tests the clean up method.
        The excel sheet is TransactionListingSale30.xlsx
        it should output 68 failed rows
        """
        rslt = cleanup('TransactionListingSale30.xlsx','test-output-I.csv')
        self.assertEqual(68,len(rslt))

if __name__ == '__main__':
    unittest.main()