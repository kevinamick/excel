import sys
import xlrd
import logging
import os
from convert_file import ConvertFile


def print_help():
    print("""
        Usage:\n
        excel.py <input_file> <schema_file> <export file> \n
        Please be sure to have all files in excel directory
        """)


class Excel(object):

    def __init__(self, args):
        self.args = args
        self.input_file = None
        self.schema_file = None
        self.export_file_name = None

    def run(self):
        if len(self.args) == 1:
            print_help()
            sys.exit(1)

        if len(self.args) < 4:
            print('Please enter proper input file name, schema file name, and export file name')
            sys.exit(1)

        logging.info('running')
        self.input_file = os.path.join(os.getcwd(), self.args[1])
        self.schema_file = self.args[2]
        self.export_file_name = self.args[3]

        writer = ConvertFile(self.input_file, self.schema_file, self.export_file_name)
        writer.write_to()


def main():
    logging.basicConfig(level=logging.INFO)
    logging.info('Started')
    args = sys.argv
    excel = Excel(args)
    excel.run()
    logging.info('Finished')

if __name__ == '__main__':
    main()
