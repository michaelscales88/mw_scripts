import os
import pyexcel as pe


def main(base_dir):
    sheet = pe.get_book(file_name=base_dir + "\\CONFIG.xlsx")
    print(type(sheet))
    print(sheet)

if __name__ == '__main__':
    main(os.getcwd())
