'''
python script that renames files from an Excel sheet

example Excel sheet included in the parent repo, setup looks as below
with fullpaths to the files you want to rename in column A
and the filename you want to rename it to in column B
__________________________________________
|         colA           |     colB      |
|fullpath1/oldname1.docx |newname1.docx  |
|fullpath2/oldname2.wav  |newname2.wav   |

'''
try:
    from openpyxl import load_workbook
except:
    print("Please install the openpyxl library for Python by typing:")
    print("pip3 install openpyxl")
    print("then, try again")
import os
import argparse
from pprint import pprint

def load_excel_workbook(path):
    '''
    loads an excel .xlsx file located at path
    '''
    workbook = load_workbook(filename=path)
    worksheet = workbook.active
    return worksheet

def iterate_worksheet(worksheet):
    '''
    goes through list of fullpaths in columnA
    tries to rename them with filename in column B
    '''
    oldpaths = []
    newnames = []
    _oldpaths = worksheet['A'][1:]
    _newnames = worksheet['B'][1:]
    for oldpath in _oldpaths:
        oldpaths.append(oldpath.value)
    for newname in _newnames:
        newnames.append(newname.value)
    for oldpath in oldpaths:
        path,oldname = os.path.split(oldpath)
        newname = newnames[oldpaths.index(oldpath)]
        newpath = os.path.join(path,newname)
        print("renaming: " + oldpath)
        print("newname :  " + newname)
        try:
            os.rename(oldpath, newpath)
        except PermissionError:
            print("It seems you don't have permission to rename that file")
            exit()
        except Exception as e:
            print("There was an error, whoops")
            print(e)
            exit()

def init():
    '''
    initializes variables and arguments
    '''
    parser = argparse.ArgumentParser(description="Renames files from an Excel file")
    parser.add_argument('-i', '--input', dest='i', help='the input Excel file')
    args = parser.parse_args()
    if not args.i:
        print("Please provide an input Excel file, .xlsx")
        exit()
    if args.i.endswith(".xls"):
        print("Please convert your file to the latest Office format, .xslx, and retry")
        exit()
    return args

def main():
    '''
    do the thing
    '''
    args = init()
    worksheet = load_excel_workbook(args.i)
    iterate_worksheet(worksheet)

if __name__ == "__main__":
    main()
