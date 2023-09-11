from pathlib import Path
import xlwings
from utils import ConvertTo

def retrieve_source_path():
    book_path = ConvertTo.trimm_str(input('Please provide the complete source path : '))

    if not Path(book_path).exists():
        raise FileNotFoundError(f'This source path \'{book_path}\' doesn\'t exists')
    
    return book_path

def retrieve_drop_path(book_path: str):
    with xlwings.App() as xw_app:
        souce_book = xw_app.books.open(book_path)
        source_sheet = souce_book.sheets[0]
        if not isinstance(source_sheet, xlwings.Sheet):
            raise TypeError(f'The first sheet of this \'{book_path}\' doesn\'t return a Sheet Object')
        
        drop_path = ConvertTo.trimm_str(source_sheet['DROP_PATH'].value)
        
        if Path(drop_path).exists():
            raise FileExistsError(f'This \'{drop_path}\' drop path directory exists, kindly delete it first if not needed')
        
        Path(drop_path).mkdir(parents=True)
        
        return drop_path