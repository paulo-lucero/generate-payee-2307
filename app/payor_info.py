import xlwings
from wtax_item import EntityItem

def generate_payor_info(book_path: str):
    with xlwings.App() as xw_app:
        xw_app = xw_app.books.open(book_path)
        source_sheet = xw_app.sheets[0]
        if not isinstance(source_sheet, xlwings.Sheet):
            raise TypeError(f'First sheet at source book isn\'t a Sheet object')
        
        return EntityItem(
            tin=source_sheet['PAYOR_TIN'].value,
            org_name=source_sheet['PAYOR_ORG_NAME'].value,
            last_name=source_sheet['PAYOR_LAST_NAME'].value,
            first_name=source_sheet['PAYOR_FIRST_NAME'].value,
            mid_name=source_sheet['PAYOR_MID_NAME'].value,
            address=source_sheet['PAYOR_ADDRESS'].value,
            zip_code=source_sheet['PAYOR_ZIP_CODE'].value
        ).add_signor(
            signor_name=source_sheet['SIGNOR_NAME'].value,
            signor_position=source_sheet['SIGNOR_POSITION'].value,
            signor_tin=source_sheet['SIGNOR_TIN'].value
        )