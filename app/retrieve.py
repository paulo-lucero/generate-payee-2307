from wtax_info import PayeeInfoDict
from wtax_item import InCompletePayeeInfo, EntityItem, WithholdingTaxItem
import xlwings as xw
from xlwings import Sheet, Range
import traceback
import sys

def process_src_sheet(source_sheet: Sheet) -> PayeeInfoDict:
    payee_info_dict = PayeeInfoDict()
    
    current_row = 14
    while True:
        current_row += 1

        current_ref = f'A{current_row}:O{current_row}'
        current_item = source_sheet[current_ref]
        if not isinstance(current_item, Range):
            raise ValueError(f'This {current_ref} row isn\'t found')
        
        raw_item = current_item.value
        if not isinstance(raw_item, list):
            raise ValueError(f'This {current_ref} row doesn\'t return a list')
        
        try:
            payee_item = EntityItem(
                tin=raw_item[0],
                org_name=raw_item[1],
                last_name=raw_item[2],
                first_name=raw_item[3],
                mid_name=raw_item[4],
                address=raw_item[5],
                zip_code=raw_item[6],
            )
            wtax_item = WithholdingTaxItem(
                date=raw_item[7],
                atc_code=raw_item[8],
                atc_description=raw_item[9],
                base=raw_item[10],
                tax=raw_item[11]
            )
            payee_item.add_signor(
                signor_name=raw_item[12],
                signor_position=raw_item[13],
                signor_tin=raw_item[14]
            )

            payee_info_dict.process_item(payee_item=payee_item, wtax_item=wtax_item)
        except InCompletePayeeInfo:
            break
        except Exception as e:
            print(f'{traceback.format_exc()}\nRetrieve Phase - Error on - {current_ref}')
            sys.exit()

    return payee_info_dict

def generate_payees_infos(src_path: str) -> PayeeInfoDict:
    with xw.App() as xw_app:
        source_book = xw_app.books.open(src_path)
        source_sheet = source_book.sheets[0]
        if not isinstance(source_sheet, Sheet):
            raise ValueError(f'First sheet isn\'t found at {src_path} or not a Sheet Object')
        
        payee_info_dict = process_src_sheet(source_sheet)

        return payee_info_dict