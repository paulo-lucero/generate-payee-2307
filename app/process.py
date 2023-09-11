import xlwings
from pathlib import Path
from wtax_item import EntityItem
from wtax_info import PayeeInfoDict, PayeeInfo, WithholdingTaxDict, WtaxCellRef
from typing import TypedDict, Literal
import re
from calendar import monthrange
import traceback
import sys

def convert_to_double_digit(value: str):
    return ('0' + value) if len(value) == 1 else value


class PayeeFormatInput:
    @staticmethod
    def period_mmdd(month: str, day: int):
        double_digit_month = convert_to_double_digit(month)
        double_digit_day = convert_to_double_digit(str(day))
        return f' {double_digit_month[0]}  {double_digit_month[1]}   {double_digit_day[0]}   {double_digit_day[1]}'
    
    @staticmethod
    def period_yyyy(year: str):
        return f' {year[0]}  {year[1]}   {year[2]}  {year[3]}'
    
    @staticmethod
    def tin_unit(unit: str):
        return f' {unit[0]}  {unit[1]}   {unit[2]}'
    
    @staticmethod
    def branch_unit(unit: str):
        return f'  {unit[0]}   {unit[1]}   {unit[2]}   {unit[3]}   {unit[4]}'
    
    @staticmethod
    def zip_code(zip_code: str):
        if len(zip_code) < 4: return ''
        return f' {zip_code[0]}  {zip_code[1]}   {zip_code[2]}  {zip_code[3]}'
    

def perform_write(source_sheet: xlwings.Sheet, ref_name: str, value: str | int, sheet_type: Literal['range'] | Literal['shape']):
    if sheet_type == 'range':
        xw_range = source_sheet[ref_name]
        if not isinstance(xw_range, xlwings.Range):
            raise TypeError(f'This range name or reference {ref_name} doesn\'t return a Range object')
        xw_range.value = value
    elif sheet_type == 'shape':
        xw_shape = source_sheet.shapes[ref_name]
        if not isinstance(xw_shape, xlwings.Shape):
            raise TypeError(f'This range name or reference {ref_name} doesn\'t return a Shape object')
        xw_shape.text = value


def to_2dec(value: float | int):
    return '{:.2f}'.format(value)


class SettingsProcessWTaxInfos(TypedDict):
    source_sheet: xlwings.Sheet
    wtax_dict: WithholdingTaxDict


def process_wtax_infos(settings: SettingsProcessWTaxInfos):
    source_sheet = settings['source_sheet']
    wtax_dict = settings['wtax_dict']

    wtax_count = 0
    period_month = None
    for wtax_info in wtax_dict:
        wtax_count += 1
        period_month = wtax_info.quarter_month

        wtax_item = wtax_info.info
        base = to_2dec(wtax_item.base)
        tax = to_2dec(wtax_item.tax)

        perform_write(source_sheet, WtaxCellRef.atc_description(wtax_count), wtax_item.atc_description, 'range')
        perform_write(source_sheet, WtaxCellRef.atc_code(wtax_count), wtax_item.atc_code, 'range')
        perform_write(source_sheet, WtaxCellRef.month_period(period_month, wtax_count), base, 'range')
        perform_write(source_sheet, WtaxCellRef.total_base(wtax_count), base, 'range')
        perform_write(source_sheet, WtaxCellRef.total_tax(wtax_count), tax, 'range')
    
    if not isinstance(period_month, int):
        raise ValueError('Period Month shouldn\'t be \'None\'')
    
    total_base = to_2dec(wtax_dict.total_base)
    total_tax = to_2dec(wtax_dict.total_tax)

    perform_write(source_sheet, WtaxCellRef.total_qtr_period(period_month), total_base, 'range')
    perform_write(source_sheet, 'Total_Base', total_base, 'range')
    perform_write(source_sheet, 'Total_Tax', total_tax, 'range')


class SettingsWriteSignor(TypedDict):
    source_sheet: xlwings.Sheet
    ref_signor_info: str
    ref_signor_tin: str
    signor_item: EntityItem


def write_signor(settings: SettingsWriteSignor):
    source_sheet = settings['source_sheet']
    signor_item = settings['signor_item']
    
    perform_write(source_sheet, settings['ref_signor_info'], signor_item.signor_info, 'range')
    perform_write(source_sheet, settings['ref_signor_tin'], signor_item.signor_tin, 'range')


def write_tin_segments(source_sheet: xlwings.Sheet, tin_segments: list[str], tin_segment_refs: list[str]):
    for segment_index in range(len(tin_segment_refs)):
        perform_write(
            source_sheet,
            tin_segment_refs[segment_index],
            PayeeFormatInput.tin_unit(tin_segments[segment_index]),
            'shape'
        )


class SettingsWriteEntityInfo(TypedDict):
    source_sheet: xlwings.Sheet
    entity_item: EntityItem
    ref_tin_segments: list[str]
    ref_branch: str
    ref_entity_name: str
    ref_entity_address: str
    ref_zip_code: str


def write_entity_info(settings: SettingsWriteEntityInfo):
    source_sheet = settings['source_sheet']
    entity_item = settings['entity_item']
    tin_segments = entity_item.tin_segments

    write_tin_segments(source_sheet, tin_segments, settings['ref_tin_segments'])
    perform_write(source_sheet, settings['ref_branch'], PayeeFormatInput.branch_unit(tin_segments[3]), 'shape')
    perform_write(source_sheet, settings['ref_entity_name'], entity_item.prod_entity_name, 'shape')
    perform_write(source_sheet, settings['ref_entity_address'], entity_item.address, 'shape')
    perform_write(source_sheet, settings['ref_zip_code'], PayeeFormatInput.zip_code(entity_item.zip_code), 'shape')


class SettingsWritePayee(TypedDict):
    file_path: str
    payee_info: PayeeInfo
    payor_item: EntityItem


def write_payee(settings: SettingsWritePayee):
    payor_item = settings['payor_item']
    payee_info = settings['payee_info']
    payee_item = payee_info.info
    month = payee_info.month
    year = payee_info.year
    last_day = monthrange(int(year), int(month))[1]

    source_path = str((Path(__file__) / '../../format/2307.xlsx').resolve())

    if not Path(source_path).exists():
        raise FileNotFoundError(f'The source 2307 form path: \'{source_path}\' doesn\'t exist')
    
    with xlwings.App() as xw_app:
        source_book = xw_app.books.open(source_path)
        source_sheet = source_book.sheets[0]
        if not isinstance(source_sheet, xlwings.Sheet):
            raise TypeError(f'First sheet isn\'t at \'{source_path}\' a Sheet Type ')
        
        perform_write(source_sheet, 'Return_Period_From_mmdd', PayeeFormatInput.period_mmdd(month, 1), 'shape')
        perform_write(source_sheet, 'Return_Period_From_yyyy', PayeeFormatInput.period_yyyy(year), 'shape')
        perform_write(source_sheet, 'Return_Period_To_mmdd', PayeeFormatInput.period_mmdd(month, last_day), 'shape')
        perform_write(source_sheet, 'Return_Period_To_yyyy', PayeeFormatInput.period_yyyy(year), 'shape')

        write_entity_info({
            'source_sheet': source_sheet,
            'entity_item': payee_item,
            'ref_tin_segments': ['Payee_Tin_1', 'Payee_Tin_2', 'Payee_Tin_3'],
            'ref_branch': 'Payee_Branch',
            'ref_entity_name': 'Payee_Name',
            'ref_entity_address': 'Payee_Address',
            'ref_zip_code': 'Payee_Zip_Code'
        })

        write_entity_info({
            'source_sheet': source_sheet,
            'entity_item': payor_item,
            'ref_tin_segments': ['Payor_Tin_1', 'Payor_Tin_2', 'Payor_Tin_3'],
            'ref_branch': 'Payor_Branch',
            'ref_entity_name': 'Payor_Name',
            'ref_entity_address': 'Payor_Address',
            'ref_zip_code': 'Payor_Zip_Code'
        })

        process_wtax_infos({
            'source_sheet': source_sheet,
            'wtax_dict': payee_info.wtax_dict
        })

        write_signor({
            'source_sheet': source_sheet,
            'signor_item': payor_item,
            'ref_signor_info': 'Signor_Payor_Info',
            'ref_signor_tin': 'Signor_Payor_Tin'
        })

        write_signor({
            'source_sheet': source_sheet,
            'signor_item': payee_item,
            'ref_signor_info': 'Signor_Payee_Info',
            'ref_signor_tin': 'Signor_Payee_Tin'
        })

        source_book.save(path=settings['file_path'])


class SettingsProcessPayee(TypedDict):
    drop_path: str
    payee_dict: PayeeInfoDict
    payor_item: EntityItem


def generate_file_name(payee_info: PayeeInfo, count: int) -> str:
    raw_entity_name = payee_info.info.raw_entity_name

    proc_org_name = payee_info.info.tin + '-' + re.sub('\\W', '', raw_entity_name)

    return '_'.join((payee_info.year, payee_info.month, proc_org_name, str(count))) + '.xlsx'


def process_payees(settings: SettingsProcessPayee):
    for payee_set in settings['payee_dict']:
        payee_info = payee_set[1]
        file_name = generate_file_name(payee_info, payee_set[0])
        file_path = str((Path(settings['drop_path']) / file_name).resolve())
        
        try:
            write_payee({
                'file_path': file_path,
                'payee_info': payee_info,
                'payor_item': settings['payor_item']
            })
        except Exception as e:
            print(f'{traceback.format_exc()}\nProcess Phase - Error on: {file_path}')
            sys.exit()


def generate_forms(drop_path: str, payee_dict: PayeeInfoDict, payor_item: EntityItem):
    process_payees({
        'drop_path': drop_path,
        'payee_dict': payee_dict,
        'payor_item': payor_item
    })