import math
from utils import check_instance
from wtax_item import WithholdingTaxItem, EntityItem
from typing import Any, Iterator

class WtaxCellRef:
    default_row = 37
    qtr_periods = ['O%s', 'T%s', 'Y%s']
    total_qtr_periods = ['Total_1M', 'Total_2M', 'Total_3M']

    @staticmethod
    def atc_description(row_inc: int):
        return f'A{WtaxCellRef.default_row + row_inc}'
    
    @staticmethod
    def atc_code(row_inc: int):
        return f'L{WtaxCellRef.default_row + row_inc}'
    
    @staticmethod
    def month_period(period: int, row_inc: int):
        return WtaxCellRef.qtr_periods[period - 1] % (WtaxCellRef.default_row + row_inc)
    
    @staticmethod
    def total_base(row_inc: int):
        return f'AD{WtaxCellRef.default_row + row_inc}'
    
    @staticmethod
    def total_tax(row_inc: int):
        return f'AI{WtaxCellRef.default_row + row_inc}'
    
    @staticmethod
    def total_qtr_period(period: int):
        return WtaxCellRef.total_qtr_periods[period - 1]

class WithholdingTaxInfo:
    def __init__(self, wtax_item: WithholdingTaxItem) -> None:
        self.__info = wtax_item

        month = self.info.month

        quarter = math.ceil(month / 3)
        self.__quarter_month: int = month - ((quarter * 3) - 3)

    def add_info(self, wtax_item: WithholdingTaxItem):
        self.info.base = self.info.base + wtax_item.base
        self.info.tax = self.info.tax + wtax_item.tax

    @property
    def info(self):
        return self.__info

    @property
    def quarter_month(self) -> int:
        return self.__quarter_month
    
class WithholdingTaxDict:
    def __init__(self) -> None:
        self.__withholding_taxes: dict[str, WithholdingTaxInfo] = {}

    def __getitem__(self, key: str) -> WithholdingTaxInfo | None:
        return self.__withholding_taxes.get(key)
    
    def add_info(self, wt_item: WithholdingTaxItem):
        atc_code = wt_item.atc_code
        wtax_info = self[atc_code]

        if wtax_info is None:
            wtax_info = WithholdingTaxInfo(wt_item)
            self.__withholding_taxes[atc_code] = wtax_info
        else:
            wtax_info.add_info(wt_item)

    def __iter__(self):
        return iter(self.__withholding_taxes.values())
    
    def __len__(self):
        return len(self.__withholding_taxes)
    
    @property
    def is_full(self) -> bool:
        limit = 10
        return len(self.__withholding_taxes) >= limit
    
    @property
    def total_base(self):
        total = 0
        for wtax_info in self:
            total += wtax_info.info.base

        return total
    
    @property
    def total_tax(self):
        total = 0
        for wtax_info in self:
            total += wtax_info.info.tax

        return total

class PayeeInfo:
    def __init__(self, payee_item: EntityItem, wtax_item: WithholdingTaxItem) -> None:
        self.__info = payee_item
        self.__month = str(wtax_item.month)
        self.__year = str(wtax_item.year)
        self.__wtax_dict = WithholdingTaxDict()

    def add_wtax_info(self, w_raw: WithholdingTaxItem):
        check_instance(
            ('Withholding Tax Data - Raw', WithholdingTaxItem, w_raw)
        )
        self.wtax_dict.add_info(w_raw)

    @property
    def info(self):
        return self.__info
    
    @property
    def month(self):
        return self.__month
    
    @property
    def year(self):
        return self.__year

    @property
    def wtax_dict(self):
        return self.__wtax_dict
    
    @property
    def is_wtax_full(self) -> bool:
        return self.__wtax_dict.is_full

class PayeeInfoDict:
    def __init__(self) -> None:
        self.__payees_info: dict[str, list[PayeeInfo]] = {}

    def get_recent_info(self, payee_str: str):
        payee_list = self.__payees_info.get(payee_str)
        if payee_list is None or payee_list[-1].is_wtax_full: return

        return payee_list[-1]
    
    def process_item(
            self,
            payee_item: EntityItem,
            wtax_item: WithholdingTaxItem
        ):
        payee_key = payee_item.tin + '--' + str(wtax_item.month) + '-' + str(wtax_item.year)
        payee_info = self.get_recent_info(payee_key)

        if payee_info is None:
            payee_info = PayeeInfo(payee_item, wtax_item)
            self[payee_key].append(payee_info)

        payee_info.add_wtax_info(wtax_item)

    def __iter__(self) -> Iterator[tuple[int, PayeeInfo]]:
        payee_infos_with_count: list[tuple[int, PayeeInfo]] = []

        for payee_list in self.__payees_info.values():
            payee_info_holder = tuple((payee_count + 1, payee_list[payee_count]) for payee_count in range(len(payee_list)))

            payee_infos_with_count = [*payee_infos_with_count, *payee_info_holder]

        return iter(payee_infos_with_count)
    
    def __getitem__(self, key: str):
        payee_info_list = self.__payees_info.get(key)

        if payee_info_list is None:
            payee_info_list = []
            self.__payees_info[key] = payee_info_list

        return payee_info_list