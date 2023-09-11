from typing import Any, Literal, Callable
from datetime import datetime
from utils import check_instance, merge_str_if, merge_str_if_not_empty, ConvertTo
import re

class InCompletePayeeInfo(Exception):
    pass

def get_invalid(*values: tuple[str, Any, Literal['is_empty', 'is_positive']]):
    for header, value, option in values:
        if (
                option == 'is_empty' and (
                    (not type(value) is str) or (not value.strip())
                )
            ): return (header, value, 'Invalid String')
        elif (
                option == 'is_positive' and (
                    (not isinstance(value, (int, float))) or value <= 0
                )
            ): return (header, value, 'Invalid Number')

        
def throw_error(arg: tuple[str, Any, str] | None):
    if not arg: return

    raise ValueError(f'{arg[2]} : {arg[0]} has a value of \'{arg[1]}\'')

class WithholdingTaxItem:
    def __init__(
        self,
        atc_code: Any,
        atc_description: Any,
        date: Any,
        base: Any,
        tax: Any
    ) -> None:
        processed_atc_code = ConvertTo.cap_str(atc_code)
        processed_atc_description = ConvertTo.trimm_str(atc_description)

        check_instance(
            ('Date', datetime, date)
        )
        processed_date: datetime = date
        processed_month = processed_date.month
        processed_year = processed_date.year
        
        processed_base = float(base)
        processed_tax = float(tax)

        throw_error(get_invalid(
            ('ATC Code', processed_atc_code, 'is_empty'),
            ('ATC Description', processed_atc_description, 'is_empty'),
            ('Tax Base', processed_base, 'is_positive'),
            ('Tax Amount', processed_tax, 'is_positive')
        ))

        self.atc_code: str = processed_atc_code
        self.atc_description: str = processed_atc_description
        self.month: int = processed_month
        self.year: int = processed_year
        self.base: float = processed_base
        self.tax: float = processed_tax

def validate_value_len(value: str, fixed_len: int, msg_func: Callable[[str, int, int], str], is_limit_mode: bool = False):
    '''
        args:
            value: str - value to be validated
            fixed_len: int - required number of value length
            msg_func: callable
                value: str - value to be validated
                fixed_len: int - required number of value length
                value_len: int - length of the value
    '''
    value_len = len(value)
    if (
        not is_limit_mode and value_len != fixed_len
    ) or (
        is_limit_mode and value_len > fixed_len
    ):
        raise ValueError(msg_func(value, fixed_len, value_len))
    
def tin_segment_msg_invalid(position: int):
    def msg_func(value: str, fixed_len: int, value_len: int):
        return f'Invalid length of tin segment at position of \'{position}\' and value of \'{value}\': it should be only {fixed_len} not {value_len}'
    return msg_func
    
def validate_tin_units(tin_units: list[str]):
    tin_units_len = len(tin_units)
    if tin_units_len != 3:
        raise ValueError('Invalid number of tin units: it should be only 3 not ' + str(tin_units_len))
    
    for unit_count in range(tin_units_len):
        validate_value_len(
            tin_units[unit_count],
            3,
            tin_segment_msg_invalid(unit_count + 1)
        )

def validate_tin_format(tin: str, pattern: str, branch_optional: bool = False):
    tin_list = re.split(pattern, tin)

    not_length_four = len(tin_list) != 4

    if (
        (not branch_optional and not_length_four) or
        (branch_optional and len(tin_list) != 3 and not_length_four)
    ):
        raise ValueError('Invalid Tin: ' + tin)
    
    validate_tin_units(tin_list[:3])

    if not not_length_four: validate_value_len(tin_list[3], 5,tin_segment_msg_invalid(4))

def entity_name_msg_invalid(value: str, fixed_len: int, value_len: int):
    return f'Invalid length of entity name with a value of \'{value}\': it should be only {fixed_len} not {value_len}'

def address_msg_invalid(value: str, fixed_len: int, value_len: int):
    return f'Invalid length of address with a value of \'{value}\': it should be only {fixed_len} not {value_len}'

def zip_code_msg_invalid(value: str, fixed_len: int, value_len: int):
    return f'Invalid length of zip codes digits with a value of \'{value}\': it should be only {fixed_len} not {value_len}'

class EntityItem:
    __tin_reg_ex = '\\s*-\\s*|\\s+|\\D'

    def __init__(
        self,
        tin: Any,
        org_name: Any,
        last_name: Any,
        first_name: Any,
        mid_name: Any,
        address: Any,
        zip_code: Any
    ) -> None:
        processed_tin = ConvertTo.trimm_str(tin)

        if get_invalid(('Tin', processed_tin, 'is_empty')):
            raise InCompletePayeeInfo('Empty Tin')
        
        validate_tin_format(processed_tin, EntityItem.__tin_reg_ex)


        processed_org_name = ConvertTo.cap_str(org_name)
        processed_last_name = ConvertTo.cap_str(last_name)
        processed_first_name = ConvertTo.cap_str(first_name)
        processed_mid_name = ConvertTo.cap_str(mid_name)


        processed_address = ConvertTo.cap_str(address)
        validate_value_len(
            processed_address,
            84,
            address_msg_invalid,
            True
        )


        processed_zip = ConvertTo.integer_str(zip_code)
        if processed_zip: validate_value_len(processed_zip, 4, zip_code_msg_invalid)
        

        org_result = get_invalid(('Organization Name', processed_org_name, 'is_empty'))
        name_result = get_invalid(
            ('Last Name', processed_last_name, 'is_empty'),
            ('First Name', processed_first_name, 'is_empty')
        )

        if org_result and name_result:
            header = f'{org_result[0]} and {name_result[0]}'
            value_result = f'None or empty string'
            category_result = 'Empty'
            throw_error((header, value_result, category_result))
            
        
        self.tin: str = processed_tin
        self.org_name: str = processed_org_name
        self.last_name: str = processed_last_name
        self.first_name: str = processed_first_name
        self.mid_name: str = processed_mid_name
        self.address: str = processed_address
        self.zip_code: str = processed_zip
        self.__signor_info: None | str = None
        self.__signor_tin: None | str = None

        validate_value_len(
            self.prod_entity_name,
            91,
            entity_name_msg_invalid,
            True
        )

    def add_signor(self, signor_name, signor_tin, signor_position):
        proc_signor_name = ConvertTo.cap_str(signor_name)
        proc_signor_position = ConvertTo.cap_str(signor_position)

        has_signor_name = bool(proc_signor_name)
        has_signor_position = bool(proc_signor_position)

        proc_signor_tin = ConvertTo.trimm_str(signor_tin)
        if proc_signor_tin: validate_tin_format(proc_signor_tin, EntityItem.__tin_reg_ex, True)

        self.__signor_info = merge_str_if(
            (proc_signor_name, has_signor_name),
            (' - ', has_signor_name and has_signor_position),
            (proc_signor_position, has_signor_position)
        )
        self.__signor_tin = proc_signor_tin

        return self

    @property
    def signor_info(self):
        return ConvertTo.cap_str(self.__signor_info)
    
    @property
    def signor_tin(self):
        return ConvertTo.trimm_str(self.__signor_tin)

    @property
    def raw_entity_name(self):
        return self.org_name + self.last_name + self.first_name + self.mid_name
    
    @property
    def indiv_name(self):
        last_name = self.last_name
        has_last_name = bool(last_name)

        given_name = merge_str_if_not_empty(self.first_name, self.mid_name)
        has_given_name = bool(given_name)

        return merge_str_if(
            (last_name, has_last_name),
            (', ', has_last_name and has_given_name),
            (given_name, has_given_name)
        )
    
    @property
    def prod_entity_name(self):
        org_name = self.org_name
        has_org_name = bool(org_name)

        indiv_name = self.indiv_name
        has_indiv_name = bool(indiv_name)

        return merge_str_if(
            (org_name, has_org_name),
            (' - ', has_org_name and has_indiv_name),
            (indiv_name, has_indiv_name)
        )
    
    @property
    def tin_segments(self):
        return re.split(EntityItem.__tin_reg_ex, self.tin)