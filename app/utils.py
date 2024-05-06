import typing
from typing import Any

def check_type(*args: tuple[str, typing.Any, typing.Any]):
    '''
        args: (message header: str, type object: Any, value: Any)
    '''
    for header, t_type, value in args:
        if type(value) is t_type: continue
        raise TypeError(f'{header} should be {t_type.__name__}, not {type(value).__name__}')
    
def check_instance(*args: tuple[str, typing.Any, typing.Any]):
    '''
        args: (message header: str, class object: Any, value: Any)
    '''
    for header, t_type, value in args:
        if isinstance(value, t_type): continue
        raise TypeError(f'{header} should be {t_type.__name__}, not {type(value).__name__}')
    
def merge_str_if(*args: tuple[str, bool]):
    merged_value = ''

    for value, is_merge in args:
        if not is_merge: continue
        merged_value += value
    
    return merged_value

def merge_str_if_not_empty(*values: str):
    return merge_str_if(*[(value, bool(value)) for value in values])

def process_string(value: Any, is_lower = True) -> str:
    processed_value = '' if value is None else str(value).strip()
    return processed_value.lower() if is_lower else processed_value

class ConvertTo:
    @staticmethod
    def trimm_str(value: Any):
        return process_string(value, False)
    
    @staticmethod
    def lower_str(value: Any):
        return process_string(value)
    
    @staticmethod
    def cap_str(value: Any):
        return process_string(value).upper()
    
    @staticmethod
    def proper_str(value: Any):
        return process_string(value).title()
    
    @staticmethod
    def integer_str(value: Any):
        try:
            return process_string(int(value))
        except:
            return ''