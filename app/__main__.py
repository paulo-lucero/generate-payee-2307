from retrieve import generate_payees_infos
from process import generate_forms
from generate_path import retrieve_source_path, retrieve_drop_path
from payor_info import generate_payor_info

book_path = retrieve_source_path()
drop_path = retrieve_drop_path(book_path)

payor_item = generate_payor_info(book_path)

payee_info_dict = generate_payees_infos(book_path)

generate_forms(drop_path, payee_info_dict, payor_item)

print('Test Finished')