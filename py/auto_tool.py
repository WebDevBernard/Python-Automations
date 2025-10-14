import SIMPLEGUI as fsg
from SORT_RENEWAL_LIST import sort_renewal_list
from MANUAL_RENEWAL_LETTER import manual_renewal_letter
from READ_CONFIG import config
from ICBC_E_STAMP_TOOL import icbc_e_stamp_tool

gui = {
    "Sort Renewal List": sort_renewal_list,
    "Manual Renewal Letter": lambda: manual_renewal_letter(config),
    "ICBC E-Stamp Tool": icbc_e_stamp_tool,
}

fsg.run(gui)  # Only now does the GUI open
