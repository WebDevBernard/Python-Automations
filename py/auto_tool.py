import SIMPLEGUI as fsg
from SORT_RENEWAL_LIST import sort_renewal_list
from MANUAL_RENEWAL_LETTER import manual_renewal_letter
from READ_CONFIG import config

gui = {
    'Sort Renewal List': sort_renewal_list,
    'Manual Renewal Letter': lambda: manual_renewal_letter(config),
}

fsg.run(gui)  # Only now does the GUI open