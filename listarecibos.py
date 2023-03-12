# Python3

# pylint: disable=missing-function-docstring

""" Dump old XLS format
"""

import sys
import xlibre.xolder
from waxpage.redit import char_map

def main():
    goodies(sys.argv[1:])

def goodies(args):
    param = args if args else ["ListaRecibos.xls"]
    if len(param) > 1:
        print("Only one parameter!")
        return None
    return process(param)

def simpler_head(astr, only_ascii=True):
    alist = astr.split(' ')
    astr = ''.join(simpler(alist) if only_ascii else alist)
    return astr

def simpler(astr):
    return char_map.simpler_ascii(astr)

def process(param) -> int:
    code = get_them(param[0])
    return code

def get_them(fname:str):
    """ Dump old XLS format """
    wbk = xlibre.xolder.ABook("ListaRecibos")
    print(":::", fname, "; as:", wbk.get_aname(), "; is:", wbk.get_book_type())
    wbk.load(fname)
    if wbk.ibook is None:
        print("Cannot load:", fname)
        return 2
    wbk.ibook.first()
    print("# Displaying:", wbk.ibook.current.get_aname())
    cont = wbk.ibook.current.lines()
    dgci_dump(cont, wbk)
    return 0

def dgci_dump(there, wbk):
    assert wbk is not None, "wbk"
    head, cont = there[0], there[1:]
    print("#", [simpler_head(field) for field in head])
    for line in cont[::-1]:
        print(line)
    return 0

if __name__ == "__main__":
    main()
