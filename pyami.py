import sys
import pprint
import json
import sg_excel.app as export_exl

def main(args):
    # Make sure we have only one arg, the URL
    if len(args) != 1:
        return 1
    export_exl.export(args[0])
    
if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))