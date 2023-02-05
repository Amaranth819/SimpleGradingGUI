import os
import glob
from collections import defaultdict

def extract_netids(root_path):
    netids = defaultdict(lambda: 0)
    txt_fns = glob.glob(root_path + '*.txt')
    
    for fn in txt_fns:
        _, fn = os.path.split(fn)
        netids[fn[11:-32]] += 1

    return sorted(list(netids.keys()))


if __name__ == '__main__':
    res = extract_netids('E://Syracuse//TA//CIS675 Spring2023//hw1//homework1//')
    print(res)
    print(len(res))