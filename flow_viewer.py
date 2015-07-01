from pyExcelerator import *
from json import *

place_dict = {"cookie": 1,
              "duration": 2,
              "table": 3,
              "n_packets": 4,
              "n_bytes": 5,
              "priority": 6,
              "in_port": 7,
              "dl_src": 8,
              "dl_dst": 9,
              "nw_src": 10,
              "nw_dst": 11,
              "protocol": 12,
              "tp_src": 13,
              "tp_dst": 14,
              "actions": 15
              }

def convert_tags(line):
    entry = {}
    if not line.strip().startswith("cookie"):
        return None
    pos = line.index("actions")
    match = line[:pos]
    action = line[pos:]

    item_list = match.split(",")
    for item in item_list:
        item = item.strip()
        index = item.find("=")
        if index == -1:
            if item in set(['tcp', 'udp']):
                tag = 'protocol'
                val = item
            else:
                continue
        else:
            tag = item[:index]
            val = item[index+1:]
        entry[tag] = val
    entry['actions'] = action[action.find("=")+1:]

    return entry

def write_excel(file_out, items):
    w = Workbook()
    ws = w.add_sheet('Open Flow')
    line = 0
    for tag in place_dict:
        ws.write(line, place_dict[tag]-1, tag)
    line += 1
    for item in items:
        for tag in item:
            ws.write(line, place_dict[tag]-1, item[tag])
        line += 1
    w.save(file_out)

def main(file_in, file_out):
    fd = open(file_in)
    items = []
    try:
        for line in fd:
            entry = convert_tags(line)
            if entry is not None:
                items.append(entry)
        dump(items, open(file_out+".json", 'w'))
        write_excel(file_out+".xls", items)
    finally:
        fd.close()

if __name__ == '__main__':
    main('C:\Users\henry\Desktop\origin.txt', 'C:\Users\henry\Desktop\openflow')