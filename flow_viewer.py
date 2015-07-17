from xlwt import *
from json import *

place_dict = {"cookie": 1,
              "duration": 2,
              "table": 3,
              "n_packets": 4,
              "n_bytes": 5,
              "hard_age": 6,
              "priority": 7,
              "in_port": 8,
              "dl_src": 9,
              "dl_dst": 10,
              "nw_src": 11,
              "nw_dst": 12,
              "protocol": 13,
              "tp_src": 14,
              "tp_dst": 15,
              "actions": 16
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
    style = easyxf('font: bold on, colour red; pattern: pattern solid, fore-colour grey25')
    line = 0
    for tag in place_dict:
        ws.write(line, place_dict[tag]-1, tag, style)
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
