import datetime
import json
import sys
import os

cls_dict = {}
cls_list = []
port_list = []

def parse_line(i, line_list, idx):
    props = i.split(",")
    if len(props) == 9:
        c = props[0]
#         if len(c) > 25:
#             if "OMUX" in c:
#                 c = c[:15] + "_OMUX"
#             elif "CMUX" in c:
#                 c = c[:15] + "_CMUX"
#             elif "IMUX" in c:
#                 c = c[:15] + "_IMUX"
#                 #print(c)
        att = int(props[7].replace(" ", "").replace("att:", ""))
        ch = int(props[8].replace(" ", "").replace(';\n', '').replace("ch:", ""))
        last_line = idx + att + ch + 1

        if c not in cls_dict:
            if "    " in c:
                pass
            else:
                av_ports = line_list[idx + att + 1:last_line]
                pts = [port_list.append(l.split(",")[0].strip()) for l in av_ports]
                cls_dict[c] = {"ports": pts}
                cls_list.append(c)

series = sys.argv[1]
scid_dir = "series" + "\\" + series
print(scid_dir)
scid_list = os.listdir(scid_dir)

def parse_file(d):
    for f in os.listdir(d):
        f = d + "\\" + f
        fi = open(f, "r")
        line_list = fi.readlines()
        idx = 0
        for i in line_list:
            parse_line(i, line_list, idx)
            idx += 1

# Parsing each file
for scid in scid_list:
    parse_file(scid_dir + "\\" + scid)

# Remove duplicates by converting lists to sets
u_list = set(cls_list)
port_list = set(port_list)

# Append unique classes to 'class_list.txt'
if os.path.exists("class_list.txt"):
    with open("class_list.txt", "r+") as fil:
        existing_classes = set([line.strip() for line in fil.readlines()])
else:
    existing_classes = set()

new_classes = u_list - existing_classes  # Find new classes

with open("class_list.txt", "a") as fil:
    for u in new_classes:
        fil.write(u + "\n")

# Append unique ports to 'available_ports.txt'
if os.path.exists("available_ports.txt"):
    with open("available_ports.txt", "r+") as ap:
        existing_ports = set([line.strip() for line in ap.readlines()])
else:
    existing_ports = set()

new_ports = port_list - existing_ports  # Find new ports

with open("available_ports.txt", "a") as ap:
    for s in new_ports:
        ap.write(s + "\n")
