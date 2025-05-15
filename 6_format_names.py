import json
import math
import os
import sys

# Read the JSON file to get display name
with open('config.json', 'r') as file:
    config = json.load(file)

with open("filePathConfig.json", "r") as fconfig_file:
    fconfig = json.load(fconfig_file)

if len(sys.argv) > 1:
    file_path = sys.argv[1]
    path_parts = file_path.split(fconfig["pathSplitChar"])

    #if len(path_parts) > 3:
    # file_name = path_parts[3]

file_name = path_parts[3]
scid = file_name.split(".")[1]
display_type = file_name.split(".")[0]
svg_name = "./output/" + config + ".svg"
json_name = f"./output/{scid}_{display_type}_obj.json"
#print(json_name)

with open(json_name, "r") as f:
    json_file = f.read()
f.close()

json_file = json.loads(json_file)

ss = open(svg_name, "r")
svg_lines = ss.readlines()
#print(svg_lines)
final_svg = []
obj_dict = {}

for i in json_file:
    name = i["instance"]
    #print(name)
    if name not in obj_dict:
        print(name)
        obj_dict[name] = True
        #print(obj_dict)

#check lines for shapes and replace them if exists
for line in svg_lines:
    for key in obj_dict:
        if key in line:
            #print(key)
            idx = line.find(key)
            if idx > -1:
                replace_val = key.replace("-","_")
                line = line.replace(key, replace_val)
    final_svg.append(line)



with open(svg_name, "w") as f:
    f.writelines(final_svg)
f.close()