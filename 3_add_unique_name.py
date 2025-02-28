import json
import sys

#take command line argument of flat file name
if len(sys.argv) > 1:
    file_path = sys.argv[1]
    path_parts = file_path.split('\\')

    #if len(path_parts) > 3:
        #file_name = path_parts[3]
file_name = path_parts[3] #sys.argv[1].split('\\')[3]#"cboddtx.5130"
scid = file_name.split(".")[1]
display_type = file_name.split(".")[0]
json_file = f"./output/{scid}_{display_type}_obj.json"

with open(json_file, "r") as f:
    j_file = json.loads(f.read())
f.close()

unique_file = file_name.replace(".","-") + ".txt"
f = open("./output/" + unique_file, "r")
t_file = f.readlines()
f.close()

def lookup_file(j_file, shp_def, shp_id):
    for i in j_file:
        #print(i)
        if i["lines"]:
            if len(i["lines"]) > 0:
                if shp_def in i["lines"][0]:
                    i["instance"] = shp_id
                    #print(i["instance"])
    with open(json_file, "w") as f:
        f.write(json.dumps(j_file, indent=4))
    f.close()

shp_dict = {}
for i in t_file:

    if "|" in i:
        sp = i.split("|")
        #sp[0] + sp[1]
        if len(sp) >0:
            shp_id = sp[0]
            shp_def = sp[1]
            lookup_file(j_file, shp_def, shp_id)
