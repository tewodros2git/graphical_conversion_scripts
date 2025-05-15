import datetime
import json
import sys

with open("filePathConfig.json", "r") as fconfig_file:
    fconfig = json.load(fconfig_file)

#take command line argument of flat file name
if len(sys.argv) > 1:
    file_path = sys.argv[1]
    path_parts = file_path.split(fconfig["pathSplitChar"])

    #if len(path_parts) > 3:
        #file_name = path_parts[3]
file_name = path_parts[3] #sys.argv[1].split('\\')[3]#"cboddtx.5130"
#file_path = sys.argv[2]
#print(file_name)
cls_list = open("class_list.txt","r")
cls_list = cls_list.readlines()
cls_list = [x.replace("\n","") for x in cls_list]
cls_list = set(cls_list)

#wildcard lookups with their type of RCAP object
with open("obj_types.json","r") as ot:
    obj_types = json.loads(ot.read())


scid = file_name.split(".")[1]
display_type = file_name.split(".")[0]

def obj_type_lookup(cls_name):
    for key in obj_types:
        if key in cls_name:
            return obj_types[key]
    return "NOTFOUND"

def unique_ports(lines, obj_type):
    ret_list = []

    if obj_type in ("BBE", "CMDRX"):
        none_counter = 0
        prt_in = 1
        prt_out = 1

        for line in lines:
            if "NONE," in line and "LEFT" in line:
                prt_in += 1
                new_prt = "IN" + str(prt_in) + ","
                new_line = line.replace("NONE,", new_prt)

            elif "NONE," in line and ("RIGHT" in line or "BOTTOM" in line):
                prt_out += 1
                new_prt = "OUT" + str(prt_out) + ","
                new_line = line.replace("NONE,", new_prt)
            else:
                new_line = line
            ret_list.append(new_line)


    elif obj_type == "OMNI-ANTENNA":
        none_counter = 0
        for line in lines:
            if "NONE," in line:
                none_counter += 1
                new_prt = "IN" + str(none_counter) + ","
                new_line = line.replace("NONE,",new_prt)
                ret_list.append(new_line)
            else:
                ret_list.append(line)
    else:
        ret_list = lines

    return ret_list

def parse_line(i, line_list, idx):
    try:
        props = i.split(",")

        if len(props) == 9 :
            #print(props)
            if props[0].replace(" ","") in cls_list:
                cls = props[0].replace(" ","")
                instance = props[1].replace(" ","")
                #print(instance)
                nickname = props[2].replace(" ","").replace('"','')
                att  = int(props[7].replace(" ","").replace("att:",""))
                ch = int(props[8].replace(" ","").replace(';\n','').replace("ch:",""))
                x = int(props[3])
                y =int(props[4])
                w =int(props[5])
                h = int(props[6])

                last_line = (idx + att + ch) + 1
                obj_t = obj_type_lookup(cls)

                line_dict = {
                        "obj_type":obj_t,
                        "cls":cls,
                        "instance":instance,
                        "nickname":nickname,
                        "att":att,
                        "ch":ch,
                        "x":x,
                        "y":y,
                        "h":h,
                        "w":w,
                        "lines":unique_ports(line_list[idx:last_line],obj_t)
                    }

                return line_dict
    except Exception as err:
        print(err)


dt = datetime.datetime.now()
dt = dt.strftime("%d%m%Y_%H.%M.%S")




f = open(file_path,"r")
line_list = f.readlines()

new_file_list = []
#for each line in the file

idx = 0
for i in line_list:
    #Check for Changes to line data
    line = parse_line(i,line_list,idx)
    if line is not None:
        new_file_list.append(line)

    idx += 1

f.close()

with open(f"./output/{scid}_{display_type}_obj.json","w") as f:
    f.write(json.dumps(new_file_list, indent=4))
    print("DONE")
f.close()
