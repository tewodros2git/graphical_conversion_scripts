import json
import math
import os
import sys
write_outputs_to_file = True

if len(sys.argv) > 1:
    file_path = sys.argv[1]
    path_parts = file_path.split('\\')

file_name = path_parts[3]
scid = file_name.split(".")[1]
display_type = file_name.split(".")[0]
json_file = f"./output/{scid}_{display_type}_obj.json"

with open(json_file, "r") as f:
    j_file = f.read()
f.close()

j_file = json.loads(j_file)

def get_offset(obj_type, name, port, j):
    #print("Processing line:", j)
    if "hot-btn" not in j:
        try:
            offset = j.split(",")[2].replace(" ","")
            offset_side = j.split(",")[1].replace(" ","")
            num_of_lines = j.split(",")[4].replace(" ","")
            line_split = j.split(",")
            len_split = len(line_split)
            lines = []
            for tt in range(5, len_split):
                lines.append(line_split[tt].replace(" ","").replace("\n","").replace(";",""))

            #print("Offset:", offset)
            #print("Offset Side:", offset_side)
           # print("Number of Lines:", num_of_lines)
           # print("Lines:", lines)

            return (offset, offset_side, num_of_lines, lines)
        except Exception as e:
            v = 2
            #print(e)
            #print(obj_name, obj_type,port)


def calc_start_end(x, y, h, w, port, offset, offset_side, num_of_lines, lines, obj_name):
    """
    Finds the x,y of all the ports of a shape
    """
    if lines[0] !="0":

        x1_half = math.floor(w / 2)

        y1_half = math.floor(h / 2)

        if offset_side == "NONE":  # Handle the case where port details are "NONE" for "HOT-ARROW" objects
                x1_set = x
                y1_set = y
                x1_start = x1_set
                y1_start = y1_set
        else:
            try:
                offset_1 = int(offset)
            except Exception as e:
                #print(obj_name, e, offset, lines)
                return None
            if offset_side == "RIGHT":
                x1_set = (x + x1_half)

                y1_set = (y + y1_half) - offset_1
                x1_start = x1_set
                y1_start = y1_set
                for jj in range(len(lines)):
                    #even iter means calc on x, odd means y
                    if jj % 2 == 0:
                        x1_set += int(lines[jj])
                    else:
                        y1_set += int(lines[jj])
            elif offset_side == "LEFT":
                y1_set = (y + y1_half) - offset_1
                x1_set = (x - x1_half)
                x1_start = x1_set
                y1_start = y1_set
                for jj in range(len(lines)):
                    #even iter means calc on x, odd means y
                    if jj % 2 == 0:
                        x1_set += int(lines[jj])
                    else:
                        y1_set += int(lines[jj])
            elif offset_side == "TOP":
                y1_set = (y + y1_half)
                if "COMM_DUAL" in obj_name:
                    x1_set = (x + x1_half) + offset_1
                else:
                    x1_set = (x + x1_half) - offset_1
                x1_start = x1_set
                y1_start = y1_set
                for jj in range(len(lines)):
                    #even iter means calc on y, odd means x
                    if jj % 2 == 0:
                        y1_set += int(lines[jj])
                    else:
                        x1_set += int(lines[jj])
            elif offset_side == "BOTTOM":
                y1_set = (y - y1_half)
                if "COMM_DUAL" in obj_name:
                    x1_set = (x - x1_half) - offset_1
                else:
                    x1_set = (x - x1_half) + offset_1
                x1_start = x1_set
                y1_start = y1_set
                for jj in range(len(lines)):
                    #even iter means calc on x, odd means y
                    if jj % 2 == 0:
                        y1_set += int(lines[jj])
                    else:
                        x1_set += int(lines[jj])
        #print(x, y, h, w, port, offset, offset_side, num_of_lines, lines, obj_name)
        return {
            "obj_name":obj_name,
            "prt":port,
            "start_x": x1_start,
            "start_y": y1_start,
            "end_x":x1_set,
            "end_y":y1_set,
            "offset_side": offset_side
        }
    else :
        return None


def dist(end, start):
    if (end == None or start == None):
        return 100000
    x = math.pow(end["end_x"] - start["start_x"], 2)
    y = math.pow(end["end_y"] - start["start_y"], 2)
    return math.sqrt(x + y)

def prune_duplicate(li):
    seen = set()
    used_ports = set()
    unique_arr = []

    for item in li:
        key = (item['shp1'], item['shp1_port'], item['shp2'], item['shp2_port'])
        key2 = (item['shp2'], item['shp2_port'], item['shp1'], item['shp1_port'])
        if key not in seen and key2 not in seen:
            p1 = item['shp1'] + item['shp1_port']
            p2 = item['shp2'] + item['shp2_port']
            used_ports.add(p1)
            used_ports.add(p2)
            unique_arr.append(item)
            seen.add(key)

    ret = []
    for item in unique_arr:
        found1 = False
        found2 = False
        i1 = item['shp1'] + item['shp1_port']
        i2 = item['shp2'] + item['shp2_port']
        for item2 in unique_arr:
            if item == item2:
                continue
            p1 = item2['shp1'] + item2['shp1_port']
            p2 = item2['shp2'] + item2['shp2_port']
            if i1 == p1 or i1 == p2:
                found1 = True
            if i2 == p1 or i2 == p2:
                found2 = True
        if found1 == False or found2 == False:
            ret.append(item)

    return ret

def get_cxn_obj(cxn_list):
    """
    Finds the connection points based on x,y \n
    Writes to json file used in node script
    """
    all_cxns = []

    shp1 = cxn_list
    shp2 = cxn_list
    logger = {

    }
    icnt = 0
    for i in shp1:
        #print(i)
        jcnt = 0 
        found_conn = False
        if i is not None and ("_365" in i["obj_name"] or "_353" in i["obj_name"]):
            print("Process " + i["obj_name"] + " " + i["prt"])
        for j in shp2:
            #if (i["start_x"] == j["end_x"] or i["start_x"] - j["end_x"] <=1)  and (i["start_y"] == j["end_y"] or i["start_y"] - j["end_y"] <=1):
            if i and j is not None:#"start_x" in i and "end_x" in j and "start_y" in i and "end_y" in j:
                if (i["prt"] == "HOT-ARROW" and i["offset_side"] != "NONE") or (j["prt"] == "HOT-ARROW" and j["offset_side"] != "NONE"):
                    continue
                x_diff = i["start_x"] - j["end_x"]
                y_diff = i["start_y"] - j["end_y"]
                #print(i, j)
                try: 
                    if ((-6 <= x_diff <= 6 )  and (-1 <= y_diff <= 1 ) and i["obj_name"] != j["obj_name"]):
                        cxns = {
                            "shp1": i["obj_name"],
                            "shp1_port": i["prt"],#convert_port_val(i["prt"],i["obj_name"]),
                            "shp1_side": i["offset_side"],
                            "shp2": j["obj_name"],
                            "shp2_port": j["prt"],#convert_port_val(j["prt"],j["obj_name"]),
                            "shp2_side": j["offset_side"]
                        }
                        if "COMM_DUAL" in cxns["shp1"]:
                            cxns["shp1_port"] = cxns["shp1_side"]
                        if "COMM_DUAL" in cxns["shp2"]:
                            cxns["shp2_port"] = cxns["shp2_side"]
                        if "COMM" in cxns["shp1"]:
                            print(i)
                        if "COMM" in cxns["shp2"]:
                            print(j)

                        o1_name = i["obj_name"] + i["prt"]
                        o2_name = j["obj_name"] + j["prt"]
                        f_key = o1_name + o2_name
                        b_key = o2_name + o1_name
                        x_key = o1_name
                        y_key = o2_name
                        
                        if (f_key not in logger and b_key not in logger and "STUB" not in cxns["shp1"] and "STUB" not in cxns["shp2"]):
                            if i is not None and ("_365" in i["obj_name"] or "_353" in i["obj_name"]):
                                print("found " + i["obj_name"] + " " + i["prt"] + " -> " + j["obj_name"] + " " + j["prt"])
                            found_conn = True
                            all_cxns.append(cxns)
                            logger[f_key] = True
                            logger[x_key] = True
                            logger[y_key] = True
                    
                except Exception as e:
                    print(e, i, j)
            else:
                #print(icnt,shp1[icnt], jcnt,shp2[jcnt], i , j)
                x_diff = 0
                y_diff = 0
            jcnt += 1
        if found_conn == False:
            # no connection found, increasing radius, dont find closest for hot arrows
            if i is not None and "HOT-ARROW" not in i["obj_name"]:
                if "_365" in i["obj_name"] or "_353" in i["obj_name"]:
                    print("Need to find closest to " + i["obj_name"])
                closest = None
                closest_dist = 100000
                for k in shp2:
                    if i and k is not None:
                        if (i["prt"] == "HOT-ARROW" and i["offset_side"] != "NONE") or (k["prt"] == "HOT-ARROW" and k["offset_side"] != "NONE"):
                            continue
                        if k != None and "STUB" not in k["obj_name"] and k["obj_name"] != i["obj_name"]:
                            d = dist(i, k)
                            if d < closest_dist:
                                closest_dist = d
                                closest = k
                if closest != None:
                    if "_365" in i["obj_name"] or "_353" in i["obj_name"]:
                        print("   Closest is " + i["obj_name"] + " " + i["prt"] + " -> " + closest["obj_name"] + " " + closest["prt"])
                    cxns = {
                        "shp1": i["obj_name"],
                        "shp1_port": i["prt"],#convert_port_val(i["prt"],i["obj_name"]),
                        "shp1_side": i["offset_side"],
                        "shp2": closest["obj_name"],
                        "shp2_port": closest["prt"],#convert_port_val(j["prt"],j["obj_name"]),
                        "shp2_side": closest["offset_side"]
                    }
                    if "COMM_DUAL" in cxns["shp1"]:
                        cxns["shp1_port"] = cxns["shp1_side"]
                    if "COMM_DUAL" in cxns["shp2"]:
                        cxns["shp2_port"] = cxns["shp2_side"]
                            
                    o1_name = i["obj_name"] + i["prt"]
                    o2_name = closest["obj_name"] + closest["prt"]
                    f_key = o1_name + o2_name
                    b_key = o2_name + o1_name
                    x_key = o1_name
                    y_key = o2_name
                    
                    if (f_key not in logger and b_key not in logger and "STUB" not in cxns["shp1"] and "STUB" not in cxns["shp2"]):
                        found_conn = True
                        all_cxns.append(cxns)
                        logger[f_key] = True
                        logger[x_key] = True
                        logger[y_key] = True
        icnt += 1

    all_cxns = prune_duplicate(all_cxns)

    no_stubs = []
    stubs = []
    #remove channel stubs
    idx = 0
    stu = {}
    for ll in all_cxns:
        if "CHANNEL-STUB"  in ll["shp1"] :
            #print("channel stub1 ", ll["shp1"])
            k = ll["shp1"]
            if k not in stu:

                stu[k] = []
                stu[k].append({"shp":ll["shp2"], "prt":ll["shp2_port"], "side":ll["shp2_side"]})
            else:
                stu[k].append({"shp":ll["shp2"], "prt":ll["shp2_port"], "side":ll["shp2_side"]})
        if "CHANNEL-STUB"  in ll["shp2"]:
            #print("channel stub2 ", ll["shp2"])
            k = ll["shp2"]
            if k not in stu:

                stu[k] =[]
                stu[k].append({"shp":ll["shp1"], "prt":ll["shp1_port"], "side":ll["shp1_side"]})
            else:
                stu[k].append({"shp":ll["shp1"], "prt":ll["shp1_port"], "side":ll["shp1_side"]})
        else:
            no_stubs.append(ll)
        idx +=0

    if len(stu) > 0:
        #print(json.dumps(stu, indent=4))
        for k in stu:
            if len(stu[k]) > 1:
                shp_1 = stu[k][0]["shp"]
                prt_1 = stu[k][0]["prt"]
                side_1 = stu[k][0]["side"]
                shp_2 = stu[k][1]["shp"]
                prt_2 = stu[k][1]["prt"]
                side_2 = stu[k][1]["side"]
                dic = {
                    "shp1": shp_1,
                    "shp1_port": prt_1,
                    "shp1_side": side_1,
                    "shp2": shp_2,
                    "shp2_port": prt_2,
                    "shp2_side": side_2
                }

                no_stubs.append(dic)
                #if "COMM-DUAL*" in shp1:
                #   print(json.dumps(no_stubs, indent=4))

    with open("./output/cxn_file.json","w+") as aa:
        aa.write(json.dumps(no_stubs, indent=4))

    aa.close()

av_ports = open("available_ports.txt", "r")
av_ports = [j.replace("\n","") for j in av_ports.readlines()]

#print(j_file)
cxn_list = []
for shp in j_file:
    obj_type = shp["obj_type"]
    obj_name = shp["instance"]
    att = shp["att"]  # Get the attribute att for offset
    ll = shp["ch"] + att + 1
    ofs = att
    #print("Object:", obj_name)
    #print("Offset:", ofs)
    #print("Line Limit:", ll)
    for i in shp["lines"][ofs:ll]:
        gp = i.split(",")[0].strip()
        if gp in av_ports or obj_type == "HOT-ARROW":  # Check if port is in available ports or "NONE"
            #print("Processing port:", gp)
            shp_det = get_offset(obj_type, obj_name, gp, i)
            #print("Shape details:", shp_det)
            if shp_det:
                start_end = calc_start_end(shp["x"],
                            shp["y"],
                            shp["h"],
                            shp["w"],
                            gp,
                            shp_det[0],
                            shp_det[1],
                            shp_det[2],
                            shp_det[3],
                            obj_name
                            )

                cxn_list.append(start_end)
            #else:
            #    print("no details", obj_name, obj_type, i)
        #else:
        #    print("cant find port", obj_name, obj_type, ofs, i)


if write_outputs_to_file:
    with open("./output/start_end.json","w+") as pp:
        pp.write(json.dumps(cxn_list, indent=4))
    pp.close()
#print(cxn_list)
get_cxn_obj(cxn_list)