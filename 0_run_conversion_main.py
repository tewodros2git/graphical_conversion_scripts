import sys
import os
import subprocess



if not os.path.isdir("./output"):
    print("creating output folder")
    os.mkdir("output")

file_name = sys.argv[1]
#file_path = sys.argv[2]
print(f"using file: {file_name} ........starting")
print("MAKE SURE YOU TO ALWAYS RUN parse_class_name.py FOR NEW SERIES")
# 1
subprocess.run(['python', "1_flat_to_dict.py", file_name], shell=False,
               stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)

print("step 1 done")
# 2
nn = subprocess.run(['node', "2_get_unique_names.js", file_name], shell=False,
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE)
print(nn.stderr, nn.stdout)

print("step 2 done")
# 3
subprocess.run(['python', "3_add_unique_name.py", file_name], shell=False,
               stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)

print("step 3 done")
# 4
nn = subprocess.run(['python', "4_get_cxn_loc_edited.py", file_name], shell=False,
               stdout=subprocess.PIPE, stderr=subprocess.PIPE)
out = nn.stdout.decode('utf-8').split("\n")
for o in out:
    print(o)
print()
print(nn.stderr)
print("step 4 done")
# 5
nn = subprocess.run(['node', "5_convert_rcap_to_svg.js", file_name], shell=False,
               stdout=subprocess.PIPE, stderr=subprocess.PIPE)
#out = nn.stdout.decode('utf-8').split("\n")
#for o in out:
#    print(o)
#print()
#print(nn.stderr)
print("step 5 done")

# 6
six = subprocess.run(['python', "6_format_names.py", file_name], shell=False,
                     stdout=subprocess.PIPE, stderr=subprocess.PIPE)
print(six.stderr)
print("step 6 done")

'''
def process_directory(directory):
    for filename in os.listdir(directory):
        if filename.endswith('.json'):  # Adjust the condition as per your file type
            file_path = os.path.join(directory, filename)
            process_file(file_path)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python process_looper.py <directory>")
        sys.exit(1)

    directory_path = sys.argv[1]
    if not os.path.isdir(directory_path):
        print("Invalid directory path.")
        sys.exit(1)

    process_directory(directory_path)
'''