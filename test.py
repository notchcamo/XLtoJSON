import sys
import pandas as pd
from convert import convert_xl_to_json, convert_json_to_xl

# Parse command line args
print("args: ", sys.argv)
source = sys.argv[1]
dest = sys.argv[2]

# Convert
if source.endswith(".xlsx"):
    print("xlsx -> json")
    convert_xl_to_json(source, dest)
    print("Done.")
elif source.endswith(".json"):
    print("json -> xlsx")
    convert_json_to_xl(source, dest)
    print("Done.")