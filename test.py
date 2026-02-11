import sys
import pandas as pd
from convert import convertXLtoJSON, convertJSONtoXL

# Parse command line args
print("args: ", sys.argv)
source = sys.argv[1]
dest = sys.argv[2]

# Convert
if source.endswith(".xlsx"):
    print("xlsx -> json")
    convertXLtoJSON(source, dest)
    print("Done.")
elif source.endswith(".json"):
    print("json -> xlsx")
    convertJSONtoXL(source, dest)
    print("Done.")