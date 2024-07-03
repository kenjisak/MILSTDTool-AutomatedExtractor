import re

if re.match(re.compile(r'^\d+(\.\d+)+\.?$'), "5.12.3.4.16.1"):
    print(True)
else:
    print(False)