import uuid

print(uuid.uuid4().hex)


# import subprocess

# subprocess.call("pdftk cert_template.pdf dump_data_fields output fields.txt")

# field_list = list()
# with open("fields.txt", "r") as fields_file:
#     lines = fields_file.readlines()

# linesList = list()
# myDict = dict()
# for eachLine in lines:
#         if "---" in eachLine:
#            continue
#         if "FieldType" in eachLine:
#             myDict = dict()
#             linesList.append(myDict)
#             start = eachLine.find(":")
#             field_type = eachLine[start + 1:-1]
#             myDict["FieldType"] = field_type
#         if "FieldName" in eachLine:
#             start = eachLine.find(":")
#             field_name = eachLine[start + 1: -1]
#             myDict["FieldName"] = field_name

# print(linesList)
        
        
        