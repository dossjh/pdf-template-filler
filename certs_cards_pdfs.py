#!python
from fdfgen import forge_fdf
import subprocess
import time
import csv
import sys
import os


# fields = [('Name', 'John Smith'), ('Phone', '555-1234')]
# fdf = forge_fdf("", fields, [], [], [])
# fdf_file = open("data.fdf", "wb")
# fdf_file.write(fdf)
# fdf_file.close()

error_no_file = False
if len(sys.argv) > 1:
    filename = sys.argv[1]
else:
    filename = "names.csv"
    print("Attempting to use default names file 'names.csv'")
    
if not os.path.isfile(filename):
    print("The file: {} does not exist, please run with a csv file containing names.")
else:    
    names_list = list()
    with open(filename, "r") as f:
        reader = csv.reader(f)
        for row in reader:
            for each in row:
                names_list.append(each)
                print(each)

    namesArray = list()
    for groupNum in range(0, len(names_list), 4):
        newList = list()
        for eachNum in range(groupNum, groupNum + 4):
            if eachNum > (len(names_list) - 1):
                break
            #generate certificate
            fields = [('name', names_list[eachNum])]
            fdf = forge_fdf("", fields, [], [], [])
            out_filename = "certificates\\" + names_list[eachNum].replace(" ", "_") + ".pdf"
            fdf_file_name = out_filename + ".fdf"
            fdf_file = open(fdf_file_name, "wb")
            fdf_file.write(fdf)
            fdf_file.close()
            subprocess.Popen("pdftk Certificate.pdf fill_form " + fdf_file_name + " output " +
                     out_filename)
            newList.append(names_list[eachNum])
        namesArray.append(newList)
    number = 1
    for each in namesArray:
        for eachvalue in each:
            print(number, eachvalue)
            number = number + 1
        lenEach = len(each)
        if lenEach < 4:
            for i in range(4 - lenEach):
                each.append("not used")

        fields = [('name1', each[0]), ('name2', each[1]), ('name3', each[2]), ('name4', each[3])]
        print(fields)
        fdf = forge_fdf("", fields, [], [], [])
        filename = "certcards\\" + each[0].replace(" ", "_") + ".pdf"
        fdf_file_name = filename + ".fdf"
        fdf_file = open(fdf_file_name, "wb")
        fdf_file.write(fdf)
        fdf_file.close()
        subprocess.Popen("pdftk cert_template.pdf fill_form " + fdf_file_name + " output " + filename)

    print("waiting 5 second to delete temp files...")
    time.sleep(5)
    workingDirectory = os.path.dirname(os.path.realpath(__file__))
    workingDirectory = workingDirectory + "\\certcards"

    subprocess.Popen("del *.fdf", cwd=workingDirectory, shell=True)

    certsWorking = os.path.dirname(os.path.realpath(__file__)) + "\\certificates"
    subprocess.Popen("del *.fdf", cwd=certsWorking, shell=True)
    