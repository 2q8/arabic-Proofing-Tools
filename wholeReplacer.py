import re
import xlrd
import os

def replace(filename):
    #Open txt file

    output="fixed_"+filename #prefix the file name for output

    f=open(filename,"r",encoding="utf-8") #open source file
    txt=f.read()

    out=open(output,'w',encoding="utf-8")   #create output file

    #Replace each word
    for i in range(len(wordlist)):
        txt=re.sub(r"\b%s\b"%wordlist[i][0],wordlist[i][1],txt)
        #txt=txt.replace(wordlist[i][0],wordlist[i][1])

    print("Done replacing.")

    out.write(txt)
    out.close()

#Import wordlist
wordlist=[]
wb = xlrd.open_workbook('Replacer v2 .xlsm')
sh = wb.sheet_by_index(0)

for i in range(sh.nrows):
    #print sh.cell_value(i,0)
    wordlist.append([sh.cell_value(i,0),sh.cell_value(i,1)])

wordlist=wordlist[1:] #remove headers
print("Words loaded: "+str(len(wordlist)))



#main
files=os.listdir(".") #find all files in current folder

for file in files:
    if file.endswith(".txt"): #if it is a txt file
        print("Replacing "+file+"...")
        replace(file)
#replace("3.txt")
