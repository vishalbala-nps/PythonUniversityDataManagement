from pyexcel_ods import get_data
import json
from prettytable import PrettyTable

acadamics = {}
ielts = {}
interview = {}
fresult = PrettyTable(['Student Name','% in Academics','% in IELTS','% in Interview','Total %'])


def getavg(t1,t2,t3,t4):
   global acadamics
   t1 = [i for i in t1 if i != []]
   t2 = [i for i in t2 if i != []]
   t3 = [i for i in t3 if i != []]
   t4 = [i for i in t4 if i != []]

   for i in range(3,len(t1)):
      sum = t1[i][1]*2+t1[i][2]*2+t1[i][3]+t1[i][4]+t1[i][5]*2+t1[i][6]+t1[i][7]
      acadamics[t1[i][0]] = sum
      
   for i in range(3,len(t2)):
      sum = t2[i][1]*2+t2[i][2]*2+t2[i][3]+t2[i][4]+t2[i][5]*2+t2[i][6]+t2[i][7]
      acadamics[t2[i][0]] += sum

   for i in range(3,len(t3)):
      sum = t3[i][1]*2+t3[i][2]*2+t3[i][3]+t3[i][4]+t3[i][5]*2+t3[i][6]+t3[i][7]
      acadamics[t3[i][0]] += sum

   for i in range(3,len(t4)):
      sum = t4[i][1]*2+t4[i][2]*2+t4[i][3]+t4[i][4]+t4[i][5]*2+t4[i][6]+t4[i][7]
      acadamics[t4[i][0]] += sum

   for name,mark in acadamics.items():
      acadamics[name] = round((mark/4000)*100,2)

def getielts(data):
   global ielts
   data = [i for i in data if i != []]
   for i in range(3,len(data)):
      ielts[data[i][0]] = round(((data[i][1]+data[i][2]+data[i][3]+data[i][4])/36)*100,2)

def getinterview(data):
   global interview
   data = [i for i in data if i != []]
   for i in range(3,len(data)):
      interview[data[i][0]] = round(((data[i][1]+data[i][2]+data[i][3]+data[i][4]+data[i][5])/50)*100,2)

#t1 = input("Please Enter the location of the Term 1 scores ODS file:")
#t2 = input("Please Enter the location of the Term 2 scores ODS file:")
#t3 = input("Please Enter the location of the Term 3 scores ODS file:")
#t4 = input("Please Enter the location of the Term 4 scores ODS file:")
#ielts = input("Please Enter the location of the IELTS scores ODS file:")
#interview = input("Please Enter the location of the Interview scores ODS file:")

t1 = "Data1.ods"
t2 = "Data2.ods"
t3 = "Data3.ods"
t4 = "Data4.ods"
ieltsf = "IELTS.ods"
interviewf = "Interview.ods"

try:
   t1d = json.dumps(get_data(t1))
   t1d = json.loads(t1d)['Sheet1']
   t2d = json.dumps(get_data(t2))
   t2d = json.loads(t2d)['Sheet1']
   t3d = json.dumps(get_data(t3))
   t3d = json.loads(t3d)['Sheet1']
   t4d = json.dumps(get_data(t4))
   t4d = json.loads(t4d)['Sheet1']
   ieltsd = json.dumps(get_data(ieltsf))
   ieltsd = json.loads(ieltsd)['Sheet1']
   interviewd = json.dumps(get_data(interviewf))
   interviewd = json.loads(interviewd)['Sheet1']
except:
   print("Failed to load ODS File!")
   exit(1)

getavg(t1d,t2d,t3d,t4d)
getielts(ieltsd)
getinterview(interviewd)

for name,mark in acadamics.items():
   fresult.add_row([name, mark, ielts[name],interview[name],round(mark+interview[name]+ielts[name],2)])
   
print(fresult.get_string(sortby=("Total %"), reversesort=True))