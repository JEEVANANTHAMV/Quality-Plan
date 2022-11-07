import os
name=""

f1=open("name_1.txt","r")
f2=open("inspection_qp.txt","w")
f3=open("search_location.txt","r")
name=f1.read()
#need to change based on given location
path = f3.read()
for r , d,f in os.walk(path):
    for file in f:
        if name in file:
            path=os.path.join(r,file)
            f2.write(path)
            f2.write("\n")
f1.close()

f2.close()
os.remove("name_1.txt")
