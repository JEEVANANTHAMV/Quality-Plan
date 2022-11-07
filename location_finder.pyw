import os
name=""
f1=open("name.txt","r")
f2=open("location.txt","w")
name=f1.read()
path = 'E:\\'
for r , d,f in os.walk(path):
    for file in f:
        if name in file:
            f2.write(os.path.join(r,file))
f1.close()
f2.close()
os.remove("name.txt")
