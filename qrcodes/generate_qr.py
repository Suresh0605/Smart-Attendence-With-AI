import qrcode
import os

path = "C:/Users/lenovo/PycharmProjects/faceidbasic/faces"
mylist = os.listdir(path)
for cl in mylist:
    name = os.path.splitext(cl)[0]
    qr = qrcode.make(name)
    qr.save(name+".png")