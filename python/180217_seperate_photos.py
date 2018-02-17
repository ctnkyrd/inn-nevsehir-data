import os
from shutil import copyfile

photo_directory = "E:\\excele\\fotograflar"
new_directory= "E:\\excele\\all_photo"

if not os.path.exists(new_directory):
    os.makedirs(new_directory)


for r,d,f in os.walk(photo_directory):
    for photo in f:
        if photo.endswith('jpg'):
            try:
                src = os.path.join(r,photo.replace('Ayvali','Ayvalı').decode('utf-8'))
                if(photo.split('_')[1].decode('utf-8')!=u"Taşkınpaşa"):
                    dest = new_directory +"\\"+"-1_"+photo.split('_')[3]+".jpg"
                else:
                    dest = new_directory +"\\"+photo.split('_')[2]+photo.split('_')[3]+".jpg"
                copyfile(src, dest)
                print photo 
            except BaseException as be:
                print be.message

print "Completed!"