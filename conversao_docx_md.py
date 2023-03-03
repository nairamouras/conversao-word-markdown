from os import walk
import os
import pypandoc

files = []
path = 'temp'
for (dirpath, dirnames, filenames) in walk(path):
    files.extend(filenames)
    break
tamanho = len(files)
for i in range(tamanho):
    if files[i].endswith('.docx'):
        md = os.path.splitext(files[i])[0]
        pypandoc.convert_file(path+'/'+files[i], to='md', outputfile=path+'/'+md+'.md')