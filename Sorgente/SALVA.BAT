del *.zip
del *.exe
pkzip -a m113216 *.*
copy m113216.zip a: /v
copy avlas.bat a: /v
dir a:

