# importiranje potrebnih modulov in knjiznjic
readmode=True
from PIL import Image
import numpy as np
import pandas as pd
# import openpyxl
import xlsxwriter


# definirajmo slikovno datoteko s katero bomo manipulirali
image = input("Vpišite ime slikovne datoteke: ")

# preverimo če sploh datoteka obstaja in uporabnika ponovno prosimo za vnos, če je bil napačen
while readmode:
    try:
        image = Image.open(image).convert('RGB')
        readmode = False

    except:
        print("Ne najdem", image)
        print("Prosim, poskusite ponovno. ")
        image = input("Vpišite ime slikovne datoteke (tudi končnico): ")

# ugotovimo dimenzije slike
# https://stackoverflow.com/questions/6444548/how-do-i-get-the-picture-size-with-pil
width, height = image.size

# širino množimo s 3, ker imamo RGB - tri vrednosti za eno celico
width = width*3

# sliko spremenimo v array RGB
image_sequence = image.getdata()
image_array = np.array(image_sequence)

# spremenimo dimenzijo matrike, da bo ustrezala resoluciji slike
# vir: https://stackoverflow.com/questions/12575421/convert-a-1d-array-to-a-2d-array-in-numpy
image_array.shape = (height, width)

# ustvari dataframe iz arraya
df = pd.DataFrame(image_array)

# shranimo v xlsx datoteko po izbiri
save_file_name = input("Prosim vpišite ime datoteke, ki jo želite ustvariti (BREZ končnice): ")

# ustvari Pandas Excel writer z XlsxWriter kot orodje
# vir: https://xlsxwriter.readthedocs.io/example_pandas_conditional.html
writer = pd.ExcelWriter(str(save_file_name) + ".xlsx", engine='xlsxwriter')

# konvertira dataframe v XlsxWriter Excel objekt.
df.to_excel(writer, sheet_name='Sheet69')

# definiramo xlsxwriter workbook in worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Sheet69']

# zapišemo dimenzije dataframe-a.
(max_row, max_col) = df.shape

# df.to_excel(str(save_file_name) + ".xlsx", sheet_name='Sheet1') -> brez barvanja se lahko shrani xlsx s tem

# definirajmo array, ki ga kasneje uporabimo v zanki za RGB
rgb = ["#ff0000", "#00ff00", "#0000ff"]

# z zanko bomo vsak tretji stolpec formatirali z rdečo, zeleno in modro
for x in range(width):
    worksheet.conditional_format(1, x, height, x+1, {'type': '2_color_scale', 'min_value': '0', 'max_value': '255', 'min_color': '#000000', 'max_color': rgb[x%3]})


# za prikaz večjih slik je potrebno vse pomanjšati, da že ob zagonu npr Excela postane pregledno
for x in range(height+1):
    worksheet.set_row_pixels(x, 12) #visina v pt
worksheet.set_column_pixels(0, width+1, 4) #sirina v character units
worksheet.hide_row_col_headers()
worksheet.hide_gridlines(2)
worksheet.set_zoom(32)

# funkcija, ki shrani xlsx datoteko v trenutno mapo
writer.save()

# print(df)