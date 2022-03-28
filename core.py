# importiranje potrebnih modulov in knjiznjic
readmode=True
from PIL import Image
import numpy as np
import pandas as pd
import xlsxwriter


# definirajmo slikovno datoteko s katero bomo manipulirali
image = input("Vpišite ime slikovne datoteke: ")

# preverimo če sploh datoteka obstaja in uporabnika ponovno prosimo za vnos, če je bil napačen. Konvertiramo tudi barve v RGB (24 bit).
while readmode:
    try:
        image = Image.open(image).convert('RGB')
        readmode = False
        

    except:
        print("Ne najdem", image)
        print("Prosim, poskusite ponovno. ")
        image = input("Vpišite ime slikovne datoteke (tudi končnico): ")


# če je slika razmeroma velika, uporabniku damo možnost da jo zmanjša na širino 300 pikslov in ohrani razmerje
while max(image.size) > 300:
    question = input("Slikovna datoteka je velika in bo zahtevna za obdelavo. Jo želite zmanjšati? da / ne: ")
    if question == ("ne"):
        break
    while question == ("da"):
        basewidth = 300
        wpercent = (basewidth/float(image.size[0]))
        hsize = int((float(image.size[1])*float(wpercent)))
        image = image.resize((basewidth,hsize), Image.ANTIALIAS)

        # uporabniku damo možnost da shrani pomanjsano verzijo slike za referenco
        question_save = input("Želite shraniti pomanjšano verzijo slike? da / ne: ")
        if question_save == ("ne"):
            break
        elif question_save == ("da"):
            small_image = input("Kako želite poimenovati pomanjšano sliko (brez končnice)? ")
            imgform = ".png"
            small_image_form = small_image+imgform
            small_image_path = "./pomanjsane_slike"
            image.save(f"{small_image_path}/{small_image_form}")
            break
        else:
            print("Prosim odgovorite z 'da' ali 'ne'")
        
    else:
        print("Prosim odgovorite z 'da' ali 'ne'")
    
    # break vstavimo zaradi slik, ki so vertikalnega formata in bo po zmanjšanju širine višina vseeno večja od 300px
    break


# shranimo dimenzije slike
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
writer = pd.ExcelWriter("./xlsx_datoteke/" + str(save_file_name) + ".xlsx", engine='xlsxwriter')

# konvertira dataframe v XlsxWriter Excel objekt.
df.to_excel(writer, sheet_name='Sheet69')

# definiramo xlsxwriter workbook in worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Sheet69']

# zapišemo dimenzije dataframe-a.
(max_row, max_col) = df.shape


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

# funkcija, ki shrani xlsx datoteko v podmapo "xlsx_datoteke"
writer.save()
