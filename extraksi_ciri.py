import os, cv2, xlsxwriter
import pandas as pd, imageio.v2 as imageio, numpy as np


# define the function
def color_extraction(folder_path, writer):
    #List Kosong
    rows = []

    # Cari ke mana data yang ingin di ektraksi
    filenames = os.listdir(folder_path)

    # Pengulangan untuk pengambilan data, pemisahan menjadi 3 channels, penghitungan rata - rata, Pemasukan data ke dalam list Kosong 
    for file in filenames:
        image = imageio.imread(os.path.join(folder_path, file))

        # Pemisahan Channel Gambar
        red_channel, blue_channel, green_channel = cv2.split(image)

        # Penghitungan Rata - Rata
        # Pastikan urutan penghitungan rata - rata sama dengan urutan pemisahan channel
        red_average = np.average(red_channel)
        blue_average = np.average(blue_channel)
        green_average = np.average(green_channel)

        # Filename tanpa Path
        # file_name = os.path.basename(file)

        # Buat dan tambahkan data tiap row
        row = {'Image Label': file, 'Red channel': red_average, 'Green channel': green_average, 'Blue channel': blue_average}
        rows.append(row)

        # Pemasukan data ke Excel
        df = pd.DataFrame(rows)
        df.to_excel(writer, index=False)

        # Edit Excel
        workbook = writer.book

        # format variabel
        formats = workbook.add_format({'align': 'center', 'text_wrap': True})
        # jalnankan format ke excel
        for worksheet in workbook.worksheets():
            worksheet.set_column('A:D', 11.78, formats)
            # formats.set_align('left')
            # worksheet.set_column('A:A', 11.78, formats)

# Variabel berisi PATH menuju lokasi dataset
folder_path = 'D:\Kuliah\Semester 3\Pengolahan Citra Digital\Dataset Daging\Dataset Daging'

# Overwrite Data Spreadsheet Tiap Run untuk menghindari error Pemrission Denied
with pd.ExcelWriter('Color Extractions.xlsx',  engine='xlsxwriter') as writer:
    # Pemanggilan fungsi def
    color_extraction(folder_path, writer)