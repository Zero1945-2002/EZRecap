import PySimpleGUI as sg
from openpyxl import load_workbook
import os
import subprocess
import win32com.client

file_types = [("Excel Workbook (*.xlsx)", "*.xlsx"),
              ("Excel 97-2003 Workbook (*.xls)", "*.xls"),
              ("All files (*.*)", "*.*")]

halaman_awal = [[
            sg.Text("Excel File"),
            sg.Input(tooltip='klik tombol browse untuk memilih file',default_text='',size=(25, 1), key="-FILE-"),
            sg.FileBrowse(file_types=file_types),
            sg.Button("Buka Excel")
        ]]

window = sg.Window("Rekap Nilai", halaman_awal)

def is_file_open(filename):
    try:
        # Mencoba membuka file dalam mode penulisan (write mode)
        with open(filename, 'a') as f:
            pass
        return
    except PermissionError:
        excel_app = win32com.client.Dispatch("Excel.Application")
        wb = excel_app.Workbooks.Open(filename)
        wb.Save()
        wb.Close(SaveChanges=False)
        excel_app.Quit()
        return 

def cek(var,text):
    if var == False:
        sg.popup(text)
        return False
    return True

def update_nilai_tr():
    nama = False
    nim = False
    p_nim = False
    jenis_tr = False

    # Kelas
    if len(values['-KELAS-']) == 0:
        sg.popup('Mohon pilih kelasnya terlebih dahulu!')
        return

    # NIM
    if str(values['-NIM-']) == ' ' or str(values['-NIM-']) == '':
        sg.popup('Masukkan NIM terlebih dahulu!')
        return
    
    # Jenis TR
    if len(values['-TR-']) == 0:
        sg.popup('Mohon pilih TR terlebih dahulu!')
        return

    # Nilai TR
    if str(values['-NILAI-TR-']).isdigit():
        nilai = int(values['-NILAI-TR-'])
        if nilai < 0 or nilai > 100:
            sg.popup('Masukkan nilai dengan range 0 s/d 100')
            return
    else:
        sg.popup('Masukkan nilai dengan range 0 s/d 100')
        return

    # Pencarian cell
    kelas = str(values['-KELAS-'][0])
    ws = workbook[kelas]
    for k in ws.iter_cols():
        for c in k:
            if c.value == "NIM" or c.value == "nim" or c.value == "Nim":
                nim_col = c.column
                nim = True
                for row in ws.iter_rows(min_col=nim_col, max_col=nim_col):
                    for cell in row:
                        if str(cell.value) == str(values['-NIM-']):
                            nim_row = cell.row
                            p_nim = True
            if c.value == values['-TR-'][0]:
                tr_col = c.column
                jenis_tr = True
            if c.value == "Nama" or c.value == "NAMA" or c.value == "nama":
                nama_col = c.column
                nama = True
    if cek(nim,"Pastikan kolom 'NIM' tersedia!") == False:
        return
    if cek(p_nim,'Praktikan dengan nim '+str(values['-NIM-'])+' tidak ditemukan!') == False:
        return
    if cek(jenis_tr,"Pastikan kolom '"+str(values['-TR-'][0])+"' tersedia!") == False:
        return
    if cek(nama,"Pastikan kolom 'Nama' tersedia!") == False:
        return

    # Ubah nilai cell
    ws.cell(row=nim_row, column=tr_col, value=int(values['-NILAI-TR-']))
    workbook.save(filename=filename)
    nama_praktikan = ws.cell(row=nim_row, column=nama_col).value
    nim_praktikan = ws.cell(row=nim_row, column=nim_col).value
    teks = 'Nilai '+str(values['-TR-'][0])+' praktikan '+str(nama_praktikan)+' ('+str(nim_praktikan)+') telah di update!'
    sg.popup(teks)
    return

def update_nilai_ta():
    nama = False
    nim = False
    p_nim = False
    jenis_ta = False

    # Kelas
    if len(values['-KELAS-']) == 0:
        sg.popup('Mohon pilih kelasnya terlebih dahulu!')
        return

    # NIM
    if str(values['-NIM-']) == ' ' or str(values['-NIM-']) == '':
        sg.popup('Masukkan NIM terlebih dahulu!')
        return
    
    # Jenis TR
    if len(values['-TA-']) == 0:
        sg.popup('Mohon pilih TA terlebih dahulu!')
        return

    # Nilai TR
    if str(values['-NILAI-TA-']).isdigit():
        nilai = int(values['-NILAI-TA-'])
        if nilai < 0 or nilai > 100:
            sg.popup('Masukkan nilai dengan range 0 s/d 100')
            return
    else:
        sg.popup('Masukkan nilai dengan range 0 s/d 100')
        return

    # Pencarian cell
    kelas = str(values['-KELAS-'][0])
    ws = workbook[kelas]
    for k in ws.iter_cols():
        for c in k:
            if c.value == "NIM" or c.value == "nim" or c.value == "Nim":
                nim_col = c.column
                nim = True
                for row in ws.iter_rows(min_col=nim_col, max_col=nim_col):
                    for cell in row:
                        if str(cell.value) == str(values['-NIM-']):
                            nim_row = cell.row
                            p_nim = True
            if c.value == values['-TA-'][0]:
                ta_col = c.column
                jenis_ta = True
            if c.value == "Nama" or c.value == "NAMA" or c.value == "nama":
                nama_col = c.column
                nama = True
    if cek(nim,"Pastikan kolom 'NIM' tersedia!") == False:
        return
    if cek(p_nim,'Praktikan dengan nim '+str(values['-NIM-'])+' tidak ditemukan!') == False:
        return
    if cek(jenis_ta,"Pastikan kolom '"+str(values['-TA-'][0])+"' tersedia!") == False:
        return
    if cek(nama,"Pastikan kolom 'Nama' tersedia!") == False:
        return

    # Ubah nilai cell
    ws.cell(row=nim_row, column=ta_col, value=int(values['-NILAI-TA-']))
    workbook.save(filename=filename)
    nama_praktikan = ws.cell(row=nim_row, column=nama_col).value
    nim_praktikan = ws.cell(row=nim_row, column=nim_col).value
    teks = 'Nilai '+str(values['-TA-'][0])+' praktikan '+str(nama_praktikan)+' ('+str(nim_praktikan)+') telah di update!'
    sg.popup(teks)
    return
            
while True:
    event, values = window.read()
    print(event, values)

    if event == sg.WIN_CLOSED:
        break
    elif event == "Buka Excel":
        if values['-FILE-'] == ' ' or values['-FILE-'] == '':
            sg.popup("Klik tombol 'Browse' untuk memilih file!")
        else:
            filename = values["-FILE-"]
            if os.path.isfile(filename):
                nama_file = filename
                is_file_open(nama_file)
                workbook = load_workbook(filename=nama_file)
                sheet = workbook.sheetnames
                col1 = sg.Column([
                                [sg.Frame('Kelas:',[[sg.Text('Kelas:'),sg.Listbox(sheet, size=(20, 2), enable_events=True, key='-KELAS-'), sg.Text('    *klik di sini untuk memilih kelas', text_color="#3e3e3e")],],size=(505,70))]], size=(515,75), pad=(0,0))                                  
                halaman_lanjut = [[col1]]
                
                window.close()
                window = sg.Window('Rekap Nilai', halaman_lanjut)
            else:
                sg.popup('File tidak ditemukan!')
    elif event == 'Simpan Nilai TR':
        is_file_open(nama_file)
        update_nilai_tr()
    elif event == "Simpan Nilai TA":
        is_file_open(nama_file)
        update_nilai_ta()
    elif event == 'Kembali':
        halaman_awal = [[
            sg.Text("Excel File"),
            sg.Input(tooltip='klik tombol browse untuk memilih file',default_text='',size=(25, 1), key="-FILE-"),
            sg.FileBrowse(file_types=file_types),
            sg.Button("Buka Excel")
        ]]
        window.close()
        window = sg.Window('Rekap Nilai', halaman_awal)
    elif event == 'Buka File':
        try:
            # Membuka file Excel dengan aplikasi yang terkait
            subprocess.Popen(['start', '', nama_file], shell=True)
            sg.popup('File berhasil dibuka!')
            
        except FileNotFoundError:
            print(f"File tidak ditemukan")
        except Exception as e:
            print("Terjadi kesalahan:", str(e))
    elif event == '-KELAS-': 
        kelas = str(values['-KELAS-'][0])
        ws = workbook[kelas]
        cek_tr = 0
        cek_ta = 0
        TR = []
        TA = []
        for k in ws.iter_cols():
            for c in k:
                teks = str(c.value)
                if teks[0].upper() == "T" and teks[1].upper() == "R":
                    TR.append(c.value)
                    cek_tr = cek_tr + 1
                if teks[0].upper() == "T" and teks[1].upper() == "A":
                    TA.append(c.value)
                    cek_ta = cek_ta + 1
        is_file_open(nama_file)
        workbook = load_workbook(filename=nama_file)
        sheet = workbook.sheetnames
        col1 = sg.Column([
                        [sg.Frame('Kelas:',[[sg.Text('Kelas:'),sg.Listbox(sheet,default_values=kelas, size=(20, 2), enable_events=True, key='-KELAS-'), sg.Text('    *klik di sini untuk mengganti kelas', text_color="#3e3e3e")],],size=(505,70))],
                        [sg.Frame('NIM:',[[sg.Text('NIM:'),sg.Input(tooltip='ex : 202111129', default_text='', key='-NIM-', size=(19,1))]],size=(505,50))]], size=(515,140), pad=(0,0))                                  
        col2 = sg.Column([[sg.Frame('Tugas Rumah:', [[sg.Text(), sg.Column([
                                                                            [sg.Listbox(TR, size=(20, 5), enable_events=True, key='-TR-')],
                                                                            [sg.Text('Nilai:'),sg.Input(tooltip='masukkan nilai 1 s/d 100', default_text='', key='-NILAI-TR-', size=(19,1))],
                                                                            [sg.Button('Simpan Nilai TR')]], size=(230,155), pad=(0,0))]])] ], pad=(0,0))
        col3 = sg.Column([[sg.Frame('Test Awal:', [[sg.Text(), sg.Column([
                                                                        [sg.Listbox(TA, size=(20, 5), enable_events=True, key='-TA-')],
                                                                        [sg.Text('Nilai:'),sg.Input(tooltip='masukkan nilai 1 s/d 100', default_text='', key='-NILAI-TA-', size=(19,1))],
                                                                        [sg.Button('Simpan Nilai TA')]], size=(230,155), pad=(0,0))]])] ], pad=(0,0))
        col4 = sg.Column([[sg.Frame('Actions:',
                                            [[sg.Column([[sg.Button('Kembali'), sg.Button('Buka File')]],size=(460,45), pad=(0,0))]],size=(505,60))]], pad=(0,0))
        halaman_lanjut = [[col1],
                [col2, col3],
                [col4]]
        
        window.close()
        window = sg.Window('Rekap Nilai Kelas '+kelas, halaman_lanjut)


window.close()
