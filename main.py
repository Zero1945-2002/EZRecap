import PySimpleGUI as sg
from openpyxl import load_workbook
import os
import subprocess

file_types = [
    ("Excel Workbook (*.xlsx)", "*.xlsx"),  # Format file Excel xlsx
    ("Excel 97-2003 Workbook (*.xls)", "*.xls"),  # Format file Excel xls (versi lama)
    ("All files (*.*)", "*.*"),  # Semua jenis file
    ("Excel Macro-Enabled Workbook (*.xlsm)", "*.xlsm"),  # Format file Excel xlsm (mengandung macro)
    ("Excel Template (*.xltx)", "*.xltx"),  # Format file Excel xltx (template)
    ("Excel Macro-Enabled Template (*.xltm)", "*.xltm"),  # Format file Excel xltm (template mengandung macro)
    ("CSV (Comma Delimited) (*.csv)", "*.csv"),  # Format file Excel csv (Comma Separated Values)
    ("Text File (*.txt)", "*.txt"),  # Format file Excel txt (Text File)
    ("PRN (Printable) (*.prn)", "*.prn")  # Format file Excel prn (Printable)
]


sheet = []
TR = []
TA = []

def popup():
    layout = [[
                sg.Text("Excel File"),
                sg.Input(tooltip='klik tombol browse untuk memilih file',size=(25, 1), key="-FILE-"),
                sg.FileBrowse(file_types=file_types),
                sg.Button("Confirm")
            ]]
    event, values = sg.Window('Pilih File', layout, modal=True).read(close=True)
    return None if event == sg.WIN_CLOSED else str(values['-FILE-'])

menu_def = [['&File', ['&Pilih File','&Keluar']],
            ['&Help', ['&Cara Penggunaan','&About']] ]

col1 = sg.Column([
                        [sg.Frame('Kelas:',[[sg.Text('Kelas:'),sg.Combo(sheet, size=(20, 10), enable_events=True, key='-KELAS-', readonly=True), sg.Text('    *klik di sini untuk mengganti kelas', text_color="#3e3e3e")],],size=(505,60))],
                        [sg.Frame('NIM:',[[sg.Text('NIM:'),sg.Input(tooltip='ex : 202111129', default_text='', key='-NIM-', size=(19,1))]],size=(505,50))]], size=(515,130), pad=(0,0))                                  
col2 = sg.Column([[sg.Frame('Tugas Rumah:', [[sg.Text(), sg.Column([
                                                                    [sg.Combo(TR, tooltip='klik disini untuk memilih jenis TR', size=(20, 10), enable_events=True, key='-TR-', readonly=True)],
                                                                    [sg.Text('Nilai:'),sg.Input(tooltip='masukkan nilai 1 s/d 100', default_text='', key='-NILAI-TR-', size=(19,1))],
                                                                    [sg.Button('Simpan Nilai TR')]], size=(230,105), pad=(0,0))]])] ], pad=(0,0))
col3 = sg.Column([[sg.Frame('Test Awal:', [[sg.Text(), sg.Column([
                                                                [sg.Combo(TA, tooltip='klik disini untuk memilih jenis TA', size=(20, 10), enable_events=True, key='-TA-', readonly=True)],
                                                                [sg.Text('Nilai:'),sg.Input(tooltip='masukkan nilai 1 s/d 100', default_text='', key='-NILAI-TA-', size=(19,1))],
                                                                [sg.Button('Simpan Nilai TA')]], size=(230,105), pad=(0,0))]])] ], pad=(0,0))
col4 = sg.Column([[sg.Frame('Actions:',[[sg.Column([[sg.Button('Buka File')]],size=(460,45), pad=(0,0))]],size=(505,60))]], pad=(0,0))

layout = [ [sg.MenubarCustom(menu_def, key='-MENU-', tearoff=False)]]  

layout += [[col1],
        [col2, col3],
        [col4]]

window = sg.Window('EZRecap', layout, margins=(0,0))


def cek(var,text):
    if var == False:
        sg.popup(text)
        return False
    return True

def update_nilai(type):
    if type == 'TR':
        t_type = "-TR-"
        n_type = "-NILAI-TR-"
    elif type == 'TA':
        t_type = "-TA-"
        n_type = "-NILAI-TA-"
    nama = False
    nim = False
    p_nim = False
    jenis = False

    # Kelas
    if len(values['-KELAS-']) == 0:
        sg.popup('Mohon pilih kelasnya terlebih dahulu!')
        return

    # NIM
    if str(values['-NIM-']) == ' ' or str(values['-NIM-']) == '':
        sg.popup('Masukkan NIM terlebih dahulu!')
        return
    
    # Jenis T
    if len(values[t_type]) == 0:
        sg.popup('Mohon pilih jenis tugas '+t_type+' terlebih dahulu!')
        return

    # Nilai TR
    if str(values[n_type]).isdigit():
        nilai = int(values[n_type])
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
            if c.value == values[t_type]:
                t_col = c.column
                jenis = True
            if c.value == "Nama" or c.value == "NAMA" or c.value == "nama":
                nama_col = c.column
                nama = True
    if cek(nim,"Pastikan kolom 'NIM' tersedia!") == False:
        return
    if cek(p_nim,'Praktikan dengan nim '+str(values['-NIM-'])+' tidak ditemukan!') == False:
        return
    if cek(jenis,"Pastikan kolom '"+str(values[t_type])+"' tersedia!") == False:
        return
    if cek(nama,"Pastikan kolom 'Nama' tersedia!") == False:
        return

    # Ubah nilai cell
    ws.cell(row=nim_row, column=t_col, value=int(values[n_type]))
    try:
        workbook.save(filename=filename)
    except PermissionError:
        sg.popup('Harap tutup terlebih dahulu file yang terbuka!')
        return
    nama_praktikan = ws.cell(row=nim_row, column=nama_col).value
    nim_praktikan = ws.cell(row=nim_row, column=nim_col).value
    teks = 'Nilai '+str(values[t_type])+' praktikan '+str(nama_praktikan)+' ('+str(nim_praktikan)+') berhasil dimasukkan!'
    sg.popup(teks)
    return
   
while True:
    event, values = window.read()
    print(event, values)
    if event == sg.WIN_CLOSED or event == "Keluar":
        break

    elif event == "Pilih File":
        filename = popup()
        if filename != None:
            window.write_event_value('proses file', None)

    elif event == "proses file":
        if filename == ' ' or filename == '':
            sg.popup("Klik tombol 'Browse' untuk memilih file!")
            window.write_event_value("Pilih File", None)
        else:
            if os.path.isfile(filename):
                sg.popup('File berhasil diakses!') 
                nama_file = filename
                workbook = load_workbook(filename=nama_file)
                sheet = workbook.sheetnames
                window['-KELAS-'].update(values=sheet)
            else:
                sg.popup('File tidak ditemukan!')
                window.write_event_value("Pilih File", None)

    elif event == 'Simpan Nilai TR':
        update_nilai("TR")

    elif event == 'Simpan Nilai TA':
        update_nilai("TA")

    elif event == 'Buka File':
        try:
            # Membuka file Excel dengan aplikasi yang terkait
            result = subprocess.run(['start', '', nama_file], shell=True, capture_output=True, text=True)
            if result.stderr:
                error_message = result.stderr.strip()
                if "The process cannot access the file because it is being used by another process." in error_message:
                    sg.popup('File sedang digunakan oleh proses lain dan telah terbuka!')
                else:
                    sg.popup('Error:', error_message)
            else:
                sg.popup('File berhasil dibuka!')
        except PermissionError as e:
                sg.popup('PermissionError:', str(e))
        except FileNotFoundError:
            sg.popup(f"File tidak ditemukan")
        except Exception as e:
            sg.popup("Mohon pilih filenya terlebih dahulu!")

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
        window['-TR-'].update(values=TR)
        window['-TA-'].update(values=TA)

    elif event == "Cara Penggunaan":
        sg.popup('Cara Penggunaan',
                 '1- Pilih file excel terlebih dahulu dengan memilih menu ',
                 '   "File" lalu tekan "Pilih File"',
                 '2- Klik "Browse" dan pilih file excel yang akan diakses,',
                 '   lalu klik "Confirm"',
                 '3- Jika file berhasil diakses, pilih kelas lalu isi NIM',
                 '   kemudian pilih jenis tugas dan isi nilainya',
                 '4- Terakhir klik "Update Nilai TA" atau "Update Nilai TR,"',
                 '   sesuai jenis tugas yang diisi',
                 '5- Dan nilai berhasil dimasukkan jika tidak terjadi kesalahan',non_blocking=True)
        
    elif event == "About":
        sg.popup('EZRecap',
                 'Alat bantu asisten lab dalam menginput nilai praktikan',
                 'dibuat oleh Zaid Immaduddin Abdurrahman',
                 '',
                 'contact us :',
                 'Email : zaidimmaduddin56@gmail.com',
                 'GitHub : https://github.com/zaid-2002/EZRecap.git')


window.close()
