from fpdf.enums import XPos, YPos
from function import *
import pandas as pd
from datetime import datetime


def createpdf(WaliKelas, NIK_WaliKelas, Kelas, Semester, Thn, Description, Ekskul, absen_file, output_dir):

    absen = pd.read_excel(f'{absen_file}', header = 1, sheet_name = 'Profile', index_col = 'No', na_filter = False)
    for i in absen.index:
        profile = absen.loc[i]
        Nama = profile.loc['Nama Siswa']
        NoInduk = profile.loc['Nomor Induk']
        Nisn = profile.loc['NISN']
        Jurusan = profile.loc['Jurusan']
        Ekskul1 = profile.loc['Ekskul 1']
        Ekskul2 = profile.loc['Ekskul 2']
        sakit = profile.loc['Sakit']
        izin = profile.loc['Izin']
        tanpa_keterangan = profile.loc['Tanpa Keterangan']
        notes = profile.loc['Catatan Wali Kelas']

        pdf = students(nama = Nama, kelas = Kelas, jurusan = Jurusan, ekskul_dir = Ekskul, semester = Semester, desc_path = Description)
        lmargin=42.55
        tmargin=90
        rmargin=42.55
        pdf.set_margins(lmargin, tmargin, rmargin)
        pdf.set_auto_page_break(True, margin = 56.7)
        #additional fonts

        class Title:
            pdf.set_y(tmargin)
            pdf.set_font("helvetica", "B", 13)
            pdf.cell(w = 0,h = 17, txt = "PEMERINTAH PROVINSI KALIMANTAN TIMUR", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = "C")
            pdf.cell(w = 0,h = 17, txt = "DINAS PENDIDIKAN DAN KEBUDAYAAN", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = "C")

            pdf.set_font_size(15)
            pdf.cell(w = 0,h = 17, txt = "SMA KRISTEN HARAPAN BANGSA", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = "C")

            pdf.set_font("arialN", "", 12)
            pdf.cell(w = 0,h = 16, txt = "Jl. Indrakila No. 99 G Kelurahan Gunung Samarinda Baru Kec. Balikpapan Utara", border = 0,  new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = "C")
            pdf.ln()

        class TopLine:
            LineYTop = pdf.get_y() - 16/3
            pdf.set_line_width(2.25)
            pdf.line(0, LineYTop, 597.6, LineYTop)
            pdf.set_line_width(1)

        class Profile:
            pdf.set_font("arialN", "", 12)
            borderbool = 0 #1 for show, 0 for not show
            #first row
            pdf.set_x(6)
            pdf.cell(w = 130.5, h = 16, txt = "Nama Peserta Didik", border = borderbool,  new_x = XPos.RIGHT, new_y = YPos.TOP, align = 'L')
            pdf.cell(w = 265.5, h = 16, txt = ": " + Nama, border = borderbool,  new_x = XPos.RIGHT, new_y = YPos.TOP, align = 'L')
            pdf.cell(w = 90, h = 16, txt = "Kelas", border = borderbool,  new_x = XPos.RIGHT, new_y = YPos.TOP, align = 'L')
            pdf.cell(w = 108, h = 16, txt = ": " + Kelas, border = borderbool,  new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            #second row
            pdf.set_x(6)
            pdf.cell(w = 130.5, h = 16, txt = "No. Induk", border = borderbool, new_x = XPos.RIGHT, new_y = YPos.TOP, align = 'L')
            pdf.cell(w = 265.5, h = 16, txt = ": " + str(NoInduk), border = borderbool,  new_x = XPos.RIGHT, new_y = YPos.TOP, align = 'L')
            pdf.cell(w = 90, h = 16, txt = "Semester", border = borderbool,  new_x = XPos.RIGHT, new_y = YPos.TOP, align = 'L')
            pdf.cell(w = 108, h = 16, txt = ": " + Semester, border = borderbool, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            #third row
            pdf.set_x(6)
            pdf.cell(w = 130.5, h = 16, txt = "NISN", border = borderbool, new_x = XPos.RIGHT, new_y = YPos.TOP, align = 'L')
            pdf.cell(w = 265.5, h = 16, txt = ": " + str(Nisn), border = borderbool, new_x = XPos.RIGHT, new_y = YPos.TOP, align = 'L')
            pdf.cell(w = 90, h = 16, txt = "Tahun Pelajaran", border = borderbool, new_x = XPos.RIGHT, new_y = YPos.TOP, align = 'L')
            pdf.cell(w = 108, h = 16, txt = ": " + Thn, border = borderbool, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            
        class BotLine:
            LineYBot = pdf.get_y() + 16/3
            pdf.set_line_width(2.25)
            pdf.line(0, LineYBot, 597.6, LineYBot)
            pdf.set_line_width(1)
            
        pdf.ln()
        pdf.set_font('arialN', 'B', 14)
        pdf.cell(w = 0, h = 17, txt = "CAPAIAN HASIL BELAJAR", border = 0, new_x = XPos.RIGHT, new_y = YPos.NEXT, align = 'C')

        class Sikap:
            def tabelsikap(self, desc, np_nk):
                score_df = pd.read_csv(f'reportcard/{Jurusan.upper()}.csv', index_col = 'Nama')
                NPNK_df = score_df[(score_df['NP/NK'] == f'{np_nk.upper()}')]['Agama']
                score = NPNK_df.loc[self.nama]
                if score >= 92:
                    predicate = "Sangat Baik"
                    desc = "Siswa selalu " + desc
                elif score >= 83:
                    predicate = "Baik"
                    desc = "Siswa sering " + desc
                elif score >= 75:
                    predicate = 'Cukup Baik'
                    desc = "Siswa " + desc
                else:
                    predicate = "Kurang"
                    desc = "Siswa jarang " + desc
                    
                self.set_x(lmargin + 5.6)   
                self.set_fill_color(217,217,217)
                self.set_line_width(1.5)
                self.set_font('arialN', 'B', 12)
                self.cell(w = 95.4, h = 21.6, txt = "Predikat", border = 1, align = 'C', fill = True)
                self.cell(w = 431.1, h = 21.6, txt = "Deskripsi", border = 1, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'C', fill = True)
                self.set_x(lmargin + 5.6)
                self.set_font_size(10)
                self.cell(w = 95.4, h = 72, txt = predicate.title(), border = 1, align = 'C', fill = False)
                padding(self, 431.1, 16, desc.capitalize(), 72, "J")
                self.ln(72)
                self.ln(12)
            
            pdf.set_x(lmargin - 31.5)
            pdf.set_font_size(12)
            pdf.cell(w = 0, h = 16, txt = "A.   SIKAP", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            pdf.set_x(lmargin - 13.5)
            pdf.cell(w = 0, h = 16, txt = "1.   Sikap Spiritual", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            descSpiritual = "menunjukkan sikap ketaatan beribadah, berprilaku syukur, berdoa sebelum dan sesudah melakukan kegiatan, merefleksikan diri dalam beribadah, serta toleransi dalam beragama."
            tabelsikap(pdf, descSpiritual, 'NP')
            
            pdf.set_x(lmargin - 13.5)
            pdf.set_font_size(12)
            descSosial = "menunjukkan sikap tanggung jawab, kerjasama, peduli, serta selalu aktif memerhatikan pelajaran dan penjelasan guru dalam proses pembelajaran."
            pdf.cell(w = 0, h = 16, txt = "2.   Sosial", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            tabelsikap(pdf, descSosial, 'NK')

        class Pengetahuan:
            pdf.add_page()
            pdf.set_x(lmargin - 31.5)
            pdf.set_font_size(12)
            pdf.cell(w = 0, h = 16, txt = "B.   PENGETAHUAN", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            pdf.set_x(lmargin - 13.5)
            pdf.cell(w = 0, h = 16, txt = "Kriteria Ketuntasan Minimal: 75", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            pdf.ln(6)
            pdf.nilai_subject('np')
            pdf.ln()
            pdf.ln(12)

        class Keterampilan:
            pdf.add_page()
            pdf.set_x(lmargin - 31.5)
            pdf.set_font_size(12)
            pdf.cell(w = 0, h = 16, txt = "C.   KETERAMPILAN", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            pdf.set_x(lmargin - 13.5)
            pdf.cell(w = 0, h = 16, txt = "Kriteria Ketuntasan Minimal: 75", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            pdf.ln(6)
            pdf.nilai_subject('nk')
            pdf.ln()
            pdf.ln()
            pdf.total_table()
            
            
        class tabel_predikat():
            pdf.add_page()
            pdf.set_x(lmargin - 13.5)
            pdf.cell(w = 0, h = 16, txt = "Kriteria Ketuntasan Minimal: 75", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            pdf.ln(6)
            pdf.set_x(lmargin - 13.5)
            pdf.set_font('arialN', 'B', 12)
            pdf.cell(w = 104.6, h = 21.6*2, txt = 'KKM', border = 1, fill = True, align = 'C')
            pdf.cell(w = 439.2, h = 21.6, txt = 'Predikat', border = 1, fill = True, align = 'C', new_x = XPos.LEFT, new_y = YPos.NEXT)
            pdf.cell(w = 104.2, h = 21.6, txt = 'Kurang (D)', border = 1, fill = True, align = 'C')
            pdf.cell(w = 104.2, h = 21.6, txt = 'Cukup (C)', border = 1, fill = True, align = 'C')
            pdf.cell(w = 104.2, h = 21.6, txt = 'Baik (B)', border = 1, fill = True, align = 'C')
            pdf.cell(w = 126.6, h = 21.6, txt = 'Sangat Baik (A)', border = 1, fill = True, align = 'C', new_x = XPos.LMARGIN, new_y = YPos.NEXT)
            pdf.set_x(lmargin - 13.5)
            pdf.cell(w = 104.6, h = 21.6, txt = '75', border = 1, align = 'C')
            pdf.cell(w = 104.2, h = 21.6, txt = '<75', border = 1, align = 'C')
            pdf.cell(w = 104.2, h = 21.6, txt = '75 - 82', border = 1, align = 'C')
            pdf.cell(w = 104.2, h = 21.6, txt = '83 - 91', border = 1, align = 'C')
            pdf.cell(w = 126.6, h = 21.6, txt = '92 - 100', border = 1, align = 'C',  new_x = XPos.LMARGIN, new_y = YPos.NEXT)
            pdf.ln()
            
        class ekstrakurikuler():
            pdf.set_x(lmargin - 31.5)
            pdf.set_font_size(12)
            pdf.cell(w = 0, h = 16, txt = "D.   EKSTRAKURIKULER", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            pdf.ln(6)
            pdf.set_x(lmargin - 13.5)
            pdf.cell(w = 181.5, h = 21.6, txt = 'Kegiatan Ekstrakurikuler', border = 1, align = 'C', fill = True)
            pdf.cell(w = 117.1, h = 21.6, txt = 'Nilai', border = 1, align = 'C', fill = True)
            pdf.cell(w = 246.9, h = 21.6, txt = 'Keterangan', border = 1, align = 'C', fill = True)
            pdf.ln()
            pdf.set_font('arialN', '', 12)
            pdf.nilai_ekskul('1.', Ekskul1)
            pdf.nilai_ekskul('2.', Ekskul2)
            pdf.ln()
            
        class prestasi():
            pdf.set_x(lmargin - 31.5)
            pdf.set_font('arialN', 'B', 12)
            pdf.set_font_size(12)
            pdf.cell(w = 0, h = 16, txt = "E.   PRESTASI", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            pdf.ln(6)
            pdf.set_x (lmargin - 13.5)
            pdf.cell(w = 33, h = 21.6, txt = 'No.', border = 1, align = 'C', fill = True)
            pdf.cell(w = 170.7, h = 21.6, txt = 'Jenis Prestasi', border = 1, align = 'C', fill = True)
            pdf.cell(w = 340.7, h = 21.6, txt = 'Keterangan', border = 1, align = 'C', fill = True)
            pdf.ln()
            pdf.set_x(lmargin - 13.5)
            pdf.set_font('arialN', '', 12)
            yprestasi = pdf.get_y()
            pdf.cell(w = 33, h = 21.6, txt = '1.', border = 1, align = 'C')
            pdf.cell(w = 170.7, h = 21.6, txt = '', border = 1, align = 'L')
            pdf.cell(w = 340.7, h = 21.6, txt = '', border = 1, align = 'L')
            pdf.ln()
            
            prestasi_df = pd.read_excel(f'{absen_file}', header = 1, sheet_name = 'Prestasi', index_col = 'No', na_filter = False)
            no_prestasi = 0
            pdf.set_y(yprestasi)
            for i in prestasi_df.index:
                peraih_prestasi = prestasi_df.loc[i]['Nama']
                if peraih_prestasi == Nama:
                    no_prestasi += 1
                    jenis_prestasi = prestasi_df.loc[i]['Jenis Prestasi']
                    keterangan = prestasi_df.loc[i]['Keterangan']
                    pdf.set_x(lmargin - 13.5)
                    pdf.cell(w = 33, h = 21.6, txt = f'{str(no_prestasi)}.', border = 1, align = 'C')
                    pdf.cell(w = 170.7, h = 21.6, txt = jenis_prestasi, border = 1, align = 'L')
                    pdf.cell(w = 340.7, h = 21.6, txt = keterangan, border = 1, align = 'L')
                    pdf.ln()
                    
        class ketidakhadiran():
            pdf.set_font('arialN', 'B', 12)
            pdf.ln()
            pdf.set_x(lmargin - 31.5)
            pdf.set_font_size(12)
            pdf.cell(w = 0, h = 16, txt = "F.   KETIDAKHADIRAN", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            pdf.ln(6)
            pdf.set_x(lmargin - 13.5)
            pdf.cell(w = (33 + 166.5 + 40.5 + 58.5), h = 21.6, txt = 'Ketidakhadiran', border = 1, align = 'C', fill = True)
            pdf.ln()
            pdf.set_x (lmargin - 13.5)
            pdf.set_font('arialN', '', 12)
            pdf.cell(w = 33, h = 21.6, txt = '1.', border = 1, align = 'C')
            pdf.cell(w = 166.5, h = 21.6, txt = 'Sakit', border = 1, align = 'L')
            pdf.cell(w = 40.5, h = 21.6, txt = str(sakit), border = 'LTB', align = 'C')
            pdf.cell(w = 58.5, h = 21.6, txt = 'Hari', border = 'RTB', align = 'L')
            pdf.ln()
            pdf.set_x (lmargin - 13.5)
            pdf.cell(w = 33, h = 21.6, txt = '2.', border = 1, align = 'C')
            pdf.cell(w = 166.5, h = 21.6, txt = 'Izin', border = 1, align = 'L')
            pdf.cell(w = 40.5, h = 21.6, txt = str(izin), border = 'LTB', align = 'C')
            pdf.cell(w = 58.5, h = 21.6, txt = 'Hari', border = 'RTB', align = 'L')
            pdf.ln()
            pdf.set_x (lmargin - 13.5)
            pdf.cell(w = 33, h = 21.6, txt = '3.', border = 1, align = 'C')
            pdf.cell(w = 166.5, h = 21.6, txt = 'Tanpa Keterangan', border = 1, align = 'L')
            pdf.cell(w = 40.5, h = 21.6, txt = str(tanpa_keterangan), border = 'LTB', align = 'C')
            pdf.cell(w = 58.5, h = 21.6, txt = 'Hari', border = 'RTB', align = 'L')
            pdf.ln()
            pdf.set_x (lmargin - 13.5)
            pdf.set_font('arialN', 'B', 12)
            pdf.cell(w = 33 + 166.5, h = 21.6, txt = 'Jumlah', border = 1, align = 'C')
            pdf.cell(w = 40.5, h = 21.6, txt = str(sakit + izin + tanpa_keterangan), border = ':TB', align = 'C')
            pdf.cell(w = 58.5, h = 21.6, txt = 'Hari', border = 'RTB', align = 'L')
            pdf.ln()
            
        class catatan():
            pdf.ln()
            pdf.set_x(lmargin - 31.5)
            pdf.set_font('arialN', 'B', 12)
            pdf.set_font_size(12)
            pdf.cell(w = 0, h = 16, txt = "G.   CATATAN WALI KELAS", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            pdf.ln(6)
            pdf.set_x(lmargin - 13.5)
            padding(pdf, w = 544.4, h_cell = 21.6, txt = notes, h_wanted = 62.3, align = 'L')
            pdf.ln()
            
            pdf.add_page()
            pdf.set_x(lmargin - 31.5)
            pdf.set_font('arialN', 'B', 12)
            pdf.set_font_size(12)
            pdf.cell(w = 0, h = 16, txt = "H.   TANGGAPAN ORANG TUA/WALI", border = 0, new_x = XPos.LMARGIN, new_y = YPos.NEXT, align = 'L')
            pdf.ln(6)
            pdf.set_x(lmargin - 13.5)
            pdf.cell(w = 544.4, h = 62.3, txt = '', border = 1, align = 'L')
            pdf.ln()
            pdf.ln(86.4)
            yttd = pdf.get_y()
            pdf.multi_cell(w = 333.9, h = 21.6, txt = 'Mengetahui,\nOrang Tua/Wali \n\n\n\n ……………………………', align = 'L')
            pdf.set_xy(lmargin+333.9, yttd)
            
            now = datetime.now()
            pdf.multi_cell(w = 333.9, h = 21.6, txt = f'Balikpapan, {now.strftime("%e %B %Y")}\nWali Kelas \n\n\n\n{WaliKelas} \nNIK. {NIK_WaliKelas}', align = 'L')
            pdf.ln()
            pdf.cell(w = 162, h = 129.6, txt = '')
            pdf.multi_cell(w = 540.9-162-lmargin, h = 21.6, txt = 'Mengetahui, \nKepala Sekolah, \n\n\n\nRuth Murwani Dumasthary, S.Sos., M.Pd. \nNIK. 06.01.010714', align = 'L')
            
        pdf.output(f'{output_dir}/{Nama}_{Kelas}_{Jurusan}.pdf')
        print(f'{Nama}\'s reportcard PRINTED')
    print('All Done!!!')
    
if __name__ == "__main__":
    WaliKelas = ''
    NIK_WaliKelas = ''
    Kelas = '12.1'
    Semester = '1'
    Thn = '2022/2023'
    Description = 'Description.xlsx'
    Ekskul = 'Ekskul'
    absen_file = 'Absen Kelas_12.1.xlsx'
    output_dir = 'Output'
    createpdf(WaliKelas, NIK_WaliKelas, Kelas, Semester, Thn, Description, Ekskul, absen_file, output_dir)