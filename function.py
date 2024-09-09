from fpdf import FPDF
from fpdf.enums import XPos, YPos
import pandas as pd

def h_multi(self: str, w: int, txt: str):
        with self.offset_rendering() as dummy:
            y = dummy.get_y()
            dummy.multi_cell(w = w, h = 1, txt = txt)
            num_line = (dummy.get_y()-y)
        return num_line

def padding(self: str, w: int, h_cell: int, txt: str, h_wanted: int, align: str):
        y1 = self.get_y()
        self.cell(w = w, h = (h_wanted - h_cell*h_multi(self, w, txt))/2, border = 'LTR', new_x = XPos.LEFT, new_y = YPos.NEXT)
        self.multi_cell(w = w, h = h_cell, txt = txt, border = 'LR', align = align, new_x = XPos.LEFT, new_y = YPos.NEXT)
        self.cell(w = w, h = (h_wanted - h_cell*h_multi(self, w, txt))/2, border = 'LBR')
        self.set_xy(self.get_x(),y1)
        
class students(FPDF):
    def __init__(self, nama, kelas, jurusan, semester: int, desc_path, ekskul_dir):
        self.nama = nama.title()
        self.kelas = kelas
        self.jurusan = jurusan.upper()
        self.ekskul_dir = ekskul_dir
        self.semester = semester
        self.desc_path = desc_path
        super().__init__('P', 'pt', 'A4')
        global lmargin
        lmargin=42.55
        tmargin=90
        rmargin=42.55
        self.set_margins(lmargin, tmargin, rmargin)
        self.add_page()
        self.add_font('arialN', '', 'reportcard/fonts/ArialN.ttf')
        self.add_font('arialN', 'B', 'reportcard/fonts/ArialNB.ttf')
        
        #groups and subjects
        group_a = ['Kelompok A (Umum)',
                    {'Pendidikan Agama dan Budi Pekerti':'Agama',
                    'Pendidikan Pancasila dan Kewarganegaraan':'PKN',
                    'Bahasa Indonesia':'Bahasa Indonesia',
                    'Matematika (Wajib)':'Matematika (Wajib)',
                    'Sejarah Indonesia':'Sejarah Indonesia',
                    'Bahasa Inggris':'Bahasa Inggris'}]
        group_b = ['Kelompok B (Umum)',
                    {'Seni Budaya':'Seni Budaya',
                    'Pendidikan Jasmani, Olah Raga dan Kesehatan':'PJOK',
                    'Prakarya dan Kewirausahaan':'PKWU'}]
        muatan_lokal = ['Muatan Lokal',
                    {'Teknologi dan Ilmu Komunikasi':'ICT'}]
        group_cIPA = ['Kelompok C (Peminatan)', 
                    {'Matematika (Peminatan)':'Matematika (Peminatan)',
                    'Biologi':'Biologi',
                    'Kimia':'Kimia',
                    'Fisika':'Fisika',
                    'Pemilihan Lintas Minat/Pendalaman Materi: Sastra Inggris':'Literature'}]
        group_cIPS = ['Kelompok C (Peminatan)',
                    {'Ekonomi':'Ekonomi', 
                    'Geografi':'Geografi', 
                    'Sejarah Internasional':'Sejarah Internasional', 
                    'Sosiologi':'Sosiologi', 
                    'Pemilihan Lintas Minat/Pendalaman Materi: Sastra Inggris':'Literature'}]
        
        #create overall subject groups which depends on stream (IPA/IPS)
        if self.jurusan == 'IPA':
            self.subj_groups = (group_a, group_b, muatan_lokal, group_cIPA)
        elif self.jurusan == 'IPS':
            self.subj_groups = (group_a, group_b, muatan_lokal, group_cIPS)
    
    def header(self):
        self.image('reportcard/watermark.png', 0, 0, 595, 842)
        self.image('reportcard/header.png', 390, 5, 189.36, 76.32)
    
    def footer(self):
        footertxt = "      " + self.nama + " || Halaman " + str(self.page_no()) + " dari {nb}"
        self.image('reportcard/footer.png', -16.75, 817, 628.5, 24.8)
        self.set_y(785.3)
        self.set_font('Times', 'B', 10)
        lenline = (595 - self.get_string_width(footertxt))/2
        self.set_line_width(0.25)
        self.set_draw_color(79, 129, 189)
        self.line(42.55, 792.9, lenline, 792.9)
        self.line(595 - lenline, 792.9, 595 - 42.55, 792.9)
        self.cell(w = 0, h = 15.2, txt = footertxt, border = 0, align = 'C')
        
    def nilai_subject(self, np_nk):
        #width of cells
        lmargin = 42.55
        w_no = 34.1
        w_subj = 163
        w_score = 42.5
        w_pred = 56.8
        w_desc = 249.2
        
        #creating dataframe of a student
        self.main_df = pd.read_csv(f'reportcard/{self.jurusan}.csv') #read IPA.csv or IPS.csv
        index = self.main_df[(self.main_df['Nama'] == self.nama) & (self.main_df['NP/NK'] == np_nk.upper())].index[0] #get index which has the same name
        student_df = self.main_df.loc[index] #get row from dataframe with the same index
        subj_df = student_df.drop(['Nama', 'NP/NK']) #remove 'Nama' and 'NP/NK' from dataframe
        
        #create header
        self.set_font('arialN', 'B', 12)
        self.set_x(lmargin - 13.5)
        self.set_fill_color(217,217,217)
        self.cell(w = w_no, h = 21.6, txt = 'No', border = 1, align = 'C', fill = True)
        self.cell(w = w_subj, h = 21.6, txt = 'Mata Pelajaran', border = 1, align = 'C', fill = True)
        self.cell(w = w_score, h = 21.6, txt = 'Nilai', border = 1, align = 'C', fill = True)
        self.cell(w = w_pred, h = 21.6, txt = 'Predikat', border = 1, align = 'C', fill = True)
        self.cell(w = w_desc, h = 21.6, txt = 'Deskripsi', border = 1, align = 'C', fill = True)
        self.ln()
        self.total = 0
        list_avail_group = []
        for group in self.subj_groups: #iteration, for groups in list of groups
            self.set_x(lmargin - 13.5)
            self.set_font('arialN', 'B', 12)
            self.cell(w = w_no + w_subj + w_score + w_pred + w_desc, h = 21.6, txt = str(group[0]), border = 1, align = 'L', new_x = XPos.LEFT, new_y = YPos.NEXT) #place groups cells
            no = 0
            for subj in subj_df.index: #iteration, for each subjects index in the dataframe
                if subj in group[1].values(): #to check whether the subjects index is in the subjects list
                    no += 1
                    list_avail_group.append(subj)
                    score = int(subj_df[subj])
                    self.total += score
                    #predicate & desc if
                    if score >= 92:
                        pred = 'A'
                        append_desc = 'sangat baik'
                    elif score >= 83:
                        pred = 'B'
                        append_desc = 'baik'
                    elif score >= 75:
                        pred = 'C'
                        append_desc = 'cukup baik'
                    else:
                        pred = 'D'
                        append_desc = 'namun masih memerlukan bimbingan'
                        
                    desc_df = pd.read_excel(f'{self.desc_path}', header = 3, sheet_name = f'{self.kelas[:2]} - Semester {self.semester}')
                    desc_df_renew = desc_df.set_index('Nama Mapel')
                    translate = {'np':'KD PENGETAHUAN', 'nk':'KD KETERAMPILAN'}
                    desc = f'Siswa menunjukkan kemampuannya {append_desc} dalam {desc_df_renew.loc[subj][translate[np_nk]]}'
                    desc = desc.capitalize()
                    
                    values = list(group[1].values())
                    ind = values.index(subj)
                    keys = list(group[1].keys())
                    subj_full = keys[ind]
                    with self.offset_rendering() as dummy:
                        self.set_font('arialN', 'B', 12)
                        h1 = h_multi(self, w = 163, txt = subj_full) #returns subject cell height
                        self.set_font('arialN', '', 10)
                        h2 = h_multi(self, w = 249.2, txt = desc)  #returns description cell height
                        self.set_font('arialN', 'B', 12)
                        num_line = max(h1, h2) #check which one is taller in height
                        padding(dummy, w = 163, h_cell = 21.6, txt = str(subj_full), h_wanted = 21.6 * num_line, align = 'L')
                        padding(dummy, w = 249.2, h_cell = 21.6, txt = str(desc), h_wanted = 21.6 * num_line, align = 'J')
                    if dummy.page_break_triggered:
                        self.add_page() #add page if dummy cell trigger page break
                    self.set_x(lmargin - 13.5)
                    self.set_font('arialN', '', 12)
                    self.cell(w = w_no, h = 21.6 * num_line, txt = str(no), border = 1, align = 'L') #number
                    padding(self, w = w_subj, h_cell = 21.6, txt = str(subj_full), h_wanted = 21.6 * num_line, align = 'L') #subject
                    self.cell(w = w_score, h = 21.6 * num_line, txt = str(round(score)), border = 1, align = 'C') #score
                    self.cell(w = w_pred, h = 21.6 * num_line, txt = str(pred), border = 1, align = 'C') #predicate
                    self.set_font('arialN', '', 10)
                    padding(self, w = w_desc, h_cell = 21.6, txt = str(desc), h_wanted = 21.6 * num_line, align = 'J') #description
                    self.set_font('arialN', 'B', 12)
                    self.ln(21.6 * num_line) #line break
                    self.set_font('arialN', 'B', 12)
                    
        self.set_font('arialN', 'B', 12)
        self.set_x(lmargin - 13.5)
        self.cell(w = w_no + w_subj, h = 21.6, txt = 'Jumlah', border = 1, align = 'L')
        self.cell(w = w_score, h = 21.6, txt = str(round(self.total)), border = 1, align = 'C')
        self.cell(w = w_pred + w_desc, h = 21.6, txt = '', border = 1, align = 'L')
        self.ln()
        self.set_x(lmargin - 13.5)
        self.cell(w = w_no + w_subj, h = 21.6, txt = 'Rata-Rata', border = 1, align = 'L')
        self.cell(w = w_score, h = 21.6, txt = str(round(self.total/len(list_avail_group))), border = 1, align = 'C')
        self.cell(w = w_pred + w_desc, h = 21.6, txt = '', border = 1, align = 'L')
        
    def total_table(self):
        main_df = pd.read_csv(f'reportcard/{self.jurusan}.csv') #read IPA.csv or IPS.csv
        score = 0
        np_nk = ['NP', 'NK']
        for nilai in np_nk:
            index = main_df[(main_df['Nama'] == self.nama) & (main_df['NP/NK'] == nilai.upper())].index[0] #get index which has the same name
            student_df = main_df.loc[index]
            subj_df = student_df.drop(['Nama', 'NP/NK'])
            for subj in subj_df.index:
                score = score + subj_df[subj]    
        avg_score = (score)/(2*len(subj_df.index))
        
        with self.offset_rendering() as dummy:
            dummy.cell(w = 34.1 + 163, h = 21.6, txt = 'Jumlah Total', border = 1, align = 'L')
            dummy.cell(w = 42.5, h = 21.6, txt = '', border = 1, align = 'C')
            dummy.cell(w = 56.8 + 249.2, h = 21.6, txt = '', border = 1, align = 'L')
            dummy.ln()
            dummy.cell(w = 34.1 + 163, h = 21.6, txt = 'Rata-Rata', border = 1, align = 'L')
            dummy.cell(w = 42.5, h = 21.6, txt = '', border = 1, align = 'C')
            dummy.cell(w = 56.8 + 249.2, h = 21.6, txt = '', border = 1, align = 'L')
            dummy.ln()
        if dummy.page_break_triggered:
            self.add_page()
        self.set_x(lmargin - 13.5)
        self.cell(w = 34.1 + 163, h = 21.6, txt = 'Jumlah Total', border = 1, align = 'L')
        self.cell(w = 42.5, h = 21.6, txt = str(round(score)), border = 1, align = 'C')
        self.cell(w = 56.8 + 249.2, h = 21.6, txt = '', border = 1, align = 'L')
        self.ln()
        self.set_x(lmargin - 13.5)
        self.cell(w = 34.1 + 163, h = 21.6, txt = 'Rata-Rata Total', border = 1, align = 'L')
        self.cell(w = 42.5, h = 21.6, txt = str(round(avg_score)), border = 1, align = 'C')
        self.cell(w = 56.8 + 249.2, h = 21.6, txt = '', new_x = XPos.START, new_y = YPos.NEXT, border = 1, align = 'L')
        self.ln()
    
    def nilai_ekskul(self, no, ekskul):
        ekskul_list = {'':'',
                       'Archery':'Panahan',
                       'Badminton':'Bulu Tangkis',
                       'Band':'Band',
                       'Basketball':'Basket',
                       'Content Creator':'Pencipta Konten',
                       'Cooking':'Tata Boga',
                       'Dance':'Menari',
                       'Graphic Design':'Desain Grafis',
                       'Illustration':'Ilustrasi',
                       'Knitting':'Merajut',
                       'Mini Soccer':'Futsal',
                       'Public Speaking':'Keterampilan Berbicara',
                       'Robotic':'Robotik',
                       'Bowling':'Boling',
                       'Coding':'Pemrograman',
                       'Swimming':'Renang',
                       'Tae Kwon Do': 'Tae Kwon Do'}
        if ekskul != '':
            desc_df = pd.read_excel(f'{self.ekskul_dir}/{ekskul}_2022-2023.xlsx', header = 2, usecols = ['Name', 'Total Score', 'Total Predicate'], index_col = 'Name')
            desc_df = desc_df.loc[self.nama]
            pred = desc_df['Total Predicate']
            if pred == 'A':
                keterangan_append = 'Sangat baik'
            elif pred == 'B':
                keterangan_append = 'Baik'
            elif pred == 'C':
                keterangan_append = 'Cukup baik'
            elif pred == 'D':
                keterangan_append = 'Kurang baik'
            keterangan = f'{keterangan_append}, dalam mengikuti ekstrakurikuler {ekskul_list[ekskul]} yang diadakan sekolah.'
        else:
            pred = ''
            keterangan = ''
        self.set_x(lmargin - 13.5)
        num_line = h_multi(self, w = 246.9, txt = str(keterangan))
        self.cell(w = 27.7, h = 21.6*num_line, txt = str(no), border = 1, align = 'C')
        self.cell(w = (181.5 - 27.7), h = 21.6*num_line, txt = f'{ekskul_list[ekskul]}', border = 1, align = 'L')
        self.cell(w = 117.1, h = 21.6*num_line, txt = str(pred), border = 1, align = 'C')
        self.multi_cell(w = 246.9, h = 21.6, txt = str(keterangan), border = 1, align = 'L')
    
if __name__ == '__main__':
    pdf = students('Wangi Aminah', '12.1', 'ipa', '1', 'Description.xlsx', 'Ekskul')
    
    pdf.set_font('arialN', '', 12)
    pdf.nilai_subject('nk')
    pdf.nilai_subject('np')
    pdf.ln()
    pdf.nilai_ekskul('1.', 'Badminton')
    pdf.output('text.pdf')