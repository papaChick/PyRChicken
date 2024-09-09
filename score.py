import pandas as pd

groupA = ['Agama', 'Bahasa Indonesia', 'Bahasa Inggris', 'Matematika (Wajib)', 'PKN', 'Sejarah Indonesia'] #Kelompok A (Umum)
groupB = ['Seni Budaya', 'PJOK', 'PKWU'] #Kelompok B (Umum)
mutlok = ['ICT'] #Muatan Lokal
groupC_IPA = ['Matematika (Peminatan)', 'Biologi', 'Kimia', 'Fisika', 'Literature'] #Kelompok C (Peminatan) - IPA
groupC_IPS = ['Ekonomi', 'Sosiologi', 'Geografi', 'Sejarah Internasional', 'Literature'] #Kelompok C (Peminatan) - IPS
subj_IPA = groupA + groupB + mutlok + groupC_IPA #Mata Pelajaran Siswa IPA
subj_IPS = groupA + groupB + mutlok + groupC_IPS #Mata Pelajaran Siswa IPS

def collect_data(absen_file, kelas, semester: int, subj_path):
    col_IPA = subj_IPA.copy()
    col_IPA.insert(0, 'Nama')
    col_IPA.insert(1, 'NP/NK')
    col_IPS = subj_IPS.copy()
    col_IPS.insert(0, 'Nama')
    col_IPS.insert(1, 'NP/NK')
    df_IPA = pd.DataFrame(columns = col_IPA)
    df_IPS = pd.DataFrame(columns = col_IPS)
    df_absen = pd.read_excel(f'{absen_file}', header = 1, sheet_name = 'Profile')
    num = 0
   
    for i in df_absen.index:
        name = df_absen.loc[i]['Nama Siswa']
        stream = df_absen.loc[i]['Jurusan']
        subjects = ''
        list_pengetahuan = [name, 'NP']
        list_keterampilan = [name, 'NK']
        try:
            if stream == 'IPA':
                subjects = subj_IPA
                df_main = df_IPA
            elif stream == 'IPS':
                subjects = subj_IPS  
                df_main = df_IPS
                
        except ValueError:
            print ('Please input stream as "IPA" or "IPS"')
        
        except:
            print ('Please check the "Stream" column')
        
        for subj in subjects:
            df_pengetahuan = pd.read_excel(f'{subj_path}\{subj}_{kelas}_semester {semester}.xlsx', sheet_name = 'NP (S)', header = 2, index_col = 'Nama Siswa') #Nilai Pengetahuan
            #index = df_pengetahuan[df_pengetahuan['Nama Siswa'] == name].index[0] #name index
            nilai_pengetahuan = df_pengetahuan.loc[name]['NRap.S'] #filter nilai
            list_pengetahuan.append(nilai_pengetahuan)
            df_keterampilan = pd.read_excel(f'{subj_path}\{subj}_{kelas}_semester {semester}.xlsx', sheet_name = 'NK (S)', header = 1, index_col = 'Nama Siswa') #Nilai Keterampilan
            #index = df_keterampilan[df_keterampilan['Nama Siswa'] == name].index[0] #name index
            nilai_keterampilan = df_keterampilan.loc[name]['Rata2 NK'] #filter nilai
            list_keterampilan.append(nilai_keterampilan)
        num += 1
        df_main.loc[num] = list_pengetahuan
        num += 1
        df_main.loc[num] = list_keterampilan
    print (df_IPA)
    print (df_IPS)
    df_IPA.to_csv('reportcard/IPA.csv', index = False)
    #df_IPA.to_csv('IPA.csv', index = True)
    df_IPS.to_csv('reportcard/IPS.csv', index = False)
    #df_IPS.to_csv('IPA.csv', index = True)

if __name__ == "__main__":
    collect_data('Absen Kelas_12.1.xlsx', 12.1, 1, 'Subjects')