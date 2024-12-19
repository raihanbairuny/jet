import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import streamlit as st
import matplotlib
from io import BytesIO

pd.set_option('display.max_columns', None)

# Fungsi untuk mengunggah file
def upload_files():
    gl_file = st.file_uploader("Upload GL File", type=["xlsx"])
    log_file = st.file_uploader("Upload Log File (optional)", type=["xlsx"])
    return gl_file, log_file

# Fungsi untuk memproses GL dan Log
def process_files(gl_file, log_file, client_name):
    gl = pd.read_excel(gl_file)
    
    # Rename kolom GL
    glmap = {
        'Description Awal': ['Document Number', 'Debit/Credit ind', 'Posting period', 'Amount in Doc. Curr.', 'G/L Account', 'Document Date', 'Entry Date', 'Time of Entry', 'User Name'],
        'Jet': ['Journal_ID', 'Amount_Credit_Debit_Indicator', 'Period', 'Amount', 'GL_Account_Number', 'Document_Date', 'Entered_Date', 'Entered_Time', 'Entered_By']
    }
    colz = dict(zip(glmap['Description Awal'], glmap['Jet']))
    gl.rename(columns=colz, inplace=True)

    # Cek dan hitung kolom Net
    if 'Amount' in gl.columns:
        if not ((gl['Amount'] > 0).any() and (gl['Amount'] < 0).any()):
            gl['Net'] = gl.apply(lambda x: x['Amount'] * -1 if x['Amount_Credit_Debit_Indicator'] == 'H' else x['Amount'], axis=1)
        else:
            gl['Net'] = gl['Amount']  # Jika sudah ada positif dan negatif, set Net sama dengan Amount

    # Proses Log jika ada
    if log_file is not None:
        log = pd.read_excel(log_file)
        logmap = {
            'Description Awal': ['Document Number', 'Debit/Credit ind', 'Amount'],
            'Jet': ['Journal_ID', 'Amount_Credit_Debit_Indicator', 'Amount']
        }
        log_colz = dict(zip(logmap['Description Awal'], logmap['Jet']))
        log.rename(columns=log_colz, inplace=True)

        # Jalankan semua fungsi
        rg = check_for_gaps_in_JE_ID(gl)
        ec = comparison_of_entries_of_GL_and_log_file(gl, log)
        ac = comparison_of_amounts_of_GL_and_log_file(gl, log)
        igl = check_for_incomplete_entries(gl, amount_ready=True)
        de = check_for_duplicate_entries(gl, amount_ready=True)
        rae = check_for_round_dollar_entries(gl, amount_ready=True)
        lp = check_for_post_date_entries(gl, amount_ready=True)
        we = check_for_weekend_entries(gl, date_format='%Y-%m-%d%H:%M:%S')
        ne = check_for_nights_entries(gl, date_format='%Y-%m-%d%H:%M:%S')
        ru = check_for_rare_users(gl)
        ra = check_for_rare_accounts(gl)
        bf = benford(gl)

        # Simpan hasil ke file Excel
        save_results_to_excel(rg, ec, ac, igl, de, rae, lp, we, ne, ru, ra, bf, client_name)
    else:
        # Jalankan fungsi tanpa Log
        rg = check_for_gaps_in_JE_ID(gl)
        igl = check_for_incomplete_entries(gl, amount_ready=False)
        de = check_for_duplicate_entries(gl, amount_ready=False)
        rae = check_for_round_dollar_entries(gl, amount_ready=False)
        lp = check_for_post_date_entries(gl, amount_ready=False)
        we = check_for_weekend_entries(gl, date_format='%Y-%m-%d%H:%M:%S')
        ne = check_for_nights_entries(gl, date_format='%Y-%m-%d%H:%M:%S')
        ru = check_for_rare_users(gl)
        ra = check_for_rare_accounts(gl)
        bf = benford(gl)

        # Simpan hasil ke file Excel
        save_results_to_excel(rg, None, None, igl , de, rae, lp, we, ne, ru, ra, bf, client_name)

# Fungsi untuk menyimpan hasil ke file Excel
def save_results_to_excel(rg, ec, ac, igl, de, rae, lp, we, ne, ru, ra, bf, client_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # if rg:
        #     pd.DataFrame({'range_gap_awal': [i[0] for i in rg], 'range_gap_akhir': [i[-1] for i in rg]}).to_excel(writer, sheet_name="check_for_gaps", index=False)
        rg.to_excel(writer, sheet_name='Range_Gaps', index=False)
        if ec is not None:
            pd.DataFrame(ec, index=[0]).to_excel(writer, sheet_name="entry_comparison", index=False)
        if ac is not None:
            ac.to_excel(writer, sheet_name="amount_comparison", index=False)
        igl.to_excel(writer, sheet_name="incomplete_gl_entries", index=False)
        rae.to_excel(writer, sheet_name="round_amount_entries", index=False)
        de.to_excel(writer, sheet_name="duplicate_entries", index=False)
        we.to_excel(writer, sheet_name="weekend_entries", index=False)
        ne.to_excel(writer, sheet_name='night_entries', index=False)
        ru.to_excel(writer, sheet_name="rare_users", index=False)
        ra.to_excel(writer, sheet_name="rare_accounts", index=False)
        bf.to_excel(writer, sheet_name="benford's_law", index=False)

    output.seek(0)
    return output

# Fungsi untuk memeriksa celah dalam ID Jurnal
def check_for_gaps_in_JE_ID(GL_Detail, Journal_ID_Column='Journal_ID'):
    gaps = []
    previous = None
    for item in GL_Detail[Journal_ID_Column]:
        if previous and (item - previous > 1):
            gaps.append([previous, item])
        previous = item
    if gaps:
        return pd.DataFrame({'range_gap_awal': [i[0] for i in gaps], 'range_gap_akhir': [i[1] for i in gaps]}).astype({'range_gap_awal': int, 'range_gap_akhir': int})
    else:
        return pd.DataFrame(columns=['range_gap_awal', 'range_gap_akhir'])

# Fungsi perbandingan entri GL dan Log
def comparison_of_entries_of_GL_and_log_file(GL_Detail_YYYYMMDD_YYYYMMDD, Log_File_YYYYMMDD_YYYYMMDD):
    In_GL_not_in_LOG = set(GL_Detail_YYYYMMDD_YYYYMMDD['Journal_ID']) - set(Log_File_YYYYMMDD_YYYYMMDD['Journal_ID'])
    In_LOG_not_in_GL = set(Log_File_YYYYMMDD_YYYYMMDD['Journal_ID']) - set(GL_Detail_YYYYMMDD_YYYYMMDD['Journal_ID'])
    return {
        "In_GL_not_in_LOG": len(In_GL_not_in_LOG),
        "In_LOG_not_in_GL": len(In_LOG_not_in_GL),
        "In_both": len(set(GL_Detail_YYYYMMDD_YYYYMMDD['Journal_ID']) & set(Log_File_YYYYMMDD_YYYYMMDD['Journal_ID']))
    }

# Fungsi perbandingan jumlah GL dan Log
def comparison_of_amounts_of_GL_and_log_file(GL_Detail_YYYYMMDD_YYYYMMDD, Log_File_YYYYMMDD_YYYYMMDD):
    gl_totals_pivot = GL_Detail_YYYYMMDD_YYYYMMDD.pivot_table(index=['Journal_ID', 'Amount_Credit_Debit_Indicator'], 
                  values='Net', 
                  aggfunc=sum).reset_index()
    if 'Total' not in Log_File_YYYYMMDD_YYYYMMDD.columns:
        Log_File_YYYYMMDD_YYYYMMDD['Total'] = 0
    recon_gl_to_log = gl_totals_pivot.merge(Log_File_YYYYMMDD_YYYYMMDD, on=['Journal_ID', 'Amount_Credit_Debit_Indicator'], 
                                            how='outer').fillna(0)
    recon_gl_to_log['Comparison'] = round(abs(recon_gl_to_log['Net']), 2) - round(abs(recon_gl_to_log['Total']), 2)
    failed_test = recon_gl_to_log.loc[recon_gl_to_log['Comparison'] != 0]
    return failed_test

# Fungsi untuk memeriksa entri yang tidak lengkap
def check_for_incomplete_entries(GL_Detail_YYYYMMDD_YYYYMMDD, amount_ready=True):
    if 'Net' not in GL_Detail_YYYYMMDD_YYYYMMDD.columns:
        st.warning("Kolom 'Net' tidak ditemukan. Pastikan kolom 'Amount' sudah dihitung dengan benar.")
        return pd.DataFrame()  # Kembalikan DataFrame kosong jika kolom tidak ada

    if amount_ready:
        GL_Pivot = GL_Detail_YYYYMMDD_YYYYMMDD.pivot_table(index='Journal_ID', values='Net', aggfunc=sum)
        failed_test = GL_Pivot.loc[round(GL_Pivot['Net'], 2) != 0]
    else:
        GL_Pivot = GL_Detail_YYYYMMDD_YYYYMMDD.pivot_table(index='Journal_ID', values='Amount', aggfunc=sum)
        failed_test = GL_Pivot.loc[round(GL_Pivot['Amount'], 2) != 0]
    return pd.DataFrame(failed_test.to_records())

# Fungsi untuk memeriksa entri duplikat
def check_for_duplicate_entries(GL_Detail_YYYYMMDD_YYYYMMDD, amount_ready=True):
    if amount_ready:
        net = 'Amount'
    else:
        net = 'Net'

    GL_Pivot = GL_Detail_YYYYMMDD_YYYYMMDD.pivot_table(index=['GL_Account_Number', 'Period', net], 
                                                        values='Journal_ID', aggfunc=np.count_nonzero)
    GL_Copy = GL_Detail_YYYYMMDD_YYYYMMDD[['Journal_ID', 'GL_Account_Number', 'Period', net]].copy()
    GL_Pivot.columns = ['Journal_Entry_Count']
    Duplicates = GL_Pivot.loc[GL_Pivot['Journal_Entry_Count'] != 1]
    Duplicates = pd.DataFrame(Duplicates.to_records())

    failed_test = GL_Copy.merge(Duplicates, on=['GL_Account_Number', 'Period', net], how='right').fillna(0)
    print('%d instances detected' % len(failed_test['Journal_ID']))
    return failed_test

# Fungsi untuk memeriksa entri jumlah bulat
def check_for_round_dollar_entries(GL_Detail_YYYYMMDD_YYYYMMDD, amount_ready=True):
    if amount_ready:
        net = 'Amount'
    else:
        net = 'Net'
    GL_Copy = GL_Detail_YYYYMMDD_YYYYMMDD[['Journal_ID', 'GL_Account_Number', 'Period', net]].copy()
    GL_Copy['1000s Remainder'] = GL_Copy[net] % 1000
    failed_test = GL_Copy.loc[GL_Copy['1000s Remainder'] == 0]
    print('%d instances detected' % len(failed_test['Journal_ID']))
    return failed_test

# Fungsi untuk memeriksa entri tanggal post
def check_for_post_date_entries(GL_Detail_YYYYMMDD_YYYYMMDD, amount_ready=True):
    if amount_ready:
        net = 'Amount'
    else:
        net = 'Net'


    GL_Copy = GL_Detail_YYYYMMDD_YYYYMMDD[['Journal_ID', 'Document_Date', 'Entered_Date', 'Period', net]].copy()
    if GL_Copy['Entered_Date'].dtypes.name == 'datetime64[ns]':
        failed_test = GL_Copy.loc[GL_Copy['Document_Date'] > (GL_Copy['Entered_Date'] + timedelta(days=100))]
    else:
        failed_test = GL_Copy.loc[GL_Copy['Document_Date'] > (GL_Copy['Entered_Date'] + 100)]
    print('%d instances detected' % len(failed_test['Journal_ID']))
    return failed_test

# Fungsi untuk memeriksa entri akhir pekan
def check_for_weekend_entries(GL_Detail_YYYYMMDD_YYYYMMDD, date_format='%Y%m%d%H%M%S'):
    GL_Copy = GL_Detail_YYYYMMDD_YYYYMMDD[['Journal_ID', 'Entered_Date', 'Entered_Time']].copy()
    ed = GL_Copy['Entered_Date']
    et = GL_Copy['Entered_Time']
    if type(GL_Copy['Entered_Date'].iloc[0]) != str:
        ed = GL_Copy['Entered_Date'].astype(str)
    if type(GL_Copy['Entered_Time'].iloc[0]) != str:
        et = GL_Copy['Entered_Time'].astype(str)

    ed.fillna(value='0000-00-00', inplace=True)
    et.fillna(value='00:00:00', inplace=True)

    GL_Copy['Entry_Date_Time_Formatted'] = pd.to_datetime(ed + et, format=date_format, errors='coerce')
    GL_Copy['WeekDayNo'] = GL_Copy['Entry_Date_Time_Formatted'].apply(lambda x: x.isoweekday())
    failed_test = GL_Copy.loc[GL_Copy['WeekDayNo'] >= 6]
    print('%d instances detected' % len(failed_test['Journal_ID']))
    return failed_test

# Fungsi untuk memeriksa entri malam
def check_for_nights_entries(GL_Detail_YYYYMMDD_YYYYMMDD, date_format='%Y%m%d%H%M%S'):
    print('Checking for Night Entries is started')
    GL_Copy = GL_Detail_YYYYMMDD_YYYYMMDD[['Journal_ID', 'Entered_Date', 'Entered_Time']].copy()
    ed = GL_Copy['Entered_Date']
    et = GL_Copy['Entered_Time']
    if type(GL_Copy['Entered_Date'].iloc[0]) != str:
        ed = GL_Copy['Entered_Date'].astype(str)
    if type(GL_Copy['Entered_Time'].iloc[0]) != str:
        et = GL_Copy['Entered_Date'].astype(str)
    ed.fillna(value='0000-00-00', inplace=True)
    et.fillna(value='00:00:00', inplace=True)
    GL_Copy['Entry_Date_Time_Formatted'] = pd.to_datetime(GL_Copy['Entered_Date'].astype(str) + 
                                                          GL_Copy['Entered_Time'].astype(str), format=date_format, errors='coerce')
    GL_Copy['Hour'] = GL_Copy['Entry_Date_Time_Formatted'].dt.hour
    failed_test = GL_Copy.loc[(GL_Copy['Hour'] >= 20) | (GL_Copy['Hour'] <= 5)]
    print('%d instances detected' % len(failed_test['Journal_ID']))
    return failed_test

# Fungsi untuk memeriksa pengguna langka
# def check_for_rare_users(GL_Detail_YYYYMMDD_YYYYMMDD):
#     print('Checking for Rare Users is started')
#     GL_Pivot = GL_Detail_YYYYMMDD_YYYYMMDD.pivot_table(index=['Entered_By'], values='Journal_ID', 
#                                                        aggfunc=np.count_nonzero).fillna(0)
#     Rare_Users = GL_Pivot.loc[GL_Pivot['Journal_ID'] <= 10]
#     Rare_Users = pd.DataFrame(Rare_Users.to_records())
#     GL_Copy = GL_Detail_YYYYMMDD_YYYYMMDD[['Journal_ID', 'GL_Account_Number', 'Entered_By']].copy()
#     failed_test = GL_Copy.merge(Rare_Users, on=['Entered_By'], how='right').fillna(0)
#     failed_test = failed_test.rename(columns={'Journal_ID_x': 'Journal_ID', 'Journal_ID_y': 'Entered_Count'})
#     print('%d instances detected' % len(failed_test['Entered_By']))
#     return failed_test

def check_for_rare_users(GL_Detail_YYYYMMDD_YYYYMMDD):
    print('Checking for Rare Users is started')
    
    # Membuat pivot untuk menghitung jumlah entri berdasarkan 'Entered_By'
    GL_Pivot = GL_Detail_YYYYMMDD_YYYYMMDD.pivot_table(
        index=['Entered_By'], 
        values='Journal_ID', 
        aggfunc=np.count_nonzero
    ).fillna(0)
    
    # Memilih pengguna dengan entri <= 10
    Rare_Users = GL_Pivot.loc[GL_Pivot['Journal_ID'] <= 10]
    Rare_Users = pd.DataFrame(Rare_Users.to_records())
    
    # Membuat salinan data yang diperlukan dari GL_Detail
    GL_Copy = GL_Detail_YYYYMMDD_YYYYMMDD[['Journal_ID', 'GL_Account_Number', 'Entered_By']].copy()
    
    # Menggabungkan data GL_Copy dengan Rare_Users berdasarkan 'Entered_By'
    failed_test = GL_Copy.merge(Rare_Users, on=['Entered_By'], how='right').fillna(0)
    
    # Menjamin bahwa hanya satu kolom Journal_ID yang ada
    failed_test = failed_test.rename(columns={'Journal_ID_x': 'Journal_ID', 'Journal_ID_y': 'Entered_Count'})
    
    # Menampilkan jumlah entri yang terdeteksi
    print('%d instances detected' % len(failed_test['Entered_By']))
    
    return failed_test


# Fungsi untuk memeriksa akun langka
# def check_for_rare_accounts(GL_Detail_YYYYMMDD_YYYYMMDD):
#     print('Checking for Rare Accounts is started')
#     GL_Pivot = GL_Detail_YYYYMMDD_YYYYMMDD.pivot_table(index=['GL_Account_Number'], values='Journal_ID', 
#                                                         aggfunc=np.count_nonzero).fillna(0)
#     Rare_Accounts = GL_Pivot.loc[GL_Pivot['Journal_ID'] <= 3]
#     Rare_Accounts = pd.DataFrame(Rare_Accounts.to_records())
#     GL_Copy = GL_Detail_YYYYMMDD_YYYYMMDD[['Journal_ID', 'GL_Account_Number', 'Entered_By']].copy()
#     failed_test = GL_Copy.merge(Rare_Accounts, on=['GL_Account_Number'], how='right').fillna(0)
#     print('%d instances detected' % len(failed_test['GL_Account_Number']))
#     return failed_test

def check_for_rare_accounts(GL_Detail_YYYYMMDD_YYYYMMDD):
    print('Checking for Rare Accounts is started')
    
    # Membuat pivot untuk menghitung jumlah entri berdasarkan 'GL_Account_Number'
    GL_Pivot = GL_Detail_YYYYMMDD_YYYYMMDD.pivot_table(
        index=['GL_Account_Number'], 
        values='Journal_ID', 
        aggfunc=np.count_nonzero
    ).fillna(0)
    
    # Memilih akun dengan entri <= 3
    Rare_Accounts = GL_Pivot.loc[GL_Pivot['Journal_ID'] <= 3]
    Rare_Accounts = pd.DataFrame(Rare_Accounts.to_records())
    
    # Membuat salinan data yang diperlukan dari GL_Detail
    GL_Copy = GL_Detail_YYYYMMDD_YYYYMMDD[['Journal_ID', 'GL_Account_Number', 'Entered_By']].copy()
    
    # Menggabungkan data GL_Copy dengan Rare_Accounts berdasarkan 'GL_Account_Number'
    failed_test = GL_Copy.merge(Rare_Accounts, on=['GL_Account_Number'], how='right').fillna(0)
    
    # Mengganti nama kolom agar lebih jelas dan sesuai dengan kolom yang diinginkan
    failed_test = failed_test.rename(columns={'Journal_ID_x': 'Journal_ID', 'Journal_ID_y': 'Entered_Count'})
    
    # Menampilkan jumlah entri yang terdeteksi
    print('%d instances detected' % len(failed_test['GL_Account_Number']))
    
    return failed_test


# Fungsi untuk menerapkan hukum Benford
def benford(GL_Detail_YYYYMMDD_YYYYMMDD):
    GL_Detail_YYYYMMDD_YYYYMMDD['first_digit'] = GL_Detail_YYYYMMDD_YYYYMMDD['Amount'].astype(str).apply(lambda x: x[0] if x[0] != '-' else x[1])
    bf = GL_Detail_YYYYMMDD_YYYYMMDD['first_digit'].value_counts().reset_index(name='Countz')
    bf = bf[bf['first_digit'] != '0']
    bf['percentz'] = bf['Countz'].apply(lambda x: x / bf['Countz'].sum() * 100)
    bf.plot.bar(x='first_digit', y='percentz')
    return bf

# Fungsi utama untuk menjalankan aplikasi Streamlit
def main():
    st.title("GL and Log File Processor")
    client_name = st.text_input("Enter client name: ")
    gl_file, log_file = upload_files()
    
    # if gl_file is not None and client_name:
    #     process_files(gl_file, log_file, client_name)
    if st.button("Submit"):
        if client_name and gl_file:
            process_files(gl_file, log_file, client_name)
            st.success("Processing completed!")
            st.download_button(
                label="Download Report",
                data=save_results_to_excel(),
                file_name=f"report_jet_{client_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("Please upload the GL file and enter the client name.")

if __name__ == "__main__":
    main()
