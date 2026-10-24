#!/usr/bin/env python3
"""
=============================================================================
GENERATOR LAPORAN OTOMATIS HASIL TRYOUT SISWA (REVISI: PEMISAHAN IRT & BIASA)
=============================================================================
Fitur Utama:
1. Pemisahan Grafik Skala IRT (0 - 1000) dan Skala Biasa (0 - 100)
   agar pergerakan nilai Biasa tetap terlihat jelas dan proporsional.
2. Analisis Detail Per Siswa dengan rincian khusus mata pelajaran yang diambil
   (menyembunyikan mata pelajaran yang kosong/NaN).
3. Grafik per mapel terpisah untuk IRT dan Biasa per siswa.
4. Peringkat Top Siswa dinamis (pilihan dropdown 5, 10, 15, Semua).
5. Tabel Kinerja Mata Pelajaran pembanding IRT vs Biasa.
=============================================================================
"""

import os
import sys
import glob
import re
import json
import pandas as pd
import numpy as np

# ---------------------------------------------------------------------------
# KONFIGURASI MATA PELAJARAN & ALIAS NIS
# ---------------------------------------------------------------------------
KNOWN_SUBJECTS = {
    'INDONESIA': ('INDONESIA', 'Wajib'),
    'BAHASA INDONESIA': ('INDONESIA', 'Wajib'),
    'INGGRIS': ('INGGRIS', 'Wajib'),
    'BAHASA INGGRIS': ('INGGRIS', 'Wajib'),
    'MATEMATIKA': ('MATEMATIKA', 'Wajib'),
    'MATEMATIKA-LANJUT': ('MATEMATIKA-LANJUT', 'Pilihan'),
    'MATLANJUT': ('MATEMATIKA-LANJUT', 'Pilihan'),
    'MATEMATIKA LANJUT': ('MATEMATIKA-LANJUT', 'Pilihan'),
    'FISIKA': ('FISIKA', 'Pilihan'),
    'KIMIA': ('KIMIA', 'Pilihan'),
    'BIOLOGI': ('BIOLOGI', 'Pilihan'),
    'EKONOMI': ('EKONOMI', 'Pilihan'),
    'GEOGRAFI': ('GEOGRAFI', 'Pilihan'),
    'SOSIOLOGI': ('SOSIOLOGI', 'Pilihan'),
    'PPKN': ('PPKN', 'Pilihan'),
    'PENDIDIKAN PANCASILA': ('PPKN', 'Pilihan'),
    'PENDIDIKAN': ('PPKN', 'Pilihan'),
    'ARAB': ('ARAB', 'Pilihan'),
    'BAHASA ARAB': ('ARAB', 'Pilihan'),
    'INGGRIS-LANJUT': ('INGGRIS-LANJUT', 'Pilihan'),
    'INGGRIS LANJUT': ('INGGRIS-LANJUT', 'Pilihan'),
    'BAHASA INGGRIS LANJUT': ('INGGRIS-LANJUT', 'Pilihan'),
}

# Merger otomatis data siswa jika ada salah input NIS di masa lalu
NIS_ALIAS_MAP = {
    '250062': ('240060', 'ZAIM MUKAARIM'),
}

# ---------------------------------------------------------------------------
# 1. PARSING NAMA FILE
# ---------------------------------------------------------------------------
def parse_filename(filepath):
    """
    Mengekstrak 3 informasi dari nama file:
    - Jenis Nilai : 'IRT' atau 'Biasa'
    - Kategori    : 'WAJIB' atau 'PILIHAN'
    - Tryout ke-X : integer
    Contoh: 'IRT PILIHAN 1.xlsx' -> ('IRT', 'PILIHAN', 1)
    """
    basename = os.path.basename(filepath)
    stem, _ = os.path.splitext(basename)
    stem_upper = stem.upper()

    # 1. Jenis Nilai
    if 'BIASA' in stem_upper:
        score_type = 'Biasa'
    elif 'IRT' in stem_upper:
        score_type = 'IRT'
    else:
        score_type = 'IRT'  # Default fallback jika tidak ada keterangan

    # 2. Kategori
    if 'WAJIB' in stem_upper:
        category = 'WAJIB'
    elif 'PILIHAN' in stem_upper:
        category = 'PILIHAN'
    else:
        category = 'UMUM'

    # 3. Tryout ke-berapa
    m = re.search(r'(?:TO|TRYOUT|WAJIB|PILIHAN|IRT|BIASA)?\s*(\d+)', stem_upper)
    if m:
        tryout_num = int(m.group(1))
    else:
        m2 = re.search(r'(\d+)', stem_upper)
        tryout_num = int(m2.group(1)) if m2 else 1

    return score_type, category, tryout_num


# ---------------------------------------------------------------------------
# 2. DETEKSI SHEET TERBAIK DALAM FILE EXCEL
# ---------------------------------------------------------------------------
def find_best_sheet(xl):
    """
    Menemukan sheet yang berisi kolom Rekap Nilai paling lengkap.
    """
    best_sheet = None
    best_count = -1
    for sheet_name in xl.sheet_names:
        try:
            df_head = xl.parse(sheet_name, nrows=3)
            cols = [str(c).strip().upper() for c in df_head.columns]
            has_nis = any('NIS' in c for c in cols)
            has_nama = any('NAMA' in c for c in cols)
            if not (has_nis and has_nama):
                continue
            count = sum(1 for c in cols if c in KNOWN_SUBJECTS)
            if count > best_count:
                best_count = count
                best_sheet = sheet_name
        except Exception:
            continue

    if best_sheet is None:
        best_sheet = xl.sheet_names[0]
    return best_sheet


# ---------------------------------------------------------------------------
# 3. MEMUAT & MENSTANDARKAN DATA
# ---------------------------------------------------------------------------
def load_all_scores(data_dir="data_nilai"):
    """
    Membaca seluruh file .xlsx dari folder data_dir (atau fallback ke root jika kosong).
    Mengembalikan DataFrame dengan kolom:
    ['nis', 'nama', 'mapel', 'kategori', 'tryout', 'jenis', 'score']
    """
    files = []
    if os.path.exists(data_dir):
        files = sorted(glob.glob(os.path.join(data_dir, '*.xlsx')))
        if not files:
            files = sorted(glob.glob(os.path.join(data_dir, '*.xls')))

    if not files:
        print(f"[INFO] Folder '{data_dir}' kosong/tidak ditemukan. Mencari file di direktori saat ini...")
        files = sorted(glob.glob('*.xlsx'))

    if not files:
        raise FileNotFoundError(f"Tidak ditemukan file .xlsx di '{data_dir}' maupun di direktori saat ini.")

    print(f"[INFO] Ditemukan {len(files)} file Excel untuk diproses:")
    records = []

    for filepath in files:
        score_type, category, tryout_num = parse_filename(filepath)
        print(f"  -> Memproses: {os.path.basename(filepath)} [Jenis: {score_type}, Kategori: {category}, TO: {tryout_num}]")

        try:
            xl = pd.ExcelFile(filepath)
            sheet_name = find_best_sheet(xl)
            df = xl.parse(sheet_name)
        except Exception as e:
            print(f"     [PERINGATAN] Gagal membaca {filepath}: {e}")
            continue

        # Cari kolom NIS dan NAMA
        col_nis = next((c for c in df.columns if 'NIS' in str(c).upper()), None)
        col_nama = next((c for c in df.columns if 'NAMA' in str(c).upper()), None)

        if not col_nis or not col_nama:
            print(f"     [PERINGATAN] Kolom NIS atau NAMA tidak ditemukan di sheet '{sheet_name}'. Dilewati.")
            continue

        # Ekstrak data nilai per baris
        for _, row in df.iterrows():
            raw_nis = str(row[col_nis]).strip().replace('.0', '')
            raw_nama = str(row[col_nama]).strip()

            if not raw_nis or raw_nis.lower() in ('nan', 'none', ''):
                continue

            # Perbaiki alias NIS jika ada
            if raw_nis in NIS_ALIAS_MAP:
                raw_nis, raw_nama = NIS_ALIAS_MAP[raw_nis]

            for col in df.columns:
                col_norm = str(col).strip().upper()
                if col_norm in KNOWN_SUBJECTS:
                    std_mapel, std_cat = KNOWN_SUBJECTS[col_norm]
                    val = row[col]
                    try:
                        num_val = float(val)
                        if pd.notna(num_val):
                            records.append({
                                'nis': str(raw_nis),
                                'nama': str(raw_nama),
                                'mapel': std_mapel,
                                'kategori': std_cat,
                                'tryout': tryout_num,
                                'jenis': score_type,
                                'score': round(num_val, 2)
                            })
                    except (ValueError, TypeError):
                        pass

    if not records:
        raise ValueError("Tidak ada data nilai yang berhasil diekstrak dari file Excel.")

    df_scores = pd.DataFrame(records)
    print(f"\n[SUKSES] Total {len(df_scores)} catatan nilai dimuat untuk {df_scores['nis'].nunique()} siswa.")
    return df_scores


# ---------------------------------------------------------------------------
# 4. KALKULASI STATISTIK DENGAN PANDAS
# ---------------------------------------------------------------------------
def compute_dashboard_data(df):
    """
    Menghitung seluruh metrik analitik:
    - Ringkasan Eksekutif
    - Peringkat Siswa (IRT vs Biasa)
    - Kinerja Mata Pelajaran (IRT vs Biasa)
    - Tren Antar Tryout (IRT vs Biasa)
    - Detail per Mata Pelajaran
    - Detail per Siswa (hanya mapel yang nilainya valid/tidak kosong)
    """
    tryouts = sorted(df['tryout'].unique().tolist())
    has_irt = 'IRT' in df['jenis'].values
    has_biasa = 'Biasa' in df['jenis'].values

    # 1. Ringkasan Eksekutif
    total_students = df['nis'].nunique()
    irt_df = df[df['jenis'] == 'IRT']
    biasa_df = df[df['jenis'] == 'Biasa']

    overall_avg_irt = round(float(irt_df['score'].mean()), 2) if has_irt and len(irt_df) else None
    overall_avg_biasa = round(float(biasa_df['score'].mean()), 2) if has_biasa and len(biasa_df) else None

    unique_wajib = sorted(df[df['kategori'] == 'Wajib']['mapel'].unique().tolist())
    unique_pilihan = sorted(df[df['kategori'] == 'Pilihan']['mapel'].unique().tolist())
    all_subjects = sorted(df['mapel'].unique().tolist())

    # 2. Peringkat Siswa
    student_meta = df.groupby('nis')['nama'].last().to_dict()

    st_irt = irt_df.groupby('nis')['score'].agg(avg_irt='mean', count_irt='count').reset_index() if has_irt and len(irt_df) else pd.DataFrame(columns=['nis', 'avg_irt', 'count_irt'])
    st_biasa = biasa_df.groupby('nis')['score'].agg(avg_biasa='mean', count_biasa='count').reset_index() if has_biasa and len(biasa_df) else pd.DataFrame(columns=['nis', 'avg_biasa', 'count_biasa'])

    all_nis = pd.DataFrame({'nis': list(student_meta.keys())})
    st_merged = all_nis.merge(st_irt, on='nis', how='left').merge(st_biasa, on='nis', how='left')
    st_merged['nama'] = st_merged['nis'].map(student_meta)

    # Rank IRT & Biasa langsung menggunakan pandas rank (mengabaikan NaN)
    st_merged['avg_irt'] = pd.to_numeric(st_merged['avg_irt'], errors='coerce')
    st_merged['avg_biasa'] = pd.to_numeric(st_merged['avg_biasa'], errors='coerce')

    st_merged['rank_irt'] = st_merged['avg_irt'].rank(ascending=False, method='min', na_option='bottom')
    st_merged.loc[st_merged['avg_irt'].isna(), 'rank_irt'] = np.nan

    st_merged['rank_biasa'] = st_merged['avg_biasa'].rank(ascending=False, method='min', na_option='bottom')
    st_merged.loc[st_merged['avg_biasa'].isna(), 'rank_biasa'] = np.nan

    # 3. Kinerja Mata Pelajaran
    subj_list = []
    for m in all_subjects:
        m_df = df[df['mapel'] == m]
        cat = m_df['kategori'].iloc[0]
        m_irt = m_df[m_df['jenis'] == 'IRT']['score']
        m_biasa = m_df[m_df['jenis'] == 'Biasa']['score']

        avg_irt = round(float(m_irt.mean()), 2) if len(m_irt) else None
        max_irt = round(float(m_irt.max()), 2) if len(m_irt) else None
        min_irt = round(float(m_irt.min()), 2) if len(m_irt) else None

        avg_biasa = round(float(m_biasa.mean()), 2) if len(m_biasa) else None
        max_biasa = round(float(m_biasa.max()), 2) if len(m_biasa) else None
        min_biasa = round(float(m_biasa.min()), 2) if len(m_biasa) else None

        subj_list.append({
            'mapel': m,
            'kategori': cat,
            'avg_irt': avg_irt,
            'max_irt': max_irt,
            'min_irt': min_irt,
            'avg_biasa': avg_biasa,
            'max_biasa': max_biasa,
            'min_biasa': min_biasa,
            'count_peserta': int(m_df['nis'].nunique())
        })

    subj_df = pd.DataFrame(subj_list)

    # 4. Tren Antar Tryout (Keseluruhan, Wajib, Pilihan)
    def get_tryout_trend(sub_df):
        trend = {'tryouts': tryouts, 'avg_irt': {}, 'avg_biasa': {}}
        for t in tryouts:
            t_irt = sub_df[(sub_df['tryout'] == t) & (sub_df['jenis'] == 'IRT')]['score']
            t_biasa = sub_df[(sub_df['tryout'] == t) & (sub_df['jenis'] == 'Biasa')]['score']
            trend['avg_irt'][str(t)] = round(float(t_irt.mean()), 2) if len(t_irt) else None
            trend['avg_biasa'][str(t)] = round(float(t_biasa.mean()), 2) if len(t_biasa) else None
        return trend

    overall_trend = {
        'all': get_tryout_trend(df),
        'wajib': get_tryout_trend(df[df['kategori'] == 'Wajib']),
        'pilihan': get_tryout_trend(df[df['kategori'] == 'Pilihan'])
    }

    # 5. Detail per Mata Pelajaran
    subject_details = {}
    for m in all_subjects:
        m_df = df[df['mapel'] == m]
        cat = m_df['kategori'].iloc[0]
        m_tryouts = sorted(m_df['tryout'].unique().tolist())

        # Trend per TO untuk mapel ini
        to_trend_irt = {}
        to_trend_biasa = {}
        for t in m_tryouts:
            s_irt = m_df[(m_df['tryout'] == t) & (m_df['jenis'] == 'IRT')]['score']
            s_biasa = m_df[(m_df['tryout'] == t) & (m_df['jenis'] == 'Biasa')]['score']
            to_trend_irt[str(t)] = round(float(s_irt.mean()), 2) if len(s_irt) else None
            to_trend_biasa[str(t)] = round(float(s_biasa.mean()), 2) if len(s_biasa) else None

        # Data per siswa di mapel ini
        st_in_subj = []
        for nis, st_group in m_df.groupby('nis'):
            st_name = st_group['nama'].iloc[0]
            scores_irt = {}
            scores_biasa = {}
            for _, r in st_group.iterrows():
                t_str = str(r['tryout'])
                if r['jenis'] == 'IRT':
                    scores_irt[t_str] = r['score']
                elif r['jenis'] == 'Biasa':
                    scores_biasa[t_str] = r['score']

            avg_i = round(float(pd.Series(list(scores_irt.values())).mean()), 2) if scores_irt else None
            avg_b = round(float(pd.Series(list(scores_biasa.values())).mean()), 2) if scores_biasa else None

            st_in_subj.append({
                'nis': nis,
                'nama': st_name,
                'scores_irt': scores_irt,
                'avg_irt': avg_i,
                'scores_biasa': scores_biasa,
                'avg_biasa': avg_b,
            })

        st_in_subj.sort(key=lambda s: (s['avg_irt'] if s['avg_irt'] is not None else -9999, s['avg_biasa'] if s['avg_biasa'] is not None else -9999), reverse=True)

        subject_details[m] = {
            'name': m,
            'category': cat,
            'tryout_nums': m_tryouts,
            'avg_per_to_irt': to_trend_irt,
            'avg_per_to_biasa': to_trend_biasa,
            'students': st_in_subj
        }

    # 6. Detail per Siswa (Hanya mata pelajaran yang diambil / tidak NaN)
    student_details = {}
    for nis, st_group in df.groupby('nis'):
        st_name = st_group['nama'].iloc[0]
        st_info = st_merged[st_merged['nis'] == nis].iloc[0]

        # Rata-rata per tryout siswa ini
        st_to_avgs_irt = {}
        st_to_avgs_biasa = {}
        for t in tryouts:
            sub_to_irt = st_group[(st_group['tryout'] == t) & (st_group['jenis'] == 'IRT')]['score']
            sub_to_biasa = st_group[(st_group['tryout'] == t) & (st_group['jenis'] == 'Biasa')]['score']
            if len(sub_to_irt):
                st_to_avgs_irt[str(t)] = round(float(sub_to_irt.mean()), 2)
            if len(sub_to_biasa):
                st_to_avgs_biasa[str(t)] = round(float(sub_to_biasa.mean()), 2)

        # Mapel yang diambil siswa (Filter: hanya mapel yang memiliki minimal 1 nilai valid)
        st_subjects = []
        for m, m_group in st_group.groupby('mapel'):
            cat = m_group['kategori'].iloc[0]
            sc_irt = {}
            sc_biasa = {}
            for _, r in m_group.iterrows():
                t_str = str(r['tryout'])
                if r['jenis'] == 'IRT':
                    sc_irt[t_str] = r['score']
                elif r['jenis'] == 'Biasa':
                    sc_biasa[t_str] = r['score']

            # Abaikan mapel jika kosong/NaN di kedua jenis penilaian
            if not sc_irt and not sc_biasa:
                continue

            m_avg_irt = round(float(pd.Series(list(sc_irt.values())).mean()), 2) if sc_irt else None
            m_avg_biasa = round(float(pd.Series(list(sc_biasa.values())).mean()), 2) if sc_biasa else None

            st_subjects.append({
                'name': m,
                'category': cat,
                'scores_irt': sc_irt,
                'avg_irt': m_avg_irt,
                'has_irt': len(sc_irt) > 0,
                'scores_biasa': sc_biasa,
                'avg_biasa': m_avg_biasa,
                'has_biasa': len(sc_biasa) > 0
            })

        # Urutkan: Wajib dulu, lalu Pilihan berdasarkan rata-rata nilai
        st_subjects.sort(key=lambda s: (0 if s['category'] == 'Wajib' else 1, -(s['avg_irt'] if s['avg_irt'] is not None else -999)))

        student_details[nis] = {
            'nis': nis,
            'nama': st_name,
            'rank_irt': int(st_info['rank_irt']) if pd.notna(st_info['rank_irt']) else None,
            'rank_biasa': int(st_info['rank_biasa']) if pd.notna(st_info['rank_biasa']) else None,
            'avg_irt': round(float(st_info['avg_irt']), 2) if pd.notna(st_info['avg_irt']) else None,
            'avg_biasa': round(float(st_info['avg_biasa']), 2) if pd.notna(st_info['avg_biasa']) else None,
            'count_irt': int(st_info['count_irt']) if pd.notna(st_info['count_irt']) else 0,
            'count_biasa': int(st_info['count_biasa']) if pd.notna(st_info['count_biasa']) else 0,
            'tryout_avgs_irt': st_to_avgs_irt,
            'tryout_avgs_biasa': st_to_avgs_biasa,
            'subjects': st_subjects
        }

    return {
        'summary': {
            'total_students': total_students,
            'overall_avg_irt': overall_avg_irt,
            'overall_avg_biasa': overall_avg_biasa,
            'total_scores_irt': len(irt_df),
            'total_scores_biasa': len(biasa_df),
            'total_wajib': len(unique_wajib),
            'total_pilihan': len(unique_pilihan),
            'tryouts': tryouts,
            'has_irt': has_irt,
            'has_biasa': has_biasa
        },
        'students_ranking': st_merged.to_dict(orient='records'),
        'subjects_summary': subj_df.to_dict(orient='records'),
        'overall_trend': overall_trend,
        'subject_details': subject_details,
        'student_details': student_details
    }


# ---------------------------------------------------------------------------
# 5. GENERATE RAW HTML SINGLE FILE (DENGAN GRAFIK TERPISAH IRT & BIASA)
# ---------------------------------------------------------------------------
def generate_html_report(data, output_path="index.html"):
    """
    Menghasilkan file HTML tunggal dengan desain UI/UX kelas dunia,
    CSS Glassmorphism responsif, dan Chart.js interaktif.
    """
    summary = data['summary']
    overall_trend = data['overall_trend']
    subjects_summary = data['subjects_summary']
    students_ranking = data['students_ranking']
    subject_details = data['subject_details']
    student_details = data['student_details']

    # Serialisasi data ke JSON untuk Chart.js & skrip interaktif
    json_summary = json.dumps(summary, ensure_ascii=False)
    json_overall_trend = json.dumps(overall_trend, ensure_ascii=False)
    json_subjects_summary = json.dumps(subjects_summary, ensure_ascii=False)
    json_students_ranking = json.dumps(students_ranking, ensure_ascii=False)
    json_subject_details = json.dumps(subject_details, ensure_ascii=False)
    json_student_details = json.dumps(student_details, ensure_ascii=False)

    irt_avg_display = f"{summary['overall_avg_irt']:.2f}" if summary['overall_avg_irt'] is not None else "—"
    biasa_avg_display = f"{summary['overall_avg_biasa']:.2f}" if summary['overall_avg_biasa'] is not None else "—"
    tryouts_list = ", ".join([f"TO {t}" for t in summary['tryouts']])

    sorted_students_for_dropdown = sorted(student_details.values(), key=lambda s: s['nama'])
    student_options_html = "".join([
        f'<option value="{s["nis"]}">{s["nama"]} (NIS: {s["nis"]})</option>'
        for s in sorted_students_for_dropdown
    ])

    sorted_mapel_for_dropdown = sorted(subject_details.keys())
    mapel_options_html = "".join([
        f'<option value="{m}">{m} ({subject_details[m]["category"]})</option>'
        for m in sorted_mapel_for_dropdown
    ])

    html_content = f"""<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Laporan Hasil Tryout Siswa — Pemisahan Grafik Nilai IRT & Nilai Biasa</title>
    <meta name="description" content="Dashboard interaktif evaluasi tryout siswa dengan visualisasi grafik terpisah untuk Nilai IRT (0-1000) dan Nilai Biasa (0-100).">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
    <style>
        :root {{
            --bg-primary: #0b0f19;
            --bg-secondary: #131b2e;
            --bg-card: rgba(19, 27, 46, 0.75);
            --bg-card-hover: rgba(30, 41, 69, 0.9);
            --text-primary: #f8fafc;
            --text-secondary: #94a3b8;
            --text-muted: #64748b;
            --accent-irt: #818cf8;
            --accent-irt-glow: rgba(99, 102, 241, 0.25);
            --accent-biasa: #10b981;
            --accent-biasa-glow: rgba(16, 185, 129, 0.25);
            --accent-amber: #f59e0b;
            --accent-rose: #f43f5e;
            --border: rgba(51, 65, 85, 0.5);
            --border-highlight: rgba(99, 102, 241, 0.35);
            --radius-sm: 8px;
            --radius-md: 12px;
            --radius-lg: 16px;
            --radius-xl: 20px;
            --shadow-card: 0 10px 25px -5px rgba(0, 0, 0, 0.4), 0 8px 10px -6px rgba(0, 0, 0, 0.3);
            --transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        }}

        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        body {{
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
            background: var(--bg-primary);
            color: var(--text-primary);
            min-height: 100vh;
            line-height: 1.6;
            overflow-x: hidden;
            background-image: 
                radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.12) 0px, transparent 50%),
                radial-gradient(at 100% 100%, rgba(16, 185, 129, 0.08) 0px, transparent 50%),
                radial-gradient(at 50% 50%, rgba(15, 23, 42, 0.8) 0px, transparent 100%);
            background-attachment: fixed;
        }}

        /* Header Navigation */
        .top-navbar {{
            position: sticky;
            top: 0;
            z-index: 100;
            background: rgba(11, 15, 25, 0.85);
            backdrop-filter: blur(16px);
            border-bottom: 1px solid var(--border);
            padding: 14px 28px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}

        .brand {{
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        .brand-badge {{
            background: linear-gradient(135deg, #6366f1, #8b5cf6);
            color: white;
            font-weight: 800;
            font-size: 14px;
            padding: 6px 12px;
            border-radius: var(--radius-sm);
            letter-spacing: 0.5px;
            box-shadow: 0 0 15px rgba(99, 102, 241, 0.4);
        }}

        .brand-title {{
            font-size: 17px;
            font-weight: 700;
            letter-spacing: -0.3px;
        }}

        .nav-links {{
            display: flex;
            gap: 16px;
        }}

        .nav-links a {{
            color: var(--text-secondary);
            text-decoration: none;
            font-size: 13.5px;
            font-weight: 500;
            transition: var(--transition);
            padding: 6px 12px;
            border-radius: var(--radius-sm);
        }}

        .nav-links a:hover {{
            color: var(--text-primary);
            background: rgba(255, 255, 255, 0.06);
        }}

        /* Container */
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            padding: 32px 24px 60px;
        }}

        /* Hero Header */
        .hero {{
            text-align: center;
            padding: 24px 16px 36px;
        }}

        .hero-badge {{
            display: inline-flex;
            align-items: center;
            gap: 8px;
            background: rgba(99, 102, 241, 0.12);
            border: 1px solid rgba(99, 102, 241, 0.3);
            color: var(--accent-irt);
            padding: 6px 16px;
            border-radius: 999px;
            font-size: 13px;
            font-weight: 600;
            margin-bottom: 16px;
        }}

        .hero h1 {{
            font-size: clamp(26px, 4vw, 38px);
            font-weight: 900;
            letter-spacing: -1px;
            margin-bottom: 12px;
            background: linear-gradient(135deg, #ffffff 40%, #94a3b8 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}

        .hero p {{
            color: var(--text-secondary);
            font-size: 15px;
            max-width: 760px;
            margin: 0 auto;
        }}

        /* Badges */
        .badge {{
            display: inline-flex;
            align-items: center;
            gap: 4px;
            font-size: 11px;
            font-weight: 700;
            padding: 3px 8px;
            border-radius: 6px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }}

        .badge-irt {{
            background: rgba(99, 102, 241, 0.18);
            color: var(--accent-irt);
            border: 1px solid rgba(99, 102, 241, 0.4);
        }}

        .badge-biasa {{
            background: rgba(16, 185, 129, 0.18);
            color: var(--accent-biasa);
            border: 1px solid rgba(16, 185, 129, 0.4);
        }}

        .badge-wajib {{
            background: rgba(59, 130, 246, 0.15);
            color: #60a5fa;
            border: 1px solid rgba(59, 130, 246, 0.3);
        }}

        .badge-pilihan {{
            background: rgba(168, 85, 247, 0.15);
            color: #c084fc;
            border: 1px solid rgba(168, 85, 247, 0.3);
        }}

        /* Section Layout */
        .section {{
            margin-bottom: 48px;
        }}

        .section-header {{
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
            margin-bottom: 20px;
            flex-wrap: wrap;
            gap: 16px;
        }}

        .section-title {{
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        .section-icon {{
            width: 40px;
            height: 40px;
            border-radius: var(--radius-md);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
            background: var(--bg-card);
            border: 1px solid var(--border);
        }}

        .section-title h2 {{
            font-size: 22px;
            font-weight: 800;
            letter-spacing: -0.5px;
        }}

        .section-title p {{
            font-size: 13px;
            color: var(--text-muted);
        }}

        /* Cards Grid */
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 18px;
            margin-bottom: 24px;
        }}

        .stat-card {{
            background: var(--bg-card);
            backdrop-filter: blur(12px);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            padding: 20px;
            box-shadow: var(--shadow-card);
            transition: var(--transition);
            position: relative;
            overflow: hidden;
        }}

        .stat-card:hover {{
            transform: translateY(-2px);
            border-color: var(--border-highlight);
            background: var(--bg-card-hover);
        }}

        .stat-card::before {{
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: linear-gradient(90deg, transparent, rgba(99, 102, 241, 0.6), transparent);
        }}

        .stat-label {{
            font-size: 12px;
            font-weight: 600;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 8px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .stat-value {{
            font-size: 28px;
            font-weight: 800;
            letter-spacing: -0.5px;
            margin-bottom: 6px;
        }}

        .stat-detail {{
            font-size: 12px;
            color: var(--text-muted);
        }}

        /* Table Styling */
        .table-card {{
            background: var(--bg-card);
            backdrop-filter: blur(12px);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow-card);
            overflow: hidden;
        }}

        .table-responsive {{
            overflow-x: auto;
            width: 100%;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13.5px;
            text-align: left;
        }}

        thead tr {{
            background: rgba(15, 23, 42, 0.7);
            border-bottom: 1px solid var(--border);
        }}

        th {{
            padding: 14px 16px;
            font-weight: 600;
            font-size: 12px;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            white-space: nowrap;
        }}

        td {{
            padding: 13px 16px;
            border-bottom: 1px solid rgba(51, 65, 85, 0.25);
            color: var(--text-primary);
            white-space: nowrap;
        }}

        tbody tr {{
            transition: background 0.15s ease;
        }}

        tbody tr:hover {{
            background: rgba(99, 102, 241, 0.05);
        }}

        .student-link {{
            color: var(--text-primary);
            text-decoration: none;
            font-weight: 600;
            cursor: pointer;
            transition: var(--transition);
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }}

        .student-link:hover {{
            color: var(--accent-irt);
            transform: translateX(2px);
        }}

        /* Forms & Selects */
        .select-control, .input-control {{
            background: rgba(15, 23, 42, 0.85);
            border: 1px solid var(--border);
            color: var(--text-primary);
            padding: 10px 16px;
            border-radius: var(--radius-md);
            font-size: 13.5px;
            font-family: inherit;
            outline: none;
            transition: var(--transition);
        }}

        .select-control:focus, .input-control:focus {{
            border-color: var(--accent-irt);
            box-shadow: 0 0 0 3px var(--accent-irt-glow);
        }}

        /* Dual Rankings & Dual Charts Grid */
        .grid-2col {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(480px, 1fr));
            gap: 24px;
        }}

        @media (max-width: 640px) {{
            .grid-2col {{
                grid-template-columns: 1fr;
            }}
        }}

        /* Chart Card */
        .chart-box {{
            background: var(--bg-card);
            backdrop-filter: blur(12px);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            padding: 22px;
            box-shadow: var(--shadow-card);
            display: flex;
            flex-direction: column;
        }}

        .chart-box-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 16px;
            flex-wrap: wrap;
            gap: 8px;
        }}

        .chart-box-title {{
            font-weight: 700;
            font-size: 14.5px;
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .chart-wrapper {{
            position: relative;
            height: 330px;
            width: 100%;
        }}

        /* Buttons & Filter Pills */
        .btn-group {{
            display: inline-flex;
            background: rgba(15, 23, 42, 0.8);
            border: 1px solid var(--border);
            border-radius: var(--radius-md);
            padding: 4px;
            gap: 4px;
        }}

        .btn-pill {{
            background: transparent;
            border: none;
            color: var(--text-secondary);
            font-size: 12.5px;
            font-weight: 600;
            padding: 6px 14px;
            border-radius: var(--radius-sm);
            cursor: pointer;
            transition: var(--transition);
        }}

        .btn-pill.active {{
            background: var(--accent-irt);
            color: white;
            box-shadow: 0 0 10px rgba(99, 102, 241, 0.4);
        }}

        .text-irt {{ color: var(--accent-irt); font-weight: 700; }}
        .text-biasa {{ color: var(--accent-biasa); font-weight: 700; }}

        /* Profile Banner in Detail */
        .profile-banner {{
            background: linear-gradient(135deg, rgba(99, 102, 241, 0.15), rgba(16, 185, 129, 0.15));
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            padding: 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 18px;
            margin-bottom: 24px;
        }}

        .profile-info h3 {{
            font-size: 22px;
            font-weight: 800;
            margin-bottom: 4px;
        }}

        .profile-badges {{
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
            margin-top: 8px;
        }}

        .rank-pill {{
            background: rgba(15, 23, 42, 0.6);
            border: 1px solid var(--border);
            padding: 6px 14px;
            border-radius: 999px;
            font-size: 12.5px;
            font-weight: 700;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }}

        .empty-notice {{
            display: flex;
            align-items: center;
            justify-content: center;
            height: 100%;
            color: var(--text-muted);
            font-size: 13.5px;
            text-align: center;
            padding: 30px;
            border: 2px dashed rgba(51, 65, 85, 0.4);
            border-radius: var(--radius-md);
        }}

        /* Footer */
        footer {{
            text-align: center;
            padding: 40px 20px;
            color: var(--text-muted);
            font-size: 13px;
            border-top: 1px solid var(--border);
            margin-top: 60px;
        }}
    </style>
</head>
<body>

    <!-- Top Navigation -->
    <header class="top-navbar">
        <div class="brand">
            <div class="brand-badge">TRYOUT ANALYTICS</div>
            <div class="brand-title">Laporan Nilai IRT & Biasa</div>
        </div>
        <nav class="nav-links">
            <a href="#dashboard">Dashboard</a>
            <a href="#rankings">Peringkat</a>
            <a href="#subjects">Kinerja Mapel</a>
            <a href="#trends">Tren Antar TO</a>
            <a href="#subject-detail">Detail Mapel</a>
            <a href="#student-detail">Detail Siswa</a>
        </nav>
    </header>

    <main class="container">

        <!-- Hero Title -->
        <section class="hero" id="dashboard">
            <div class="hero-badge">📊 Evaluasi Hasil Tryout Siswa</div>
            <h1>Dashboard Kinerja Tryout: Nilai IRT & Nilai Biasa</h1>
            <p>Visualisasi grafik skala independen: Nilai IRT (rentang 0 - 1000) dan Nilai Biasa (rentang 0 - 100) serta analisis detail per siswa dengan rincian mapel yang aktif dikerjakan.</p>
        </section>

        <!-- 1. Executive Summary Cards -->
        <section class="section">
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-label">👥 Total Siswa</div>
                    <div class="stat-value">{summary['total_students']}</div>
                    <div class="stat-detail">Siswa terdaftar aktif</div>
                </div>

                <div class="stat-card" style="border-top: 3px solid var(--accent-irt);">
                    <div class="stat-label">
                        <span>🔮 Rata-rata IRT Umum</span>
                        <span class="badge badge-irt">IRT</span>
                    </div>
                    <div class="stat-value text-irt">{irt_avg_display}</div>
                    <div class="stat-detail">Skala 0 - 1000 &bull; {summary['total_scores_irt']} nilai</div>
                </div>

                <div class="stat-card" style="border-top: 3px solid var(--accent-biasa);">
                    <div class="stat-label">
                        <span>🟢 Rata-rata Biasa Umum</span>
                        <span class="badge badge-biasa">Biasa</span>
                    </div>
                    <div class="stat-value text-biasa">{biasa_avg_display}</div>
                    <div class="stat-detail">{"Skala 0 - 100 • " + str(summary['total_scores_biasa']) + " nilai" if summary['total_scores_biasa'] > 0 else "Belum ada file data Biasa"}</div>
                </div>

                <div class="stat-card">
                    <div class="stat-label">📚 Cakupan Mapel</div>
                    <div class="stat-value">{summary['total_wajib'] + summary['total_pilihan']} <span style="font-size: 14px; color: var(--text-muted); font-weight: 500;">Mapel</span></div>
                    <div class="stat-detail">{summary['total_wajib']} Wajib &bull; {summary['total_pilihan']} Pilihan</div>
                </div>

                <div class="stat-card">
                    <div class="stat-label">🚀 Tryout Terdata</div>
                    <div class="stat-value">{len(summary['tryouts'])} <span style="font-size: 14px; color: var(--text-muted); font-weight: 500;">Sesi</span></div>
                    <div class="stat-detail">{tryouts_list}</div>
                </div>
            </div>
        </section>

        <!-- 2. Dual Rankings Section (Dinamis 5 / 10 / 15 / Semua) -->
        <section class="section" id="rankings">
            <div class="section-header">
                <div class="section-title">
                    <div class="section-icon">🏆</div>
                    <div>
                        <h2>Peringkat Siswa Terbaik</h2>
                        <p>Daftar peringkat terbaik berdasarkan Nilai IRT vs Nilai Biasa (Dinamis 5 / 10 / 15 / Semua)</p>
                    </div>
                </div>
                <div>
                    <label for="rankLimitSelect" style="font-size: 13px; color: var(--text-secondary); margin-right: 8px; font-weight: 600;">Tampilkan:</label>
                    <select id="rankLimitSelect" class="select-control" onchange="updateRankingsTable()">
                        <option value="5">5 Siswa</option>
                        <option value="10" selected>10 Siswa (Default)</option>
                        <option value="15">15 Siswa</option>
                        <option value="all">Semua Siswa</option>
                    </select>
                </div>
            </div>

            <div class="grid-2col">
                <!-- Tabel Peringkat IRT -->
                <div class="table-card">
                    <div style="padding: 16px 20px; background: rgba(99, 102, 241, 0.08); border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center;">
                        <div style="font-weight: 700; font-size: 15px; display: flex; align-items: center; gap: 8px;">
                            <span>🏆 Top Siswa (Berdasarkan Nilai IRT)</span>
                            <span class="badge badge-irt">IRT</span>
                        </div>
                    </div>
                    <div class="table-responsive">
                        <table>
                            <thead>
                                <tr>
                                    <th style="width: 50px;">Rank</th>
                                    <th>Nama Siswa</th>
                                    <th>NIS</th>
                                    <th style="text-align: right;">Rata-rata IRT</th>
                                    <th style="text-align: right;">Rata-rata Biasa</th>
                                    <th style="text-align: center;">Jml Nilai</th>
                                </tr>
                            </thead>
                            <tbody id="rankingIrtBody"></tbody>
                        </table>
                    </div>
                </div>

                <!-- Tabel Peringkat Biasa -->
                <div class="table-card">
                    <div style="padding: 16px 20px; background: rgba(16, 185, 129, 0.08); border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center;">
                        <div style="font-weight: 700; font-size: 15px; display: flex; align-items: center; gap: 8px;">
                            <span>🎯 Top Siswa (Berdasarkan Nilai Biasa)</span>
                            <span class="badge badge-biasa">Biasa</span>
                        </div>
                    </div>
                    <div class="table-responsive">
                        <table>
                            <thead>
                                <tr>
                                    <th style="width: 50px;">Rank</th>
                                    <th>Nama Siswa</th>
                                    <th>NIS</th>
                                    <th style="text-align: right;">Rata-rata Biasa</th>
                                    <th style="text-align: right;">Rata-rata IRT</th>
                                    <th style="text-align: center;">Jml Nilai</th>
                                </tr>
                            </thead>
                            <tbody id="rankingBiasaBody"></tbody>
                        </table>
                    </div>
                </div>
            </div>
        </section>

        <!-- 3. Tabel Kinerja Mata Pelajaran (IRT vs Biasa) -->
        <section class="section" id="subjects">
            <div class="section-header">
                <div class="section-title">
                    <div class="section-icon">📚</div>
                    <div>
                        <h2>Tabel Kinerja Mata Pelajaran</h2>
                        <p>Komparasi performa rata-rata, nilai tertinggi, dan nilai terendah antara IRT vs Biasa</p>
                    </div>
                </div>
                <div class="btn-group">
                    <button class="btn-pill active" onclick="filterSubjectTable('all', this)">Semua Mapel</button>
                    <button class="btn-pill" onclick="filterSubjectTable('Wajib', this)">Mapel Wajib</button>
                    <button class="btn-pill" onclick="filterSubjectTable('Pilihan', this)">Mapel Pilihan</button>
                </div>
            </div>

            <div class="table-card">
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr>
                                <th style="width: 50px;">#</th>
                                <th>Mata Pelajaran</th>
                                <th>Kategori</th>
                                <th style="text-align: right;" class="text-irt">Rata-rata IRT</th>
                                <th style="text-align: right;" class="text-biasa">Rata-rata Biasa</th>
                                <th style="text-align: right;">Tertinggi IRT</th>
                                <th style="text-align: right;">Terendah IRT</th>
                                <th style="text-align: right;">Tertinggi Biasa</th>
                                <th style="text-align: right;">Terendah Biasa</th>
                                <th style="text-align: center;">Peserta</th>
                            </tr>
                        </thead>
                        <tbody id="subjectTableBody"></tbody>
                    </table>
                </div>
            </div>
        </section>

        <!-- 4. Grafik Tren Tryout (TERPISAH: IRT & BIASA) -->
        <section class="section" id="trends">
            <div class="section-header">
                <div class="section-title">
                    <div class="section-icon">📈</div>
                    <div>
                        <h2>Grafik Tren Nilai Antar Tryout</h2>
                        <p>Visualisasi grafik terpisah dengan skala sumbu Y independen untuk Nilai IRT dan Nilai Biasa</p>
                    </div>
                </div>
                <div class="btn-group">
                    <button class="btn-pill active" onclick="switchOverallTrend('all', this)">Semua Mapel</button>
                    <button class="btn-pill" onclick="switchOverallTrend('wajib', this)">Mapel Wajib</button>
                    <button class="btn-pill" onclick="switchOverallTrend('pilihan', this)">Mapel Pilihan</button>
                </div>
            </div>

            <div class="grid-2col">
                <!-- Grafik 1: Tren IRT -->
                <div class="chart-box" style="border-top: 3px solid var(--accent-irt);">
                    <div class="chart-box-header">
                        <div class="chart-box-title">
                            <span>🔮 Tren Rata-rata Nilai IRT</span>
                            <span class="badge badge-irt">Skala 0 - 1000</span>
                        </div>
                    </div>
                    <div class="chart-wrapper">
                        <canvas id="overallTrendIrtCanvas"></canvas>
                    </div>
                </div>

                <!-- Grafik 2: Tren Biasa -->
                <div class="chart-box" style="border-top: 3px solid var(--accent-biasa);">
                    <div class="chart-box-header">
                        <div class="chart-box-title">
                            <span>🟢 Tren Rata-rata Nilai Biasa</span>
                            <span class="badge badge-biasa">Skala 0 - 100</span>
                        </div>
                    </div>
                    <div class="chart-wrapper" id="overallTrendBiasaWrapper">
                        <canvas id="overallTrendBiasaCanvas"></canvas>
                    </div>
                </div>
            </div>
        </section>

        <!-- 5. Laporan Detail Per Mata Pelajaran (TERPISAH: IRT & BIASA) -->
        <section class="section" id="subject-detail">
            <div class="section-header">
                <div class="section-title">
                    <div class="section-icon">📖</div>
                    <div>
                        <h2>Laporan Detail Per Mata Pelajaran</h2>
                        <p>Grafik tren terpisah dan daftar lengkap nilai siswa pada mata pelajaran terpilih</p>
                    </div>
                </div>
                <div>
                    <select id="subjectSelect" class="select-control" onchange="renderSubjectDetail(this.value)">
                        {mapel_options_html}
                    </select>
                </div>
            </div>

            <div id="subjectDetailContent"></div>
        </section>

        <!-- 6. Analisis Detail Per Siswa (HANYA MAPEL YANG DIAMBIL SISWA) -->
        <section class="section" id="student-detail">
            <div class="section-header">
                <div class="section-title">
                    <div class="section-icon">👤</div>
                    <div>
                        <h2>Analisis Detail Per Siswa</h2>
                        <p>Pilih siswa untuk melihat grafik per mapel yang aktif diambil serta riwayat nilainya (mapel kosong otomatis disembunyikan)</p>
                    </div>
                </div>
                <div style="display: flex; gap: 10px; flex-wrap: wrap;">
                    <input type="text" id="studentSearchInput" class="input-control" placeholder="🔍 Cari nama / NIS siswa..." oninput="filterStudentDropdown(this.value)">
                    <select id="studentSelect" class="select-control" onchange="renderStudentDetail(this.value)">
                        {student_options_html}
                    </select>
                </div>
            </div>

            <div id="studentDetailContent"></div>
        </section>

    </main>

    <footer>
        <p>Laporan Hasil Tryout Siswa &bull; Dibuat secara otomatis dengan Python & Pandas &bull; Grafik Skala Terpisah IRT & Biasa</p>
    </footer>

    <!-- INJEKSI DATASET JSON KE CLIENT-SIDE JAVASCRIPT -->
    <script>
        const summaryData = {json_summary};
        const overallTrendData = {json_overall_trend};
        const subjectsSummary = {json_subjects_summary};
        const studentsRanking = {json_students_ranking};
        const subjectDetails = {json_subject_details};
        const studentDetails = {json_student_details};

        // Palet warna cerah untuk membedakan mata pelajaran siswa
        const subjectPalette = [
            '#38bdf8', '#818cf8', '#c084fc', '#f472b6', '#fb7185',
            '#f97316', '#facc15', '#4ade80', '#2dd4bf', '#a78bfa',
            '#60a5fa', '#34d399', '#e879f9'
        ];

        // Global Chart instances
        let overallTrendIrtChart = null;
        let overallTrendBiasaChart = null;
        let subjectTrendIrtChart = null;
        let subjectTrendBiasaChart = null;
        let studentTrendIrtChart = null;
        let studentTrendBiasaChart = null;

        const tooltipStyle = {{
            backgroundColor: '#0f172a',
            titleColor: '#f8fafc',
            bodyColor: '#cbd5e1',
            borderColor: '#334155',
            borderWidth: 1,
            padding: 12,
            cornerRadius: 8,
        }};

        // Inisialisasi Saat Halaman Dimuat
        document.addEventListener('DOMContentLoaded', () => {{
            updateRankingsTable();
            filterSubjectTable('all');
            initOverallTrendCharts('all');
            
            const firstSubject = document.getElementById('subjectSelect').value;
            if (firstSubject) renderSubjectDetail(firstSubject);

            const firstStudent = document.getElementById('studentSelect').value;
            if (firstStudent) renderStudentDetail(firstStudent);
        }});

        // 1. UPDATE RANKINGS TABLE (Dinamis: 5, 10, 15, Semua)
        function updateRankingsTable() {{
            const limitVal = document.getElementById('rankLimitSelect').value;
            const limit = limitVal === 'all' ? studentsRanking.length : parseInt(limitVal, 10);

            // Sort Top IRT
            const irtRanked = [...studentsRanking]
                .filter(s => s.avg_irt !== null && s.avg_irt !== undefined)
                .sort((a, b) => b.avg_irt - a.avg_irt)
                .slice(0, limit);

            // Sort Top Biasa
            const biasaRanked = [...studentsRanking]
                .filter(s => s.avg_biasa !== null && s.avg_biasa !== undefined)
                .sort((a, b) => b.avg_biasa - a.avg_biasa)
                .slice(0, limit);

            // Render IRT Table
            const irtBody = document.getElementById('rankingIrtBody');
            if (irtRanked.length === 0) {{
                irtBody.innerHTML = `<tr><td colspan="6" style="text-align: center; color: var(--text-muted); padding: 24px;">Belum ada data nilai IRT</td></tr>`;
            }} else {{
                irtBody.innerHTML = irtRanked.map((s, idx) => `
                    <tr>
                        <td><strong>${{idx + 1}}</strong></td>
                        <td>
                            <a class="student-link" onclick="selectStudent('${{s.nis}}')">
                                ${{s.nama}} <span style="font-size: 11px; opacity: 0.7;">↗</span>
                            </a>
                        </td>
                        <td style="color: var(--text-muted);">${{s.nis}}</td>
                        <td style="text-align: right;" class="text-irt">${{s.avg_irt ? s.avg_irt.toFixed(2) : '—'}}</td>
                        <td style="text-align: right;" class="text-biasa">${{s.avg_biasa ? s.avg_biasa.toFixed(2) : '—'}}</td>
                        <td style="text-align: center; color: var(--text-muted);">${{s.count_irt || 0}}</td>
                    </tr>
                `).join('');
            }}

            // Render Biasa Table
            const biasaBody = document.getElementById('rankingBiasaBody');
            if (biasaRanked.length === 0) {{
                biasaBody.innerHTML = `<tr><td colspan="6" style="text-align: center; color: var(--text-muted); padding: 24px;">Data nilai 'Biasa' belum tersedia.<br><span style="font-size: 12px;">Tambahkan file Excel dengan kata 'BIASA' pada nama filenya untuk melihat peringkat ini.</span></td></tr>`;
            }} else {{
                biasaBody.innerHTML = biasaRanked.map((s, idx) => `
                    <tr>
                        <td><strong>${{idx + 1}}</strong></td>
                        <td>
                            <a class="student-link" onclick="selectStudent('${{s.nis}}')">
                                ${{s.nama}} <span style="font-size: 11px; opacity: 0.7;">↗</span>
                            </a>
                        </td>
                        <td style="color: var(--text-muted);">${{s.nis}}</td>
                        <td style="text-align: right;" class="text-biasa">${{s.avg_biasa ? s.avg_biasa.toFixed(2) : '—'}}</td>
                        <td style="text-align: right;" class="text-irt">${{s.avg_irt ? s.avg_irt.toFixed(2) : '—'}}</td>
                        <td style="text-align: center; color: var(--text-muted);">${{s.count_biasa || 0}}</td>
                    </tr>
                `).join('');
            }}
        }}

        // 2. FILTER SUBJECT TABLE
        function filterSubjectTable(category, btnElement) {{
            if (btnElement) {{
                const btns = btnElement.parentElement.querySelectorAll('.btn-pill');
                btns.forEach(b => b.classList.remove('active'));
                btnElement.classList.add('active');
            }}

            const filtered = category === 'all' 
                ? subjectsSummary 
                : subjectsSummary.filter(s => s.kategori === category);

            filtered.sort((a, b) => (b.avg_irt || 0) - (a.avg_irt || 0));

            const tbody = document.getElementById('subjectTableBody');
            tbody.innerHTML = filtered.map((s, idx) => `
                <tr>
                    <td>${{idx + 1}}</td>
                    <td>
                        <a class="student-link" onclick="selectSubject('${{s.mapel}}')">
                            ${{s.mapel}} <span style="font-size: 11px; opacity: 0.7;">↗</span>
                        </a>
                    </td>
                    <td>
                        <span class="badge ${{s.kategori === 'Wajib' ? 'badge-wajib' : 'badge-pilihan'}}">${{s.kategori}}</span>
                    </td>
                    <td style="text-align: right;" class="text-irt">${{s.avg_irt !== null ? s.avg_irt.toFixed(2) : '—'}}</td>
                    <td style="text-align: right;" class="text-biasa">${{s.avg_biasa !== null ? s.avg_biasa.toFixed(2) : '—'}}</td>
                    <td style="text-align: right; color: var(--text-secondary);">${{s.max_irt !== null ? s.max_irt.toFixed(1) : '—'}}</td>
                    <td style="text-align: right; color: var(--text-secondary);">${{s.min_irt !== null ? s.min_irt.toFixed(1) : '—'}}</td>
                    <td style="text-align: right; color: var(--text-secondary);">${{s.max_biasa !== null ? s.max_biasa.toFixed(1) : '—'}}</td>
                    <td style="text-align: right; color: var(--text-secondary);">${{s.min_biasa !== null ? s.min_biasa.toFixed(1) : '—'}}</td>
                    <td style="text-align: center; color: var(--text-muted);">${{s.count_peserta}}</td>
                </tr>
            `).join('');
        }}

        // 3. OVERALL TREND CHARTS (DUAL CHARTS: IRT & BIASA)
        function initOverallTrendCharts(category) {{
            const trend = overallTrendData[category];
            const labels = trend.tryouts.map(t => `Tryout ${{t}}`);
            const dataIrt = trend.tryouts.map(t => trend.avg_irt[String(t)] ?? null);
            const dataBiasa = trend.tryouts.map(t => trend.avg_biasa[String(t)] ?? null);

            // Chart IRT
            const ctxIrt = document.getElementById('overallTrendIrtCanvas');
            if (overallTrendIrtChart) overallTrendIrtChart.destroy();
            overallTrendIrtChart = new Chart(ctxIrt, {{
                type: 'line',
                data: {{
                    labels: labels,
                    datasets: [{{
                        label: 'Rata-rata Nilai IRT',
                        data: dataIrt,
                        borderColor: '#818cf8',
                        backgroundColor: 'rgba(99, 102, 241, 0.12)',
                        borderWidth: 3,
                        pointRadius: 6,
                        pointHoverRadius: 9,
                        tension: 0.25,
                        spanGaps: true,
                        fill: true
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{ tooltip: tooltipStyle }},
                    scales: {{
                        y: {{
                            suggestedMin: 200,
                            suggestedMax: 1000,
                            grid: {{ color: 'rgba(51, 65, 85, 0.3)' }},
                            title: {{ display: true, text: 'Skor IRT (Skala 0 - 1000)', color: '#818cf8', font: {{ weight: '600' }} }}
                        }},
                        x: {{ grid: {{ color: 'rgba(51, 65, 85, 0.15)' }} }}
                    }}
                }}
            }});

            // Chart Biasa
            const ctxBiasa = document.getElementById('overallTrendBiasaCanvas');
            if (overallTrendBiasaChart) overallTrendBiasaChart.destroy();
            overallTrendBiasaChart = new Chart(ctxBiasa, {{
                type: 'line',
                data: {{
                    labels: labels,
                    datasets: [{{
                        label: 'Rata-rata Nilai Biasa',
                        data: dataBiasa,
                        borderColor: '#10b981',
                        backgroundColor: 'rgba(16, 185, 129, 0.12)',
                        borderWidth: 3,
                        pointRadius: 6,
                        pointHoverRadius: 9,
                        tension: 0.25,
                        spanGaps: true,
                        fill: true
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{ tooltip: tooltipStyle }},
                    scales: {{
                        y: {{
                            min: 0,
                            max: 100,
                            grid: {{ color: 'rgba(51, 65, 85, 0.3)' }},
                            title: {{ display: true, text: 'Nilai Biasa (Skala 0 - 100)', color: '#10b981', font: {{ weight: '600' }} }}
                        }},
                        x: {{ grid: {{ color: 'rgba(51, 65, 85, 0.15)' }} }}
                    }}
                }}
            }});
        }}

        function switchOverallTrend(cat, btn) {{
            const btns = btn.parentElement.querySelectorAll('.btn-pill');
            btns.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            initOverallTrendCharts(cat);
        }}

        // 4. RENDER DETAIL PER MATA PELAJARAN (DUAL CHARTS)
        function renderSubjectDetail(mapelName) {{
            const s = subjectDetails[mapelName];
            if (!s) return;

            const container = document.getElementById('subjectDetailContent');
            const tryoutLabels = s.tryout_nums.map(n => `TO ${{n}}`);
            const irtTrend = s.tryout_nums.map(n => s.avg_per_to_irt[String(n)] ?? null);
            const biasaTrend = s.tryout_nums.map(n => s.avg_per_to_biasa[String(n)] ?? null);

            const allIrtScores = s.students.map(st => st.avg_irt).filter(v => v !== null);
            const allBiasaScores = s.students.map(st => st.avg_biasa).filter(v => v !== null);

            const avgIrt = allIrtScores.length ? (allIrtScores.reduce((a, b) => a + b, 0) / allIrtScores.length).toFixed(2) : '—';
            const avgBiasa = allBiasaScores.length ? (allBiasaScores.reduce((a, b) => a + b, 0) / allBiasaScores.length).toFixed(2) : '—';

            container.innerHTML = `
                <div class="stats-grid" style="margin-bottom: 20px;">
                    <div class="stat-card">
                        <div class="stat-label">Kategori Mapel</div>
                        <div class="stat-value" style="font-size: 20px;">${{s.category}}</div>
                        <div class="stat-detail">${{s.students.length}} siswa terdaftar</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-label text-irt">Rata-rata IRT</div>
                        <div class="stat-value text-irt">${{avgIrt}}</div>
                        <div class="stat-detail">Skala 0 - 1000</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-label text-biasa">Rata-rata Biasa</div>
                        <div class="stat-value text-biasa">${{avgBiasa}}</div>
                        <div class="stat-detail">Skala 0 - 100</div>
                    </div>
                </div>

                <!-- Dual Charts untuk Mapel -->
                <div class="grid-2col" style="margin-bottom: 24px;">
                    <div class="chart-box" style="border-top: 3px solid var(--accent-irt);">
                        <div class="chart-box-header">
                            <div class="chart-box-title">
                                <span>🔮 Tren Nilai IRT — ${{s.name}}</span>
                                <span class="badge badge-irt">Skala 0 - 1000</span>
                            </div>
                        </div>
                        <div class="chart-wrapper">
                            <canvas id="subjectTrendIrtCanvas"></canvas>
                        </div>
                    </div>

                    <div class="chart-box" style="border-top: 3px solid var(--accent-biasa);">
                        <div class="chart-box-header">
                            <div class="chart-box-title">
                                <span>🟢 Tren Nilai Biasa — ${{s.name}}</span>
                                <span class="badge badge-biasa">Skala 0 - 100</span>
                            </div>
                        </div>
                        <div class="chart-wrapper">
                            <canvas id="subjectTrendBiasaCanvas"></canvas>
                        </div>
                    </div>
                </div>

                <div class="table-card">
                    <div style="padding: 14px 20px; font-weight: 700; border-bottom: 1px solid var(--border);">
                        📋 Daftar Siswa & Riwayat Nilai — ${{s.name}}
                    </div>
                    <div class="table-responsive">
                        <table>
                            <thead>
                                <tr>
                                    <th style="width: 40px;">#</th>
                                    <th>Nama Siswa</th>
                                    <th>NIS</th>
                                    ${{s.tryout_nums.map(n => `<th style="text-align: right;">TO ${{n}} (IRT)</th>`).join('')}}
                                    <th style="text-align: right;" class="text-irt">Rata IRT</th>
                                    ${{summaryData.has_biasa ? s.tryout_nums.map(n => `<th style="text-align: right;">TO ${{n}} (Biasa)</th>`).join('') : ''}}
                                    <th style="text-align: right;" class="text-biasa">Rata Biasa</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${{s.students.map((st, idx) => `
                                    <tr>
                                        <td>${{idx + 1}}</td>
                                        <td><a class="student-link" onclick="selectStudent('${{st.nis}}')">${{st.nama}} ↗</a></td>
                                        <td style="color: var(--text-muted);">${{st.nis}}</td>
                                        ${{s.tryout_nums.map(n => `<td style="text-align: right;">${{st.scores_irt[String(n)] !== undefined ? st.scores_irt[String(n)] : '—'}}</td>`).join('')}}
                                        <td style="text-align: right;" class="text-irt">${{st.avg_irt !== null ? st.avg_irt.toFixed(2) : '—'}}</td>
                                        ${{summaryData.has_biasa ? s.tryout_nums.map(n => `<td style="text-align: right;">${{st.scores_biasa[String(n)] !== undefined ? st.scores_biasa[String(n)] : '—'}}</td>`).join('') : ''}}
                                        <td style="text-align: right;" class="text-biasa">${{st.avg_biasa !== null ? st.avg_biasa.toFixed(2) : '—'}}</td>
                                    </tr>
                                `).join('')}}
                            </tbody>
                        </table>
                    </div>
                </div>
            `;

            // Chart IRT Mapel
            const ctxIrt = document.getElementById('subjectTrendIrtCanvas');
            if (subjectTrendIrtChart) subjectTrendIrtChart.destroy();
            subjectTrendIrtChart = new Chart(ctxIrt, {{
                type: 'line',
                data: {{
                    labels: tryoutLabels,
                    datasets: [{{
                        label: `Rata-rata IRT (${{s.name}})`,
                        data: irtTrend,
                        borderColor: '#818cf8',
                        backgroundColor: 'rgba(99, 102, 241, 0.12)',
                        borderWidth: 3,
                        pointRadius: 6,
                        tension: 0.2,
                        spanGaps: true,
                        fill: true
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{ tooltip: tooltipStyle }},
                    scales: {{
                        y: {{
                            suggestedMin: 200,
                            suggestedMax: 1000,
                            grid: {{ color: 'rgba(51, 65, 85, 0.3)' }},
                            title: {{ display: true, text: 'Skor IRT (0 - 1000)', color: '#818cf8', font: {{ weight: '600' }} }}
                        }},
                        x: {{ grid: {{ color: 'rgba(51, 65, 85, 0.15)' }} }}
                    }}
                }}
            }});

            // Chart Biasa Mapel
            const ctxBiasa = document.getElementById('subjectTrendBiasaCanvas');
            if (subjectTrendBiasaChart) subjectTrendBiasaChart.destroy();
            subjectTrendBiasaChart = new Chart(ctxBiasa, {{
                type: 'line',
                data: {{
                    labels: tryoutLabels,
                    datasets: [{{
                        label: `Rata-rata Biasa (${{s.name}})`,
                        data: biasaTrend,
                        borderColor: '#10b981',
                        backgroundColor: 'rgba(16, 185, 129, 0.12)',
                        borderWidth: 3,
                        pointRadius: 6,
                        tension: 0.2,
                        spanGaps: true,
                        fill: true
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{ tooltip: tooltipStyle }},
                    scales: {{
                        y: {{
                            min: 0,
                            max: 100,
                            grid: {{ color: 'rgba(51, 65, 85, 0.3)' }},
                            title: {{ display: true, text: 'Nilai Biasa (0 - 100)', color: '#10b981', font: {{ weight: '600' }} }}
                        }},
                        x: {{ grid: {{ color: 'rgba(51, 65, 85, 0.15)' }} }}
                    }}
                }}
            }});
        }}

        // 5. ANALISIS DETAIL PER SISWA (HANYA MAPEL YANG DIAMBIL SISWA)
        function renderStudentDetail(nis) {{
            const st = studentDetails[nis];
            if (!st) return;

            const container = document.getElementById('studentDetailContent');
            const allTryouts = summaryData.tryouts;
            const tryoutLabels = allTryouts.map(n => `TO ${{n}}`);

            const rankIrtDisplay = st.rank_irt ? `#${{st.rank_irt}}` : '—';
            const rankBiasaDisplay = st.rank_biasa ? `#${{st.rank_biasa}}` : '—';

            // Siapkan Datasets per Mapel yang diambil untuk IRT
            const irtDatasets = [];
            st.subjects.forEach((sub, idx) => {{
                if (sub.has_irt) {{
                    const color = subjectPalette[idx % subjectPalette.length];
                    const pts = allTryouts.map(n => sub.scores_irt[String(n)] ?? null);
                    irtDatasets.push({{
                        label: `${{sub.name}} (${{sub.category}})`,
                        data: pts,
                        borderColor: color,
                        backgroundColor: color + '20',
                        borderWidth: 2.5,
                        pointRadius: 5,
                        pointHoverRadius: 8,
                        tension: 0.2,
                        spanGaps: true
                    }});
                }}
            }});

            // Tambahkan garis rata-rata IRT siswa
            const avgIrtPts = allTryouts.map(n => st.tryout_avgs_irt[String(n)] ?? null);
            irtDatasets.push({{
                label: '⭐ Rata-rata IRT Siswa',
                data: avgIrtPts,
                borderColor: '#ffffff',
                borderDash: [6, 4],
                borderWidth: 3,
                pointRadius: 6,
                pointBackgroundColor: '#ffffff',
                tension: 0.2,
                spanGaps: true,
                order: -1
            }});

            // Siapkan Datasets per Mapel yang diambil untuk Biasa
            const biasaDatasets = [];
            st.subjects.forEach((sub, idx) => {{
                if (sub.has_biasa) {{
                    const color = subjectPalette[idx % subjectPalette.length];
                    const pts = allTryouts.map(n => sub.scores_biasa[String(n)] ?? null);
                    biasaDatasets.push({{
                        label: `${{sub.name}} (${{sub.category}})`,
                        data: pts,
                        borderColor: color,
                        backgroundColor: color + '20',
                        borderWidth: 2.5,
                        pointRadius: 5,
                        pointHoverRadius: 8,
                        tension: 0.2,
                        spanGaps: true
                    }});
                }}
            }});

            // Tambahkan garis rata-rata Biasa siswa
            if (st.count_biasa > 0) {{
                const avgBiasaPts = allTryouts.map(n => st.tryout_avgs_biasa[String(n)] ?? null);
                biasaDatasets.push({{
                    label: '⭐ Rata-rata Biasa Siswa',
                    data: avgBiasaPts,
                    borderColor: '#ffffff',
                    borderDash: [6, 4],
                    borderWidth: 3,
                    pointRadius: 6,
                    pointBackgroundColor: '#ffffff',
                    tension: 0.2,
                    spanGaps: true,
                    order: -1
                }});
            }}

            container.innerHTML = `
                <div class="profile-banner">
                    <div class="profile-info">
                        <h3>${{st.nama}}</h3>
                        <p style="color: var(--text-secondary); font-size: 13.5px;">Nomor Induk Siswa (NIS): <strong>${{st.nis}}</strong> &bull; Mengambil <strong>${{st.subjects.length}} Mata Pelajaran</strong></p>
                        <div class="profile-badges">
                            <span class="rank-pill text-irt">🏆 Peringkat IRT: ${{rankIrtDisplay}} dari ${{summaryData.total_students}}</span>
                            <span class="rank-pill text-biasa">🎯 Peringkat Biasa: ${{rankBiasaDisplay}} dari ${{summaryData.total_students}}</span>
                        </div>
                    </div>
                    <div style="display: flex; gap: 16px;">
                        <div style="text-align: right;">
                            <div style="font-size: 12px; color: var(--text-muted); text-transform: uppercase;">Rata-rata IRT</div>
                            <div style="font-size: 26px; font-weight: 800;" class="text-irt">${{st.avg_irt ? st.avg_irt.toFixed(2) : '—'}}</div>
                        </div>
                        <div style="text-align: right;">
                            <div style="font-size: 12px; color: var(--text-muted); text-transform: uppercase;">Rata-rata Biasa</div>
                            <div style="font-size: 26px; font-weight: 800;" class="text-biasa">${{st.avg_biasa ? st.avg_biasa.toFixed(2) : '—'}}</div>
                        </div>
                    </div>
                </div>

                <!-- Dual Charts Siswa: IRT per Mapel vs Biasa per Mapel -->
                <div class="grid-2col" style="margin-bottom: 24px;">
                    <!-- Grafik 1: IRT per Mapel -->
                    <div class="chart-box" style="border-top: 3px solid var(--accent-irt);">
                        <div class="chart-box-header">
                            <div class="chart-box-title">
                                <span>🔮 Perkembangan Nilai IRT per Mapel</span>
                                <span class="badge badge-irt">Skala 0 - 1000</span>
                            </div>
                        </div>
                        <div class="chart-wrapper">
                            <canvas id="studentTrendIrtCanvas"></canvas>
                        </div>
                    </div>

                    <!-- Grafik 2: Biasa per Mapel -->
                    <div class="chart-box" style="border-top: 3px solid var(--accent-biasa);">
                        <div class="chart-box-header">
                            <div class="chart-box-title">
                                <span>🟢 Perkembangan Nilai Biasa per Mapel</span>
                                <span class="badge badge-biasa">Skala 0 - 100</span>
                            </div>
                        </div>
                        <div class="chart-wrapper" id="studentTrendBiasaWrapper">
                            <canvas id="studentTrendBiasaCanvas"></canvas>
                        </div>
                    </div>
                </div>

                <!-- Tabel Rincian Mapel yang Diambil Siswa -->
                <div class="table-card">
                    <div style="padding: 14px 20px; font-weight: 700; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center;">
                        <span>📋 Rincian Mata Pelajaran yang Dikerjakan (${{st.subjects.length}} Mapel)</span>
                        <span style="font-size: 12px; color: var(--text-muted); font-weight: 400;">Mata pelajaran pilihan yang tidak diambil otomatis disembunyikan</span>
                    </div>
                    <div class="table-responsive">
                        <table>
                            <thead>
                                <tr>
                                    <th style="width: 40px;">#</th>
                                    <th>Mata Pelajaran yang Diambil</th>
                                    <th>Kategori</th>
                                    ${{allTryouts.map(n => `<th style="text-align: right;">TO ${{n}} (IRT)</th>`).join('')}}
                                    <th style="text-align: right;" class="text-irt">Rata IRT</th>
                                    ${{summaryData.has_biasa ? allTryouts.map(n => `<th style="text-align: right;">TO ${{n}} (Biasa)</th>`).join('') : ''}}
                                    <th style="text-align: right;" class="text-biasa">Rata Biasa</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${{st.subjects.map((sub, idx) => `
                                    <tr>
                                        <td>${{idx + 1}}</td>
                                        <td><strong>${{sub.name}}</strong></td>
                                        <td><span class="badge ${{sub.category === 'Wajib' ? 'badge-wajib' : 'badge-pilihan'}}">${{sub.category}}</span></td>
                                        ${{allTryouts.map(n => `<td style="text-align: right;">${{sub.scores_irt[String(n)] !== undefined ? sub.scores_irt[String(n)] : '—'}}</td>`).join('')}}
                                        <td style="text-align: right;" class="text-irt">${{sub.avg_irt !== null ? sub.avg_irt.toFixed(2) : '—'}}</td>
                                        ${{summaryData.has_biasa ? allTryouts.map(n => `<td style="text-align: right;">${{sub.scores_biasa[String(n)] !== undefined ? sub.scores_biasa[String(n)] : '—'}}</td>`).join('') : ''}}
                                        <td style="text-align: right;" class="text-biasa">${{sub.avg_biasa !== null ? sub.avg_biasa.toFixed(2) : '—'}}</td>
                                    </tr>
                                `).join('')}}
                            </tbody>
                        </table>
                    </div>
                </div>
            `;

            // Chart IRT Siswa
            const ctxIrt = document.getElementById('studentTrendIrtCanvas');
            if (studentTrendIrtChart) studentTrendIrtChart.destroy();
            studentTrendIrtChart = new Chart(ctxIrt, {{
                type: 'line',
                data: {{ labels: tryoutLabels, datasets: irtDatasets }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    interaction: {{ mode: 'index', intersect: false }},
                    plugins: {{
                        tooltip: tooltipStyle,
                        legend: {{ labels: {{ color: '#cbd5e1', usePointStyle: true, font: {{ size: 11 }} }} }}
                    }},
                    scales: {{
                        y: {{
                            suggestedMin: 200,
                            suggestedMax: 1000,
                            grid: {{ color: 'rgba(51, 65, 85, 0.3)' }},
                            title: {{ display: true, text: 'Skor IRT (0 - 1000)', color: '#818cf8', font: {{ weight: '600' }} }}
                        }},
                        x: {{ grid: {{ color: 'rgba(51, 65, 85, 0.15)' }} }}
                    }}
                }}
            }});

            // Chart Biasa Siswa
            const ctxBiasa = document.getElementById('studentTrendBiasaCanvas');
            if (studentTrendBiasaChart) studentTrendBiasaChart.destroy();

            if (biasaDatasets.length === 0) {{
                document.getElementById('studentTrendBiasaWrapper').innerHTML = `
                    <div class="empty-notice">
                        Belum ada data nilai Biasa untuk siswa ini.<br>
                        (Data nilai Biasa akan otomatis muncul saat file 'BIASA ...' dimasukkan ke folder data_nilai)
                    </div>
                `;
            }} else {{
                studentTrendBiasaChart = new Chart(ctxBiasa, {{
                    type: 'line',
                    data: {{ labels: tryoutLabels, datasets: biasaDatasets }},
                    options: {{
                        responsive: true,
                        maintainAspectRatio: false,
                        interaction: {{ mode: 'index', intersect: false }},
                        plugins: {{
                            tooltip: tooltipStyle,
                            legend: {{ labels: {{ color: '#cbd5e1', usePointStyle: true, font: {{ size: 11 }} }} }}
                        }},
                        scales: {{
                            y: {{
                                min: 0,
                                max: 100,
                                grid: {{ color: 'rgba(51, 65, 85, 0.3)' }},
                                title: {{ display: true, text: 'Nilai Biasa (0 - 100)', color: '#10b981', font: {{ weight: '600' }} }}
                            }},
                            x: {{ grid: {{ color: 'rgba(51, 65, 85, 0.15)' }} }}
                        }}
                    }}
                }});
            }}
        }}

        // Helper navigasi klik
        function selectStudent(nis) {{
            const select = document.getElementById('studentSelect');
            select.value = nis;
            renderStudentDetail(nis);
            document.getElementById('student-detail').scrollIntoView({{ behavior: 'smooth' }});
        }}

        function selectSubject(mapelName) {{
            const select = document.getElementById('subjectSelect');
            select.value = mapelName;
            renderSubjectDetail(mapelName);
            document.getElementById('subject-detail').scrollIntoView({{ behavior: 'smooth' }});
        }}

        function filterStudentDropdown(query) {{
            const q = query.toLowerCase().trim();
            const select = document.getElementById('studentSelect');
            const options = select.options;
            let firstMatch = null;

            for (let i = 0; i < options.length; i++) {{
                const opt = options[i];
                const matches = opt.text.toLowerCase().includes(q);
                opt.style.display = matches ? '' : 'none';
                if (matches && !firstMatch) firstMatch = opt.value;
            }}

            if (firstMatch) {{
                select.value = firstMatch;
                renderStudentDetail(firstMatch);
            }}
        }}
    </script>
</body>
</html>
"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"\n[SELESAI] File laporan berhasil dibuat di: {output_path}")


# ---------------------------------------------------------------------------
# MAIN EXECUTION
# ---------------------------------------------------------------------------
def main():
    data_dir = sys.argv[1] if len(sys.argv) > 1 else "data_nilai"
    output_html = sys.argv[2] if len(sys.argv) > 2 else "index.html"

    print("=" * 70)
    print("🚀 GENERATOR LAPORAN TRYOUT SISWA (GRAFIK TERPISAH IRT & BIASA)")
    print("=" * 70)
    print(f"Folder Sumber Data : {data_dir}")
    print(f"Output File HTML   : {output_html}\n")

    try:
        df_scores = load_all_scores(data_dir=data_dir)
        analytics_data = compute_dashboard_data(df_scores)
        generate_html_report(analytics_data, output_path=output_html)
        print("\n✨ Pembuatan laporan selesai dengan sukses!")
    except Exception as e:
        print(f"\n❌ [ERROR] Terjadi kesalahan saat memproses data: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
