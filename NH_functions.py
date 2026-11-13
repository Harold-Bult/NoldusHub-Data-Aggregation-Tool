# -*- coding: utf-8 -*-
"""
Created on Wed Apr 30 12:29:49 2025

@author: HAR@Noldus I.T.

--- FIXES applied (FBA + GBA, 2026) ---
1. _parse_iso_flexible: robust timestamp parser. Handles the 'T' separator
   (ISO 8601), 6- or 7-digit second fractions (.NET format) and timezone.
   Used both on header fields (starttime, ...) and in get_time_from_line
   (timestamps of the data rows). The old rigid strptime failed on both
   formats.
2. get_xls_prospect: robust to None/'-'/empty durations.
3. get_exp_contents: uses the FULL path from os.walk (x[0]); a folder is a
   session only if it contains at least one signal CSV. Also returns the
   {session_name: full_path} mapping.
4. show_sel_session_export / aggregate_data: receive the full paths and have
   a guard if the session header is empty. Timestamps normalized to naive.
5. NH_resample: (a) an all-NaN column -> filled with NaN and skipped;
   (b) y = to_numpy(copy=True) to avoid the "assignment destination is
   read-only" crash (recent pandas returns a non-writable view).
"""

import numpy as np
import pandas as pd
from scipy import signal

import datetime
import os, configparser, time, base64, pickle, re
import tkinter as tk
from tkinter import filedialog


DATE_COLUMN = 'date/time'
SELECTIONS  = './NH_appselections.cfg'
date_format = '%Y-%m-%d %H:%M:%S.%f'
time_format = '%H:%M:%S.%f'
fn_export   = './aggregate.xlsx'
p_download  = './downloads/'


# ------------------------------------------------------------------ FIX 1
def _parse_iso_flexible(s):
    """robust datetime.fromisoformat: normalizes the second fraction to at
    most 6 digits, so it accepts both the ISO 8601 'T' separator and the
    7-digit .NET timestamps (e.g. 2024-05-06T09:17:00.4030893+02:00) that
    fromisoformat would otherwise reject."""
    if not isinstance(s, str):
        return s
    s2 = re.sub(r'\.(\d+)', lambda m: '.' + m.group(1)[:6], s)
    return datetime.datetime.fromisoformat(s2)
# ------------------------------------------------------------------------


def read_mem():
    d = dict()
    if os.path.isfile(SELECTIONS):
        with open(SELECTIONS, 'rb') as f:
            d = pickle.loads(f.read())
    return d


def write_mem(d):
    success = False
    with open(SELECTIONS, 'wb') as f:
        pickle.dump(d, f)
        success = True
    return success


def get_cfg(fname):
    getOK = False
    if os.path.isfile(fname):
        getOK = True
        config = configparser.ConfigParser()
        config.read(fname)
    else:
        config['configuration']['ftpath'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), '\\')
        getOK = False
    return config, getOK


def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()


def set_png_as_page_bg(png_file):
    bin_str = get_base64_of_bin_file(png_file)
    page_bg_img = '''
    <style>
    body {
    background-image: url("data:image/png;base64,%s");
    background-size: cover;
    }
    </style>
    ''' % bin_str
    return page_bg_img


def read_line_from_file(file_name, line_number):
    with open(file_name, 'r') as f:
        for line_no, line in enumerate(f):
            if line_no == line_number:
                f.close()
                return line
        f.close()
        raise ValueError('line %s does not exist in file %s' % (line_number, file_name))


# ------------------------------------------------------------------ FIX 1 (data rows)
def get_time_from_line(line, fcfg, sname):
    """Returns the Datetime of the data row. Uses the flexible parser:
    the old rigid-format strptime fails if the timestamp is ISO 8601
    with 'T' instead of a space, or has 7-digit fractions."""
    datum_col_nr = fcfg['datatypes'].split(fcfg['colsep'].strip('"')).index('datetime')
    raw = line.split(fcfg['colsep'].strip('"'))[datum_col_nr].strip()
    t = _parse_iso_flexible(raw)
    if getattr(t, 'tzinfo', None) is not None:
        t = t.replace(tzinfo=None)
    return t
# ------------------------------------------------------------------------


def get_physdata(PATH, datatype, cfg, includedata):
    if os.path.isfile(PATH+datatype+'.csv'):
        signals = [cfg[str(x)]["name"] for x in [*cfg._sections.keys()] if x.startswith("signal_")]
        signals_i = [x for x in [*cfg._sections.keys()] if x.startswith("signal_")]
        if signals.index(datatype) > -1:
            sname = signals_i[signals.index(datatype)]
            hdrline = int(cfg._sections[sname]['headerline'])

            fcfg = dict()
            fcfg['headerline'] = cfg._sections[sname]['headerline']
            fcfg['name'] = cfg._sections[sname]['name']
            with open(PATH+datatype+'.csv') as input_file:
                for nlines, line in enumerate(input_file):
                    pass
            with open(PATH+datatype+'.csv') as input_file:
                hdr = [next(input_file).strip(';\n') for _ in range(hdrline-1)]
                hdrs = next(input_file).strip(';\n').split(';')
                fcfg['headers'] = hdrs
                fcfg['nlines'] = nlines

                hdr = [x for x in hdr if not x.startswith("# ")]
                k = [*cfg._sections[sname].keys()]
                v = [*cfg._sections[sname].values()]
                for i in range(len(hdr)):
                    keyinhdr = hdr[i].split(';', 1)[0]
                    j = v.index(keyinhdr)
                    fcfg[k[j]] = hdr[i].split(';', 1)[1].split('\n')[0]

                if 'sessionduration' in fcfg:
                    t = datetime.datetime.strptime(fcfg['sessionduration'][:15], time_format)
                    fcfg['sessionduration'] = (t-datetime.datetime(1900, 1, 1, 0, 0, 0, 0)).total_seconds()
                # ---- FIX 1: robust parsing of header timestamps ----
                if 'starttime' in fcfg:
                    fcfg['starttime'] = _parse_iso_flexible(fcfg['starttime'])
                if 'videostarttimescreen' in fcfg:
                    fcfg['videostarttimescreen'] = _parse_iso_flexible(fcfg['videostarttimescreen'])
                if 'videostarttimewebcam' in fcfg:
                    fcfg['videostarttimewebcam'] = _parse_iso_flexible(fcfg['videostarttimewebcam'])
                # --------------------------------------------------------

                if nlines > hdrline:
                    line1st = read_line_from_file(PATH+datatype+'.csv', hdrline)
                    linelast = read_line_from_file(PATH+datatype+'.csv', nlines)
                    t_start = get_time_from_line(line1st, fcfg, sname)
                    t_end = get_time_from_line(linelast, fcfg, sname)
                    fcfg['signalduration'] = (t_end - t_start).total_seconds()
                    fcfg['sampfreq'] = (nlines-1-hdrline)/(t_end - t_start).total_seconds()
                    fcfg['signalstart'] = t_start
                else:
                    fcfg['signalduration'] = 0.0
                    fcfg['sampfreq'] = 0.0
                    fcfg['signalstart'] = None

                # fallback: if sessionduration/starttime is missing from the header
                if 'sessionduration' not in fcfg:
                    fcfg['sessionduration'] = fcfg.get('signalduration', 0.0)
                if 'starttime' not in fcfg:
                    fcfg['starttime'] = fcfg.get('signalstart')

                dtypes = fcfg['datatypes'].split(fcfg['colsep'].replace('"', ''))
                fcfg['datatypes'] = dtypes
                parse_dates = [hdrs[i].strip('\n') for i, value in enumerate(dtypes) if value == 'datetime']

            if includedata:
                physdata = pd.read_csv(PATH+datatype+'.csv',
                    header=hdrline-1,
                    usecols=hdrs,
                    parse_dates=parse_dates,
                    delimiter=fcfg['colsep'].replace('"', ''),
                    na_values=fcfg['missingvalue'].split(fcfg['colsep'].replace('"', ''))
                )
                physdata = physdata.rename(columns={parse_dates[0]: 'DateTime'})
            else:
                physdata = pd.DataFrame()
        else:
            physdata = pd.DataFrame()
            fcfg = dict()
    else:
        physdata = pd.DataFrame()
        fcfg = dict()
        print('File not found: '+PATH+datatype+'.csv')
    return physdata, fcfg


def get_folder():
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 10)
    f_path = filedialog.askdirectory(parent=root, title="Select Directory")
    root.destroy()
    return f_path


def get_rois(cfg):
    roi_names = [str(x) for x in [*cfg._sections.keys()] if x.startswith("ROI_")]
    roivals = [cfg._sections[x] for x in [*cfg._sections.keys()] if x.startswith("ROI_")]
    return roi_names, roivals


# ------------------------------------------------------------------ FIX 3
def get_exp_contents(rootfolder, cfg, prog_bar):
    """Returns: datframe, exp_name, datafound, signals_subs, session_paths.
    Uses the FULL path from os.walk; a folder is a session only if it
    contains at least one signal CSV."""
    signals = [cfg[str(x)]["name"] for x in [*cfg._sections.keys()] if x.startswith("signal_")]
    columns = [cfg[str(x)]["name"] for x in [*cfg._sections.keys()] if x.startswith("signal_")]
    columns.insert(0, 'duration (sec)')

    all_dirs = [x[0] for x in os.walk(rootfolder)]

    session_dirs = []
    for d in all_dirs:
        try:
            files = os.listdir(d)
        except OSError:
            continue
        if any((s + '.csv') in files for s in signals):
            session_dirs.append(d)

    session_names = [os.path.split(d)[1] for d in session_dirs]
    session_paths = dict(zip(session_names, session_dirs))

    datframe = pd.DataFrame(index=session_names, columns=columns, dtype='string')
    datafound = False
    signals_subs = dict()

    for i, sess_path in enumerate(session_dirs):
        df_row = [None]*len(columns)
        prog_bar.progress((i+1)/max(len(session_dirs), 1), 'preparing inventory...')
        for j in range(0, len(signals)):
            dummy, fcfg = get_physdata(sess_path+'/', signals[j], cfg, False)
            if (signals[j] in list(signals_subs)) == False:
                signals_subs[signals[j]] = []
            if len(fcfg) > 0:
                signal_duration = fcfg['signalduration']
                session_duration = fcfg['sessionduration']
                df_row[0] = str(round(session_duration, 1))
                if session_duration and session_duration > 0:
                    df_row[j+1] = str(round(100 * signal_duration/session_duration, 1))+' %'
                else:
                    df_row[j+1] = '-'
                for k, x in enumerate(fcfg['datatypes']):
                    if (x != 'datetime') and not(fcfg['headers'][k] in signals_subs[signals[j]]):
                        signals_subs[signals[j]] = signals_subs[signals[j]] + [fcfg['headers'][k]]
                datafound = True
            else:
                df_row[j+1] = '-'
        datframe.iloc[i, :] = df_row

    exp_name = os.path.split(rootfolder)[1]
    return datframe, exp_name, datafound, signals_subs, session_paths
# ------------------------------------------------------------------------


# ------------------------------------------------------------------ FIX 2
def get_xls_prospect(df, signals_subs, samp_freq):
    durations = df['duration (sec)']
    def _to_float(a):
        try:
            return float(a)
        except (TypeError, ValueError):
            return 0.0
    nr_lines = sum([_to_float(a)*samp_freq for a in durations])
    idx_map = {i: key for i, key in enumerate(signals_subs)}
    nr_MB = 16 * nr_lines * sum([len(signals_subs[idx_map[i]]) for i in list(range(0, len(idx_map)))]) / (1024*1024)
    return nr_lines, nr_MB
# ------------------------------------------------------------------------


# ------------------------------------------------------------------ FIX 5
def NH_resample(physdata, t, xnew0):
    data_types = physdata.dtypes
    headers = list(physdata)
    dtypes = []
    colnames = []

    for i in range(len(data_types)):
        if (data_types[headers[i]] == 'string') | (data_types[headers[i]] == 'object'):
            dtypes.append('str')
            colnames.append(headers[i])
        if data_types[headers[i]] == 'float64':
            dtypes.append('float64')
            colnames.append(headers[i])

    ynew = pd.DataFrame({c: pd.Series(dtype=t) for c, t in dict(zip(colnames, dtypes)).items()})

    ndatcols = len(dtypes)
    if ndatcols > 0:
        for i in range(ndatcols):
            colname = colnames[i]

            if ((dtypes[i] == 'float64') & (colname != 'DT_0')):
                # FIX 5b: copy=True -> writable array (avoids "read-only" on y[nanmask]=...)
                y = physdata[colname].to_numpy(copy=True)

                # FIX 5a: all-NaN column -> fill with NaN and skip
                if np.all(np.isnan(y)):
                    ynew[colname] = np.full_like(xnew0, np.nan, dtype=float)
                    continue

                if np.isnan(np.sum(y)):
                    nanmask = np.isnan(y)
                    ypd = pd.Series(y)
                    a = np.vstack((ypd.ffill()[pd.Series(nanmask)], ypd.bfill()[pd.Series(nanmask)]))
                    padding = np.average(a.T.tolist(), axis=1)
                    del(ypd, a)
                else:
                    nanmask = np.zeros((len(y)), dtype=bool)
                    padding = []

                tmin = np.min(np.array(t)[~nanmask])
                tmax = np.max(np.array(t)[~nanmask])
                imin = np.where(nanmask == False)[0][0]
                imax = np.where(nanmask == False)[0][-1]

                xnew_segment = ((xnew0 >= tmin) & (xnew0 <= tmax))
                num = sum(xnew_segment)
                y[nanmask] = padding

                yn = signal.resample(y[imin:imax], num)
                ynew[colname] = np.full_like(xnew0, None)
                ynew.loc[xnew_segment, colname] = yn

            if dtypes[i] == 'str':
                inputstrings = physdata[colname]
                ynew[colname] = ['']*len(xnew0)
                for k in range(len(inputstrings)):
                    tmask = xnew0 >= t[k]
                    ynew[colname] = ynew[colname].mask(tmask, inputstrings[k])
    return ynew
# ------------------------------------------------------------------------


# ------------------------------------------------------------------ FIX 4
def aggregate_data(cfg, signals_subs, session_paths, samp_freq, xls_mode, AOI_src_selected, prog_bar):
    t0 = time.time()
    signals = list(signals_subs.keys())
    rois, roivals = get_rois(cfg)

    columns = ['DateTime', 'DT_0', 'Session']
    for a in signals:
        columns += signals_subs[a]
    for a in range(len(rois)):
        columns += [roivals[a]['name']+'_'+AOI_src_selected]

    try:
        os.makedirs(p_download, exist_ok=True)
        pd.DataFrame([]).to_excel(p_download+fn_export, engine="openpyxl")
    except Exception:
        print("Could not save XLSX. Close the target file in Excel and try again.")

    session_items = list(session_paths.items())

    for i, (sess_name, sess_path) in enumerate(session_items):
        t = time.time()
        prog_bar.progress((i+1)/max(len(session_items), 1),
                          text=sess_name + ' : ' + str(round(t-t0, 3)) + ' s')

        hdr_info = dict()
        for j in range(0, len(signals)):
            dummy, hdr_info = get_physdata(sess_path+'/', signals[j], cfg, False)
            if len(hdr_info) > 0:
                break
        if not hdr_info or not hdr_info.get('sessionduration'):
            print('No valid signal in ' + sess_path + ' - session skipped')
            continue

        x0 = np.arange(0.0, hdr_info['sessionduration'], 1/samp_freq, dtype=float)
        x = [pd.to_datetime(hdr_info['starttime'], format=date_format).tz_localize(None) + datetime.timedelta(seconds=ts) for ts in x0]

        df = pd.DataFrame(columns=columns)
        df['DateTime'] = x
        df['DT_0'] = x0
        df['Session'] = sess_name

        for j in range(0, len(signals)):
            physdata, hdr_info2 = get_physdata(sess_path+'/', signals[j], cfg, True)
            if len(physdata) > 0:
                start_ref = hdr_info2['starttime']
                if getattr(start_ref, 'tzinfo', None) is not None:
                    start_ref = start_ref.replace(tzinfo=None)
                dt_col = physdata['DateTime']
                if hasattr(dt_col, 'dt') and dt_col.dt.tz is not None:
                    dt_col = dt_col.dt.tz_localize(None)
                dt_relative = [xx.total_seconds() for xx in (dt_col - start_ref)]
                ynew = NH_resample(physdata, dt_relative, x0)
            else:
                ynew = pd.DataFrame(columns=signals_subs[signals[j]])
            for xx in ynew.keys():
                df[xx] = ynew[xx]

        for j in range(0, len(rois)):
            col_x = AOI_src_selected+'X'
            col_y = AOI_src_selected+'Y'
            if col_x in df.columns and col_y in df.columns:
                xx = df[col_x]
                yy = df[col_y]
                x1 = float(roivals[j]['x1']); x2 = float(roivals[j]['x2'])
                y1 = float(roivals[j]['y1']); y2 = float(roivals[j]['y2'])
                df[roivals[j]['name']+'_'+AOI_src_selected] = ((xx > x1) & (xx < x2)) & ((yy > y1) & (yy < y2))

        if xls_mode == "1 file per session":
            try:
                mode = 'a' if os.path.isfile(p_download+fn_export) else 'w'
                kwargs = dict(engine="openpyxl", mode=mode, datetime_format=date_format)
                if mode == 'a':
                    kwargs['if_sheet_exists'] = "replace"
                with pd.ExcelWriter(p_download+fn_export, **kwargs) as writer:
                    df.to_excel(writer, sheet_name=sess_name[:31])
            except Exception:
                print("Could not save XLSX. Close the target file in Excel and try again.")

        if xls_mode == "1 tabsheet per session":
            try:
                with pd.ExcelWriter(p_download+sess_name+'.xlsx', engine="openpyxl", mode='w', datetime_format=date_format) as writer:
                    df.to_excel(writer, sheet_name=sess_name[:31])
            except Exception:
                print("Could not save XLSX. Close the target file in Excel and try again.")

        if xls_mode == "all sessions in 1 tabsheet":
            try:
                mode = 'a' if os.path.isfile(p_download+fn_export) else 'w'
                kwargs = dict(engine="openpyxl", mode=mode, datetime_format=date_format)
                if mode == 'a':
                    kwargs['if_sheet_exists'] = "overlay"
                with pd.ExcelWriter(p_download+fn_export, **kwargs) as writer:
                    startrow = writer.sheets['all data'].max_row if (mode == 'a' and 'all data' in writer.sheets) else 0
                    df.to_excel(writer, sheet_name='all data', index=False,
                                header=(startrow == 0), startrow=startrow)
            except Exception:
                print("Could not save XLSX. Close the target file in Excel and try again.")

    return True
# ------------------------------------------------------------------------


def show_sel_session_export(sessionpath, signals_subs, cfg, samp_freq, prog_bar):
    """sessionpath: FULL path to the session folder."""
    hdr_info = dict()
    signals = list(signals_subs.keys())

    for j in range(0, len(signals)):
        dummy, hdr_info = get_physdata(sessionpath+'/', signals[j], cfg, False)
        if len(hdr_info) > 0:
            break

    if not hdr_info or not hdr_info.get('sessionduration'):
        columns = ['DateTime', 'DT_0']
        for x in signals:
            columns += signals_subs.get(x, [])
        return pd.DataFrame(columns=columns)

    columns = ['DateTime', 'DT_0']
    for x in signals:
        columns += signals_subs[x]
    resampSessData = pd.DataFrame(columns=columns)
    xnew0 = np.arange(0.0, hdr_info['sessionduration'], 1/samp_freq, dtype=float)
    resampSessData['DateTime'] = [pd.to_datetime(hdr_info['starttime'], format=date_format).tz_localize(None) + datetime.timedelta(seconds=ts) for ts in xnew0]
    resampSessData['DT_0'] = xnew0

    colstep = 0
    for j in range(0, len(signals)):
        physdata, hdr_info2 = get_physdata(sessionpath+'/', signals[j], cfg, True)
        if len(hdr_info2) > 0 and len(physdata) > 0:
            start_ref = hdr_info2['starttime']
            if getattr(start_ref, 'tzinfo', None) is not None:
                start_ref = start_ref.replace(tzinfo=None)
            dt_col = physdata['DateTime']
            if hasattr(dt_col, 'dt') and dt_col.dt.tz is not None:
                dt_col = dt_col.dt.tz_localize(None)
            dt_relative = [a.total_seconds() for a in (dt_col - start_ref)]
            ynew = NH_resample(physdata, dt_relative, xnew0)
            for x in ynew.keys():
                colstep += 1
                prog_bar.progress(min(colstep/max(len(columns), 1), 1.0), text="Resampling **"+signals[j]+" - "+x+"** data now. Please wait.")
                time.sleep(0.01)
                resampSessData[x] = ynew[x]
    return resampSessData


def get_aoi(aoi, eye_x, eye_y):
    aoi_res = pd.DataFrame(columns=[aoi['name']], dtype='bool')
    return aoi_res
