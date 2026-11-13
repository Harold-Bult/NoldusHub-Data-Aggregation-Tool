# -*- coding: utf-8 -*-
"""
Created: Fri Feb 21 16:31:22 2025
version 1.0.0: Thu Apr 24 18:45:32 2025

@author: HAR@Noldus.com
The UI is defined here; the functions used are defined in NH_functions.py

--- FIXES applied (FBA, 2026) ---
1. Dynamic PROGPATH (file directory); no longer hard-coded to C:\\Users\\har.
2. Sensible defaults for rootfolder / df / datfound when no state was saved.
3. Use of session_paths (full paths) for row selection and export.
"""

import os

# ---- FIX 1: PROGPATH = directory of this file (Windows and Linux) ----
PROGPATH = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
if PROGPATH:
    os.chdir(PROGPATH)
# --------------------------------------------------------------------

CONFIG_FILE = './NH_datsegmentation.cfg'
ICON_NOLDUS = './static/noldus.ico'
BACK_NOLDUS = './static/logo_noldus.png'
max_xls_lines = 1048576

import streamlit as st
import NH_functions


st.cache_data.clear()

page = st.set_page_config(
     page_title="NoldusHub Data Aggregation Tool",
     page_icon=ICON_NOLDUS,
     layout="wide",
     initial_sidebar_state="expanded",
     menu_items={'About': "Contact Noldus at support@noldus.com"}
)


CFG, get_OK = NH_functions.get_cfg(CONFIG_FILE)

d = NH_functions.read_mem()
if 'NH_cfg' in d:
    NH_cfg = d['NH_cfg']

if not 'NH_cfg' in locals():
    if get_OK:
        NH_cfg = CFG
        sucmsg = st.sidebar.success('Configuration loaded!', icon="✅")
        NH_functions.time.sleep(1)
        sucmsg.empty()

samp_freq = float(NH_cfg["general"]["sampfreq"])

if 'aoi_sources' in NH_cfg['general']:
    AOI_sources = NH_cfg['general']['aoi_sources'].split(';')
else:
    if 'aoi_sources' in CFG['general']:
        AOI_sources = CFG['general']['aoi_sources'].split(';')
    else:
        AOI_sources = ['No AOI sources defined in config']

if 'txt_sf' in d:
    txt_sf = d['txt_sf']
    samp_freq = float(txt_sf)

# ---- FIX 2: sensible defaults if no state was saved ----
rootfolder = d.get('rootfolder', '')
title      = d.get('title', 'titel')
df         = d.get('df', None)
datfound   = d.get('datfound', False)
signals_subs = d.get('signals_subs', {})
session_paths = d.get('session_paths', {})
nr_lines   = d.get('nr_lines', 0)
nr_MB      = d.get('nr_MB', 0)
xls_mode   = d.get('xls_mode', None)
AOI_src_selected = d.get('AOI_src_selected', AOI_sources[0])
if AOI_src_selected not in AOI_sources:
    AOI_src_selected = AOI_sources[0]
# ----------------------------------------------------------


st.sidebar.title("NoldusHub Data Aggregation Tool")
st.sidebar.write("")

######### CONTAINER 1 ##########
with st.sidebar.container(height=120, border=True):
    if st.button("Select data root folder", type="secondary", icon=None, use_container_width=False, key="select"):
        rootfolder = NH_functions.get_folder()
        prog_bar = st.progress(0)
        df, title, datfound, signals_subs, session_paths = NH_functions.get_exp_contents(rootfolder, NH_cfg, prog_bar)
        prog_bar.empty()
        if datfound:
            nr_lines, nr_MB = NH_functions.get_xls_prospect(df, signals_subs, samp_freq)
        else:
            nr_lines = 0
            nr_MB = 0
    st.write(rootfolder)

######### CONTAINER 2 ##########
with st.sidebar.container(height=400, border=True):
    txt_sf = st.text_input("**Sample Frequency output file:**", key="SF", value=samp_freq)
    try:
        samp_freq = float(txt_sf)
    except (TypeError, ValueError):
        pass

    st.write("")
    AOI_src_selected = st.selectbox("**Select AOI source**", AOI_sources, AOI_sources.index(AOI_src_selected))
    st.write("")

    xlsoptions = ["1 file per session", "1 tabsheet per session", "all sessions in 1 tabsheet"]
    try:
        index = xlsoptions.index(xls_mode)
    except (ValueError, TypeError):
        index = 0
    xls_mode = st.radio("**Make-up of .xlsx export file:**", xlsoptions, index)

    if datfound and df is not None:
        nr_lines, nr_MB = NH_functions.get_xls_prospect(df, signals_subs, samp_freq)
    if 'nr_MB' in locals():
        if nr_lines/max_xls_lines > 1:
            st.write('filesize = ' + str(round(nr_MB, 1))+' MB')
            if xls_mode == 'all sessions in 1 tabsheet':
                st.write('#-datalines =  :red[' + str(int(nr_lines))+' > max lines ('+str(int(max_xls_lines))+' per xlsx sheet!]')
        else:
            st.write('filesize = ' + str(round(nr_MB, 1))+' MB')
            if xls_mode == 'all sessions in 1 tabsheet':
                st.write('#-datalines = ' + str(int(nr_lines)))

######### CONTAINER 3 ##########
with st.sidebar.container(height=130, border=True):
    if datfound:
        if st.button("Create aggregate xlsx", type="secondary", icon=None, use_container_width=False, key="aggregate"):
            prog_bar = st.progress(0, text="wait...")
            NH_functions.aggregate_data(NH_cfg, signals_subs, session_paths, samp_freq, xls_mode, AOI_src_selected, prog_bar)
            prog_bar.empty()

######### CONTAINER 4 ##########
with st.sidebar.expander("Inspect configuration"):
    st.write(NH_cfg._sections.items(), unsafe_allow_html=True)


######## main navigation UI: ######
if df is not None:
    st.write('**Experiment: ' + title + '**')
    rowselected = st.dataframe(
        df,
        hide_index=False,
        on_select='rerun',
        selection_mode=['single-row'])

    if len(rowselected.selection.rows) > 0:
        row_i = rowselected.selection.rows[0]
        sess_name = df.index[row_i]
        # ---- FIX 3: use the FULL path from the session_paths map ----
        sess_path = session_paths.get(sess_name, sess_name)
        # ---------------------------------------------------------------
        signals = [NH_cfg[str(x)]["name"] for x in [*NH_cfg._sections.keys()] if x.startswith("signal_")]
        prog_bar = st.progress(0, text="wait...")
        resampSessData = NH_functions.show_sel_session_export(sess_path, signals_subs, NH_cfg, samp_freq, prog_bar)
        prog_bar.empty()
        st.data_editor(
            resampSessData,
            column_config={
                "DateTime": st.column_config.DatetimeColumn(
                    "DateTime",
                    format="YYYY-MM-DD HH:mm:ss.SS",
                ),
            },
            hide_index=True,
        )
        prog_bar.empty()


###############################
#   PERSISTENCY FILE UPDATE   #
###############################
d = dict()
d['txt_sf'] = txt_sf
d['NH_cfg'] = NH_cfg
d['rootfolder'] = rootfolder
d['title'] = title
if df is not None:
    d['df'] = df
d['datfound'] = datfound
d['signals_subs'] = signals_subs
d['session_paths'] = session_paths
d['nr_lines'] = nr_lines
d['nr_MB'] = nr_MB
d['AOI_src_selected'] = AOI_src_selected
d['xls_mode'] = xls_mode

NH_functions.write_mem(d)
