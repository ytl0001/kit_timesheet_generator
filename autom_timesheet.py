import datetime
import pandas as pd
import pdfrw

# = = = = = = = = = = = = = = = = = = =

def get_workdays(year, month):
    
    workdays = []

    # get first workday of first week
    d = datetime.datetime(year,month,1)

    if d.weekday() > 4: # first week has no workday
        w1_workday = '-'
    else:
        w1_workday = d
    workdays.append(w1_workday)

    # get monday (first workday) of second week
    w2_workday = datetime.datetime(year,month, 8 - d.weekday())
    workdays.append(w2_workday)

    for wk_idx in range(0,6):
        w_workday = w2_workday + datetime.timedelta(days=7*(wk_idx+1))
        if w_workday.month != w2_workday.month: break
        workdays.append(w_workday)
        
    return workdays

# = = = = = = = = = = = = = = = = = = =

def add_entry_to_dict(dict_in, idx, date, start, worktime_h, worktime_m, text='IOR'):
    
    date = date.strftime("%d.%m.%y")
    start_time = start.strftime("%H:%M")
    end_time = (start + datetime.timedelta(hours=int(worktime_h), minutes=int(worktime_m))).strftime("%H:%M")
    break_time = '00:00'
    work_time = f"{worktime_h:02d}:{worktime_m:02d}"

    fieldnames = [f"Tätigkeit Stichwort ProjektRow{idx}", f"ttmmjjRow{idx}", f"hhmmRow{idx}", f"hhmmRow{idx}_2", f"hhmmRow{idx}_3", f"hhmmRow{idx}_4"]
    values = [text, date, start_time, end_time, break_time, work_time]
    
    for idx, field in enumerate(fieldnames):
        dict_in.update({field: values[idx]})

# = = = = = = = = = = = = = = = = = = =

def generate_entries(df_fname, year, month, text, max_pses=4, max_pday=8, ses1_start=8, ses2_start=14):
    
    df = pd.read_csv(df_fname)
    worktime_arr = df.loc[:,['h', 'm']].to_numpy()

    work_h = max_pses # max per session
    work_h_perday = max_pday # max per day

    entries_dict = {}
    dict_idx = 1

    start1 = datetime.datetime(year, month, 1, hour=ses1_start)
    start2 = datetime.datetime(year, month, 1, hour=ses2_start)
    
    workdays = get_workdays(year, month)

    for wk_idx, week in enumerate(worktime_arr):

        if week[0] == 0 and week[1] == 0: 
            continue

        rest_hour, rest_min = week
        date_entry = workdays[wk_idx]

        for day_idx in range((rest_hour//work_h_perday)+1):

            if rest_hour < work_h and rest_min > 0: # ONLY first session if less than 4 hours
                add_entry_to_dict(entries_dict, dict_idx, date_entry, start1, rest_hour, week[1], text[wk_idx])
                dict_idx += 1
                rest_hour, rest_min = 0, 0

            elif rest_hour >= work_h: # split worktime intra day if more than 4 hours

                # first session 08:00 to 12:00
                add_entry_to_dict(entries_dict, dict_idx, date_entry, start1, work_h, 0, text[wk_idx])
                dict_idx += 1
                rest_hour -= work_h

                # second session 14:00 to REST
                if rest_hour <= work_h:
                    add_entry_to_dict(entries_dict, dict_idx, date_entry, start2, rest_hour, week[1], text[wk_idx])
                    dict_idx += 1
                    rest_hour, rest_min = 0, 0

                else: # more than 8 hours of work needs to be split over multiple days

                    # second session 14:00 to 18:00
                    add_entry_to_dict(entries_dict, dict_idx, date_entry, start2, work_h, 0, text[wk_idx])
                    dict_idx += 1
                    rest_hour -= work_h

            date_entry = workdays[wk_idx] + datetime.timedelta(days=day_idx+1)
            
    total_time_h = df["Hours"].sum()
    sum_h = int(total_time_h//1)
    sum_m = int(total_time_h%1 * 60)
    entries_dict.update({"Summe": f"{sum_h}h {sum_m}m"})
    
    return entries_dict

# = = = = = = = = = = = = = = = = = = =

def form_filler(in_path, data, out_path):
    
    pdf = pdfrw.PdfReader(in_path)
    
    for page in pdf.pages:
        annotations = page['/Annots']
        if annotations is None: continue
            
        for annotation in annotations:
            
            if annotation['/Subtype'] == '/Widget':
                key = annotation['/T'].to_unicode()
                
                if key in data:
                    pdfstr = pdfrw.objects.pdfstring.PdfString.encode(data[key])
                    annotation.update(pdfrw.PdfDict(V=pdfstr))
                    
        pdf.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
        pdfrw.PdfWriter().write(out_path, pdf)