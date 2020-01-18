import pandas as pd
import time
from time import gmtime, strftime
from xlsxwriter.utility import xl_rowcol_to_cell
from IPython.display import display, HTML


def tmpxls(df,fn, dpath='C:/!data/tmp/', cs=[10]):
    tm(s=1)
    
    path = dpath + fn+'.xlsx'
    print ("save to :", path)
    
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    df.to_excel(writer, sheet_name = 'sheet', index=False,  startcol=0, startrow=1)

    tm('0. df.to_excel')
    
    workbook  = writer.book
    worksheet = writer.sheets['sheet']
    
    format1 = workbook.add_format({'color':'black','font_size':10, 'text_wrap':True, 'valign':'top'})
    f_link = workbook.add_format({'color':'blue','font_size':9, 'text_wrap':True, 'valign':'top'})
    format_inn = workbook.add_format({'color':'black','font_size':10, 'bold':True, 'bg_color':'#eef9ff', 'valign':'top'})
    total_fmt = workbook.add_format({'color':'black','font_size':20, 'bold':True, 'bg_color':'#ffeef9', 'align':'center'})

    header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'vcenter',
    'align': 'center',
    'fg_color': '#D7E4BC',
    'border': 1})

    # Write the column headers with the defined format.
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(1, col_num, value, header_format)
    
    cnt,inn_idx = 0,0 
    
    dsz = df.sample(1000 if df.shape[0] > 1000 else df.shape[0])
    
    link_col = -1
    
    for k in dsz.keys():
        
#         print (k, len(dsz[k].dropna()), end = ' | ')
        
        width = round( dsz[k].dropna().astype(str).str.len().mean())+1  if len(dsz[k].dropna()) > 0 else 10  
        width = width if width < 70 else 70

#         print (width)
        
        inn_list = ['inn', 'ИНН','налогоплательщика']
        link_list = ['link']
        
        if any(sub in dsz.keys()[cnt] for sub in inn_list):
            start_range = xl_rowcol_to_cell(2, cnt)
            end_range = xl_rowcol_to_cell(df.shape[0]+1, cnt)

            # Construct and write the formula
            formula = "=SUBTOTAL(3,{:s}:{:s})".format(start_range, end_range)
            f = format_inn
            
            # Set the column width and format.
            worksheet.set_column(cnt,cnt, width , f )   
            worksheet.write_formula(0, cnt, formula, total_fmt)
        elif any(sub in dsz.keys()[cnt] for sub in link_list):
            worksheet.set_column(cnt,cnt, 20 , f_link )   # Set the column width and format.
#             link_col = cnt
        else:
            f = format1 
            worksheet.set_column(cnt,cnt, width , f )   # Set the column width and format.
        cnt += 1
    
#     r = 0
#     if link_col != -1:
#         for index, row in df.iterrows():
#             formula = df.loc[index]['link']
#             if type(formula) is str:
#                 1
# #                 print (index,'|',cnt,"|",link_col,"|",formula)
# #                 worksheet.write_formula(r, link_col, formula, total_fmt)
#             r += 1

    worksheet.freeze_panes(2, 4)
    worksheet.autofilter(1,0,df.shape[0]+1, df.shape[1])
    worksheet.set_zoom(90)
    
    writer.save()
    writer.close()    

pd.DataFrame.tmpxls = tmpxls  # расширяем класс dataframe        
    
    
global tm 

def tm(txt="", s=0):
    global l, start
    
    if s: 
        l = start = time.time()
        print("","*"*80,"\n * ",strftime("%Y-%m-%d %H:%M:%S ", time.localtime()), str(txt),"\n","*"*80); 
    else: 
          print("\n[% 3.3f] " % (time.time()-l),"[% 3.0f] " % (time.time()-start),
          strftime("%Y-%m-%d %H:%M:%S ", time.localtime())," ",str(txt),"\n","*"*80); l = time.time()  
tm(s=1)



global ii

# dfr['1'] = ['22',23,5]

def ii(dii, k = 0, pref=''):    
    
#     global dfr
#     dii = dfr
    
    dfo = pd.DataFrame(columns=['name','rows','cols','memory','keys'])
    
    styler = dfo.style
    styler = styler.format("{:,.0f}")
    styler = styler.set_properties(**{'width': '100px', 'text-align': 'left'})
    styler.set_table_styles(
    [dict(selector="td", props=[("text-align", "left")])]   
    )
#     print(memus(dfr, deep = 0))
    for i in dii.keys(): 
#         print(type(dii[i]),i)
        if type(dii[i]) is pd.DataFrame:
#             if k == 1: print(dii[i].keys())
#             print ("[%s]:\t\t"%i , dii[i].shape, "{0:,}".format( dii[i].memory_usage(deep=0).sum()))
            dfo = dfo.append({'name': "dfr['" +pref+''+i+"']", 
                    'rows':  dii[i].shape[0], 
                    'cols':  dii[i].shape[1], 
                    'memory': "{0:,}".format( dii[i].memory_usage(deep=0).sum()),
                    'keys': "['" + "','".join( dii[i].keys().astype('str') )+"']" if k == 1 else str(len(dii[i].keys()))
                    }, ignore_index=True)
            
        elif isinstance(dii[i], dict):
#             print (i)
            dfo = dfo.append(ii(dii=dii[i], pref= str(i)+"']['", k = k))
    
    def text_left(s):  return ['text-align: left; width:210px' for v in s]
    def col_width(s):  return ['text-align: left; width:600px' for v in s]

    def float_format(s): return ["{0:,}".format(s) for v in s]
    
    def highlight_max(s): # highlight the maximum in a Series yellow.
        os = s
#         s = pd.to_numeric(s.str.replace(',', ''))
#         print(os, "|", s)
        is_max = s == s.max()       
        return ['background-color: #fff0cc; width:75px' if v else '' for v in is_max]

        '''     
        try: 
                float(s)
                s = float(s)
            except ValueError:
                try: 
                    int(s)
                    s = int(s)
                except ValueError:
                    s = len (s)
                    print (s)
        
''' 

    
#     dfo['memory'] = dfo['memory'].apply(lambda x: "{0:,}".format(x))

    dfo_html = dfo.style \
    .apply(text_left, subset=['name']) \
    .apply(col_width, subset=['keys']) \
    .apply(highlight_max)
    
    html = (dfo_html)
#                      )
# , subset=['B', 'C', 'D']
    display(html)
#     return dfo

tm(">>>> adds init complete!")