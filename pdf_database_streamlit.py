import pandas as pd
import streamlit as st
#import pickle

prdDFlist = [['All Products','DF_All_Products_v6B.sav'], ['Architectural Windows', 'DF_Architectural_Windows_v4.sav'], ['Curtain Wall', 'DF_Curtain_Wall_v4.sav']\
             ,['Entrances','DF_Entrances_v4.sav'],['Storefront','DF_Storefront_v4.sav'],['Window Wall','DF_Window_Wall_v4.sav'],['Sun Controls', 'DF_Sun_Controls_v4.sav']]
prdCat = ['All Products', 'Architectural Windows','Canadian Products', 'Curtain Wall','Entrances','Storefront','Window Wall','Sun Controls']
prd = 'Curtain Wall'
prdNum = 2

bolList = []
dtlList = []

dfFile = 'DF_All_Products_4-25-2024.sav'

def convert_df(df):
    
    return df.to_csv().encode('utf-8')

def run_search():

    row = 1
    ct = -1

    #pdfDir = st.session_state.fldrPth
    fOut = st.session_state.fileOut
    
    prdCat = st.session_state.pc
    
    if st.session_state.exStr == True:
      txtStr = '\\b' + st.session_state.srchStr + '\\b'
    else:
      txtStr = st.session_state.srchStr  
    
    if st.session_state.hypStr == True:
      txtStr2 = txtStr.replace('-','')
      txtStr3 = txtStr.replace('-',' ')

    #for i in range(len(prdDFlist)):
    #    if prdCat in prdDFlist[i][0]:
    #        dfFile = prdDFlist[i][1]
    #        break'

    #cols = ['Product Name','Full Name','Secondary Heading','Filename','Detail Heading 1','Detail Heading 2','Page Number','Search String']
    #workbook = xlsxwriter.Workbook(fOut)
    #worksheet = workbook.add_worksheet("Details Parts List") 
    #worksheet.write_row(0,0,cols)
    dfP = pd.read_pickle(dfFile)
    #dfP = pickle.load(open(dfFile,'rb'))
    if prdCat == 'All Products':
      dfPrd = dfP
    else:
      dfPrd = dfP.loc[(dfP['Product Category'] == prdCat)]
    #print(dfPrd)
    #print(prd)
    if prdCat == 'All Products':
      if st.session_state.hypStr == True:
        dfOut1 = dfPrd.loc[(dfPrd['Text'].str.contains('|'.join([txtStr, txtStr2, txtStr3]),case=False, na=False, regex=True))]
      else:
        dfOut1 = dfPrd.loc[(dfPrd['Text'].str.contains(txtStr,case=False, na=False, regex=True))]
    else:
        dfOut1 = dfPrd.loc[(dfPrd['Text'].str.contains(txtStr,case=False, na=False, regex=True)) & (dfPrd['Product Category'] == prdCat)]

    if st.session_state.actPrd == True:
      dfOut2 = dfOut1.loc[(dfPrd['Status'] == 'Active')]
    else:
      dfOut2 = dfOut1
    if not docList:
      dfOut = dfOut2
    else:
      dfOut = dfOut2[dfOut2['Document Type'].isin(docList)]

    dfOut['Count'] = dfOut['Count'].astype('int')
    print(list(dfOut.columns.values))
    cols1 = ['Product Category','Document Type','Status','Document Info #1','Document Info #2','Filename','Page Info #1','Page Info #2','Page #','Text','Filepath']
    cols2 = ['Product Name','Product Category','Document Type','Status','Document Info #1','Document Info #2','Page Info #1','Page Info #2','Page #','Text','Filepath']
    #cols = ['Product Category','Product Description','Filename','Page Title','Page Description','Page Description 2','Page Number']
    df3 = dfOut.groupby('Product Name').sum().drop(columns = cols1)
    df4 = dfOut.groupby('Filename').sum().drop(columns = cols2)
    
    #dfLine = pd.DataFrame(columns = cols)
    #dfOut.to_csv(fOut)
    
    csv = convert_df(dfOut)

    with st.container():
        #col1, col2 = st.columns(2)
        col1edge, col1, col2edge = st.columns((1, 9, 1))
        col1.header('Full Search Results')
        col1.download_button("Download Results", csv, st.session_state.fileOut,"text/csv", key='browser-data')
        col1.dataframe(dfOut)
        col1.header('Number of results by product')
        col1.dataframe(df3)
        col1.header('Number of results by file')
        col1.dataframe(df4)
# closing the pdf file object

def list_files():
  dfP = pd.read_pickle(dfFile)
  dfOut = dfP['Filename'].drop_duplicates()
  #dfOut = dfOut.sort_values
  with st.container():
    col1edge, col1, col2edge = st.columns((1, 12, 1))
    col1.dataframe(dfOut)
# filePath is a string that contains the path to the pdf
ct = 0

st.set_page_config(layout="wide")

with st.container():
    col1edge, col1, col2edge = st.columns((1, 12, 1))
    col1.title('Web Details Parts Search')
    col1.caption('')
    col1.caption('Contains web details, installation manuals, structural charts, and Florida Product Approvals')
    col1.caption('Last Updated: 1/9/24')
    col1.caption('')

with st.container():
    #col1, col2 = st.columns(2)
    col1edge, col1, col2, col3, col2edge = st.columns((1, 3, 3, 3, 1))
    col1.selectbox('Product Category', prdCat, index=0, key='pc', help=None, on_change=None, args=None, kwargs=None, disabled=False)
    col2.text_input("Search String", value = 'WW-110', key="srchStr")
    col3.text_input("Output File", value = 'pdf_search_results.csv', key="fileOut")
    col1.text('')
    col2.text('')
    col3.text('')
    
    col1.text('')
    col2.text('')
    col3.text('')
    col1.checkbox('Active products only',key='actPrd', value=True)
    col2.checkbox('Exact part only (e.g., GP-100 but not GP-1001)', key='exStr', value=True)
    col3.checkbox('Include results without hyphen or space', key='hypStr', value=False)
    
    #col1.button_calc = st.button(label='Calculate', key ='calc')
    with st.container():
    #col1, col2 = st.columns(2)
      col1edge, col1, col2edge = st.columns((1, 9, 1))
      docList = col1.multiselect('Types of documents to search:',['Web Details','Installation Manuals','Florida Product Approvals', 'Sell Sheet', 'Specifications', 'Structural Charts'],[])
      col1.text('')
      col1.button(label='Search', key ='calc')
      col1.text('')
      col1.button(label='List Included Files', key ='fileList')


if st.session_state.calc:
    run_search()

if st.session_state.fileList:
    list_files()
