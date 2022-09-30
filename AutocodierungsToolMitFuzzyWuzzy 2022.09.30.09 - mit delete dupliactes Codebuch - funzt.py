import streamlit as st
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import pandas as pd
import numpy as np

#rapidfuzz
from rapidfuzz import process, fuzz
from rapidfuzz.process import extractOne
from rapidfuzz.string_metric import levenshtein, normalized_levenshtein
from rapidfuzz.fuzz import ratio

#für Excel-Funktionen
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

#Für Abbildungen
import matplotlib.pyplot as plt


from st_aggrid import GridOptionsBuilder, AgGrid, GridUpdateMode, DataReturnMode #für editierbare Tabellen


st.set_page_config(layout='wide', page_title='Fuzzy Autocodierung' )

#Code um den Button-Design anzupassen
m = st.markdown("""
<style>
div.stButton > button:first-child {
    background-color: #ce1126;
    color: white;
    height: 3em;
    width: 14em;
    border-radius:10px;
    border:3px solid #000000;
    font-size:18px;
    font-weight: bold;
    margin: auto;
    display: block;
}

div.stButton > button:hover {
	background:linear-gradient(to bottom, #ce1126 5%, #ff5a5a 100%);
	background-color:#ce1126;
}

div.stButton > button:active {
	position:relative;
	top:3px;
}
</style>""", unsafe_allow_html=True)

















#Variablen
final_result = pd.DataFrame()

codebuchKategorie = []
similarity = []
Code = []

anzahlCodierteZeilen = 0


# ======== "Hier Versuche mit eigenem Datenimport, funzt !!!" =============================================================================#

dfAntworten = pd.DataFrame()
dfCodebuch = pd.DataFrame()
dfExcelExport = pd.DataFrame()

dataupload1 = 0
dataupload2 = 0


st.header("Autocodierung von offenen Antworten mit Fuzzy Wuzzy")

st.sidebar.subheader("Codeliste:")
uploaded_file2 = st.sidebar.file_uploader("Upload Excel-File mit Codeschema",type=["xlsx","xls", "xlsm"])
st.sidebar.info("Die Spalte A in Excel mit der Codeliste muss 'Name' heissen und die B-Spalte 'Codes' ")

if uploaded_file2:
    dfCodebuchImport = pd.read_excel(uploaded_file2, dtype={'Name': 'str'}) #, index_col=0
    if (dfCodebuchImport.columns[0]) != "Name":
        st.warning("Erste Spalte in der Codeliste muss Namen heissen")
    if (dfCodebuchImport.columns[1]) != "Codes":
        st.warning("Zweite Spalte in der Codeliste muss Codes heissen")

    anzahlZeilenMitDuplikate = len(dfCodebuchImport)
    #duplikate entfernen
    dfCodebuchImport = dfCodebuchImport.drop_duplicates(subset=['Name'], keep="first")
    anzahlZeilenOhneDuplikate = len(dfCodebuchImport)
    if anzahlZeilenMitDuplikate != anzahlZeilenOhneDuplikate:
        st.info("Es wurden Duplikate in der Spalte Name im Codebuch  gefunden und entfernt")

	#st.write(dfCodebuch)
    #d2 = df2import.to_dict() #orient="index"
    dataupload2 = 2
    st.write("Codebuch als editierbares Dataframe (df):")
    grid_return = AgGrid(dfCodebuchImport, editable=True,theme="streamlit", key="HalloThomas")
    dfCodebuch = grid_return['data']

    #Beispiel von https://medium.com/analytics-vidhya/matching-messy-pandas-columns-with-fuzzywuzzy-4adda6c7994f
    codeBuchExpander = st.expander('Codebuch runterladen')
    with codeBuchExpander:
                    speicherZeitpunkt = pd.to_datetime('today')
                    st.write("")
                    if len(dfCodebuch) > 0:					
                        def to_excel(dfCodebuch):
                            output = BytesIO()
                            writer = pd.ExcelWriter(output, engine='xlsxwriter')
                            dfCodebuch.to_excel(writer, index=False, sheet_name='Sheet1')
                            workbook = writer.book
                            worksheet = writer.sheets['Sheet1']
                            format1 = workbook.add_format({'num_format': '0.00'}) 
                            worksheet.set_column('A:A', None, format1)  
                            writer.save()
                            processed_data = output.getvalue()
                            return processed_data
                        df_xlsx = to_excel(dfCodebuch)
                        st.download_button(label='📥 Tabelle in Excel abspeichern?',
                            data=df_xlsx ,
                            file_name= 'Codeliste '+str(speicherZeitpunkt) +'.xlsx' )










st.sidebar.subheader("")

st.sidebar.subheader("Offene Antworten:")
uploaded_file1 = st.sidebar.file_uploader("Upload Excel-File mit offenen Antworten",type=["xlsx","xls", "xlsm"])
#st.sidebar.info("Die Spalte A in Excel mit den offenen Antworten muss 'IDNR' und Spalte B muss 'Name' heissen ")


if uploaded_file1:

    

    #dfAntworten = pd.read_excel(uploaded_file1, dtype={'Name': 'str'}) #, index_col=0
    dfAntworten = pd.read_excel(uploaded_file1) #, index_col=0
    
    originalTabellenexpander = st.expander("Rohdaten einsehen:")
    with originalTabellenexpander:
        st.dataframe(dfAntworten)


    st.subheader("")
    st.markdown("---")
    st.write("")
    
    #Nur Kolumen mit string-Variablen anzeigen df.select_dtypes(include=[object])
    #Dataframe dass nur textvariablen enthält:
    dfAntwortenNurText = dfAntworten.select_dtypes(include=[object])
    

    variablelAuswahl = st.selectbox("Text-Variable/Spalte auswählen, die codiert werden soll:", dfAntwortenNurText.columns)
    if variablelAuswahl !=[]:
        #st.write("Ausgewählte Variable: ",variablelAuswahl)
        #Name = variablelAuswahl
        dfAntworten['Name'] = dfAntworten[variablelAuswahl]
        #st.write("dfAntworten['Name'] : ",dfAntworten['Name']  )

    st.write("")

    IDvariablelAuswahl = st.selectbox("ID-Variable/Spalte auswählen:", dfAntworten.columns)
    if IDvariablelAuswahl !=[]:
        #st.write("Ausgewählte ID-Variable: ",IDvariablelAuswahl)
        #IDNR = IDvariablelAuswahl
        dfAntworten['IDNR'] = dfAntworten[IDvariablelAuswahl]
        #st.write("dfAntworten['IDNR'] : ",dfAntworten['IDNR']  )
    
    if variablelAuswahl == [] or IDvariablelAuswahl == []:
        st.warning("Bitte Variablen angeben")

    
    dfAntworten = dfAntworten[{'IDNR', 'Name'}]
	
    #Anzahl Zeichen in Name/offene Antwortspalte - einzelBuchstaben Nennungen raus bitte
    dfAntworten['AnzahlZeichen'] = dfAntworten['Name'].str.len()
    dfAntworten['Name'] = np.where(dfAntworten['AnzahlZeichen']== 1, "keine Antwort", dfAntworten.Name)
	
	#Ersetze leere Zellen
    dfAntworten['0'] = dfAntworten['Name'].fillna('keine Antwort')
    
   

    dfAntwortenMitSplit = dfAntworten

	#Spalten erstellen bei Kommagtrennten Antworten
    #dfAntwortenMitSplit = dfAntwortenMitSplit['Name'].str.split(',', expand=True)

    #Test - Split nach Komma und /oder  Leerschlag

    st.subheader("")
    st.markdown("---")

    st.write("Ev Splitting auswählen:")

    kommaButton=st.checkbox('Nach Komma splitten', value=True)
    leerschlagButton = st.checkbox('Nach Leerschlag splitten')
    if kommaButton == True and leerschlagButton == False:
        dfAntwortenMitSplit = dfAntwortenMitSplit['Name'].str.split(',', expand=True)
    
    if leerschlagButton == True and kommaButton == False:
        dfAntwortenMitSplit = dfAntwortenMitSplit['Name'].str.split(' ', expand=True)

    if leerschlagButton == True and kommaButton == True:
         dfAntwortenMitSplit = dfAntwortenMitSplit['Name'].str.split(',| ', expand=True)



	#Möglichkeit: .str.split(',', n=1, expand=True) n=1 denotes that we want to make only one split.


    #Alle Missings mit keine Antwort ersetzen
    dfAntwortenMitSplit = dfAntwortenMitSplit.fillna('nichts / keine')



    # taufen wir die Spaltenköpfe um
    #for z in dfAntwortenMitSplit.columns:

     #  dfAntwortenMitSplit.rename(columns={z:'Nennung' + str(z+1)},inplace=True)

	#hängen wir IDNR dran
    #dfAntwortenMitSplit['IDNR'] = dfAntworten['IDNR']


    st.markdown("---")




    offeneAnwtortenTabellenExpander = st.expander("Tabellen mit den offenen Antworten einsehen:")
    with offeneAnwtortenTabellenExpander:
        st.write("dfAntworten vor Splitting: ",dfAntworten)
        st.write("dfAntwortenMitSplit nach Splitting", dfAntwortenMitSplit)

        anzahlSpalten = len(dfAntwortenMitSplit.columns)


    st.markdown("---")



    SpaltenAuswahl = st.multiselect("Ev Antwort-Spalten auswählen, die codiert werden sollen -  aktuell stehen " + str(anzahlSpalten) + " Spalten zur Verfügung",dfAntwortenMitSplit.columns)
    if SpaltenAuswahl !=[]:
      dfAntwortenMitSplit = dfAntwortenMitSplit[SpaltenAuswahl]
    
    #jetzt wollen wir Dataframe umstellen, so dass links Nennung und IDNR stehen, die Nennungsspalten untereinander eingereiht werden
    

	#Leeres Dataframe mit gewünschter Zusammenstellung für  fuzzy wuzzy


    df_Antworten_Fuzzzioniert = pd.DataFrame({'AIDNR': [],'Nennung': [], 'offeneAntwort':[]})



#=========== Iterationen mit globals Variable ==========================================================#

    for z in dfAntwortenMitSplit.columns:
       #Zuerst taufen wir die Spaltenköpfe um
       dfAntwortenMitSplit.rename(columns={z:'Nennung' + str(z+1)},inplace=True)
       
	   #Hier wird je Spalte ein neues Dataframe erstellt - funzt!!!
       globals()[f"df_{z+1}"] = dfAntwortenMitSplit['Nennung' + str(z+1)]

       #st.write("globals df sieht so aus:",  globals()[f"df_{z+1}"] )

       zwischen_Serie = globals()[f"df_{z+1}"]

       zwischen_df = pd.DataFrame(zwischen_Serie)

       zwischen_df['Nennung'] = "Nennung" + str(z+1)
       
       zwischen_df['AIDNR'] = dfAntworten['IDNR']

       zwischen_df['offeneAntwort'] = dfAntwortenMitSplit['Nennung' + str(z+1)]

		#Umstellung/Auswahl der Variablen
       zwischen_df = zwischen_df[{'AIDNR','Nennung','offeneAntwort'}]

       #st.write("zwischen_df: ",zwischen_df)

	   #Und hier noch Werte zu einem wachsenden Datadrame hinzufügen...
       df_Antworten_Fuzzzioniert = df_Antworten_Fuzzzioniert.append(zwischen_df, ignore_index=True)
    
    #st.write("df_Antworten_Fuzzzioniert: ",df_Antworten_Fuzzzioniert)
    #st.write("Anzahl Zeilen: ",len(df_Antworten_Fuzzzioniert))


#======================= Hier startet fuzzying =====================================================#

    st.subheader("")
    st.subheader("Automatische Codierung mit process-extract (fuzz.WRatio):")


    codebuchKategorie = []
    similarity = []
    Code = []

    with st.form("my_form"):
        
        submitted = st.form_submit_button("Codiere!")
        if submitted:
            with st.spinner('Bin am codieren....'):
                for i in df_Antworten_Fuzzzioniert.offeneAntwort :     
                    if pd.isnull( i ) :          
                        codebuchKategorie.append(np.nan)
                        similarity.append(np.nan)
                    else :          
                        ratio = process.extract( i, dfCodebuch.Name, limit=1, scorer=fuzz.WRatio)
                        codebuchKategorie.append(ratio[0][0])
                        similarity.append(ratio[0][1])
                        df_Antworten_Fuzzzioniert['ID'] = df_Antworten_Fuzzzioniert['AIDNR']

            st.success('Fertig')           

            df_Antworten_Fuzzzioniert['codebuchKategorie'] = pd.Series(codebuchKategorie)
            df_Antworten_Fuzzzioniert['codebuchKategorie'] = df_Antworten_Fuzzzioniert['codebuchKategorie'] #+ ' im Codebuch'
            df_Antworten_Fuzzzioniert['similarity'] = pd.Series(similarity)

            #st.write("df_Antworten_Fuzzzioniert nach Fuzzy: ",df_Antworten_Fuzzzioniert)
            st.write("Anzahl Zeilen nach Fuzzyionierung: ",len(df_Antworten_Fuzzzioniert))

            #Wenn similirity >= 80, schreibe 'OK' in die Spalte Codierungsresultat:

            GrenzWert = st.number_input("Grenzwert der Similiarity einstellen?", value=79)

            df_Antworten_Fuzzzioniert.loc[df_Antworten_Fuzzzioniert['similarity'] >= GrenzWert, 'Coderungsresultat'] = "OK"
            df_Antworten_Fuzzzioniert.loc[df_Antworten_Fuzzzioniert['similarity'] < GrenzWert, 'Coderungsresultat'] = "Codierung überprüfen"



            final_result = df_Antworten_Fuzzzioniert[['AIDNR','Nennung','offeneAntwort', 'codebuchKategorie','similarity', 'Coderungsresultat']]

            #Variable Name (=Codebuchkategorie) wird für merge mit Codebuch - wollen die Codes holen - benötigt:
            final_result['Name'] = final_result['codebuchKategorie']

            #st.write("final_result Tabelle vor merge: ", final_result)
            #st.write("Anzahl Zeilen vor merge: ",len(final_result.index))

            #merge um noch codes hinzuzufügen
            final_result = pd.merge(final_result, dfCodebuch, how='inner')

            #st.write("Anzahl Zeilen nach merge: ",len(final_result.index))

            #st.write("Final Result; ",final_result)

            #genial einfach formatierte tabelle
            st.write("Kontroll-Tabelle mit einer Übersicht der offenen Antworten, den ähnlichsten Kategorien aus dem Codebuch und die Similarity:")
            st.write(final_result.style.background_gradient(subset='similarity', cmap='summer_r'))

            #st.write(final_result.describe())
            #st.write(final_result.similarity.value_counts())
            anzahlHoheWerte = final_result[final_result['similarity']>= GrenzWert]
            st.write("Anzahl Werte mit mindestens " +str(GrenzWert) + "% similarity:",len(anzahlHoheWerte))
            st.write("Prozent-Anteil der Werte mit mindestens " +str(GrenzWert) + "% similarity:",int(100*(len(anzahlHoheWerte)/len(final_result.index))))
            
        



            #Umformatierung für Excelexport
            dfExcelExport['IDNR'] = final_result['AIDNR']
            dfExcelExport['Nennung'] = final_result['Nennung']
            dfExcelExport['Autocodierungsergebnis'] = final_result['Coderungsresultat']
            dfExcelExport['Codes'] = final_result['Codes']
            dfExcelExport['codebuchKategorie'] = final_result['codebuchKategorie']
            dfExcelExport['offeneAntwort'] = final_result['offeneAntwort']
            dfExcelExport['similarity'] = final_result['similarity']
            dfExcelExport.sort_values(by=['IDNR'], inplace=True)

            #Falls Autocodierungsresultat nicht ok, soll das Autocodeergebnis nicht angezeigt werden
            dfExcelExport.loc[dfExcelExport['Autocodierungsergebnis'] != "OK", 'codebuchKategorie'] = "nicht erkannt"
            dfExcelExport.loc[dfExcelExport['Autocodierungsergebnis'] != "OK", 'Codes'] = 99




            st.write("Anzahl Zeilen: ",len(dfExcelExport))
            anzahlCodierteZeilen = len(dfExcelExport)
            #st.write("dfExcelExport: ",dfExcelExport )
        
        AutoCodierung_Expander = st.expander("Kontroll-Tabelle mit allen offenen Antworten, Codes, similarity einsehen:")

        with AutoCodierung_Expander:
            if len(dfExcelExport) > 0:
                        st.write(dfExcelExport)
                        speicherZeitpunkt = pd.to_datetime('today')
                        st.write("")				
                        def to_excel(dfExcelExport):
                                output = BytesIO()
                                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                                dfExcelExport.to_excel(writer, index=False, sheet_name='Sheet1')
                                workbook = writer.book
                                worksheet = writer.sheets['Sheet1']
                                format1 = workbook.add_format({'num_format': '0.00'}) 
                                worksheet.set_column('A:A', None, format1)  
                                writer.save()
                                processed_data = output.getvalue()
                                return processed_data




        FertigesDatenfile_Expander = st.expander("Fertige Tabelle mit allen Codes in den Spalten für direkte Anwendung in SPSS:")
        with FertigesDatenfile_Expander:
            if len(dfExcelExport) > 0:
                #Pivotierter Tabelle für direktübernahme in SPSS
                dfExcelExportPivotiert = dfExcelExport.pivot(index='IDNR', columns='Nennung')['Codes']
                dfExcelExportPivotiert['Particpant'] = dfExcelExportPivotiert.index
                st.write("dfExcelExportPivotiert: ", dfExcelExportPivotiert)
                
                speicherZeitpunkt = pd.to_datetime('today')
                st.write("")

                if len(dfExcelExportPivotiert) > 0:					
                                def to_excel(dfExcelExportPivotiert):
                                    output = BytesIO()
                                    writer = pd.ExcelWriter(output, engine='xlsxwriter')
                                    dfExcelExportPivotiert.to_excel(writer, index=False, sheet_name='Sheet1')
                                    workbook = writer.book
                                    worksheet = writer.sheets['Sheet1']
                                    format1 = workbook.add_format({'num_format': '0.00'}) 
                                    worksheet.set_column('A:A', None, format1)  
                                    writer.save()
                                    processed_data = output.getvalue()
                                    return processed_data



if anzahlCodierteZeilen > 0:
    df_xlsx = to_excel(dfExcelExport)
    st.download_button(label='📥 Kontroll-Tabelle in Excel abspeichern?',
                                    data=df_xlsx ,
                                    file_name= 'Autocodierte Kontroll-Tabelle '+str(speicherZeitpunkt) +'.xlsx' )

if anzahlCodierteZeilen > 0:
    df_xlsx = to_excel(dfExcelExportPivotiert)
    st.download_button(label='📥 Fertige Tabelle in Excel abspeichern?',
                                    data=df_xlsx ,
                                    file_name= 'Autocodierte Datentabelle '+str(speicherZeitpunkt) +'.xlsx' )
    if len(dfExcelExportPivotiert) != len(dfAntworten):
                st.warning(' Obacht - Die Anzahl Zeilen vom Exportfile stimmen nicht mit dem Originaldatenfile überein', icon="⚠️")