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
import plotly.express as px
import plotly.graph_objects as go

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
    st.info("Obacht - nan darf nicht als Kategorie im Codebuch vorkommen")
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
                            writer.close()
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
    
    st.info("Text-Variable/Spalte auswählen")
    variablelAuswahl = st.selectbox("Text-Variable/Spalte auswählen, die codiert werden soll:", dfAntwortenNurText.columns)
    if variablelAuswahl !=[]:
        #st.write("Ausgewählte Variable: ",variablelAuswahl)
        #Name = variablelAuswahl
        dfAntworten['Name'] = dfAntworten[variablelAuswahl]
        #st.write("dfAntworten['Name'] : ",dfAntworten['Name']  )

    st.write("")
    st.write("")

    st.info("ID-Variable/Spalte auswählen")
    IDvariablelAuswahl = st.selectbox("ID-Variable/Spalte auswählen. Erleichtert die Datenkontrolle:", dfAntworten.columns)
    if IDvariablelAuswahl !=[]:
        #st.write("Ausgewählte ID-Variable: ",IDvariablelAuswahl)
        #IDNR = IDvariablelAuswahl
        dfAntworten['IDNR'] = dfAntworten[IDvariablelAuswahl]
        #st.write("dfAntworten['IDNR'] : ",dfAntworten['IDNR']  )
    
    if variablelAuswahl == [] or IDvariablelAuswahl == []:
        st.warning("Bitte Variablen angeben")

    
    dfAntworten = dfAntworten[['IDNR', 'Name']]

    #Test - gewisse Wörter zusammenschreiben
    dfAntworten['Name'] = dfAntworten['Name'].str.lower()
    a = ["credit suisse", "bank cler", "migros bank"]
    b = ["credit_suisse", "bank_cler", "migros_bank"]
    dfAntworten['Name'] = dfAntworten['Name'].replace(a,b,regex=True)
    #dfAntworten['Name'] = dfAntworten['Name'].str.replace(a,b, case=False,regex=True)

	
    #Anzahl Zeichen in Name/offene Antwortspalte - einzelBuchstaben Nennungen raus bitte
    dfAntworten['AnzahlZeichen'] = dfAntworten['Name'].str.len()
    dfAntworten['Name'] = np.where(dfAntworten['AnzahlZeichen']== 1, "keine Antwort", dfAntworten.Name)

    dfAntworten['AnzahlWorte'] = dfAntworten['Name'].str.split().str.len()
    dfAntworten['AnzahlKommata'] = dfAntworten['Name'].str.count(",")

    
    st.write("")
    st.markdown("---")
    st.write("")

    #wenn mehere Worte aber kein Komma -> füge Komma(s) ein
    kommataEinfuegen = st.checkbox("Kommata einfügen? (statt Leerzeichen in Antworten mit mehr als 2 Worten aber ganz ohne Kommas)")
    if kommataEinfuegen:
        dfAntworten['AntwortenVorBearbeitung'] = dfAntworten['Name']
        dfAntworten['Name'] = np.where((dfAntworten['AnzahlKommata']<1) & (dfAntworten['AnzahlWorte']>2),dfAntworten['Name'].str.replace(" ",","), dfAntworten.Name)
        kommataEinfuegungsCheck = st.expander("Datenfile nach Kommata ergänzungen")
        with kommataEinfuegungsCheck:
             st.dataframe(dfAntworten)


	#Ersetze leere Zellen
    dfAntworten['0'] = dfAntworten['Name'].fillna('keine Antwort')
    
   

    dfAntwortenMitSplit = dfAntworten

	#Spalten erstellen bei Kommagtrennten Antworten
    #dfAntwortenMitSplit = dfAntwortenMitSplit['Name'].str.split(',', expand=True)

    #Test - Split nach Komma und /oder  Leerschlag

    st.subheader("")
    #dfAntwortenMitSplit = dfAntwortenMitSplit['Name'].str.strip()

    st.subheader("")
    st.markdown("---")

    leereZellenLoeschen = st.checkbox("Alle leere Textstellen löschen?",value=False)
    if leereZellenLoeschen:
         dfAntwortenMitSplit['Name'] = dfAntwortenMitSplit['Name'].str.replace(" ","")
         dfCodebuch['Name'] = dfCodebuch['Name'].str.replace(" ","")
         leerzellenDFsAnzeigen = st.button("Datensatz und Codebuch anzeigen?")
         if leerzellenDFsAnzeigen:
            st.dataframe(dfCodebuch)
            st.dataframe(dfAntwortenMitSplit)

    st.subheader("")
    st.markdown("---")




    st.subheader("")
    st.markdown("---")

    zeilenUmbruecheErsetzen = st.checkbox("Versteckte Zeilenumbrüche aus Excel mit Kommas ersetzen?",value=True)
    if zeilenUmbruecheErsetzen:
         dfAntwortenMitSplit['Name'] = dfAntwortenMitSplit['Name'].str.replace(r'\r\n|\r|\n', ', ')
         dfAntwortenMitSplit['Name'] = dfAntwortenMitSplit['Name'].str.replace('\n', ', ').str.replace('\r', ', ')


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

    if leerschlagButton == False and kommaButton == False:
        dfAntwortenMitSplit['AIDNR'] = dfAntwortenMitSplit['IDNR']
        dfAntwortenMitSplit['Nennung'] = "Nennung1"
        dfAntwortenMitSplit['offeneAntwort'] = dfAntwortenMitSplit['Name']
        #dfAntwortenMitSplit.drop(['Name', 'IDNR','Nennung'], axis=1)
        #st.write("dfAntwortenMitSplit: ",dfAntwortenMitSplit)

    #df_Antworten_Fuzzzioniert = pd.DataFrame({'AIDNR': [],'Nennung': [], 'offeneAntwort':[]})
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
    st.write("anzahlSpalten: ",anzahlSpalten)


    st.markdown("---")

    anzahlAusgewaehlteSpalten = 1

    SpaltenAuswahl = st.multiselect("Ev Antwort-Spalten auswählen, die codiert werden sollen -  aktuell stehen " + str(anzahlSpalten) + " Spalten zur Verfügung",dfAntwortenMitSplit.columns)
    if SpaltenAuswahl !=[]:
      dfAntwortenMitSplit = dfAntwortenMitSplit[SpaltenAuswahl]
      anzahlAusgewaehlteSpalten = len(SpaltenAuswahl)
      st.write("Anzahl ausgewählte Spalten:", anzahlAusgewaehlteSpalten)
    
    #jetzt wollen wir Dataframe umstellen, so dass links Nennung und IDNR stehen, die Nennungsspalten untereinander eingereiht werden
    

	#Leeres Dataframe mit gewünschter Zusammenstellung für  fuzzy wuzzy


    df_Antworten_Fuzzzioniert = pd.DataFrame({'AIDNR': [],'Nennung': [], 'offeneAntwort':[]})


    if leerschlagButton == False and kommaButton == False:
        df_Antworten_Fuzzzioniert = dfAntwortenMitSplit[['AIDNR','Nennung','offeneAntwort']]



#=========== Iterationen mit globals Variable ==========================================================#


    else:

        for z in dfAntwortenMitSplit.columns:
        #Zuerst taufen wir die Spaltenköpfe um

            #st.write("z:",z)     
            zaeler = z + 1
            zaelerAlsText = str(zaeler)
            #st.write("zaelerAlsText: ",zaelerAlsText)

            dfAntwortenMitSplit.rename(columns={z:'Nennung' + str(zaelerAlsText)},inplace=True)
            
            #Hier wird je Spalte ein neues Dataframe erstellt - funzt!!!
            globals()[f"df_{z+1}"] = dfAntwortenMitSplit['Nennung' + str(zaelerAlsText)]

            #st.write("globals df sieht so aus:",  globals()[f"df_{z+1}"] )

            zwischen_Serie = globals()[f"df_{zaelerAlsText}"]

            zwischen_df = pd.DataFrame(zwischen_Serie)

            zwischen_df['Nennung'] = "Nennung" + str(z+1)
            
            zwischen_df['AIDNR'] = dfAntworten['IDNR']

            zwischen_df['offeneAntwort'] = dfAntwortenMitSplit['Nennung' + str(zaelerAlsText)]

                #Umstellung/Auswahl der Variablen
            #code bis 2023.07.12 zwischen_df = zwischen_df[{'AIDNR','Nennung','offeneAntwort'}]
            zwischen_df = zwischen_df[['AIDNR','Nennung','offeneAntwort']]

            #st.write("zwischen_df: ",zwischen_df)

            #Und hier noch Werte zu einem wachsenden Datadrame hinzufügen...
            #Code bis 2023.07.12
            #df_Antworten_Fuzzzioniert = df_Antworten_Fuzzzioniert.append(zwischen_df, ignore_index=True)
            #You need to use concat instead (for most applications):

            df_Antworten_Fuzzzioniert = pd.concat([df_Antworten_Fuzzzioniert, zwischen_df]).reset_index(drop=True)
            
    
            #st.write("df_Antworten_Fuzzzioniert: ",df_Antworten_Fuzzzioniert)
            #st.write("Anzahl Zeilen: ",len(df_Antworten_Fuzzzioniert))


#======================= Hier startet fuzzying =====================================================#

    st.subheader("")
    st.subheader("Automatische Codierung mit process-extract (fuzz.WRatio):")

    #st.write("df_Antworten_Fuzzzioniert: ",df_Antworten_Fuzzzioniert)
    #st.write(df_Antworten_Fuzzzioniert.Nennung.unique())
    #st.write(len(df_Antworten_Fuzzzioniert.offeneAntwort))

    codebuchKategorie = []
    similarity = []
    Code = []


    with st.form("my_form"):
        
        submitted = st.form_submit_button("Codiere!")
        if submitted:
            with st.spinner('Bin am codieren....'):
                for i in df_Antworten_Fuzzzioniert.offeneAntwort:     
                    if pd.isnull( i ) :          
                        codebuchKategorie.append(np.nan)
                        similarity.append(np.nan)
                    else :          
                        ratio = process.extract( i, dfCodebuch.Name, limit=1, scorer=fuzz.WRatio)
                        codebuchKategorie.append(ratio[0][0])
                        similarity.append(ratio[0][1])
                        df_Antworten_Fuzzzioniert['ID'] = df_Antworten_Fuzzzioniert['AIDNR']

            st.success('Fertig!')           

            df_Antworten_Fuzzzioniert['codebuchKategorie'] = pd.Series(codebuchKategorie)
            df_Antworten_Fuzzzioniert['codebuchKategorie'] = df_Antworten_Fuzzzioniert['codebuchKategorie'] #+ ' im Codebuch'
            df_Antworten_Fuzzzioniert['similarity'] = pd.Series(similarity)

            #st.write("df_Antworten_Fuzzzioniert nach Fuzzy: ",df_Antworten_Fuzzzioniert)
            st.write("Anzahl Zeilen nach Fuzzyionierung: ",len(df_Antworten_Fuzzzioniert.index))

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



            st.write("")
            #st.write(final_result.describe())
            #st.write(final_result.similarity.value_counts())
            anzahlHoheWerte = final_result[final_result['similarity']>= GrenzWert]
            st.write("Anzahl Werte mit mindestens " +str(GrenzWert) + "% similarity:",len(anzahlHoheWerte))
            st.write("Prozent-Anteil der Werte mit mindestens " +str(GrenzWert) + "% similarity:",int(100*(len(anzahlHoheWerte)/len(final_result.index))))
            
        
            #genial einfach formatierte tabelle
            farbigeKontrollTabelle = st.expander("Farbige Kontroll-Tabelle mit einer Übersicht der offenen Antworten, leider etwas langsam")
            with farbigeKontrollTabelle:
                st.write("Kontroll-Tabelle mit einer Übersicht der offenen Antworten, den ähnlichsten Kategorien aus dem Codebuch und die Similarity:")
                st.write(final_result.style.background_gradient(subset='similarity', cmap='summer_r'))


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

        _=""" Funktioniert plötzlich nicht mehr 2023.07.12
        ChartPreviewExpander = st.expander("Chart - Vorschau der Codierungsergebnisse:")

        with ChartPreviewExpander:
            if len(dfExcelExport) > 0:


                #Erst nur Auswahl interessanten Spalten
                df_codebuchKategorie = dfExcelExport [['codebuchKategorie']]
                # #Dann Umstellung in Tabellformat das geeignet für Abbildungen ist:
                df_codebuchKategorieAnzahl = df_codebuchKategorie['codebuchKategorie'].value_counts()
                st.write("df_codebuchKategorieAnzahl:", df_codebuchKategorieAnzahl)
                df_codebuchKategorieAnzahl = df_codebuchKategorieAnzahl.reset_index(level=0)
                st.write("df_codebuchKategorieAnzahl:", df_codebuchKategorieAnzahl)
                df_codebuchKategorieAnzahl = df_codebuchKategorieAnzahl.rename(columns={'codebuchKategorie':'Anzahl_Nennungen'})
                st.write("df_codebuchKategorieAnzahl:", df_codebuchKategorieAnzahl)
                
                

                df_codebuchKategorieAnzahl['Antwort'] = df_codebuchKategorieAnzahl['index']




                #Dann nue interessante Zeilen beahlten ohne keine Nennung usw
                df_codebuchKategorieAnzahl = df_codebuchKategorieAnzahl[((df_codebuchKategorieAnzahl.Antwort != 'nichts / keine') & (df_codebuchKategorieAnzahl.Antwort != 'nicht erkannt'))]
                df_codebuchKategorieAnzahl = df_codebuchKategorieAnzahl[((df_codebuchKategorieAnzahl.Antwort != 'nichts/keine') & (df_codebuchKategorieAnzahl.Antwort != 'nichterkannt'))]
                df_codebuchKategorieAnzahl['Prozentanteile'] = 100* df_codebuchKategorieAnzahl['Anzahl_Nennungen']/len(dfAntworten.index)
                
                st.write(df_codebuchKategorieAnzahl)
                #st.bar_chart(data=df_codebuchKategorieAnzahl, y='Prozentanteile', use_container_width=True)
                
                CodierungsPreviewstapeldiagramm = px.bar(df_codebuchKategorieAnzahl, x='Antwort', y= 'Prozentanteile', text ='Prozentanteile'
					#color_discrete_map={'Radio-RW' : FARBE_RADIO,'Zattoo-RW' : FARBE_ZATTOO ,'Kino-RW' : FARBE_KINO,'DOOH-RW' : FARBE_DOOH,'OOH-RW' : FARBE_OOH,'FACEBOOK-RW' : FARBE_FACEBOOK,'YOUTUBE-RW' : FARBE_YOUTUBE,'ONLINEVIDEO-RW' : FARBE_ONLINEVIDEO,'ONLINE-RW' : FARBE_ONLINE, 'TV-RW' : FARBE_TV, 'Gesamt-RW' : FARBE_GESAMT},
			,color='Antwort', 
			#color_discrete_map={'16 - 24 J.' : FARBE_16_24 ,'25 - 34 J.' : FARBE_25_34 ,'35 - 44 J.' : FARBE_35_44,'45 - 59 J.' : FARBE_45_59 },
			title="Ungestützte Nennungen",
			hover_data=['Prozentanteile'],
		)
                CodierungsPreviewstapeldiagramm.update_traces(texttemplate='%{text:.2s}', textposition='inside')
                CodierungsPreviewstapeldiagramm.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
                CodierungsPreviewstapeldiagramm.update_layout(showlegend=False)
                #CodierungsPreviewstapeldiagramm.update_layout(width=400,height=300)
		        # Change grid color and axis colors
                # CodierungsPreviewstapeldiagramm.update_yaxes(showline=True, linewidth=1, linecolor='white', gridcolor='Black')
                st.plotly_chart(CodierungsPreviewstapeldiagramm, use_container_width=True)


            """


            
        
        AutoCodierung_Expander = st.expander("Tip: Kontroll-Tabelle mit allen offenen Antworten, Codes, similarity einsehen:")

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
                                writer.close()
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
                                    writer.close()
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