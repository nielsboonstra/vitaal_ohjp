import streamlit as st
import pandas as pd
import regex as re

st.set_page_config(page_title="VITAAL | Ultimo naar OHJP Conversie Tool", layout="wide")
st.logo("https://raw.githubusercontent.com/nielsboonstra/vitaal_ohjp/0c188e531f78210f2f59043fc362bb5b6bf8087d/vitaal_logo.png")

@st.cache_data
def load_excel(file : str, header : int = 0) -> pd.DataFrame | None:
    try:
        df = pd.read_excel(file, engine='openpyxl', header=header)
        return df
    except Exception as e:
        st.error(f"Fout bij het laden van het Excel-bestand: {e}")
        return None

def normalize_frequency(row : pd.Series) -> pd.Series:
    """Normalize the frequency of the tasks in the DataFrame.
    Frequencies "WK" and "JR" are converted to "MD".
    Column "Frequentie aantal" is changed using multiplication or division.
    
    :param row: A row of the DataFrame.
    :return: The modified row with normalized frequency.
    """

    if row["Frequentie"] == "WK":
        row["Frequentie aantal"] = row["Frequentie aantal"] * 0.25
        row["Frequentie"] = "MD"
    elif row["Frequentie"] == "JR":
        row["Frequentie aantal"] = row["Frequentie aantal"] * 12
        row["Frequentie"] = "MD"
    return row

def extract_object(traject_complex_list : list, omschrijving_list : list) -> list:
    """
    Extract the object of interest from the list of trajects and complexen.
    
    :param traject_complex_list: List of trajects and complexen.
    :param omschrijving_list: List of omschrijvingen.
    :return: List of objects.
    """

    objects = []
    is_complex  = []

    for i, item in enumerate(traject_complex_list):
        if 'complex' in item.lower(): # If string contains 'complex', add part of string after 'complex' to the list. E.g., 'Stuw- en sluiscomplex Amerongen' becomes 'Amerongen'.
            objects.append(item.split('complex')[-1].strip())
            is_complex.append(True) # Object of interest is a complex.
        elif 'verkeerscentrale' in item.lower(): # Verkeerscentrales are out of scope for the OHJP.
            if 'waalbrug' in omschrijving_list[i].lower(): # Exception: Waalbrug falls under the verkeerscentrale category in the decomposition, but is within scope.
                # Add Waalbrug to the list of objects and set is_complex to False.
                objects.append('Waalbrug')
                is_complex.append(False)
            else:
                objects.append(-1) # Mark as -1, so it can be easily filtered out later.
                is_complex.append(False)
        elif "eilandbrug" in item.lower(): # Eilandbrug is a complex, but does not contain 'complex' in the string. Add manually to complex list.
            objects.append('Eilandbrug')
            is_complex.append(True)
        else:
            #Look for object in omschrijving_list. It can be assumed that the object can be found after the first '-' in the string.
            # If the object cannot be found, raise an error.
            if '-' in omschrijving_list[i]:
                object = omschrijving_list[i].split('-')[1].strip()
                # If a '/' is present in the object string, split the string and take the first part.
                if '/' in object:
                    object = object.split('/')[0].strip()
                if object == '':
                    raise ValueError(f"Object not found in {omschrijving_list[i]}. Is the Ultimo job named correctly?") # Raise error if no object could be found in an Ultimo job
                else:
                    objects.append(object)
                    is_complex.append(False)
    return objects, is_complex

def create_heatmap_df(df : pd.DataFrame, start_week : int = 1) -> pd.DataFrame:
    """
    Create a heatmap dataframe with unique values under "Omschrijving" as index and week numbers as columns.
    The values in the dataframe are >= 1 if the planning block is present, otherwise it will be 0.
    All week numbers are present in the columns.
    """

    # Create week nr columns 1 to 52
    week_numbers = list(range(1, 53))
    for week in week_numbers:
        df[str(week)] = 0


    # Create a pivot table with "Omschrijving" as index, "Week" as columns, and count of occurrences as values
    heatmap_df = df.pivot_table(index=['Omschrijving'], columns='Week', aggfunc='size', fill_value=0)
    heatmap_df = heatmap_df.reset_index()

    #Also add week numbers that had no values
    for week in week_numbers:
        if week not in heatmap_df.columns:
            heatmap_df[week] = 0
    
    #Find metadata (Frequentie aantal, Frequentie, Uitvoerende) and add it to the heatmap_df. Using the first occurrence of the metadata in the original dataframe. Add the metadata to the heatmap_df as new columns.
    metadata = df[['Omschrijving', 'Object', 'Traject of Complex', 'Uitvoerende', 'is_complex', 'Frequentie aantal', 'Frequentie']].drop_duplicates(subset=['Omschrijving', 'Frequentie aantal'])    

    heatmap_df = heatmap_df.merge(metadata, on='Omschrijving', how='left')

    #sort week_numbers starting with start_week to 52, then 1 to start_week-1
    week_numbers = [week for week in week_numbers[start_week-1:] + week_numbers[:start_week-1]]
    # Sort columns: metadata first, then week numbers
    cols = ['Object', 'Omschrijving','Frequentie aantal', 'Traject of Complex', 'Uitvoerende', 'is_complex'] + week_numbers
    heatmap_df = heatmap_df[cols]

    return heatmap_df

st.title("🏃 VITAAL | Ultimo naar OHJP Conversie Tool")

st.write("Deze tool helpt jou met het omzetten van een OMS-export uit Ultimo naar een OHJP-worksheet" \
" voor het opzetten van een onderhoudsjaarplan.")

with st.sidebar:
    st.markdown("**1. Upload een OMS-exportbestand (Excel-format) om te beginnen:** 👇")
    uploaded_file = st.file_uploader("Kies een bestand", type=["xlsx"])
    header_row = st.number_input("Kies de rij waar de kolomnamen staan in je Excel (standaard is 0)", min_value=0, value=0)
    if uploaded_file is not None:
        df = load_excel(uploaded_file, header=header_row)
        if df is not None:
            st.session_state['df'] = df
            st.success("OMS-exportbestand succesvol geladen!")

#On the main page, give user the option to preview both uploaded files that can be hidden/collapsed by clicking on an arrow
if 'df' in st.session_state:
    with st.expander("📊 Bekijk de OMS-export data", expanded=False):
        st.dataframe(st.session_state['df'])

if 'df' in st.session_state:
    st.markdown("**Start de conversie naar OHJP:**")
    # Ask user for start year and start week for the planning
    start_year = st.number_input("Kies het startjaar voor de planning:", min_value=2025, max_value=2100)
    start_week = st.number_input("Kies de startweek voor de planning:", min_value=1, max_value=52, value=27)
    naam_export = st.text_input("Kies de naam voor het exportbestand (zonder extensie):", value="OHJP [X]e contractjaar VITAAL")
    if st.button("Start conversie"):

        with st.spinner("Bezig met het converteren van de OMS-export naar OHJP...", show_time=True):
            df = st.session_state['df']
            df = df.drop(columns=["Id", "Beheerobject"])
            df =df.rename(columns={"Traject of Complex" : "Complex"})
            object_data = extract_object(df["Complex"], df["Omschrijving"])
            df['Object'] = object_data[0]
            df['is_complex'] = object_data[1]
            df = df[df['Object'] != -1].reset_index(drop=True) #Drop rows with -1 in the 'Object' column. These are the verkeerscentrales Tiel and Nijmegen.
            # If is_complex == False, then set "Complex" to "Vaste objecten"
            df.loc[df['is_complex'] == False, 'Complex'] = 'Vaste objecten'
            df = df.apply(normalize_frequency, axis=1) # Transform all frequencies that are not in Months to Months
            #Turn Startdatum and Einddatum wk into integers
            df["Week"] = df["Start week"].astype(int)
            df["Week_end"] = df["Gereed week"].astype(int)
            df['Weeks'] = df.apply(lambda row: list(range(row['Week'], row['Week_end'] + 1)), axis=1) # To do: Implement this in create_heatmap_df
            # If there are values >= "start_year+1"+"start_week", then print a warning and remove those values
            week_threshold = int(str(start_year + 1) + str(start_week))
            st.write(week_threshold)
            df_to_remove = df[df["Week"] >= week_threshold]
            if not df_to_remove.empty:
                st.warning(f"Er zijn taken gepland na week {start_week} van {start_year + 1}. Deze worden niet meegenomen in de planning.")
            df = df[df["Week"] < week_threshold] # Remove rows

            #Create Heatmap df's per object.
            complexes = df['Complex'].unique()
            heatmap_dfs_complex = {}

            for complex in complexes:
                # Filter the dataframe for the current traject
                filtered_df = df[df['Complex'] == complex]
                complex_df = create_heatmap_df(filtered_df,start_week=start_week)
                heatmap_dfs_complex[complex] = complex_df

            # Save the heatmap dataframes to separate Excel worksheets with the name of the traject
            with pd.ExcelWriter(f'{naam_export}.xlsx') as writer:
                for complex, heatmap_df in heatmap_dfs_complex.items():
                    heatmap_df = heatmap_df.drop(columns="Complex")
                    heatmap_df.to_excel(writer, sheet_name=complex, index=False)
        st.success(f"Conversie voltooid! De OHJP-gegevens worden opgeslagen als '{naam_export}.xlsx'.")
        st.write("Je kunt het bestand downloaden via de onderstaande knop:")
        st.download_button(
            label="Download OHJP Excel-bestand",
            data=open(f'{naam_export}.xlsx', 'rb').read(),
            file_name=f'{naam_export}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        st.write("Kopieer en plak de inhoud van de Excel per worksheet naar een OHJP-template dat past bij jouw wensen.")
        st.balloons()  # Show balloons to celebrate the successful conversion