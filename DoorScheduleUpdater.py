import streamlit as st
import pandas as pd
import base64
import numpy as np
import math
import matplotlib.pyplot as plt
import seaborn as sns

st. set_page_config(layout="wide") 

def get_binary_file_downloader_html(bin_file, file_label='File'):
    '''
    Generates an HTML download link for binary files.
    Args:
    - bin_file (io.BytesIO) - The binary file.
    - file_label (str) - The name of the file.
    Returns:
    - (str) - The HTML download link.
    '''
    with open(bin_file, 'rb') as f:
        data = f.read()
    bin_str = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{file_label}.xlsx">Download {file_label} File</a>'
    return href

def try_convert_to_float(val):
    try:
        number = float(val.replace(',', '.').replace("'", ''))
        # Check if number is integer-like after conversion
        if number.is_integer():
            return int(number)
        return number
    except:
        return val
    
def convert_to_int(value):
    # Check if the value is NaN (Not a Number)
    if isinstance(value, float) and math.isnan(value):
        return np.nan  # or return '---' if you prefer
    elif isinstance(value, (float, int)):
        return int(value)
    return value

def create_door_type_chart():
    # Count the occurrences of each door type
    door_type_counts = updated["Türtyp (Optionen-Set)"].value_counts(dropna=False)
    # If NaN/None values exist, rename them to a more readable label
    if np.nan in door_type_counts.index:
        door_type_counts = door_type_counts.rename({np.nan: 'NaN/None'})

    # Use Seaborn for plotting
    fig, ax = plt.subplots(figsize=(10,6))
    sns.barplot(x=door_type_counts.index, y=door_type_counts.values, palette='viridis', ax=ax)
    
    # Annotate each bar with its count
    for p in ax.patches:
        ax.annotate(f'{int(p.get_height())}', 
                    (p.get_x() + p.get_width() / 2., p.get_height()), 
                    ha='center', va='center', 
                    fontsize=10, color='black',
                    xytext=(0, 5), 
                    textcoords='offset points')

    ax.set_title('Verteilung der Türtypen')
    ax.set_xlabel('Türtyp S&S')
    ax.set_ylabel('Anzahl')
    ax.set_xticklabels(door_type_counts.index, rotation=45, ha='right')  # Adjust rotation and alignment for better readability
    st.pyplot(fig)


def lichtes_durchgangsmass_breite():
    # Calculate clear width "Lichtes Durchgangsmass Breite"
    # Convert to float
    updated["Breite Lichtmass"] = updated["Breite Lichtmass"].apply(try_convert_to_float)
    # Multiply by 1000
    updated["Breite Lichtmass"] = pd.to_numeric(updated["Breite Lichtmass"], errors='coerce')
    updated["Lichtes Durchgangsmass Breite in mm (Zeichenfolge)"] = (updated["Breite Lichtmass"] * 1000).round()
    # Replace NaNs with a placeholder string
    updated["Lichtes Durchgangsmass Breite in mm (Zeichenfolge)"].fillna('---', inplace=True)
    # Convert the entire column to string type
    updated["Lichtes Durchgangsmass Breite in mm (Zeichenfolge)"] = updated["Lichtes Durchgangsmass Breite in mm (Zeichenfolge)"].astype(str)
    # Remove '.0' from the strings
    updated["Lichtes Durchgangsmass Breite in mm (Zeichenfolge)"] = updated["Lichtes Durchgangsmass Breite in mm (Zeichenfolge)"].str.replace('.0', '', regex=False)

def lichtes_durchgangsmass_höhe():
    # Calculate clear height "Lichtes Durchgangsmass Höhe"
    # Convert to float
    updated["Höhe Lichtmass"] = updated["Höhe Lichtmass"].apply(try_convert_to_float)
    # Multiply by 1000
    updated["Höhe Lichtmass"] = pd.to_numeric(updated["Höhe Lichtmass"], errors='coerce')
    updated["Lichtes Durchgangsmass Höhe in mm (Zeichenfolge)"] = (updated["Höhe Lichtmass"] * 1000).round()
    # Replace NaNs with a placeholder string
    updated["Lichtes Durchgangsmass Höhe in mm (Zeichenfolge)"].fillna('---', inplace=True)
    # Convert the entire column to string type
    updated["Lichtes Durchgangsmass Höhe in mm (Zeichenfolge)"] = updated["Lichtes Durchgangsmass Höhe in mm (Zeichenfolge)"].astype(str)
    # Remove '.0' from the strings
    updated["Lichtes Durchgangsmass Höhe in mm (Zeichenfolge)"] = updated["Lichtes Durchgangsmass Höhe in mm (Zeichenfolge)"].str.replace('.0', '', regex=False)

def rohbaumass_breite():
    # Calculate clear width "Lichtes Durchgangsmass Breite"
    # Convert to float
    updated["Breite Rohbau"] = updated["Breite Rohbau"].apply(try_convert_to_float)
    # Multiply by 1000
    updated["Breite Rohbau"] = pd.to_numeric(updated["Breite Rohbau"], errors='coerce')
    updated["Rohbaumass Breite in mm (Zeichenfolge)"] = (updated["Breite Rohbau"] * 1000).round()
    # Replace NaNs with a placeholder string
    updated["Rohbaumass Breite in mm (Zeichenfolge)"].fillna('---', inplace=True)
    # Convert the entire column to string type
    updated["Rohbaumass Breite in mm (Zeichenfolge)"] = updated["Rohbaumass Breite in mm (Zeichenfolge)"].astype(str)
    # Remove '.0' from the strings
    updated["Rohbaumass Breite in mm (Zeichenfolge)"] = updated["Rohbaumass Breite in mm (Zeichenfolge)"].str.replace('.0', '', regex=False)

def rohbaumass_höhe():
    # Calculate clear width "Lichtes Durchgangsmass Breite"
    # Convert to float
    updated["Höhe Rohbau"] = updated["Höhe Rohbau"].apply(try_convert_to_float)
    # Multiply by 1000
    updated["Höhe Rohbau"] = pd.to_numeric(updated["Höhe Rohbau"], errors='coerce')
    updated["Rohbaumass Höhe in mm (Zeichenfolge)"] = (updated["Höhe Rohbau"] * 1000).round()
    # Replace NaNs with a placeholder string
    updated["Rohbaumass Höhe in mm (Zeichenfolge)"].fillna('---', inplace=True)
    # Convert the entire column to string type
    updated["Rohbaumass Höhe in mm (Zeichenfolge)"] = updated["Rohbaumass Höhe in mm (Zeichenfolge)"].astype(str)
    # Remove '.0' from the strings
    updated["Rohbaumass Höhe in mm (Zeichenfolge)"] = updated["Rohbaumass Höhe in mm (Zeichenfolge)"].str.replace('.0', '', regex=False)

def update_wandtyp_based_on_wandstruktur(updated):

    def update_wandstruktur(row):
        # Function to parse Wandstruktur and update Wandtyp (Optionen-Set)
        walltypes = {
            'BE': ['BE', 'SB'],  # If contains 'BE' or 'SB' then 'BE' is added
            'MW': ['MW'],
            'LB': ['LB'],
            'GS': ['GS'],
            'NS': ['NS']
        }

        wandstruktur = row['Wandstruktur']
        if isinstance(wandstruktur, str):  # Check if the value is a string
            for walltype, keywords in walltypes.items():
                for keyword in keywords:
                    if keyword in wandstruktur:
                        return walltype

        return row['Wandtyp (Zeichenfolge)']

    updated['Wandtyp (Zeichenfolge)'] = updated.apply(update_wandstruktur, axis=1)
    updated['Wandtyp (Zeichenfolge)'] = updated['Wandtyp (Zeichenfolge)'].astype(str)
    return updated



st.title("Türmatrix Updater")

uploaded_file = st.file_uploader("Upload an XLSX file", type=["xlsx"])

if uploaded_file:
    # Read the xlsx file
    original = pd.read_excel(uploaded_file, sheet_name="Türmatrix")
    updated = pd.read_excel(uploaded_file, sheet_name="Türmatrix", skiprows=1)
    # Display the original dataframe
    with st.expander("Türmatrix - original Export aus ArchiCAD"):
        st.write(original)
    # with st.expander("Türmatrix - Vorbereitung für Aktualisierung der Daten"):
    #     st.write(updated)
    # Replace '---' strings to NaN
    updated["Breite Lichtmass"].replace('---', np.nan, inplace=True)
    # Convert the column to a string, replace commas, then to float
    for col in updated.columns:
        updated[col] = updated[col].apply(try_convert_to_float)

    lichtes_durchgangsmass_breite()

    lichtes_durchgangsmass_höhe()

    rohbaumass_breite()

    rohbaumass_höhe()

    update_wandtyp_based_on_wandstruktur(updated)

    # with st.expander("Türmatrix - Aktualisierung der Durchgangsbreiten"):
    #     st.write(updated)
    #     st.write(updated["Lichtes Durchgangsmass Breite in mm (Zeichenfolge)"].dtypes)

    with st.expander("Türmatrix - Dashboard"):
        create_door_type_chart()

   
    # Save updated dataframe as combined
    combined = updated.copy()

    # Insert the column indexes from updated as the first row of combined
    combined.loc[-1] = combined.columns.tolist()  # adding a row with labels
    combined.index = combined.index + 1  # shifting index
    combined = combined.sort_index()  # reordering

    # Set the column indexes of combined to be the same as those of original
    combined.columns = original.columns

    #remove nan string values
    combined = combined.replace("nan", "")

    # Optional: Display the combined dataframe
    with st.expander("Türmatrix - bereit für Import ins ArchiCAD"):
        st.write(combined)

    # Create an Excel writer and write the combined DataFrame to it
    output = "output.xlsx"
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        combined.to_excel(writer, sheet_name="Türmatrix", index=False)

    # Display the download link
    st.markdown(get_binary_file_downloader_html(output, 'Updated_Türmatrix'), unsafe_allow_html=True)
