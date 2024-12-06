import streamlit as st
import pandas as pd

# Funzione per generare le lettere delle colonne in stile Excel
def generate_excel_columns(num_columns):
    letters = []
    for i in range(1, num_columns + 1):
        column_name = ""
        while i > 0:
            i, remainder = divmod(i - 1, 26)
            column_name = chr(65 + remainder) + column_name
        letters.append(column_name)
    return letters

# Streamlit app
def main():
    st.title("Trasforma Dati in Formato Verticale (Range di Colonne Personalizzabile)")

    # File uploader
    uploaded_file = st.file_uploader("Carica un file Excel", type=["xlsx"])
    
    if uploaded_file is not None:
        # Leggi il primo foglio del file Excel
        try:
            data = pd.read_excel(uploaded_file)
            num_columns = data.shape[1]
            
            # Genera i riferimenti delle colonne in stile Excel
            excel_columns = generate_excel_columns(num_columns)

            # Aggiungi i riferimenti stile Excel come intestazione
            data_with_references = data.copy()
            data_with_references.columns = excel_columns
            data_with_references.loc[-1] = data.columns  # Aggiungi la riga originale come prima riga
            data_with_references.index = data_with_references.index + 1
            data_with_references.sort_index(inplace=True)

            # Mostra l'anteprima dei dati con riferimenti stile Excel
            st.write("Anteprima dei dati caricati (con riferimenti stile Excel):")
            st.dataframe(data_with_references)

            # Input per il range di colonne
            range_input = st.text_input(
                "Inserisci il range di colonne da trasporre (es. 'C:V')", value="C:V"
            )

            if range_input:
                try:
                    # Converte il range in indici di colonne
                    start_col, end_col = range_input.split(":")
                    start_idx = excel_columns.index(start_col)
                    end_idx = excel_columns.index(end_col) + 1
                    cols_to_melt = data.iloc[:, start_idx:end_idx].columns

                    # Identifica le colonne non selezionate per il range
                    id_vars = data.drop(columns=cols_to_melt).columns

                    # Trasforma i dati
                    reshaped_data = data.melt(
                        id_vars=id_vars,  # Colonne non trasposte
                        value_vars=cols_to_melt,
                        var_name='SIZE',  # Nome per le colonne trasposte
                        value_name='VALUE'  # Nome per i valori trasposti
                    )

                    # Mostra l'output
                    st.write("Dati trasformati:")
                    st.dataframe(reshaped_data)

                    # Scarica il file riorganizzato
                    st.write("Scarica il file Excel riorganizzato:")
                    file_name = "Reshaped_Data.xlsx"
                    reshaped_data.to_excel(file_name, index=False)
                    with open(file_name, "rb") as file:
                        st.download_button(
                            label="Download",
                            data=file,
                            file_name="Reshaped_Data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except ValueError:
                    st.error("Errore: il range di colonne specificato non è valido. Assicurati di usare colonne esistenti.")
                except Exception as e:
                    st.error(f"Errore durante la trasformazione: {e}")
        except Exception as e:
            st.error(f"Errore durante la lettura del file: {e}")

if __name__ == "__main__":
    main()
