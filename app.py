import streamlit as st
import pandas as pd

# Streamlit app
def main():
    st.title("Trasforma Dati Taglie in Formato Verticale")

    # File uploader
    uploaded_file = st.file_uploader("Carica un file Excel", type=["xlsx"])
    
    if uploaded_file is not None:
        # Leggi il primo foglio del file Excel
        try:
            data = pd.read_excel(uploaded_file)
            st.write("Anteprima dei dati caricati (con nomi delle colonne):")
            st.dataframe(data)  # Visualizza i dati con nomi di colonne
            
            # Mostra le colonne disponibili
            st.write("Colonne disponibili:")
            st.write(list(data.columns))

            # Input per il range di colonne
            range_input = st.text_input(
                "Inserisci il range di colonne per le taglie (es. 'C:V')", value="C:V"
            )

            if range_input:
                try:
                    # Converte il range in indici di colonne
                    start_col, end_col = range_input.split(":")
                    cols_to_melt = data.loc[:, start_col:end_col].columns
                    
                    # Trasforma i dati
                    reshaped_data = data.melt(
                        id_vars=['STYLE', 'CODE', 'WHL'],  # Campi identificativi
                        value_vars=cols_to_melt,
                        var_name='SIZE',
                        value_name='QUANTITY'
                    )

                    # Rimuove righe con quantità NaN
                    reshaped_data = reshaped_data.dropna(subset=['QUANTITY'])

                    # Mostra l'output
                    st.write("Dati trasformati:")
                    st.dataframe(reshaped_data)

                    # Scarica il file riorganizzato
                    st.write("Scarica il file Excel riorganizzato:")
                    file_name = "Reshaped_Size_Quantity_and_WHL_Data.xlsx"
                    reshaped_data.to_excel(file_name, index=False)
                    with open(file_name, "rb") as file:
                        st.download_button(
                            label="Download",
                            data=file,
                            file_name="Reshaped_Size_Quantity_and_WHL_Data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except KeyError:
                    st.error("Errore: il range di colonne specificato non è valido. Assicurati di usare colonne esistenti.")
                except Exception as e:
                    st.error(f"Errore durante la trasformazione: {e}")
        except Exception as e:
            st.error(f"Errore durante la lettura del file: {e}")

if __name__ == "__main__":
    main()
