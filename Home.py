import streamlit as st
import Summary
import SummaryAbacus
import SummaryBrinks
import SummaryTAG

st.set_page_config(page_title="Parser Dashboard", layout="wide")
st.title("Parser Dashboard")

# Buat menu di sidebar
menu = st.sidebar.radio(
    "Pilih Halaman Parser", ["Advantage", "Abacus", "Brinks", "TAG"]
)

# Panggil parser sesuai pilihan
if menu == "Advantage":
    Summary.run()
elif menu == "Abacus":
    SummaryAbacus.run()
elif menu == "Brinks":
    SummaryBrinks.run()
elif menu == "TAG":
    SummaryTAG.run()
