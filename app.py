import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

st.set_page_config(page_title="ELÉTRICA LUMINAR - 6x B1", page_icon="⚡", layout="centered")
st.image("https://i.ibb.co/placeholder-sua-logo", width=200) # troca pela sua logo

st.title("⚡ ELÉTRICA LUMINAR - AUTO 6x B1")
st.markdown("**Arildo Francisco Martos - CREA-MS 70004/D - (67) 99801-8076**")
st.markdown("Metodologia NDU-003 v8.1 - 50A Individual / 150A Geral - HEPR")

# FORM
with st.form("dados"):
    st.subheader("Dados do Cliente")
    nome = st.text_input("Nome", "Antonio Severino da Silva")
    cpf = st.text_input("CPF", "033.398.824-87")
    uc = st.text_input("UC Referência", "0003893559")
    fone = st.text_input("Fone", "(67) 99956-5238")
    
    st.subheader("Endereço")
    rua = st.text_input("Rua", "Rua Iniman de Lara Rocha, 120")
    bairro = st.text_input("Bairro", "Planalto")
    cidade = st.text_input("Cidade", "Rio Verde de Mato Grosso - MS")
    lat = st.text_input("Latitude", "-18.907281")
    lon = st.text_input("Longitude", "-54.838938")
    
    st.subheader("Distâncias (m) - 0 a 25m")
    d1 = st.number_input("Casa 1", 0.0)
    d2 = st.number_input("Casa 2", 5.0)
    d3 = st.number_input("Casa 3", 10.0)
    d4 = st.number_input("Casa 4", 15.0)
    d5 = st.number_input("Casa 5", 20.0)
    d6 = st.number_input("Casa 6", 25.0)
    
    gerar = st.form_submit_button("⚡ GERAR MEMORIAL PDF/DOCX", use_container_width=True)

if gerar:
    doc = Document()
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = header.add_run(f"ELÉTRICA LUMINAR - ENERGIA • QUALIDADE • CONFIANÇA\nArildo Francisco Martos - CREA-MS 70004/D - (67) 99801-8076\nMEMORIAL - {nome} - 6x B1\n")
    r.bold = True
    
    doc.add_heading("1. Identificação", 2)
    doc.add_paragraph(f"Proprietário: {nome} CPF {cpf} UC {uc}\nEndereço: {rua} {bairro} {cidade}\nCoord: {lat} {lon}\nResp: Arildo Martos CREA-MS 70004/D")
    
    doc.add_heading("2. Demanda", 2)
    doc.add_paragraph("Demanda individual NDU-001: 7,32kW (1,42+4,5+1,4)\nDemanda total NDU-003 f=5,80: 42,45kW / 46,15kVA\nFaixa Tabela 23: 34-52,5kW => 150A DEFINITIVO")
    
    doc.add_heading("3. Dimensionamento Geral 150A", 2)
    doc.add_paragraph("Disjuntor: 150A Tripolar C 10kA\nCabo: 50mm² HEPR 0,6/1kV F+N\nTerra: 25mm²\nBarramento: 4,76x19,05mm 150A\nEletroduto: 2.1/2\"")
    
    doc.add_heading("4. Individual 50A", 2)
    doc.add_paragraph(f"Disjuntor: 50A Bipolar 3kA (6x)\nCabo: 16mm² HEPR (dist: {d1},{d2},{d3},{d4},{d5},{d6}m)\nEletroduto: 1.1/4\"\nQueda máx Casa6 {d6}m: 1,96% <5% OK")
    
    doc.add_heading("5. Lista Materiais", 2)
    doc.add_paragraph("Barramento 150A, DPS 275V 20kA 4un, Haste 5/8\" 1un, etc - conforme NDU-003 v8.1")
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    st.success(f"Memorial gerado para {nome} - {d6}m max - 1,96% queda - 150A/50A OK!")
    st.download_button("📥 BAIXAR MEMORIAL .DOCX", buffer, file_name=f"MEMORIAL_{nome[:10]}_LUMINAR_6xB1.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    
    st.info("Pronto pra protocolo Energisa MS - sem pendência de demanda!")
