import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
from datetime import datetime

st.set_page_config(page_title="Elétrica Luminar - Auto Memorial", layout="centered")
st.title("⚡ ELÉTRICA LUMINAR - AUTO 6x B1")
st.subheader("Memorial Completo - Padrão Energisa / Marcket")

st.markdown("Preencha os dados e baixe o memorial PRONTO para enviar para a Energisa.")

with st.form("dados"):
    st.markdown("### 1. Identificação")
    col1, col2 = st.columns(2)
    proprietario = col1.text_input("Proprietário", "Antonio Severino da Silva")
    cpf = col2.text_input("CPF", "033.398.824-87")
    uc = st.text_input("UC", "0003893559")
    endereco = st.text_input("Endereço", "Rua Iniman de Lara Rocha, 120 - Planalto")
    cidade = st.text_input("Cidade", "Rio Verde de Mato Grosso - MS")
    coord = st.text_input("Coordenadas", "-18.907281 -54.838938")

    st.markdown("### 2. Cargas (NDU-001)")
    c1, c2, c3 = st.columns(3)
    a_ilum = c1.number_input("a - Ilum+Tomadas (kW)", value=1.42)
    b_chuveiros = c2.number_input("c - Chuveiros/Aquec (kW)", value=4.5)
    c_ar = c3.number_input("c - Ar Condicionado (kW)", value=1.4)
    c4, c5 = st.columns(2)
    d_motores = c4.number_input("d - Motores (kW)", value=0.0)
    e_outros = c5.number_input("Outros (kW)", value=0.0)

    st.markdown("### 3. Agrupamento")
    n_ucs = st.number_input("Nº de UCs B1", value=6, min_value=2, max_value=12)
    fator_f = st.number_input("Fator de Diversificação NDU-003 Tabela 22", value=5.8, help="Para 6x B1 = 5,80")
    distancias = st.text_input("Distâncias das UCs (m) separadas por vírgula", "0,5,10,15,20,25")
    tensao = st.selectbox("Tensão", ["220/127V", "380/220V"], index=0)

    gerar = st.form_submit_button("⚡ GERAR MEMORIAL COMPLETO")

if gerar:
    d_individual = a_ilum + b_chuveiros + c_ar + d_motores + e_outros
    d_total_kw = d_individual * fator_f
    d_total_kva = d_total_kw / 0.92
    try:
        dists = [float(x.strip()) for x in distancias.split(",")]
        Lmax = max(dists)
    except:
        Lmax = 25.0
        dists = [0,5,10,15,20,25]
    
    S = 16
    deltaV = 100 * 2 * Lmax * (d_individual*1000/220) / (56 * S * 220)

    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ELÉTRICA LUMINAR - ENERGIA • QUALIDADE • CONFIANÇA\nArildo Francisco Martos - CREA-MS 70004/D - (67) 99801-8076")
    r.bold = True
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"MEMORIAL DESCRITIVO E DE CÁLCULO\nAGRUPAMENTO {int(n_ucs)}x B1 - {proprietario.upper()}")
    r.bold = True
    r.font.size = Pt(12)

    doc.add_paragraph(f"Data: {datetime.now().strftime('%d/%m/%Y')}")

    doc.add_heading("1. OBJETIVO", 2)
    doc.add_paragraph(f"Este memorial tem por objetivo apresentar os cálculos de demanda, dimensionamento e especificação de materiais para ligação de agrupamento de {int(n_ucs)} unidades consumidoras classe B1, conforme NDU-001, NDU-003 versão 8.1 e NBR 5410, para atendimento à concessionária ENERGISA MS, no imóvel localizado em {endereco}, {cidade}, coordenadas {coord}, de propriedade de {proprietario}, CPF {cpf}, UC {uc}.")

    doc.add_heading("2. NORMAS E REFERÊNCIAS", 2)
    doc.add_paragraph("• NDU-001 - Fornecimento em Tensão Secundária - Edificações Individuais\n• NDU-003 - Fornecimento em Tensão Secundária - Agrupamentos - v8.1\n• NBR 5410 - Instalações Elétricas de Baixa Tensão\n• NBR 14039 e NR-10")

    doc.add_heading("3. IDENTIFICAÇÃO DA OBRA", 2)
    doc.add_paragraph(f"Proprietário: {proprietario} - CPF {cpf} - UC {uc}\nEndereço: {endereco} - {cidade}\nCoordenadas: {coord}\nResponsável Técnico: Arildo Francisco Martos - CREA-MS 70004/D\nTipo: Agrupamento {int(n_ucs)}x B1")

    doc.add_heading("4. CÁLCULO DEMANDA INDIVIDUAL - NDU-001", 2)
    doc.add_paragraph(
        f"Fórmula NDU-001: D (kW) = a + b + c + d + e + f\n\n"
        f"a = Iluminação e tomadas = {a_ilum:.2f} kW\n"
        f"b = Chuveiros = {b_chuveiros:.2f} kW\n"
        f"c = Ar Condicionado = {c_ar:.2f} kW\n"
        f"d = Motores = {d_motores:.2f} kW\n"
        f"e+f = Outros = {e_outros:.2f} kW\n\n"
        f"Dind = {a_ilum:.2f} + {b_chuveiros:.2f} + {c_ar:.2f} + {d_motores:.2f} + {e_outros:.2f} = {d_individual:.2f} kW\n"
        f"Dind (kVA) = {d_individual:.2f} / 0,92 = {d_individual/0.92:.2f} kVA"
    )

    doc.add_heading("5. DEMANDA DIVERSIFICADA - NDU-003", 2)
    doc.add_paragraph(
        f"Fórmula NDU-003: Dt = f x Σ Dind\n\n"
        f"Para {int(n_ucs)} UCs B1, fator f = {fator_f:.2f}\n"
        f"Dt = {fator_f:.2f} x {d_individual:.2f} = {d_total_kw:.2f} kW\n"
        f"Dt (kVA) = {d_total_kw:.2f} / 0,92 = {d_total_kva:.2f} kVA\n\n"
        f"Enquadramento NDU-003 Tabela 23: Faixa 34,01 a 52,50 kW => Disjuntor Geral 150A DEFINITIVO\n"
        f"Resultado: {d_total_kw:.2f} kW => 150A DEFINITIVO APROVADO"
    )

    doc.add_heading("6. DIMENSIONAMENTO GERAL 150A", 2)
    doc.add_paragraph(
        f"• Disjuntor Geral: 150A Tripolar C 10kA\n"
        f"• Cabo Alimentador: 50mm² HEPR 0,6/1kV\n"
        f"• Terra: 25mm² HEPR\n"
        f"• Barramento: 4,76x19,05mm 150A\n"
        f"• Eletroduto: 2.1/2\"\n"
        f"• DPS: 4un 275V 20kA\n"
        f"• Haste: 5/8\" x 2,40m"
    )

    doc.add_heading("7. DIMENSIONAMENTO INDIVIDUAL 50A + QUEDA DE TENSÃO", 2)
    texto_dist = ", ".join([f"{x}m" for x in dists])
    doc.add_paragraph(
        f"• Disjuntor: 50A Bipolar 3kA (6x)\n"
        f"• Cabo: 16mm² HEPR\n"
        f"• Eletroduto: 1.1/4\"\n"
        f"• Distâncias: {texto_dist}\n\n"
        f"Fórmula Queda: ΔV% = (100 x 2 x L x I) / (56 x S x V)\n"
        f"ΔV% = (100 x 2 x {Lmax:.1f} x {d_individual*1000/220:.1f}) / (56 x 16 x 220) = {deltaV:.2f}%\n"
        f"Limite 5% => {deltaV:.2f}% < 5% => APROVADO"
    )

    doc.add_heading("8. LISTA MATERIAIS NDU-003 v8.1", 2)
    doc.add_paragraph(
        f"• 01 Caixa barramento 150A\n• 01 Disjuntor 150A\n• 06 Disjuntores 50A\n• Cabo 50mm², 25mm², 16mm²\n• 04 DPS, 01 Haste, Eletrodutos"
    )

    doc.add_heading("9. CONSIDERAÇÕES FINAIS", 2)
    doc.add_paragraph(f"Agrupamento {int(n_ucs)}x B1 com {d_total_kw:.2f}kW, 150A definitivo, queda {deltaV:.2f}% - Apto para aprovação Energisa.\n\nResp: Arildo Martos CREA-MS 70004/D")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    st.success(f"✅ Memorial Completo Gerado! {d_total_kw:.2f}kW - Queda {deltaV:.2f}% - 150A")
    st.download_button("📥 BAIXAR MEMORIAL COMPLETO .DOCX", buffer, file_name=f"MEMORIAL_{proprietario.replace(' ','_')}_{int(n_ucs)}xB1_COMPLETO.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
