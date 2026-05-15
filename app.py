import streamlit as st
import datetime, io, smtplib
from email.mime.text import MIMEText
from gerar_pauta import gerar_pauta

st.set_page_config(
    page_title="Pauta de Audiências — Controladoria Jurídica",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #F4F6F9; }
.header-box {
    background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
    color: white; padding: 24px 32px; border-radius: 12px;
    margin-bottom: 24px;
}
.header-box h1 { margin: 0; font-size: 1.8rem; font-weight: 700; }
.header-box p  { margin: 4px 0 0; opacity: .8; font-size: .95rem; }
.upload-card {
    background: white; border-radius: 10px; padding: 20px 24px;
    border: 2px dashed #CBD5E0; box-shadow: 0 2px 6px rgba(0,0,0,.06);
}
.upload-card.filled { border-color: #2E75B6; border-style: solid; }
.upload-card h3 { margin: 0 0 4px; font-size: 1rem; color: #1F3864; }
.upload-card p  { margin: 0 0 12px; color: #718096; font-size: .85rem; }
.metric-card {
    background: white; border-radius: 8px; padding: 16px;
    text-align: center; box-shadow: 0 2px 6px rgba(0,0,0,.06);
}
.metric-card .num { font-size: 2.2rem; font-weight: 700; color: #1F3864; }
.metric-card .lbl { font-size: .75rem; color: #718096; text-transform: uppercase; letter-spacing: .5px; }
.alert-div {
    background: #FFF5F5; border: 2px solid #C53030;
    border-radius: 8px; padding: 16px 20px; margin: 16px 0;
}
.alert-div h3 { color: #C53030; margin: 0 0 8px; }
.alert-ok {
    background: #F0FFF4; border: 2px solid #276749;
    border-radius: 8px; padding: 12px 20px; margin: 16px 0;
}
.alert-ok p { color: #276749; margin: 0; font-weight: 600; }
.div-table { width: 100%; border-collapse: collapse; font-size: .85rem; }
.div-table th { background: #1F3864; color: white; padding: 8px 12px; text-align: left; }
.div-table td { padding: 8px 12px; border-bottom: 1px solid #E2E8F0; }
.div-table tr.alta  td { background: #FFF5F5; }
.div-table tr.media td { background: #FFFFF0; }
.div-table tr.baixa td { background: #F0FFF4; }
.btn-gerar > button {
    background: linear-gradient(135deg, #1F3864, #2E75B6) !important;
    color: white !important; font-weight: 700 !important;
    font-size: 1.1rem !important; padding: 14px 40px !important;
    border-radius: 8px !important; border: none !important;
    width: 100% !important; cursor: pointer !important;
}
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-box">
  <h1>⚖️ Pauta de Audiências — Controladoria Jurídica</h1>
  <p>Geração automática do relatório semanal com preservação de dados e detecção de divergências</p>
</div>
""", unsafe_allow_html=True)

# ── Uploads ────────────────────────────────────────────────────────────────────
st.markdown("### 📂 Arquivos de Entrada")
col1, col2 = st.columns(2, gap="large")

with col1:
    st.markdown("""
    <div class="upload-card">
      <h3>📋 Relatório Anterior</h3>
      <p>Arquivo da semana passada — dados preenchidos pelos advogados e Controladoria serão preservados</p>
    </div>
    """, unsafe_allow_html=True)
    file_old = st.file_uploader(
        "Selecione o arquivo da semana anterior (.xlsx)",
        type=['xlsx'],
        key='file_old',
        label_visibility='collapsed',
    )
    if file_old:
        st.success(f"✅ {file_old.name}")

with col2:
    st.markdown("""
    <div class="upload-card">
      <h3>📊 Exportação CJ (Base)</h3>
      <p>Arquivo exportado do sistema — CJ REUNIÃO DE PAUTA ATUALIZADA — será a base do novo relatório</p>
    </div>
    """, unsafe_allow_html=True)
    file_new = st.file_uploader(
        "Selecione a exportação CJ (.xlsx)",
        type=['xlsx'],
        key='file_new',
        label_visibility='collapsed',
    )
    if file_new:
        st.success(f"✅ {file_new.name}")

st.markdown("---")

# ── Botão Gerar ────────────────────────────────────────────────────────────────
if not file_new:
    st.warning("⚠️ Selecione ao menos o arquivo de **Exportação CJ** para gerar o relatório.")
else:
    col_btn, col_info = st.columns([1,2])
    with col_btn:
        gerar = st.button("⚡ Gerar Relatório", use_container_width=True, type="primary")

    if gerar or st.session_state.get('resultado'):
        with st.spinner("Processando dados e gerando planilha..."):
            try:
                src_new_bytes = file_new.read()
                src_old_bytes = file_old.read() if file_old else None

                output_bytes, resumo, divs = gerar_pauta(src_new_bytes, src_old_bytes)
                st.session_state['resultado'] = (output_bytes, resumo, divs)

            except Exception as e:
                st.error(f"❌ Erro ao gerar: {e}")
                st.stop()

        output_bytes, resumo, divs = st.session_state['resultado']

        # ── Métricas ────────────────────────────────────────────────────────────
        st.markdown("### 📊 Resumo")
        m1,m2,m3,m4,m5,m6 = st.columns(6)
        def metric(col, num, lbl, color="#1F3864"):
            col.markdown(f"""
            <div class="metric-card">
              <div class="num" style="color:{color}">{num}</div>
              <div class="lbl">{lbl}</div>
            </div>""", unsafe_allow_html=True)

        metric(m1, resumo['total_pendentes'],  'Audiências Pendentes')
        metric(m2, resumo['total_canceladas'], 'Canceladas')
        metric(m3, resumo['ctrl_pendentes'],   'Precisam Corresp.', '#7D4300')
        metric(m4, resumo['preservados_chave'],'Dados Preservados',  '#145A32')
        metric(m5, resumo['divergencias'],     'Divergências',
               '#C53030' if resumo['divergencias']>0 else '#145A32')
        metric(m6, f"{resumo['s1_start']}", f"S1: até {resumo['s1_end']}")

        st.caption(f"Semana 1: {resumo['s1_start']} → {resumo['s1_end']}  |  "
                   f"Prévia S2: {resumo['s2_start']} → {resumo['s2_end']}  |  "
                   f"Gerado em: {resumo['gerado_em']}")

        # ── Alertas de divergência ───────────────────────────────────────────────
        st.markdown("### 🔍 Verificação de Divergências")
        if divs:
            n_alta  = sum(1 for d in divs if d['gravidade']=='ALTA')
            n_media = sum(1 for d in divs if d['gravidade']=='MÉDIA')
            n_baixa = sum(1 for d in divs if d['gravidade']=='BAIXA')

            st.markdown(f"""
            <div class="alert-div">
              <h3>⚠️ {len(divs)} divergência(s) encontrada(s) — verifique antes de distribuir</h3>
              <p>🔴 Alta: {n_alta} &nbsp;|&nbsp; 🟡 Média: {n_media} &nbsp;|&nbsp; 🟢 Baixa: {n_baixa}</p>
            </div>
            """, unsafe_allow_html=True)

            # Tabela de divergências
            rows_html=""
            for d in divs:
                cls=d['gravidade'].lower()
                grav_icon={'ALTA':'🔴','MÉDIA':'🟡','BAIXA':'🟢'}.get(d['gravidade'],'')
                rows_html+=f"""<tr class="{cls}">
                  <td><b>{d['tipo']}</b></td>
                  <td>{d['campo']}</td>
                  <td><code>{d['valor']}</code></td>
                  <td>{d.get('valor_sugerido','') or '—'}</td>
                  <td>{grav_icon} {d['gravidade']}</td>
                  <td>{d['descricao']}</td>
                  <td style="text-align:center"><b>{d['ocorrencias']}</b></td>
                </tr>"""

            st.markdown(f"""
            <table class="div-table">
              <tr>
                <th>TIPO</th><th>CAMPO</th><th>VALOR ENCONTRADO</th>
                <th>SUGERIDO</th><th>GRAVIDADE</th><th>DESCRIÇÃO</th><th>QTD</th>
              </tr>
              {rows_html}
            </table>
            """, unsafe_allow_html=True)

            # Alerta para gestor
            st.markdown("---")
            st.markdown("#### 📧 Notificar Gestor do Painel")
            with st.expander("Configurar notificação por e-mail"):
                email_gestor = st.text_input("E-mail do gestor", placeholder="gestor@escritorio.com.br")
                msg_extra = st.text_area("Observação adicional (opcional)", height=80)
                if st.button("📨 Enviar alerta ao gestor"):
                    if email_gestor:
                        div_lines="\n".join(
                            f"  • [{d['gravidade']}] {d['tipo']}: {d['valor']} ({d['ocorrencias']}x)" for d in divs
                        )
                        corpo=f"""Prezado(a) gestor(a),

O relatório de Pauta de Audiências gerado em {resumo['gerado_em']} apresenta {len(divs)} divergência(s) que requer(em) atenção antes da distribuição.

DIVERGÊNCIAS ENCONTRADAS:
{div_lines}

RESUMO DO RELATÓRIO:
• Total pendentes: {resumo['total_pendentes']}
• Semana 1: {resumo['s1_start']} a {resumo['s1_end']}
• Correspondentes necessários: {resumo['ctrl_pendentes']}

{('OBSERVAÇÃO: ' + msg_extra) if msg_extra else ''}

Por favor, acesse o painel para verificar e corrigir antes de distribuir o arquivo.

— Sistema de Pauta de Audiências · Controladoria Jurídica
"""
                        st.code(corpo, language=None)
                        st.info("📋 Copie o texto acima e envie manualmente, ou configure as credenciais SMTP para envio automático.")
                    else:
                        st.warning("Informe o e-mail do gestor.")
        else:
            st.markdown("""
            <div class="alert-ok">
              <p>✅ Nenhuma divergência detectada. O arquivo está pronto para distribuição.</p>
            </div>
            """, unsafe_allow_html=True)

        # ── Download ─────────────────────────────────────────────────────────────
        st.markdown("---")
        st.markdown("### 📥 Download")
        nome_arquivo = f"REUNIAO_PAUTA_{datetime.date.today().strftime('%d_%m_%Y')}.xlsx"

        col_dl, col_info2 = st.columns([1,2])
        with col_dl:
            st.download_button(
                label="⬇️ Baixar Relatório",
                data=output_bytes,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary" if not divs else "secondary",
            )
        with col_info2:
            if divs:
                st.warning(f"⚠️ O arquivo contém a aba **DIVERGÊNCIAS** com {len(divs)} item(s) para revisão.")
            else:
                st.info(f"📄 Arquivo: `{nome_arquivo}`")

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Controladoria Jurídica · Imaculada Gordiano Advogados · Sistema de Pauta de Audiências")
