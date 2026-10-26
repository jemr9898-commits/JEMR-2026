import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io

# Impresiones y PDFs con ReportLab
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Configuración de la página web
st.set_page_config(page_title="ESTADO DE CUENTA", layout="wide")

file_path = "ESTADO DE CUENTA DETALLADO.xlsx"

# ==========================================
# MÓDULO DE GESTIÓN DE USUARIOS EN EXCEL
# ==========================================
def cargar_usuarios():
    try:
        df_usr = pd.read_excel(file_path, sheet_name="USUARIOS")
        df_usr = df_usr.dropna(subset=['USUARIO'])
        df_usr['USUARIO'] = df_usr['USUARIO'].astype(str).str.strip().str.lower()
        df_usr['ROL'] = df_usr['ROL'].astype(str).str.strip().str.upper()
        return df_usr
    except Exception:
        # Crear hoja por defecto si no existe
        df_def = pd.DataFrame([
            {"USUARIO": "admin", "PASSWORD": "123", "ROL": "ADMINISTRADOR", "NOMBRE": "ADMINISTRADOR DEL SISTEMA"},
            {"USUARIO": "jesus", "PASSWORD": "123", "ROL": "USUARIO", "NOMBRE": "JESUS MARQUEZ"}
        ])
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_def.to_excel(writer, sheet_name='USUARIOS', index=False)
        except Exception:
            pass
        return df_def

def registrar_auditoria(usuario, accion, detalle):
    try:
        try:
            df_log = pd.read_excel(file_path, sheet_name="AUDITORIA")
        except Exception:
            df_log = pd.DataFrame(columns=["FECHA_HORA", "USUARIO", "ACCION", "DETALLE"])

        nuevo_log = {
            "FECHA_HORA": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "USUARIO": str(usuario).upper(),
            "ACCION": str(accion).upper(),
            "DETALLE": str(detalle).upper()
        }
        df_log = pd.concat([df_log, pd.DataFrame([nuevo_log])], ignore_index=True)

        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_log.to_excel(writer, sheet_name='AUDITORIA', index=False)
    except Exception as e:
        st.error(f"ERROR REGISTRANDO AUDITORÍA: {e}")

# Inicializar sesión
if "AUTENTICADO" not in st.session_state:
    st.session_state["AUTENTICADO"] = False
if "USUARIO_ACTUAL" not in st.session_state:
    st.session_state["USUARIO_ACTUAL"] = ""
if "ROL_ACTUAL" not in st.session_state:
    st.session_state["ROL_ACTUAL"] = ""
if "NOMBRE_ACTUAL" not in st.session_state:
    st.session_state["NOMBRE_ACTUAL"] = ""

def iniciar_sesion(user_input, pass_input):
    user_clean = str(user_input).strip().lower()
    df_usr = cargar_usuarios()
    row = df_usr[df_usr['USUARIO'] == user_clean]
    
    if not row.empty:
        pass_real = str(row.iloc[0]['PASSWORD']).strip()
        if str(pass_input).strip() == pass_real:
            st.session_state["AUTENTICADO"] = True
            st.session_state["USUARIO_ACTUAL"] = user_clean.upper()
            st.session_state["ROL_ACTUAL"] = str(row.iloc[0]['ROL']).upper()
            st.session_state["NOMBRE_ACTUAL"] = str(row.iloc[0]['NOMBRE']).upper()
            registrar_auditoria(st.session_state["USUARIO_ACTUAL"], "INICIO DE SESIÓN", "ACCESO EXITOSO AL SISTEMA")
            st.success("¡INICIO DE SESIÓN EXITOSO!")
            st.rerun()
            return
            
    st.error("USUARIO O CONTRASEÑA INCORRECTOS")

def cerrar_sesion():
    if st.session_state["AUTENTICADO"]:
        registrar_auditoria(st.session_state["USUARIO_ACTUAL"], "CIERRE DE SESIÓN", "SALIDA DEL SISTEMA")
    st.session_state["AUTENTICADO"] = False
    st.session_state["USUARIO_ACTUAL"] = ""
    st.session_state["ROL_ACTUAL"] = ""
    st.session_state["NOMBRE_ACTUAL"] = ""
    st.rerun()

# PANTALLA DE LOGIN
if not st.session_state["AUTENTICADO"]:
    st.title("🔒 CONTROL DE ACCESO — ESTADO DE CUENTA")
    st.markdown("### INGRESE SUS CREDENCIALES PARA ACCEDER AL SISTEMA")
    
    col_login, _ = st.columns([1, 1])
    with col_login:
        with st.form("FORM_LOGIN"):
            usr = st.text_input("USUARIO")
            pwd = st.text_input("CONTRASEÑA", type="password")
            btn_login = st.form_submit_button("INGRESAR AL SISTEMA")
            if btn_login:
                iniciar_sesion(usr, pwd)
    st.stop()

# ==========================================
# BARRA SUPERIOR E IDENTIFICACIÓN DEL USUARIO
# ==========================================
col_header_title, col_logout = st.columns([4, 1])
with col_header_title:
    st.title("💼 ESTADO DE CUENTA")
    st.info(f"👤 USUARIO CONECTADO: **{st.session_state['NOMBRE_ACTUAL']}** | USUARIO: **{st.session_state['USUARIO_ACTUAL']}** | ROL: **{st.session_state['ROL_ACTUAL']}**")
with col_logout:
    st.write("")
    st.write("")
    if st.button("🚪 CERRAR SESIÓN"):
        cerrar_sesion()

st.markdown("---")

# FUNCIONES AUXILIARES DE FORMATO Y CÁLCULO
def formato_moneda(val, decimales=2):
    if pd.isna(val) or val == "" or val is None or str(val).lower() == 'nan':
        return "0,00" if decimales == 2 else "0,0000"
    try:
        val_float = float(val)
        fmt = f"{{:,.{decimales}f}}"
        return fmt.format(val_float).replace(',', 'X').replace('.', ',').replace('X', '.')
    except Exception:
        return str(val)

def formato_fecha(val):
    if pd.isna(val) or val == "" or val is None or str(val).lower() == 'nan':
        return ""
    try:
        return pd.to_datetime(val).strftime('%Y-%m-%d')
    except Exception:
        return str(val)[:10]

def calcular_fecha_pp(f_recepcion, dias_pp):
    if pd.isna(dias_pp) or dias_pp is None or str(dias_pp).lower() == 'nan':
        return "N/A"
    try:
        dias = int(float(dias_pp))
        if dias > 0 and pd.notna(f_recepcion):
            f_lim = pd.to_datetime(f_recepcion) + pd.Timedelta(days=dias)
            return f_lim.strftime('%Y-%m-%d')
        return "N/A"
    except Exception:
        return "N/A"

def df_a_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='EXPORTADO')
    return output.getvalue()

def generar_pdf_comprobante(row_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=18, leading=22, textColor=colors.HexColor('#1E3A8A'), alignment=1, spaceAfter=8)
    subtitle_style = ParagraphStyle('SubtitleStyle', parent=styles['Normal'], fontSize=11, leading=14, textColor=colors.HexColor('#4B5563'), alignment=1, spaceAfter=15)
    label_style = ParagraphStyle('LabelStyle', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor('#374151'), fontName='Helvetica-Bold')
    val_style = ParagraphStyle('ValStyle', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor('#111827'))

    story.append(Paragraph("<b>COMPROBANTE DE REGISTRO Y RECEPCIÓN DE DOCUMENTO</b>", title_style))
    story.append(Paragraph("SISTEMA DE GESTIÓN DE ESTADO DE CUENTA — CONTROL FÍSICO", subtitle_style))
    story.append(HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#1E3A8A'), spaceAfter=15))

    tipo_doc_nombre = "FACTURA (FC)" if str(row_data.get('TIPO')) == 'FC' else "NOTA DE CRÉDITO (NC)"
    
    data_general = [
        [Paragraph("N° REGISTRO SISTEMA:", label_style), Paragraph(str(row_data.get('N°', '')), val_style),
         Paragraph("FECHA REGISTRO:", label_style), Paragraph(datetime.now().strftime('%Y-%m-%d %H:%M'), val_style)],
        [Paragraph("TIPO DE DOCUMENTO:", label_style), Paragraph(tipo_doc_nombre, val_style),
         Paragraph("NÚMERO DE DOCUMENTO:", label_style), Paragraph(str(row_data.get('FACTURA', '')), val_style)],
        [Paragraph("PROVEEDOR:", label_style), Paragraph(str(row_data.get('PROVEEDOR', '')), val_style),
         Paragraph("OCD ICG:", label_style), Paragraph(str(row_data.get('OCD ICG', '') or 'N/A'), val_style)],
        [Paragraph("FECHA EMISIÓN:", label_style), Paragraph(formato_fecha(row_data.get('F FACTURA', '')), val_style),
         Paragraph("FECHA RECEPCIÓN:", label_style), Paragraph(formato_fecha(row_data.get('F RECEPCION', '')), val_style)],
        [Paragraph("DÍAS DE CRÉDITO:", label_style), Paragraph(f"{int(float(row_data.get('CREDITO', 0)))} DÍAS", val_style),
         Paragraph("FECHA VENCIMIENTO:", label_style), Paragraph(formato_fecha(row_data.get('VENCIMIENTO', '')), val_style)]
    ]

    t_gen = Table(data_general, colWidths=[120, 150, 120, 150])
    t_gen.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F9FAFB')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#E5E7EB')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#F3F4F6')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_gen)
    story.append(Spacer(1, 15))

    story.append(Paragraph("<b>PLAZOS DE PRONTO PAGO (PP)</b>", label_style))
    f_rec = row_data.get('F RECEPCION')
    pp1_v = 0 if pd.isna(row_data.get('PP1')) else row_data.get('PP1')
    pp2_v = 0 if pd.isna(row_data.get('PP2')) else row_data.get('PP2')
    pp3_v = 0 if pd.isna(row_data.get('PP3')) else row_data.get('PP3')
    
    pp1_f = calcular_fecha_pp(f_rec, pp1_v)
    pp2_f = calcular_fecha_pp(f_rec, pp2_v)
    pp3_f = calcular_fecha_pp(f_rec, pp3_v)

    data_pp = [
        [Paragraph("<b>PRONTO PAGO 1 (PP1)</b>", label_style), Paragraph(f"{int(float(pp1_v if str(pp1_v).isdigit() or type(pp1_v) in [int, float] else 0))} DÍAS", val_style), Paragraph(f"FECHA LÍMITE: {pp1_f}", val_style)],
        [Paragraph("<b>PRONTO PAGO 2 (PP2)</b>", label_style), Paragraph(f"{int(float(pp2_v if str(pp2_v).isdigit() or type(pp2_v) in [int, float] else 0))} DÍAS", val_style), Paragraph(f"FECHA LÍMITE: {pp2_f}", val_style)],
        [Paragraph("<b>PRONTO PAGO 3 (PP3)</b>", label_style), Paragraph(f"{int(float(pp3_v if str(pp3_v).isdigit() or type(pp3_v) in [int, float] else 0))} DÍAS", val_style), Paragraph(f"FECHA LÍMITE: {pp3_f}", val_style)]
    ]
    t_pp = Table(data_pp, colWidths=[180, 180, 180])
    t_pp.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_pp)
    story.append(Spacer(1, 15))

    story.append(Paragraph("<b>DETALLE FINANCIERO Y LIQUIDACIÓN</b>", label_style))
    
    tot_usd = float(row_data.get('TOTAL', 0)) if pd.notna(row_data.get('TOTAL')) else 0.0
    tasa_val = float(row_data.get('TASA', 0)) if pd.notna(row_data.get('TASA')) else 0.0
    tot_bs_calc = tot_usd * tasa_val

    data_fin = [
        [Paragraph("<b>CONCEPTO</b>", label_style), Paragraph("<b>MONTO USD ($)</b>", label_style)],
        [Paragraph("SUB-TOTAL", val_style), Paragraph(formato_moneda(row_data.get('SUB-TOTAL', 0)), val_style)],
        [Paragraph("MONTO EXENTO", val_style), Paragraph(formato_moneda(row_data.get('EXENTO', 0)), val_style)],
        [Paragraph("BASE IMPONIBLE", val_style), Paragraph(formato_moneda(row_data.get('BASE', 0)), val_style)],
        [Paragraph("IVA (16%)", val_style), Paragraph(formato_moneda(row_data.get('IVA', 0)), val_style)],
        [Paragraph("IVA NO RETENIDO (25% DEL IVA)", val_style), Paragraph(formato_moneda(row_data.get('IVA NR', 0)), val_style)],
        [Paragraph("<b>TOTAL GENERAL (USD)</b>", label_style), Paragraph(f"<b>$ {formato_moneda(tot_usd)}</b>", label_style)],
        [Paragraph("TASA DE CAMBIO REFERENCIAL", val_style), Paragraph(f"BS. {formato_moneda(tasa_val, decimales=4)}", val_style)],
        [Paragraph("<b>TOTAL EQUIVALENTE EN BOLÍVARES (BS)</b>", label_style), Paragraph(f"<b>BS. {formato_moneda(tot_bs_calc)}</b>", label_style)]
    ]
    t_fin = Table(data_fin, colWidths=[320, 220])
    t_fin.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#E5E7EB')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#9CA3AF')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E5E7EB')),
        ('BACKGROUND', (0,6), (-1,6), colors.HexColor('#FEF3C7')),
        ('BACKGROUND', (0,8), (-1,8), colors.HexColor('#E0E7FF')),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_fin)
    story.append(Spacer(1, 40))

    data_firmas = [
        [Paragraph("__________________________________<br/><b>RECIBIDO POR (FIRMA / SELLO)</b>", val_style),
         Paragraph("__________________________________<br/><b>REVISADO / APROBADO (ADMINISTRACIÓN)</b>", val_style)]
    ]
    t_firmas = Table(data_firmas, colWidths=[270, 270])
    t_firmas.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.append(t_firmas)

    doc.build(story)
    return buffer.getvalue()

def cargar_datos():
    try:
        df_datos = pd.read_excel(file_path, sheet_name="DATOS")
        df_datos = df_datos.dropna(subset=['N°', 'PROVEEDOR'], how='all')
        numeric_cols = ['CREDITO', 'DIAS', 'TASA', 'SUB-TOTAL', 'EXENTO', 'BASE', 'IVA', 'IVA NR', 'TOTAL', 'TOTAL BS']
        for col in numeric_cols:
            if col in df_datos.columns:
                df_datos[col] = pd.to_numeric(df_datos[col], errors='coerce').fillna(0)
        
        if 'FACTURA' in df_datos.columns:
            df_datos['FACTURA'] = df_datos['FACTURA'].astype(str).str.replace('.0', '', regex=False)
        if 'N°' in df_datos.columns:
            df_datos['N°'] = df_datos['N°'].astype(str).str.replace('.0', '', regex=False)
            
        return df_datos
    except Exception:
        cols_def = ['N°', 'OCD ICG', 'PROVEEDOR', 'TIPO', 'FACTURA', 'F FACTURA', 'F RECEPCION', 'CREDITO', 'PP1', 'PP2', 'PP3', 'VENCIMIENTO', 'DIAS', 'TASA', 'SUB-TOTAL', 'EXENTO', 'BASE', 'IVA', 'IVA NR', 'TOTAL', 'TOTAL BS', 'ESTADO']
        return pd.DataFrame(columns=cols_def)

def cargar_condiciones():
    try:
        df_cond = pd.read_excel(file_path, sheet_name="CONDICIONES")
        return df_cond.dropna(how='all')
    except Exception:
        return pd.DataFrame(columns=['Proveedores', 'Factura digital', 'Condicion', 'Departamento', 'Nombre', 'Correo', 'Numero de telefono', 'comentarios'])

try:
    df_datos = cargar_datos()
    df_condiciones = cargar_condiciones()

    lista_proveedores_datos = list(df_datos['PROVEEDOR'].dropna().unique()) if 'PROVEEDOR' in df_datos.columns else []
    lista_proveedores_cond = list(df_condiciones['Proveedores'].dropna().unique()) if 'Proveedores' in df_condiciones.columns else []
    todos_proveedores = sorted(list(set(lista_proveedores_datos + lista_proveedores_cond)))

    titulos_pestanas = [
        "📊 DASHBOARD & REPORTES", 
        "📝 REGISTRAR FACTURA", 
        "🏢 GESTIÓN DE PROVEEDORES", 
        "📋 CONDICIONES COMERCIALES"
    ]
    if st.session_state["ROL_ACTUAL"] == "ADMINISTRADOR":
        titulos_pestanas.extend(["👥 GESTIÓN DE USUARIOS", "🗑️ GESTIÓN Y BORRADO DE REGISTROS", "📜 AUDITORÍA Y MODIFICACIONES"])

    pestanas = st.tabs(titulos_pestanas)
    
    tab_dashboard = pestanas[0]
    tab_nueva_factura = pestanas[1]
    tab_proveedores = pestanas[2]
    tab_condiciones = pestanas[3]
    
    if st.session_state["ROL_ACTUAL"] == "ADMINISTRADOR":
        tab_usuarios = pestanas[4]
        tab_borrado = pestanas[5]
        tab_auditoria = pestanas[6]

    # ==========================================
    # PESTAÑA 1: DASHBOARD Y REPORTES
    # ==========================================
    with tab_dashboard:
        st.subheader("🔍 FILTROS DE CONSULTA")
        col_fil, col_exp = st.columns([3, 1])
        with col_fil:
            proveedores_sel = st.multiselect("SELECCIONAR PROVEEDORES", options=todos_proveedores, default=todos_proveedores)
        
        df_filtrado = df_datos.copy()
        if proveedores_sel and 'PROVEEDOR' in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado['PROVEEDOR'].isin(proveedores_sel)]

        with col_exp:
            st.markdown("### EXPORTAR")
            excel_bytes_dash = df_a_excel_bytes(df_filtrado)
            st.download_button(
                label="📥 EXPORTAR REPORTE A EXCEL",
                data=excel_bytes_dash,
                file_name=f"ESTADO_DE_CUENTA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        col1, col2 = st.columns(2)
        with col1:
            if 'TOTAL' in df_filtrado.columns:
                st.metric("💰 TOTAL GENERAL ACUMULADO ($)", formato_moneda(df_filtrado['TOTAL'].sum()))
        with col2:
            st.metric("📄 REGISTROS MOSTRADOS", len(df_filtrado))

        st.markdown("---")
        
        st.subheader("🖨️ REIMPRIMIR COMPROBANTES PDF (REGISTROS ANTERIORES)")
        if not df_filtrado.empty:
            opciones_facturas = []
            dict_facturas = {}
            for idx, row in df_filtrado.iterrows():
                label_op = f"REG #{row.get('N°')} | PROV: {row.get('PROVEEDOR')} | DOC: {row.get('TIPO')} {row.get('FACTURA')} | TOTAL: ${formato_moneda(row.get('TOTAL'))}"
                opciones_facturas.append(label_op)
                dict_facturas[label_op] = row
            
            factura_seleccionada_str = st.selectbox("SELEACCIONE UN DOCUMENTO REGISTRADO PARA IMPRIMIR COMPROBANTE:", options=opciones_facturas)
            if factura_seleccionada_str:
                row_sel = dict_facturas[factura_seleccionada_str]
                pdf_bytes_hist = generar_pdf_comprobante(row_sel)
                st.download_button(
                    label=f"🖨️ DESCARGAR COMPROBANTE PDF (DOCUMENTO {row_sel.get('FACTURA')})",
                    data=pdf_bytes_hist,
                    file_name=f"COMPROBANTE_{row_sel.get('TIPO')}_{row_sel.get('FACTURA')}.pdf",
                    mime="application/pdf"
                )
        else:
            st.info("NO HAY REGISTROS DISPONIBLES PARA REIMPRIMIR CON LOS FILTROS ACTUALES.")

        st.markdown("---")
        st.subheader("📈 ANÁLISIS VISUAL")
        col_g1, col_g2 = st.columns(2)

        with col_g1:
            if 'PROVEEDOR' in df_filtrado.columns and 'TOTAL' in df_filtrado.columns and not df_filtrado.empty:
                df_prov = df_filtrado.groupby('PROVEEDOR')['TOTAL'].sum().reset_index()
                fig_prov = px.bar(df_prov, x='PROVEEDOR', y='TOTAL', title="MONTO TOTAL POR PROVEEDOR ($)", text_auto='.2s', color='PROVEEDOR')
                fig_prov.update_layout(showlegend=False)
                st.plotly_chart(fig_prov, use_container_width=True)

        with col_g2:
            if 'ESTADO' in df_filtrado.columns and 'TOTAL' in df_filtrado.columns and not df_filtrado.empty:
                df_est = df_filtrado.groupby('ESTADO')['TOTAL'].sum().reset_index()
                fig_est = px.pie(df_est, names='ESTADO', values='TOTAL', title="DISTRIBUCIÓN POR ESTADO DE CUENTA", hole=0.4)
                st.plotly_chart(fig_est, use_container_width=True)

        st.subheader("📋 DETALLE DE DOCUMENTOS")
        df_tabla_display = df_filtrado.copy()
        
        for col_f in ['F FACTURA', 'F RECEPCION', 'VENCIMIENTO']:
            if col_f in df_tabla_display.columns:
                df_tabla_display[col_f] = df_tabla_display[col_f].apply(formato_fecha)

        if 'F RECEPCION' in df_tabla_display.columns:
            df_tabla_display['LÍMITE PP1'] = df_filtrado.apply(lambda r: calcular_fecha_pp(r.get('F RECEPCION'), r.get('PP1', 0)), axis=1)
            df_tabla_display['LÍMITE PP2'] = df_filtrado.apply(lambda r: calcular_fecha_pp(r.get('F RECEPCION'), r.get('PP2', 0)), axis=1)
            df_tabla_display['LÍMITE PP3'] = df_filtrado.apply(lambda r: calcular_fecha_pp(r.get('F RECEPCION'), r.get('PP3', 0)), axis=1)

        cols_moneda = ['SUB-TOTAL', 'EXENTO', 'BASE', 'IVA', 'IVA NR', 'TOTAL', 'TOTAL BS']
        for c in cols_moneda:
            if c in df_tabla_display.columns:
                df_tabla_display[c] = df_tabla_display[c].apply(formato_moneda)
        if 'TASA' in df_tabla_display.columns:
            df_tabla_display['TASA'] = df_tabla_display['TASA'].apply(lambda x: formato_moneda(x, decimales=4))

        st.dataframe(df_tabla_display, use_container_width=True)

    # ==========================================
    # PESTAÑA 2: REGISTRAR NUEVA FACTURA
    # ==========================================
    with tab_nueva_factura:
        st.subheader("➕ REGISTRO DE DOCUMENTOS (FC / NC)")
        
        with st.form("FORM_REGISTRO_FACTURA", clear_on_submit=False):
            col_a1, col_a2, col_a3 = st.columns(3)
            with col_a1:
                sel_prov = st.selectbox("PROVEEDOR", options=todos_proveedores if todos_proveedores else ["NUEVO PROVEEDOR"])
            with col_a2:
                sel_tipo_doc = st.selectbox("TIPO DE DOCUMENTO", ["FC", "NC"], help="FC: FACTURA / NC: NOTA DE CRÉDITO")
            with col_a3:
                num_fact = st.text_input("NÚMERO DE DOCUMENTO / FACTURA")

            col_b1, col_b2, col_b3, col_b4 = st.columns(4)
            with col_b1:
                f_fac = st.date_input("FECHA EMISIÓN", datetime.today())
            with col_b2:
                f_rec = st.date_input("FECHA RECEPCIÓN", datetime.today())
            with col_b3:
                d_cred = st.number_input("DÍAS DE CRÉDITO", min_value=0, value=21, step=1)
            with col_b4:
                ocd_val = st.text_input("OCD ICG (OPCIONAL)")

            st.markdown("**DÍAS DE PRONTO PAGO (PP)**")
            col_pp1, col_pp2, col_pp3 = st.columns(3)
            with col_pp1:
                pp1_val = st.number_input("DÍAS PRONTO PAGO 1 (PP1)", min_value=0, value=0, step=1)
            with col_pp2:
                pp2_val = st.number_input("DÍAS PRONTO PAGO 2 (PP2)", min_value=0, value=0, step=1)
            with col_pp3:
                pp3_val = st.number_input("DÍAS PRONTO PAGO 3 (PP3)", min_value=0, value=0, step=1)

            st.markdown("**VALORES FINANCIEROS Y CÁLCULOS AUTOMÁTICOS**")
            col_c1, col_c2, col_c3, col_c4 = st.columns(4)
            with col_c1:
                tasa_val = st.number_input("TASA DE CAMBIO (BS/$)", min_value=0.0, value=875.6505, format="%.4f", step=0.0001)
            with col_c2:
                sub_total_val = st.number_input("SUB-TOTAL ($)", min_value=0.0, value=0.0, step=1.0)
            with col_c3:
                exento_val = st.number_input("MONTO EXENTO ($)", min_value=0.0, value=0.0, step=1.0)
            with col_c4:
                base_val = st.number_input("BASE IMPONIBLE ($)", min_value=0.0, value=0.0, step=1.0)

            iva_calc = base_val * 0.16
            iva_nr_calc = iva_calc * 0.25
            total_calc = sub_total_val + iva_nr_calc
            tot_bs_calc = total_calc * tasa_val

            col_d1, col_d2, col_d3, col_d4 = st.columns(4)
            with col_d1:
                st.text_input("IVA ($) [16% DE BASE]", value=formato_moneda(iva_calc), disabled=True)
            with col_d2:
                st.text_input("IVA NO RETENIDO ($) [25% DE IVA]", value=formato_moneda(iva_nr_calc), disabled=True)
            with col_d3:
                st.text_input("TOTAL GENERAL ($) [SUB-TOTAL + IVA NR]", value=formato_moneda(total_calc), disabled=True)
            with col_d4:
                st.text_input("TOTAL EQUIVALENTE EN BS [TOTAL $ * TASA]", value=formato_moneda(tot_bs_calc), disabled=True)

            est_val = st.selectbox("ESTADO", ["VIGENTE", "PENDIENTE", "PAGADA"])

            btn_guardar_factura = st.form_submit_button("💾 GUARDAR REGISTRO Y GENERAR COMPROBANTE PDF")

            if btn_guardar_factura:
                f_venc = pd.to_datetime(f_fac) + pd.Timedelta(days=int(d_cred))
                try:
                    num_nuevo = int(pd.to_numeric(df_datos['N°'], errors='coerce').max() + 1)
                except Exception:
                    num_nuevo = 1

                nueva_row = {
                    'N°': str(num_nuevo),
                    'OCD ICG': str(ocd_val),
                    'PROVEEDOR': sel_prov,
                    'TIPO': sel_tipo_doc,
                    'FACTURA': str(num_fact),
                    'F FACTURA': pd.to_datetime(f_fac),
                    'F RECEPCION': pd.to_datetime(f_rec),
                    'CREDITO': float(d_cred),
                    'PP1': str(int(pp1_val)),
                    'PP2': str(int(pp2_val)),
                    'PP3': str(int(pp3_val)),
                    'VENCIMIENTO': f_venc,
                    'DIAS': 0.0,
                    'TASA': float(tasa_val),
                    'SUB-TOTAL': float(sub_total_val),
                    'EXENTO': float(exento_val),
                    'BASE': float(base_val),
                    'IVA': float(iva_calc),
                    'IVA NR': float(iva_nr_calc),
                    'TOTAL': float(total_calc),
                    'TOTAL BS': float(tot_bs_calc),
                    'ESTADO': est_val
                }

                df_full_datos = pd.read_excel(file_path, sheet_name="DATOS")
                df_full_datos = pd.concat([df_full_datos, pd.DataFrame([nueva_row])], ignore_index=True)

                with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df_full_datos.to_excel(writer, sheet_name='DATOS', index=False)

                registrar_auditoria(st.session_state["USUARIO_ACTUAL"], "REGISTRO DE FACTURA", f"REGISTRADA FACTURA N° {num_fact} DE {sel_prov} POR ${formato_moneda(total_calc)}")

                st.session_state['ultimo_registro'] = nueva_row
                st.success("¡REGISTRO GUARDADO EXITOSAMENTE EN EL EXCEL!")

        if 'ultimo_registro' in st.session_state and st.session_state['ultimo_registro']:
            reg = st.session_state['ultimo_registro']
            pdf_bytes = generar_pdf_comprobante(reg)
            st.markdown("---")
            st.info("📄 **COMPROBANTE PDF GENERADO LISTO PARA IMPRESIÓN:**")
            st.download_button(
                label=f"🖨️ DESCARGAR E IMPRIMIR COMPROBANTE PDF (FACTURA N° {reg.get('FACTURA')})",
                data=pdf_bytes,
                file_name=f"COMPROBANTE_FC_{reg.get('FACTURA')}_{datetime.now().strftime('%Y%m%d')}.pdf",
                mime="application/pdf"
            )

    # ==========================================
    # PESTAÑA 3: GESTIÓN DE PROVEEDORES
    # ==========================================
    with tab_proveedores:
        st.subheader("🏢 DIRECTORIO DE PROVEEDORES")
        
        col_p1, col_p2 = st.columns([1, 2])

        with col_p1:
            st.markdown("### AGREGAR NUEVO PROVEEDOR")
            with st.form("FORM_NUEVO_PROVEEDOR"):
                nombre_nuevo_prov = st.text_input("NOMBRE DE LA EMPRESA / PROVEEDOR")
                btn_crear_prov = st.form_submit_button("AÑADIR PROVEEDOR")

                if btn_crear_prov and nombre_nuevo_prov.strip():
                    nombre_limpio = nombre_nuevo_prov.strip().upper()
                    
                    df_cond_full = cargar_condiciones()
                    if 'Proveedores' in df_cond_full.columns and nombre_limpio not in df_cond_full['Proveedores'].values:
                        nueva_cond_row = {
                            'Proveedores': nombre_limpio,
                            'Factura digital': 'NO',
                            'Condicion': 'CONTADO',
                            'Departamento': '',
                            'Nombre': '',
                            'Correo': '',
                            'Numero de telefono': '',
                            'comentarios': ''
                        }
                        df_cond_full = pd.concat([df_cond_full, pd.DataFrame([nueva_cond_row])], ignore_index=True)
                        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            df_cond_full.to_excel(writer, sheet_name='CONDICIONES', index=False)
                        
                        registrar_auditoria(st.session_state["USUARIO_ACTUAL"], "CREACIÓN DE PROVEEDOR", f"NUEVO PROVEEDOR CREADO: {nombre_limpio}")
                        st.success(f"¡PROVEEDOR '{nombre_limpio}' REGISTRADO CON ÉXITO!")
                        st.rerun()

        with col_p2:
            st.markdown("### LISTA DE PROVEEDORES REGISTRADOS")
            df_prov_export = pd.DataFrame({"PROVEEDORES ACTIVOS": todos_proveedores})
            st.dataframe(df_prov_export, use_container_width=True)
            
            excel_bytes_prov = df_a_excel_bytes(df_prov_export)
            st.download_button(
                label="📥 EXPORTAR LISTA DE PROVEEDORES A EXCEL",
                data=excel_bytes_prov,
                file_name="LISTADO_PROVEEDORES.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # ==========================================
    # PESTAÑA 4: CONDICIONES COMERCIALES
    # ==========================================
    with tab_condiciones:
        st.subheader("📋 MÓDULO DE CONDICIONES COMERCIALES POR PROVEEDOR")
        
        col_c_left, col_c_right = st.columns([3, 1])
        with col_c_right:
            excel_bytes_cond = df_a_excel_bytes(df_condiciones)
            st.download_button(
                label="📥 EXPORTAR CONDICIONES A EXCEL",
                data=excel_bytes_cond,
                file_name="CONDICIONES_COMERCIALES.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        edited_cond_df = st.data_editor(
            df_condiciones,
            use_container_width=True,
            num_rows="dynamic",
            key="EDITOR_CONDICIONES"
        )

        if st.button("💾 GUARDAR CAMBIOS EN CONDICIONES"):
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                edited_cond_df.to_excel(writer, sheet_name='CONDICIONES', index=False)
            
            registrar_auditoria(st.session_state["USUARIO_ACTUAL"], "MODIFICACIÓN DE CONDICIONES", "ACTUALIZACIÓN GENERAL EN TABLA DE CONDICIONES COMERCIALES")
            st.success("¡CONDICIONES COMERCIALES ACTUALIZADAS E INTEGRADAS EN EXCEL!")
            st.rerun()

    # ==========================================
    # PESTAÑAS ADMINISTRATIVAS (SOLO ADMIN)
    # ==========================================
    if st.session_state["ROL_ACTUAL"] == "ADMINISTRADOR":
        # 👥 PESTAÑA 5: GESTIÓN DE USUARIOS
        with tab_usuarios:
            st.subheader("👥 CONTROL Y CREACIÓN DE USUARIOS")
            
            col_u1, col_u2 = st.columns([1, 1])
            
            with col_u1:
                st.markdown("### CREAR NUEVO USUARIO")
                with st.form("FORM_NUEVO_USUARIO"):
                    nuevo_nombre = st.text_input("NOMBRE Y APELLIDO COMPLETO")
                    nuevo_usuario = st.text_input("NOMBRE DE USUARIO (LOGIN)")
                    nueva_clave = st.text_input("CONTRASEÑA", type="password")
                    nuevo_rol = st.selectbox("ROL DE ACCESO", ["USUARIO", "ADMINISTRADOR"])
                    btn_crear_usuario = st.form_submit_button("💾 REGISTRAR NUEVO USUARIO")
                    
                    if btn_crear_usuario:
                        if nuevo_nombre.strip() and nuevo_usuario.strip() and nueva_clave.strip():
                            df_usr_curr = cargar_usuarios()
                            usr_clean = nuevo_usuario.strip().lower()
                            
                            if usr_clean in df_usr_curr['USUARIO'].values:
                                st.error("¡EL NOMBRE DE USUARIO YA EXISTE! ELIJA OTRO.")
                            else:
                                nueva_usr_row = {
                                    "USUARIO": usr_clean,
                                    "PASSWORD": nueva_clave.strip(),
                                    "ROL": nuevo_rol.upper(),
                                    "NOMBRE": nuevo_nombre.strip().upper()
                                }
                                df_usr_curr = pd.concat([df_usr_curr, pd.DataFrame([nueva_usr_row])], ignore_index=True)
                                
                                with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                    df_usr_curr.to_excel(writer, sheet_name='USUARIOS', index=False)
                                
                                registrar_auditoria(st.session_state["USUARIO_ACTUAL"], "CREACIÓN DE USUARIO", f"CREADO NUEVO USUARIO '{usr_clean.upper()}' CON ROL '{nuevo_rol.upper()}'")
                                st.success(f"¡USUARIO '{usr_clean.upper()}' REGISTRADO CON ÉXITO!")
                                st.rerun()
                        else:
                            st.warning("POR FAVOR COMPLETE TODOS LOS CAMPOS DEL FORMULARIO.")

                st.markdown("---")
                st.markdown("### 🗑️ ELIMINAR USUARIO")
                df_usr_curr = cargar_usuarios()
                usuarios_borrables = [u for u in df_usr_curr['USUARIO'].tolist() if u != 'admin']
                
                with st.form("FORM_ELIMINAR_USUARIO"):
                    usuario_a_borrar = st.selectbox("SELECCIONE USUARIO A ELIMINAR", options=usuarios_borrables if usuarios_borrables else [""])
                    btn_eliminar = st.form_submit_button("❌ ELIMINAR USUARIO SELECCIONADO")
                    
                    if btn_eliminar and usuario_a_borrar:
                        df_usr_nuevo = df_usr_curr[df_usr_curr['USUARIO'] != usuario_a_borrar]
                        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            df_usr_nuevo.to_excel(writer, sheet_name='USUARIOS', index=False)
                        
                        registrar_auditoria(st.session_state["USUARIO_ACTUAL"], "ELIMINACIÓN DE USUARIO", f"ELIMINADO EL USUARIO '{usuario_a_borrar.upper()}'")
                        st.success(f"¡USUARIO '{usuario_a_borrar.upper()}' ELIMINADO EXITOSAMENTE!")
                        st.rerun()
            
            with col_u2:
                st.markdown("### USUARIOS REGISTRADOS EN EL SISTEMA")
                df_usr_view = cargar_usuarios()
                df_usr_display = df_usr_view[['NOMBRE', 'USUARIO', 'ROL']].copy()
                df_usr_display.columns = ['NOMBRE Y APELLIDO', 'USUARIO DE ACCESO', 'ROL DEL SISTEMA']
                st.dataframe(df_usr_display, use_container_width=True)

        # 🗑️ PESTAÑA 6: GESTIÓN Y BORRADO DE REGISTROS DE FACTURAS / NC
        with tab_borrado:
            st.subheader("🗑️ MÓDULO DE GESTIÓN Y BORRADO DE FACTURAS / NC")
            st.markdown("ESTE MÓDULO PERMITE ELIMINAR REGISTROS INDIVIDUALES O LIMPIAR TODO EL HISTORIAL PARA EMPEZAR EN LIMPIO.")

            col_b1, col_b2 = st.columns(2)

            with col_b1:
                st.markdown("### 📌 ELIMINAR UN REGISTRO ESPECÍFICO")
                df_datos_curr = cargar_datos()
                
                if not df_datos_curr.empty:
                    opciones_borrado = []
                    dict_borrado = {}
                    for idx, row in df_datos_curr.iterrows():
                        label_b = f"REG #{row.get('N°')} | PROV: {row.get('PROVEEDOR')} | DOC: {row.get('TIPO')} {row.get('FACTURA')} | TOTAL: ${formato_moneda(row.get('TOTAL'))}"
                        opciones_borrado.append(label_b)
                        dict_borrado[label_b] = row.get('N°')

                    with st.form("FORM_BORRAR_REGISTRO_INDIVIDUAL"):
                        reg_sel_borrar = st.selectbox("SELEACCIONE EL DOCUMENTO A ELIMINAR", options=opciones_borrado)
                        btn_borrar_ind = st.form_submit_button("🗑️ ELIMINAR DOCUMENTO SELECCIONADO")

                        if btn_borrar_ind and reg_sel_borrar:
                            num_reg_target = dict_borrado[reg_sel_borrar]
                            df_nuevo_datos = df_datos_curr[df_datos_curr['N°'].astype(str) != str(num_reg_target)]
                            
                            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                df_nuevo_datos.to_excel(writer, sheet_name='DATOS', index=False)

                            registrar_auditoria(st.session_state["USUARIO_ACTUAL"], "ELIMINACIÓN DE REGISTRO", f"ELIMINADO DOCUMENTO N° REGISTRO {num_reg_target} DE LA HOJA DATOS")
                            st.success(f"¡DOCUMENTO REGISTRO N° {num_reg_target} ELIMINADO CON ÉXITO!")
                            st.rerun()
                else:
                    st.info("NO HAY DOCUMENTOS REGISTRADOS PARA ELIMINAR.")

            with col_b2:
                st.markdown("### ⚠️ VACIAR TODA LA BASE DE DATOS (REINICIAR DESDE CERO)")
                st.warning("¡ATENCIÓN! ESTA ACCIÓN BORRARÁ TODOS LOS REGISTROS DE FACTURAS Y NOTAS DE CRÉDITO DE LA HOJA 'DATOS' PARA EMPEZAR A GRABAR INFORMACIÓN REAL DESDE CERO.")
                
                confirmacion = st.checkbox("ENTIENDO QUE ESTA ACCIÓN ELIMINARÁ TODA LA INFORMACIÓN DE FACTURAS.")
                
                if st.button("🚨 REINICIAR Y BORRAR TODOS LOS REGISTROS", disabled=not confirmacion):
                    cols_def = ['N°', 'OCD ICG', 'PROVEEDOR', 'TIPO', 'FACTURA', 'F FACTURA', 'F RECEPCION', 'CREDITO', 'PP1', 'PP2', 'PP3', 'VENCIMIENTO', 'DIAS', 'TASA', 'SUB-TOTAL', 'EXENTO', 'BASE', 'IVA', 'IVA NR', 'TOTAL', 'TOTAL BS', 'ESTADO']
                    df_vacio = pd.DataFrame(columns=cols_def)
                    
                    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df_vacio.to_excel(writer, sheet_name='DATOS', index=False)
                    
                    registrar_auditoria(st.session_state["USUARIO_ACTUAL"], "REINICIO DE BASE DE DATOS", "SE HA LIMPIADO COMPLETAMENTE LA HOJA DE DATOS DE FACTURAS")
                    st.success("¡BASE DE DATOS REINICIADA EXITOSAMENTE! AHORA PUEDES EMPEZAR A GRABAR DATOS REALES DESDE CERO.")
                    st.rerun()

        # 📜 PESTAÑA 7: AUDITORÍA DE ACCIONES
        with tab_auditoria:
            st.subheader("📜 REGISTRO DE AUDITORÍA Y MODIFICACIONES DE USUARIOS")
            st.markdown("ESTE MÓDULO REGISTRA CADA ACCIÓN, MODIFICACIÓN Y REGISTRO REALIZADO POR CUALQUIER EMPLEADO EN EL SISTEMA.")
            
            try:
                df_aud = pd.read_excel(file_path, sheet_name="AUDITORIA")
                st.dataframe(df_aud.sort_index(ascending=False), use_container_width=True)
                
                excel_bytes_aud = df_a_excel_bytes(df_aud)
                st.download_button(
                    label="📥 EXPORTAR HISTORIAL DE AUDITORÍA A EXCEL",
                    data=excel_bytes_aud,
                    file_name=f"AUDITORIA_SISTEMA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception:
                st.info("AÚN NO HAY REGISTROS DE AUDITORÍA ALMACENADOS.")

except Exception as e:
    st.error(f"ERROR AL PROCESAR LA APLICACIÓN: {e}")