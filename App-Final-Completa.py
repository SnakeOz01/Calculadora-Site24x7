import json
import math
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Site24x7 Calculator LicenseAmounts", layout="wide", page_icon="🔍")

CONFIG_PATH = "Config-Fixed-V3.json"
# Fallback si no existe con ese nombre
try:
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        CONFIG = json.load(f)
except:
    with open("config_fixed_v3.json", "r", encoding="utf-8") as f:
        CONFIG = json.load(f)

# --- HELPERS (sin cambios de App-Final-Completa) ---
def addon_effective_size(opt):
    size = opt.get("size", 0)
    unit = str(opt.get("unit", "")).upper().strip()
    if unit in ["K", ""]:
        return int(size)
    elif unit == "M":
        return int(size * 1000)
    elif unit == "B":
        return int(size * 1_000_000)
    elif unit == "GB":
        return int(size)
    elif unit == "TB":
        return int(size * 1000)
    return int(size)

def bom_effective_qty(name, qty):
    if qty == 0:
        return 0
    qty = int(qty)
    low = name.lower().replace(" ", "")
    if "10gb" in low:
        return qty * 10
    if "100gb" in low:
        return qty * 100
    if "1tb" in low:
        return qty * 1000
    if "500k" in low:
        return qty * 500
    if "10m" in low and ("pageview" in low or "synthetic" in low):
        return qty * 10000
    if "5m" in low and "pageview" in low:
        return qty * 5000
    if "50m" in low:
        return qty * 50000
    if "100m" in low:
        return qty * 100000
    if "1b" in low:
        return qty * 1000000
    if "m(millones)" in low:
        return qty * 1000
    return qty

# --- LECTOR CON VALIDACIÓN (TRAÍDO DE App-Final) ---
def leer_bom_excel(uploaded_file):
    errores = []
    try:
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None, engine="openpyxl")
        return df, "openpyxl"
    except Exception as e:
        errores.append(f"openpyxl: {type(e).__name__}: {str(e)}")
    try:
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None, engine="calamine")
        return df, "calamine"
    except Exception as e:
        errores.append(f"calamine: {type(e).__name__}: {str(e)}")
    # Fallback XML manual (para archivos corruptos como Boomplain_2_columns.xlsx)
    try:
        uploaded_file.seek(0)
        data_bytes = uploaded_file.read()
        with zipfile.ZipFile(BytesIO(data_bytes)) as z:
            try:
                ss_xml = z.read("xl/sharedStrings.xml")
                root_ss = ET.fromstring(ss_xml)
                ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                strings = []
                for si in root_ss.findall('main:si', ns):
                    t_elems = si.findall('main:t', ns)
                    text = ''.join([t.text if t.text else '' for t in t_elems])
                    strings.append(text)
            except:
                strings = []
            sheet_xml = z.read("xl/worksheets/sheet1.xml")
            root_sheet = ET.fromstring(sheet_xml)
            ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            sheet_data = root_sheet.find('main:sheetData', ns)
            rows = []
            for row in sheet_data.findall('main:row', ns):
                cells = row.findall('main:c', ns)
                row_vals = []
                for c in cells:
                    t = c.get('t')
                    v_elem = c.find('main:v', ns)
                    v = v_elem.text if v_elem is not None else ''
                    if t == 's':
                        try:
                            idx = int(v)
                            val = strings[idx] if idx < len(strings) else ""
                        except:
                            val = ""
                        row_vals.append(val)
                    else:
                        row_vals.append(v)
                rows.append(row_vals)
        df = pd.DataFrame(rows)
        return df, "xml_manual"
    except Exception as e:
        errores.append(f"xml_manual: {type(e).__name__}: {str(e)}")

    raise RuntimeError("No fue posible leer el Excel.\n\n" + "\n".join(errores))

def procesar_bom_site24x7_con_validacion(df):
    if df is None or df.empty:
        raise ValueError("El archivo Excel está vacío.")
    if df.shape[1] < 2:
        raise ValueError("El BoM debe tener al menos dos columnas: Recurso y Cantidad.")

    df = df.iloc[:, :2].copy()
    df.columns = ["recurso", "cantidad"]
    df["recurso"] = df["recurso"].astype("string").str.strip()

    seccion_actual = None
    registros = []
    errores_validacion = []

    for idx, row in df.iterrows():
        recurso = row["recurso"]
        cantidad = row["cantidad"]

        if pd.isna(recurso) and pd.isna(cantidad):
            continue

        cantidad_texto = ""
        if not pd.isna(cantidad):
            cantidad_texto = str(cantidad).strip()

        if not pd.isna(recurso) and cantidad_texto.lower() == "cantidades":
            seccion_actual = str(recurso).strip()
            continue

        if str(recurso).strip().lower() in ["basic monitors", "host monitors", "advanced monitors", "network component license", "includes gb applogs", "webpage views", "recursos", "recursos (site24x7)"]:
            continue

        if pd.isna(recurso) or str(recurso).strip() == "" or str(recurso).strip().lower() == "nan":
            continue

        # --- VALIDACIÓN ESTRICTA DE CANTIDAD (DE App-Final) ---
        if pd.isna(cantidad) or str(cantidad).strip() == "":
            errores_validacion.append(f"❌ La cantidad del recurso '{recurso}' está vacía. Debe ingresar un número entero >=0. (Fila Excel {idx+1})")
            continue

        cantidad_texto = str(cantidad).strip()

        # Validación: debe ser número entero
        try:
            cantidad_float = float(cantidad_texto.replace(",", "."))
            if not cantidad_float.is_integer():
                errores_validacion.append(f'❌ El Valor: **"{cantidad_texto}"**, no es un número entero válido para el Item **{recurso}**. (Fila {idx+1}) - Debe ser 0,1,2...')
                continue
            cantidad_num = int(cantidad_float)
        except:
            errores_validacion.append(f'❌ El Valor: \n**"{cantidad_texto}"**, no es un número entero válido para el Item: \n**"{recurso}"**\n. \nContiene letras o caracteres especiales en (Fila {idx+1}). \nEjemplos válidos: 0, 1, 10, 100.')
            continue

        if cantidad_num < 0:
            errores_validacion.append(f"❌ La cantidad del recurso '{recurso}' es negativa ({cantidad_num}). Debe ser >=0. (Fila {idx+1})")
            continue

        registros.append({
            "seccion": seccion_actual or "Sin sección",
            "recurso": str(recurso).strip(),
            "cantidad": cantidad_num,
            "fila_excel": idx + 1  # Fila exacta del Excel (1-indexed) para que coincida con validación
        })

    if errores_validacion:
        mensaje = "Se encontraron errores en la columna Cantidades:\n\n" + "\n".join(errores_validacion)
        raise ValueError(mensaje)

    resultado = pd.DataFrame(registros, columns=["seccion", "recurso", "cantidad", "fila_excel"])
    return resultado

def sum_by_category(bom_dict):
    totals = {"basic": 0, "host": 0, "advanced": 0, "network": 0, "rum_pageviews_k": 0, "applogs_gb": 0, "synthetic_runs_k": 0}
    mapping = CONFIG.get("mapping_bom_to_category", {})
    for rn, qty in bom_dict.items():
        if qty <= 0:
            continue
        rn_clean = rn.strip()
        if rn_clean.lower() in ["basic monitors", "host monitors", "advanced monitors", "network component license", "includes gb applogs", "webpage views", "cantidades"]:
            continue
        if "retention" in rn_clean.lower():
            continue
        cat = mapping.get(rn_clean)
        if not cat:
            rn_nospace = rn_clean.lower().replace(" ", "")
            for k, v in mapping.items():
                k_nospace = k.lower().replace(" ", "")
                if k_nospace in rn_nospace or rn_nospace in k_nospace:
                    cat = v
                    break
        if not cat:
            low = rn_clean.lower()
            if any(x in low for x in ["website", "websocket", "dns", "ping", "ssl", "brand", "soap", "rest", "port", "blocklist", "file upload", "grpc", "hadoop", "cron", "kubernetes pod", "smart disk", "zookeeper", "statsd", "plugin", "nutanix storage", "vmware resource", "datastore", "snapshot", "azure monitoring", "backup monitoring"]):
                cat = "basic"
            elif any(x in low for x in ["windows, linux", "docker hosts", "host containers", "kubernetes nod", "podman", "vmware vcenter", "vmware vm", "vmware esx", "xen", "ec2", "azure vms", "gcp compute", "nutanix cluster", "nutanix host", "nutanix vm", "vmware nsx"]):
                cat = "host"
            elif any(x in low for x in ["web transaction", "ftp transfer", "mail delivery", "web page speed", "defacement", "apm-insight", "biztalk", "sharepoint", "office 365", "active directory", "exchange", "hyper-v", "sql monitoring", "failover", "horizon", "isp latency", "mysql", "postgresql", "oracle", "slo"]):
                cat = "advanced"
            elif any(x in low for x in ["network device", "network monitoring", "netflow", "ncm", "meraki", "voip", "wan rtt", "ipam", "cisco aci", "wireless lan", "access points", "velocloud"]):
                cat = "network"
            elif "log" in low:
                cat = "applogs_gb"
            elif "pageview" in low:
                cat = "rum_pageviews_k"
            elif "synthetic" in low:
                cat = "synthetic_runs_k"
        if cat in totals:
            totals[cat] += bom_effective_qty(rn_clean, qty)
    return totals

def calculate_addons_needed_optimized(required, included, addon_options):
    deficit = max(0, int(required) - int(included))
    if deficit == 0:
        return {"deficit": 0, "packs": [], "total_extra": 0, "cost": 0}
    if not addon_options:
        return {"deficit": deficit, "packs": [], "total_extra": 0, "cost": 0}
    normalized = []
    for opt in addon_options:
        eff = addon_effective_size(opt)
        if eff <= 0:
            continue
        normalized.append({"orig_size": opt.get("size"), "unit": opt.get("unit", ""), "eff_size": eff, "price": opt.get("price", 0)})
    if not normalized:
        return {"deficit": deficit, "packs": [], "total_extra": 0, "cost": 0}
    normalized.sort(key=lambda x: x["eff_size"])
    max_size = max(o["eff_size"] for o in normalized)
    limit = deficit + max_size
    if deficit > 500000 and limit > 2000000:
        cheapest = min(normalized, key=lambda x: x["price"] / x["eff_size"])
        packs_needed = math.ceil(deficit / cheapest["eff_size"])
        return {"deficit": deficit, "packs": [{"size": cheapest["orig_size"], "unit": cheapest["unit"], "eff_size": cheapest["eff_size"], "qty": packs_needed, "price": cheapest["price"], "total_price": packs_needed * cheapest["price"]}], "total_extra": packs_needed * cheapest["eff_size"], "cost": packs_needed * cheapest["price"]}
    if limit > 2000000:
        limit = deficit + 100000
    INF = float('inf')
    dp_cost = [INF] * (limit + 1)
    dp_choice = [None] * (limit + 1)
    dp_cost[0] = 0
    for i in range(limit + 1):
        if dp_cost[i] == INF:
            continue
        for opt in normalized:
            nxt = i + opt["eff_size"]
            if nxt <= limit and dp_cost[i] + opt["price"] < dp_cost[nxt]:
                dp_cost[nxt] = dp_cost[i] + opt["price"]
                dp_choice[nxt] = (i, opt)
    best_cost = INF
    best_idx = -1
    for i in range(deficit, limit + 1):
        if dp_cost[i] < best_cost:
            best_cost = dp_cost[i]
            best_idx = i
    if best_idx == -1 or best_cost == INF:
        return {"deficit": deficit, "packs": [], "total_extra": 0, "cost": 0}
    packs_count = {}
    cur = best_idx
    while cur > 0 and dp_choice[cur] is not None:
        prev, opt = dp_choice[cur]
        key = (opt["orig_size"], opt["unit"], opt["price"], opt["eff_size"])
        packs_count[key] = packs_count.get(key, 0) + 1
        cur = prev
    packs_list = []
    total_extra = 0
    for (orig_size, unit, price, eff_size), qty in packs_count.items():
        packs_list.append({"size": orig_size, "unit": unit, "eff_size": eff_size, "qty": qty, "price": price, "total_price": qty * price})
        total_extra += qty * eff_size
    return {"deficit": deficit, "packs": packs_list, "total_extra": total_extra, "cost": best_cost, "covered": best_idx}

def evaluate_all_plans(totals):
    results = {}
    for plan_name, plan_data in CONFIG["plans"].items():
        included = plan_data["included"]
        addons_def = plan_data["addons"]
        plan_result = {}
        total_cost = 0
        for cat_key, inc_key in [("basic", "basic"), ("host", "host"), ("advanced", "advanced"), ("network", "network")]:
            req = totals.get(inc_key, 0)
            inc = included.get(inc_key, 0)
            opts = addons_def.get(cat_key, [])
            calc = calculate_addons_needed_optimized(req, inc, opts)
            plan_result[inc_key] = calc
            total_cost += calc["cost"]
        for addon_key, total_key in [("rum", "rum_pageviews_k"), ("applogs", "applogs_gb"), ("synthetic", "synthetic_runs_k")]:
            req = totals.get(total_key, 0)
            inc = included.get(total_key, 0)
            opts = addons_def.get(addon_key, [])
            calc = calculate_addons_needed_optimized(req, inc, opts)
            plan_result[total_key] = calc
            total_cost += calc["cost"]
        results[plan_name] = {"details": plan_result, "addon_cost": total_cost, "base_price": plan_data.get("base_price_usd", 0), "total_price": total_cost + plan_data.get("base_price_usd", 0), "included": included}
    return results

# --- SIDEBAR (igual que App-Final-Completa) ---
with st.sidebar:
    st.header("1️⃣ Plan de licenciamiento")
    st.write("Selecciona el plan que deseas cotizar")
    plans = list(CONFIG["plans"].keys())
    selected_plan = st.selectbox("Plan", plans, index=1)
    inc = CONFIG["plans"][selected_plan]["included"]
    st.subheader(f"Incluidos - {selected_plan}")
    col_a, col_b = st.columns(2)
    col_a.metric("Basic", inc.get("basic",0))
    col_b.metric("Host", inc.get("host",0))
    col_a.metric("Advanced", inc.get("advanced",0))
    col_b.metric("Network", inc.get("network",0))
    col_a.metric("RUM K", inc.get("rum_pageviews_k",0))
    col_b.metric("Logs GB", inc.get("applogs_gb",0))
    st.metric("Synthetic K", inc.get("synthetic_runs_k",0))
    st.divider()
    st.info("💡 El plan **Enterprise** incluye Anomaly Detection, Event Correlation, NCM Compliance (AIOps)")
    st.caption("Modelo de licenciamiento 2026")

st.title("🔍 Site24x7 - Calculadora de Licenciamiento")
st.caption("FInal Version Calculator V1.0 - ManageEngine LATAM")

# --- TABS (igual que App-Final-Completa) ---
tab1, tab2 = st.tabs(["2️⃣ Paso 2: Cargar BoM / Manual", "3️⃣ Paso 3: Resultados y Recomendación"])

with tab1:
    col1, col2 = st.columns([1,1], gap="large")
    
    with col1:
        st.subheader("Opción A: Subir (BoM) 📁")
        st.write("Carga el archivo Excel con las cantidades de recursos diligenciados.")
        uploaded = st.file_uploader("Excel BoM", type=["xlsx", "xls"], key="bom_upload")

        if uploaded:
            try:
                # VALIDACIÓN DE App-Final
                df_raw, motor = leer_bom_excel(uploaded)
                df_bom = procesar_bom_site24x7_con_validacion(df_raw)
                
                # COMPORTAMIENTO DE App-Final-Completa: mostrar solo con cantidad >0
                df_bom_filtrado = df_bom[df_bom["cantidad"] > 0].copy()
                
                st.success(f"✅ Archivo BoM válido: {len(df_bom_filtrado)} recursos con cantidades diligenciadas correctamente.")
                
                # CUADRO RESUMEN SOLO CON RECURSOS CON CANTIDADES CORRECTAS (Recurso, Cantidad, Fila Excel) - EN ORDEN DE FILA ORIGINAL
                df_resumen = df_bom_filtrado[["fila_excel", "recurso", "cantidad"]].copy()
                df_resumen.columns = [" # ", "Recurso", "Cantidad" ]
                # Sin sort_values para mantener orden de fila del Excel como antes - ya viene en orden de lectura
                
                st.subheader("📋 Resumen - Recursos con cantidades detectadas")
                # Mostrar con Fila Excel exacta para que coincida con validación (idx+1)
                st.dataframe(df_resumen, use_container_width=True, height=350, hide_index=True)
                
                # Para cálculo, usar solo los >0 con cantidades correctas
                bom_dict = dict(zip(df_resumen["Recurso"], df_resumen["Cantidad"]))
                totals = sum_by_category(bom_dict)
                st.session_state["totals"] = totals
                st.session_state["bom_dict"] = bom_dict
                
            except ValueError as ve:
                st.error(f"⚠️ Error de validación detectado:\n\n{str(ve)}")
                st.warning("Corrige el Excel en las filas indicadas y vuelve a subirlo. No se calcularán totales hasta que corrijas los caracteres inválidos.")
                st.stop()
            except Exception as e:
                st.error(f"Error leyendo Excel: {e}")

    with col2:
        st.subheader("Opción B: Entrada de cantidades (Manual)  ⌨️")
        st.write("Si no tienes Excel, ingresa totales por categoría directamente.")
        
        basic_m = st.number_input("**Basic Monitors** (Web, Ping, SSL, Brand...)", 0, 100000, 0, key="basic_m")
        host_m = st.number_input("**Host Monitors** (Servers, EC2, Azure, GCP...)", 0, 100000, 0, key="host_m")
        adv_m = st.number_input("**Advanced Monitors** (APM, Transacciones...)", 0, 100000, 0, key="adv_m")
        net_m = st.number_input("**Network Components** (Devices, NCM...)", 0, 100000, 0, key="net_m")
        rum_k = st.number_input("**RUM Pageviews en K** (750 = 750K, 5000 = 5M)", 0, 10000000, 0, key="rum_m")
        logs_gb = st.number_input("**AppLogs GB/mes** (100 = 100GB)", 0, 1000000, 0, key="logs_m")
        synth_k = st.number_input("**Synthetic Runs en K** (480 = 480K)", 0, 10000000, 0, key="synth_m")
        
        if st.button("Usar valores manuales"):
            totals_manual = {"basic": basic_m, "host": host_m, "advanced": adv_m, "network": net_m, "rum_pageviews_k": rum_k, "applogs_gb": logs_gb, "synthetic_runs_k": synth_k}
            st.session_state["totals"] = totals_manual
            st.success("Valores manuales guardados")
            totals = totals_manual

    # Mostrar totales si existen
    if "totals" in st.session_state:
        totals = st.session_state["totals"]
        st.divider()
        st.subheader("Totales detectados (categoría ✅)")
        
        c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
        c1.metric("Basic", totals["basic"], help="Website, SSL, Brand...")
        c2.metric("Host", totals["host"], help="Servers, EC2, Azure...")
        c3.metric("Advanced", totals["advanced"], help="APM, Transactions...")
        c4.metric("Network", totals["network"], help="Devices, NCM...")
        c5.metric("RUM K", totals["rum_pageviews_k"], help="Pageviews en miles")
        c6.metric("Logs GB", totals["applogs_gb"], help="AppLogs GB")
        c7.metric("Synthetic K", totals["synthetic_runs_k"], help="Synthetic runs en miles")
        
        st.subheader("📊 Distribución de recursos solicitados")
        chart_data = pd.DataFrame({
            "Categoria": ["Basic", "Host", "Advanced", "Network", "RUM K", "Logs GB", "Synthetic K"],
            "Requerido": [totals["basic"], totals["host"], totals["advanced"], totals["network"], totals["rum_pageviews_k"], totals["applogs_gb"], totals["synthetic_runs_k"]]
        })
        st.bar_chart(chart_data, x="Categoria", y="Requerido", use_container_width=True)

with tab2:
    if "totals" not in st.session_state:
        st.warning("⚠️ Primero carga un BoM o usa entrada manual en Paso 2")
    else:
        totals = st.session_state["totals"]
        results = evaluate_all_plans(totals)
        best_plan = min(results, key=lambda p: results[p]['total_price'])
        
        st.subheader("Totales detectados por categoría ✅")
        c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
        c1.metric("Basic", totals["basic"])
        c2.metric("Host", totals["host"])
        c3.metric("Advanced", totals["advanced"])
        c4.metric("Network", totals["network"])
        c5.metric("RUM K", totals["rum_pageviews_k"])
        c6.metric("Logs GB", totals["applogs_gb"])
        c7.metric("Synthetic K", totals["synthetic_runs_k"])

        #st.success(f"✅ **Recomendación automática:** {best_plan} - Total ${results[best_plan]['total_price']} USD/mes (Base ${results[best_plan]['base_price']} + Add-ons ${results[best_plan]['addon_cost']})")
        st.info(f"Nota: Si necesitas AIOps, elige Enterprise aunque {best_plan} sea más barato.")
        
        st.divider()
        st.subheader(f"🔍 Detalle Add-ons para el plan **{selected_plan}**.")
        
        det = results[selected_plan]["details"]
        quote_rows=[]
        quote_rows.append({"Items & Description": f"{selected_plan} \n Plan", "Quantity": 1, "Unit Price": results[selected_plan]["base_price"], "Total Price": results[selected_plan]["base_price"]})
        
        detail_data=[]
        for cat_label, key in [("Basic Monitors","basic"),("Host Monitors","host"),("Advanced Monitors","advanced"),("Network Components","network"),("RUM Pageviews","rum_pageviews_k"),("AppLogs","applogs_gb"),("Synthetic Runs","synthetic_runs_k")]:
            calc = det.get(key)
            if not calc:
                continue
            if calc["deficit"]>0:
                #packs_str = "\n  ||  ".join([f"{p['qty']}x {p['size']}{p['unit']} (${p['price']}/pack)" for p in calc["packs"]])
                packs_str = "\n   ||   ".join([f"{p['qty']}x {p['size']}{p['unit']}" for p in calc["packs"]])#  (${p['price']}/pack)" 
                detail_data.append({"Categoria": cat_label,"Cantidades incluidas en plan": results[selected_plan]["included"].get(key,0), "Requeridas": totals.get(key,0), "Deficit": calc["deficit"], "Add-ons necesarios": packs_str})#:, "Costo": f"${calc['cost']}"})
                for p in calc["packs"]:
                    desc = f"Additional {p['eff_size']}{p['unit']} {cat_label}"
                    quote_rows.append({"Items & Description": desc, "Quantity": p["qty"], "Unit Price": p["price"], "Total Price": p["total_price"]})
            else:
                #detail_data.append({"Categoria": cat_label,"Cantidades incluidas en plan": results[selected_plan]["included"].get(key,0), "Requerido": totals.get(key,0), "Deficit": 0, "Add-ons necesarios": "✅ Cubierto", "Costo": "$0"})
                detail_data.append({"Categoria": cat_label,"Cantidades incluidas en plan": results[selected_plan]["included"].get(key,0), "Deficit": 0, "Add-ons necesarios": "✅ Cubierto"})
        
        st.dataframe(pd.DataFrame(detail_data), use_container_width=True)
        
        st.divider()
        st.subheader("📄 Tabla Quote Final")
        st.write("Esta tabla es la que se envía a Sales")
        
        df_quote = pd.DataFrame(quote_rows)
        total_quote = df_quote["Total Price"].sum()
        
        st.dataframe(df_quote, use_container_width=True, height=400)

        st.markdown(f"### 📩 Opciones de descarga")
        #st.markdown(f"### 💵 Total Estimado: **${total_quote} USD/mes** (pago anual)")
        st.caption("Precios de lista USD sin impuestos.")
        
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            csv = df_quote.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Descargar CSV Quote (con Logs y Runs)", csv, "quote_site24x7_final_completa.csv", "text/csv", use_container_width=True)
        with col_dl2:
            csv_comp = pd.DataFrame(detail_data).to_csv(index=False).encode('utf-8')
            st.download_button("📊 Descargar Detalle por Categoría", csv_comp, "detalle_categorias.csv", "text/csv", use_container_width=True)

st.markdown("---")
st.caption("Created by Diego Gonzalez - LATAM Technical Presales Consultant- Site24x7 - All Rights Reserved - 2026 ©")
