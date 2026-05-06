# ── Sidebar ────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=600, show_spinner="Sincronizando con SharePoint...")
def fetch_sharepoint_files():
    if "sharepoint" not in st.secrets:
        return []
    
    cfg = st.secrets["sharepoint"]
    client_id = cfg.get("client_id")
    client_secret = cfg.get("client_secret")
    tenant_id = cfg.get("tenant_id")
    tenant_name = cfg.get("tenant_name")
    site_url = cfg.get("site_url")
    folder_path = cfg.get("folder_path")
    
    # Si sigue teniendo el ID de ejemplo, no intentamos conectar
    if client_id == "d1e8a3e0-bf9d-4f97-94e7-8c0c50f4d551" or not folder_path:
        return []
        
    try:
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.ConfidentialClientApplication(
            client_id, authority=authority, client_credential=client_secret
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        
        if "access_token" not in result:
            st.sidebar.error("Error autenticando con SharePoint. Revisa secrets.toml.")
            return []
            
        token = result["access_token"]
        headers = {"Authorization": f"Bearer {token}"}
        
        hostname = f"{tenant_name}.sharepoint.com"
        site_path = site_url.replace(f"https://{hostname}", "")
        site_api_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
        
        resp = requests.get(site_api_url, headers=headers)
        if not resp.ok:
            st.sidebar.error("No se pudo acceder al sitio de SharePoint.")
            return []
            
        site_id = resp.json().get("id")
        
        folder_api_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder_path}:/children"
        resp = requests.get(folder_api_url, headers=headers)
        if not resp.ok:
            st.sidebar.error(f"No se encontró la carpeta: {folder_path}")
            return []
            
        files_data = resp.json().get("value", [])
        downloaded = []
        
        for item in files_data:
            name = item.get("name", "")
            if name.endswith(".xlsx") and not name.startswith("~"):
                d_url = item.get("@microsoft.graph.downloadUrl")
                if d_url:
                    f_resp = requests.get(d_url)
                    if f_resp.ok:
                        f_obj = io.BytesIO(f_resp.content)
                        f_obj.name = name
                        downloaded.append(f_obj)
                        
        if downloaded:
            st.sidebar.success(f"✓ {len(downloaded)} archivos cargados de SharePoint")
        return downloaded
    except Exception as e:
        st.sidebar.error(f"Error de conexión: {e}")
        return []

def sidebar():
    with st.sidebar:
        import os
        import base64
        if os.path.exists("logo-pcc-hd.png"):
            with open("logo-pcc-hd.png", "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            st.markdown(f'''
            <div style="padding: 1rem 0 0 0; text-align: left;">
                <img src="data:image/png;base64,{b64}" style="width: 180px; max-width: 100%; height: auto; image-rendering: -webkit-optimize-contrast;">
            </div>
            ''', unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div style="padding:1.5rem 1rem 1rem; border-bottom:1px solid #E2E8F0; margin-bottom: 1rem;">
              <div style="font-size:1.4rem;font-weight:800;color:#D50032;letter-spacing:-0.5px;display:flex;align-items:center;gap:0.5rem;line-height:1.2;">
                Protección <span style="font-weight:400;color:#0F172A;">Catódica de Colombia</span>
              </div>
              <div style="font-size:0.75rem;color:#64748B;letter-spacing:0.08em;margin-top:4px;font-weight:600;">
                CATHODIC PROTECTION DASHBOARD
              </div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown('<p style="font-size:0.8rem;font-weight:600;color:#475569;margin:1rem 0 0.5rem;">CARGAR ARCHIVOS (OPCIONAL)</p>',
                    unsafe_allow_html=True)
        uploaded = st.file_uploader("Excel de FastField", type=["xlsx"],
                                     accept_multiple_files=True,
                                     label_visibility="collapsed")
        
        sp_files = fetch_sharepoint_files()
        all_files = (uploaded or []) + sp_files

        inspecciones = []
        if all_files:
            vistos = set()
            for f in all_files:
                if f.name in vistos: continue
                vistos.add(f.name)
                try:
                    data = load_excel(f); data["filename"] = f.name
                    inspecciones.append(data)
                except Exception as e:
                    st.error(f"{f.name[:22]}…: {e}")

        if inspecciones:
            st.markdown('<hr style="border-color:#E0E0E0;margin:0.6rem 0;">', unsafe_allow_html=True)
            by_tramo = {}
            for d in inspecciones:
                by_tramo.setdefault(d["meta"]["tramo"], []).append(d)
            for tramo, items in by_tramo.items():
                st.markdown(f'<p style="font-size:0.67rem;color:#999;text-transform:uppercase;'
                             f'letter-spacing:0.08em;margin:0.7rem 0 0.2rem;">{tramo}</p>',
                             unsafe_allow_html=True)
                for d in items:
                    icon = "📋" if d["tipo"]=="PAP" else "📡"
                    color = "#1565C0" if d["tipo"]=="PAP" else "#1B5E20"
                    st.markdown(f"""
                    <div style="background:white;border:1px solid #E2E8F0;border-radius:6px;padding:0.6rem 0.8rem;
                                margin:4px 0;border-left:4px solid {color};box-shadow:0 1px 2px rgba(0,0,0,0.02);">
                      <div style="font-size:0.85rem;font-weight:600;color:#0F172A;display:flex;align-items:center;gap:4px;">
                        {d['tipo']} <span style="font-weight:400;color:#64748B;font-size:0.8rem;">— {d['meta']['fecha']}</span></div>
                      <div style="font-size:0.75rem;color:#64748B;margin-top:4px;font-weight:500;">
                        {d['meta']['inspector']} • {len(d['df'])} pts</div>
                    </div>""", unsafe_allow_html=True)

        return inspecciones


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    inspecciones = sidebar()

    if not inspecciones:
        st.markdown(f"""
        <div style="margin-top:4rem;text-align:center;padding:4rem;background:white;border-radius:16px;border:1px dashed #CBD5E1;box-shadow:0 4px 6px rgba(0,0,0,0.02);animation:fadeUp 0.5s ease-out forwards;">
          <h2 style="color:#0F172A;margin-bottom:0.8rem;font-weight:700;">Dashboard de Inspecciones</h2>
          <p style="color:#64748B;font-size:1.1rem;line-height:1.6;max-width:500px;margin:0 auto;">
            Sube los archivos <b>.xlsx</b> exportados desde FastField usando el panel lateral.<br>
            La aplicación detectará automáticamente si es <b>PAP</b> o <b>DCVG</b> y organizará los datos por tramo.
          </p>
        </div>
        """, unsafe_allow_html=True)
        return

    pap_list  = [d for d in inspecciones if d["tipo"]=="PAP"]
    dcvg_list = [d for d in inspecciones if d["tipo"]=="DCVG"]

    # Si hay múltiples archivos → mostrar resumen primero
    if len(inspecciones) > 1:
        render_resumen(inspecciones)
        st.markdown('<div class="sec-div" style="margin:2rem 0;border-top:2px solid #EBEBEB;"></div>',
                    unsafe_allow_html=True)

    # PAP: selector si hay múltiples
    if pap_list:
        sel_pap = pap_list[0]
        if len(pap_list) > 1:
            opts = {f"{d['meta']['tramo']} — {d['meta']['inspector']} ({d['meta']['fecha']})": d
                    for d in pap_list}
            with st.sidebar:
                st.markdown('<hr style="border-color:#E0E0E0;margin:0.6rem 0;">', unsafe_allow_html=True)
                st.markdown('<p style="font-size:0.72rem;color:#888;margin-bottom:0.2rem;">SELECCIONAR PAP</p>',
                             unsafe_allow_html=True)
                sel_pap = opts[st.selectbox("PAP", list(opts.keys()),
                                             label_visibility="collapsed")]
        render_pap(sel_pap)

    # DCVG: selector si hay múltiples
    if dcvg_list:
        if pap_list:
            st.markdown('<div style="margin:1.5rem 0;border-top:2px solid #EBEBEB;"></div>',
                        unsafe_allow_html=True)
            pbi_title("Dashboard DCVG")
        
        sel_dcvg = dcvg_list[0]
        if len(dcvg_list) > 1:
            opts = {f"{d['meta']['tramo']} — {d['meta']['inspector']} ({d['meta']['fecha']})": d
                    for d in dcvg_list}
            with st.sidebar:
                st.markdown('<hr style="border-color:#E0E0E0;margin:0.6rem 0;">', unsafe_allow_html=True)
                st.markdown('<p style="font-size:0.72rem;color:#888;margin-bottom:0.2rem;">SELECCIONAR DCVG</p>',
                             unsafe_allow_html=True)
                sel_dcvg = opts[st.selectbox("DCVG", list(opts.keys()),
                                              label_visibility="collapsed")]
        render_dcvg(sel_dcvg)

if __name__ == "__main__":
    main()
