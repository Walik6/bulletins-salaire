import re, unicodedata, zipfile, io
import pdfplumber, openpyxl
import streamlit as st
from pypdf import PdfReader, PdfWriter
from rapidfuzz import process, fuzz
from datetime import datetime

# ── Mot de passe ───────────────────────────────────────────────────────────
MOT_DE_PASSE = st.secrets["mot_de_passe"]

def verifier_mdp():
    if "authentifie" not in st.session_state:
        st.session_state.authentifie = False

    if not st.session_state.authentifie:
        st.set_page_config(page_title="Connexion", page_icon="🔒", layout="centered")

        st.markdown("""
            <style>
            div[data-testid="stTextInput"] input { margin-bottom: 0 !important; }
            div[data-testid="stTextInput"] > label { margin-bottom: 4px !important; }
            .login-box {
                max-width: 380px;
                margin: 80px auto 0 auto;
                padding: 2rem 2rem 1.5rem;
                border: 1px solid #e0e0e0;
                border-radius: 12px;
                background: white;
            }
            </style>
        """, unsafe_allow_html=True)

        st.markdown('<div class="login-box">', unsafe_allow_html=True)
        st.markdown("### 🔒 Accès sécurisé")
        st.markdown("Entrez le mot de passe pour continuer.")
        mdp = st.text_input("Mot de passe", type="password", label_visibility="collapsed",
                            placeholder="Mot de passe")
        if st.button("Se connecter", type="primary", use_container_width=True):
            if mdp == MOT_DE_PASSE:
                st.session_state.authentifie = True
                st.rerun()
            else:
                st.error("❌ Mot de passe incorrect.")
        st.markdown('</div>', unsafe_allow_html=True)
        st.stop()

verifier_mdp()

# ── Fonctions ──────────────────────────────────────────────────────────────
MOIS_FR = {
    'janvier':'01_Janvier','fevrier':'02_Fevrier','février':'02_Fevrier',
    'mars':'03_Mars','avril':'04_Avril','mai':'05_Mai','juin':'06_Juin',
    'juillet':'07_Juillet','aout':'08_Aout','août':'08_Aout',
    'septembre':'09_Septembre','octobre':'10_Octobre',
    'novembre':'11_Novembre','decembre':'12_Decembre','décembre':'12_Decembre',
}

def normaliser(t):
    if not t: return ''
    t = unicodedata.normalize('NFD', str(t).strip())
    t = ''.join(c for c in t if unicodedata.category(c) != 'Mn').lower()
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9\s]', ' ', t)).strip()

def detecter_mois_annee(texte):
    n = normaliser(texte)
    mois = next((v for k,v in MOIS_FR.items() if re.search(r'\b'+re.escape(k)+r'\b', n)), 'Mois_Inconnu')
    a = re.search(r'\b(20\d{2})\b', texte)
    return mois, (a.group(1) if a else datetime.now().strftime('%Y'))

def charger_employes(fichier_excel, col_id, col_nom, col_prenom, ligne_debut):
    wb = openpyxl.load_workbook(fichier_excel, data_only=True)
    ws = wb.active
    employes = []
    for row in ws.iter_rows(min_row=ligne_debut, values_only=True):
        id_emp, nom, prenom = row[col_id-1], row[col_nom-1], row[col_prenom-1]
        if not id_emp or not nom or not prenom: continue
        nf = re.sub(r'[<>:"/\\|?*]', '_', f'{id_emp}_{nom}_{prenom}')
        employes.append({'nom_fichier': nf,
                         'cle':         normaliser(f'{nom} {prenom}'),
                         'cle_nom':     normaliser(str(nom)),
                         'cle_prenom':  normaliser(str(prenom))})
    return employes

def trouver_employe(texte, employes, seuil=72):
    n = normaliser(texte)
    for emp in employes:
        if emp['cle'] in n: return emp, 95, 'exact'
    for emp in employes:
        if emp['cle_nom'] in n and emp['cle_prenom'] in n: return emp, 90, 'séparé'
    cles = [emp['cle'] for emp in employes]
    mots = n.split()
    best_score, best_emp = 0, None
    for taille in [4, 3, 5, 2]:
        for i in range(len(mots) - taille + 1):
            seg = ' '.join(mots[i:i+taille])
            r = process.extractOne(seg, cles, scorer=fuzz.token_sort_ratio, score_cutoff=seuil)
            if r and r[1] > best_score:
                best_score, best_emp = r[1], employes[cles.index(r[0])]
    return (best_emp, best_score, 'fuzzy') if best_emp else (None, 0, '')

# ── Page config ────────────────────────────────────────────────────────────
st.set_page_config(page_title="Bulletins de Salaire", page_icon="📄", layout="centered")

# ── CSS global ─────────────────────────────────────────────────────────────
st.markdown("""
    <style>
    /* Bouton déconnexion discret en haut à droite */
    .logout-btn {
        position: fixed;
        top: 14px;
        right: 16px;
        z-index: 9999;
        background: transparent;
        border: 1px solid #ccc;
        border-radius: 20px;
        padding: 4px 14px;
        font-size: 13px;
        color: #555;
        cursor: pointer;
        white-space: nowrap;
    }
    .logout-btn:hover { background: #f5f5f5; border-color: #999; color: #222; }
    </style>
""", unsafe_allow_html=True)

# ── Bouton déconnexion ─────────────────────────────────────────────────────
if st.button("⎋  Déconnexion", key="logout"):
    st.session_state.authentifie = False
    st.rerun()

st.markdown("""
    <style>
    /* Rend le bouton déconnexion pill et positionné en haut à droite */
    div[data-testid="stButton"]:has(button[kind="secondary"]:first-child) {
        position: fixed;
        top: 12px;
        right: 16px;
        z-index: 9999;
    }
    div[data-testid="stButton"]:has(button[kind="secondary"]:first-child) button {
        border-radius: 20px !important;
        padding: 4px 16px !important;
        font-size: 13px !important;
        height: auto !important;
        min-height: 0 !important;
        white-space: nowrap !important;
        border: 1px solid #ccc !important;
        background: transparent !important;
        color: #555 !important;
    }
    div[data-testid="stButton"]:has(button[kind="secondary"]:first-child) button:hover {
        background: #f5f5f5 !important;
        color: #222 !important;
    }
    </style>
""", unsafe_allow_html=True)

# ── Titre ──────────────────────────────────────────────────────────────────
st.title("📄 Séparation Bulletins de Salaire")
st.markdown("Uploadez vos fichiers, ajustez les colonnes si besoin, puis cliquez sur **Lancer**.")
st.divider()

# ── Upload fichiers ────────────────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    fichier_pdf   = st.file_uploader("📕 Fichier PDF (bulletins)", type=["pdf"])
with col2:
    fichier_excel = st.file_uploader("📗 Fichier Excel (employés)", type=["xlsx", "xls"])

st.divider()

# ── Colonnes Excel ─────────────────────────────────────────────────────────
st.markdown("**Colonnes de votre Excel**")
c1, c2, c3, c4 = st.columns(4)
with c1: col_id      = st.number_input("Colonne ID",         min_value=1, max_value=20, value=1)
with c2: col_nom     = st.number_input("Colonne Nom",        min_value=1, max_value=20, value=2)
with c3: col_prenom  = st.number_input("Colonne Prénom",     min_value=1, max_value=20, value=3)
with c4: ligne_debut = st.number_input("1ère ligne données", min_value=1, max_value=20, value=2)

st.divider()

# ── Bouton lancer ──────────────────────────────────────────────────────────
if st.button("▶  Lancer le traitement", type="primary", use_container_width=True):
    if not fichier_pdf or not fichier_excel:
        st.warning("⚠️ Veuillez uploader le PDF et le fichier Excel avant de lancer.")
    else:
        with st.spinner("Traitement en cours..."):
            try:
                employes = charger_employes(fichier_excel, int(col_id), int(col_nom),
                                            int(col_prenom), int(ligne_debut))
                st.info(f"✅ {len(employes)} employés chargés depuis l'Excel")

                reader   = PdfReader(fichier_pdf)
                nb_pages = len(reader.pages)
                st.info(f"📄 {nb_pages} pages dans le PDF")

                with pdfplumber.open(fichier_pdf) as pdf_tmp:
                    texte_p1 = pdf_tmp.pages[0].extract_text() or ''
                mois, annee = detecter_mois_annee(texte_p1)
                st.info(f"📅 Période détectée : {mois} {annee}")

                zip_buffer = io.BytesIO()
                succes, echecs, log_lines = [], [], []
                progress = st.progress(0, text="Traitement des pages...")

                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                    with pdfplumber.open(fichier_pdf) as pdf_plumber:
                        for i, (page_pypdf, page_plumber) in enumerate(
                            zip(reader.pages, pdf_plumber.pages)
                        ):
                            num = i + 1
                            progress.progress(num / nb_pages, text=f"Page {num} / {nb_pages}")
                            texte = page_plumber.extract_text() or ''
                            emp, score, methode = trouver_employe(texte, employes)
                            writer = PdfWriter()
                            writer.add_page(page_pypdf)
                            buf = io.BytesIO()
                            writer.write(buf)
                            buf.seek(0)
                            if emp:
                                zf.writestr(f'{annee}/{mois}/{emp["nom_fichier"]}.pdf', buf.read())
                                log_lines.append(f'✅ Page {num:3d} → {emp["nom_fichier"]}.pdf')
                                succes.append(num)
                            else:
                                zf.writestr(f'{annee}/{mois}/INCONNU/page_{num:03d}.pdf', buf.read())
                                lignes = [l.strip() for l in texte.splitlines() if l.strip()]
                                apercu = ' | '.join(lignes[:5])
                                log_lines.append(f'⚠️  Page {num:3d} → INCONNU  ({apercu})')
                                echecs.append(num)

                progress.empty()
                st.divider()

                col_ok, col_err = st.columns(2)
                col_ok.metric("Pages traitées ✅", len(succes))
                col_err.metric("Pages inconnues ⚠️", len(echecs))

                if echecs:
                    st.warning(f"Pages non reconnues : {echecs}\n\nVérifiez le journal ci-dessous.")

                with st.expander("📋 Journal détaillé", expanded=bool(echecs)):
                    st.code('\n'.join(log_lines))

                zip_buffer.seek(0)
                nom_zip = f'Bulletins_{mois}_{annee}.zip'
                st.success("✅ Traitement terminé ! Cliquez ci-dessous pour télécharger.")
                st.download_button(
                    label=f"📥 Télécharger {nom_zip}",
                    data=zip_buffer.getvalue(),
                    file_name=nom_zip,
                    mime="application/zip",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ Erreur : {e}")
