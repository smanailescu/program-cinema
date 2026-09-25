import streamlit as st
import mysql.connector
import random
import time
from datetime import datetime, timedelta

st.set_page_config(layout="wide", page_title="Panou Control Manager")

if "manager_logat" not in st.session_state:
    st.session_state.manager_logat = False
if "id_angajat_de_editat" not in st.session_state:
    st.session_state.id_angajat_de_editat = None

def get_conexiune():
    # Conectare Cloud (Streamlit Secrets)
    try:
        if "DB_HOST" in st.secrets:
            return mysql.connector.connect(
                host=st.secrets["DB_HOST"],
                port=int(st.secrets["DB_PORT"]),
                user=st.secrets["DB_USER"],
                password=st.secrets["DB_PASSWORD"],
                database=st.secrets["DB_NAME"]
            )
    except Exception:
        pass

    # Conectare locală (XAMPP)
    return mysql.connector.connect(
        host="localhost",
        user="root",
        password="",
        database="nume_angajati"
    )

# Verifică dacă o tură din șablon se potrivește cu opțiunea aleasă de angajat
def potrivire_tura(pref_tura, ora_start_td, ora_end_td):
    if not pref_tura:
        return True
    
    pref = pref_tura.strip().lower()
    
    if pref in ["indisponibil", "liber", "nu pot"]:
        return False
    if pref in ["oricand", "obligatoriu", "toata ziua", "all"]:
        return True

    start_h = ora_start_td.total_seconds() / 3600.0
    end_h = ora_end_td.total_seconds() / 3600.0
    if end_h <= start_h:
        end_h += 24.0

    if pref == "dimineata":
        return start_h <= 12.0
    elif pref == "middle":
        return 12.0 < start_h <= 17.0 and end_h <= 23.5
    elif pref in ["seara", "inchidere", "close"]:
        return start_h >= 16.0 or end_h >= 24.0

    return True

# Parola poate fi setată în Secrets sau rămâne implicit cinema2026
parola_corecta = "cinema2026"
try:
    if "PAROLA_MANAGER" in st.secrets:
        parola_corecta = st.secrets["PAROLA_MANAGER"]
except Exception:
    pass

if not st.session_state.manager_logat:
    st.title("👔 Acces Manager")
    parola_introdusa = st.text_input("Parolă:", type="password")
    if st.button("Intră în cont"):
        if parola_introdusa == parola_corecta: 
            st.session_state.manager_logat = True
            st.rerun()
        else:
            st.error("❌ Parolă incorectă!")
            
else:
    with st.sidebar:
        st.write("👔 **Mod Manager Activ**")
        st.divider()
        if st.button("🚪 Ieși din cont"):
            st.session_state.manager_logat = False
            st.rerun()

    try:
        con = get_conexiune()
        cursor = con.cursor()

        # Ne asigurăm că tabelul program_final are coloana mod_trafic
        try:
            cursor.execute("ALTER TABLE program_final ADD COLUMN mod_trafic VARCHAR(20) DEFAULT 'Verde'")
            con.commit()
        except Exception:
            pass

        tab1, tab2, tab3, tab4 = st.tabs(["📅 Programul Final", "🤖 Generare Automată", "👥 Echipa", "➕ Adaugă Angajat"])

        # ==========================================
        # TAB 1: VIZUALIZARE PROGRAM GRUPAT (GRID)
        # ==========================================
        with tab1:
            st.title("📅 Programul Întregii Săptămâni")
            
            cursor.execute("SELECT DISTINCT data_zi FROM program_final ORDER BY data_zi ASC")
            zile_generate = cursor.fetchall()
            
            if zile_generate:
                col_act1, col_act2 = st.columns([1, 2])
                with col_act1:
                    if st.button("🗑️ Resetează / Șterge Programul Generat", type="primary"):
                        cursor.execute("TRUNCATE TABLE program_final")
                        con.commit()
                        st.success("Programul a fost șters!")
                        time.sleep(1)
                        st.rerun()
                with col_act2:
                    zile_pe_rand = st.radio(
                        "📐 Câte zile vrei să vezi grupate pe un rând?",
                        options=[2, 3, 4],
                        index=1,
                        horizontal=True
                    )
                    
                st.write("---")

                traducere_zile = {
                    'Monday': 'Luni', 'Tuesday': 'Marti', 'Wednesday': 'Miercuri',
                    'Thursday': 'Joi', 'Friday': 'Vineri', 'Saturday': 'Sambata', 'Sunday': 'Duminica'
                }

                for idx_start in range(0, len(zile_generate), zile_pe_rand):
                    grup_zile = zile_generate[idx_start : idx_start + zile_pe_rand]
                    coloane_grid = st.columns(zile_pe_rand, gap="medium")

                    for col_idx, zi_tuple in enumerate(grup_zile):
                        zi_selectata = zi_tuple[0]
                        nume_zi_sapt = zi_selectata.strftime('%A')
                        zi_ro = traducere_zile.get(nume_zi_sapt, 'Luni')

                        with coloane_grid[col_idx]:
                            cursor.execute("""
                                SELECT zona, departament, ora_start, ora_end, mod_trafic, necesar_oameni 
                                FROM sabloane_necesar 
                                WHERE zi_saptamana = %s 
                                ORDER BY 
                                    FIELD(departament, 'VIP Host', 'VIP Plasator', 'Cafe', 'Bilete', 'Bar Mare', 'Bar IMAX', 'Plasator 1-9', 'Plasator 10-15', 'Plasator 16-20', 'Plasator IMAX'),
                                    FIELD(mod_trafic, 'Verde', 'Galben', 'Rosu'),
                                    ora_start ASC
                            """, (zi_ro,))
                            sabloane_zi = cursor.fetchall()

                            # Luăm toate alocările din ziua respectivă și le consumăm pe rând (fără dubluri)
                            cursor.execute("""
                                SELECT pf.departament, pf.ora_start, pf.ora_end, pf.mod_trafic, a.nume_angajat 
                                FROM program_final pf 
                                LEFT JOIN angajati a ON pf.id_angajat = a.id_angajat 
                                WHERE pf.data_zi = %s
                            """, (zi_selectata,))
                            alocari_zi = cursor.fetchall()

                            pool_alocari = {}
                            for al in alocari_zi:
                                cheie_al = (al[0], al[1], al[2], al[3])
                                if cheie_al not in pool_alocari:
                                    pool_alocari[cheie_al] = []
                                pool_alocari[cheie_al].append(al[4])
                            
                            if sabloane_zi:
                                html_table = "<div style='margin-bottom: 25px; box-shadow: 0 2px 6px rgba(0,0,0,0.15);'>"
                                html_table += "<table style='width:100%; border-collapse: collapse; text-align: center; font-family: sans-serif; font-size: 12.5px; color: black; border: 2px solid black;'>"
                                html_table += "<thead>"
                                html_table += f"<tr style='background-color: #262730; color: white; font-size: 14px; font-weight: bold; border: 2px solid black;'>"
                                html_table += f"<th colspan='4' style='padding: 8px; text-transform: uppercase; letter-spacing: 0.5px;'>📅 {zi_ro} — {zi_selectata.strftime('%d.%m.%Y')}</th>"
                                html_table += "</tr>"
                                html_table += "<tr style='background-color: #f4b084; color: black; font-weight: bold; border: 2px solid black;'>"
                                html_table += "<th style='border: 1px solid black; padding: 5px;'>Departament</th>"
                                html_table += "<th style='border: 1px solid black; padding: 5px;'>Start</th>"
                                html_table += "<th style='border: 1px solid black; padding: 5px;'>End</th>"
                                html_table += "<th style='border: 1px solid black; padding: 5px; width: 38%;'>Angajat</th>"
                                html_table += "</tr>"
                                html_table += "</thead>"
                                html_table += "<tbody>"

                                for sablon in sabloane_zi:
                                    zona, dep, o_start, o_end, mod, necesar = sablon[0], sablon[1], sablon[2], sablon[3], sablon[4], sablon[5]
                                    cheie_sablon = (dep, o_start, o_end, mod)

                                    if mod == 'Verde':
                                        bg_color = "#00b050"
                                    elif mod == 'Galben':
                                        bg_color = "#ffff00"
                                    else:
                                        bg_color = "#ff0000"

                                    start_str = (datetime.min + o_start).time().strftime('%H:%M')
                                    end_str = (datetime.min + o_end).time().strftime('%H:%M')

                                    for _ in range(necesar):
                                        nume_afisat = ""
                                        if cheie_sablon in pool_alocari and len(pool_alocari[cheie_sablon]) > 0:
                                            nume_extras = pool_alocari[cheie_sablon].pop(0)
                                            nume_afisat = nume_extras if nume_extras else "-"

                                        html_table += f"<tr style='background-color: {bg_color}; color: black; font-weight: bold; border: 1px solid black;'>"
                                        html_table += f"<td style='border: 1px solid black; padding: 4px; white-space: nowrap;'>{dep}</td>"
                                        html_table += f"<td style='border: 1px solid black; padding: 4px;'>{start_str}</td>"
                                        html_table += f"<td style='border: 1px solid black; padding: 4px;'>{end_str}</td>"
                                        html_table += f"<td style='border: 1px solid black; padding: 4px; background-color: white; color: black;'>{nume_afisat}</td>"
                                        html_table += "</tr>"
                                    
                                html_table += "</tbody></table></div>"
                                st.markdown(html_table, unsafe_allow_html=True)
                            else:
                                st.warning(f"Nu există șabloane pentru {zi_ro}.")
            else:
                st.info("Niciun program generat încă. Mergi la tab-ul 'Generare Automată' pentru a crea programul săptămânii.")

        # ==========================================
        # TAB 2: GENERARE AUTOMATĂ (2 - 5 TURE STRICT)
        # ==========================================
        with tab2:
            st.title("🤖 Generare Program Inteligent")
            st.info("⚖️ **Reguli active:** Fiecare angajat primește **minim 2 ture** și **maxim 5 ture** pe săptămână (maxim o tură pe zi).")
            
            col1, col2, col3 = st.columns(3)
            moduri = ["Verde (Bază)", "Galben (Mediu)", "Roșu (Aglomerat)"]
            
            mod_vip = col1.selectbox("Mod VIP:", moduri)
            mod_imax = col2.selectbox("Mod IMAX:", moduri)
            mod_cinema = col3.selectbox("Mod Cinema:", moduri)
            
            data_start = st.date_input("Data de început a săptămânii (Joi):")
            
            if st.button("🚀 Generează Programul", use_container_width=True, type="primary"):
                MIN_TURE = 2
                MAX_TURE = 5

                mapare_mod = {
                    "Verde (Bază)": ["Verde"],
                    "Galben (Mediu)": ["Verde", "Galben"],
                    "Roșu (Aglomerat)": ["Verde", "Galben", "Rosu"]
                }
                
                mapare_abilitati = {
                    "Bar Mare": "Bar",
                    "Bar IMAX": "Bar",
                    "Plasator 1-9": "Plasator",
                    "Plasator 10-15": "Plasator",
                    "Plasator 16-20": "Plasator",
                    "Plasator IMAX": "Plasator",
                    "VIP Host": "VIP",
                    "VIP Plasator": "VIP",
                    "Cafe": "Cafe",
                    "Bilete": "Casier" 
                }

                zile_saptamana = ['Joi', 'Vineri', 'Sambata', 'Duminica', 'Luni', 'Marti', 'Miercuri']
                zile_obligatorii = {'Sambata', 'Duminica', 'Marti'}
                data_end = data_start + timedelta(days=6)

                cursor.execute("TRUNCATE TABLE program_final")
                
                cursor.execute("""
                    SELECT a.id_angajat, a.nume_angajat, a.gen, a.rating, GROUP_CONCAT(ad.departament) 
                    FROM angajati a JOIN angajat_departament ad ON a.id_angajat = ad.id_angajat 
                    GROUP BY a.id_angajat
                """)
                toti_angajatii = cursor.fetchall()
                ang_dict = {
                    ang[0]: {
                        "id": ang[0],
                        "nume": ang[1],
                        "gen": ang[2],
                        "rating": float(ang[3]) if ang[3] else 5.0,
                        "deps": [d.strip() for d in ang[4].split(",")] if ang[4] else []
                    }
                    for ang in toti_angajatii
                }
                
                # Citim preferințele săptămânii curente din cand_pot_lucra
                cursor.execute("""
                    SELECT id_angajat, data_zi, tura 
                    FROM cand_pot_lucra 
                    WHERE data_zi BETWEEN %s AND %s
                """, (data_start, data_end))
                toate_preferintele = cursor.fetchall()
                
                pref_dict = {}
                angajati_cu_formular = set()
                for p in toate_preferintele:
                    id_a, d_zi, tura_pref = p[0], p[1], p[2]
                    angajati_cu_formular.add(id_a)
                    cheie = (id_a, d_zi)
                    if cheie not in pref_dict:
                        pref_dict[cheie] = []
                    pref_dict[cheie].append(tura_pref)

                # Verificăm în ce zile și pe ce ture poate lucra un angajat
                def poate_lucra(id_ang, data_zi, nume_zi, ora_s, ora_e, fortat_minim=False):
                    # Zilele de Sâmbătă, Duminică și Marți sunt mereu disponibile (obligatorii)
                    if nume_zi in zile_obligatorii:
                        return True, True
                    
                    # Dacă angajatul a completat formularul pe săptămâna aceasta:
                    if id_ang in angajati_cu_formular:
                         optiuni = pref_dict.get((id_ang, data_zi))
                        if not optiuni:
                            # Nu are înregistrare în acea zi flexibilă => a bifat "Liber"
                            return False, False
                        if any(potrivire_tura(opt, ora_s, ora_e) for opt in optiuni):
                            return True, True
                        # Dacă forțăm atingerea minimului de 2 ture și omul a zis că poate veni în acea zi
                        if fortat_minim:
                            return True, False
                        return False, False
                    else:
                        # Dacă nu a completat formularul, îl considerăm disponibil
                        return True, False

                # Colectăm toate locurile (sloturile) de muncă din întreaga săptămână
                filtre_vip = "','".join(mapare_mod[mod_vip])
                filtre_imax = "','".join(mapare_mod[mod_imax])
                filtre_cinema = "','".join(mapare_mod[mod_cinema])

                sloturi_saptamana = []
                for i, nume_zi in enumerate(zile_saptamana):
                    data_curenta = data_start + timedelta(days=i)
                    cursor.execute(f"""
                        SELECT id_necesar, departament, ora_start, ora_end, necesar_oameni, zona, mod_trafic 
                        FROM sabloane_necesar 
                        WHERE zi_saptamana = '{nume_zi}' AND (
                            (zona = 'VIP' AND mod_trafic IN ('{filtre_vip}')) OR
                            (zona = 'IMAX' AND mod_trafic IN ('{filtre_imax}')) OR
                            (zona = 'Cinema' AND mod_trafic IN ('{filtre_cinema}'))
                        ) ORDER BY ora_start ASC
                    """)
                    ture_zi = cursor.fetchall()
                    for tura in ture_zi:
                        dep_necesar = tura[1]
                        ora_s, ora_e = tura[2], tura[3]
                        oameni_necesari = tura[4]
                        mod_tura = tura[6]
                        abilitate = mapare_abilitati.get(dep_necesar, dep_necesar)
                        for _ in range(oameni_necesari):
                            sloturi_saptamana.append({
                                "data_zi": data_curenta,
                                "nume_zi": nume_zi,
                                "departament": dep_necesar,
                                "abilitate": abilitate,
                                "ora_start": ora_s,
                                "ora_end": ora_e,
                                "mod_trafic": mod_tura,
                                "id_angajat": None
                            })

                istoric_saptamanal = {id_a: 0 for id_a in ang_dict}
                zile_lucrate_angajat = {id_a: set() for id_a in ang_dict}

                # Calculăm câte zile disponibile are fiecare angajat în total (pentru a-i prioritiza pe cei cu opțiuni puține)
                zile_disp_count = {}
                for id_a in ang_dict:
                    cnt = 0
                    for i, n_zi in enumerate(zile_saptamana):
                        d_zi = data_start + timedelta(days=i)
                        if n_zi in zile_obligatorii or (id_a in angajati_cu_formular and (id_a, d_zi) in pref_dict) or (id_a not in angajati_cu_formular):
                            cnt += 1
                    zile_disp_count[id_a] = cnt

                # =========================================================
                # FAZA 1 & 2: ALOCAREA INIȚIALĂ (CU PLAFON MAXIM 5 TURE)
                # =========================================================
                for slot in sloturi_saptamana:
                    d_zi = slot["data_zi"]
                    n_zi = slot["nume_zi"]
                    abil = slot["abilitate"]
                    ora_s = slot["ora_start"]
                    ora_e = slot["ora_end"]

                    candidati = []
                    for id_a, info in ang_dict.items():
                        # REGULĂ STRICTĂ: Maxim 5 ture pe săptămână și maxim 1 tură pe zi
                        if istoric_saptamanal[id_a] >= MAX_TURE:
                            continue
                        if d_zi in zile_lucrate_angajat[id_a]:
                            continue
                        if abil not in info["deps"]:
                            continue

                        disp, pref_exact = poate_lucra(id_a, d_zi, n_zi, ora_s, ora_e, fortat_minim=False)
                        if not disp:
                            continue

                        candidati.append((id_a, pref_exact, info["rating"]))

                    if candidati:
                        # ORDINEA DE PRIORITATE:
                        # 1. Angajații care au SUB 2 TURE au prioritate absolută (0 ture înaintea celor cu 1 tură)
                        # 2. Cei cu mai puține zile disponibile în săptămână (ca să nu își piardă șansa la minim 2 ture)
                        # 3. Potrivirea exactă pe preferință
                        # 4. Echilibrarea turelor (cine are 2 ture primește înaintea celui cu 3 sau 4 ture)
                        # 5. Rating-ul (descrescător)
                        candidati.sort(key=lambda x: (
                            0 if istoric_saptamanal[x[0]] < MIN_TURE else 1,
                            istoric_saptamanal[x[0]],
                            zile_disp_count[x[0]],
                            not x[1],
                            -x[2]
                        ))

                        ales_id = candidati[0][0]
                        slot["id_angajat"] = ales_id
                        istoric_saptamanal[ales_id] += 1
                        zile_lucrate_angajat[ales_id].add(d_zi)

                # =========================================================
                # FAZA 3: CORECȚIE AUTOMATĂ (SWAP) PENTRU MINIM 2 TURE
                # =========================================================
                # Dacă cineva a rămas cu 0 sau 1 tură, îi căutăm loc liber sau preluăm o tură de la un coleg cu 3, 4 sau 5 ture
                for id_sub, info_sub in ang_dict.items():
                    incercari = 0
                    while istoric_saptamanal[id_sub] < MIN_TURE and incercari < 10:
                        incercari += 1
                        gasit_loc = False

                        # Pas 3A: Căutăm mai întâi un slot rămas neocupat compatibil
                        for slot in sloturi_saptamana:
                            if slot["id_angajat"] is None and slot["data_zi"] not in zile_lucrate_angajat[id_sub]:
                                if slot["abilitate"] in info_sub["deps"]:
                                    disp, _ = poate_lucra(id_sub, slot["data_zi"], slot["nume_zi"], slot["ora_start"], slot["ora_end"], fortat_minim=True)
                                    if disp:
                                        slot["id_angajat"] = id_sub
                                        istoric_saptamanal[id_sub] += 1
                                        zile_lucrate_angajat[id_sub].add(slot["data_zi"])
                                        gasit_loc = True
                                        break

                        if gasit_loc:
                            continue

                        # Pas 3B: Facem rocadă (Swap) cu un coleg care are deja > 2 ture (ex: 5, 4 sau 3 ture)
                        sloturi_eligibile_swap = []
                        for slot in sloturi_saptamana:
                            id_actual = slot["id_angajat"]
                            if id_actual is not None and id_actual != id_sub:
                                # Luăm doar de la colegii care au cel puțin 3 ture (ca să rămână și ei cu minim 2!)
                                if istoric_saptamanal[id_actual] > MIN_TURE and slot["data_zi"] not in zile_lucrate_angajat[id_sub]:
                                    if slot["abilitate"] in info_sub["deps"]:
                                        disp, pref_ex = poate_lucra(id_sub, slot["data_zi"], slot["nume_zi"], slot["ora_start"], slot["ora_end"], fortat_minim=True)
                                        if disp:
                                            sloturi_eligibile_swap.append((slot, istoric_saptamanal[id_actual], pref_ex))

                        if sloturi_eligibile_swap:
                            # Luăm tura de la colegul cu cele mai multe ture (ex: 5 ture -> scade la 4 ture)
                            sloturi_eligibile_swap.sort(key=lambda x: (-x[1], not x[2]))
                            slot_ales = sloturi_eligibile_swap[0][0]
                            id_vechi = slot_ales["id_angajat"]

                            # Efectuăm schimbul
                            istoric_saptamanal[id_vechi] -= 1
                            zile_lucrate_angajat[id_vechi].remove(slot_ales["data_zi"])

                            slot_ales["id_angajat"] = id_sub
                            istoric_saptamanal[id_sub] += 1
                            zile_lucrate_angajat[id_sub].add(slot_ales["data_zi"])
                        else:
                            break

                # Salvăm toate sloturile în baza de date
                for slot in sloturi_saptamana:
                    cursor.execute("""
                        INSERT INTO program_final (id_angajat, data_zi, departament, ora_start, ora_end, mod_trafic) 
                        VALUES (%s, %s, %s, %s, %s, %s)
                    """, (slot["id_angajat"], slot["data_zi"], slot["departament"], slot["ora_start"], slot["ora_end"], slot["mod_trafic"]))

                con.commit()
                st.success("✅ Programul a fost generat! Toți angajații eligibili au între minim 2 și maxim 5 ture.")
                time.sleep(1.5)
                st.rerun()

        # ==========================================
        # TAB 3: GESTIONARE ECHIPĂ & STATISTICI
        # ==========================================
        departamente_existente = ["Bar", "Plasator", "VIP", "Cafe", "Casier"] 

        with tab3:
            st.title("👥 Gestionare Echipa")
            
            cursor.execute("SELECT COUNT(id_angajat) FROM angajati")
            total_angajati = cursor.fetchone()[0]
            st.write(f"Număr total de angajați activi în sistem: **{total_angajati}**")
            
            with st.expander("📊 Vezi statistici ture angajați (Săptămâna curentă)", expanded=True):
                cursor.execute("""
                    SELECT a.id_angajat, a.nume_angajat, pf.departament 
                    FROM angajati a
                    LEFT JOIN program_final pf ON a.id_angajat = pf.id_angajat
                """)
                randuri_stat = cursor.fetchall()
                
                if randuri_stat:
                    statistici = {}
                    for id_a, nume, dep in randuri_stat:
                        if nume not in statistici:
                            statistici[nume] = {}
                        if dep:
                            statistici[nume][dep] = statistici[nume].get(dep, 0) + 1
                        
                    tabel_statistici = []
                    for nume, deps in statistici.items():
                        total_ture = sum(deps.values())
                        detalii_ture = ", ".join([f"{d} ({c}x)" for d, c in deps.items()]) if deps else "Fără ture alocate"
                        status_regula = "✅ OK (2-5 ture)" if 2 <= total_ture <= 5 else ("⚠️ Sub 2 ture" if total_ture < 2 else "❌ Peste 5 ture")
                        tabel_statistici.append({
                            "Angajat": nume,
                            "Total Ture": total_ture,
                            "Status": status_regula,
                            "Repartizare": detalii_ture
                        })
                        
                    tabel_statistici = sorted(tabel_statistici, key=lambda x: x["Total Ture"], reverse=True)
                    st.dataframe(tabel_statistici, use_container_width=True)
                else:
                    st.info("Nu există angajați sau ture alocate momentan.")

            st.write("---")
            
            col_stanga, col_dreapta = st.columns([1, 1.2], gap="large")
            with col_stanga:
                dep_ales = st.selectbox("Filtrează după abilitate:", departamente_existente)
                cursor.execute("""
                    SELECT a.id_angajat, a.nume_angajat, a.adresa_email, a.rating, a.gen 
                    FROM angajati a JOIN angajat_departament ad ON a.id_angajat = ad.id_angajat 
                    WHERE ad.departament = %s
                """, (dep_ales,))
                angajati_filtrati = cursor.fetchall()
                
                if angajati_filtrati:
                    st.write("---")
                    for rand in angajati_filtrati:
                        id_ang, nume, email, rating, gen = rand[0], rand[1], rand[2], rand[3], rand[4] if rand[4] else "Nespecificat"
                        c_detalii, c_buton = st.columns([4, 1])
                        c_detalii.markdown(f"**{nume}** ({gen})  \n⭐ {rating} | ✉️ {email}")
                        if c_buton.button("✏️", key=f"edit_{id_ang}"):
                            st.session_state.id_angajat_de_editat = id_ang
                            st.rerun()
                        st.write("---")
                else:
                    st.info("Niciun angajat alocat pentru această abilitate.")

            with col_dreapta:
                if st.session_state.id_angajat_de_editat is not None:
                    id_ales = st.session_state.id_angajat_de_editat
                    cursor.execute("SELECT nume_angajat, adresa_email, rating, gen FROM angajati WHERE id_angajat = %s", (id_ales,))
                    date_ang = cursor.fetchone()
                    if date_ang:
                        cursor.execute("SELECT departament FROM angajat_departament WHERE id_angajat = %s", (id_ales,))
                        deps_cur = [r[0] for r in cursor.fetchall()]
                        
                        st.subheader(f"Editează: {date_ang[0]}")
                        nou_email = st.text_input("Email:", value=date_ang[1])
                        nou_gen = st.selectbox("Gen:", ["M", "F"], index=0 if date_ang[3] == "M" else 1)
                        nou_rating = st.number_input("Rating (Max 5.0):", min_value=1.0, max_value=5.0, value=float(date_ang[2]), step=0.1)
                        nou_deps = st.multiselect("Abilități (Departamente):", departamente_existente, default=[d for d in deps_cur if d in departamente_existente])
                        
                        c1, c2 = st.columns(2)
                        if c1.button("💾 Salvează"):
                            cursor.execute("UPDATE angajati SET adresa_email = %s, rating = %s, gen = %s WHERE id_angajat = %s", (nou_email, nou_rating, nou_gen, id_ales))
                            cursor.execute("DELETE FROM angajat_departament WHERE id_angajat = %s", (id_ales,))
                            for dep in nou_deps:
                                cursor.execute("INSERT INTO angajat_departament (id_angajat, departament) VALUES (%s, %s)", (id_ales, dep))
                            con.commit()
                            st.session_state.id_angajat_de_editat = None 
                            st.rerun()
                        if c2.button("❌ Șterge"):
                            cursor.execute("DELETE FROM cand_pot_lucra WHERE id_angajat = %s", (id_ales,))
                            cursor.execute("DELETE FROM angajat_departament WHERE id_angajat = %s", (id_ales,))
                            cursor.execute("DELETE FROM program_final WHERE id_angajat = %s", (id_ales,))
                            cursor.execute("DELETE FROM angajati WHERE id_angajat = %s", (id_ales,))
                            con.commit()
                            st.session_state.id_angajat_de_editat = None 
                            st.rerun()

        # ==========================================
        # TAB 4: ADĂUGARE ANGAJAT
        # ==========================================
        with tab4:
            st.title("➕ Adaugă Angajat Nou")
            n_nume = st.text_input("Nume complet:")
            n_email = st.text_input("Adresa de email:")
            n_gen = st.selectbox("Gen:", ["M", "F"], key="gen_nou")
            n_rating = st.number_input("Rating inițial (1 la 5):", min_value=1.0, max_value=5.0, value=5.0, step=0.1)
            deps_alese = st.multiselect("Abilități (Departamente):", departamente_existente, key="deps_nou")
            
            if st.button("💾 Înregistrează Angajatul"):
                if n_nume and n_email and deps_alese:
                    nou_id = random.randint(1000, 99999)
                    cursor.execute("INSERT INTO angajati (id_angajat, nume_angajat, adresa_email, rating, gen) VALUES (%s, %s, %s, %s, %s)", (nou_id, n_nume, n_email, n_rating, n_gen))
                    for dep in deps_alese:
                        cursor.execute("INSERT INTO angajat_departament (id_angajat, departament) VALUES (%s, %s)", (nou_id, dep))
                    con.commit()
                    st.success("Adăugat!")
                    time.sleep(1)
                    st.rerun()

        cursor.close()
        con.close()

    except Exception as e:
        st.error(f"Eroare: {e}")
