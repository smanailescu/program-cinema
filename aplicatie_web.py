import streamlit as st
import mysql.connector
import datetime
import random
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(layout="centered", page_title="Portal Angajați - Cinema")

if "id_logat" not in st.session_state:
    st.session_state.id_logat = None
    st.session_state.nume_logat = None

TRADUCERE_ZILE = {
    'Monday': 'Luni', 'Tuesday': 'Marti', 'Wednesday': 'Miercuri',
    'Thursday': 'Joi', 'Friday': 'Vineri', 'Saturday': 'Sambata', 'Sunday': 'Duminica'
}


def get_conexiune():
    # Dacă aplicația este urcată online, ia datele din Cloud:
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

    # Altfel, dacă o rulezi de pe PC-ul tău, se conectează automat la XAMPP:
    return mysql.connector.connect(
        host="localhost",
        user="root",
        password="",
        database="nume_angajati"
    )


# Construiește fișierul Excel cu programul complet (identic cu cel de la Manager)
def genereaza_excel_program(date_zile, zile_pe_rand=3):
    wb = Workbook()
    ws = wb.active
    ws.title = "Program Saptamanal"

    fill_titlu_zi = PatternFill(start_color="262730", end_color="262730", fill_type="solid")
    fill_header = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
    fill_verde = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
    fill_galben = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    fill_rosu = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    fill_alb = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    font_titlu_zi = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    font_bold_negru = Font(name="Calibri", size=10, bold=True, color="000000")
    align_center = Alignment(horizontal="center", vertical="center")

    thin_side = Side(border_style="thin", color="000000")
    med_side = Side(border_style="medium", color="000000")
    border_celula = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_header = Border(left=thin_side, right=thin_side, top=med_side, bottom=med_side)

    rand_start_grup = 2

    for idx_start in range(0, len(date_zile), zile_pe_rand):
        grup = date_zile[idx_start: idx_start + zile_pe_rand]
        max_rand_in_grup = rand_start_grup

        for col_idx, info_zi in enumerate(grup):
            col_baza = col_idx * 5 + 1
            rand_curent = rand_start_grup

            ws.column_dimensions[get_column_letter(col_baza)].width = 18
            ws.column_dimensions[get_column_letter(col_baza + 1)].width = 9
            ws.column_dimensions[get_column_letter(col_baza + 2)].width = 9
            ws.column_dimensions[get_column_letter(col_baza + 3)].width = 24
            ws.column_dimensions[get_column_letter(col_baza + 4)].width = 4

            ws.merge_cells(
                start_row=rand_curent, start_column=col_baza,
                end_row=rand_curent, end_column=col_baza + 3
            )
            titlu_text = f"{info_zi['zi_ro'].upper()} — {info_zi['data_str']}"
            for c in range(col_baza, col_baza + 4):
                cel = ws.cell(row=rand_curent, column=c)
                if c == col_baza:
                    cel.value = titlu_text
                cel.fill = fill_titlu_zi
                cel.font = font_titlu_zi
                cel.alignment = align_center
                cel.border = border_header
            ws.row_dimensions[rand_curent].height = 24
            rand_curent += 1

            headers = ["Departament", "Start", "End", "Angajat"]
            for offset, h_text in enumerate(headers):
                cel = ws.cell(row=rand_curent, column=col_baza + offset, value=h_text)
                cel.fill = fill_header
                cel.font = font_bold_negru
                cel.alignment = align_center
                cel.border = border_header
            ws.row_dimensions[rand_curent].height = 20
            rand_curent += 1

            for rand_tura in info_zi["randuri"]:
                dep, start_str, end_str, nume_ang, mod = rand_tura

                if mod == "Verde":
                    fill_tura = fill_verde
                elif mod == "Galben":
                    fill_tura = fill_galben
                else:
                    fill_tura = fill_rosu

                valori = [dep, start_str, end_str, nume_ang]
                for offset, val in enumerate(valori):
                    cel = ws.cell(row=rand_curent, column=col_baza + offset, value=val)
                    cel.font = font_bold_negru
                    cel.alignment = align_center
                    cel.border = border_celula
                    cel.fill = fill_alb if offset == 3 else fill_tura

                ws.row_dimensions[rand_curent].height = 18
                rand_curent += 1

            if rand_curent > max_rand_in_grup:
                max_rand_in_grup = rand_curent

        rand_start_grup = max_rand_in_grup + 2

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# --- ECRANUL DE LOGARE ---
if st.session_state.id_logat is None:
    st.title("🔐 Logare Angajați")

    email_introdus = st.text_input("Email:")

    if st.button("Intră în cont", type="primary"):
        if email_introdus.strip() != "":
            try:
                con = get_conexiune()
                cursor = con.cursor()

                comanda = "SELECT id_angajat, nume_angajat FROM angajati WHERE adresa_email = %s"
                cursor.execute(comanda, (email_introdus.strip(),))
                rezultat = cursor.fetchone()

                if rezultat:
                    st.session_state.id_logat = rezultat[0]
                    st.session_state.nume_logat = rezultat[1]
                    st.rerun()
                else:
                    st.error("❌ Email-ul nu a fost găsit în baza de date!")

                cursor.close()
                con.close()
            except Exception as e:
                st.error(f"Eroare la conectare: {e}")
        else:
            st.warning("Te rog să introduci un email.")

# --- ECRANUL PRINCIPAL (DUPĂ LOGARE) ---
else:
    with st.sidebar:
        st.write(f"👤 Logat ca: **{st.session_state.nume_logat}**")
        st.divider()
        if st.button("🚪 Deconectare"):
            st.session_state.id_logat = None
            st.session_state.nume_logat = None
            st.rerun()

    st.title(f"👋 Salut, {st.session_state.nume_logat}!")

    tab_program, tab_preferinte = st.tabs([
        "📅 Turele Mele & Program Complet",
        "✍️ Alege Preferințele"
    ])

    # =========================================================================
    # TAB 1: TURELE ANGAJATULUI & DESCĂRCARE PROGRAM COMPLET
    # =========================================================================
    with tab_program:
        try:
            con = get_conexiune()
            cursor = con.cursor()

            cursor.execute("SELECT DISTINCT data_zi FROM program_final ORDER BY data_zi ASC")
            zile_generate = cursor.fetchall()

            if zile_generate:
                prima_zi = zile_generate[0][0]
                ultima_zi = zile_generate[-1][0]

                st.subheader(f"📅 Programul Publicat: {prima_zi.strftime('%d.%m.%Y')} — {ultima_zi.strftime('%d.%m.%Y')}")

                # 1. Căutăm turele alocate special angajatului logat
                cursor.execute("""
                    SELECT data_zi, departament, ora_start, ora_end
                    FROM program_final
                    WHERE id_angajat = %s
                    ORDER BY data_zi ASC, ora_start ASC
                """, (st.session_state.id_logat,))
                turele_mele = cursor.fetchall()

                st.markdown("#### 🎯 Turele tale în această săptămână:")
                if turele_mele:
                    st.success(f"Ai **{len(turele_mele)} ture** repartizate în această săptămână:")

                    html_mele = "<table style='width:100%; border-collapse: collapse; text-align: center; font-family: sans-serif; font-size: 14px; margin-bottom: 20px; border: 2px solid #262730;'>"
                    html_mele += "<thead style='background-color: #262730; color: white;'>"
                    html_mele += "<tr>"
                    html_mele += "<th style='padding: 10px; border: 1px solid #444;'>Ziua</th>"
                    html_mele += "<th style='padding: 10px; border: 1px solid #444;'>Data</th>"
                    html_mele += "<th style='padding: 10px; border: 1px solid #444;'>Departament</th>"
                    html_mele += "<th style='padding: 10px; border: 1px solid #444;'>Interval Orar</th>"
                    html_mele += "</tr></thead><tbody>"

                    for t in turele_mele:
                        d_zi, dep, o_s, o_e = t[0], t[1], t[2], t[3]
                        zi_ro = TRADUCERE_ZILE.get(d_zi.strftime('%A'), '')
                        start_str = (datetime.datetime.min + o_s).time().strftime('%H:%M')
                        end_str = (datetime.datetime.min + o_e).time().strftime('%H:%M')

                        html_mele += "<tr style='background-color: #f9f9f9; color: black; font-weight: bold;'>"
                        html_mele += f"<td style='padding: 8px; border: 1px solid #ccc;'>{zi_ro}</td>"
                        html_mele += f"<td style='padding: 8px; border: 1px solid #ccc;'>{d_zi.strftime('%d.%m.%Y')}</td>"
                        html_mele += f"<td style='padding: 8px; border: 1px solid #ccc; color: #008037;'>{dep}</td>"
                        html_mele += f"<td style='padding: 8px; border: 1px solid #ccc;'>{start_str} - {end_str}</td>"
                        html_mele += "</tr>"

                    html_mele += "</tbody></table>"
                    st.markdown(html_mele, unsafe_allow_html=True)
                else:
                    st.info("ℹ️ Nu ai nicio tură alocată în programul publicat pentru această săptămână.")

                st.divider()

                # 2. Construim datele pentru întregul program (pentru descărcarea Excel și vizualizare)
                date_zile_procesate = []
                for zi_tuple in zile_generate:
                    zi_selectata = zi_tuple[0]
                    nume_zi_sapt = zi_selectata.strftime('%A')
                    zi_ro = TRADUCERE_ZILE.get(nume_zi_sapt, 'Luni')

                    cursor.execute("""
                        SELECT zona, departament, ora_start, ora_end, mod_trafic, necesar_oameni
                        FROM sabloane_necesar
                        WHERE zi_saptamana = %s
                        ORDER BY FIELD(departament, 'VIP Host', 'VIP Plasator', 'Cafe', 'Bilete', 'Bar Mare',
                                       'Bar IMAX', 'Plasator 1-9', 'Plasator 10-15', 'Plasator 16-20',
                                       'Plasator IMAX'),
                                 FIELD(mod_trafic, 'Verde', 'Galben', 'Rosu'),
                                 ora_start ASC
                    """, (zi_ro,))
                    sabloane_zi = cursor.fetchall()

                    cursor.execute("""
                        SELECT pf.departament, pf.ora_start, pf.ora_end, pf.mod_trafic, a.nume_angajat
                        FROM program_final pf
                        LEFT JOIN angajati a ON pf.id_angajat = a.id_angajat
                        WHERE pf.data_zi = %s
                    """, (zi_selectata,))
                    alocati_zi = cursor.fetchall()

                    alocati_map = {}
                    for al in alocati_zi:
                        cheie_al = (al[0], al[1], al[2], al[3])
                        if cheie_al not in alocati_map:
                            alocati_map[cheie_al] = []
                        alocati_map[cheie_al].append(al[4] if al[4] else "-")

                    randuri_zi = []
                    for sablon in sabloane_zi:
                        zona, dep, o_start, o_end, mod, necesar = (
                            sablon[0], sablon[1], sablon[2], sablon[3], sablon[4], sablon[5]
                        )
                        cheie_s = (dep, o_start, o_end, mod)
                        start_str = (datetime.datetime.min + o_start).time().strftime('%H:%M')
                        end_str = (datetime.datetime.min + o_end).time().strftime('%H:%M')

                        if cheie_s in alocati_map:
                            lista_nume = alocati_map[cheie_s]
                            for _ in range(necesar):
                                nume_angajat = lista_nume.pop(0) if lista_nume else "-"
                                randuri_zi.append((dep, start_str, end_str, nume_angajat, mod))
                        else:
                            for _ in range(necesar):
                                randuri_zi.append((dep, start_str, end_str, "", mod))

                    date_zile_procesate.append({
                        "zi_selectata": zi_selectata,
                        "zi_ro": zi_ro,
                        "data_str": zi_selectata.strftime('%d.%m.%Y'),
                        "randuri": randuri_zi
                    })

                st.markdown("#### 📋 Programul Complet al Echipei")
                excel_bytes = genereaza_excel_program(date_zile_procesate, zile_pe_rand=3)
                prima_zi_str = prima_zi.strftime('%d_%m_%Y')

                st.download_button(
                    label="📥 Descarcă Programul Complet (.xlsx)",
                    data=excel_bytes,
                    file_name=f"Program_Cinema_{prima_zi_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )

                with st.expander("👀 Vezi programul complet direct pe site"):
                    for info_zi in date_zile_procesate:
                        if info_zi["randuri"]:
                            html_table = "<div style='margin-bottom: 20px;'>"
                            html_table += "<table style='width:100%; border-collapse: collapse; text-align: center; font-family: sans-serif; font-size: 12.5px; color: black; border: 2px solid black;'>"
                            html_table += "<thead>"
                            html_table += "<tr style='background-color: #262730; color: white; font-size: 14px; font-weight: bold; border: 2px solid black;'>"
                            html_table += f"<th colspan='4' style='padding: 8px; text-transform: uppercase;'>📅 {info_zi['zi_ro']} — {info_zi['data_str']}</th>"
                            html_table += "</tr>"
                            html_table += "<tr style='background-color: #f4b084; color: black; font-weight: bold; border: 2px solid black;'>"
                            html_table += "<th style='border: 1px solid black; padding: 5px;'>Departament</th>"
                            html_table += "<th style='border: 1px solid black; padding: 5px;'>Start</th>"
                            html_table += "<th style='border: 1px solid black; padding: 5px;'>End</th>"
                            html_table += "<th style='border: 1px solid black; padding: 5px; width: 40%;'>Angajat</th>"
                            html_table += "</tr></thead><tbody>"

                            for rand_tura in info_zi["randuri"]:
                                dep, start_str, end_str, nume_angajat, mod = rand_tura
                                if mod == 'Verde':
                                    bg_color = "#00b050"
                                elif mod == 'Galben':
                                    bg_color = "#ffff00"
                                else:
                                    bg_color = "#ff0000"

                                # Evidențiem cu galben deschis rândul pe care apare angajatul logat
                                bg_nume = "#d4edda" if nume_angajat == st.session_state.nume_logat else "white"

                                html_table += f"<tr style='background-color: {bg_color}; color: black; font-weight: bold; border: 1px solid black;'>"
                                html_table += f"<td style='border: 1px solid black; padding: 4px;'>{dep}</td>"
                                html_table += f"<td style='border: 1px solid black; padding: 4px;'>{start_str}</td>"
                                html_table += f"<td style='border: 1px solid black; padding: 4px;'>{end_str}</td>"
                                html_table += f"<td style='border: 1px solid black; padding: 4px; background-color: {bg_nume}; color: black;'>{nume_angajat}</td>"
                                html_table += "</tr>"

                            html_table += "</tbody></table></div>"
                            st.markdown(html_table, unsafe_allow_html=True)
            else:
                st.info("⏳ Managerul nu a publicat încă programul final pentru această săptămână.")

            cursor.close()
            con.close()
        except Exception as e:
            st.error(f"Eroare la încărcarea programului: {e}")

    # =========================================================================
    # TAB 2: ALEGEREA PREFERINȚELOR PENTRU SĂPTĂMÂNA URMĂTOARE
    # =========================================================================
    with tab_preferinte:
        st.subheader("🎥 Alege-ți disponibilitatea")

        acum = datetime.datetime.now()
        azi = acum.date()

        zile_pana_la_joi = (3 - azi.weekday()) % 7
        if zile_pana_la_joi == 0:
            zile_pana_la_joi = 7

        data_joi = azi + datetime.timedelta(days=zile_pana_la_joi)
        data_vineri = data_joi + datetime.timedelta(days=1)
        data_sambata = data_joi + datetime.timedelta(days=2)
        data_duminica = data_joi + datetime.timedelta(days=3)
        data_luni = data_joi + datetime.timedelta(days=4)
        data_marti = data_joi + datetime.timedelta(days=5)
        data_miercuri = data_joi + datetime.timedelta(days=6)

        st.write(f"### 📅 Săptămâna: *{data_joi.strftime('%d.%m.%Y')}  -  {data_miercuri.strftime('%d.%m.%Y')}*")

        try:
            con = get_conexiune()
            cursor = con.cursor()

            comanda_istoric = """
                SELECT data_zi, tura 
                FROM cand_pot_lucra 
                WHERE id_angajat = %s AND data_zi BETWEEN %s AND %s 
                ORDER BY data_zi
            """
            cursor.execute(comanda_istoric, (st.session_state.id_logat, data_joi, data_miercuri))
            optiuni_existente = cursor.fetchall()

            if optiuni_existente:
                with st.expander("👀 Vezi preferințele tale salvate pentru această săptămână", expanded=True):
                    for rand in optiuni_existente:
                        zi_nume = TRADUCERE_ZILE.get(rand[0].strftime('%A'), '')
                        st.write(f"• **{zi_nume} ({rand[0].strftime('%d-%m-%Y')})**: {rand[1].capitalize()}")
            else:
                st.info("Nu ai introdus încă nicio preferință pentru această săptămână.")

            cursor.close()
            con.close()
        except Exception as e:
            st.error(f"Eroare la verificarea istoricului: {e}")

        st.divider()

        deadline_depasit = (azi.weekday() == 2 and acum.hour >= 12)

        if deadline_depasit:
            st.error(
                "🔒 **Termenul a expirat!** Este trecut de Miercuri ora 12:00. Formularul a fost blocat. Te rugăm să contactezi managerul pentru modificări."
            )
        else:
            st.warning("⏳ **Reminder:** Poți adăuga sau modifica programul până cel târziu **Miercuri la ora 12:00**.")
            st.info("🔒 **Sâmbătă, Duminică și Marți** sunt zile cu prezență obligatorie. Se salvează automat ca 'oricând'.")
            st.divider()

            optiuni_ture = ["oricand", "dimineata", "middle", "seara"]

            st.write(f"**JOI** ({data_joi.strftime('%d-%m-%Y')})")
            col1, col2 = st.columns(2)
            disp_joi = col1.radio("Disponibilitate joi:", ["Liber", "Pot veni"], key="radio_joi")
            tura_joi = col2.selectbox("Alege tura:", optiuni_ture, key="box_joi") if disp_joi == "Pot veni" else "liber"
            st.divider()

            st.write(f"**VINERI** ({data_vineri.strftime('%d-%m-%Y')})")
            col3, col4 = st.columns(2)
            disp_vineri = col3.radio("Disponibilitate vineri:", ["Liber", "Pot veni"], key="radio_vin")
            tura_vineri = col4.selectbox("Alege tura:", optiuni_ture, key="box_vin") if disp_vineri == "Pot veni" else "liber"
            st.divider()

            st.write(f"**LUNI** ({data_luni.strftime('%d-%m-%Y')})")
            col5, col6 = st.columns(2)
            disp_luni = col5.radio("Disponibilitate luni:", ["Liber", "Pot veni"], key="radio_luni")
            tura_luni = col6.selectbox("Alege tura:", optiuni_ture, key="box_luni") if disp_luni == "Pot veni" else "liber"
            st.divider()

            st.write(f"**MIERCURI** ({data_miercuri.strftime('%d-%m-%Y')})")
            col7, col8 = st.columns(2)
            disp_miercuri = col7.radio("Disponibilitate miercuri:", ["Liber", "Pot veni"], key="radio_mie")
            tura_miercuri = col8.selectbox("Alege tura:", optiuni_ture, key="box_mie") if disp_miercuri == "Pot veni" else "liber"
            st.divider()

            if st.button("💾 Trimite preferințele", type="primary", use_container_width=True):
                try:
                    con = get_conexiune()
                    cursor = con.cursor()

                    comanda_delete = "DELETE FROM cand_pot_lucra WHERE id_angajat = %s AND data_zi BETWEEN %s AND %s"
                    cursor.execute(comanda_delete, (st.session_state.id_logat, data_joi, data_miercuri))

                    comanda_insert = """
                        INSERT INTO cand_pot_lucra (id_disponibilitate, id_angajat, data_zi, tura) 
                        VALUES (%s, %s, %s, %s)
                    """

                    # Zilele obligatorii
                    cursor.execute(comanda_insert, (random.randint(10000, 999999), st.session_state.id_logat, data_sambata, "oricand"))
                    cursor.execute(comanda_insert, (random.randint(10000, 999999), st.session_state.id_logat, data_duminica, "oricand"))
                    cursor.execute(comanda_insert, (random.randint(10000, 999999), st.session_state.id_logat, data_marti, "oricand"))

                    # Zilele opționale
                    cursor.execute(comanda_insert, (random.randint(10000, 999999), st.session_state.id_logat, data_joi, tura_joi))
                    cursor.execute(comanda_insert, (random.randint(10000, 999999), st.session_state.id_logat, data_vineri, tura_vineri))
                    cursor.execute(comanda_insert, (random.randint(10000, 999999), st.session_state.id_logat, data_luni, tura_luni))
                    cursor.execute(comanda_insert, (random.randint(10000, 999999), st.session_state.id_logat, data_miercuri, tura_miercuri))

                    con.commit()
                    st.success("✅ Opțiunile au fost salvate!")

                    cursor.close()
                    con.close()

                    st.rerun()
                except Exception as e:
                    st.error(f"Eroare la salvare: {e}")