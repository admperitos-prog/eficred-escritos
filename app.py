from flask import Flask, request, jsonify
from flask_cors import CORS
import zipfile, re, os, io, base64, tempfile
from lxml import etree
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2 import service_account

app = Flask(__name__)
CORS(app)

# ── Google Drive setup ──────────────────────────────────────────
SCOPES = ['https://www.googleapis.com/auth/drive']

def get_drive_service():
    creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if not creds_json:
        raise Exception('GOOGLE_CREDENTIALS_JSON no configurado')
    import json
    creds_info = json.loads(creds_json)
    creds = service_account.Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

def descargar_template(drive, file_id):
    """Descarga un template del Drive como bytes .docx"""
    # Verificar el tipo de archivo
    meta = drive.files().get(fileId=file_id, fields='mimeType,name', supportsAllDrives=True).execute()
    mime = meta.get('mimeType', '')
    
    if mime == 'application/vnd.google-apps.document':
        # Es un Google Doc — exportar como .docx
        request = drive.files().export_media(
            fileId=file_id,
            mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    else:
        # Es un .docx nativo — descargar directo
        request = drive.files().get_media(fileId=file_id, supportsAllDrives=True)
    
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf.read()

def subir_pdf(drive, folder_id, nombre, pdf_bytes):
    """Sube un PDF al Drive y devuelve el ID"""
    buf = io.BytesIO(pdf_bytes)
    media = MediaIoBaseUpload(buf, mimetype='application/pdf')
    # supportsAllDrives=True permite subir a carpetas compartidas
    file_meta = {'name': nombre, 'parents': [folder_id]}
    resultado = drive.files().create(
        body=file_meta,
        media_body=media,
        fields='id,webViewLink',
        supportsAllDrives=True
    ).execute()
    return resultado

# ── IDs de los templates en Drive ──────────────────────────────
TEMPLATES = {
    'informa_ope':  os.environ.get('TPL_INFORMA_OPE',  '1rwuXVlE3dyWqSHe7ivMPahNpB_L6reoA0x0b9pe3QOQ'),
    'liq_apod':     os.environ.get('TPL_LIQ_APOD',     '18IDwJjSW6_Rt2QORzRTAdSrxxONv0aJ7Ir-7IRqx8us'),
    'liq_patr':     os.environ.get('TPL_LIQ_PATR',     '11LyQGVLu6opFqy1LgfqRqvpOCZx9vwWOD5cJl4XnoNQ'),
    'aprob_liq':    os.environ.get('TPL_APROB_LIQ',    '1mZlc9pAghDb5gfa6xanOrWNrqA3gDwOHPgmCRH6iRlg'),
    'aprob_ope':    os.environ.get('TPL_APROB_OPE',    '128px-ptOggJpKb_LelIm38XYZU3T8DU9-YRhhGJerzk'),
    'amplia':       os.environ.get('TPL_AMPLIA',       '1SN00MVBdDxjrFeSb8hyQ9zDGGWjTFsSZ7XrVHbpWV3M'),
    'oficio':       os.environ.get('TPL_OFICIO',       '10eliAnuSN9G1n7rb8C5Wy95mw7kMTuB5S6RXQgU4KFA'),
    'oficio_reit':  os.environ.get('TPL_OFICIO_REIT',  '11PU44Fsu3R0rf5YUyvYFIOAG1pb9M0sTD-TK0-7EbBo'),
    'trance_apod':  os.environ.get('TPL_TRANCE_APOD',  '1KX5w-XRmlzPxOQCFCMFf_XyXamwg67BoAX26zn6hkVU'),
    'trance_patr':  os.environ.get('TPL_TRANCE_PATR',  '1fOkkODMFVPaTTWOdjsEsFNTSKw4ljvWzbc0ib4OTE0o'),
    'modifica_embargo': os.environ.get('TPL_MODIFICA', '1Cjp9QIydc1j0juFnM9_cTNAaFnxprXjB'),
}

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def tag(name): return f'{{{W}}}{name}'

def make_run(texto, bold=False):
    r = etree.Element(tag('r'))
    rpr = etree.SubElement(r, tag('rPr'))
    if bold:
        etree.SubElement(rpr, tag('b'))
        etree.SubElement(rpr, tag('bCs'))
    t = etree.SubElement(r, tag('t'))
    t.text = texto
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r

def reemplazar_runs(p, runs_data):
    """Reemplaza todos los runs de un párrafo con nuevos runs"""
    for r in list(p.findall(tag('r'))):
        p.remove(r)
    for texto, bold in runs_data:
        if texto:
            p.append(make_run(texto, bold))

def procesar_docx(template_bytes, tipo, d):
    """Procesa el .docx del template con los datos y devuelve bytes del docx modificado"""
    
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as z:
        archivos = {n: z.read(n) for n in z.namelist()}
    
    xml_bytes = archivos['word/document.xml']
    root = etree.fromstring(xml_bytes)
    body = root.find(tag('body'))
    parrafos = list(body.findall(tag('p')))
    
    p       = d['perito']
    car     = d.get('caratula', '')
    expte   = d.get('expte', '')
    rep     = p.get('rep', 'patrocinio')
    esApod  = rep == 'apoderado'
    
    car_fmt  = f'\u201c{car}\u201d'
    ex_fmt   = f'Expte. N\u00b0 {expte}'
    car_ex   = f'{car_fmt} {ex_fmt}'

    # ── Reemplazos genéricos de texto en todos los runs ──
    reemplazos_texto = build_reemplazos(tipo, d)
    
    for para in parrafos:
        for r in para.findall(f'.//{tag("r")}'):
            for t in r.findall(tag('t')):
                if t.text:
                    for buscar, reemplazar in reemplazos_texto.items():
                        if buscar in t.text:
                            t.text = t.text.replace(buscar, reemplazar)

    # ── Reemplazar párrafo del encabezado ──
    for para in parrafos:
        texts = ''.join(t.text or '' for t in para.findall(f'.//{tag("t")}'))
        if 'patrocinio letrado' in texts or ('Ingeniero Mec' in texts and 'GABRIEL' in texts) or ('FRIEDRICH' in texts and 'patrocinio' in texts):
            if esApod:
                runs_data = [
                    ('MARTÍN MANGINI', True),
                    (' y ', False),
                    ('FAUSTO L. GRIPPALDI', True),
                    (', abogados, conforme la participación acordada en estos autos caratulados: ', False),
                    (car_ex, True),
                    (', ante V.S. nos presentamos y respetuosamente decimos que:', False),
                ]
            else:
                runs_data = [
                    (p['nombre'], True),
                    (f', {p["prof"]}, D.N.I. N\u00ba {p["dni"]}, con el patrocinio letrado de los Dres. MARTÍN MANGINI, abogado, Matrícula C.A.E.R. N° 11.042, T° I F° 299, y FAUSTO L. GRIPPALDI, abogado, Matrícula C.A.E.R. N° 10.555, T° I F° 286, por la participación en estos autos: ', False),
                    (car_ex, True),
                    (', ante V.S. respetuosamente digo QUE:', False),
                ]
            reemplazar_runs(para, runs_data)
            break

    # ── Para apoderado: eliminar manifestación y líneas sueltas al final ──
    if esApod:
        encontro_fin = False
        for para in list(body.findall(tag('p'))):
            texts = ''.join(t.text or '' for t in para.findall(f'.//{tag("t")}'))
            if 'MANIFESTACI' in texts:
                body.remove(para)
                continue
            if 'SERA JUSTICIA' in texts or 'SERÁ JUSTICIA' in texts:
                encontro_fin = True
                continue
            if encontro_fin:
                elem_str = etree.tostring(para).decode()
                if 'v:rect' in elem_str and not texts.strip():
                    body.remove(para)

    nuevo_xml = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    archivos['word/document.xml'] = nuevo_xml
    
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in archivos.items():
            zout.writestr(name, data)
    buf.seek(0)
    return buf.read()

def build_reemplazos(tipo, d):
    """Construye el mapa de reemplazos de texto según el tipo de escrito"""
    p   = d['perito']
    car = d.get('caratula', '')
    ex  = d.get('expte', '')
    rep = p.get('rep', 'patrocinio')
    esApod = rep == 'apoderado'
    car_ex = f'\u201c{car}\u201d Expte. N\u00b0 {ex}'

    r = {
        # Carátulas del template
        '\u201cCABRERA GABRIEL DAVID C/ RIO URUGUAY COOPERATIVA DE SEGUROS LTDA. S/ EJECUCION DE HONORARIOS\u201d Expte.: 12573': f'\u201c{car}\u201d Expte. N\u00b0 {ex}',
        '\u201cMODAD, REN\u00c9E ALINE C/PROVINCIA ART SA S/ EJECUCION DE HONORARIOS\u201d (Expte. N\u00ba 8805)': f'\u201c{car}\u201d (Expte. N\u00b0 {ex})',
        '\u201cMODAD RENEE ALINE C/ PREVENCION ART SA S/ EJECUCION DE HONORARIOS (CONEX.LABORAL N\u00ba 4)\u201d - Expte. N\u00ba 18482': f'\u201c{car}\u201d - Expte. N\u00b0 {ex}',
        'PEREZ MIRIAM CRISTINA C/ LIBRA COMPA\u00d1\u00cdA DE SEGUROS S.A. S/ EJECUCION DE HONORARIOS (CONEXIDAD CYC 3) N\u00ba37879': car,
        '\u201cMUSSINO CORRADO JUAN MARIA C/ MUNICIPALIDAD DE GUALEGUAYCHU S/EJECUCION DE HONORARIOS\u201d Expte. N\u00ba 492/25': f'\u201c{car}\u201d Expte. N\u00b0 {ex}',
        '\u201cMUSSINO CORRADO JUAN MARIA C/ COMPA\u00d1IA DE SEGUROS LA MERCANTIL ANDINA S.A. S/EJECUCION DE HONORARIOS\u201d Expte. N\u00ba 484/25.': f'\u201c{car}\u201d Expte. N\u00b0 {ex}',
        '\u201cFRIEDRICH, MAURICIO FRANCISCO C/ INSTITUTO AUT\u00c1RQUICO PROVINCIAL DEL SEGURO S/ EJECUCION DE HONORARIOS\u201d Exp.N\u00b09113': f'\u201c{car}\u201d Expte. N\u00b0 {ex}',
        '\u201cFRIEDRICH MAURICIO FRANCISCO C/ SAN CRISTOBAL SOCIEDAD MUTUAL DE SEGUROS GENERALES S/ EJECUCION DE HONORARIOS\u201d Expte. N\u00ba 7194': f'\u201c{car}\u201d Expte. N\u00b0 {ex}',
        # Expedientes
        'Expte.: 12573': f'Expte. N\u00b0 {ex}',
        'Expte. N\u00ba 8805': f'Expte. N\u00b0 {ex}',
        'Expte. N\u00ba 18482': f'Expte. N\u00b0 {ex}',
        'Exp.N\u00b09113': f'Expte. N\u00b0 {ex}',
        'Expte. N\u00ba 492/25': f'Expte. N\u00b0 {ex}',
        'Expte. N\u00ba 484/25.': f'Expte. N\u00b0 {ex}',
        'Expte. N\u00ba 7194': f'Expte. N\u00b0 {ex}',
        # Nombres perito
        'GABRIEL DAVID CABRERA': p['nombre'],
        'RENEE ALINE MODAD': p['nombre'],
        'RENÉE ALINE MODAD': p['nombre'],
        'MODAD RENÉE ALINE': p['nombre'],
        'FRIEDRICH MAURICIO FRANCISCO': p['nombre'],
        'MUSSINO CORRADO JUAN MARIA': p['nombre'],
        'PEREZ MIRIAM CRISTINA': p['nombre'],
        'el Sr. GABRIEL DAVID CABRERA': f'el/la {p["nombre"]}',
        # DNI
        'DNI N\u00ba 23.696.067': f'DNI N\u00b0 {p["dni"]}',
        'D.N.I. N\u00ba23.696.067': f'D.N.I. N\u00b0 {p["dni"]}',
        'DNI N\u00ba 23.450.195': f'DNI N\u00b0 {p["dni"]}',
        # CUIT perito
        'CUIT 27-22305073-0': f'CUIT {p["cuit"]}',
        'CUIT N\u00ba 20-23696067-7': f'CUIT N\u00b0 {p["cuit"]}',
        'CUIT N\u00ba 20-23450195-0': f'CUIT N\u00b0 {p["cuit"]}',
        'CUIT 20-118585942-2': f'CUIT {p["cuit"]}',
        # Profesión
        'Ingeniero Mecánico': p['prof'],
        # CBU / alias perito
        'CAUSA.SABLE.GRIFO': p['alias'],
        'CABEZA.TEMPLO.CHACO': p['alias'],
        'BURRO.SOJA.AULA': p['alias'],
        'CBU N\u00ba 0110223120022300419074': f'CBU N\u00b0 {p["cbu"]}',
        'CBU: 0720192588000037160986': f'CBU: {p["cbu"]}',
        'Número de CBU: 0720192588000037160986': f'Número de CBU: {p["cbu"]}',
        # Condición fiscal
        'Monotributista, inscripta como profesional liberal en ATER': p['cond'],
        # Demandado
        'PROVINCIA A.R.T. S.A.': d.get('demandado', ''),
        'PROVINCIA ART SA': d.get('demandado', ''),
        'LIBRA COMPAÑÍA ARGENTINA DE SEGUROS S.A.': d.get('demandado', ''),
        'RIO URUGUAY COOPERATIVA DE SEGUROS LTDA.': d.get('demandado', ''),
        'MUNICIPALIDAD DE GUALEGUAYCHU': d.get('demandado', ''),
        # CUIT demandado
        'CUIT 30-68825409-0': f'CUIT {d.get("cuitDem", "")}',
        'CUIT N\u00ba30-68825409-0': f'CUIT N\u00b0 {d.get("cuitDem", "")}',
        'CUIT 30-71233282-0': f'CUIT {d.get("cuitDem", "")}',
        'CUIT 30-50003691-1': f'CUIT {d.get("cuitDem", "")}',
    }

    # Plural para apoderado
    if esApod:
        r['Vengo por medio'] = 'Venimos por medio'
        r['intereso se sirva'] = 'interesamos se sirva'
        r['respetuosamente digo QUE'] = 'respetuosamente decimos que'

    # Reemplazos específicos por tipo
    if tipo == 'informa_ope':
        r['26/03/2026'] = d.get('fecha_prov', '')

    elif tipo in ('liq_apod', 'liq_patr'):
        r.update({
            '30/09/2022': d.get('fecha_reg', ''),
            '11/02/2026': d.get('fecha_reg', ''),
            '(30/09/2022 - 11/03/2026)': f'({d.get("int_desde","")} - {d.get("int_hasta","")})',
            '(12/08/24 - 30/03/26)': f'({d.get("int_desde","")} - {d.get("int_hasta","")})',
            '$ 126.487,73': f'$ {d.get("hon_ejec","")}',
            '$126.487,73': f'${d.get("hon_ejec","")}',
            '$ 318.951.63': f'$ {d.get("intereses","")}',
            '$ 445.439,36': f'$ {d.get("subtotal_hon","")}',
            '10 juristas': f'{d.get("jur_m","")} juristas',
            '5 juristas': f'{d.get("jur_m","")} juristas',
            '$ 793.916,1': f'$ {d.get("mm","")}',
            '$ 79.864,11': f'$ {d.get("gastos","")}',
            '$ 1.677.146,31': f'$ {d.get("subtotal_ejec","")}',
            '$2.122.585,67': f'${d.get("total","")}',
            '$321.362': f'${d.get("saldo","")}',
        })

    elif tipo == 'aprob_ope':
        r.update({
            '$570.022,02': f'${d.get("mp","")}',
            '$24.268,54 en concepto de gastos, con destino a la cuenta que se individualiza como Caja de Ahorro N\u00ba 599018198202':
                f'${d.get("mm","")} en concepto de honorarios regulados, con destino a la cuenta que se individualiza como Caja de Ahorro N\u00b0 599018198202',
            '$24.268,54 en concepto de gastos, siendo la cuenta bancaria':
                f'${d.get("mg","")} en concepto de honorarios regulados, siendo la cuenta bancaria',
        })

    elif tipo == 'amplia':
        r.update({
            '18/02/2026': d.get('fecha_ope', ''),
            '19/02/2026': d.get('fecha_ope', ''),
            '$1.513.473,11': f'${d.get("monto_amp","")}',
            '$1.741.181,2': f'${d.get("saldo_pend","")}',
        })

    elif tipo == 'oficio':
        r.update({
            'Dr. Arturo Mc. Loughlin': d.get('juez', ''),
            'Dr. Luciano Amoroto': d.get('secretaria', ''),
            '$11.539.549,73': f'${d.get("monto","")}',
            'N\u00b016-100492/8': f'N\u00b0 {d.get("cta_judicial","")}',
            'CBU de la Cuenta:3860016403000010049281': f'CBU: {d.get("cbu_judicial","")}',
            '<jdocyclab-ssdor@jusentrerios.gov.ar>': f'<{d.get("email_juzgado","")}>',
        })

    elif tipo == 'oficio_reit':
        r.update({
            'Dr. Francisco Unamunzaga': d.get('juez', ''),
            'Dr. Luciano G. Bernigaud': d.get('secretaria', ''),
            'N\u00b07-102656/2': f'N\u00b0 {d.get("cta_judicial","")}',
            'CBU de la Cuenta:3860007203000010265629': f'CBU: {d.get("cbu_judicial","")}',
            '$6.450,00': f'${d.get("monto","")}',
        })

    return r

def docx_to_pdf(docx_bytes):
    """Convierte docx a PDF usando LibreOffice"""
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
        f.write(docx_bytes)
        tmp_docx = f.name
    
    tmp_dir = tempfile.mkdtemp()
    os.system(f'libreoffice --headless --convert-to pdf --outdir {tmp_dir} {tmp_docx}')
    
    pdf_name = os.path.basename(tmp_docx).replace('.docx', '.pdf')
    pdf_path = os.path.join(tmp_dir, pdf_name)
    
    if os.path.exists(pdf_path):
        with open(pdf_path, 'rb') as f:
            pdf_bytes = f.read()
        os.remove(tmp_docx)
        os.remove(pdf_path)
        return pdf_bytes
    
    os.remove(tmp_docx)
    raise Exception('Error al convertir a PDF')

# ── Endpoint principal ──────────────────────────────────────────
@app.route('/generar', methods=['POST', 'OPTIONS'])
def generar():
    if request.method == 'OPTIONS':
        return '', 200
    
    try:
        datos = request.json
        tipo      = datos.get('tipo', '')
        folder_id = datos.get('folder_id', '')
        
        if tipo not in TEMPLATES:
            return jsonify({'error': f'Tipo no soportado: {tipo}'}), 400
        
        drive = get_drive_service()
        
        # Descargar template
        template_id = TEMPLATES[tipo]
        template_bytes = descargar_template(drive, template_id)
        
        # Generar docx
        docx_bytes = procesar_docx(template_bytes, tipo, datos)
        
        # Convertir a PDF
        pdf_bytes = docx_to_pdf(docx_bytes)
        
        # Subir PDF al Drive
        from datetime import datetime
        fecha = datetime.now().strftime('%d-%m-%Y')
        car = datos.get('caratula', '')[:40]
        nombre = f"{tipo.upper()} - {car} - {fecha}.pdf"
        
        resultado = subir_pdf(drive, folder_id, nombre, pdf_bytes)
        
        return jsonify({
            'ok': True,
            'pdf_id': resultado['id'],
            'url': resultado['webViewLink'],
            'nombre': nombre,
        })
    
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'ok': True, 'version': '1.0'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
