"""Workbook batch quotes using the same rate finder as the single-shipment UI."""
from io import BytesIO
from decimal import Decimal, ROUND_HALF_UP
import math
import re

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

FIELDS = {'postcode':'Postcode losadres', 'kg':'Totaalgewicht',
          'ldm':'Totaal laadmeter', 'pallets':'Aantal Pallets'}


def numeric(value):
    if isinstance(value, bool) or value is None:
        raise ValueError('Ontbrekende of ongeldige numerieke invoer')
    text = str(value).strip().replace(' ', '')
    if ',' in text:
        text = text.replace('.', '').replace(',', '.')
    try:
        result = float(text)
    except (ValueError, TypeError):
        raise ValueError('Ontbrekende of ongeldige numerieke invoer') from None
    if not math.isfinite(result) or result <= 0:
        raise ValueError('Gewicht, laadmeters en pallets moeten positief zijn')
    return result


def money(value):
    return float(Decimal(str(value)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))


def normalize_postcode(value, country):
    if isinstance(value,(int,float)) and not isinstance(value,bool):
        if not math.isfinite(value) or int(value)!=value:
            raise ValueError('Ongeldige postcode')
        value = str(int(value))
    postcode = str(value or '').strip().upper().replace(' ','')
    if country == 'FR':
        if not re.fullmatch(r'\d{1,5}',postcode):
            raise ValueError('Franse postcode moet vijf cijfers bevatten')
        postcode = postcode.zfill(5)
    if not postcode:
        raise ValueError('Postcode ontbreekt')
    return postcode


def enrich_workbook(raw, sheet, columns, country, pallet_type, rates, diesel, road,
                    margin, include_margin, find_rates):
    book = openpyxl.load_workbook(BytesIO(raw))
    values = openpyxl.load_workbook(BytesIO(raw), data_only=True)
    ws, source = book[sheet], values[sheet]
    if source.max_row > 5001:
        raise ValueError('Maximaal 5.000 zendingregels per bestand.')
    country = country.upper()
    rates = [r for r in rates if r.country.upper() == country]
    carriers = sorted({r.carrier for r in rates})
    if not carriers:
        raise ValueError('Geen vervoerders met een staffel voor het geselecteerde land.')
    # Unknown pallet types cannot safely use pallet-count-specific brackets.
    eligible = rates if pallet_type != 'Onbekend' else [
        r for r in rates if r.max_eu_pallets is None and r.max_block_pallets is None]
    fields = ['Totale inkoop EUR']
    if include_margin:
        fields += ['Verkoop EUR']
    start = ws.max_column + 1
    headers = [
        f'{carrier} - {field}' for carrier in carriers for field in fields]
    for col, header in enumerate(headers, start):
        cell = ws.cell(1,col,header)
        cell.font = Font(bold=True,color='FFFFFF')
        cell.fill = PatternFill('solid',fgColor='205864')
        cell.alignment = Alignment(wrap_text=True,vertical='top')
        ws.column_dimensions[get_column_letter(col)].width = 23
    ws.row_dimensions[1].height = max(ws.row_dimensions[1].height or 15,60)
    preview, counts = [], {'regels':0,'met_tarief':0,'zonder_tarief':0,'ongeldig':0}
    for row in range(2,source.max_row+1):
        if all(source.cell(row,c).value is None for c in range(1,start)):
            continue
        counts['regels'] += 1
        output, matches, error = [], {}, None
        try:
            postcode = normalize_postcode(source.cell(row,columns['postcode']).value,country)
            kg, ldm, pallets = (numeric(source.cell(row,columns[k]).value) for k in ('kg','ldm','pallets'))
            if not pallets.is_integer():
                raise ValueError('Aantal pallets moet een geheel getal zijn')
            eu = pallets if pallet_type=='Europallet' else 0
            block = pallets if pallet_type=='Blokpallet' else 0
            _, found = find_rates(eligible,country,postcode,kg,ldm,eu,block)
            matches = {r.carrier:r for r in found}
        except ValueError as exc:
            error = str(exc)
        counts['ongeldig' if error else 'met_tarief' if matches else 'zonder_tarief'] += 1
        for carrier in carriers:
            rate = matches.get(carrier)
            if error or rate is None:
                output += [('Ongeldige invoer: '+error) if error else 'Geen tarief'] * len(fields)
                continue
            base = Decimal(str(rate.base_price))
            fuel_pct = Decimal(str(diesel.get(carrier,0)))
            road_pct = Decimal(str(road.get(carrier,0)))
            fuel_amount, road_amount = base*fuel_pct/100, base*road_pct/100
            purchase = base+fuel_amount+road_amount
            output += [money(purchase)]
            if include_margin:
                profit = purchase*Decimal(str(margin))/100
                output += [money(purchase+profit)]
        for col, value in enumerate(output,start):
            cell = ws.cell(row,col,value)
            if headers[col-start].endswith('%'):
                cell.number_format = '0.00%'
            elif headers[col-start].endswith('EUR'):
                cell.number_format = '#,##0.00'
            if isinstance(value,str):
                cell.data_type = 's'
        if len(preview)<25:
            preview.append({'Excel-rij':row,**dict(zip(headers,output))})
    buffer = BytesIO()
    book.save(buffer)
    book.close()
    values.close()
    return buffer.getvalue(),preview,counts


def add_pallet_scale(raw, sheet, postcode_column, country, kg_each, rates, diesel,
                     road, margin, include_margin, find_rates, source_raw=None):
    kg_each = numeric(kg_each)
    book = openpyxl.load_workbook(BytesIO(raw))
    values = openpyxl.load_workbook(BytesIO(source_raw if source_raw is not None else raw),data_only=True)
    destinations, skipped = set(), 0
    for row in values[sheet].iter_rows(min_row=2):
        if all(cell.value is None for cell in row):
            continue
        try:
            destinations.add(normalize_postcode(row[postcode_column-1].value,country))
        except ValueError:
            skipped += 1
    if len(destinations)>1250:
        raise ValueError('Maximaal 1.250 unieke bestemmingen voor een 1-4 palletstaffel.')
    if not destinations:
        raise ValueError('Geen geldige bestemmingen voor een palletstaffel.')
    scale = book.create_sheet('Staffel 1-4 EP')
    scale.append(['Land','Postcode','Europallets','Totaalgewicht kg','Laadmeters','Kg per europallet'])
    for postcode in sorted(destinations):
        for qty in range(1,5):
            scale.append([country,postcode,qty,qty*kg_each,round(qty*0.4,2),kg_each])
    scale.freeze_panes = 'A2'
    scale.auto_filter.ref = scale.dimensions
    for column in range(1,7):
        scale.column_dimensions[get_column_letter(column)].width=22
        scale.cell(1,column).font=Font(bold=True)
    name=scale.title
    buffer=BytesIO()
    book.save(buffer)
    book.close()
    values.close()
    output,_,counts = enrich_workbook(buffer.getvalue(),name,
        {'postcode':2,'kg':4,'ldm':5,'pallets':3},country,'Europallet',rates,
        diesel,road,margin,include_margin,find_rates)
    return output,{'bestemmingen':len(destinations),'overgeslagen_postcodes':skipped,**counts}


def render_batch_ui(st, rates, diesel, road, margin, find_rates, countries):
    import hashlib
    import json
    st.subheader('Excel-zendingen')
    upload = st.file_uploader('Zendingenbestand',type=['xlsx'],key='shipment_batch')
    country = st.selectbox('Leverland voor alle regels',countries,index=countries.index('FR'),key='batch_country')
    pallet_type = st.selectbox('Palletsoort voor alle regels',['Onbekend','Europallet','Blokpallet'],key='batch_pallet_type')
    include_margin = st.checkbox('Ingestelde marge meenemen',value=False,key='batch_margin')
    if include_margin:
        st.metric('Marge over totale inkoop',f'{margin:g}%')
    include_scale = st.checkbox('Staffel voor 1-4 europallets toevoegen',key='batch_scale')
    kg_each = 0.0
    if include_scale:
        kg_each = st.number_input('Gewicht per europallet (kg)',min_value=0.0,value=0.0,step=25.0,key='batch_scale_kg')
        st.caption('Apart werkblad per unieke postcode: 1-4 europallets, 0,4 laadmeter per pallet. Het gekozen gewicht geldt voor alle bestemmingen.')
        if kg_each<=0:
            st.warning('Vul een positief gewicht per europallet in voor de staffel.')
    st.warning('Berekening met de nu geladen staffels en ingestelde toeslagen, niet met historische tarieven op de losdatum. Controleer de geldigheid van de staffels.')
    if pallet_type=='Onbekend':
        st.caption('Palletspecifieke staffels worden uitgesloten zolang de palletsoort onbekend is.')
    if upload is None:
        return
    if upload.size > 30*1024*1024:
        st.error('Bestand te groot: maximaal 30 MB.')
        return
    raw = upload.getvalue()
    try:
        book = openpyxl.load_workbook(BytesIO(raw),read_only=True,data_only=True)
        sheet = st.selectbox('Werkblad',book.sheetnames,key='batch_sheet')
        headers = next(book[sheet].values,())
        book.close()
        options = {i:f'{get_column_letter(i)} - {value}' for i,value in enumerate(headers,1) if value is not None}
        if not options:
            raise ValueError('Geen kolomkoppen gevonden op de eerste rij.')
        columns = {}
        widgets = st.columns(4)
        for widget,(field,label) in zip(widgets,FIELDS.items()):
            selected = next((i for i in options if str(headers[i-1]).strip().lower()==label.lower()),next(iter(options)))
            columns[field] = widget.selectbox(label,list(options),index=list(options).index(selected),format_func=options.get,key='batch_col_'+field)
        if len(set(columns.values()))!=4:
            st.error('Selecteer voor elk invoerveld een aparte kolom.')
            return
        fingerprint = hashlib.sha256(b'compact-v2'+raw+repr(rates).encode()+json.dumps([sheet,columns,country,pallet_type,diesel,road,margin,include_margin,include_scale,kg_each],sort_keys=True).encode()).hexdigest()
        if st.button('Tarieven per vervoerder berekenen',disabled=not rates or (include_scale and kg_each<=0),key='batch_calculate'):
            with st.spinner('Zendingen berekenen...'):
                result = enrich_workbook(raw,sheet,columns,country,pallet_type,rates,diesel,road,margin,include_margin,find_rates)
                if include_scale:
                    output,scale_counts = add_pallet_scale(result[0],sheet,columns['postcode'],country,kg_each,rates,diesel,road,margin,include_margin,find_rates,source_raw=raw)
                    result = (output,result[1],{**result[2],'staffel':scale_counts})
                st.session_state['batch_result'] = (fingerprint,result)
        saved = st.session_state.get('batch_result')
        if saved and saved[0]==fingerprint:
            output, preview, counts = saved[1]
            st.write(f"{counts['regels']} zendingen: {counts['met_tarief']} met tarief, {counts['zonder_tarief']} zonder tarief, {counts['ongeldig']} met ongeldige invoer.")
            if 'staffel' in counts:
                scale = counts['staffel']
                st.write(f"Palletstaffel: {scale['bestemmingen']} bestemmingen, {scale['regels']} scenario's. {scale['zonder_tarief']} zonder tarief; {scale['overgeslagen_postcodes']} regels met ongeldige postcode overgeslagen.")
            st.download_button('Excel met tarieven downloaden',output,file_name=upload.name.rsplit('.',1)[0]+'_tarieven.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            st.dataframe(preview,hide_index=True,width='stretch')
    except Exception as exc:
        st.error(f'Zendingenbestand kon niet worden verwerkt: {exc}')
