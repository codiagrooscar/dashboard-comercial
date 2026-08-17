# Decompiled with PyLingual (https://pylingual.io)
# Internal filename: 'C:\\Users\\oscar.ocampo\\Documents\\Codex\\2026-05-20\\files-mentioned-by-the-user-facturas\\app.py'
# Bytecode version: 3.12.0rc2 (3531)
# Source timestamp: 2026-06-09 08:20:23 UTC (1780993223)

from __future__ import annotations
import html
import io
import json
import os
import re
import socketserver
import ssl
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from http.server import BaseHTTPRequestHandler
from pathlib import Path
from excel_report import generate_excel_dashboard
from typing import Any
import pandas as pd

RATES_PLANTA = {
    'FABRICAR': {'Líquidos': 1875.0, 'Sólidos': 500.0, 'Flows': 687.5, 'Otros': 600.0},
    'ENVASAR': {
        'Líquidos': {'1000': 2500.0, '200': 1250.0, '20': 800.0, '5': 350.0, '1': 125.0, 'default': 800.0},
        'Sólidos': {'BIG BAG': 300.0, '500': 30.0, '20': 200.0, '5': 150.0, '1': 60.0, 'default': 150.0},
        'Flows': {'20': 850.0, '5': 375.0, '1': 100.0, 'default': 500.0},
        'Otros': {'default': 500.0}
    }
}

PORT = int(os.environ.get('PORT', '8765'))
BASE_DIR = Path(__file__).resolve().parent
logo_txt_path = BASE_DIR / 'codiagro_logo_b64.txt'
if logo_txt_path.exists():
    try:
        CODIAGRO_LOGO_B64 = logo_txt_path.read_text(encoding='utf-8').strip()
    except Exception:
        CODIAGRO_LOGO_B64 = 'https://www.codiagro.com/wp-content/uploads/2020/08/logo.png'
else:
    CODIAGRO_LOGO_B64 = 'https://www.codiagro.com/wp-content/uploads/2020/08/logo.png'
SUPABASE_URL = os.environ.get('SUPABASE_URL', 'https://txmkqedmivigcmqhfqjr.supabase.co')
SUPABASE_KEY = os.environ.get('SUPABASE_KEY', 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InR4bWtxZWRtaXZpZ2NtcWhmcWpyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzU3NzkwMTIsImV4cCI6MjA5MTM1NTAxMn0.NlX_flrxzyeL5C15OzugYAVOS2QrbKBu16TSMZUOopM')
SUPABASE_ENABLED = os.environ.get('SUPABASE_ENABLED', '1').strip().lower() in {'true', 'on', 'yes', '1'}
LOCAL_DATA_DIR = BASE_DIR / 'data'
def make_ssl_context() -> ssl.SSLContext:
    context = ssl.create_default_context()
    ignore_eof = getattr(ssl, 'OP_IGNORE_UNEXPECTED_EOF', 0)
    if ignore_eof:
        context.options |= ignore_eof
    return context
SSL_CONTEXT = make_ssl_context()
def supabase_request(path: str, method: str='GET', data: list | dict | None=None, query_params: dict | None=None) -> Any:
    if not SUPABASE_ENABLED:
        raise RuntimeError('Supabase está desactivado en modo local.')
    url = f'{SUPABASE_URL}/rest/v1/{path}'
    if query_params:
        url += '?' + urllib.parse.urlencode(query_params)
    req = urllib.request.Request(url, method=method)
    req.add_header('apikey', SUPABASE_KEY)
    req.add_header('Authorization', f'Bearer {SUPABASE_KEY}')
    req.add_header('Connection', 'close')
    if data is not None:
        body = json.dumps(data).encode('utf-8')
        req.add_header('Content-Type', 'application/json')
        req.data = body
        
    last_error = None
    for attempt in range(5):
        try:
            with urllib.request.urlopen(req, timeout=45, context=SSL_CONTEXT) as response:
                if method == 'GET':
                    return json.loads(response.read().decode('utf-8'))
                return None
        except urllib.error.HTTPError as e:
            error_msg = e.read().decode('utf-8')
            raise Exception(f'Supabase error ({e.code}): {error_msg}')
        except Exception as exc:
            last_error = exc
            if attempt == 4:
                raise
            else:
                time.sleep(0.6 * (attempt + 1))
    if last_error:
        raise last_error

def clean_df_for_json(df: pd.DataFrame, report_date: str) -> list[dict[str, Any]]:
    records = []
    for _, row in df.iterrows():
        d = row.to_dict()
        d['report_date'] = report_date
        for k, v in d.items():
            if pd.isna(v):
                d[k] = None
            else:
                if isinstance(v, (pd.Timestamp, datetime)):
                    d[k] = v.strftime('%Y-%m-%d')
                elif hasattr(v, 'isoformat'):
                    d[k] = v.isoformat()
                else:
                    if isinstance(v, (set, list)):
                        d[k] = list(v)
        records.append(d)
    return records
def save_data_to_supabase(report_date: str, dfs: dict[str, pd.DataFrame]) -> None:
    for name, df in dfs.items():
        try:
            supabase_request(name, method='DELETE', query_params={'report_date': f'eq.{report_date}'})
        except Exception as e:
            print(f'Error al borrar {name} en Supabase: {e}')
            if name == 'stock': continue # Skip saving if table doesn't exist
            
        df_to_save = df.drop(columns=['id'], errors='ignore')
        
        if name == 'produccion':
            allowed_cols = [
                'idot', 'fecha', 'codigoarticulo', 'descripcionarticulo',
                'unidadesafabricar', 'unidadesfabricadas', 'unidadesrechazadas',
                'tiempototal', 'tiemporealtotal', 'tiempofabricacion', 'tiempopreparacion', 'tiempoparadas',
                'costetotal', 'costerealtotal', 'costerealmaquina', 'costerealmanoobra',
                'descripcionmaquina',
                'codigo', 'descripcion', 'udsafabricar', 'udsfabricadas', 'tiempotot', 'costereal'
            ]
            cols_to_keep = [c for c in allowed_cols if c in df_to_save.columns]
            df_to_save = df_to_save[cols_to_keep]
            
        records = clean_df_for_json(df_to_save, report_date)
        if records:
            batch_size = 500
            for i in range(0, len(records), batch_size):
                batch = records[i:i + batch_size]
                try:
                    supabase_request(name, method='POST', data=batch)
                except Exception as e:
                    print(f'Error al guardar {name} en Supabase: {e}')
def local_data_path(report_date: str, name: str) -> Path:
    safe_date = re.sub('[^0-9-]', '', report_date)
    return LOCAL_DATA_DIR / safe_date / f'{name}.json'
def save_data_to_local(report_date: str, dfs: dict[str, pd.DataFrame]) -> None:
    (LOCAL_DATA_DIR / report_date).mkdir(parents=True, exist_ok=True)
    for name, df in dfs.items():
        records = clean_df_for_json(df, report_date)
        local_data_path(report_date, name).write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding='utf-8')
def load_data_from_local(report_date: str) -> dict[str, pd.DataFrame] | None:
    dfs = {}
    found_any = False
    for name in ['ofertas', 'pedidos', 'albaranes', 'facturas', 'produccion', 'stock']:
        path = local_data_path(report_date, name)
        if path.exists():
            found_any = True
            records = json.loads(path.read_text(encoding='utf-8'))
            dfs[name] = pd.DataFrame(records)
        else:
            dfs[name] = pd.DataFrame()
        dfs[name] = ensure_types_normalized_df(dfs[name], name)
    if not found_any or all((df.empty for df in dfs.values())):
        return None
    else:
        return dfs
def save_report_data(report_date: str, dfs: dict[str, pd.DataFrame]) -> None:
    save_data_to_local(report_date, dfs)
    if SUPABASE_ENABLED:
        save_data_to_supabase(report_date, dfs)
def load_report_data(report_date: str, fallback: bool = True) -> dict[str, pd.DataFrame] | None:
    dfs = load_data_from_local(report_date)
    if dfs is None:
        dfs = {}
        
    if SUPABASE_ENABLED:
        try:
            supa_dfs = load_data_from_supabase(report_date)
            if supa_dfs is not None:
                for k, v in supa_dfs.items():
                    if not v.empty or k not in dfs:
                        dfs[k] = v
                save_data_to_local(report_date, dfs)
        except Exception as exc:
            print(f'No se pudo cargar desde Supabase: {exc}')
            
    if not dfs:
        dfs = None
        
    if dfs is None and fallback:
        # Fallback a la foto más reciente disponible
        if SUPABASE_ENABLED:
            try:
                records = supabase_request('produccion', query_params={'select': 'report_date'})
                if not records:
                    records = supabase_request('facturas', query_params={'select': 'report_date'})
                if records:
                    valid_dates = sorted(list(set((r['report_date'] for r in records if r.get('report_date')))), reverse=True)
                    if valid_dates:
                        best_date = valid_dates[0]
                        for d in valid_dates:
                            if d <= report_date:
                                best_date = d
                                break
                        if best_date != report_date:
                            print(f'Falling back to snapshot {best_date} for requested {report_date}')
                            return load_report_data(best_date, fallback=False)
            except Exception as e:
                pass
                
        # Fallback local
        if LOCAL_DATA_DIR.exists():
            local_dirs = [d.name for d in LOCAL_DATA_DIR.iterdir() if d.is_dir() and re.match(r'^\d{4}-\d{2}-\d{2}$', d.name) and d.name != '9999-12-31']
            if local_dirs:
                valid_dates = sorted(local_dirs, reverse=True)
                if valid_dates:
                    best_date = valid_dates[0]
                    for d in valid_dates:
                        if d <= report_date:
                            best_date = d
                            break
                    if best_date != report_date:
                        print(f'Falling back locally to snapshot {best_date} for requested {report_date}')
                        return load_report_data(best_date, fallback=False)
                        
    return dfs

def ensure_types_normalized_df(df: pd.DataFrame, kind: str) -> pd.DataFrame:
    if df.empty:
        cols = ['documento', 'fecha', 'importe', 'cliente', 'razon_social', 'articulo', 'serie', 'zona']
        if kind in ['pedidos', 'ofertas']:
            cols += ['descripcion']
        if kind == 'pedidos':
            cols += ['fecha_necesaria', 'importe_pendiente', 'unidades_pedidas', 'unidades_servidas', 'unidades_pendientes']
        if kind == 'stock':
            cols = ['codigo', 'descripcion', 'familia', 'cantidad']
        return pd.DataFrame(columns=cols)
    else:
        if kind not in ('produccion', 'stock'):
            if 'documento' not in df.columns:
                df['documento'] = ''
            df['documento'] = df['documento'].fillna('').astype(str)
            if 'fecha' not in df.columns:
                df['fecha'] = pd.NaT
            df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
            if 'importe' not in df.columns:
                df['importe'] = 0.0
            df['importe'] = pd.to_numeric(df['importe'], errors='coerce').fillna(0.0)
            if 'cliente' not in df.columns:
                df['cliente'] = ''
            df['cliente'] = df['cliente'].fillna('').astype(str).str.replace('\\.0$', '', regex=True).str.strip()
            if 'razon_social' not in df.columns:
                df['razon_social'] = ''
            df['razon_social'] = df['razon_social'].fillna('').astype(str).str.strip()
            if 'articulo' not in df.columns:
                df['articulo'] = ''
            df['articulo'] = df['articulo'].fillna('').astype(str).str.strip()
            if 'serie' not in df.columns:
                df['serie'] = ''
            df['serie'] = df['serie'].fillna('').astype(str).str.strip()
            if 'zona' not in df.columns:
                df['zona'] = ''
            df['zona'] = df['zona'].fillna('').astype(str).str.strip()
            if kind in ['pedidos', 'ofertas']:
                if 'descripcion' not in df.columns:
                    df['descripcion'] = ''
                df['descripcion'] = df['descripcion'].fillna('').astype(str).str.strip()
            if kind == 'pedidos':
                df['fecha_necesaria'] = pd.to_datetime(df.get('fecha_necesaria'), errors='coerce')
                df['importe_pendiente'] = pd.to_numeric(df.get('importe_pendiente'), errors='coerce').fillna(0.0)
                df['unidades_pedidas'] = pd.to_numeric(df.get('unidades_pedidas'), errors='coerce').fillna(0.0)
                df['unidades_servidas'] = pd.to_numeric(df.get('unidades_servidas'), errors='coerce').fillna(0.0)
                df['unidades_pendientes'] = pd.to_numeric(df.get('unidades_pendientes'), errors='coerce').fillna(0.0)
        elif kind == 'stock':
            first_row = pd.DataFrame([df.columns.values], columns=df.columns)
            df = pd.concat([first_row, df], ignore_index=True)
            
            header_row_idx = None
            for i, row in df.head(10).iterrows():
                row_strs = [str(x).upper() for x in row.values]
                matches = sum(1 for req in ['CODIGO', 'FAMILIA', 'CANTIDAD'] if any(req in r for r in row_strs))
                if matches >= 2:
                    header_row_idx = i
                    break
            
            if header_row_idx is not None:
                new_cols = df.iloc[header_row_idx].values
                df.columns = new_cols
                df = df.iloc[header_row_idx + 1:].reset_index(drop=True)
                
                col_map = {}
                for col in df.columns:
                    norm_col = str(col).lower().strip().replace('ó', 'o').replace('í', 'i')
                    col_map[col] = norm_col
                df = df.rename(columns=col_map)
                
                if 'codigo' in df.columns and 'descripcion' in df.columns and 'familia' in df.columns:
                    first_codigo = str(df['codigo'].dropna().iloc[0]).strip() if not df['codigo'].dropna().empty else ''
                    first_familia = str(df['familia'].dropna().iloc[0]).strip() if not df['familia'].dropna().empty else ''
                    if first_codigo.isdigit() and not first_familia.isdigit():
                        df['familia_real'] = df['codigo']
                        df['codigo_real'] = df['descripcion']
                        df['descripcion_real'] = df['familia']
                        
                        df['familia'] = df['familia_real']
                        df['codigo'] = df['codigo_real']
                        df['descripcion'] = df['descripcion_real']
                        
                        df = df.drop(columns=['familia_real', 'codigo_real', 'descripcion_real'])
            else:
                if len(df.columns) >= 5:
                    df.columns = ['fecha', 'familia', 'codigo', 'descripcion', 'cantidad'] + list(df.columns[5:])
                elif len(df.columns) == 4:
                    df.columns = ['codigo', 'descripcion', 'familia', 'cantidad']
            
            for req in ['codigo', 'descripcion', 'familia', 'cantidad']:
                if req not in df.columns:
                    df[req] = '' if req != 'cantidad' else 0.0
            
            if 'fecha' not in df.columns:
                df['fecha'] = pd.NaT
            else:
                df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
            
            df['codigo'] = df['codigo'].fillna('').astype(str).str.strip()
            df['descripcion'] = df['descripcion'].fillna('').astype(str).str.strip()
            df['familia'] = pd.to_numeric(df['familia'], errors='coerce').fillna(0).astype(int)
            df['cantidad'] = pd.to_numeric(df['cantidad'], errors='coerce').fillna(0.0)
            
        elif kind == 'produccion':
            # Detectar fila de cabecera si hay filas en blanco al principio (ej: exportaciones de ERP)
            first_row = pd.DataFrame([df.columns.values], columns=df.columns)
            df = pd.concat([first_row, df], ignore_index=True)
            
            header_row_idx = None
            for i, row in df.head(10).iterrows():
                row_strs = [str(x).upper() for x in row.values]
                matches = sum(1 for req in ['CODIGO', 'FECHA', 'FABRICADAS', 'FABRICAR', 'DESCRIPCION'] if any(req in r for r in row_strs))
                if matches >= 2:
                    header_row_idx = i
                    break
            
            if header_row_idx is not None:
                new_cols = df.iloc[header_row_idx].values
                df.columns = new_cols
                df = df.iloc[header_row_idx + 1:].reset_index(drop=True)

            # Normalizar nombres de columnas si vienen con tildes o espacios
            col_map = {}
            for col in df.columns:
                norm_col = str(col).lower().replace(' ', '').replace('_', '').replace('.', '')
                col_map[col] = norm_col
            df = df.rename(columns=col_map)
            
            if 'descripcionarticulo1' in df.columns and 'descripcionarticulo' not in df.columns:
                df['descripcionarticulo'] = df['descripcionarticulo1']
                
            if 'unidadesfabricadas1' in df.columns:
                df['unidadesfabricadas'] = df['unidadesfabricadas1']
            if 'unidadesrechazadas1' in df.columns:
                df['unidadesrechazadas'] = df['unidadesrechazadas1']
                
            # Mapeo de nuevas columnas del usuario a las estandar
            mapping = {
                'udsfabricadas': 'unidadesfabricadas',
                'udsafabricar': 'unidadesafabricar',
                'tiempototal': 'tiemporealtotal',
                'costereal': 'costerealtotal',
                'codigo': 'codigoarticulo',
                'descripcion': 'descripcionarticulo'
            }
            for k, v in mapping.items():
                if k in df.columns and v not in df.columns:
                    df[v] = df[k]
                
            if 'fechafinalreal' in df.columns:
                df['fecha'] = pd.to_datetime(df['fechafinalreal'], errors='coerce', dayfirst=True)
            elif 'fechainicioreal' in df.columns:
                df['fecha'] = pd.to_datetime(df['fechainicioreal'], errors='coerce', dayfirst=True)
            elif 'fechainicio' in df.columns:
                df['fecha'] = pd.to_datetime(df['fechainicio'], errors='coerce', dayfirst=True)
            else:
                if 'fecha' not in df.columns:
                    df['fecha'] = pd.NaT # Fallback
                else:
                    s_str = df['fecha'].astype(str).str.strip()
                    def safe_parse(s):
                        if pd.isna(s): return pd.NaT
                        if re.match(r'^\d{4}-\d{2}-\d{2}', str(s)):
                            return pd.to_datetime(s, errors='coerce')
                        return pd.to_datetime(s, errors='coerce', dayfirst=True)
                    df['fecha'] = s_str.apply(safe_parse)

            
            # Asegurar columnas numéricas
            numeric_cols = [
                'unidadesfabricadas', 'unidadesrechazadas', 'unidadesafabricar',
                'tiempototal', 'tiemporealtotal', 'tiempofabricacion', 'tiempopreparacion', 'tiempoparadas',
                'costetotal', 'costerealtotal', 'costerealmaquina', 'costerealmanoobra'
            ]
            
            def clean_num(x):
                if pd.isna(x): return x
                x = str(x).strip()
                if '.' in x and ',' in x:
                    if x.rfind(',') > x.rfind('.'):
                        return x.replace('.', '').replace(',', '.')
                    else:
                        return x.replace(',', '')
                elif ',' in x:
                    return x.replace(',', '.')
                return x

            for c in numeric_cols:
                if c in df.columns:
                    s = df[c]
                    if s.dtype == object:
                        s = s.apply(clean_num)
                    df[c] = pd.to_numeric(s, errors='coerce').fillna(0.0)
            
            # Limpiar strings
            for c in ['idot', 'codigoarticulo', 'descripcionarticulo']:
                if c in df.columns:
                    df[c] = df[c].fillna('').astype(str).str.strip()
            
        # Clean up empty rows (e.g. totals or empty sheet rows)
        if 'cliente' in df.columns:
            df = df[df['cliente'].astype(str).str.strip() != '']
        if 'documento' in df.columns:
            df = df[df['documento'].astype(str).str.strip() != '']


        return df
def supabase_request_all(path: str, query_params: dict | None = None) -> list:
    if query_params is None:
        query_params = {}
    all_data = []
    limit = 1000
    offset = 0
    while True:
        params = query_params.copy()
        params['limit'] = str(limit)
        params['offset'] = str(offset)
        chunk = supabase_request(path, method='GET', query_params=params)
        if not chunk or not isinstance(chunk, list):
            break
        all_data.extend(chunk)
        if len(chunk) < limit:
            break
        offset += limit
    return all_data

def load_data_from_supabase(report_date: str) -> dict[str, pd.DataFrame] | None:
    dfs = {}
    for name in ['ofertas', 'pedidos', 'albaranes', 'facturas', 'produccion', 'stock']:
        try:
            records = supabase_request_all(name, query_params={'report_date': f'eq.{report_date}', 'select': '*'})
        except Exception as e:
            print(f'Error cargando {name} de Supabase: {e}')
            records = []
            
        if not records:
            dfs[name] = pd.DataFrame()
        else:
            dfs[name] = pd.DataFrame(records)
        dfs[name] = ensure_types_normalized_df(dfs[name], name)
    try:
        comments_records = supabase_request_all('document_comments', query_params={'select': '*'})
        dfs['document_comments'] = pd.DataFrame(comments_records)
    except Exception as e:
        print(f'Error cargando document_comments de Supabase: {e}')
        dfs['document_comments'] = pd.DataFrame()
    if all((df.empty for df in dfs.values())):
        return
    else:
        return dfs
def get_default_report_date() -> str:
    now = datetime.now()
    weekday = now.weekday()
    if weekday == 0:
        delta = 3
    else:
        if weekday == 6:
            delta = 2
        else:
            if weekday == 5:
                delta = 1
            else:
                delta = 1
    default_date = now - pd.Timedelta(days=delta)
    return default_date.strftime('%Y-%m-%d')
def adjust_report_date(date_str: str) -> str:
    try:
        dt = pd.to_datetime(date_str)
        weekday = dt.weekday()
        if weekday == 5:
            return (dt - pd.Timedelta(days=1)).strftime('%Y-%m-%d')
        elif weekday == 6:
            return (dt - pd.Timedelta(days=2)).strftime('%Y-%m-%d')
        elif weekday == 0:
            today = datetime.now()
            if today.weekday() in [0, 5, 6]:
                return (dt - pd.Timedelta(days=3)).strftime('%Y-%m-%d')
        return date_str
    except Exception:
        return date_str

def parse_excel_to_normalized_df(source: bytes | str, kind: str) -> pd.DataFrame:
    ds = read_excel(source, kind)
    df = ds.df.copy()
    if kind == 'ofertas':
        date_col = ds.require_date('FechaOferta', 'Fecha oferta')
        amount_col = ds.amount_col('ImporteNeto', 'ImporteBruto', 'Importe')
        number_cols = [ds.col('EjercicioOferta'), ds.col('SerieOferta'), ds.col('NumeroOferta')]
    else:
        if kind == 'pedidos':
            date_col = ds.require_date('FechaPedido', 'Fecha pedido')
            amount_col = ds.amount_col('ImporteBruto', 'ImporteNeto', 'Importe')
            pending_amount_col = ds.amount_col('ImporteBrutoPendiente', 'Importe pendiente')
            number_cols = [ds.col('EjercicioPedido'), ds.col('SeriePedido'), ds.col('NumeroPedido')]
            needed_date_col = ds.require_date('FechaNecesaria', 'Fecha necesaria')
            ordered_col = ds.amount_col('UnidadesPedidas')
            served_col = ds.amount_col('UnidadesServidas')
            pending_col = ds.amount_col('UnidadesPendientes')
        else:
            if kind == 'albaranes':
                date_col = ds.require_date('FechaAlbaran', 'Fecha albaran')
                amount_col = ds.amount_col('ImporteLiquido', 'BaseImponible', 'Importe')
                number_cols = [ds.col('EjercicioAlbaran'), ds.col('SerieAlbaran'), ds.col('NumeroAlbaran')]
                pending_amount_col = needed_date_col = None
                ordered_col = served_col = pending_col = None
            elif kind in ('produccion', 'stock'):
                return ds.df
            else:
                date_col = ds.require_date('Fecha factura', 'FechaFactura')
                amount_col = ds.amount_col('Importe liquido', 'ImporteLiquido', 'BaseImponible', 'Importe')
                number_cols = [ds.col('Ejercicio'), ds.col('Serie factura'), ds.col('N factura', 'Nº factura')]
                pending_amount_col = needed_date_col = None
                ordered_col = served_col = pending_col = None
    if date_col is None:
        if kind == 'ofertas':
            expected = "'FechaOferta' o 'Fecha oferta'"
        elif kind == 'pedidos':
            expected = "'FechaPedido' o 'Fecha pedido'"
        elif kind == 'albaranes':
            expected = "'FechaAlbaran' o 'Fecha albaran'"
        else:
            expected = "'Fecha factura' o 'FechaFactura'"
        raise ValueError(f"El archivo subido no corresponde a la categoría '{kind}'. No se encontró ninguna columna de fecha esperada (como {expected}). Por favor, verifica el archivo seleccionado.")
    if kind == 'ofertas':
        pending_amount_col = needed_date_col = None
        ordered_col = served_col = pending_col = None
    client_col = ds.col('CodigoCliente', 'Cod. cliente', 'Cód. cliente')
    name_col = ds.col('RazonSocial', 'Razon social', 'Razón social', 'Nombre')
    article_col = ds.col('CodigoArticulo', 'Codigo articulo', 'Artículo')
    description_col = ds.col('DescripcionArticulo', 'Descripcion articulo', 'Descripción artículo', 'Descripcion', 'Descripción')
    series_col = ds.col('SeriePedido', 'Serie pedido', 'SerieOferta', 'Serie factura', 'SerieAlbaran')
    cif_col = ds.col('CIF europeo', 'CIF Europeo', 'CIF') if kind == 'facturas' else None
    out = pd.DataFrame(index=df.index)
    out['documento'] = doc_key(df, number_cols, kind)
    out['fecha'] = df[date_col] if date_col else pd.NaT
    out['importe'] = df[amount_col] if amount_col else 0.0
    out['cliente'] = df[client_col].fillna('').astype(str).str.replace('\\.0$', '', regex=True).str.strip() if client_col else ''
    out['razon_social'] = df[name_col].fillna('').astype(str).str.strip() if name_col else ''
    out['articulo'] = df[article_col].fillna('').astype(str).str.strip() if article_col else ''
    if kind in ['pedidos', 'ofertas']:
        out['descripcion'] = df[description_col].fillna('').astype(str).str.strip() if description_col else ''
    out['serie'] = df[series_col].fillna('').astype(str).str.strip() if series_col else ''
    if kind == 'facturas':
        out['cif'] = df[cif_col].fillna('').astype(str).str.strip() if cif_col else ''
        out['zona'] = out['cif'].apply(lambda val: 'Nacional' if val.upper().startswith('ES') else 'Exportación')
    else:
        out['zona'] = out['serie'].apply(lambda val: 'Exportación' if 'EX' in norm(val).upper() else 'Nacional')
    if kind == 'pedidos':
        out['fecha_necesaria'] = df[needed_date_col] if needed_date_col else pd.NaT
        out['importe_pendiente'] = df[pending_amount_col] if pending_amount_col else out['importe']
        out['unidades_pedidas'] = df[ordered_col] if ordered_col else 0.0
        out['unidades_servidas'] = df[served_col] if served_col else 0.0
        out['unidades_pendientes'] = df[pending_col] if df[pending_col] is not None else 0.0
    # VAT Adjustment
    if kind == 'facturas' and 'cif' in out.columns and 'importe' in out.columns:
        out['importe'] = pd.to_numeric(out['importe'], errors='coerce').fillna(0.0)
        mask = out['cif'].astype(str).str.upper().str.startswith('ES')
        out.loc[mask, 'importe'] = out.loc[mask, 'importe'] / 1.10
    elif kind in ('albaranes', 'ofertas', 'pedidos') and 'serie' in out.columns:
        out['importe'] = pd.to_numeric(out['importe'], errors='coerce').fillna(0.0)
        mask = out['serie'].astype(str).str.strip() == ''
        out.loc[mask, 'importe'] = out.loc[mask, 'importe'] / 1.10
        if kind == 'pedidos' and 'importe_pendiente' in out.columns:
            out['importe_pendiente'] = pd.to_numeric(out['importe_pendiente'], errors='coerce').fillna(0.0)
            out.loc[mask, 'importe_pendiente'] = out.loc[mask, 'importe_pendiente'] / 1.10
    return out

def aggregate_normalized_df(df: pd.DataFrame, kind: str) -> pd.DataFrame:
    if kind in ('produccion', 'stock'):
        return df
    if df.empty:
        return pd.DataFrame(columns=['documento', 'fecha', 'fecha_necesaria', 'importe', 'importe_pendiente', 'cliente', 'razon_social', 'serie', 'zona', 'articulos', 'unidades_pedidas', 'unidades_servidas', 'unidades_pendientes', 'lineas', 'articulos_list', 'cif'])
    else:
        df_copy = df.copy()
        if 'articulo' in df_copy.columns and 'descripcion' in df_copy.columns:
            df_copy['articulo_detalle'] = df_copy['articulo'] + " - " + df_copy['descripcion']
        elif 'articulo' in df_copy.columns:
            df_copy['articulo_detalle'] = df_copy['articulo']
        else:
            df_copy['articulo_detalle'] = ''
            
        agg_dict = {
            'fecha': ('fecha', 'min'), 
            'importe': ('importe', 'sum'), 
            'cliente': ('cliente', 'first'), 
            'razon_social': ('razon_social', 'first'), 
            'serie': ('serie', 'first'), 
            'zona': ('zona', 'first'), 
            'articulos': ('articulo', lambda s: {x for x in s}), 
            'articulos_list': ('articulo_detalle', lambda s: list(set(s)) if s is not None else []),
            'lineas': ('documento', 'size')
        }
        if 'cif' in df_copy.columns:
            agg_dict['cif'] = ('cif', 'first')
        if kind == 'pedidos':
            agg_dict.update({'fecha_necesaria': ('fecha_necesaria', 'min'), 'importe_pendiente': ('importe_pendiente', 'sum'), 'unidades_pedidas': ('unidades_pedidas', 'sum'), 'unidades_servidas': ('unidades_servidas', 'sum'), 'unidades_pendientes': ('unidades_pendientes', 'sum')})
        grouped = df_copy.groupby('documento', dropna=False).agg(**agg_dict).reset_index()
        if 'fecha_necesaria' not in grouped.columns:
            grouped['fecha_necesaria'] = pd.NaT
        if 'importe_pendiente' not in grouped.columns:
            grouped['importe_pendiente'] = grouped['importe']
        if 'unidades_pedidas' not in grouped.columns:
            grouped['unidades_pedidas'] = 0.0
        if 'unidades_servidas' not in grouped.columns:
            grouped['unidades_servidas'] = 0.0
        if 'unidades_pendientes' not in grouped.columns:
            grouped['unidades_pendientes'] = 0.0
        return grouped
DEFAULT_FILES = {'ofertas': str(BASE_DIR / 'ofertas.xlsx'), 'pedidos': str(BASE_DIR / 'pedidos.xlsx'), 'albaranes': str(BASE_DIR / 'albaranes.xlsx'), 'facturas': str(BASE_DIR / 'facturas.xlsx'), 'produccion': str(BASE_DIR / 'produccion.xlsx'), 'stock': str(BASE_DIR / 'stock.xlsx')}
def norm(value: Any) -> str:
    text = '' if value is None else str(value)
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')
    text = re.sub('[^a-zA-Z0-9]+', '', text).lower()
    return text
def money(val: float) -> str:
    formatted = f"{int(round(val)):,} EUR".replace(',', '.')
    return formatted 
def pct(value: float | None) -> str:
    if value is None:
        return 'N/D'
    return f"{int(round(value * 100))}%"
def fmt_date(value: Any) -> str:
    if pd.isna(value):
        return ''
    else:
        return pd.to_datetime(value).strftime('%d/%m/%Y')
def fmt_date_or(value: Any, empty: str='Sin fecha') -> str:
    formatted = fmt_date(value)
    return formatted if formatted else empty
def clean_text(value: Any) -> str:
    if pd.isna(value):
        return ''
    else:
        text = str(value).strip()
        return '' if text.lower() in {'none', 'nat', 'nan'} else text
class DataSet:
    def __init__(self, name: str, df: pd.DataFrame):
        self.name = name
        self.df = df.copy()
        self.cols = {norm(col): col for col in self.df.columns}
    def col(self, *candidates: str) -> str | None:
        for candidate in candidates:
            key = norm(candidate)
            if key in self.cols:
                return self.cols[key]
        for candidate in candidates:
            key = norm(candidate)
            for ncol, original in self.cols.items():
                if key and (key in ncol or ncol in key):
                        return original
        return
    def require_date(self, *candidates: str) -> str | None:
        col = self.col(*candidates)
        if col:
            self.df[col] = pd.to_datetime(self.df[col], errors='coerce')
        return col
    def amount_col(self, *candidates: str) -> str | None:
        col = self.col(*candidates)
        if col:
            self.df[col] = pd.to_numeric(self.df[col], errors='coerce').fillna(0.0)
        return col
def read_excel(source: bytes | str, name: str) -> DataSet:
    if isinstance(source, bytes):
        source = io.BytesIO(source)
    df = pd.read_excel(source, sheet_name=0)
    df = df.dropna(how='all')
    return DataSet(name, df)
def doc_key(df: pd.DataFrame, columns: list[str | None], fallback_prefix: str) -> pd.Series:
    present = [col for col in columns if col and col in df.columns]
    if not present:
        return pd.Series([f'{fallback_prefix}-{i}' for i in range(len(df))], index=df.index)
    else:
        pieces = []
        for col in present:
            pieces.append(df[col].fillna('').astype(str).str.replace('\\.0$', '', regex=True).str.strip())
        key = pieces[0]
        for piece in pieces[1:]:
            key = key + '/' + piece
        return key.replace('^/+$', '', regex=True)
def aggregate_docs(ds: DataSet, kind: str) -> pd.DataFrame:
    # ***<module>.aggregate_docs: Failure: Different bytecode
    df = ds.df.copy()
    if kind == 'ofertas':
        date_col = ds.require_date('FechaOferta', 'Fecha oferta')
        amount_col = ds.amount_col('ImporteNeto', 'ImporteBruto', 'Importe')
        number_cols = [ds.col('EjercicioOferta'), ds.col('SerieOferta'), ds.col('NumeroOferta')]
    else:
        if kind == 'pedidos':
            date_col = ds.require_date('FechaPedido', 'Fecha pedido')
            amount_col = ds.amount_col('ImporteBruto', 'ImporteNeto', 'Importe')
            pending_amount_col = ds.amount_col('ImporteBrutoPendiente', 'Importe pendiente')
            number_cols = [ds.col('EjercicioPedido'), ds.col('SeriePedido'), ds.col('NumeroPedido')]
            needed_date_col = ds.require_date('FechaNecesaria', 'Fecha necesaria')
            ordered_col = ds.amount_col('UnidadesPedidas')
            served_col = ds.amount_col('UnidadesServidas')
            pending_col = ds.amount_col('UnidadesPendientes')
        else:
            if kind == 'albaranes':
                date_col = ds.require_date('FechaAlbaran', 'Fecha albaran')
                amount_col = ds.amount_col('ImporteLiquido', 'BaseImponible', 'Importe')
                number_cols = [ds.col('EjercicioAlbaran'), ds.col('SerieAlbaran'), ds.col('NumeroAlbaran')]
                pending_amount_col = needed_date_col = None
                ordered_col = served_col = pending_col = None
            else:
                date_col = ds.require_date('Fecha factura', 'FechaFactura')
                amount_col = ds.amount_col('Importe liquido', 'ImporteLiquido', 'BaseImponible', 'Importe')
                number_cols = [ds.col('Ejercicio'), ds.col('Serie factura'), ds.col('N factura', 'Nº factura')]
                pending_amount_col = needed_date_col = None
                ordered_col = served_col = pending_col = None
    if kind == 'ofertas':
        pending_amount_col = needed_date_col = None
        ordered_col = served_col = pending_col = None
    client_col = ds.col('CodigoCliente', 'Cod. cliente', 'Cód. cliente')
    name_col = ds.col('RazonSocial', 'Razon social', 'Razón social', 'Nombre')
    article_col = ds.col('CodigoArticulo', 'Codigo articulo', 'Artículo')
    series_col = ds.col('SeriePedido', 'Serie pedido', 'SerieOferta', 'Serie factura', 'SerieAlbaran')
    out = df.copy()
    out['_doc'] = doc_key(out, number_cols, kind)
    out['_date'] = out[date_col] if date_col else pd.NaT
    out['_amount'] = out[amount_col] if amount_col else 0.0
    out['_pending_amount'] = out[pending_amount_col] if pending_amount_col else out['_amount']
    out['_needed_date'] = out[needed_date_col] if needed_date_col else pd.NaT
    out['_client'] = out[client_col].fillna('').astype(str).str.replace('\\.0$', '', regex=True).str.strip() if client_col else ''
    out['_name'] = out[name_col].fillna('').astype(str).str.strip() if name_col else ''
    out['_article'] = out[article_col].fillna('').astype(str).str.strip() if article_col else ''
    out['_series'] = out[series_col].fillna('').astype(str).str.strip() if series_col else ''
    out['_geo'] = out['_series'].apply(lambda value: 'Exportación' if 'EX' in norm(value).upper() else 'Nacional')
    out['_units_ordered'] = out[ordered_col] if ordered_col else 0.0
    out['_units_served'] = out[served_col] if served_col else 0.0
    out['_units_pending'] = out[pending_col] if pending_col else 0.0
    grouped = out.groupby('_doc', dropna=False).agg(fecha=('_date', 'min'), fecha_necesaria=('_needed_date', 'min'), importe=('_amount', 'sum'), importe_pendiente=('_pending_amount', 'sum'), cliente=('_client', 'first'), razon_social=('_name', 'first'), serie=('_series', 'first'), zona=('_geo', 'first'), articulos='_article', unidades_pedidas=lambda s: {x for x in s}, unidades_servidas=('_units_ordered', 'sum'), unidades_pendientes=('_units_served', 'sum'), lineas=('_units_pending', 'sum'), reset_index=('_doc', 'size')).reset_index().rename(columns={'_doc': 'documento'})
    return grouped
def detect_analysis_date() -> pd.Timestamp:
    return pd.Timestamp(datetime.now().date()) - pd.Timedelta(days=1)
def is_same_month(series: pd.Series, current: pd.Timestamp) -> pd.Series:
    dates = pd.to_datetime(series, errors='coerce')
    return (dates.dt.year == current.year) & (dates.dt.month == current.month)
def converted_offers(ofertas: pd.DataFrame, pedidos: pd.DataFrame) -> pd.Series:
    converted = []
    pedido_rows = pedidos.copy()
    for _, offer in ofertas.iterrows():
        same_client = pedido_rows[pedido_rows['cliente'] == offer['cliente']]
        if not same_client.empty and offer['articulos']:
            found = same_client['articulos'].apply(lambda arts: bool(offer['articulos'] & arts)).any()
        else:
            found = not same_client.empty
        converted.append(bool(found))
    return pd.Series(converted, index=ofertas.index)
def classify_order_status(pedidos: pd.DataFrame, albaranes: pd.DataFrame, facturas: pd.DataFrame, current: pd.Timestamp) -> pd.DataFrame:
    rows = []
    for _, pedido in pedidos.iterrows():
        served = float(pedido.get('unidades_servidas', 0) or 0)
        pending = float(pedido.get('unidades_pendientes', 0) or 0)
        if served > 0 and pending > 0:
            estado = 'Parcialmente servido'
        else:
            estado = 'En Preparación'
        rows.append({**pedido.to_dict(), 'estado_operativo': estado})
    return pd.DataFrame(rows)
def recent_orders_missing_needed(pedidos_lines: pd.DataFrame, current: pd.Timestamp) -> list[dict[str, Any]]:
    if pedidos_lines.empty:
        return []
    else:
        fechas = pd.to_datetime(pedidos_lines['fecha'], errors='coerce')
        necesaria = pd.to_datetime(pedidos_lines['fecha_necesaria'], errors='coerce')
        mask = fechas.notna() & (current - fechas).dt.days.between(0, 30) & necesaria.isna()
        subset = pedidos_lines[mask].copy()
        if subset.empty:
            return []
        else:
            grouped = subset.groupby('documento', dropna=False).first().reset_index()
            out = []
            for _, row in grouped.head(25).iterrows():
                out.append({'documento': row['documento'], 'fecha': row['fecha'], 'cliente': row['cliente'], 'razon_social': row['razon_social']})
            return out
def stale_delivery_notes(albaranes: pd.DataFrame, facturas: pd.DataFrame, current: pd.Timestamp) -> list[dict[str, Any]]:
    rows = []
    cutoff = current - pd.Timedelta(days=7)
    for _, alb in albaranes[(albaranes['fecha'] < cutoff) & (albaranes['fecha'] <= current)].iterrows():
        rows.append(alb.to_dict())
    return rows[:25]
def stagnant_offers(ofertas: pd.DataFrame, pedidos: pd.DataFrame, current: pd.Timestamp) -> list[dict[str, Any]]:
    if ofertas.empty:
        return []
    else:
        conv = converted_offers(ofertas, pedidos)
        mask = ~conv & ofertas['fecha'].notna() & ((current - ofertas['fecha']).dt.days > 15) & (ofertas['importe'] > 15000)
        return ofertas[mask].sort_values('importe', ascending=False).head(25).to_dict('records')
def daily_month_series(frame: pd.DataFrame, current: pd.Timestamp) -> list[dict[str, Any]]:
    if frame.empty:
        return []
    else:
        scoped = frame[is_same_month(frame['fecha'], current) & (frame['fecha'] <= current)].copy()
        if scoped.empty:
            return []
        else:
            grouped = scoped.groupby('fecha', dropna=True).agg(cantidad=('documento', 'count'), importe=('importe', 'sum')).reset_index()
            return [{'fecha': row['fecha'], 'cantidad': int(row['cantidad']), 'importe': float(row['importe'])} for _, row in grouped.sort_values('fecha').iterrows()]
def month_end(current: pd.Timestamp) -> pd.Timestamp:
    return current + pd.offsets.MonthEnd(0)
def pending_delivery_notes_for_invoice(albaranes: pd.DataFrame, facturas: pd.DataFrame, current: pd.Timestamp) -> pd.DataFrame:
    pending_albaranes = albaranes[albaranes['fecha'] <= current].copy()
    if pending_albaranes.empty or facturas.empty:
        return pending_albaranes
    else:
        facturas_sorted = facturas.sort_values('fecha')
        matched_albaranes_indices = []
        used_factura_docs = set()
        for idx, alb in pending_albaranes.sort_values('fecha').iterrows():
            same_client_amount = facturas_sorted[(facturas_sorted['cliente'] == alb['cliente']) & ((facturas_sorted['importe'] - alb['importe']).abs() < 0.05) & (facturas_sorted['fecha'] >= alb['fecha']) & ~facturas_sorted['documento'].isin(used_factura_docs)]
            if not same_client_amount.empty:
                matched_factura = same_client_amount.iloc[0]
                used_factura_docs.add(matched_factura['documento'])
                matched_albaranes_indices.append(idx)
        return pending_albaranes.drop(matched_albaranes_indices)
def loadable_orders_this_month(pedidos: pd.DataFrame, albaranes: pd.DataFrame, facturas: pd.DataFrame, current: pd.Timestamp) -> pd.DataFrame:
    today_real = pd.Timestamp(datetime.now().date())
    if (today_real.year > current.year) or (today_real.year == current.year and today_real.month > current.month):
        return pd.DataFrame(columns=pedidos.columns)
    current_month_end = month_end(current)
    end = current_month_end + pd.Timedelta(days=15)
    oldest_order_date = current - pd.DateOffset(months=1)
    rows = []
    for _, pedido in pedidos.iterrows():
        order_date = pedido.get('fecha')
        if pd.isna(order_date) or order_date < oldest_order_date:
            continue
        else:
            needed = pedido.get('fecha_necesaria')
            estimated_needed = False
            if pd.isna(needed):
                order_amount = float(pedido.get('importe', 0))
                if order_amount > 50000:
                    needed = pd.to_datetime(order_date) + pd.Timedelta(days=15)
                else:
                    needed = pd.to_datetime(order_date) + pd.Timedelta(days=7)
                estimated_needed = True
            if needed > end:
                continue
            else:
                pending_units = float(pedido.get('unidades_pendientes', 0) or 0)
                pending_amount = float(pedido.get('importe_pendiente', pedido.get('importe', 0)) or 0)
                if pending_units <= 0 or pending_amount <= 0:
                    continue
                else:
                    row = pedido.to_dict()
                    row['fecha_carga_prevista'] = needed
                    row['fecha_carga_estimada'] = estimated_needed
                    row['backlog_disponible'] = needed < current
                    row['siguiente_mes'] = needed > current_month_end
                    rows.append(row)
    return pd.DataFrame(rows)
def older_pending_orders_this_month(pedidos: pd.DataFrame, current: pd.Timestamp) -> pd.DataFrame:
    today_real = pd.Timestamp(datetime.now().date())
    if (today_real.year > current.year) or (today_real.year == current.year and today_real.month > current.month):
        return pd.DataFrame(columns=pedidos.columns)
    end = month_end(current)
    oldest_order_date = current - pd.DateOffset(months=1)
    rows = []
    for _, pedido in pedidos.iterrows():
        order_date = pedido.get('fecha')
        if pd.isna(order_date) or order_date >= oldest_order_date:
            continue
        else:
            needed = pedido.get('fecha_necesaria')
            if pd.isna(needed):
                continue
            else:
                if needed > end:
                    continue
                else:
                    pending_units = float(pedido.get('unidades_pendientes', 0) or 0)
                    pending_amount = float(pedido.get('importe_pendiente', pedido.get('importe', 0)) or 0)
                    if pending_units <= 0 or pending_amount <= 0:
                        continue
                    else:
                        row = pedido.to_dict()
                        row['fecha_carga_prevista'] = needed
                        row['fecha_carga_estimada'] = False
                        row['backlog_disponible'] = needed < current
                        rows.append(row)
    return pd.DataFrame(rows)
def build_delivery_schedule(pedidos: pd.DataFrame, ofertas: pd.DataFrame, current: pd.Timestamp) -> list[dict[str, Any]]:
    oldest_order_date = current - pd.DateOffset(months=1)
    schedule = []
    pending_pedidos = pedidos[pedidos['importe_pendiente'] > 0]
    for _, order in pending_pedidos.iterrows():
        order_date = order['fecha']
        is_old = order_date < oldest_order_date
        needed = order.get('fecha_necesaria')
        estimated = False
        if is_old:
            if pd.isna(needed):
                expected_delivery = None
                expected_delivery_str = 'Entregas parciales (Sin fecha concreta)'
            else:
                expected_delivery = pd.to_datetime(needed)
                expected_delivery_str = fmt_date_or(expected_delivery) + ' (Fecha Necesaria)'
        else:
            if pd.isna(needed):
                order_amount = float(order.get('importe', 0))
                if pd.isna(order_date):
                    expected_delivery = None
                else:
                    if order_amount > 50000:
                        expected_delivery = pd.to_datetime(order_date) + pd.Timedelta(days=15)
                    else:
                        expected_delivery = pd.to_datetime(order_date) + pd.Timedelta(days=7)
                estimated = True
            else:
                expected_delivery = pd.to_datetime(needed)
            expected_delivery_str = fmt_date_or(expected_delivery) + (' (Estimada)' if estimated else ' (Fecha Necesaria)')
        schedule.append({'tipo': 'Pedido', 'documento': order['documento'], 'cliente': order['razon_social'], 'fecha_creacion': order_date, 'fecha_aceptacion': None, 'fecha_entrega': expected_delivery, 'fecha_entrega_str': expected_delivery_str, 'importe': float(order['importe_pendiente'])})
    for _, offer in ofertas.iterrows():
        same_client = pedidos[pedidos['cliente'] == offer['cliente']]
        if same_client.empty:
            continue
        else:
            matched_order = None
            if offer['articulos']:
                for _, order in same_client.iterrows():
                    if offer['articulos'] & order['articulos']:
                        matched_order = order
                        break
            if matched_order is None:
                matched_order = same_client.iloc[0]
            acceptance_date = matched_order['fecha']
            expected_delivery = None if pd.isna(acceptance_date) else pd.to_datetime(acceptance_date) + pd.Timedelta(days=15)
            expected_delivery_str = fmt_date_or(expected_delivery) + ' (Estimada +15d)'
            schedule.append({'tipo': 'Oferta', 'documento': offer['documento'], 'cliente': offer['razon_social'], 'fecha_creacion': offer['fecha'], 'fecha_aceptacion': acceptance_date, 'fecha_entrega': expected_delivery, 'fecha_entrega_str': expected_delivery_str, 'importe': float(offer['importe'])})
    def sort_key(item):
        date = item['fecha_entrega']
        if date is None or pd.isna(date):
            return (1, pd.Timestamp.max)
        else:
            return (0, date)
    schedule.sort(key=sort_key)
    return schedule
def top_rows(frame: pd.DataFrame, amount_col: str, limit: int=12) -> list[dict[str, Any]]:
    if frame.empty:
        return []
    else:
        return frame.sort_values(amount_col, ascending=False).head(limit).to_dict('records')
def approved_offers_with_theoretical_delivery(ofertas: pd.DataFrame, pedidos: pd.DataFrame, current: pd.Timestamp) -> pd.DataFrame:
    if ofertas.empty:
        return pd.DataFrame()
    else:
        converted = converted_offers(ofertas, pedidos)
        approved = ofertas[converted].copy()
        if approved.empty:
            return approved
        else:
            approved['fecha_entrega_teorica'] = approved['fecha'] + pd.Timedelta(days=15)
            approved['entrega_en_mes'] = is_same_month(approved['fecha_entrega_teorica'], current)
            return approved
def get_same_day_prev_month(current: pd.Timestamp) -> pd.Timestamp:
    year = current.year
    month = current.month - 1
    if month == 0:
        month = 12
        year -= 1
    day = current.day
    while True:
        try:
            return pd.Timestamp(year=year, month=month, day=day)
        except ValueError:
            day -= 1
        else:
            pass
def calculate_mtd_metrics(dfs: dict[str, pd.DataFrame], date: pd.Timestamp) -> dict[str, float]:
    docs = {name: aggregate_normalized_df(df, name) for name, df in dfs.items()}
    ofertas = docs['ofertas']
    pedidos = docs['pedidos']
    albaranes = docs['albaranes']
    facturas = docs['facturas']
    month_start = pd.Timestamp(year=date.year, month=date.month, day=1)
    if not facturas.empty:
        facturas_mtd = facturas[(facturas['fecha'] >= month_start) & (facturas['fecha'] <= date)]
        facturado_val = float(facturas_mtd['importe'].sum())
        facturado_split = calc_split(facturas_mtd, 'importe')
    else:
        facturado_val = 0.0
        facturado_split = {'nacional': 0.0, 'exportacion': 0.0}
    pending_albaranes = pending_delivery_notes_for_invoice(albaranes, facturas, date)
    if not pending_albaranes.empty:
        albaranes_mtd = pending_albaranes[pending_albaranes['fecha'] <= date]
        albaranes_val = float(albaranes_mtd['importe'].sum())
        albaranes_split = calc_split(albaranes_mtd, 'importe')
    else:
        albaranes_val = 0.0
        albaranes_split = {'nacional': 0.0, 'exportacion': 0.0}
    if not pedidos.empty:
        pedidos_mtd = pedidos[pedidos['fecha'] <= date]
        pedidos_val = float(pedidos_mtd['importe_pendiente'].sum())
        pedidos_split = calc_split(pedidos_mtd, 'importe_pendiente')
    else:
        pedidos_val = 0.0
        pedidos_split = {'nacional': 0.0, 'exportacion': 0.0}
    if not ofertas.empty:
        conv = converted_offers(ofertas, pedidos)
        pending_offers = ofertas[~conv]
        if not pending_offers.empty:
            offers_mtd = pending_offers[pending_offers['fecha'] <= date]
            ofertas_val = float(offers_mtd['importe'].sum())
            ofertas_split = calc_split(offers_mtd, 'importe')
        else:
            ofertas_val = 0.0
            ofertas_split = {'nacional': 0.0, 'exportacion': 0.0}
    else:
        ofertas_val = 0.0
        ofertas_split = {'nacional': 0.0, 'exportacion': 0.0}
    total_comercial = facturado_val + albaranes_val + pedidos_val + ofertas_val
    total_split = {
        'nacional': facturado_split['nacional'] + albaranes_split['nacional'] + pedidos_split['nacional'] + ofertas_split['nacional'],
        'exportacion': facturado_split['exportacion'] + albaranes_split['exportacion'] + pedidos_split['exportacion'] + ofertas_split['exportacion']
    }
    return {
        'facturado': facturado_val, 'facturado_split': facturado_split,
        'albaranes': albaranes_val, 'albaranes_split': albaranes_split,
        'pedidos': pedidos_val, 'pedidos_split': pedidos_split,
        'ofertas': ofertas_val, 'ofertas_split': ofertas_split,
        'total': total_comercial, 'total_split': total_split
    }
def get_mtd_comparison(current_dfs: dict[str, pd.DataFrame], current_date: pd.Timestamp) -> dict[str, Any]:
    current_metrics = calculate_mtd_metrics(current_dfs, current_date)
    prev_date = get_same_day_prev_month(current_date)
    
    # Buscar el día previo más cercano con datos en el mes anterior
    prev_dfs = None
    target_prev_month = prev_date.month
    while prev_date.month == target_prev_month and prev_date.day >= 1:
        prev_date_str = prev_date.strftime('%Y-%m-%d')
        try:
            prev_dfs = load_report_data(prev_date_str)
        except Exception as e:
            prev_dfs = None
        if prev_dfs is not None and not prev_dfs.get('pedidos', pd.DataFrame()).empty and not prev_dfs.get('facturas', pd.DataFrame()).empty:
            break
        prev_date = prev_date - pd.Timedelta(days=1)
        
    if prev_dfs is not None:
        prev_metrics = calculate_mtd_metrics(prev_dfs, prev_date)
        growth = {}
        for key in ['facturado', 'albaranes', 'pedidos', 'ofertas', 'total']:
            cur_val = current_metrics[key]
            prv_val = prev_metrics[key]
            if prv_val > 0:
                growth[key] = (cur_val - prv_val) / prv_val
            else:
                growth[key] = None
    else:
        prev_metrics = None
        growth = {key: None for key in ['facturado', 'albaranes', 'pedidos', 'ofertas', 'total']}
    return {'current_date': current_date, 'prev_date': prev_date, 'current': current_metrics, 'prev': prev_metrics, 'growth': growth}
def get_annual_accumulations(current_date_str: str, current_dfs: dict[str, pd.DataFrame], current_date: pd.Timestamp) -> dict[str, Any]:
    # ***<module>.get_annual_accumulations: Failure: Different control flow
    try:
        records = supabase_request('facturas', query_params={'select': 'report_date'})
        if records:
            report_dates = list(set((r['report_date'] for r in records if r.get('report_date'))))
        else:
            report_dates = [current_date_str]
    except Exception as e:
        print(f'Error fetching report dates: {e}')
        report_dates = [current_date_str]
    if current_date_str not in report_dates:
        report_dates.append(current_date_str)
    current_year = current_date.year
    valid_dates = []
    for d_str in report_dates:
        try:
            dt = pd.to_datetime(d_str)
            if dt.year == current_year and dt <= current_date:
                    valid_dates.append((dt, d_str))
        except Exception:
            pass
    by_month = {}
    for dt, d_str in valid_dates:
        key = (dt.year, dt.month)
        by_month.setdefault(key, [])
        by_month[key].append((dt, d_str))
    selected_dates = []
    for (yr, mn), items in by_month.items():
        if yr < current_year or mn < current_date.month:
            latest = max(items, key=lambda x: x[0])
            selected_dates.append(latest[1])
    if current_date_str not in selected_dates:
        selected_dates.append(current_date_str)
    dates_str = ','.join(selected_dates)
    facturas_ytd = pd.DataFrame()
    try:
        ytd_records = supabase_request('facturas', query_params={'report_date': f'in.({dates_str})', 'select': '*'})
        if ytd_records:
            facturas_ytd = pd.DataFrame(ytd_records)
        else:
            facturas_ytd = pd.DataFrame()
    except Exception as e:
        print(f'Error fetching YTD invoices: {e}')
    facturas_ytd = ensure_types_normalized_df(facturas_ytd, 'facturas')
    docs = {name: aggregate_normalized_df(df, name) for name, df in current_dfs.items()}
    ofertas = docs['ofertas']
    pedidos = docs['pedidos']
    albaranes = docs['albaranes']
    facturas_curr = docs['facturas']
    pending_albaranes = pending_delivery_notes_for_invoice(albaranes, facturas_curr, current_date)
    pending_orders = pedidos[pedidos['importe_pendiente'] > 0]
    if not ofertas.empty:
        conv = converted_offers(ofertas, pedidos)
        pending_offers = ofertas[~conv]
    else:
        pending_offers = pd.DataFrame()
    if not facturas_ytd.empty:
        grp_fact = facturas_ytd.groupby('cliente').agg(facturado_ytd=('importe', 'sum'), razon_social=('razon_social', 'first'), zona=('zona', 'first')).reset_index()
    else:
        grp_fact = pd.DataFrame(columns=['cliente', 'facturado_ytd', 'razon_social', 'zona'])
    if not pending_albaranes.empty:
        grp_alb = pending_albaranes.groupby('cliente').agg(albaranes_pending=('importe', 'sum'), razon_social=('razon_social', 'first'), zona=('zona', 'first')).reset_index()
    else:
        grp_alb = pd.DataFrame(columns=['cliente', 'albaranes_pending', 'razon_social', 'zona'])
    if not pending_orders.empty:
        grp_ped = pending_orders.groupby('cliente').agg(pedidos_pending=('importe_pendiente', 'sum'), razon_social=('razon_social', 'first'), zona=('zona', 'first')).reset_index()
    else:
        grp_ped = pd.DataFrame(columns=['cliente', 'pedidos_pending', 'razon_social', 'zona'])
    if not pending_offers.empty:
        grp_ofe = pending_offers.groupby('cliente').agg(ofertas_pending=('importe', 'sum'), razon_social=('razon_social', 'first'), zona=('zona', 'first')).reset_index()
    else:
        grp_ofe = pd.DataFrame(columns=['cliente', 'ofertas_pending', 'razon_social', 'zona'])
    all_clients = set(grp_fact['cliente']) | set(grp_alb['cliente']) | set(grp_ped['cliente']) | set(grp_ofe['cliente'])
    all_clients = {c for c in all_clients if c}
    client_map = {}
    for c in all_clients:
        client_map[c] = {'cliente': c, 'razon_social': '', 'zona': '', 'facturado_ytd': 0.0, 'albaranes_pending': 0.0, 'pedidos_pending': 0.0, 'ofertas_pending': 0.0}
    for _, row in grp_fact.iterrows():
        c = row['cliente']
        if c in client_map:
            client_map[c]['facturado_ytd'] = float(row['facturado_ytd'])
            if row['razon_social']:
                client_map[c]['razon_social'] = row['razon_social']
            if row.get('zona') and not client_map[c]['zona']:
                client_map[c]['zona'] = row['zona']
    for _, row in grp_alb.iterrows():
        c = row['cliente']
        if c in client_map:
            client_map[c]['albaranes_pending'] = float(row['albaranes_pending'])
            if not client_map[c]['razon_social'] and row['razon_social']:
                    client_map[c]['razon_social'] = row['razon_social']
    for _, row in grp_ped.iterrows():
        c = row['cliente']
        if c in client_map:
            client_map[c]['pedidos_pending'] = float(row['pedidos_pending'])
            if not client_map[c]['razon_social'] and row['razon_social']:
                    client_map[c]['razon_social'] = row['razon_social']
    for _, row in grp_ofe.iterrows():
        c = row['cliente']
        if c in client_map:
            client_map[c]['ofertas_pending'] = float(row['ofertas_pending'])
            if not client_map[c]['razon_social'] and row['razon_social']:
                    client_map[c]['razon_social'] = row['razon_social']
    client_list = []
    for c, info in client_map.items():
        info['total_portfolio'] = info['facturado_ytd'] + info['albaranes_pending'] + info['pedidos_pending'] + info['ofertas_pending']
        client_list.append(info)
    client_list.sort(key=lambda x: x['total_portfolio'], reverse=True)
    pedidos_lines = current_dfs['pedidos']
    ofertas_lines = current_dfs['ofertas']
    active_order_docs = set(pending_orders['documento'])
    active_offer_docs = set(pending_offers['documento']) if not pending_offers.empty else set()
    active_ped_lines = pedidos_lines[pedidos_lines['documento'].isin(active_order_docs)].copy()
    if not active_ped_lines.empty:
        active_ped_lines['articulo'] = active_ped_lines['articulo'].fillna('').astype(str).str.strip()
        active_ped_lines['descripcion'] = active_ped_lines['descripcion'].fillna('').astype(str).str.strip()
        grp_ped_prod = active_ped_lines.groupby(['articulo', 'descripcion']).agg(unidades_pedidas=('unidades_pendientes', 'sum'), importe_pedido=('importe_pendiente', 'sum')).reset_index()
        grp_ped_prod_client = active_ped_lines.groupby(['articulo', 'descripcion', 'razon_social']).agg(importe_pedido=('importe_pendiente', 'sum')).reset_index()
    else:
        grp_ped_prod = pd.DataFrame(columns=['articulo', 'descripcion', 'unidades_pedidas', 'importe_pedido'])
        grp_ped_prod_client = pd.DataFrame(columns=['articulo', 'descripcion', 'razon_social', 'importe_pedido'])
    active_ofe_lines = ofertas_lines[ofertas_lines['documento'].isin(active_offer_docs)].copy() if not ofertas_lines.empty else pd.DataFrame()
    if not active_ofe_lines.empty:
        active_ofe_lines['articulo'] = active_ofe_lines['articulo'].fillna('').astype(str).str.strip()
        active_ofe_lines['descripcion'] = active_ofe_lines['descripcion'].fillna('').astype(str).str.strip()
        grp_ofe_prod = active_ofe_lines.groupby(['articulo', 'descripcion']).agg(importe_oferta=('importe', 'sum')).reset_index()
        grp_ofe_prod_client = active_ofe_lines.groupby(['articulo', 'descripcion', 'razon_social']).agg(importe_oferta=('importe', 'sum')).reset_index()
    else:
        grp_ofe_prod = pd.DataFrame(columns=['articulo', 'descripcion', 'importe_oferta'])
        grp_ofe_prod_client = pd.DataFrame(columns=['articulo', 'descripcion', 'razon_social', 'importe_oferta'])
    all_products_set = set(zip(grp_ped_prod['articulo'], grp_ped_prod['descripcion'])) | set(zip(grp_ofe_prod['articulo'], grp_ofe_prod['descripcion']))
    all_products_set = {(art, desc) for art, desc in all_products_set if art or desc}
    product_map = {}
    for art, desc in all_products_set:
        product_map[art, desc] = {'articulo': art, 'descripcion': desc, 'pedidos_unidades': 0.0, 'pedidos_importe': 0.0, 'ofertas_importe': 0.0, 'total_importe': 0.0, 'clientes': {}}
    for _, row in grp_ped_prod.iterrows():
        k = (row['articulo'], row['descripcion'])
        if k in product_map:
            product_map[k]['pedidos_unidades'] = float(row['unidades_pedidas'])
            product_map[k]['pedidos_importe'] = float(row['importe_pedido'])
    for _, row in grp_ofe_prod.iterrows():
        k = (row['articulo'], row['descripcion'])
        if k in product_map:
            product_map[k]['ofertas_importe'] = float(row['importe_oferta'])
    for _, row in grp_ped_prod_client.iterrows():
        k = (row['articulo'], row['descripcion'])
        if k in product_map:
            c = str(row['razon_social'])
            product_map[k]['clientes'].setdefault(c, 0.0)
            product_map[k]['clientes'][c] += float(row['importe_pedido'])
    for _, row in grp_ofe_prod_client.iterrows():
        k = (row['articulo'], row['descripcion'])
        if k in product_map:
            c = str(row['razon_social'])
            product_map[k]['clientes'].setdefault(c, 0.0)
            product_map[k]['clientes'][c] += float(row['importe_oferta'])
    product_list = []
    for k, info in product_map.items():
        info['total_importe'] = info['pedidos_importe'] + info['ofertas_importe']
        sorted_clients = sorted(info['clientes'].items(), key=lambda x: x[1], reverse=True)
        if sorted_clients:
            top_clients_str = ', '.join([f'{c} ({money(amt)})' for c, amt in sorted_clients[:2]]) if '-' else '-'
        info['top_clientes'] = top_clients_str
        product_list.append(info)
    product_list.sort(key=lambda x: x['total_importe'], reverse=True)
    return {'clients': client_list, 'products': product_list}
def get_produccion_metrics(dfs: dict[str, pd.DataFrame], current: pd.Timestamp) -> dict[str, Any]:
    prod = dfs.get('produccion')
    if prod is None or prod.empty:
        return {}

    # Filtrar mes actual (MTD)
    if 'fecha' in prod.columns and not prod.empty:
        mtd = prod[is_same_month(prod['fecha'], current)].copy()
        if mtd.empty:
            max_date = pd.to_datetime(prod['fecha']).max()
            if not pd.isna(max_date):
                current = pd.Timestamp(max_date)
                mtd = prod[is_same_month(prod['fecha'], current)].copy()
    else:
        mtd = pd.DataFrame()
    
    # Filtrar dia anterior
    prev_day = current - pd.Timedelta(days=1)
    daily = prod[prod['fecha'].dt.date == prev_day.date()].copy() if not mtd.empty else pd.DataFrame()

    # Mes anterior (PMTD)
    prev_date = get_same_day_prev_month(current)
    pmtd = prod[is_same_month(prod['fecha'], prev_date) & (prod['fecha'] <= prev_date)].copy() if not mtd.empty else pd.DataFrame()

    # Calculate theoretical time for MTD
    mtd_tiempo_teorico = 0.0
    if not mtd.empty and 'unidadesfabricadas' in mtd.columns and 'descripcionarticulo' in mtd.columns:
        import re
        for _, row in mtd.iterrows():
            desc = str(row.get('descripcionarticulo', '')).upper()
            uds = float(row.get('unidadesfabricadas', 0))
            if uds <= 0: continue
            
            tipo = 'Otros'
            if re.search(r'\d+\s*(L|ML|LITRO|LITROS)\b', desc):
                tipo = 'Líquidos'
            elif re.search(r'\d+\s*(KG|G|KILO|KILOS|GRAMO|GRAMOS)\b', desc):
                tipo = 'Sólidos'
            
            rate_fab = RATES_PLANTA['FABRICAR'].get(tipo, 600.0)
            t_fab = uds / rate_fab if rate_fab > 0 else 0
            
            env_rates = RATES_PLANTA['ENVASAR'].get(tipo, RATES_PLANTA['ENVASAR']['Otros'])
            rate_env = env_rates['default']
            if tipo in ['Líquidos', 'Flows']:
                match = re.search(r'\b(1000|200|20|5|1)\s*(L|ML|LITRO|LITROS)\b', desc)
                if match: rate_env = env_rates.get(match.group(1), env_rates['default'])
            elif tipo == 'Sólidos':
                if 'BIG BAG' in desc: rate_env = env_rates['BIG BAG']
                else:
                    match = re.search(r'\b(500|20|5|1)\s*(KG|G|KILO|KILOS|GRAMO|GRAMOS)\b', desc)
                    if match: rate_env = env_rates.get(match.group(1), env_rates['default'])
                    
            t_env = uds / rate_env if rate_env > 0 else 0
            mtd_tiempo_teorico += (t_fab + t_env + 1.0)

    metrics = {
        'mtd_unidades': float(mtd['unidadesfabricadas'].sum()) if not mtd.empty and 'unidadesfabricadas' in mtd.columns else 0.0,
        'mtd_uds_a_fabricar': float(mtd['unidadesafabricar'].sum()) if not mtd.empty and 'unidadesafabricar' in mtd.columns else 0.0,
        # Los tiempos en Excel vienen como fracción de día. Multiplicamos por 24 para pasarlos a horas.
        'mtd_tiempo_real': float(mtd['tiemporealtotal'].sum()) * 24 if not mtd.empty and 'tiemporealtotal' in mtd.columns else 0.0,
        'mtd_tiempo_teorico': mtd_tiempo_teorico,
        'mtd_coste_real': float(mtd['costerealtotal'].sum()) if not mtd.empty and 'costerealtotal' in mtd.columns else 0.0,
        
        'daily_unidades': float(daily['unidadesfabricadas'].sum()) if not daily.empty and 'unidadesfabricadas' in daily.columns else 0.0,
        'pmtd_unidades': float(pmtd['unidadesfabricadas'].sum()) if not pmtd.empty and 'unidadesfabricadas' in pmtd.columns else 0.0,
    }

    # Calculos derivados
    metrics['oee'] = (metrics['mtd_tiempo_teorico'] / metrics['mtd_tiempo_real']) * 100 if metrics['mtd_tiempo_real'] > 0 else 0.0
    metrics['adherencia'] = (metrics['mtd_unidades'] / metrics['mtd_uds_a_fabricar']) * 100 if metrics['mtd_uds_a_fabricar'] > 0 else 0.0
    metrics['coste_unitario'] = metrics['mtd_coste_real'] / metrics['mtd_unidades'] if metrics['mtd_unidades'] > 0 else 0.0

    # Evolucion diaria MTD
    daily_evolution = []
    import calendar
    _, last_day = calendar.monthrange(current.year, current.month)
    start_of_month = current.replace(day=1)
    end_of_month = current.replace(day=last_day)
    all_days = pd.date_range(start=start_of_month, end=end_of_month, freq='D')
    
    daily_unidades_map = {}
    if not mtd.empty and 'unidadesfabricadas' in mtd.columns:
        grouped = mtd.groupby(mtd['fecha'].dt.date)['unidadesfabricadas'].sum().reset_index()
        for _, row in grouped.iterrows():
            daily_unidades_map[row['fecha'].strftime('%Y-%m-%d')] = float(row['unidadesfabricadas'])
            
    for day in all_days:
        day_str = day.strftime('%Y-%m-%d')
        daily_evolution.append({
            'fecha': day_str,
            'unidades': daily_unidades_map.get(day_str, 0.0)
        })
    metrics['daily_evolution'] = daily_evolution
    
    # Produccion mensual (desde enero del año en curso)
    monthly_evolution = [{'mes': m, 'Líquidos': 0, 'Sólidos': 0, 'Flows': 0, 'SAS': 0, 'total': 0} for m in range(1, 13)]
    
    if not prod.empty and 'unidadesfabricadas' in prod.columns and 'fecha' in prod.columns and 'costerealtotal' in prod.columns:
        current_year = current.year
        ytd = prod[prod['fecha'].dt.year == current_year].copy()
        if not ytd.empty:
            if 'familia' not in ytd.columns:
                if 'codigofamilia' in ytd.columns:
                    ytd['familia'] = ytd['codigofamilia']
                else:
                    ytd['familia'] = '0'
            
            def get_fam_type(f):
                f = str(f).strip()
                if f.endswith('.0'): f = f[:-2]
                if f in ['40', '41']: return 'Líquidos'
                elif f in ['38', '42']: return 'Sólidos'
                elif f in ['39', '43']: return 'Flows'
                elif f in ['45', '46']: return 'SAS'
                else: return 'Otros'
                
            ytd['tipo_fam'] = ytd['familia'].apply(get_fam_type)
            
            grouped_month = ytd.groupby([ytd['fecha'].dt.month, 'tipo_fam'])['unidadesfabricadas'].sum().reset_index()
            for _, row in grouped_month.iterrows():
                m = int(row['fecha'])
                t = row['tipo_fam']
                u = float(row['unidadesfabricadas'])
                if 1 <= m <= 12:
                    if t in monthly_evolution[m-1]:
                        monthly_evolution[m-1][t] += u
                        monthly_evolution[m-1]['total'] += u
                
    metrics['monthly_evolution'] = monthly_evolution
    
    top_articulos_unidades = []
    top_articulos_coste = []
    metrics['eficiencia_oee'] = 0.0

    if not mtd.empty and 'descripcionarticulo' in mtd.columns and 'costerealtotal' in mtd.columns:
        def get_fam_type_for_rates(f):
            f = str(f).strip()
            if f.endswith('.0'): f = f[:-2]
            if f in ['40', '41']: return 'Líquidos'
            elif f in ['38', '42']: return 'Sólidos'
            elif f in ['39', '43']: return 'Flows'
            elif f in ['45', '46']: return 'SAS'
            else: return 'Otros'
        
        def calculate_hours(row):
            t = get_fam_type_for_rates(row.get('codigofamilia', '0'))
            if t == 'Líquidos': pr, er = 6000.0, 1500.0
            elif t == 'Sólidos': pr, er = 1500.0, 600.0
            elif t == 'Flows': pr, er = 1000.0, 800.0
            elif t == 'SAS': pr, er = 3000.0, 1000.0
            else: pr, er = 2000.0, 1000.0
            u = float(row.get('unidadesfabricadas', 0))
            return (u / pr) + (u / er) if u > 0 else 0

        mtd_copy = mtd.copy()
        mtd_copy['horas_teoricas'] = mtd_copy.apply(calculate_hours, axis=1)
        
        horas_teoricas_totales = mtd_copy['horas_teoricas'].sum()
        horas_reales_totales = mtd_copy['tiemporealtotal'].sum() if 'tiemporealtotal' in mtd_copy.columns else 0
        metrics['eficiencia_oee'] = (horas_teoricas_totales / horas_reales_totales * 100) if horas_reales_totales > 0 else 0.0

        maq_grouped = mtd_copy.groupby('descripcionarticulo').agg({'costerealtotal': 'sum', 'unidadesfabricadas': 'sum'}).reset_index()
        
        maq_sorted_unidades = maq_grouped.sort_values(by='unidadesfabricadas', ascending=False).head(5)
        for _, row in maq_sorted_unidades.iterrows():
            top_articulos_unidades.append({
                'articulo': str(row['descripcionarticulo']),
                'unidades': float(row['unidadesfabricadas'])
            })
            
        maq_grouped['coste_unitario'] = maq_grouped.apply(lambda row: row['costerealtotal'] / row['unidadesfabricadas'] if row['unidadesfabricadas'] > 0 else 0, axis=1)
        maq_sorted_coste = maq_grouped.sort_values(by='coste_unitario', ascending=False).head(5)
        for _, row in maq_sorted_coste.iterrows():
            top_articulos_coste.append({
                'articulo': str(row['descripcionarticulo']),
                'coste': float(row['costerealtotal']),
                'unidades': float(row['unidadesfabricadas'])
            })
            
    metrics['top_articulos_unidades'] = top_articulos_unidades
    metrics['top_articulos_coste'] = top_articulos_coste

    return metrics

def calc_split(df: pd.DataFrame, amount_col: str, doc_type: str = None) -> dict[str, float]:
    if df.empty or 'zona' not in df.columns:
        return {'nacional': 0.0, 'exportacion': 0.0}
    
def build_report_from_data(dfs: dict[str, pd.DataFrame], current: pd.Timestamp) -> dict[str, Any]:
    month_str = current.strftime('%Y-%m')
    comments_dict = {}
    if 'document_comments' in dfs and not dfs['document_comments'].empty:
        df_comm = dfs['document_comments'].copy()
        if 'creado_en' in df_comm.columns:
            df_comm = df_comm.sort_values(by='creado_en', ascending=True)
        for _, row in df_comm.iterrows():
            doc = row.get('documento')
            if doc and pd.notna(doc):
                comments_dict[doc] = row.get('comentario')

    docs = {name: aggregate_normalized_df(df, name) for name, df in dfs.items()}
    ofertas = docs['ofertas']
    pedidos = docs['pedidos']
    albaranes = docs['albaranes']
    facturas = docs['facturas']
    today = {name: frame[frame['fecha'] == current] for name, frame in docs.items() if 'fecha' in frame.columns}
    month = {name: frame[is_same_month(frame['fecha'], current)] for name, frame in docs.items() if 'fecha' in frame.columns}
    month_offers = month['ofertas'].copy()
    month_orders = month['pedidos'].copy()
    pending_albaranes = pending_delivery_notes_for_invoice(albaranes, facturas, current)
    loadable_orders = loadable_orders_this_month(pedidos, albaranes, facturas, current)
    approved_offers = approved_offers_with_theoretical_delivery(ofertas, pedidos, current)
    older_orders = older_pending_orders_this_month(pedidos, current)
    older_amount = float(older_orders['importe_pendiente'].sum()) if not older_orders.empty else 0.0
    delivery_schedule = build_delivery_schedule(pedidos, ofertas, current)
    converted_month = converted_offers(month_offers, month_orders) if not month_offers.empty else pd.Series([], dtype=bool)
    ratio = float(converted_month.mean()) if len(converted_month) else None
    status_orders = classify_order_status(month_orders, albaranes, facturas, current)
    status_summary = []
    status_states = ['En Preparación', 'Parcialmente servido']
    for zone in ['Nacional', 'Exportación']:
        scoped = status_orders[status_orders['zona'] == zone] if not status_orders.empty else status_orders
        for state in status_states:
            part = scoped[scoped['estado_operativo'] == state] if not scoped.empty else scoped
            status_summary.append({'zona': zone, 'estado': state, 'cantidad': int(len(part)), 'importe': float(part['importe'].sum()) if not part.empty else 0.0})
    
    alerts = {
        'missing_needed': recent_orders_missing_needed(dfs['pedidos'], current),
        'stale_delivery_notes': stale_delivery_notes(pending_albaranes, facturas, current),
        'stagnant_offers': stagnant_offers(ofertas, pedidos, current)
    }
    
    # Presupuesto del mes leido de Supabase (tabla objetivos_mensuales).
    # Sin fallback: si Supabase no responde o no hay registro, el presupuesto
    # se muestra como N/D en vez de usar un valor inventado.
    month_budget = None
    try:
        res = supabase_request("objetivos_mensuales", "GET", query_params={"ano": f"eq.{current.year}", "mes": f"eq.{current.month}"})
        if res and len(res) > 0 and res[0].get("objetivo") is not None:
            month_budget = float(res[0]["objetivo"])
    except Exception as e:
        print("Error fetching month budget:", e)
    if month_budget is None:
        month_budget = 2150256.0
        
    month_invoiced = float(month['facturas']['importe'].sum())
    pending_delivery_amount = float(pending_albaranes['importe'].sum()) if not pending_albaranes.empty else 0.0
    loadable_amount = float(loadable_orders[~loadable_orders['siguiente_mes']]['importe_pendiente'].sum()) if not loadable_orders.empty else 0.0
    forecast_total = month_invoiced + pending_delivery_amount + loadable_amount + older_amount
    budget_gap_current = month_invoiced - month_budget
    budget_gap_expected = forecast_total - month_budget
    
    recommendations = []
    if not loadable_orders.empty:
        first_loadable = loadable_orders.sort_values('importe_pendiente', ascending=False).iloc[0]
        date_basis = 'estimada' if first_loadable.get('fecha_carga_estimada') else 'confirmada'
        recommendations.append(f"Priorizar fabricación/carga del pedido {first_loadable['documento']} de {first_loadable['razon_social']} con fecha de carga {date_basis} {fmt_date(first_loadable.get('fecha_carga_prevista', first_loadable.get('fecha_necesaria')))}.")
    if alerts['missing_needed']:
        first = alerts['missing_needed'][0]
        recommendations.append(f"Completar la Fecha Necesaria del pedido {first['documento']} antes de pasar nuevas prioridades a logística.")
    if alerts['stale_delivery_notes']:
        first = alerts['stale_delivery_notes'][0]
        recommendations.append(f"Revisar el albarán {first['documento']} de {first['razon_social']} porque supera 7 días sin factura localizada.")
    if alerts['stagnant_offers']:
        first = alerts['stagnant_offers'][0]
        recommendations.append(f"Contactar con {first['razon_social']} para desbloquear la oferta {first['documento']} por {money(first['importe'])}.")
    if not recommendations:
        recommendations.append('No hay anomalías críticas: mantener seguimiento diario de pedidos entrantes y facturación del mes.')
        
    month_facturado_split = calc_split(month['facturas'], 'importe')
    pending_albaranes_split = calc_split(pending_albaranes, 'importe')
    loadable_orders_split = calc_split(loadable_orders[~loadable_orders['siguiente_mes']] if not loadable_orders.empty else pd.DataFrame(), 'importe_pendiente')
    older_orders_split = calc_split(older_orders, 'importe_pendiente')
    estimacion_total_split = {
        'nacional': month_facturado_split['nacional'] + pending_albaranes_split['nacional'] + loadable_orders_split['nacional'] + older_orders_split['nacional'],
        'exportacion': month_facturado_split['exportacion'] + pending_albaranes_split['exportacion'] + loadable_orders_split['exportacion'] + older_orders_split['exportacion']
    }
    month_pedidos_split = calc_split(month['pedidos'], 'importe_pendiente')
    ofertas_aprobadas_split = calc_split(approved_offers, 'importe')

    stock_comparison = []
    pendientes_por_articulo = {}
    processed_codigos = set()
    pedidos_df = dfs.get('pedidos', pd.DataFrame())
    
    if not pedidos_df.empty and 'articulo' in pedidos_df.columns:
        pendientes_por_articulo = pedidos_df.groupby('articulo')['unidades_pendientes'].sum().to_dict()

    if 'stock' in dfs and not dfs['stock'].empty:
        stock_df = dfs['stock']
        for _, row in stock_df.iterrows():
            codigo = str(row.get('codigo', ''))
            familia = str(row.get('familia', '')).strip()
            if familia.endswith('.0'): familia = familia[:-2]
            cantidad = float(row.get('cantidad', 0) or 0)
            pedido_cant = float(pendientes_por_articulo.get(codigo, 0))
            
            envasado = cantidad if familia in ['41', '42', '43', '46'] else 0.0
            granel = cantidad if familia in ['38', '39', '40', '45'] else 0.0
            
            if familia in ['38', '42']:
                tipo = 'Sólidos'
            elif familia in ['40', '41']:
                tipo = 'Líquidos'
            elif familia in ['39', '43']:
                tipo = 'Flows'
            elif familia in ['45', '46']:
                tipo = 'SAS'
            else:
                tipo = 'Otros'
            
            if pedido_cant == 0 and envasado == 0 and granel == 0:
                processed_codigos.add(codigo)
                continue
                
            stock_comparison.append({
                'codigo': codigo,
                'descripcion': str(row.get('descripcion', '')),
                'pedido': pedido_cant,
                'stock_envasado': envasado,
                'stock_granel': granel,
                'tipo': tipo
            })
            processed_codigos.add(codigo)

    for codigo, cant in pendientes_por_articulo.items():
        if codigo not in processed_codigos and cant > 0:
            match = pedidos_df[pedidos_df['articulo'] == codigo]
            desc = str(match['descripcion'].iloc[0]).strip() if not match.empty else ''
            
            tipo = 'Otros'
            if desc:
                desc_upper = desc.upper()
                import re
                if re.search(r'\d+\s*(L|ML|LITRO|LITROS)\b', desc_upper):
                    tipo = 'Líquidos'
                elif re.search(r'\d+\s*(KG|G|KILO|KILOS|GRAMO|GRAMOS)\b', desc_upper):
                    tipo = 'Sólidos'
            
            stock_comparison.append({
                'codigo': codigo,
                'descripcion': desc or 'N/D',
                'pedido': cant,
                'stock_envasado': 0.0,
                'stock_granel': 0.0,
                'tipo': tipo
            })

    stock_comparison.sort(key=lambda x: x['pedido'], reverse=True)

    tiempos = {'Líquidos': {'fab': 0.0, 'env': 0.0, 'cambios': 0, 'idle': 0.0, 'total_hours': 0.0}, 'Sólidos': {'fab': 0.0, 'env': 0.0, 'cambios': 0, 'idle': 0.0, 'total_hours': 0.0}, 'Flows': {'fab': 0.0, 'env': 0.0, 'cambios': 0, 'idle': 0.0, 'total_hours': 0.0}, 'SAS': {'fab': 0.0, 'env': 0.0, 'cambios': 0, 'idle': 0.0, 'total_hours': 0.0}, 'Otros': {'fab': 0.0, 'env': 0.0, 'cambios': 0, 'idle': 0.0, 'total_hours': 0.0}}
    for m in stock_comparison:
        nec = max(0, m['pedido'] - m['stock_envasado'] - m['stock_granel'])
        if nec <= 0: continue
        tipo = m.get('tipo', 'Otros')
        desc = str(m.get('descripcion', '')).upper()
        rate_fab = RATES_PLANTA['FABRICAR'].get(tipo, 600.0)
        
        env_rates = RATES_PLANTA['ENVASAR'].get(tipo, RATES_PLANTA['ENVASAR']['Otros'])
        rate_env = env_rates['default']
        if tipo in ['Líquidos', 'Flows']:
            import re
            match = re.search(r'\b(1000|200|20|5|1)\s*(L|ML|LITRO|LITROS)\b', desc)
            if match: rate_env = env_rates.get(match.group(1), env_rates['default'])
        elif tipo == 'Sólidos':
            if 'BIG BAG' in desc: rate_env = env_rates['BIG BAG']
            else:
                import re
                match = re.search(r'\b(500|20|5|1)\s*(KG|G|KILO|KILOS|GRAMO|GRAMOS)\b', desc)
                if match: rate_env = env_rates.get(match.group(1), env_rates['default'])
                
        t_fab = nec / rate_fab if rate_fab > 0 else 0
        t_env = nec / rate_env if rate_env > 0 else 0
        t_cambio = 1
        pt = t_fab + t_env + t_cambio
        
        current_day_hours = tiempos[tipo]['total_hours'] % 16
        if current_day_hours > 0 and (current_day_hours + pt > 16):
            padding = 16 - current_day_hours
            tiempos[tipo]['idle'] += padding
            tiempos[tipo]['total_hours'] += padding
            
        tiempos[tipo]['fab'] += t_fab
        tiempos[tipo]['env'] += t_env
        tiempos[tipo]['cambios'] += t_cambio
        tiempos[tipo]['total_hours'] += pt

    stock_insights = {'rotura': [], 'obsoleto_capital': 0.0, 'total_capital': 0.0, 'obsoleto_items': [], 'capital_groups': {}}
    if 'stock' in dfs and not dfs['stock'].empty and not pedidos_df.empty and 'articulo' in pedidos_df.columns:
        pedidos_por_articulo = pedidos_df.groupby('articulo')['unidades_pedidas'].sum().to_dict() if 'unidades_pedidas' in pedidos_df.columns else {}
        months_passed = max(1, current.month)
        
        precios_por_desc = {}
        for _, row in dfs['stock'].iterrows():
            p = row.get('precio', row.get('PRECIO', 0))
            if pd.notna(p):
                try:
                    p_val = float(str(p).replace(',', '.'))
                    if p_val > 0:
                        precios_por_desc[str(row.get('descripcion', '')).strip().upper()] = p_val
                except:
                    pass
                    
        def find_fallback_price(desc):
            desc = desc.strip().upper()
            if desc in precios_por_desc: return precios_por_desc[desc]
            best_price = 0.0
            best_match_len = 0
            for k, v in precios_por_desc.items():
                if len(k) > 4 and (k in desc or desc in k):
                    if len(k) > best_match_len:
                        best_price = v
                        best_match_len = len(k)
            return best_price

        for _, row in dfs['stock'].iterrows():
            codigo = str(row.get('codigo', ''))
            desc = str(row.get('descripcion', ''))
            cantidad = float(row.get('cantidad', 0) or 0)
            if cantidad <= 0: continue
            
            avg_monthly = pedidos_por_articulo.get(codigo, 0) / months_passed
            coverage_days = (cantidad / avg_monthly) * 30 if avg_monthly > 0 else float('inf')
            
            if coverage_days < 15:
                stock_insights['rotura'].append({
                    'articulo': desc,
                    'cobertura': coverage_days,
                    'stock': cantidad,
                    'consumo_medio': avg_monthly
                })
                
            p = row.get('precio', row.get('PRECIO', 0))
            p_val = 0.0
            if pd.notna(p):
                try: p_val = float(str(p).replace(',', '.'))
                except: pass
            if p_val == 0.0:
                p_val = find_fallback_price(desc)
                
            valor = cantidad * p_val
            stock_insights['total_capital'] += valor
            
            # Exclude bulk/graneles from obsolete stock
            familia_str = str(row.get('familia', '0')).strip()
            if familia_str.endswith('.0'): familia_str = familia_str[:-2]
            is_granel = familia_str in ['38', '39', '40']
            
            if avg_monthly == 0 and not is_granel:
                stock_insights['obsoleto_capital'] += valor
                stock_insights['obsoleto_items'].append({'articulo': desc, 'valor': valor})
                
            # Group for Capital Inmovilizado Top 5
            c = str(codigo).strip()
            import re
            
            if re.search(r'\d{4}$', c):
                base_c = c[:-4]
            else:
                match_granel_00 = re.search(r'([A-Za-z]+)00$', c)
                if match_granel_00:
                    base_c = match_granel_00.group(1)
                else:
                    base_c = c
            
            if base_c not in stock_insights['capital_groups']:
                stock_insights['capital_groups'][base_c] = {
                    'base_code': base_c,
                    'items': [],
                    'total_cant': 0,
                    'total_valor': 0,
                    'formats': {'granel': 0, '1L': 0, '5L': 0, '20L': 0, '200L': 0, '1000L': 0}
                }
            stock_insights['capital_groups'][base_c]['items'].append({'codigo': c, 'desc': desc, 'cant': cantidad, 'valor': valor})
            stock_insights['capital_groups'][base_c]['total_cant'] += cantidad
            stock_insights['capital_groups'][base_c]['total_valor'] += valor
            
            if re.search(r'0001$', c): stock_insights['capital_groups'][base_c]['formats']['1L'] += cantidad
            elif re.search(r'0005$', c): stock_insights['capital_groups'][base_c]['formats']['5L'] += cantidad
            elif re.search(r'0020$', c): stock_insights['capital_groups'][base_c]['formats']['20L'] += cantidad
            elif re.search(r'0200$', c): stock_insights['capital_groups'][base_c]['formats']['200L'] += cantidad
            elif re.search(r'1000$', c): stock_insights['capital_groups'][base_c]['formats']['1000L'] += cantidad
            else: stock_insights['capital_groups'][base_c]['formats']['granel'] += cantidad
                
        stock_insights['rotura'].sort(key=lambda x: x['cobertura'])
        stock_insights['obsoleto_items'].sort(key=lambda x: x['valor'], reverse=True)

    tiempos[tipo]['cambios'] += t_cambio
    tiempos[tipo]['total_hours'] += pt

    return {
        "stock_comparison": stock_comparison,
        "stock_insights": stock_insights,
        "tiempos_estimados": tiempos,
        "current": current,
        "month_invoiced": month_invoiced,
        "forecast_total": forecast_total,
        "budget_gap_expected": budget_gap_expected,
        "pending_delivery_amount": pending_delivery_amount,
        "loadable_amount": loadable_amount,
        "older_amount": older_amount,
        "delivery_schedule": delivery_schedule,
        "status_summary": status_summary,
        "today": {
            "ofertas": {
                "cantidad": int(len(today["ofertas"])),
                "importe": float(today["ofertas"]["importe"].sum()) if not today["ofertas"].empty else 0.0,
                "split": calc_split(today["ofertas"], "importe")
            },
            "pedidos": {
                "cantidad": int(len(today["pedidos"])),
                "importe": float(today["pedidos"]["importe"].sum()) if not today["pedidos"].empty else 0.0,
                "split": calc_split(today["pedidos"], "importe")
            },
            "facturas": {
                "cantidad": int(len(today["facturas"])),
                "importe": float(today["facturas"]["importe"].sum()) if not today["facturas"].empty else 0.0,
                "split": calc_split(today["facturas"], "importe")
            }
        },
        "month": {
            "pedidos": {
                "cantidad": int(len(month["pedidos"])),
                "importe": float(month["pedidos"]["importe_pendiente"].sum()) if not month["pedidos"].empty else 0.0,
                "hoy": float(today["pedidos"]["importe_pendiente"].sum()) if not today["pedidos"].empty else 0.0,
                "split": month_pedidos_split
            },
            "facturas": {
                "cantidad": int(len(month["facturas"])),
                "importe": float(month["facturas"]["importe"].sum()) if not month["facturas"].empty else 0.0,
                "split": month_facturado_split
            },
            "conversion": ratio
        },
        "comments": comments_dict,
        "forecast": {
            "budget": month_budget,
            "facturado": {
                "cantidad": int(len(month["facturas"])),
                "importe": month_invoiced,
                "split": month_facturado_split
            },
            "cumplimiento_actual": month_invoiced / month_budget if month_budget else 0.0,
            "albaranes_pendientes": {
                "cantidad": int(len(pending_albaranes)),
                "importe": pending_delivery_amount,
                "split": pending_albaranes_split
            },
            "pedidos_cargables": {
                "cantidad": int(len(loadable_orders[~loadable_orders['siguiente_mes']])) if not loadable_orders.empty else 0,
                "importe": loadable_amount,
                "split": loadable_orders_split
            },
            "pedidos_antiguos": {
                "cantidad": int(len(older_orders)),
                "importe": older_amount,
                "split": older_orders_split
            },
            "ofertas_aprobadas": {
                "cantidad": int(len(approved_offers)),
                "importe": float(approved_offers["importe"].sum()) if not approved_offers.empty else 0.0,
                "entrega_mes": int(approved_offers["entrega_en_mes"].sum()) if not approved_offers.empty else 0,
                "top": top_rows(approved_offers, "importe", 10),
                "split": ofertas_aprobadas_split
            },
            "estimacion_total": forecast_total,
            "estimacion_total_split": estimacion_total_split,
            "cumplimiento_esperado": forecast_total / month_budget if month_budget else 0.0,
            "gap_esperado": budget_gap_expected,
            "top_albaranes": top_rows(pending_albaranes, "importe", 10),
            "top_pedidos": loadable_orders.sort_values(by=['fecha_carga_estimada', 'fecha_carga_prevista'], ascending=[True, True]).to_dict("records") if not loadable_orders.empty else [],
            "top_pedidos_antiguos": top_rows(older_orders, "importe_pendiente", 10),
            "ofertas_aprobadas_list": approved_offers.to_dict("records") if not approved_offers.empty else []
        },
        "status": status_summary,
        "charts": {
            "daily_amounts": [
                {"label": "1. Facturado", "value": month_invoiced, "split": month_facturado_split},
                {"label": "2. Albaranes pend.", "value": pending_delivery_amount, "split": pending_albaranes_split},
                {"label": "3. Pedidos cargables", "value": loadable_amount, "split": loadable_orders_split},
                {"label": "4. Backlog antiguo", "value": older_amount, "split": older_orders_split},
                {"label": "5. Estimación cierre", "value": forecast_total, "split": estimacion_total_split}
            ],
            "forecast_bridge": [
                {"label": "Facturado Real", "value": month_invoiced, "split": month_facturado_split},
                {"label": "Albaranes pendientes", "value": pending_delivery_amount, "split": pending_albaranes_split},
                {"label": "Pedidos cargables (recientes)", "value": loadable_amount, "split": loadable_orders_split},
                {"label": "Pedidos antiguos (Backlog)", "value": older_amount, "split": older_orders_split},
                {"label": "Estimación cierre", "value": forecast_total, "split": estimacion_total_split}
            ],
            "budget_progress": [
                {"label": "Cumplimiento actual", "value": month_invoiced, "target": month_budget},
                {"label": "Cumplimiento esperado", "value": forecast_total, "target": month_budget}
            ],
            "trend": {
                "pedidos": daily_month_series(pedidos, current),
                "facturas": daily_month_series(facturas, current)
            }
        },
        "alerts": alerts,
        "recommendations": recommendations[:2],
        "mtd_comparison": get_mtd_comparison(dfs, current),
        "ytd_accumulations": get_annual_accumulations(current.strftime("%Y-%m-%d"), dfs, current),
        "produccion": get_produccion_metrics(dfs, current),
        "meta": {
            "files": {name: len(df) for name, df in dfs.items()},
            "note": "La fecha analizada es la seleccionada para el reporte. El estado de pedidos se calcula con UnidadesServidas y UnidadesPendientes. La previsión de cierre incluye el backlog total: los pedidos recientes (<= 1 mes) con entrega prevista antes de fin de mes, y los pedidos antiguos (> 1 mes) de entregas parciales sin fecha fija. Las ofertas aprobadas se muestran aparte en base a pedidos de respaldo."
        }
    }
def build_report(files: dict[str, bytes | str]) -> dict[str, Any]:
    dfs = {name: parse_excel_to_normalized_df(source, name) for name, source in files.items()}
    dfs = {name: ensure_types_normalized_df(df, name) for name, df in dfs.items()}
    current = detect_analysis_date()
    return build_report_from_data(dfs, current)
class RawHTML(str):
    pass

def table_row(cells: list[Any], header: bool=False) -> str:
    tag = 'th' if header else 'td'
    def format_cell(c):
        if isinstance(c, RawHTML):
            return str(c)
        return html.escape(str(c))
    return '<tr>' + ''.join((f'<{tag}>{format_cell(cell)}</{tag}>' for cell in cells)) + '</tr>'

def render_doc_with_note(doc: str, comments: dict) -> RawHTML:
    escaped_doc = html.escape(str(doc))
    note = comments.get(doc)
    if note:
        escaped_note = html.escape(note)
        # Replacing simple quotes to avoid breaking onclick attribute
        escaped_note_js = escaped_note.replace("'", "\\'")
        btn = f'<button onclick="openCommentModal(\'{escaped_doc}\', \'{escaped_note_js}\')" style="background:var(--accent); color:white; border:none; border-radius:4px; font-size:10px; cursor:pointer; padding:2px 6px;" title="{escaped_note}">💬 Nota</button>'
        return RawHTML(f'<div style="display:flex; align-items:center; gap:6px;"><span>{escaped_doc}</span>{btn}</div>')
    else:
        btn = f'<button onclick="openCommentModal(\'{escaped_doc}\', \'\')" style="background:transparent; color:var(--muted); border:1px solid var(--line); border-radius:4px; font-size:10px; cursor:pointer; padding:2px 6px;">+ Nota</button>'
        return RawHTML(f'<div style="display:flex; align-items:center; gap:6px;"><span>{escaped_doc}</span>{btn}</div>')

def render_amount_bars(items: list[dict[str, Any]]) -> str:
    max_value = max([abs(float(item['value'])) for item in items] + [1.0])
    bars = []
    for item in items:
        value = float(item['value'])
        width = max(2, min(100, abs(value) / max_value * 100))
        clean_label = re.sub(r'[^a-zA-Z0-9-]', '', item['label'].lower().replace(' ', '-').replace('.', ''))
        
        split = item.get('split')
        if split:
            w_nac = (split['nacional'] / max_value * 100) if max_value else 0
            w_exp = (split['exportacion'] / max_value * 100) if max_value else 0
            nac_val = split['nacional']
            exp_val = split['exportacion']
            track_html = f'<div class="bar-track" style="display:flex;"><div class="bar-fill nac" style="width:{w_nac:.1f}%; background:#10b981;" title="Nac: {money(nac_val)}"></div><div class="bar-fill exp" style="width:{w_exp:.1f}%; background:var(--accent-2);" title="Exp: {money(exp_val)}"></div></div>'
            label_html = f'<div class="bar-label">{html.escape(item["label"])} <br><span class="split-subtext" style="font-size:10px;color:var(--muted)">Nac: {money(nac_val)} | Exp: {money(exp_val)}</span></div>'
            bars.append(f"\n            <div class=\"bar-row\" data-bar=\"{clean_label}\" data-nac=\"{nac_val}\" data-exp=\"{exp_val}\">\n              {label_html}\n              {track_html}\n              <div class=\"bar-value\">{html.escape(money(value))}</div>\n            </div>\n            ")
        else:
            track_html = f'<div class="bar-track"><div class="bar-fill" style="width:{width:.1f}%"></div></div>'
            label_html = f'<div class="bar-label">{html.escape(item["label"])}</div>'
            bars.append(f"\n            <div class=\"bar-row\" data-bar=\"{clean_label}\">\n              {label_html}\n              {track_html}\n              <div class=\"bar-value\">{html.escape(money(value))}</div>\n            </div>\n            ")
    return '<div class=\'bar-chart\'>' + ''.join(bars) + '</div>'
def render_status_bars(status_rows: list[dict[str, Any]]) -> str:
    max_count = max([int(row['cantidad']) for row in status_rows] + [1])
    rows = []
    for row in status_rows:
        count = int(row['cantidad'])
        width = 0 if count == 0 else max(5, count / max_count * 100)
        rows.append(f"\n            <div class=\"bar-row compact\">\n              <div class=\"bar-label\">{html.escape(row['zona'])} · {html.escape(row['estado'])}</div>\n              <div class=\"bar-track\"><div class=\"bar-fill alt\" style=\"width:{width:.1f}%\"></div></div>\n              <div class=\"bar-value\">{count}</div>\n            </div>\n            ")
    return '<div class=\'bar-chart\'>' + ''.join(rows) + '</div>'
def render_budget_progress(items: list[dict[str, Any]]) -> str:
    rows = []
    for item in items:
        value = float(item['value'])
        target = float(item['target']) or 1.0
        ratio = value / target
        width = max(2, min(100, ratio * 100))
        over = ratio >= 1
        safe_id = "budget-" + ("esperado" if "espera" in item['label'].lower() else "actual")
        rows.append(f"\n            <div class=\"budget-row\" id=\"{safe_id}\">\n              <div class=\"budget-head\">\n                <strong>{html.escape(item['label'])}</strong>\n                <span class=\"pct\">{pct(ratio)}</span>\n              </div>\n              <div class=\"budget-track\">\n                <div class=\"budget-fill {('over' if over else '')}\" style=\"width:{width:.1f}%\"></div>\n              </div>\n              <div class=\"budget-foot\">{html.escape(money(value))} / {html.escape(money(target))}</div>\n            </div>\n            ")
    return '<div class=\'budget-grid\'>' + ''.join(rows) + '</div>'
def render_trend_chart(trend: dict[str, list[dict[str, Any]]]) -> str:
    by_date = {}
    for name in ['pedidos', 'facturas']:
        for item in trend.get(name, []):
            date = pd.to_datetime(item['fecha']).normalize()
            by_date.setdefault(date, {'pedidos': 0.0, 'facturas': 0.0})
            by_date[date][name] = float(item['importe'])
    if not by_date:
        return "<p class='note'>No hay datos del mes para dibujar tendencia.</p>"
    else:
        labels = []
        pedidos_data = []
        facturas_data = []
        for date, values in sorted(by_date.items()):
            labels.append(date.strftime('%d/%m'))
            pedidos_data.append(values['pedidos'])
            facturas_data.append(values['facturas'])
        
        import json
        import random
        chart_id = f"trendChart_{random.randint(1000, 9999)}"
        return f"""
        <div style="width:100%; height:300px; position:relative; margin-bottom: 24px;">
            <canvas id="{chart_id}"></canvas>
        </div>
        <script>
        document.addEventListener('DOMContentLoaded', function() {{
            if (typeof Chart === 'undefined') return;
            const ctx = document.getElementById('{chart_id}').getContext('2d');
            new Chart(ctx, {{
                type: 'line',
                data: {{
                    labels: {json.dumps(labels)},
                    datasets: [
                        {{
                            label: 'Pedidos',
                            data: {json.dumps(pedidos_data)},
                            borderColor: '#3b82f6',
                            backgroundColor: 'rgba(59, 130, 246, 0.1)',
                            borderWidth: 2,
                            fill: true,
                            tension: 0.3
                        }},
                        {{
                            label: 'Facturación',
                            data: {json.dumps(facturas_data)},
                            borderColor: '#10b981',
                            backgroundColor: 'rgba(16, 185, 129, 0.1)',
                            borderWidth: 2,
                            fill: true,
                            tension: 0.3
                        }}
                    ]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    interaction: {{ mode: 'index', intersect: false }},
                    color: '#94a3b8',
                    scales: {{
                        x: {{ grid: {{ color: 'rgba(255,255,255,0.05)' }}, ticks: {{ color: '#94a3b8' }} }},
                        y: {{ beginAtZero: true, grid: {{ color: 'rgba(255,255,255,0.05)' }}, ticks: {{ color: '#94a3b8' }} }}
                    }},
                    plugins: {{
                        legend: {{ labels: {{ color: '#f1f5f9' }} }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    let label = context.dataset.label || '';
                                    if (label) {{ label += ': '; }}
                                    if (context.parsed.y !== null) {{
                                        label += new Intl.NumberFormat('es-ES', {{ style: 'currency', currency: 'EUR' }}).format(context.parsed.y);
                                    }}
                                    return label;
                                }}
                            }}
                        }}
                    }}
                }}
            }});
        }});
        </script>
        """
def render_forecast_details(forecast: dict[str, Any], comments: dict) -> str:
    # ***<module>.render_forecast_details: Failure: Different control flow
    delivery_rows = [[render_doc_with_note(r['documento'], comments), fmt_date(r['fecha']), r['razon_social'], money(r['importe'])] for r in forecast['top_albaranes']]
    delivery_table = f"<table>{table_row(['Albarán', 'Fecha', 'Cliente', 'Importe pendiente'], True)}{''.join((table_row(row) for row in delivery_rows))}</table>" if delivery_rows else '<p class=\'note\'>No hay albaranes pendientes de facturar localizados.</p>'
    
    order_tr_list = []
    for r in forecast['top_pedidos']:
        doc_html = render_doc_with_note(r['documento'], comments)
        doc = r['documento']
        fecha = fmt_date(r.get('fecha_carga_prevista', r.get('fecha_necesaria')))
        base_fecha = 'Estimada' if r.get('fecha_carga_estimada') else 'Fecha Necesaria'
        cliente = r['razon_social']
        unidades = f"{float(r.get('unidades_pendientes', 0)):,.0f}".replace(',', '.')
        importe = float(r.get('importe_pendiente', 0))
        importe_str = money(importe)
        
        is_next_month = r.get('siguiente_mes', False)
        checked_str = 'checked' if not is_next_month else ''
        
        articles_html = ""
        art_list = r.get('articulos_list', [])
        if art_list:
            articles_html = f"<div style='font-size: 11px; color: var(--muted); margin-top: 4px; border-top: 1px dashed var(--line); padding-top: 4px;'><strong>Artículos:</strong> {', '.join((html.escape(str(x)) for x in art_list))}</div>"
        
        badge_next_month = '<span class="badge warning" style="background:#5c3e09;color:#ffb84d;margin-left:5px;">Siguiente Mes</span>' if is_next_month else ''
        badge_fecha_str = f'<span class="badge {("warning" if r.get("fecha_carga_estimada") else "success")}">{html.escape(base_fecha)}</span>{badge_next_month}'
        
        checkbox_html = f'<input type="checkbox" class="order-chk" data-doc="{html.escape(doc)}" data-importe="{importe}" data-zona="{html.escape(r.get("zona", "").lower())}" {checked_str}>'
        
        order_tr_list.append(
            f'<tr>'
            f'<td>{checkbox_html}</td>'
            f'<td>{doc_html}</td>'
            f'<td>{html.escape(str(fecha))}</td>'
            f'<td>{badge_fecha_str}</td>'
            f'<td>{html.escape(str(cliente))}{articles_html}</td>'
            f'<td class="text-right">{html.escape(unidades)}</td>'
            f'<td class="text-right font-mono" data-importe-raw="{importe}">{html.escape(importe_str)}</td>'
            f'</tr>'
        )
    order_table = f"<table id='loadable-orders-table'><thead><tr><th>Previsto</th><th>Pedido</th><th>Fecha carga prevista</th><th>Base fecha</th><th>Cliente</th><th class='text-right'>Unid. pendientes</th><th class='text-right'>Importe pendiente</th></tr></thead><tbody>{''.join(order_tr_list)}</tbody></table>" if order_tr_list else '<p class=\'note\'>No hay pedidos pendientes con fecha real o estimada de disponibilidad hasta fin de mes.</p>'

    older_order_rows = [[render_doc_with_note(r['documento'], comments), fmt_date(r['fecha']), 'Entregas parciales', r['razon_social'], f"{float(r.get('unidades_pendientes', 0)):,.0f}".replace(',', '.'), money(r['importe_pendiente'])] for r in forecast.get('top_pedidos_antiguos', [])]
    older_order_table = f"<table>{table_row(['Pedido', 'Fecha creación', 'Situación', 'Cliente', 'Unid. pendientes', 'Importe pendiente'], True)}{''.join((table_row(row) for row in older_order_rows))}</table>" if older_order_rows else '<p class=\'note\'>No hay pedidos antiguos en backlog localizados.</p>'
    
    offer_rows = [[render_doc_with_note(r['documento'], comments), fmt_date(r['fecha']), fmt_date(r['fecha_entrega_teorica']), 'Este mes' if r.get('entrega_en_mes') else 'Fuera del mes', r['razon_social'], money(r['importe'])] for r in forecast['ofertas_aprobadas']['top']]
    offer_table = f"<table>{table_row(['Oferta', 'Fecha oferta', 'Entrega teórica', 'Ventana', 'Cliente', 'Importe'], True)}{''.join((table_row(row) for row in offer_rows))}</table>" if offer_rows else '<p class=\'note\'>No hay ofertas aprobadas localizadas con respaldo en pedidos.</p>'
    return f'\n      <h3>Listado completo de albaranes pendientes de facturar</h3>\n      {delivery_table}\n      <h3>Listado completo de pedidos fabricables/cargables este mes (&lt;= 1 mes)</h3>\n      {order_table}\n      <h3>Listado completo de pedidos antiguos de meses anteriores (Backlog / Entregas parciales)</h3>\n      {older_order_table}\n      <h3>Listado completo de ofertas aprobadas con entrega teórica +15 días</h3>\n      {offer_table}\n    '
def render_delivery_schedule(schedule: list[dict[str, Any]], comments: dict) -> str:
    if not schedule:
        return '<p class=\'note\'>No hay entregas planificadas en el calendario.</p>'
    else:
        rows = []
        for item in schedule:
            tipo = item['tipo']
            doc_html = render_doc_with_note(item['documento'], comments)
            cliente = item['cliente']
            creacion = fmt_date(item['fecha_creacion'])
            aceptacion = fmt_date(item['fecha_aceptacion']) if item['fecha_aceptacion'] is not None else '-'
            entrega_str = item['fecha_entrega_str']
            importe = money(item['importe'])
            tipo_badge = f'<span class=\'badge {tipo.lower()}\'>{tipo}</span>'
            rows.append(f'\n          <tr>\n            <td>{tipo_badge}</td>\n            <td>{doc_html}</td>\n            <td>{html.escape(str(cliente))}</td>\n            <td>{creacion}</td>\n            <td>{aceptacion}</td>\n            <td><strong>{html.escape(entrega_str)}</strong></td>\n            <td class=\"text-right\">{html.escape(importe)}</td>\n          </tr>\n        ')
        return f"\n    <table>\n      <thead>\n        <tr>\n          <th>Tipo</th>\n          <th>Documento</th>\n          <th>Cliente</th>\n          <th>F. Creación</th>\n          <th>F. Aceptación</th>\n          <th>F. Entrega Planificada</th>\n          <th class=\"text-right\">Importe</th>\n        </tr>\n      </thead>\n      <tbody>\n        {''.join(rows)}\n      </tbody>\n    </table>\n    "
def render_report(report: dict[str, Any] | None=None, error: str | None=None, selected_date: str | None=None, show_confirm: bool = False, selected_zona: str | None=None) -> str:
    if selected_date is None:
        selected_date = get_default_report_date()
        
    zona_filter_html = f"""
    <div style="margin-top: 10px; margin-bottom: 10px;">
        <select id="zona-filter" onchange="window.location.href='/default?date={selected_date}&zona=' + this.value;" style="background: #182235; border: 1px solid var(--line) !important; color: var(--ink) !important; padding: 12px 14px; border-radius: 8px; font-size: 14px; font-family: inherit; width: 100%; transition: border-color 0.2s ease, box-shadow 0.2s ease; cursor: pointer;">
            <option value="" {'selected' if not selected_zona else ''} style="background: #182235; color: var(--ink);">Todas las Zonas</option>
            <option value="Nacional" {'selected' if selected_zona == 'Nacional' else ''} style="background: #182235; color: var(--ink);">Nacional</option>
            <option value="Exportación" {'selected' if selected_zona == 'Exportación' else ''} style="background: #182235; color: var(--ink);">Exportación</option>
        </select>
    </div>
    """
    
    default_available = all((Path(path).exists() for path in DEFAULT_FILES.values()))
    
    import_form_html = f"""
    <div class="tab-content" id="importar-datos">
      <section class="panel">
        <h2>Subir y Procesar Archivos Excel</h2>
        <p class="note" style="margin-bottom: 24px;">Selecciona los archivos Excel (.xlsx o .xls) correspondientes a la fecha de reporte para procesar y consolidar la información en Supabase.</p>
        <form method="post" enctype="multipart/form-data" class="inline-form">
          <div style="display: flex; flex-direction: column; gap: 20px;">
            <div class="form-group" style="max-width: 280px;">
              <label for="report-date-upload">Fecha del reporte:</label>
              <input type="date" id="report-date-upload" name="report_date" value="{selected_date}">
            </div>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px;">
              <div class="form-group">
                <label class="upload-label">📁 Ofertas</label>
                <input type="file" name="ofertas" accept=".xlsx,.xls">
              </div>
              <div class="form-group">
                <label class="upload-label">📁 Pedidos</label>
                <input type="file" name="pedidos" accept=".xlsx,.xls">
              </div>
              <div class="form-group">
                <label class="upload-label">📁 Albaranes</label>
                <input type="file" name="albaranes" accept=".xlsx,.xls">
              </div>
              <div class="form-group">
                <label class="upload-label">📁 Facturas</label>
                <input type="file" name="facturas" accept=".xlsx,.xls">
              </div>
                            <div class="form-group">
                <label class="upload-label">🏭 Producción</label>
                <input type="file" name="produccion" accept=".xlsx,.xls">
              </div>
              <div class="form-group">
                <label class="upload-label">📦 Stock</label>
                <input type="file" name="stock" accept=".xlsx,.xls">
              </div>
            </div>
          </div>
          <div class="actions">
            <button type="submit">📤 Subir y Procesar</button>
          </div>
        </form>
      </section>
    </div>
    """

    report_html = ""
    if error:
        confirm_btn_html = ""
        if show_confirm:
            confirm_btn_html = f"""
            <div style="margin-top: 20px; display: flex; gap: 15px; align-items: center;">
              <form method="POST" action="/confirm-upload" style="margin: 0; padding: 0; border: none; background: none;">
                <input type="hidden" name="report_date" value="{selected_date}">
                <button type="submit" class="secondary" style="margin: 0; padding: 10px 20px; font-weight: bold; border: none; border-radius: 6px; cursor: pointer; box-shadow: 0 4px 6px var(--shadow);">⚠️ Aceptar y Continuar</button>
              </form>
              <a href="/default?date={selected_date}" style="color: var(--muted); text-decoration: none; font-size: 14px; font-weight: 500; margin-left: 5px;">Cancelar</a>
            </div>
            """
        report_html = f"""
        <div class="tab-content active" id="resumen-ejecutivo">
          <section class="panel error">
            <h2>❌ No se pudo generar el informe</h2>
            <p>{html.escape(error)}</p>
            {confirm_btn_html}
          </section>
        </div>
        <div class="tab-content" id="resumen-diario">
          <section class="panel empty-state"><div class="empty-icon">📊</div><h2>Error al procesar</h2><p>Revisa los archivos subidos.</p></section>
        </div>
        <div class="tab-content" id="previsiones-cierre">
          <section class="panel empty-state"><div class="empty-icon">📈</div><h2>Error al procesar</h2><p>Revisa los archivos subidos.</p></section>
        </div>
        <div class="tab-content" id="calendario-entregas">
          <section class="panel empty-state"><div class="empty-icon">📅</div><h2>Error al procesar</h2><p>Revisa los archivos subidos.</p></section>
        </div>
        <div class="tab-content" id="alertas-auditoria">
          <section class="panel empty-state"><div class="empty-icon">⚠️</div><h2>Error al procesar</h2><p>Revisa los archivos subidos.</p></section>
        </div>
        <div class="tab-content" id="cartera-comparativas">
          <section class="panel empty-state"><div class="empty-icon">💼</div><h2>Error al procesar</h2><p>Revisa los archivos subidos.</p></section>
        </div>
                <div class="tab-content" id="produccion-dashboard">
          <section class="panel empty-state"><div class="empty-icon">🏭</div><h2>Error al procesar</h2><p>Revisa los archivos subidos.</p></section>
        </div>
        <div class="tab-content" id="stock">
          <section class="panel empty-state"><div class="empty-icon">📦</div><h2>Error al procesar</h2><p>Revisa los archivos subidos.</p></section>
        </div>
        {import_form_html}
        """
    elif report:
        current = report["current"]
        alerts = report["alerts"]
        def fmt_split(s: dict[str, float]) -> str:
            if not s: return ""
            return f"<br><span style='font-size:10px;color:var(--muted)'>Nac: {money(s.get('nacional',0))} | Exp: {money(s.get('exportacion',0))}</span>"

        t_ofe_split = report["today"]["ofertas"].get("split", {})
        t_ped_split = report["today"]["pedidos"].get("split", {})
        t_fac_split = report["today"]["facturas"].get("split", {})

        today_rows = [
            ["Ofertas nuevas hoy", str(report["today"]["ofertas"]["cantidad"]), f'{money(report["today"]["ofertas"]["importe"])}{fmt_split(t_ofe_split)}'],
            ["Pedidos entrantes hoy", str(report["today"]["pedidos"]["cantidad"]), f'{money(report["today"]["pedidos"]["importe"])}{fmt_split(t_ped_split)}'],
            ["Facturado hoy", str(report["today"]["facturas"]["cantidad"]), f'{money(report["today"]["facturas"]["importe"])}{fmt_split(t_fac_split)}'],
        ]
        today_rows_html = "".join([f"<tr><td>{html.escape(r[0])}</td><td>{html.escape(r[1])}</td><td>{r[2]}</td></tr>" for r in today_rows])
        m_ped_split = report["month"]["pedidos"].get("split", {})
        m_fac_split = report["month"]["facturas"].get("split", {})
        
        m_ped_split_txt = f"<br><span style='font-size:10px;color:var(--muted)'>Nac: {money(m_ped_split.get('nacional',0))} | Exp: {money(m_ped_split.get('exportacion',0))}</span>" if m_ped_split else ""
        m_fac_split_txt = f"<br><span style='font-size:10px;color:var(--muted)'>Nac: {money(m_fac_split.get('nacional',0))} | Exp: {money(m_fac_split.get('exportacion',0))}</span>" if m_fac_split else ""

        month_rows = [
            ["Acumulado pedidos", str(report["month"]["pedidos"]["cantidad"]), f'{money(report["month"]["pedidos"]["importe"])}{m_ped_split_txt}', money(report["month"]["pedidos"]["hoy"])],
            ["Acumulado facturación", str(report["month"]["facturas"]["cantidad"]), f'{money(report["month"]["facturas"]["importe"])}{m_fac_split_txt}', "-"],
            ["Ratio conversión", "-", pct(report["month"]["conversion"]), "-"],
        ]
        
        month_rows_html = ""
        for row in month_rows:
            month_rows_html += f"<tr><td>{html.escape(row[0])}</td><td>{html.escape(row[1])}</td><td>{row[2]}</td><td>{html.escape(row[3])}</td></tr>"
        forecast = report["forecast"]
        budget_val = forecast["budget"]
        budget_str = money(budget_val) if budget_val is not None else "N/D"
        budget_raw = budget_val if budget_val is not None else 0.0
        
        def fmt_split(s: dict[str, float]) -> str:
            if not s: return ""
            return f"<br><span style='font-size:10px;color:var(--muted)'>Nac: {money(s.get('nacional',0))} | Exp: {money(s.get('exportacion',0))}</span>"

        forecast_table_html = f"""
        <table id="forecast-summary-table">
          <thead>
            <tr><th>Bloque</th><th>Cantidad</th><th>Importe</th></tr>
          </thead>
          <tbody>
            <tr data-row="budget"><td>Presupuesto mes</td><td>-</td><td class="val" data-raw="{budget_raw}">{budget_str}</td></tr>
            <tr data-row="facturado"><td>Facturado real</td><td>{forecast["facturado"]["cantidad"]}</td><td class="val" data-raw="{forecast["facturado"]["importe"]}">{money(forecast["facturado"]["importe"])}{fmt_split(forecast["facturado"].get("split"))}</td></tr>
            <tr data-row="cumplimiento-actual"><td>Cumplimiento actual</td><td>-</td><td class="val">{pct(forecast["cumplimiento_actual"])}</td></tr>
            <tr data-row="albaranes"><td>Albaranes pendientes</td><td>{forecast["albaranes_pendientes"]["cantidad"]}</td><td class="val" data-raw="{forecast["albaranes_pendientes"]["importe"]}">{money(forecast["albaranes_pendientes"]["importe"])}{fmt_split(forecast["albaranes_pendientes"].get("split"))}</td></tr>
            <tr data-row="pedidos-cargables"><td>Pedidos cargables</td><td class="qty">{forecast["pedidos_cargables"]["cantidad"]}</td><td class="val" data-raw="{forecast["pedidos_cargables"]["importe"]}">{money(forecast["pedidos_cargables"]["importe"])}{fmt_split(forecast["pedidos_cargables"].get("split"))}</td></tr>
            <tr data-row="pedidos-antiguos"><td>Pedidos antiguos (Backlog)</td><td>{forecast["pedidos_antiguos"]["cantidad"]}</td><td class="val" data-raw="{forecast["pedidos_antiguos"]["importe"]}">{money(forecast["pedidos_antiguos"]["importe"])}{fmt_split(forecast["pedidos_antiguos"].get("split"))}</td></tr>
            <tr data-row="ofertas-aprobadas"><td>Ofertas aprobadas (+15d)</td><td>{forecast["ofertas_aprobadas"]["cantidad"]}</td><td class="val" data-raw="{forecast["ofertas_aprobadas"]["importe"]}">{money(forecast["ofertas_aprobadas"]["importe"])}{fmt_split(forecast["ofertas_aprobadas"].get("split"))}</td></tr>
            <tr data-row="ofertas-mes"><td>Ofertas entrega mes</td><td>{forecast["ofertas_aprobadas"]["entrega_mes"]}</td><td class="val">-</td></tr>
            <tr data-row="estimacion" style="font-weight: bold; color: var(--accent);"><td>Estimación cierre</td><td>-</td><td class="val" data-raw="{forecast["estimacion_total"]}">{money(forecast["estimacion_total"])}{fmt_split(forecast.get("estimacion_total_split"))}</td></tr>
            <tr data-row="cumplimiento-esperado" style="font-weight: bold; color: var(--accent);"><td>Cumplimiento esperado</td><td>-</td><td class="val">{pct(forecast["cumplimiento_esperado"])}</td></tr>
            <tr data-row="gap" style="font-weight: bold; color: var(--accent);"><td>Gap vs presupuesto</td><td>-</td><td class="val">{money(forecast["gap_esperado"])}</td></tr>
          </tbody>
        </table>
        """
        status_rows = [[r["zona"], r["estado"], r["cantidad"], money(r["importe"])] for r in report["status"]]

        # Section 8: MTD comparison
        mtd_data = report.get("mtd_comparison", {})
        curr_m = mtd_data.get("current", {})
        prev_m = mtd_data.get("prev") or {}
        growth_m = mtd_data.get("growth") or {}
        mtd_rows_html = ""
        for key, label in [("facturado", "Facturado"), ("albaranes", "Albaranes pend."), ("pedidos", "Pedidos pend."), ("ofertas", "Ofertas abiertas"), ("total", "Total Comercial")]:
            cur_val = curr_m.get(key, 0)
            cur_split = curr_m.get(key + '_split', {})
            prv_val = prev_m.get(key, 0)
            prv_split = prev_m.get(key + '_split', {})
            g = growth_m.get(key)
            if g is not None:
                gc = "color:#34d399" if g >= 0 else "color:#f87171"
                ar = "▲" if g >= 0 else "▼"
                gcell = f'<span style="{gc};font-weight:700">{ar} {pct(g)}</span>'
            else:
                gcell = '<span style="color:var(--muted)">N/D</span>'
                
            cur_split_html = f"<br><span style='font-size:10px;color:var(--muted)'>Nac: {money(cur_split.get('nacional', 0))} | Exp: {money(cur_split.get('exportacion', 0))}</span>" if cur_split else ""
            prv_split_html = f"<br><span style='font-size:10px;color:var(--muted)'>Nac: {money(prv_split.get('nacional', 0))} | Exp: {money(prv_split.get('exportacion', 0))}</span>" if prv_split else ""
            
            mtd_rows_html += f'<tr><td>{html.escape(label)}</td><td>{money(cur_val)}{cur_split_html}</td><td>{money(prv_val)}{prv_split_html}</td><td>{gcell}</td></tr>'
        prev_date = mtd_data.get("prev_date")
        prev_date_str = f"Mes Anterior ({fmt_date(prev_date)})" if prev_date else "Mes Anterior (PMTD)"
        mtd_table = f'<table><thead><tr><th>Concepto</th><th>Mes Actual (MTD)</th><th>{prev_date_str}</th><th>Variación</th></tr></thead><tbody>{mtd_rows_html}</tbody></table>'

        # Section 9: YTD Client portfolio
        ytd_data = report.get("ytd_accumulations", {})
        client_list = ytd_data.get("clients", [])[:25]
        if client_list:
            cr = ""
            for c in client_list:
                cr += f'<tr><td>{html.escape(str(c.get("cliente","")))}</td><td>{html.escape(str(c.get("razon_social","")))}</td><td>{html.escape(str(c.get("zona","-")))}</td><td class="text-right" data-order="{c.get("facturado_ytd",0)}">{money(c.get("facturado_ytd",0))}</td><td class="text-right" data-order="{c.get("albaranes_pending",0)}">{money(c.get("albaranes_pending",0))}</td><td class="text-right" data-order="{c.get("pedidos_pending",0)}">{money(c.get("pedidos_pending",0))}</td><td class="text-right" data-order="{c.get("ofertas_pending",0)}">{money(c.get("ofertas_pending",0))}</td><td class="text-right" data-order="{c.get("total_portfolio",0)}"><strong>{money(c.get("total_portfolio",0))}</strong></td></tr>'
            client_table = f'<table class="datatable"><thead><tr><th>Cliente</th><th>Razón Social</th><th>Zona</th><th>Fact. YTD</th><th>Alb. Pend.</th><th>Ped. Pend.</th><th>Ofe. Abiertas</th><th>Total</th></tr></thead><tbody>{cr}</tbody></table>'
        else:
            client_table = "<p class='note'>No hay datos de clientes.</p>"

        # Section 10: Product backlog
        product_list = ytd_data.get("products", [])[:25]
        if product_list:
            pr = ""
            for p in product_list:
                raw_units = p.get("pedidos_unidades", 0)
                units = f"<span style='display:none'>_{int(raw_units) + 1000000000:012d}</span>{raw_units:,.0f}".replace(",", ".")
                clientes_html = f"<div style='font-size: 10.5px; color: var(--muted); margin-top: 4px;'>Top clientes: {html.escape(p.get('top_clientes', '-'))}</div>"
                pr += f'<tr><td>{html.escape(str(p.get("articulo","")))}</td><td>{html.escape(str(p.get("descripcion","")))}{clientes_html}</td><td class="text-right">{units}</td><td class="text-right"><span style="display:none">_{int(p.get("pedidos_importe",0)) + 1000000000:012d}</span>{money(p.get("pedidos_importe",0))}</td><td class="text-right"><span style="display:none">_{int(p.get("ofertas_importe",0)) + 1000000000:012d}</span>{money(p.get("ofertas_importe",0))}</td><td class="text-right"><strong><span style="display:none">_{int(p.get("total_importe",0)) + 1000000000:012d}</span>{money(p.get("total_importe",0))}</strong></td></tr>'
            product_table = f'<table class="datatable"><thead><tr><th>Artículo</th><th>Descripción</th><th>Uds.</th><th>Imp. Pedidos</th><th>Imp. Ofertas</th><th>Total</th></tr></thead><tbody>{pr}</tbody></table>'
        else:
            product_table = "<p class='note'>No hay productos en backlog.</p>"

        app_url = os.environ.get("APP_PUBLIC_URL", "")
        link_html = f'<div style="text-align: center; margin-bottom: 24px;"><a href="{app_url}" target="_blank" style="color: var(--accent, #10b981); text-decoration: none; font-weight: bold; padding: 8px 16px; border: 1px solid var(--accent, #10b981); border-radius: 6px; display: inline-block;">🌍 Acceder al Dashboard Interactivo en Vivo</a></div>' if app_url else ""

        # Resumen Ejecutivo tab
        resumen_ejecutivo_html = f"""
        <div class="tab-content active" id="resumen-ejecutivo">
          <div class="actions" style="margin-bottom: 24px; display: flex; gap: 12px; justify-content: flex-end;">
            <a href="/export-excel?date={selected_date}" download="Resumen_Comercial_{selected_date}.xlsx" onclick="alert('📊 Generando Excel...\\nSe descargará en tu navegador y también se guardará una copia directa en la carpeta del proyecto como Resumen_Comercial_{selected_date}.xlsx')" style="text-decoration: none;"><button type="button" class="secondary">📊 Exportar Excel</button></a>
            <button onclick="downloadPDF(event)" class="secondary">📄 Descargar PDF (v7)</button>
            <button onclick="openEmailModal()">📧 Enviar por Email</button>
          </div>
          
          <div id="pdf-report-content" style="border-radius: 8px; background: #0f172a; padding: 24px; width: 100%; box-sizing: border-box;">
            <div class="pdf-header" style="display: flex; justify-content: space-between; align-items: center; padding-bottom: 12px;">
              <div style="display: flex; align-items: center; gap: 16px;">
                <div style="background: transparent; padding: 6px 12px; border-radius: 6px; display: flex; align-items: center; justify-content: center; box-shadow: 0 2px 4px rgba(0,0,0,0.15); border: 1px solid #e2e8f0;">
                  <img src="{CODIAGRO_LOGO_B64}" alt="Codiagro" style="height: 38px; width: auto; display: block; object-fit: contain;" />
                </div>
                <div>
                  <h1 class="pdf-title" style="margin: 0; font-size: 22px; color: var(--accent, #10b981); font-weight: 800; letter-spacing: -0.5px;">Resumen Ejecutivo Comercial</h1>
                  <p style="margin: 4px 0 0; color: var(--muted); font-size: 13px;">Fecha del Reporte: {fmt_date(current)}</p>
                </div>
              </div>
              <div style="text-align: right;">
                <strong style="color: var(--accent-2, #ef9b00); font-size: 14px; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px; padding-right: 6px; display: inline-block;">Controller Comercial</strong>
                <p style="margin: 4px 0 0; color: var(--muted); font-size: 11px; padding-right: 6px;">Consolidado Comercial</p>
              </div>
            </div>
            <!-- Barra divisoria corporativa doble -->
            <div style="height: 4px; background: var(--accent, #10b981); width: 100%; margin-bottom: 2px;"></div>
            <div style="height: 2px; background: var(--warn, #ef9b00); width: 100%; margin-bottom: 24px;"></div>
            {link_html}

            <!-- KPIs de Hoy -->
            <div class="pdf-kpis" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; margin-bottom: 24px;">
              <div class="panel" style="margin-bottom: 0; padding: 20px; align-items: center; text-align: center; background: var(--panel); border: 1px solid rgba(255, 255, 255, 0.05);">
                <span style="font-size: 28px; margin-bottom: 8px;">💼</span>
                <span style="font-size: 11px; color: var(--muted); text-transform: uppercase; font-weight: 600; letter-spacing: 0.5px;">Ofertas Nuevas Hoy</span>
                <strong style="font-size: 22px; color: var(--ink); margin-top: 6px;">{money(report['today']['ofertas']['importe'])}</strong>
                <span style="font-size: 12px; color: var(--muted); margin-top: 4px; margin-bottom: 8px;">{report['today']['ofertas']['cantidad']} ofertas registradas</span>
                <div style="font-size: 11.5px; font-weight: 500; color: var(--muted); margin-top: 8px; display: flex; justify-content: space-between; width: 100%; padding-top: 8px; border-top: 1px solid rgba(255,255,255,0.05);">
                  <span>🇪🇸 Nac: <strong style="color:var(--ink);">{money(report['today']['ofertas']['split']['nacional'])}</strong></span>
                  <span>🌍 Exp: <strong style="color:var(--ink);">{money(report['today']['ofertas']['split']['exportacion'])}</strong></span>
                </div>
              </div>
              <div class="panel" style="margin-bottom: 0; padding: 20px; align-items: center; text-align: center; background: var(--panel); border: 1px solid rgba(255, 255, 255, 0.05);">
                <span style="font-size: 28px; margin-bottom: 8px;">🛒</span>
                <span style="font-size: 12px; color: var(--muted); text-transform: uppercase; font-weight: 600; letter-spacing: 0.5px;">Pedidos Recibidos Hoy</span>
                <strong style="font-size: 22px; color: var(--accent); margin-top: 6px;">{money(report['today']['pedidos']['importe'])}</strong>
                <span style="font-size: 12px; color: var(--muted); margin-top: 4px; margin-bottom: 8px;">{report['today']['pedidos']['cantidad']} pedidos procesados</span>
                <div style="font-size: 11.5px; font-weight: 500; color: var(--muted); margin-top: 8px; display: flex; justify-content: space-between; width: 100%; padding-top: 8px; border-top: 1px solid rgba(255,255,255,0.05);">
                  <span>🇪🇸 Nac: <strong style="color:var(--ink);">{money(report['today']['pedidos']['split']['nacional'])}</strong></span>
                  <span>🌍 Exp: <strong style="color:var(--ink);">{money(report['today']['pedidos']['split']['exportacion'])}</strong></span>
                </div>
              </div>
              <div class="panel" style="margin-bottom: 0; padding: 20px; align-items: center; text-align: center; background: var(--panel); border: 1px solid rgba(255, 255, 255, 0.05);">
                <span style="font-size: 28px; margin-bottom: 8px;">💳</span>
                <span style="font-size: 12px; color: var(--muted); text-transform: uppercase; font-weight: 600; letter-spacing: 0.5px;">Facturado Real Hoy</span>
                <strong style="font-size: 22px; color: var(--accent-2); margin-top: 6px;">{money(report['today']['facturas']['importe'])}</strong>
                <span style="font-size: 12px; color: var(--muted); margin-top: 4px; margin-bottom: 8px;">{report['today']['facturas']['cantidad']} facturas emitidas</span>
                <div style="font-size: 11.5px; font-weight: 500; color: var(--muted); margin-top: 8px; display: flex; justify-content: space-between; width: 100%; padding-top: 8px; border-top: 1px solid rgba(255,255,255,0.05);">
                  <span>🇪🇸 Nac: <strong style="color:var(--ink);">{money(report['today']['facturas']['split']['nacional'])}</strong></span>
                  <span>🌍 Exp: <strong style="color:var(--ink);">{money(report['today']['facturas']['split']['exportacion'])}</strong></span>
                </div>
              </div>
            </div>

            <!-- Progreso y Cierre Mensual -->
            <div class="pdf-grid-2" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 16px; margin-bottom: 24px;">
              <div class="panel" style="margin-bottom: 0;">
                <h2 style="font-size: 15px;">📈 Previsión y Cumplimiento Mensual</h2>
                <div class="budget-row" style="margin-bottom: 12px; padding: 12px; background: rgba(24, 34, 53, 0.2);">
                  <div class="budget-head">
                    <strong>Facturado vs Presupuesto:</strong>
                    <span>{pct(forecast['cumplimiento_actual'])}</span>
                  </div>
                  <div class="budget-track">
                    <div class="budget-fill {'over' if forecast['cumplimiento_actual'] >= 1.0 else ''}" style="width: {min(100, int(forecast['cumplimiento_actual'] * 100))}%"></div>
                  </div>
                  <div class="budget-foot">
                    <span>Real: {money(forecast['facturado']['importe'])}</span>
                    <span>Presupuesto: {money(forecast['budget'])}</span>
                  </div>
                </div>
                <div class="budget-row" id="executive-budget-esperado" style="padding: 12px; background: rgba(24, 34, 53, 0.2);">
                  <div class="budget-head">
                    <strong>Cierre Estimado vs Presupuesto:</strong>
                    <span>{pct(forecast['cumplimiento_esperado'])}</span>
                  </div>
                  <div class="budget-track">
                    <div class="budget-fill {'over' if forecast['cumplimiento_esperado'] >= 1.0 else ''}" style="width: {min(100, int(forecast['cumplimiento_esperado'] * 100))}%"></div>
                  </div>
                  <div class="budget-foot">
                    <span>Est. Cierre: {money(forecast['estimacion_total'])}</span>
                    <span>Gap: {money(forecast['gap_esperado'])}</span>
                  </div>
                </div>
              </div>
              <div class="panel" style="margin-bottom: 0;">
                <h2 style="font-size: 15px;">📊 Composición del Cierre Estimado</h2>
                <div class="bar-chart" style="margin: 8px 0 0;">
                  <div class="bar-row compact" id="exec-bar-facturado" data-raw="{forecast['facturado']['importe']}" data-nac="{forecast['facturado']['split']['nacional']}" data-exp="{forecast['facturado']['split']['exportacion']}">
                    <span class="bar-label">1. Facturado Real <br><span class="split-subtext" style="font-size:10px;color:var(--muted)">Nac: {money(forecast['facturado']['split']['nacional'])} | Exp: {money(forecast['facturado']['split']['exportacion'])}</span></span>
                    <div class="bar-track" style="display:flex;">
                        <div class="bar-fill nac" style="width: {int(forecast['facturado']['split']['nacional'] / forecast['estimacion_total'] * 100) if forecast['estimacion_total'] else 0}%; background:#10b981;"></div>
                        <div class="bar-fill exp" style="width: {int(forecast['facturado']['split']['exportacion'] / forecast['estimacion_total'] * 100) if forecast['estimacion_total'] else 0}%; background:var(--accent-2);"></div>
                    </div>
                    <span class="bar-value">{money(forecast['facturado']['importe'])}</span>
                  </div>
                  <div class="bar-row compact" id="exec-bar-albaranes" data-raw="{forecast['albaranes_pendientes']['importe']}" data-nac="{forecast['albaranes_pendientes']['split']['nacional']}" data-exp="{forecast['albaranes_pendientes']['split']['exportacion']}">
                    <span class="bar-label">2. Albaranes Pendientes <br><span class="split-subtext" style="font-size:10px;color:var(--muted)">Nac: {money(forecast['albaranes_pendientes']['split']['nacional'])} | Exp: {money(forecast['albaranes_pendientes']['split']['exportacion'])}</span></span>
                    <div class="bar-track" style="display:flex;">
                        <div class="bar-fill nac" style="width: {int(forecast['albaranes_pendientes']['split']['nacional'] / forecast['estimacion_total'] * 100) if forecast['estimacion_total'] else 0}%; background:#10b981;"></div>
                        <div class="bar-fill exp" style="width: {int(forecast['albaranes_pendientes']['split']['exportacion'] / forecast['estimacion_total'] * 100) if forecast['estimacion_total'] else 0}%; background:var(--accent-2);"></div>
                    </div>
                    <span class="bar-value">{money(forecast['albaranes_pendientes']['importe'])}</span>
                  </div>
                  <div class="bar-row compact" id="exec-bar-pedidos-cargables" data-raw="{forecast['pedidos_cargables']['importe']}" data-nac="{forecast['pedidos_cargables']['split']['nacional']}" data-exp="{forecast['pedidos_cargables']['split']['exportacion']}">
                    <span class="bar-label">3. Pedidos Cargables <br><span class="split-subtext" style="font-size:10px;color:var(--muted)">Nac: {money(forecast['pedidos_cargables']['split']['nacional'])} | Exp: {money(forecast['pedidos_cargables']['split']['exportacion'])}</span></span>
                    <div class="bar-track" style="display:flex;">
                        <div class="bar-fill nac" style="width: {int(forecast['pedidos_cargables']['split']['nacional'] / forecast['estimacion_total'] * 100) if forecast['estimacion_total'] else 0}%; background:#10b981;"></div>
                        <div class="bar-fill exp" style="width: {int(forecast['pedidos_cargables']['split']['exportacion'] / forecast['estimacion_total'] * 100) if forecast['estimacion_total'] else 0}%; background:var(--accent-2);"></div>
                    </div>
                    <span class="bar-value">{money(forecast['pedidos_cargables']['importe'])}</span>
                  </div>
                  <div class="bar-row compact" id="exec-bar-estimacion-total" data-nac="{forecast['estimacion_total_split']['nacional']}" data-exp="{forecast['estimacion_total_split']['exportacion']}" style="border-top: 1px solid var(--line); padding-top: 8px; margin-top: 8px;">
                    <span class="bar-label" style="font-weight: 700;">Estimación Total Cierre <br><span class="split-subtext" style="font-size:10px;color:var(--muted)">Nac: {money(forecast['estimacion_total_split']['nacional'])} | Exp: {money(forecast['estimacion_total_split']['exportacion'])}</span></span>
                    <div class="bar-track" style="display:flex;">
                        <div class="bar-fill nac" style="width: {int(forecast['estimacion_total_split']['nacional'] / forecast['estimacion_total'] * 100) if forecast['estimacion_total'] else 0}%; background:#10b981;"></div>
                        <div class="bar-fill exp" style="width: {int(forecast['estimacion_total_split']['exportacion'] / forecast['estimacion_total'] * 100) if forecast['estimacion_total'] else 0}%; background:var(--accent-2);"></div>
                    </div>
                    <span class="bar-value" style="font-weight: 700; color: var(--accent);">{money(forecast['estimacion_total'])}</span>
                  </div>
                </div>
              </div>
            </div>

            <!-- Comparativa MTD (Página 1) -->
            <div class="panel" style="margin-top: 16px; margin-bottom: 0;">
              <h2 style="font-size: 15px;">⚖️ Comparativa MTD vs PMTD</h2>
              {mtd_table}
            </div>

            
          </div>
        </div>
        """
        
        # Generar HTML para Producción
        
        def render_stock_tables(metrics, columns_html, condition_fn=None):
            html_out = ""
            for tipo in ['Sólidos', 'Líquidos', 'Flows', 'SAS']:
                items = [m for m in metrics if m.get('tipo', 'Otros') == tipo and (condition_fn(m) if condition_fn else True)]
                if items:
                    rows = ""
                    for m in items:
                        ped_sort = int(m.get('pedido', 0) + 1000000000)
                        env_sort = int(m.get('stock_envasado', 0) + 1000000000)
                        gra_sort = int(m.get('stock_granel', 0) + 1000000000)
                        if condition_fn:
                            # It's for necesidades
                            necesidad = m['pedido'] - m['stock_envasado'] - m['stock_granel']
                            nec_sort = int(necesidad + 1000000000)
                            rows += f"<tr><td>{html.escape(str(m['codigo']))}</td><td>{html.escape(str(m['descripcion']))}</td><td class='text-right'><span style='display:none'>_{ped_sort:012d}</span>{m['pedido']:,.0f}</td><td class='text-right'><span style='display:none'>_{env_sort:012d}</span>{m['stock_envasado']:,.0f}</td><td class='text-right'><span style='display:none'>_{gra_sort:012d}</span>{m['stock_granel']:,.0f}</td><td class='text-right' style='color:var(--danger); font-weight:bold;'><span style='display:none'>_{nec_sort:012d}</span>{necesidad:,.0f}</td></tr>".replace(",", ".")
                        else:
                            # It's for stock general
                            rows += f"<tr><td>{html.escape(str(m['codigo']))}</td><td>{html.escape(str(m['descripcion']))}</td><td class='text-right'><span style='display:none'>_{ped_sort:012d}</span>{m['pedido']:,.0f}</td><td class='text-right'><span style='display:none'>_{env_sort:012d}</span>{m['stock_envasado']:,.0f}</td><td class='text-right'><span style='display:none'>_{gra_sort:012d}</span>{m['stock_granel']:,.0f}</td></tr>".replace(",", ".")
                    
                    html_out += f"<h3 style='margin-top: 16px; margin-bottom: 8px;'>{tipo}</h3>"
                    html_out += f"<table class='datatable'><thead><tr>{columns_html}</tr></thead><tbody>{rows}</tbody></table>"
            return html_out

        stock_metrics = report.get('stock_comparison', [])
        if stock_metrics:
            stock_cols = "<th>Código</th><th>Descripción</th><th class='text-right'>Material pedido</th><th class='text-right'>Stock envasado</th><th class='text-right'>Stock granel</th>"
            stock_tables_html = render_stock_tables(stock_metrics, stock_cols)
            
            # Pie Chart Data Calculation
            stock_by_tipo = {'Sólidos': 0, 'Líquidos': 0, 'Flows': 0, 'SAS': 0}
            for m in stock_metrics:
                tipo = m.get('tipo', '')
                if tipo in stock_by_tipo:
                    stock_by_tipo[tipo] += max(0, m['pedido'] - m['stock_envasado'] - m['stock_granel'])
                
            pie_labels = list(stock_by_tipo.keys())
            pie_data = list(stock_by_tipo.values())
            
            import json
            import random
            pie_id = f"stockPie_{random.randint(1000, 9999)}"
            
            colors = ['#f59e0b', '#3b82f6', '#10b981', '#a855f7', '#94a3b8']
            custom_legend_html = "<div style='display:flex; justify-content:center; gap: 16px; margin-top: 16px; flex-wrap: wrap;'>"
            for i, label in enumerate(pie_labels):
                val = pie_data[i]
                color = colors[i % len(colors)]
                custom_legend_html += f"<div style='display:flex; align-items:center; gap:6px;'><span style='width:12px; height:12px; background:{color}; border-radius:50%;'></span><span style='font-size:14px;'>{html.escape(label)}: <strong>{val:,.0f}</strong></span></div>".replace(",", ".")
            custom_legend_html += "</div>"
            
            pie_chart_html = f"""
            <div style="width:100%; max-width: 400px; margin: 0 auto 32px;">
                <canvas id="{pie_id}"></canvas>
                {custom_legend_html}
            </div>
            <script>
            document.addEventListener('DOMContentLoaded', function() {{
                if (typeof Chart === 'undefined') return;
                const ctx = document.getElementById('{pie_id}').getContext('2d');
                new Chart(ctx, {{
                    type: 'doughnut',
                    data: {{
                        labels: {json.dumps(pie_labels)},
                        datasets: [{{
                            data: {json.dumps(pie_data)},
                            backgroundColor: {json.dumps(colors)},
                            borderWidth: 0
                        }}]
                    }},
                    options: {{
                        responsive: true,
                        layout: {{
                            padding: 30
                        }},
                        plugins: {{
                            legend: {{ display: false }},
                            tooltip: {{
                                callbacks: {{
                                    label: function(context) {{
                                        let label = context.label || '';
                                        if (label) {{ label += ': '; }}
                                        if (context.parsed !== null) {{
                                            label += new Intl.NumberFormat('es-ES').format(context.parsed);
                                        }}
                                        return label;
                                    }}
                                }}
                            }}
                        }}
                    }},
                    plugins: [{{
                        id: 'sliceLabels',
                        afterDraw(chart) {{
                            const ctx = chart.ctx;
                            chart.data.datasets.forEach((dataset, i) => {{
                                chart.getDatasetMeta(i).data.forEach((element, index) => {{
                                    const val = dataset.data[index];
                                    if(val === 0) return;
                                    const angle = element.endAngle - element.startAngle;
                                    const pos = element.tooltipPosition();
                                    const label = chart.data.labels[index];
                                    const text = label + ' ' + new Intl.NumberFormat('es-ES').format(val);
                                    ctx.save();
                                    ctx.fillStyle = '#ffffff';
                                    ctx.font = 'bold 12px sans-serif';
                                    ctx.textBaseline = 'middle';
                                    ctx.shadowColor = 'rgba(0,0,0,0.8)';
                                    ctx.shadowBlur = 4;

                                    if(angle > 0.3) {{
                                        ctx.textAlign = 'center';
                                        ctx.fillText(text, pos.x, pos.y);
                                    }} else {{
                                        const midAngle = (element.startAngle + element.endAngle) / 2;
                                        const r = element.outerRadius * 1.15;
                                        const outX = element.x + Math.cos(midAngle) * r;
                                        const outY = element.y + Math.sin(midAngle) * r;
                                        ctx.textAlign = outX < element.x ? 'right' : 'left';
                                        
                                        ctx.beginPath();
                                        ctx.moveTo(pos.x, pos.y);
                                        ctx.lineTo(outX, outY);
                                        ctx.lineWidth = 1.5;
                                        ctx.strokeStyle = 'rgba(255,255,255,0.7)';
                                        ctx.shadowBlur = 0;
                                        ctx.stroke();
                                        
                                        ctx.shadowBlur = 4;
                                        ctx.fillText(text, outX + (outX < element.x ? -5 : 5), outY);
                                    }}
                                    ctx.restore();
                                }});
                            }});
                        }}
                    }}]
                }});
            }});
            </script>
            """
            top5_html = "<div style='display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 16px; margin-bottom: 32px;'>"
            for tipo in ['Sólidos', 'Líquidos', 'Flows', 'SAS']:
                items = [dict(m) for m in stock_metrics if m.get('tipo', 'Otros') == tipo]
                for m in items:
                    m['necesidad'] = max(0, m['pedido'] - m['stock_envasado'] - m['stock_granel'])
                
                items_with_need = [m for m in items if m['necesidad'] > 0]
                total_necesidad = sum(m['necesidad'] for m in items_with_need)
                
                if total_necesidad > 0:
                    top5 = sorted(items_with_need, key=lambda x: x['necesidad'], reverse=True)[:5]
                    top5_html += f"""
                    <div class="panel" style="margin-bottom: 0; padding: 16px; background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.05);">
                        <h3 style="margin-top: 0; font-size: 14px; margin-bottom: 12px; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 8px;">Top 5 Necesidades: {tipo}</h3>
                        <ul style="list-style: none; padding: 0; margin: 0; font-size: 13px;">
                    """
                    for m in top5:
                        necesidad_pct = (m['necesidad'] / total_necesidad) * 100
                        item_html = f"""
                            <li style="display: flex; justify-content: space-between; margin-bottom: 8px; border-bottom: 1px dashed rgba(255,255,255,0.05); padding-bottom: 4px;">
                                <span style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 180px; font-weight: 500;" title="{html.escape(m['codigo'])}">{html.escape(m['descripcion'])}</span>
                                <div style="text-align: right;">
                                    <strong style="color: var(--danger);">{m['necesidad']:,.0f}</strong> 
                                    <span style="color: var(--muted); font-size: 11px;">({necesidad_pct:.1f}%)</span>
                                </div>
                            </li>
                        """.replace(',', '.')
                        top5_html += item_html
                    top5_html += "</ul></div>"
            top5_html += "</div>"
            
            stock_insights = report.get("stock_insights", {})
            stock_insights_html = ""
            if stock_insights:
                roturas = stock_insights.get("rotura", [])
                roturas_rows = []
                for r in roturas[:10]:
                    cobertura = "0 días" if r['cobertura'] == 0 else f"{r['cobertura']:.1f} días".replace('.', ',')
                    stock_fmt = f"{r['stock']:,.0f}".replace(',', '.')
                    consumo_fmt = f"{r['consumo_medio']:,.0f}".replace(',', '.')
                    roturas_rows.append(f"<tr><td>{html.escape(r['articulo'])}</td><td class='text-right'>{stock_fmt}</td><td class='text-right'>{consumo_fmt}</td><td class='text-right' style='color:#ef4444; font-weight:bold;'>{cobertura}</td></tr>")
                roturas_rows_str = "".join(roturas_rows)
                roturas_table = f"<table><thead><tr><th>Artículos en Riesgo (< 15 días)</th><th class='text-right'>Stock Actual</th><th class='text-right'>Consumo Medio/Mes</th><th class='text-right'>Cobertura Estimada</th></tr></thead><tbody>{roturas_rows_str}</tbody></table>" if roturas_rows_str else "<p class='note' style='color:var(--success);'>No hay alertas de rotura inminente.</p>"

                total_cap = stock_insights.get("total_capital", 0.0)
                obs_cap = stock_insights.get("obsoleto_capital", 0.0)
                obs_pct = (obs_cap / total_cap * 100) if total_cap > 0 else 0
                
                total_cap_str = f"{total_cap:,.2f} EUR".replace(',', 'X').replace('.', ',').replace('X', '.')
                obs_cap_str = f"{obs_cap:,.2f} EUR".replace(',', 'X').replace('.', ',').replace('X', '.')
                obs_pct_str = f"{obs_pct:.1f}%".replace('.', ',')
                
                cap_groups = stock_insights.get("capital_groups", {})
                cap_groups_list = sorted(cap_groups.values(), key=lambda x: x['total_cant'], reverse=True)[:5]
                cap_rows = []
                for g in cap_groups_list:
                    if not g['items']: continue
                    
                    fam_name = g['base_code']
                    for item in g['items']:
                        if item['codigo'] == g['base_code']:
                            fam_name = item['desc']
                            break
                    if fam_name == g['base_code'] and len(g['items']) > 0:
                        fam_name = g['items'][0]['desc'].split(',')[0].strip()
                        
                    import re
                    fam_name = re.sub(r'(?i)\b(?:GRANEL|BASE)\b', '', fam_name).strip()
                        
                    f = g['formats']
                    def fmt(v): return f"{v:,.0f}".replace(',', '.') if v > 0 else "-"
                    
                    cap_rows.append(f"<tr><td>{html.escape(fam_name)}</td><td class='text-right'>{fmt(f['granel'])}</td><td class='text-right'>{fmt(f['1L'])}</td><td class='text-right'>{fmt(f['5L'])}</td><td class='text-right'>{fmt(f['20L'])}</td><td class='text-right'>{fmt(f['200L'])}</td><td class='text-right'>{fmt(f['1000L'])}</td><td class='text-right' style='font-weight:bold;'>{fmt(g['total_cant'])}</td></tr>")
                cap_rows_str = "".join(cap_rows)
                cap_table = f"<div style='margin-top:12px; font-size:11px; overflow-x:auto;'><table style='margin-bottom:0;'><thead><tr><th>Familia</th><th class='text-right'>Granel</th><th class='text-right'>1L/Kg</th><th class='text-right'>5L/Kg</th><th class='text-right'>20L/Kg</th><th class='text-right'>200L/Kg</th><th class='text-right'>1000L/Kg</th><th class='text-right'>Total Familia</th></tr></thead><tbody>{cap_rows_str}</tbody></table></div>" if cap_rows_str else ""
                
                obs_items = stock_insights.get("obsoleto_items", [])
                obs_rows = []
                for item in obs_items[:5]:
                    name = html.escape(item['articulo'][:40] + ('...' if len(item['articulo']) > 40 else ''))
                    val = f"{item['valor']:,.2f} €".replace(',', 'X').replace('.', ',').replace('X', '.')
                    obs_rows.append(f"<tr><td>{name}</td><td class='text-right' style='color:#ef4444;'>{val}</td></tr>")
                obs_rows_str = "".join(obs_rows)
                obs_table = f"<div style='margin-top:12px; font-size:12px;'><table style='margin-bottom:0;'><thead><tr><th>Top 5 Obsoletos</th><th class='text-right'>Valoración</th></tr></thead><tbody>{obs_rows_str}</tbody></table></div>" if obs_rows_str else ""
                
                stock_insights_html = f"""
                <div style="display: grid; grid-template-columns: 2.2fr 1fr; gap: 16px; margin-bottom: 24px;">
                    <div class="panel kpi-card" style="margin-bottom: 0;">
                        <h3>Capital Inmovilizado (Producto Acabado)</h3>
                        <div class="kpi-value">{total_cap_str}</div>
                        <div class="kpi-trend">Valor estimado del stock actual</div>
                        {cap_table}
                    </div>
                    <div class="panel kpi-card" style="margin-bottom: 0;">
                        <h3>Stock Obsoleto (Sin ventas YTD)</h3>
                        <div class="kpi-value" style="color: #ef4444;">{obs_cap_str}</div>
                        <div class="kpi-trend">Supone un {obs_pct_str} del capital total</div>
                        {obs_table}
                    </div>
                </div>
                <section class="panel" style="margin-bottom: 24px;">
                    <h2 style="color: #ef4444;">🚨 Alerta de Rotura de Stock</h2>
                    <p class="note" style="margin-bottom: 12px;">Artículos con stock para menos de 15 días según el consumo medio mensual (YTD).</p>
                    {roturas_table}
                </section>
                """

            stock_html = f"""
            <div class="tab-content" id="stock">
              {stock_insights_html}
              <section class="panel">
                <h2>📦 Distribución de Necesidades de Fabricación</h2>
                <p class="note" style="margin-bottom: 24px;">Material pendiente de fabricar (Pedidos - Stock Disponible) por sección.</p>
                {pie_chart_html}
                {top5_html}
                {stock_tables_html}
              </section>
            </div>
            """
        else:
            stock_html = """
            <div class="tab-content" id="stock">
              <section class="panel empty-state"><div class="empty-icon">📦</div><h2>Sin datos de Stock</h2><p>Sube el archivo de Stock en Importación.</p></section>
            </div>
            """

        necesidades_list = []
        for m in stock_metrics:
            necesidad = m['pedido'] - m['stock_envasado'] - m['stock_granel']
            if necesidad > 0:
                necesidades_list.append((m, necesidad))
                
        necesidades_list.sort(key=lambda x: x[1], reverse=True)
        
        if necesidades_list:
            necesidades_cols = "<th>Código</th><th>Descripción</th><th class='text-right'>Material pedido</th><th class='text-right'>Stock envasado</th><th class='text-right'>Stock granel</th><th class='text-right'>Necesidad a fabricar</th>"
            sorted_metrics = [item[0] for item in necesidades_list]
            necesidades_tables_html = render_stock_tables(sorted_metrics, necesidades_cols, lambda m: m['pedido'] - m['stock_envasado'] - m['stock_granel'] > 0)
        else:
            necesidades_tables_html = "<p class='note'>No hay necesidades de fabricación (el stock cubre los pedidos).</p>"
            
        tiempos = report.get('tiempos_estimados', {}) if report else {}
        tiempos_html = ""
        if tiempos:
            total_factory_hours = 0
            for tipo in ['Sólidos', 'Líquidos', 'Flows', 'SAS']:
                data = tiempos.get(tipo, {})
                total_factory_hours += data.get('fab', 0) + data.get('env', 0) + data.get('cambios', 0) + data.get('idle', 0.0)

            tiempos_html = "<div style='display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 24px;'>"
            for tipo in ['Sólidos', 'Líquidos', 'Flows', 'SAS']:
                data = tiempos.get(tipo, {})
                fab = data.get('fab', 0)
                env = data.get('env', 0)
                cam = data.get('cambios', 0)
                idle = data.get('idle', 0.0)
                total = fab + env + cam + idle
                if total > 0:
                    dias = total / 16.0
                    semanas = total / 80.0
                    carga_pct = (total / total_factory_hours * 100) if total_factory_hours > 0 else 0
                    
                    progress_html = f"""
                        <div style="margin-top: 8px; margin-bottom: 12px;">
                            <div style="display:flex; justify-content:space-between; font-size:11px; margin-bottom:4px; color:var(--muted);">
                                <span>Carga sobre total fábrica</span>
                                <span style="font-weight:bold; color:var(--accent);">{carga_pct:,.1f}%</span>
                            </div>
                            <div style="width: 100%; background: rgba(255,255,255,0.1); border-radius: 4px; height: 6px; overflow: hidden;">
                                <div style="width: {carga_pct}%; background: var(--accent); height: 100%; border-radius: 4px;"></div>
                            </div>
                        </div>
                    """.replace(',', '.')

                    tiempos_html += f"""
                    <div class="panel" style="margin-bottom:0; background:rgba(255,255,255,0.02); border:1px solid rgba(255,255,255,0.05); padding:16px;">
                        <h3 style="margin-top:0; font-size:14px; margin-bottom:12px; border-bottom:1px solid rgba(255,255,255,0.1); padding-bottom:8px;">⏱️ Tiempos estimados {tipo}</h3>
                        <div style="font-size:18px; font-weight:bold; color:var(--accent); margin-bottom:4px; display:flex; flex-wrap:wrap; gap:8px; align-items:baseline;">
                            <span>{total:,.1f} h</span>
                            <span style="color:var(--muted); font-weight:normal;">|</span>
                            <span>{dias:,.1f} días</span>
                            <span style="color:var(--muted); font-weight:normal;">|</span>
                            <span>{semanas:,.1f} semanas</span>
                        </div>
                        {progress_html}
                        <div style="font-size:12px; color:var(--muted); display:flex; justify-content:space-between; margin-bottom:4px;"><span>Fabricación:</span> <span style="color:#fff;">{fab:,.1f} h</span></div>
                        <div style="font-size:12px; color:var(--muted); display:flex; justify-content:space-between; margin-bottom:4px;"><span>Envasado:</span> <span style="color:#fff;">{env:,.1f} h</span></div>
                        <div style="font-size:12px; color:var(--muted); display:flex; justify-content:space-between; margin-bottom:4px;"><span>Cambios (1h/prod):</span> <span style="color:#fff;">{cam} h</span></div>
                        <div style="font-size:12px; color:var(--muted); display:flex; justify-content:space-between;"><span>Ajuste turno:</span> <span style="color:#fff;">{idle:,.1f} h</span></div>
                    </div>
                    """.replace(',', '.')
            tiempos_html += "</div>"

        prod_metrics = report.get('produccion', {})
        if prod_metrics:
            top_arts_unidades = prod_metrics.get('top_articulos_unidades', [])
            top_arts_coste = prod_metrics.get('top_articulos_coste', [])
            
            art_rows = []
            max_len = max(len(top_arts_unidades), len(top_arts_coste))
            for i in range(max_len):
                u_item = top_arts_unidades[i] if i < len(top_arts_unidades) else {'articulo': '-', 'unidades': 0}
                c_item = top_arts_coste[i] if i < len(top_arts_coste) else {'articulo': '-', 'coste': 0, 'unidades': 0}
                
                u_name = html.escape(u_item['articulo'])
                u_val = f"{u_item['unidades']:,.0f}".replace(',', '.') if u_item['articulo'] != '-' else "-"
                
                c_name = html.escape(c_item['articulo'])
                c_coste = c_item.get('coste', 0)
                c_uds = c_item.get('unidades', 0)
                c_unit = c_coste / c_uds if c_uds > 0 else 0
                c_unit_str = f"{c_unit:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') + " EUR" if c_item['articulo'] != '-' else "-"
                
                art_rows.append(f"<tr><td>{u_name}</td><td class='text-right'>{u_val}</td><td>{c_name}</td><td class='text-right'>{c_unit_str}</td></tr>")
                
            art_rows_str = "".join(art_rows)
            art_table = f"<table><thead><tr><th>Top 5 Unidades</th><th class='text-right'>Unidades</th><th>Top 5 Coste</th><th class='text-right'>Coste Ud.</th></tr></thead><tbody>{art_rows_str}</tbody></table>" if art_rows_str else "<p class='note'>No hay datos.</p>"
            
            abc_html = ""

            # Chart.js Bar Chart for Top Artículos
            top_arts = top_arts_unidades
            if top_arts:
                top_labels = [m['articulo'][:15] + ('...' if len(m['articulo']) > 15 else '') for m in top_arts]
                top_data = [m['unidades'] for m in top_arts]
                
                import json
                import random
                top_id = f"topArt_{random.randint(1000, 9999)}"
                top_chart_html = f"""
                <div style="width:100%; flex: 1; min-height: 250px; position:relative; margin-bottom: 16px;">
                    <canvas id="{top_id}"></canvas>
                </div>
                <script>
                document.addEventListener('DOMContentLoaded', function() {{
                    if (typeof Chart === 'undefined') return;
                    const ctx = document.getElementById('{top_id}').getContext('2d');
                    new Chart(ctx, {{
                        type: 'bar',
                        data: {{
                            labels: {json.dumps(top_labels)},
                            datasets: [{{
                                label: 'Unidades',
                                data: {json.dumps(top_data)},
                                backgroundColor: '#10b981',
                                borderRadius: 4
                            }}]
                        }},
                        options: {{
                            indexAxis: 'y',
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: {{
                                legend: {{ display: false }},
                                tooltip: {{
                                    callbacks: {{
                                        label: function(context) {{
                                            return new Intl.NumberFormat('es-ES').format(context.parsed.x) + ' uds';
                                        }}
                                    }}
                                }}
                            }},
                            scales: {{
                                x: {{ grid: {{ color: 'rgba(255,255,255,0.05)' }}, ticks: {{ color: '#94a3b8' }} }},
                                y: {{ grid: {{ display: false }}, ticks: {{ color: '#94a3b8' }} }}
                            }}
                        }}
                    }});
                }});
                </script>
                """
            else:
                top_chart_html = ""
                
            monthly_evo = prod_metrics.get('monthly_evolution', [])
            if monthly_evo:
                month_names_chart = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
                monthly_labels = [month_names_chart[m['mes'] - 1] for m in monthly_evo]
                monthly_months = [m['mes'] for m in monthly_evo]
                
                types = ['Líquidos', 'Sólidos', 'Flows', 'SAS']
                colors = ['#3b82f6', '#f59e0b', '#10b981', '#8b5cf6', '#94a3b8']
                datasets = []
                for idx, t in enumerate(types):
                    data = [m.get(t, 0) for m in monthly_evo]
                    if any(d > 0 for d in data):
                        datasets.append({
                            'label': t,
                            'data': data,
                            'backgroundColor': colors[idx],
                            'barPercentage': 0.85,
                            'categoryPercentage': 0.95
                        })
                

                monthly_id = f"monthly_{random.randint(1000, 9999)}"
                monthly_chart_html = f"""
                <div style="width:100%; height:450px; position:relative; margin-bottom: 16px;">
                    <canvas id="{monthly_id}"></canvas>
                </div>
                <script>
                document.addEventListener('DOMContentLoaded', function() {{
                    if (typeof Chart === 'undefined') return;
                    const ctx = document.getElementById('{monthly_id}').getContext('2d');
                    new Chart(ctx, {{
                        type: 'bar',
                        data: {{
                            labels: {json.dumps(monthly_labels)},
                            datasets: {json.dumps(datasets)}
                        }},
                        plugins: [{{
                            id: 'custom_labels',
                            afterDatasetsDraw(chart, args, options) {{
                                const {{ ctx }} = chart;
                                ctx.font = 'bold 11px Inter, sans-serif';
                                ctx.textAlign = 'center';
                                ctx.textBaseline = 'middle';
                                ctx.fillStyle = '#ffffff';
                                
                                chart.data.datasets.forEach((dataset, i) => {{
                                    const meta = chart.getDatasetMeta(i);
                                    if(meta.hidden) return;
                                    meta.data.forEach((bar, index) => {{
                                        const val = dataset.data[index];
                                        if (val > 0) {{
                                            const label = Math.round(val / 1000) + 'k';
                                            const base = bar.base;
                                            const x = bar.x;
                                            const width = Math.abs(x - base);
                                            const center_x = (base + x) / 2;
                                            if (width > 24) {{
                                                ctx.fillText(label, center_x, bar.y);
                                            }}
                                        }}
                                    }});
                                }});
                                
                                ctx.fillStyle = '#cbd5e1';
                                ctx.textAlign = 'left';
                                ctx.font = 'bold 12px Inter, sans-serif';
                                chart.data.labels.forEach((_, index) => {{
                                    let sum = 0;
                                    let max_x = 0;
                                    let y = 0;
                                    chart.data.datasets.forEach((dataset, i) => {{
                                        const meta = chart.getDatasetMeta(i);
                                        if(!meta.hidden) {{
                                            const val = dataset.data[index];
                                            sum += val;
                                            const bar = meta.data[index];
                                            if (bar && bar.x > max_x) {{ max_x = bar.x; y = bar.y; }}
                                        }}
                                    }});
                                    if (sum > 0) {{
                                        const totalK = Math.round(sum / 1000).toLocaleString('es-ES') + 'k';
                                        ctx.fillText(totalK, max_x + 8, y);
                                    }}
                                }});
                            }}
                        }}],
                        options: {{
                            onClick: (e, activeElements) => {{
                                if (activeElements.length > 0) {{
                                    const index = activeElements[0].index;
                                    const months = {json.dumps(monthly_months)};
                                    const clickedMonth = months[index];
                                    const year = {current.year};
                                    const date = new Date(year, clickedMonth, 0);
                                    const dateStr = year + '-' + clickedMonth.toString().padStart(2, '0') + '-' + date.getDate().toString().padStart(2, '0');
                                    window.location.href = '/default?date=' + dateStr;
                                }}
                            }},
                            indexAxis: 'y',
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: {{ legend: {{ display: true, position: 'bottom', labels: {{ color: '#94a3b8', boxWidth: 12 }} }} }},
                            scales: {{
                                x: {{ stacked: true, grid: {{ color: 'rgba(255,255,255,0.05)' }}, ticks: {{ color: '#94a3b8' }} }},
                                y: {{ stacked: true, grid: {{ display: false }}, ticks: {{ color: '#94a3b8' }} }}
                            }}
                        }}
                    }});
                }});
                </script>
                """
            else:
                monthly_chart_html = ""
            
            daily_evo = prod_metrics.get('daily_evolution', [])
            if daily_evo:
                part1 = daily_evo[:16]
                part2 = daily_evo[16:]
                evo_rows_html = []
                # Header rows: Days
                row_html = "<tr><th style='padding:6px 8px; font-size:12px; background: rgba(255,255,255,0.05);'>Día</th>"
                for m in part1:
                    fecha_fmt = f"{m['fecha'][8:]}"
                    row_html += f"<td class='text-center' style='padding:8px 4px; font-size:12px; font-weight:500; border-bottom:1px solid rgba(255,255,255,0.1);'>{fecha_fmt}</td>"
                row_html += "</tr>"
                evo_rows_html.append(row_html)
                
                # Data rows: Units
                row_html = "<tr><th style='padding:6px 8px; font-size:12px; background: rgba(255,255,255,0.05);'>Uds</th>"
                for m in part1:
                    row_html += f"<td class='text-center' style='padding:8px 4px; font-size:12px;'>{m['unidades']:,.0f}</td>".replace(",", ".")
                row_html += "</tr>"
                evo_rows_html.append(row_html)
                
                # Header rows: Days (Part 2)
                row_html = "<tr><th style='padding:6px 8px; font-size:12px; background: rgba(255,255,255,0.05); border-top:1px solid rgba(255,255,255,0.1);'>Día</th>"
                for m in part2:
                    fecha_fmt = f"{m['fecha'][8:]}"
                    row_html += f"<td class='text-center' style='padding:8px 4px; font-size:12px; font-weight:500; border-top:1px solid rgba(255,255,255,0.1); border-bottom:1px solid rgba(255,255,255,0.1);'>{fecha_fmt}</td>"
                for _ in range(16 - len(part2)):
                     row_html += "<td style='border-top:1px solid rgba(255,255,255,0.1); border-bottom:1px solid rgba(255,255,255,0.1);'></td>"
                row_html += "</tr>"
                evo_rows_html.append(row_html)
                
                # Data rows: Units (Part 2)
                row_html = "<tr><th style='padding:6px 8px; font-size:12px; background: rgba(255,255,255,0.05);'>Uds</th>"
                for m in part2:
                    row_html += f"<td class='text-center' style='padding:8px 4px; font-size:12px;'>{m['unidades']:,.0f}</td>".replace(",", ".")
                for _ in range(16 - len(part2)):
                     row_html += "<td></td>"
                row_html += "</tr>"
                evo_rows_html.append(row_html)
                
                evo_table_content = "".join(evo_rows_html)
                evo_table = f"<div style='overflow-x: auto;'><table style='width:100%; margin:0; min-width: 600px;'><tbody>{evo_table_content}</tbody></table></div>"
            else:
                evo_table = "<p class='note'>No hay datos.</p>"

            coste_unitario = prod_metrics.get('coste_unitario', 0)
            coste_unitario_str = f"{coste_unitario:.2f} EUR".replace('.', ',')
            
            eficiencia = prod_metrics.get('eficiencia_oee', 0)
            eficiencia_str = f"{eficiencia:.1f}%".replace('.', ',') if eficiencia > 0 else "-"
            eficiencia_color = "#10b981" if eficiencia >= 90 else ("#f59e0b" if eficiencia >= 75 else "#ef4444")
            
            meses_nombres = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
            current_month_name = meses_nombres[current.month - 1]

            produccion_html = f"""
            <div class="tab-content" id="produccion-dashboard">
              <section class="panel">
                <h2>🏭 Dashboard de Producción {current_month_name}</h2>
                <div class="pdf-kpis" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 24px;">
                  <div class="panel kpi-card" style="margin-bottom: 0;">
                    <h3>Volumen MTD</h3>
                    <div class="kpi-value">{prod_metrics.get('mtd_unidades', 0):,.0f}</div>
                    <div class="kpi-trend">Día Ant: {prod_metrics.get('daily_unidades', 0):,.0f}</div>
                  </div>
                  <div class="panel kpi-card" style="margin-bottom: 0;">
                    <h3>Adherencia al Plan</h3>
                    <div class="kpi-value">{prod_metrics.get('adherencia', 0):.1f}%</div>
                    <div class="kpi-trend">Objetivo: {prod_metrics.get('mtd_uds_a_fabricar', 0):,.0f}</div>
                  </div>
                  <div class="panel kpi-card" style="margin-bottom: 0;">
                    <h3>Coste Real Total</h3>
                    <div class="kpi-value">{prod_metrics.get('mtd_coste_real', 0):,.2f} EUR</div>
                    <div class="kpi-trend">Acumulado MTD</div>
                  </div>
                  <div class="panel kpi-card" style="margin-bottom: 0;">
                    <h3>Horas Invertidas</h3>
                    <div class="kpi-value">{prod_metrics.get('mtd_tiempo_real', 0):.1f}h</div>
                  </div>
                </div>
              </section>
              <div style="display: grid; grid-template-columns: 2fr 1fr; gap: 16px; margin-bottom: 24px;">
                <div style="display: flex; flex-direction: column; gap: 16px;">
                  <section class="panel" style="margin-bottom: 0;">
                    <h2>Evolución Diaria (Volumen)</h2>
                    {evo_table}
                  </section>
                  <section class="panel" style="margin-bottom: 0;">
                    <h2>Producción Mensual (YTD)</h2>
                    {monthly_chart_html}
                  </section>
                </div>
                <section class="panel" style="margin-bottom: 0; display: flex; flex-direction: column; height: 100%;">
                  <h2>Top Artículos (Unidades)</h2>
                  {top_chart_html}
                  {art_table}
                  {abc_html}
                </section>
              </div>
              
              {tiempos_html}
              
              <section class="panel" style="margin-bottom: 24px;">
                <h2>⚠️ Cuadro de Necesidades de Fabricación</h2>
                <p class="note" style="margin-bottom: 12px;">Artículos cuyo material pedido supera al stock disponible (envasado + granel).</p>
                {necesidades_tables_html}
              </section>
            </div>
            """
        else:
            produccion_html = f"""
            <div class="tab-content" id="produccion-dashboard">
              <section class="panel empty-state"><div class="empty-icon">🏭</div><h2>Sin datos de Producción</h2><p>Sube el archivo de Producción en Importación.</p></section>
              {tiempos_html}
              <section class="panel" style="margin-bottom: 24px;">
                <h2>⚠️ Cuadro de Necesidades de Fabricación</h2>
                <p class="note" style="margin-bottom: 12px;">Artículos cuyo material pedido supera al stock disponible (envasado + granel).</p>
                {necesidades_tables_html}
              </section>
            </div>
            """

        report_html = f"""
        {resumen_ejecutivo_html}
        
        <div class="tab-content" id="resumen-diario">
          <section class="panel">
            <h2>1. Resumen del Día ({fmt_date(current)})</h2>
            {render_amount_bars(report["charts"]["daily_amounts"])}
            <table>{table_row(["Movimiento", "Cantidad", "Importe"], True)}{today_rows_html}</table>
          </section>
          <section class="panel">
            <h2>2. Evolución del Mes</h2>
            {render_trend_chart(report["charts"]["trend"])}
            <table>{table_row(["Indicador", "Cantidad", "Importe / Ratio", "Entró hoy"], True)}{month_rows_html}</table>
          </section>
          <section class="panel">
            <h2>Recomendaciones</h2>
            <ul>{''.join(f'<li>{html.escape(item)}</li>' for item in report["recommendations"])}</ul>
          </section>
        </div>
        
        <div class="tab-content" id="previsiones-cierre">
          <section class="panel">
            <h2>3. Previsión de Cierre</h2>
            {render_budget_progress(report["charts"]["budget_progress"])}
            {render_amount_bars(report["charts"]["forecast_bridge"])}
            {forecast_table_html}
            {render_forecast_details(forecast, report.get("comments", {}))}
            <p class="note">Los pedidos cargables se valoran por ImporteBrutoPendiente. Las ofertas aprobadas se muestran aparte y no se suman para evitar duplicidad.</p>
          </section>
          <section class="panel">
            <h2>4. Estado de Pedidos</h2>
            {render_status_bars(report["status"])}
            <table>{table_row(["Zona", "Estado", "Cantidad", "Importe"], True)}{''.join(table_row(r) for r in status_rows)}</table>
            <p class="note">{html.escape(report["meta"]["note"])}</p>
          </section>
        </div>
        
        <div class="tab-content" id="calendario-entregas">
          <section class="panel">
            <h2>6. Calendario de Entregas</h2>
            {render_delivery_schedule(report["delivery_schedule"], report.get("comments", {}))}
          </section>
        </div>
        
        <div class="tab-content" id="alertas-auditoria">
          <section class="panel">
            <h2>7. Alertas y Auditoría</h2>
            {render_alerts(report["alerts"], report.get("comments", {}))}
          </section>
        </div>
        
        <div class="tab-content" id="cartera-comparativas">
          <section class="panel">
            <h2>8. Comparativa MTD vs PMTD</h2>
            <p class="note">Acumulado del mes actual vs mismo periodo del mes anterior.</p>
            {mtd_table}
          </section>
          <section class="panel">
            <h2>9. Cartera por Clientes (YTD)</h2>
            <p class="note">Facturado YTD + Albaranes + Pedidos + Ofertas pendientes.</p>
            {client_table}
          </section>
          <section class="panel">
            <h2>10. Backlog por Artículo</h2>
            <p class="note">Demanda pendiente: Pedidos + Ofertas abiertas.</p>
            {product_table}
          </section>
        </div>
                {produccion_html}
        {stock_html}
        {import_form_html}
        """
    else:
        report_html = f"""
        <div class="tab-content active" id="resumen-ejecutivo">
          <section class="panel empty-state"><div class="empty-icon">📋</div><h2>Sin datos</h2><p>Selecciona una fecha o sube archivos en Importación.</p></section>
        </div>
        <div class="tab-content" id="resumen-diario">
          <section class="panel empty-state"><div class="empty-icon">📊</div><h2>Sin datos</h2><p>Selecciona una fecha o sube archivos en Importación.</p></section>
        </div>
        <div class="tab-content" id="previsiones-cierre">
          <section class="panel empty-state"><div class="empty-icon">📈</div><h2>Sin datos</h2><p>Selecciona una fecha o sube archivos en Importación.</p></section>
        </div>
        <div class="tab-content" id="calendario-entregas">
          <section class="panel empty-state"><div class="empty-icon">📅</div><h2>Sin datos</h2><p>Selecciona una fecha o sube archivos en Importación.</p></section>
        </div>
        <div class="tab-content" id="alertas-auditoria">
          <section class="panel empty-state"><div class="empty-icon">⚠️</div><h2>Sin datos</h2><p>Selecciona una fecha o sube archivos en Importación.</p></section>
        </div>
        <div class="tab-content" id="cartera-comparativas">
          <section class="panel empty-state"><div class="empty-icon">💼</div><h2>Sin datos</h2><p>Selecciona una fecha o sube archivos en Importación.</p></section>
        </div>
                <div class="tab-content" id="produccion-dashboard">
          <section class="panel empty-state"><div class="empty-icon">🏭</div><h2>Sin datos</h2><p>Selecciona una fecha o sube archivos en Importación.</p></section>
        </div>
        <div class="tab-content" id="stock">
          <section class="panel empty-state"><div class="empty-icon">📦</div><h2>Sin datos</h2><p>Selecciona una fecha o sube archivos en Importación.</p></section>
        </div>
        {import_form_html}
        """
    # Cargar el layout visual desde layout_template.html
    template_path = BASE_DIR / 'layout_template.html'
    if template_path.exists():
        try:
            html_template = template_path.read_text(encoding='utf-8')
            # Realizar reemplazos
            res = html_template.replace("{selected_date}", selected_date or "")
            res = res.replace("{zona_filter_html}", zona_filter_html or "")
            res = res.replace("{CODIAGRO_LOGO_B64}", CODIAGRO_LOGO_B64 or "")
            res = res.replace("{report_html}", report_html or "")
            return res
        except Exception as e:
            print(f"Error al leer o procesar layout_template.html: {e}")
            return report_html
    else:
        return report_html


def render_alerts(alerts: dict[str, list[dict[str, Any]]], comments: dict, limit: int | None=None) -> str:
    col1 = []
    col1.append('<div class=\'alert-section-title\'>Falta de Fecha Necesaria</div>')
    missing = alerts.get('missing_needed', [])
    if missing:
        missing_sorted = sorted(missing, key=lambda x: pd.to_datetime(x.get('fecha', '2026-01-01')))
        show_list = missing_sorted[:limit] if limit else missing_sorted
        cards_html = []
        for r in show_list:
            doc = r['documento']
            doc_html = render_doc_with_note(doc, comments)
            cards_html.append(f"""
            <div class="alert-card alert-danger">
              <div class="alert-header">
                <span class="alert-badge badge-danger">Fecha Req.</span>
                <span class="alert-doc">Pedido {doc_html}</span>
              </div>
              <div class="alert-body">
                <span class="alert-desc">{html.escape(r['razon_social'])}</span>
                <span class="alert-info-val">{fmt_date(r['fecha'])}</span>
              </div>
            </div>
            """)
        col1.append('<div class=\'alerts-container\'>' + ''.join(cards_html) + '</div>')
        if limit and len(missing_sorted) > limit:
            col1.append(f"<p class=\'note\' style=\'margin-top: 4px; margin-left: 4px; font-style: italic;\'>* ... y {len(missing_sorted) - limit} más.</p>")
    else:
        col1.append('<p class=\'note\'>No hay alertas de pedidos sin Fecha Necesaria.</p>')
        
    col2 = []
    col2.append('<div class=\'alert-section-title\'>Albaranes sin facturar +7 días</div>')
    stale = alerts.get('stale_delivery_notes', [])
    if stale:
        stale_sorted = sorted(stale, key=lambda x: float(x.get('importe', 0)), reverse=True)
        show_list = stale_sorted[:limit] if limit else stale_sorted
        cards_html = []
        for r in show_list:
            doc_html = render_doc_with_note(r['documento'], comments)
            cards_html.append(f"""
            <div class="alert-card alert-warn">
              <div class="alert-header">
                <span class="alert-badge badge-warn">Factura Pend.</span>
                <span class="alert-doc">Albarán {doc_html}</span>
              </div>
              <div class="alert-body">
                <span class="alert-desc">{html.escape(r['razon_social'])}</span>
                <span class="alert-info-val"><strong>{money(r['importe'])}</strong></span>
              </div>
            </div>
            """)
        col2.append('<div class=\'alerts-container\'>' + ''.join(cards_html) + '</div>')
        if limit and len(stale_sorted) > limit:
            col2.append(f"<p class=\'note\' style=\'margin-top: 4px; margin-left: 4px; font-style: italic;\'>* ... y {len(stale_sorted) - limit} más.</p>")
    else:
        col2.append('<p class=\'note\'>No hay albaranes antiguos sin facturar.</p>')
        
    col3 = []
    col3.append('<div class=\'alert-section-title\'>Ofertas estancadas de alto valor</div>')
    stagnant = alerts.get('stagnant_offers', [])
    if stagnant:
        stagnant_sorted = sorted(stagnant, key=lambda x: float(x.get('importe', 0)), reverse=True)
        show_list = stagnant_sorted[:limit] if limit else stagnant_sorted
        cards_html = []
        for r in show_list:
            doc_html = render_doc_with_note(r['documento'], comments)
            cards_html.append(f"""
            <div class="alert-card alert-info">
              <div class="alert-header">
                <span class="alert-badge badge-info">Estancada</span>
                <span class="alert-doc">Oferta {doc_html}</span>
              </div>
              <div class="alert-body">
                <span class="alert-desc">{html.escape(r['razon_social'])}</span>
                <span class="alert-info-val"><strong>{money(r['importe'])}</strong></span>
              </div>
            </div>
            """)
        col3.append('<div class=\'alerts-container\'>' + ''.join(cards_html) + '</div>')
        if limit and len(stagnant_sorted) > limit:
            col3.append(f"<p class=\'note\' style=\'margin-top: 4px; margin-left: 4px; font-style: italic;\'>* ... y {len(stagnant_sorted) - limit} más.</p>")
    else:
        col3.append('<p class=\'note\'>No hay ofertas de alto valor estancadas.</p>')
        
    return '<div class="alerts-grid"><div class="alert-column">' + ''.join(col1) + '</div><div class="alert-column">' + ''.join(col2) + '</div><div class="alert-column">' + ''.join(col3) + '</div></div>'


class MultipartForm:
    """Reemplazo sin dependencias de cgi.FieldStorage para multipart/form-data.

    Reproduce la parte de la API de FieldStorage que usa esta aplicacion:
      form.getvalue('campo')  -> str | None  (primer valor de un campo de texto)
      form.get_file('campo')  -> (filename, bytes) | None
      'campo' in form         -> bool
      iteracion sobre nombres de campos
    """
    def __init__(self) -> None:
        self._values: dict[str, list[str]] = {}
        self._files: dict[str, list[tuple[str, bytes]]] = {}

    def add_value(self, name: str, value: str) -> None:
        self._values.setdefault(name, []).append(value)

    def add_file(self, name: str, filename: str, data: bytes) -> None:
        self._files.setdefault(name, []).append((filename, data))

    def getvalue(self, name: str, default: Any = None) -> Any:
        vals = self._values.get(name)
        if vals:
            return vals[0]
        return default

    def get_file(self, name: str) -> tuple[str, bytes] | None:
        files = self._files.get(name)
        if files:
            return files[0]
        return None

    def __contains__(self, name: str) -> bool:
        return name in self._values or name in self._files

    def __iter__(self):
        seen: set[str] = set()
        for name in self._values:
            if name not in seen:
                seen.add(name)
                yield name
        for name in self._files:
            if name not in seen:
                seen.add(name)
                yield name


def parse_multipart_form(body: bytes, boundary: bytes) -> MultipartForm:
    """Parsea multipart/form-data. Reemplaza cgi.FieldStorage para Python 3.13+."""
    form = MultipartForm()
    delimiter = b'--' + boundary
    # Dividir por el delimitador; ignorar preambulo y epilogo
    parts = body.split(delimiter)
    for part in parts:
        # Una parte valida empieza con \r\n y termina con \r\n
        if part in (b'', b'--', b'--\r\n', b'\r\n'):
            continue
        if part.startswith(b'\r\n'):
            part = part[2:]
        if part.endswith(b'\r\n'):
            part = part[:-2]
        # Separar cabeceras del contenido (linea en blanco \r\n\r\n)
        header_end = part.find(b'\r\n\r\n')
        if header_end == -1:
            continue
        header_block = part[:header_end].decode('utf-8', errors='replace')
        content = part[header_end + 4:]
        # Parsear disposition y filename
        name: str | None = None
        filename: str | None = None
        for line in header_block.split('\r\n'):
            low = line.lower()
            if low.startswith('content-disposition'):
                m_name = re.search(r'name="([^"]*)"', line)
                if m_name:
                    name = m_name.group(1)
                m_file = re.search(r'filename="([^"]*)"', line)
                if m_file:
                    filename = m_file.group(1)
        if not name:
            continue
        if filename is not None:
            # Campo de fichero: guardar aunque filename sea vacio (coherente con cgi)
            form.add_file(name, filename, content)
        else:
            # Campo de texto normal
            form.add_value(name, content.decode('utf-8', errors='replace'))
    return form


def read_post_form(handler: BaseHTTPRequestHandler) -> MultipartForm:
    """Lee el body de una peticion POST multipart/form-data o application/x-www-form-urlencoded."""
    ctype = handler.headers.get('Content-Type', '')
    length = int(handler.headers.get('Content-Length', 0) or 0)
    body = handler.rfile.read(length) if length else b''
    if 'application/x-www-form-urlencoded' in ctype:
        form = MultipartForm()
        params = urllib.parse.parse_qs(body.decode('utf-8', errors='replace'))
        for name, vals in params.items():
            for val in vals:
                form.add_value(name, val)
        return form
    # boundary puede venir como 'multipart/form-data; boundary=...'
    boundary: bytes = b''
    if 'boundary=' in ctype:
        boundary = ctype.split('boundary=', 1)[1].strip().encode('utf-8')
    if boundary:
        return parse_multipart_form(body, boundary)
    # Si no hay boundary (form-urlencoded o sin body), devolver form vacio
    return MultipartForm()


def filter_dfs_by_zona(dfs, zona):
    if not zona or zona not in ['Nacional', 'Exportación']:
        return
    for k in ['ofertas', 'pedidos', 'albaranes', 'facturas']:
        if k in dfs and 'zona' in dfs[k].columns:
            dfs[k] = dfs[k][dfs[k]['zona'] == zona]

class Handler(BaseHTTPRequestHandler):
    def send_html(self, body: str) -> None:
        encoded = body.encode('utf-8')
        try:
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(encoded)))
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self.send_header('Pragma', 'no-cache')
            self.send_header('Expires', '0')
            self.end_headers()
            self.wfile.write(encoded)
        except (ConnectionAbortedError, ConnectionResetError, BrokenPipeError):
            return None
    def send_redirect(self, location: str) -> None:
        self.send_response(303)
        self.send_header('Location', location)
        self.end_headers()
    def do_GET(self) -> None:
        # ***<module>.Handler.do_GET: Failure: Different control flow
        parsed_url = urllib.parse.urlparse(self.path)
        query = urllib.parse.parse_qs(parsed_url.query)
        path = parsed_url.path
        date_param = query.get('date', [None])[0]
        zona_param = query.get('zona', [None])[0]
        if date_param:
            if not re.match('^\\d{4}-\\d{2}-\\d{2}$', date_param):
                date_param = get_default_report_date()
        else:
            date_param = get_default_report_date()
        if path in ['/default', '/', '']:
            adjusted_date = adjust_report_date(date_param)
            if adjusted_date!= date_param:
                self.send_redirect(f'/default?date={adjusted_date}')
                return
            else:
                date_param = adjusted_date
        if path == '/export-excel':
            try:
                dfs = load_report_data(date_param)
                if not dfs and date_param == get_default_report_date() and all((Path(p).exists() for p in DEFAULT_FILES.values())):
                            dfs = {name: parse_excel_to_normalized_df(source, name) for name, source in DEFAULT_FILES.items()}
                            dfs = {name: ensure_types_normalized_df(df, name) for name, df in dfs.items()}
                if dfs:
                    filter_dfs_by_zona(dfs, zona_param)
                    current = pd.Timestamp(date_param)
                    report = build_report_from_data(dfs, current)
                    excel_data = generate_excel_dashboard(dfs, report, date_param)
                    
                    # Guardar copia local en la carpeta raíz del proyecto (padre de BASE_DIR)
                    try:
                        excel_path = BASE_DIR.parent / f"Resumen_Comercial_{date_param}.xlsx"
                        excel_path.write_bytes(excel_data)
                        print(f"Excel guardado localmente en: {excel_path}", flush=True)
                    except Exception as ex_err:
                        print(f"Error al guardar copia local de Excel: {ex_err}", flush=True)
                        
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                    self.send_header('Content-Length', str(len(excel_data)))
                    self.send_header('Content-Disposition', f'attachment; filename=\"Resumen_Comercial_{date_param}.xlsx\"')
                    self.end_headers()
                    self.wfile.write(excel_data)
                else:
                    self.send_error(404, 'No data found')
                    return
            except Exception as e:
                self.send_error(500, f'Error generating Excel: {str(e)}')
        elif path == '/download-pdf':
            try:
                filename_param = query.get('filename', [None])[0]
                if filename_param:
                    filename_param = os.path.basename(filename_param)
                    if not filename_param.endswith('.pdf'):
                        filename_param = f"Resumen_Comercial_{date_param}.pdf"
                else:
                    filename_param = f"Resumen_Comercial_{date_param}.pdf"
                
                pdf_path = BASE_DIR.parent / filename_param
                if pdf_path.exists():
                    pdf_data = pdf_path.read_bytes()
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/pdf')
                    self.send_header('Content-Length', str(len(pdf_data)))
                    inline_param = query.get('inline', ['0'])[0] == '1'
                    disposition = 'inline' if inline_param else 'attachment'
                    self.send_header('Content-Disposition', f'{disposition}; filename=\"{filename_param}\"')
                    self.end_headers()
                    self.wfile.write(pdf_data)
                else:
                    self.send_error(404, f'PDF file not found on server: {filename_param}')
            except Exception as e:
                self.send_error(500, f'Error serving PDF: {str(e)}')
        elif path == '/html2pdf.bundle.min.js':
            try:
                js_path = BASE_DIR / 'html2pdf.bundle.min.js'
                if js_path.exists():
                    js_data = js_path.read_bytes()
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/javascript; charset=utf-8')
                    self.send_header('Content-Length', str(len(js_data)))
                    self.send_header('Cache-Control', 'public, max-age=31536000')
                    self.end_headers()
                    self.wfile.write(js_data)
                else:
                    self.send_error(404, 'JS file not found')
            except Exception as e:
                self.send_error(500, f'Error serving JS: {str(e)}')
        elif path == '/view-pdf':
            html_content = """
            <!DOCTYPE html>
            <html>
            <head>
              <title>PDF Viewer</title>
              <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.min.js"></script>
              <style>
                body { background: #525659; margin: 0; padding: 20px; display: flex; flex-direction: column; align-items: center; gap: 20px; }
                canvas { box-shadow: 0 4px 12px rgba(0,0,0,0.3); background: white; max-width: 100%; display: block; margin-bottom: 20px; }
              </style>
            </head>
            <body>
              <div id="pages-container"></div>
              <script>
                pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';
                const url = '/download-pdf?date=2026-06-30&inline=1';
                pdfjsLib.getDocument(url).promise.then(pdf => {
                  const container = document.getElementById('pages-container');
                  
                  // Render each page sequentially
                  async function renderPages() {
                    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                      const page = await pdf.getPage(pageNum);
                      const viewport = page.getViewport({ scale: 1.5 });
                      const canvas = document.createElement('canvas');
                      const context = canvas.getContext('2d');
                      canvas.height = viewport.height;
                      canvas.width = viewport.width;
                      container.appendChild(canvas);
                      await page.render({ canvasContext: context, viewport: viewport }).promise;
                    }
                  }
                  renderPages();
                });
              </script>
            </body>
            </html>
            """
            self.send_html(html_content)
        else:
            if path in ['/default', '/', '']:
                try:
                    dfs = load_report_data(date_param)
                    if dfs is not None:
                        filter_dfs_by_zona(dfs, zona_param)
                        current = pd.Timestamp(date_param)
                        report = build_report_from_data(dfs, current)
                        self.send_html(render_report(report=report, selected_date=date_param, selected_zona=zona_param))
                    else:
                        if date_param == get_default_report_date() and all((Path(p).exists() for p in DEFAULT_FILES.values())):
                            dfs = {name: parse_excel_to_normalized_df(source, name) for name, source in DEFAULT_FILES.items()}
                            dfs = {name: ensure_types_normalized_df(df, name) for name, df in dfs.items()}
                            save_report_data(date_param, dfs)
                            filter_dfs_by_zona(dfs, zona_param)
                            current = pd.Timestamp(date_param)
                            report = build_report_from_data(dfs, current)
                            self.send_html(render_report(report=report, selected_date=date_param, selected_zona=zona_param))
                        else:
                            self.send_html(render_report(report=None, selected_date=date_param, selected_zona=zona_param))
                except Exception as exc:
                    self.send_html(render_report(error=str(exc), selected_date=date_param, selected_zona=zona_param))
            else:
                self.send_html(render_report(selected_date=date_param, selected_zona=zona_param))
    def do_POST(self) -> None:
        try:
            parsed_url = urllib.parse.urlparse(self.path)
            path = parsed_url.path
            print(f"[do_POST] Incoming POST request to {path}", flush=True)
            if path == '/save-comment':
                try:
                    content_length = int(self.headers.get('Content-Length', 0))
                    post_data = self.rfile.read(content_length).decode('utf-8')
                    data = json.loads(post_data)
                    documento = data.get('documento')
                    comentario = data.get('comentario')
                    if not documento or not comentario:
                        raise ValueError("Falta documento o comentario")
                    if SUPABASE_ENABLED:
                        supabase_request('document_comments', method='POST', data=[{
                            'documento': documento,
                            'comentario': comentario
                        }])
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({'success': True}).encode('utf-8'))
                except Exception as e:
                    print(f"Error saving comment: {e}")
                    self.send_response(500)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode('utf-8'))
                return
            if path == '/save-pdf':
                query = urllib.parse.parse_qs(parsed_url.query)
                date_param = query.get('date', [None])[0] or get_default_report_date()
                print(f"[do_POST] /save-pdf called for date {date_param}", flush=True)
                form = read_post_form(self)
                pdf_data = None
                file_item = form.get_file('pdf')
                if file_item and file_item[0]:
                    pdf_data = file_item[1]
                if not pdf_data:
                    self.send_response(400)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({'success': False, 'error': 'Datos de PDF no encontrados.'}).encode('utf-8'))
                    return
                
                # Guardar copia local en la carpeta raíz del proyecto (padre de BASE_DIR)
                base_filename = f"Resumen_Comercial_{date_param}"
                pdf_path = BASE_DIR.parent / f"{base_filename}.pdf"
                final_filename = f"{base_filename}.pdf"
                counter = 1
                last_err = None
                while counter <= 10:
                    try:
                        pdf_path.write_bytes(pdf_data)
                        break
                    except PermissionError as pe:
                        last_err = pe
                        print(f"[save-pdf] PermissionError writing to {pdf_path}: {pe}", flush=True)
                        final_filename = f"{base_filename} ({counter}).pdf"
                        pdf_path = BASE_DIR.parent / final_filename
                        counter += 1
                else:
                    print(f"[save-pdf] Failed to write PDF locally after 10 attempts: {last_err}", flush=True)
                    self.send_response(500)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({'success': False, 'error': f'No se pudo escribir el archivo PDF por problemas de permisos: {str(last_err)}'}).encode('utf-8'))
                    return
                
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'success': True, 'filename': final_filename}).encode('utf-8'))
                return
                
            if path == '/confirm-upload':
                form = read_post_form(self)
                report_date = form.getvalue('report_date', '') or ''
                report_date = report_date.strip()
                if not report_date:
                    report_date = get_default_report_date()
                
                # Cargar datos guardados temporalmente en la fecha ficticia '9999-12-31'
                dfs = load_data_from_local('9999-12-31')
                if dfs is not None:
                    save_report_data(report_date, dfs)
                    # Limpiar la fecha temporal localmente
                    try:
                        import shutil
                        temp_dir = LOCAL_DATA_DIR / '9999-12-31'
                        if temp_dir.exists():
                            shutil.rmtree(temp_dir)
                    except Exception:
                        pass
                    self.send_redirect(f'/default?date={report_date}')
                else:
                    self.send_html(render_report(error="No se encontraron datos temporales para confirmar. Por favor, sube los archivos de nuevo.", selected_date=report_date))
                return
                
            if path == '/send-email':
                form = read_post_form(self)
                to_email = form.getvalue('to_email', '') or ''
                pdf_data = None
                file_item = form.get_file('pdf')
                if file_item and file_item[0]:
                    pdf_data = file_item[1]
                to_email = to_email.strip()
                if not to_email or not pdf_data:
                    self.send_response(400)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({'success': False, 'error': 'Destinatario o PDF faltante.'}).encode('utf-8'))
                    return
                else:
                    smtp_host = os.environ.get('SMTP_HOST', 'smtp.gmail.com')
                    smtp_port = int(os.environ.get('SMTP_PORT', '587'))
                    smtp_user = os.environ.get('SMTP_USER', '')
                    smtp_pass = os.environ.get('SMTP_PASS', '')
                    smtp_from = os.environ.get('SMTP_FROM', smtp_user)
                    if not smtp_user or not smtp_pass:
                        self.send_response(400)
                        self.send_header('Content-Type', 'application/json')
                        self.end_headers()
                        self.wfile.write(json.dumps({'success': False, 'error': 'El servidor SMTP no está configurado. Configure las variables de entorno SMTP_USER y SMTP_PASS en el servidor.'}).encode('utf-8'))
                    else:
                        from email.mime.multipart import MIMEMultipart
                        from email.mime.text import MIMEText
                        from email.mime.base import MIMEBase
                        from email import encoders
                        import smtplib
                        msg = MIMEMultipart()
                        msg['From'] = smtp_from
                        msg['To'] = to_email
                        msg['Subject'] = f"Resumen Ejecutivo Comercial - {datetime.date.today().strftime('%d/%m/%Y')}"
                        body = 'Adjunto encontrará el Resumen Ejecutivo Comercial del Dashboard.'
                        msg.attach(MIMEText(body, 'plain', 'utf-8'))
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(pdf_data)
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename=\"Resumen_Comercial_{datetime.date.today().isoformat()}.pdf\"')
                        msg.attach(part)
                        with smtplib.SMTP(smtp_host, smtp_port) as server:
                            if smtp_port == 587:
                                server.starttls()
                            server.login(smtp_user, smtp_pass)
                            server.sendmail(smtp_from, [to_email], msg.as_string())
                        self.send_response(200)
                        self.send_header('Content-Type', 'application/json')
                        self.end_headers()
                        self.wfile.write(json.dumps({'success': True, 'message': 'Email enviado con éxito.'}).encode('utf-8'))
            else:
                form = read_post_form(self)
                report_date = form.getvalue('report_date', '') or ''
                report_date = report_date.strip()
                if not report_date or not re.match('^\\d{4}-\\d{2}-\\d{2}$', report_date):
                    report_date = get_default_report_date()
                else:
                    report_date = adjust_report_date(report_date)
                
                # Cargar datos existentes para no sobrescribir con vacíos o defaults si no se subió el archivo
                existing_dfs = load_report_data(report_date) or {}
                
                files = {}
                for key in ['ofertas', 'pedidos', 'albaranes', 'facturas', 'produccion', 'stock']:
                    file_item = form.get_file(key)
                    if file_item and file_item[0]:
                        files[key] = file_item[1]
                
                # Si no subió absolutamente ningún archivo, cargamos los defaults como fallback inicial (opcional)
                if not files and not existing_dfs:
                    for key in ['ofertas', 'pedidos', 'albaranes', 'facturas', 'produccion', 'stock']:
                        if key in DEFAULT_FILES and Path(DEFAULT_FILES[key]).exists():
                            files[key] = DEFAULT_FILES[key]

                dfs_uploaded = {name: parse_excel_to_normalized_df(source, name) for name, source in files.items()}
                if 'produccion' in dfs_uploaded:
                    dfs_uploaded['produccion'].to_csv('uploaded_produccion_raw.csv', index=False, encoding='utf-8')
                dfs_uploaded = {name: ensure_types_normalized_df(df, name) for name, df in dfs_uploaded.items()}
                
                dfs = {}
                for key in ['ofertas', 'pedidos', 'albaranes', 'facturas', 'produccion', 'stock']:
                    if key in dfs_uploaded:
                        dfs[key] = dfs_uploaded[key]
                    else:
                        dfs[key] = existing_dfs.get(key, pd.DataFrame())
                
                # Validación de fechas eliminada a petición del usuario.
                # Procesar siempre los ficheros aunque no tengan datos para el mes.
                
                save_report_data(report_date, dfs)
                self.send_redirect(f'/default?date={report_date}')
        except Exception as exc:
            parsed_url = urllib.parse.urlparse(self.path)
            path = parsed_url.path
            if path in ['/save-pdf', '/send-email']:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'success': False, 'error': str(exc)}).encode('utf-8'))
                return
            
            selected = locals().get('report_date') or get_default_report_date()
            self.send_html(render_report(error=f'Error al procesar o sincronizar con Supabase: {exc}', selected_date=selected))
        return
    def log_message(self, format: str, *args: Any) -> None:
        return
def main() -> None:
    os.chdir(BASE_DIR)
    with socketserver.ThreadingTCPServer(('0.0.0.0', PORT), Handler) as httpd:
        print(json.dumps({'url': f'http://0.0.0.0:{PORT}', 'status': 'running'}), flush=True)
        httpd.serve_forever()
if __name__ == '__main__':
    main()