#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json
import pandas as pd
import re
import sys
from pathlib import Path
from datetime import timedelta
from datetime import datetime
import hashlib
import bisect
import time
##############################Funciones version 15###################################################################################################
##############################Funciones version 15###################################################################################################
##############################Funciones version 15###################################################################################################

class IndicesCasetas:
    """
    Sistema de índices para búsqueda rápida de casetas.
    Optimiza búsquedas por tracto y rango de fechas.
    """
    
    def __init__(self):
        # Índice principal: {tracto: [(fecha_dt, id_caseta), ...]}
        # Las listas están ordenadas por fecha para búsqueda binaria
        self.por_tracto = {}
        
        # Caché de datos de casetas para evitar múltiples lookups
        self.cache_casetas = {}
    
    def construir_desde_db(self, casetas_dict):
        """
        Construye índices desde el diccionario de casetas de la BD.
        
        Parámetros:
        - casetas_dict: db_manager.data['casetas']
        """
        import time
        inicio = time.time()
        
        self.por_tracto.clear()
        self.cache_casetas.clear()
        
        print("[ÍNDICES] Construyendo índices de casetas...")
        
        for id_caseta, datos_caseta in casetas_dict.items():
            tracto = datos_caseta.get('tracto')
            if not tracto:
                continue
            
            # Parsear fecha
            fecha_dt = parse_datetime_caseta(
                datos_caseta.get('fecha'),
                datos_caseta.get('hora')
            )
            
            if not fecha_dt:
                continue
            
            # Agregar a índice por tracto
            if tracto not in self.por_tracto:
                self.por_tracto[tracto] = []
            
            self.por_tracto[tracto].append((fecha_dt, id_caseta))
            
            # Cachear datos de caseta
            self.cache_casetas[id_caseta] = {
                'tracto': tracto,
                'fecha_dt': fecha_dt,
                'datos': datos_caseta
            }
        
        # Ordenar por fecha para búsqueda binaria
        for tracto in self.por_tracto:
            self.por_tracto[tracto].sort(key=lambda x: x[0])
        
        tiempo = time.time() - inicio
        print(f"[ÍNDICES] Construidos en {tiempo:.2f}s")
        print(f"[ÍNDICES] Tractos indexados: {len(self.por_tracto)}")
        print(f"[ÍNDICES] Casetas totales: {len(self.cache_casetas)}")
    
    def buscar_por_rango(self, tracto, fecha_inicio, fecha_final):
        """
        Busca casetas de un tracto dentro de un rango de fechas.
        Usa búsqueda binaria para optimización.
        
        Retorna: lista de id_caseta
        """
        if tracto not in self.por_tracto:
            return []
        
        casetas_tracto = self.por_tracto[tracto]
        
        # Búsqueda binaria del punto de inicio
        import bisect
        idx_inicio = bisect.bisect_left(casetas_tracto, (fecha_inicio, ''))
        idx_final = bisect.bisect_right(casetas_tracto, (fecha_final, 'zzz'))
        
        # Extraer IDs en el rango
        return [id_caseta for _, id_caseta in casetas_tracto[idx_inicio:idx_final]]

class TrackingCambios:
    """
    Sistema para detectar qué viajes han cambiado y necesitan recalcular relaciones.
    Usa hashes para detectar cambios en datos relevantes.
    """
    
    def __init__(self):
        # {id_viaje: hash_datos_relevantes}
        self.hashes_viajes = {}
    
    def calcular_hash_viaje(self, datos_viaje):
        """
        Calcula hash de los datos relevantes para relaciones:
        - tracto
        - fecha_inicio
        - fecha_final
        """
        import hashlib
        
        datos_relevantes = (
            str(datos_viaje.get('tracto', '')),
            str(datos_viaje.get('fecha_inicio', '')),
            str(datos_viaje.get('fecha_final', ''))
        )
        
        return hashlib.md5(str(datos_relevantes).encode()).hexdigest()
    
    def detectar_cambios(self, viajes_dict):
        """
        Detecta qué viajes son nuevos o han cambiado.
        
        Retorna: set de id_viaje que necesitan recalcular relaciones
        """
        viajes_a_procesar = set()
        nuevos_hashes = {}
        
        for id_viaje, datos_viaje in viajes_dict.items():
            hash_actual = self.calcular_hash_viaje(datos_viaje)
            nuevos_hashes[id_viaje] = hash_actual
            
            # Verificar si es nuevo o cambió
            if id_viaje not in self.hashes_viajes:
                viajes_a_procesar.add(id_viaje)  # Nuevo
            elif self.hashes_viajes[id_viaje] != hash_actual:
                viajes_a_procesar.add(id_viaje)  # Cambió
        
        # Actualizar hashes
        self.hashes_viajes = nuevos_hashes
        
        return viajes_a_procesar
    
    def guardar_estado(self, ruta='tracking_cache.json'):
        """Guarda el estado de hashes para persistencia"""
        import json
        try:
            with open(ruta, 'w') as f:
                json.dump(self.hashes_viajes, f)
        except Exception:
            pass
    
    def cargar_estado(self, ruta='tracking_cache.json'):
        """Carga el estado de hashes guardado"""
        import json
        from pathlib import Path
        
        if Path(ruta).exists():
            try:
                with open(ruta, 'r') as f:
                    self.hashes_viajes = json.load(f)
                print(f"[TRACKING] Cargados {len(self.hashes_viajes)} hashes")
            except Exception:
                pass
            
##############################Funciones version 14###################################################################################################
##############################Funciones version 14###################################################################################################
##############################Funciones version 14###################################################################################################

def merge_gastos_data(existing_gastos, new_gastos):
    """
    Combina datos de excel de gastos (multi-hoja) siguiendo lógica: REEMPLAZAR + MANTENER
    
    Parámetros:
    - existing_gastos: dict con hojas {'GASTOS': [...], 'CIERRE 35': [...], 'MODIFICACIONES_COMISION': {...}} o None/0
    - new_gastos: dict con hojas {'GASTOS': [...], 'CIERRE 35': [...], 'MODIFICACIONES_COMISION': {...}}
    
    Retorna: dict con datos combinados
    """
    
    if existing_gastos == 0 or existing_gastos is None:
        return new_gastos if new_gastos else {}
    
    if not new_gastos:
        return existing_gastos
    
    # Procesar hoja GASTOS
    print("[MERGE GASTOS]")
    gastos_combined = merge_single_sheet(
        existing_sheet=existing_gastos.get('GASTOS', []),
        new_sheet=new_gastos.get('GASTOS', []),
        header_rows=1,
        id_col=2  # Columna C (índice 2) = ID viaje en GASTOS
    )
    
    # Procesar hoja CIERRE 35
    print("[MERGE CIERRE 35]")
    cierre35_combined = merge_single_sheet(
        existing_sheet=existing_gastos.get('CIERRE 35', []),
        new_sheet=new_gastos.get('CIERRE 35', []),
        header_rows=1,
        id_col=2  # Columna C (índice 2) = ID viaje en CIERRE 35
    )
    
    # NUEVO: Procesar modificaciones de comisiones (merge de diccionarios)
    modificaciones_existing = existing_gastos.get('MODIFICACIONES_COMISION', {})
    modificaciones_new = new_gastos.get('MODIFICACIONES_COMISION', {})
    
    # Combinar: los nuevos reemplazan a los existentes si hay conflicto
    modificaciones_combined = {**modificaciones_existing, **modificaciones_new}
    
    if modificaciones_new:
        print(f"[MERGE MODIFICACIONES] Existentes: {len(modificaciones_existing)}, Nuevas: {len(modificaciones_new)}, Total: {len(modificaciones_combined)}")
    
    resultado = {
        'GASTOS': gastos_combined,
        'CIERRE 35': cierre35_combined
    }
    
    # Solo incluir modificaciones si existen
    if modificaciones_combined:
        resultado['MODIFICACIONES_COMISION'] = modificaciones_combined
    
    return resultado

def procesar_hojas_operadores(filepath):
    """
    Procesa hojas de operadores (que inician con "O") para calcular modificaciones de comisiones.
    
    Cada hoja contiene:
    - Primera tabla (fila 14+): Columna B=ID viaje, Columna H=Comisión
    - Segunda tabla: Buscar "FONDO AHORRO" y "PAGO 7MO DIA" en columna F, obtener valores de columna H
    
    Calcula: (FONDO AHORRO + PAGO 7MO DIA) / número_viajes
    Y agrega ese valor a cada comisión.
    
    Retorna: dict {id_viaje: nueva_comision}
    """
    try:
        xls = pd.ExcelFile(filepath, engine='openpyxl')
        
        # Filtrar hojas que inician con "O"
        hojas_operadores = [name for name in xls.sheet_names if name.strip().upper().startswith('O')]
        
        if not hojas_operadores:
            return {}
        
        print(f"[GASTOS] Procesando {len(hojas_operadores)} hoja(s) de operadores...")
        
        modificaciones_comisiones = {}
        
        for hoja_nombre in hojas_operadores:
            try:
                # Leer hoja
                df = pd.read_excel(filepath, sheet_name=hoja_nombre, header=None, engine='openpyxl')
                data = df.values.tolist()
                
                if len(data) < 14:
                    continue
                
                # ===== PRIMERA TABLA: Viajes y comisiones =====
                viajes_comisiones = []
                
                # Empezar desde fila 14 (índice 13)
                for i in range(13, len(data)):
                    row = data[i]
                    
                    # Columna B (índice 1)
                    if len(row) <= 1:
                        break
                    
                    id_viaje = row[1]
                    
                    # Si está vacío, terminar lectura de tabla
                    if pd.isna(id_viaje) or str(id_viaje).strip() == '':
                        break
                    
                    # Columna H (índice 7)
                    comision = row[7] if len(row) > 7 else 0.0
                    
                    try:
                        comision_float = float(comision) if not pd.isna(comision) else 0.0
                    except:
                        comision_float = 0.0
                    
                    viajes_comisiones.append({
                        'id_viaje': str(id_viaje).strip(),
                        'comision_original': comision_float
                    })
                
                if not viajes_comisiones:
                    continue
                
                # ===== SEGUNDA TABLA: Fondo ahorro y pago 7mo día =====
                fondo_ahorro = 0.0
                pago_7mo_dia = 0.0
                
                # Buscar en columna F (índice 5) los títulos
                for i in range(len(data)):
                    row = data[i]
                    
                    if len(row) <= 7:  # Necesitamos hasta columna H
                        continue
                    
                    celda_f = str(row[5]).strip().upper() if not pd.isna(row[5]) else ''
                    
                    if 'FONDO AHORRO' in celda_f or 'FONDO DE AHORRO' in celda_f:
                        # Obtener valor de columna H (índice 7)
                        try:
                            fondo_ahorro = float(row[7]) if not pd.isna(row[7]) else 0.0
                        except:
                            fondo_ahorro = 0.0
                    
                    if 'PAGO 7MO DIA' in celda_f or 'PAGO 7MO. DIA' in celda_f or '7MO DIA' in celda_f:
                        # Obtener valor de columna H (índice 7)
                        try:
                            pago_7mo_dia = float(row[7]) if not pd.isna(row[7]) else 0.0
                        except:
                            pago_7mo_dia = 0.0
                
                # ===== CALCULAR DISTRIBUCIÓN =====
                suma_extras = fondo_ahorro + pago_7mo_dia
                num_viajes = len(viajes_comisiones)
                
                if num_viajes > 0 and suma_extras > 0:
                    distribucion_por_viaje = suma_extras / num_viajes
                    
                    print(f"  Hoja '{hoja_nombre[:30]}...': {num_viajes} viajes, Fondo=${fondo_ahorro:,.2f}, 7mo=${pago_7mo_dia:,.2f}, Dist=${distribucion_por_viaje:,.2f}")
                    
                    # Aplicar distribución a cada viaje
                    for viaje in viajes_comisiones:
                        id_viaje = viaje['id_viaje']
                        comision_original = viaje['comision_original']
                        nueva_comision = comision_original + distribucion_por_viaje
                        
                        modificaciones_comisiones[id_viaje] = nueva_comision
                
            except Exception as e:
                print(f"  Error procesando hoja '{hoja_nombre}': {e}")
                continue
        
        print(f"[GASTOS] Total de viajes con comisión modificada: {len(modificaciones_comisiones)}")
        
        return modificaciones_comisiones
        
    except Exception as e:
        print(f"[GASTOS] Error procesando hojas de operadores: {e}")
        return {}

def merge_tablero_costos(existing, new):
    """
    Combina datos de Tablero Costo Ruta siguiendo lógica: REEMPLAZAR + MANTENER
    
    IMPORTANTE: Los datos de tablero son LISTAS de registros (no dicts), donde
    cada fila puede generar múltiples registros (uno por tipo de viaje).
    
    La clave única es: id_ruta + tipo_viaje
    
    Parámetros:
    - existing: lista de dicts [{id_ruta, tipo_viaje, ...}, ...]
    - new: lista de dicts [{id_ruta, tipo_viaje, ...}, ...]
    
    Retorna: lista combinada
    """
    if not existing or existing == 0:
        return new if new else []
    
    if not new:
        return existing
    
    # Construir diccionario por clave única: id_ruta-tipo_viaje
    registros_por_clave = {}
    
    # Agregar existentes
    for registro in existing:
        id_ruta = registro.get('id_ruta')
        tipo_viaje = registro.get('tipo_viaje', 'Sin tipo')
        clave = f"{id_ruta}-{tipo_viaje}"
        registros_por_clave[clave] = registro
    
    # Procesar nuevos (reemplazar o agregar)
    nuevos = 0
    actualizados = 0
    
    for registro in new:
        id_ruta = registro.get('id_ruta')
        tipo_viaje = registro.get('tipo_viaje', 'Sin tipo')
        clave = f"{id_ruta}-{tipo_viaje}"
        
        if clave in registros_por_clave:
            # Reemplazar existente
            registros_por_clave[clave] = registro
            actualizados += 1
        else:
            # Agregar nuevo
            registros_por_clave[clave] = registro
            nuevos += 1
    
    # Convertir de vuelta a lista
    combined = list(registros_por_clave.values())
    
    print(f"Merge Tablero: {nuevos} nuevos, {actualizados} actualizados, {len(combined)} total")
    return combined

def merge_clientes_particulares(existing, new):
    """
    Combina datos de Cliente Particular siguiendo lógica: REEMPLAZAR + MANTENER
    - Si un ID existe en ambos: se reemplaza con el nuevo
    - Si un ID solo existe en existing: se mantiene
    - Si un ID solo existe en new: se agrega
    
    Parámetros:
    - existing: dict {id_ruta: datos}
    - new: dict {id_ruta: datos}
    
    Retorna: dict combinado
    """
    if not existing or existing == 0:
        return new if new else {}
    
    if not new:
        return existing
    
    # Copiar existentes
    combined = existing.copy()
    
    # Contadores
    nuevos = 0
    actualizados = 0
    
    # Procesar nuevos registros
    for id_ruta, datos in new.items():
        if id_ruta in combined:
            # Reemplazar existente
            combined[id_ruta] = datos
            actualizados += 1
        else:
            # Agregar nuevo
            combined[id_ruta] = datos
            nuevos += 1
    
    print(f"Merge Clientes: {nuevos} nuevos, {actualizados} actualizados, {len(combined)} total")
    return combined

##############################Funciones version 13###################################################################################################
##############################Funciones version 13###################################################################################################
##############################Funciones version 13###################################################################################################

def verificar_comisiones_vs_rutas(db_manager):
    """
    Verifica que las comisiones de los viajes coincidan con los pagos de operador de sus rutas.
    Marca viajes con advertencia 'ERR. COM' si no coinciden.
    
    Esta función debe ejecutarse DESPUÉS de:
    1. Enriquecer viajes con datos de gastos (para tener comisión)
    2. Enriquecer viajes con datos de cierre (para tener id_ruta)
    3. Cargar tabla de rutas
    
    Retorna: número de viajes marcados con ERR. COM
    """
    print("\n[VERIFICACIÓN] Comparando comisiones de viajes con pagos de operador...")
    
    viajes_con_error = 0
    viajes_verificados = 0
    viajes_sin_datos = 0
    
    # Obtener todas las rutas para búsqueda rápida
    todas_rutas = db_manager.obtener_todas_rutas()
    
    # Agrupar rutas por id_ruta para acceso rápido
    rutas_por_id = {}
    for ruta in todas_rutas:
        id_ruta = ruta.get('id_ruta')
        if id_ruta:
            if id_ruta not in rutas_por_id:
                rutas_por_id[id_ruta] = []
            rutas_por_id[id_ruta].append(ruta)
    
    # Verificar cada viaje
    for id_viaje, datos_viaje in db_manager.data['viajes'].items():
        id_ruta = datos_viaje.get('id_ruta')
        comision = datos_viaje.get('comision')
        
        # Solo verificar si tiene ambos datos
        if not id_ruta or comision is None:
            viajes_sin_datos += 1
            continue
        
        # Obtener rutas relacionadas
        rutas_relacionadas = rutas_por_id.get(id_ruta, [])
        
        if not rutas_relacionadas:
            # Tiene id_ruta pero no existe en tabla de rutas
            viajes_sin_datos += 1
            continue
        
        # Obtener todos los pagos de operador de las rutas
        pagos_operador = set()
        for ruta in rutas_relacionadas:
            pago = ruta.get('pago_operador')
            if pago is not None:
                # Redondear a 2 decimales para comparación
                pagos_operador.add(round(float(pago), 2))
        
        if not pagos_operador:
            # Ruta no tiene pagos de operador definidos
            viajes_sin_datos += 1
            continue
        
        # Comparar comisión con pagos
        comision_redondeada = round(float(comision), 2)
        
        if comision_redondeada not in pagos_operador:
            # No coincide - marcar error
            if 'advertencias' not in datos_viaje:
                datos_viaje['advertencias'] = []
            
            if 'ERR. COM' not in datos_viaje['advertencias']:
                datos_viaje['advertencias'].append('ERR. COM')
                viajes_con_error += 1
        
        viajes_verificados += 1
    
    print(f"[VERIFICACIÓN] Viajes verificados: {viajes_verificados}")
    print(f"[VERIFICACIÓN] Viajes con ERR. COM: {viajes_con_error}")
    print(f"[VERIFICACIÓN] Viajes sin datos para verificar: {viajes_sin_datos}")
    
    # Guardar cambios
    if viajes_con_error > 0:
        db_manager.guardar_datos()
    
    return viajes_con_error

def imprimir_rutas_relacionadas(rutas, id_ruta_base):
    """
    Imprime registros de ruta relacionados a un viaje.
    Si hay múltiples tipos de viaje, muestra el primero completo y los demás resumidos.
    
    Parámetros:
    - rutas: lista de registros de ruta
    - id_ruta_base: ID base de la ruta (sin tipo de viaje)
    """
    if not rutas:
        return
    
    print(f"\n{'='*150}")
    print(f"INFORMACIÓN DE RUTA: {id_ruta_base}")
    print(f"{'='*150}")
    
    # Ordenar por tipo de viaje para consistencia
    rutas_ordenadas = sorted(rutas, key=lambda x: x.get('tipo_viaje') or '')
    
    # Mostrar primer registro completo
    primera = rutas_ordenadas[0]
    
    print(f"\nDatos Generales:")
    print(f"  Clave: {primera.get('clave') or 'N/A'}")
    print(f"  Nombre Ruta: {primera.get('nombre_ruta') or 'N/A'}")
    print(f"  Cliente: {primera.get('cliente') or 'N/A'}")
    print(f"  Origen: {primera.get('origen') or 'N/A'}")
    print(f"  Entrega: {primera.get('entrega') or 'N/A'}")
    print(f"  Base: {primera.get('base') or 'N/A'}")
    print(f"  Ensamble: {primera.get('ensamble') or 'N/A'}")
    print(f"  Kilómetros: {primera.get('kilometros') or 'N/A'}")
    
    # Mostrar tipos de viaje
    print(f"\nTipos de Viaje ({len(rutas_ordenadas)}):")
    print(f"{'-'*150}")
    print(f"{'Tipo Viaje':15} | {'Flete':12} | {'Tipo Cobro':12} | {'Pago Operador':15} | {'Tipo Cobro Op':15} | {'Estado':12}")
    print(f"{'-'*150}")
    
    for ruta in rutas_ordenadas:
        tipo = ruta.get('tipo_viaje') or 'Sin tipo'
        tipo = str(tipo)[:15]
        
        flete = ruta.get('flete')
        flete_str = f"${flete:,.2f}" if flete is not None else ''
        flete_str = flete_str[:12]
        
        tipo_cobro = str(ruta.get('tipo_cobro') or '')[:12]
        
        pago_op = ruta.get('pago_operador')
        pago_op_str = f"${pago_op:,.2f}" if pago_op is not None else ''
        pago_op_str = pago_op_str[:15]
        
        tipo_cob_op = str(ruta.get('tipo_cobro_op') or '')[:15]
        
        origen_datos = ruta.get('origen_datos', '')
        if origen_datos == 'completo':
            estado = 'Completo'
        elif origen_datos == 'solo_cliente':
            estado = 'Solo Cliente'
        elif origen_datos == 'solo_costos':
            estado = 'Solo Costos'
        else:
            estado = ''
        estado = estado[:12]
        
        print(f"{tipo:15} | {flete_str:12} | {tipo_cobro:12} | {pago_op_str:15} | {tipo_cob_op:15} | {estado:12}")
    
    print(f"{'='*150}\n")

def imprimir_rutas(rutas, max_rutas=50):
    """
    Imprime tabla de rutas en formato tabular.
    Similar a imprimir_filtros pero para rutas.
    
    Parámetros:
    - rutas: lista de dicts con datos de rutas
    - max_rutas: máximo de rutas a mostrar (default: 50)
    """
    if not rutas:
        print("\nNo hay rutas para mostrar.\n")
        return
    
    # Limitar cantidad si es muy grande
    total_rutas = len(rutas)
    if total_rutas > max_rutas:
        print(f"\n⚠️  Se encontraron {total_rutas} rutas. Mostrando primeras {max_rutas}.")
        print(f"   Usa filtros para reducir resultados.\n")
        rutas = rutas[:max_rutas]
    
    print("\n" + "="*220)
    print(f"{'ID Ruta':35} | {'Clave':8} | {'Nombre Ruta':15} | {'Cliente':20} | {'Origen':12} | {'Entrega':12} | {'Base':8} | {'Ensamble':8} | {'KM':6} | {'Tipo Viaje':12} | {'Flete':10} | {'T.Cobro':8} | {'Pago Op':10} | {'T.Cob.Op':8} | {'Estado':12} |")
    print("-"*220)
    
    for r in rutas:
        id_ruta = str(r.get('id_ruta', ''))[:35]
        clave = str(r.get('clave') or '')[:8]
        nombre = str(r.get('nombre_ruta') or '')[:15]
        cliente = str(r.get('cliente') or '')[:20]
        origen = str(r.get('origen') or '')[:12]
        entrega = str(r.get('entrega') or '')[:12]
        base = str(r.get('base') or '')[:8]
        ensamble = str(r.get('ensamble') or '')[:8]
        
        km = r.get('kilometros')
        km_str = f"{km:.0f}" if km is not None else ''
        km_str = km_str[:6]
        
        tipo_viaje = str(r.get('tipo_viaje') or '')[:12]
        
        flete = r.get('flete')
        flete_str = f"${flete:,.0f}" if flete is not None else ''
        flete_str = flete_str[:10]
        
        tipo_cobro = str(r.get('tipo_cobro') or '')[:8]
        
        pago_op = r.get('pago_operador')
        pago_op_str = f"${pago_op:,.0f}" if pago_op is not None else ''
        pago_op_str = pago_op_str[:10]
        
        tipo_cob_op = str(r.get('tipo_cobro_op') or '')[:8]
        
        # Estado del registro
        origen_datos = r.get('origen_datos', '')
        if origen_datos == 'completo':
            estado = 'Completo'
        elif origen_datos == 'solo_cliente':
            estado = 'Solo Cliente'
        elif origen_datos == 'solo_costos':
            estado = 'Solo Costos'
        else:
            estado = ''
        estado = estado[:12]
        
        print(f"{id_ruta:35} | {clave:8} | {nombre:15} | {cliente:20} | {origen:12} | {entrega:12} | {base:8} | {ensamble:8} | {km_str:6} | {tipo_viaje:12} | {flete_str:10} | {tipo_cobro:8} | {pago_op_str:10} | {tipo_cob_op:8} | {estado:12} |")
    
    print("="*220)
    print(f"Total mostrado: {len(rutas)} de {total_rutas}\n")
    
def listar_todas_rutas_db(db_manager, max_mostrar=50):
    """
    Lista todas las rutas en la base de datos.
    
    Parámetros:
    - db_manager: instancia de DatabaseManager
    - max_mostrar: máximo de registros a mostrar
    """
    rutas = db_manager.obtener_todas_rutas()
    
    if not rutas:
        print("\nNo hay rutas en la base de datos.\n")
        return []
    
    print(f"\n=== TODAS LAS RUTAS ({len(rutas)} registros) ===")
    imprimir_rutas(rutas, max_rutas=max_mostrar)
    
    return rutas


def buscar_rutas_db(db_manager, filtros):
    """
    Busca rutas aplicando filtros.
    
    Filtros soportados:
    - clave: busca en clave de ruta
    - nombre: busca en nombre de ruta
    - cliente: busca en nombre de cliente
    - origen: busca en origen
    - destino/entrega: busca en entrega
    - tipo: busca en tipo de viaje
    - ensamble: busca en tipo de ensamble
    
    Parámetros:
    - db_manager: instancia de DatabaseManager
    - filtros: dict con filtros a aplicar
    
    Retorna: lista de rutas que cumplen los filtros
    """
    todas_rutas = db_manager.obtener_todas_rutas()
    
    if not todas_rutas:
        print("\nNo hay rutas en la base de datos.\n")
        return []
    
    def cumple_filtros(ruta):
        # Filtro clave
        if 'clave' in filtros and filtros['clave'] is not None:
            clave_f = str(filtros['clave']).strip().lower()
            val = str(ruta.get('clave') or '').lower()
            if clave_f not in val:
                return False
        
        # Filtro nombre
        if 'nombre' in filtros and filtros['nombre'] is not None:
            nombre_f = str(filtros['nombre']).strip().lower()
            val = str(ruta.get('nombre_ruta') or '').lower()
            if nombre_f not in val:
                return False
        
        # Filtro cliente
        if 'cliente' in filtros and filtros['cliente'] is not None:
            cliente_f = str(filtros['cliente']).strip().lower()
            val = str(ruta.get('cliente') or '').lower()
            if cliente_f not in val:
                return False
        
        # Filtro origen
        if 'origen' in filtros and filtros['origen'] is not None:
            origen_f = str(filtros['origen']).strip().lower()
            val = str(ruta.get('origen') or '').lower()
            if origen_f not in val:
                return False
        
        # Filtro destino/entrega
        if 'destino' in filtros and filtros['destino'] is not None:
            destino_f = str(filtros['destino']).strip().lower()
            val = str(ruta.get('entrega') or '').lower()
            if destino_f not in val:
                return False
        
        # Filtro tipo de viaje
        if 'tipo' in filtros and filtros['tipo'] is not None:
            tipo_f = str(filtros['tipo']).strip().lower()
            val = str(ruta.get('tipo_viaje') or '').lower()
            if tipo_f not in val:
                return False
        
        # Filtro ensamble
        if 'ensamble' in filtros and filtros['ensamble'] is not None:
            ensamble_f = str(filtros['ensamble']).strip().lower()
            val = str(ruta.get('ensamble') or '').lower()
            if ensamble_f not in val:
                return False
        
        return True
    
    filtradas = [r for r in todas_rutas if cumple_filtros(r)]
    
    if not filtradas:
        print("\nNo se encontraron rutas que cumplan los filtros.\n")
        return []
    
    imprimir_rutas(filtradas)
    return filtradas


def buscar_ruta_por_id_db(db_manager, id_ruta_buscar):
    """
    Busca rutas por ID exacto o parcial.
    
    Parámetros:
    - db_manager: instancia de DatabaseManager
    - id_ruta_buscar: ID completo o parcial
    """
    todas_rutas = db_manager.obtener_todas_rutas()
    
    id_lower = id_ruta_buscar.lower()
    encontradas = [r for r in todas_rutas if id_lower in r.get('id_ruta', '').lower()]
    
    if not encontradas:
        print(f"\nNo se encontró ruta con ID '{id_ruta_buscar}'.\n")
        return []
    
    print(f"\n=== Rutas encontradas para '{id_ruta_buscar}' ===")
    imprimir_rutas(encontradas)
    
    return encontradas


def estadisticas_rutas_db(db_manager):
    """
    Muestra estadísticas de la tabla de rutas.
    """
    rutas = db_manager.obtener_todas_rutas()
    
    if not rutas:
        print("\nNo hay rutas en la base de datos.\n")
        return
    
    total = len(rutas)
    completas = sum(1 for r in rutas if r.get('origen_datos') == 'completo')
    solo_cliente = sum(1 for r in rutas if r.get('origen_datos') == 'solo_cliente')
    solo_costos = sum(1 for r in rutas if r.get('origen_datos') == 'solo_costos')
    
    # Contar tipos de viaje
    tipos_viaje = {}
    for r in rutas:
        tipo = r.get('tipo_viaje')
        if tipo:
            tipos_viaje[tipo] = tipos_viaje.get(tipo, 0) + 1
    
    # Contar clientes únicos
    clientes_unicos = set()
    for r in rutas:
        cliente = r.get('cliente')
        if cliente:
            clientes_unicos.add(cliente)
    
    # IDs de ruta únicos
    ids_unicos = set(r.get('id_ruta') for r in rutas)
    
    print("\n" + "="*60)
    print("ESTADÍSTICAS DE RUTAS")
    print("="*60)
    print(f"Total de registros: {total}")
    print(f"IDs de ruta únicos: {len(ids_unicos)}")
    print(f"\nPor completitud:")
    print(f"  - Completos (cliente + costo): {completas} ({completas/total*100:.1f}%)")
    print(f"  - Solo cliente: {solo_cliente} ({solo_cliente/total*100:.1f}%)")
    print(f"  - Solo costos: {solo_costos} ({solo_costos/total*100:.1f}%)")
    
    if tipos_viaje:
        print(f"\nPor tipo de viaje:")
        for tipo, count in sorted(tipos_viaje.items(), key=lambda x: -x[1]):
            print(f"  - {tipo}: {count} ({count/total*100:.1f}%)")
    
    print(f"\nClientes únicos: {len(clientes_unicos)}")
    print("="*60 + "\n")

##############################Funciones version 12###################################################################################################
##############################Funciones version 12###################################################################################################
##############################Funciones version 12###################################################################################################

def detectar_excel_rutas(data):
    """
    Detecta si un excel es de tipo 'cliente_particular' o 'tablero_costo_ruta'
    basándose en celdas específicas.
    
    Retorna: 'cliente_particular', 'tablero_costo_ruta', o None
    """
    if not data or len(data) == 0:
        return None
    
    # Verificar primera fila
    primera_fila = data[0] if len(data) > 0 else []
    
    # Cliente Particular: celda D1 (índice 3) = "Base Origen"
    if len(primera_fila) > 3:
        celda_d1 = str(primera_fila[3]).strip() if primera_fila[3] is not None else ''
        if celda_d1 == "Base Origen":
            return 'cliente_particular'
    
    # Tablero Costo Ruta: celda E1 (índice 4) = "Tipo Ensamble"
    if len(primera_fila) > 4:
        celda_e1 = str(primera_fila[4]).strip() if primera_fila[4] is not None else ''
        if celda_e1 == "Tipo Ensamble":
            return 'tablero_costo_ruta'
    
    return None


def crear_id_ruta(clave, nombre_ruta):
    """
    Crea un ID combinado para rutas: 'CLAVE-NOMBRE_RUTA'
    
    Retorna: string con ID combinado o None si faltan datos
    """
    if pd.isna(clave) or pd.isna(nombre_ruta):
        return None
    
    clave_str = str(clave).strip()
    nombre_str = str(nombre_ruta).strip()
    
    if not clave_str or not nombre_str:
        return None
    
    return f"{clave_str}-{nombre_str}"


##############################Funciones de Parseo de Excel Rutas###################################################################################################

def parsear_cliente_particular(data, header_rows=1):
    """
    Parsea excel de Cliente Particular y retorna diccionario por ID_RUTA.
    Detecta y alerta sobre duplicados.
    
    Columnas: A (Clave), B (Nombre ruta), C (Cliente), D (Origen), H (Entrega), I (Base)
    
    Retorna: dict {ID_RUTA: {datos}}
    """
    if not data or len(data) <= header_rows:
        return {}
    
    clientes = {}
    duplicados_encontrados = []
    
    for idx, row in enumerate(data[header_rows:], start=header_rows+1):
        try:
            clave = row[0] if len(row) > 0 else None
            nombre_ruta = row[1] if len(row) > 1 else None
            
            # Crear ID_RUTA
            id_ruta = crear_id_ruta(clave, nombre_ruta)
            
            if not id_ruta:
                continue
            
            # Detectar duplicado
            if id_ruta in clientes:
                duplicados_encontrados.append((id_ruta, idx))
                print(f"  ⚠️  DUPLICADO detectado en Cliente Particular: '{id_ruta}' (fila {idx})")
                continue  # Ignorar duplicado, mantener el primero
            
            cliente = row[2] if len(row) > 2 else None
            origen = row[3] if len(row) > 3 else None      # Columna D
            entrega = row[7] if len(row) > 7 else None     # Columna H
            base = row[8] if len(row) > 8 else None        # Columna I
            
            clientes[id_ruta] = {
                'id_ruta': id_ruta,
                'clave': str(clave).strip() if clave and not pd.isna(clave) else None,
                'nombre_ruta': str(nombre_ruta).strip() if nombre_ruta and not pd.isna(nombre_ruta) else None,
                'cliente': str(cliente).strip() if cliente and not pd.isna(cliente) else None,
                'origen': str(origen).strip() if origen and not pd.isna(origen) else None,
                'entrega': str(entrega).strip() if entrega and not pd.isna(entrega) else None,
                'base': str(base).strip() if base and not pd.isna(base) else None,
            }
        except Exception as e:
            print(f"[PARSE ERROR] Cliente Particular fila {idx}: {e}")
            continue
    
    if duplicados_encontrados:
        print(f"  ⚠️  Total duplicados en Cliente Particular: {len(duplicados_encontrados)}")
    
    print(f"[PARSE] Cliente Particular: {len(clientes)} registros únicos")
    return clientes


def parsear_tablero_costo_ruta(data, header_rows=1):
    """
    Parsea excel de Tablero Costo Ruta y retorna lista de registros.
    IMPORTANTE: Crea un registro POR CADA tipo de viaje detectado.
    Detecta y alerta sobre duplicados de ID_RUTA.
    
    Retorna: list [{id_ruta, clave, nombre_ruta, ensamble, kilometros, tipo_viaje, flete, tipo_cobro, pago_operador, tipo_cobro_op}]
    """
    if not data or len(data) <= header_rows:
        return []
    
    # Mapeo de columnas por tipo de viaje (índices 0-based)
    # Tipo: (col_flete, col_tipo_cobro, col_pago_op, col_tipo_cobro_op)
    tipos_viaje_cols = {
        'Ensacado': (16, 14, 41, 39),      # Q, O, AP, AN
        'Granel': (21, 19, 45, 43),        # V, T, AT, AR
        'Otro': (26, 24, 47, 49),          # AA, Y, AV, AX   - original AX-AV / 49-47, girados ya que en excel estan alrevez
        'Piedra caliza': (31, 29, 53, 51), # AF, AD, BB, AZ
        'BigBag': (36, 34, 57, 55)         # AK, AI, BF, BD
    }
    
    registros = []
    ids_vistos = set()
    duplicados_encontrados = []
    
    for idx, row in enumerate(data[header_rows:], start=header_rows+1):
        try:
            clave = row[0] if len(row) > 0 else None
            nombre_ruta = row[1] if len(row) > 1 else None
            
            # Crear ID_RUTA
            id_ruta = crear_id_ruta(clave, nombre_ruta)
            
            if not id_ruta:
                continue
            
            # Detectar duplicado de fila (no de registro final)
            if id_ruta in ids_vistos:
                duplicados_encontrados.append((id_ruta, idx))
                print(f"  ⚠️  DUPLICADO detectado en Tablero Costo: '{id_ruta}' (fila {idx})")
                continue  # Ignorar fila duplicada completa
            
            ids_vistos.add(id_ruta)
            
            ensamble = row[4] if len(row) > 4 else None      # Columna E
            kilometros = row[10] if len(row) > 10 else None  # Columna K
            
            # Datos base comunes
            datos_base = {
                'id_ruta': id_ruta,
                'clave': str(clave).strip() if clave and not pd.isna(clave) else None,
                'nombre_ruta': str(nombre_ruta).strip() if nombre_ruta and not pd.isna(nombre_ruta) else None,
                'ensamble': str(ensamble).strip() if ensamble and not pd.isna(ensamble) else None,
                'kilometros': float(kilometros) if kilometros and not pd.isna(kilometros) else None,
            }
            
            # Detectar tipos de viaje y crear un registro por cada uno
            tiene_algun_tipo = False
            
            for tipo_nombre, (col_flete, col_tipo_cobro, col_pago_op, col_tipo_cobro_op) in tipos_viaje_cols.items():
                # Extraer valores
                flete = row[col_flete] if len(row) > col_flete else None
                tipo_cobro = row[col_tipo_cobro] if len(row) > col_tipo_cobro else None
                pago_op = row[col_pago_op] if len(row) > col_pago_op else None
                tipo_cobro_op = row[col_tipo_cobro_op] if len(row) > col_tipo_cobro_op else None
                
                # Verificar si tiene datos (al menos uno de los 4 campos)
                tiene_datos = any([
                    flete is not None and not pd.isna(flete),
                    tipo_cobro is not None and not pd.isna(tipo_cobro),
                    pago_op is not None and not pd.isna(pago_op),
                    tipo_cobro_op is not None and not pd.isna(tipo_cobro_op)
                ])
                
                if tiene_datos:
                    # Crear registro para este tipo de viaje
                    registro = datos_base.copy()
                    registro['tipo_viaje'] = tipo_nombre
                    registro['flete'] = float(flete) if flete and not pd.isna(flete) else None
                    registro['tipo_cobro'] = str(tipo_cobro).strip() if tipo_cobro and not pd.isna(tipo_cobro) else None
                    registro['pago_operador'] = float(pago_op) if pago_op and not pd.isna(pago_op) else None
                    registro['tipo_cobro_op'] = str(tipo_cobro_op).strip() if tipo_cobro_op and not pd.isna(tipo_cobro_op) else None
                    
                    registros.append(registro)
                    tiene_algun_tipo = True
            
            # Si la fila no tiene ningún tipo de viaje, crear registro parcial
            if not tiene_algun_tipo:
                registro = datos_base.copy()
                registro['tipo_viaje'] = None
                registro['flete'] = None
                registro['tipo_cobro'] = None
                registro['pago_operador'] = None
                registro['tipo_cobro_op'] = None
                registros.append(registro)
                
        except Exception as e:
            print(f"[PARSE ERROR] Tablero Costo Ruta fila {idx}: {e}")
            continue
    
    if duplicados_encontrados:
        print(f"  ⚠️  Total duplicados en Tablero Costo: {len(duplicados_encontrados)}")
    
    print(f"[PARSE] Tablero Costo Ruta: {len(registros)} registros (de {len(ids_vistos)} ID_RUTA únicos)")
    return registros


##############################Función para Combinar y Crear Tabla Rutas###################################################################################################

def construir_tabla_rutas(clientes_dict, costos_list):
    """
    Combina datos de clientes y costos para crear la tabla final de Rutas.
    
    Proceso:
    1. Para cada ID_RUTA en costos, buscar en clientes y combinar
    2. Para IDs en clientes pero no en costos, crear registros parciales
    3. Retorna dict {id_ruta_tipo: {registro_completo}}
    
    La clave del dict será ID_RUTA-TIPO_VIAJE para manejar múltiples tipos por ruta
    
    Retorna: dict de registros de rutas
    """
    tabla_rutas = {}
    ids_procesados = set()
    
    print("\n[CONSTRUCCIÓN] Combinando datos de clientes y costos...")
    
    # Paso 1: Procesar registros de costos (pueden tener múltiples tipos de viaje)
    for costo in costos_list:
        id_ruta = costo['id_ruta']
        tipo_viaje = costo.get('tipo_viaje', 'Sin tipo')
        
        # Crear clave única: ID_RUTA-TIPO_VIAJE
        clave_registro = f"{id_ruta}-{tipo_viaje}" if tipo_viaje else id_ruta
        
        # Buscar cliente correspondiente
        cliente = clientes_dict.get(id_ruta, {})
        
        # Combinar datos
        registro = {
            'id_ruta': id_ruta,
            'clave': costo.get('clave') or cliente.get('clave'),
            'nombre_ruta': costo.get('nombre_ruta') or cliente.get('nombre_ruta'),
            'cliente': cliente.get('cliente'),
            'origen': cliente.get('origen'),
            'entrega': cliente.get('entrega'),
            'base': cliente.get('base'),
            'ensamble': costo.get('ensamble'),
            'kilometros': costo.get('kilometros'),
            'tipo_viaje': tipo_viaje,
            'flete': costo.get('flete'),
            'tipo_cobro': costo.get('tipo_cobro'),
            'pago_operador': costo.get('pago_operador'),
            'tipo_cobro_op': costo.get('tipo_cobro_op'),
            'origen_datos': 'completo' if cliente else 'solo_costos'
        }
        
        tabla_rutas[clave_registro] = registro
        ids_procesados.add(id_ruta)
    
    # Paso 2: Agregar clientes que no tienen costos
    for id_ruta, cliente in clientes_dict.items():
        if id_ruta not in ids_procesados:
            registro = {
                'id_ruta': id_ruta,
                'clave': cliente.get('clave'),
                'nombre_ruta': cliente.get('nombre_ruta'),
                'cliente': cliente.get('cliente'),
                'origen': cliente.get('origen'),
                'entrega': cliente.get('entrega'),
                'base': cliente.get('base'),
                'ensamble': None,
                'kilometros': None,
                'tipo_viaje': None,
                'flete': None,
                'tipo_cobro': None,
                'pago_operador': None,
                'tipo_cobro_op': None,
                'origen_datos': 'solo_cliente'
            }
            
            tabla_rutas[id_ruta] = registro
    
    # Estadísticas
    completos = sum(1 for r in tabla_rutas.values() if r['origen_datos'] == 'completo')
    solo_cliente = sum(1 for r in tabla_rutas.values() if r['origen_datos'] == 'solo_cliente')
    solo_costos = sum(1 for r in tabla_rutas.values() if r['origen_datos'] == 'solo_costos')
    
    print(f"[CONSTRUCCIÓN] Registros creados: {len(tabla_rutas)}")
    print(f"  - Completos (cliente + costo): {completos}")
    print(f"  - Solo cliente: {solo_cliente}")
    print(f"  - Solo costo: {solo_costos}")
    
    if solo_cliente > 0:
        print(f"  ⚠️  Advertencia: {solo_cliente} rutas solo tienen datos de cliente")
    if solo_costos > 0:
        print(f"  ⚠️  Advertencia: {solo_costos} rutas solo tienen datos de costo")
    
    return tabla_rutas

##############################Funciones version 11###################################################################################################
##############################Funciones version 11###################################################################################################
##############################Funciones version 11###################################################################################################

def detectar_casetas_duplicadas_db(db_manager, diferencia_minutos=5, tracto_filter=None):
    """
    Detecta casetas del mismo tracto con diferencia de tiempo menor o igual a diferencia_minutos.
    Estas casetas pueden ser errores humanos en la captura de datos.
    
    Parámetros:
    - db_manager: instancia de DatabaseManager
    - diferencia_minutos: umbral de minutos para considerar duplicado (default: 5)
    - tracto_filter: si se especifica, solo busca en ese tracto
    
    Retorna: lista de pares de casetas sospechosas
    """
    
    print(f"\n[DETECCIÓN] Buscando casetas con diferencia <= {diferencia_minutos} minutos...")
    
    # Agrupar casetas por tracto
    casetas_por_tracto = {}
    
    for id_caseta, datos_caseta in db_manager.data['casetas'].items():
        tracto = datos_caseta.get('tracto', '')
        
        # Aplicar filtro de tracto si existe
        if tracto_filter is not None and tracto_filter not in str(tracto):
            continue
        
        if not tracto or pd.isna(tracto):
            continue
        
        # Parsear fecha y hora
        fecha_cas = parse_datetime_caseta(datos_caseta.get('fecha'), datos_caseta.get('hora'))
        
        if fecha_cas is None:
            continue
        
        if tracto not in casetas_por_tracto:
            casetas_por_tracto[tracto] = []
        
        casetas_por_tracto[tracto].append({
            'id_caseta': id_caseta,
            'tracto': tracto,
            'fecha': datos_caseta.get('fecha', ''),
            'hora': datos_caseta.get('hora', ''),
            'importe': datos_caseta.get('importe', 0),
            'nombre': datos_caseta.get('nombre', ''),
            'clase': datos_caseta.get('clase', ''),
            'tarjeta': datos_caseta.get('tarjeta', ''),
            'datetime': fecha_cas
        })
    
    # Buscar pares con diferencia <= diferencia_minutos
    pares_encontrados = []
    total_tractos_procesados = 0
    
    for tracto, casetas in casetas_por_tracto.items():
        # Ordenar por fecha/hora
        casetas.sort(key=lambda x: x['datetime'])
        
        # Comparar casetas consecutivas
        for i in range(len(casetas) - 1):
            caseta_actual = casetas[i]
            caseta_siguiente = casetas[i + 1]
            
            # Calcular diferencia en minutos
            diferencia = caseta_siguiente['datetime'] - caseta_actual['datetime']
            diferencia_mins = diferencia.total_seconds() / 60
            
            # Si la diferencia es <= umbral, agregar el par
            if 0 < diferencia_mins <= diferencia_minutos:
                pares_encontrados.append({
                    'caseta1': caseta_actual,
                    'caseta2': caseta_siguiente,
                    'diferencia_minutos': diferencia_mins
                })
        
        total_tractos_procesados += 1
    
    print(f"[DETECCIÓN] Tractos procesados: {total_tractos_procesados}")
    print(f"[DETECCIÓN] Pares sospechosos encontrados: {len(pares_encontrados)}")
    
    if not pares_encontrados:
        print(f"\nNo se encontraron casetas duplicadas con diferencia <= {diferencia_minutos} minutos.\n")
        return []
    
    # Formatear para impresión usando imprimir_lista_casetas
    registros_para_imprimir = []
    
    print(f"\n{'='*120}")
    print(f"CASETAS DUPLICADAS - Diferencia <= {diferencia_minutos} minutos")
    print(f"{'='*120}\n")
    
    for par in pares_encontrados:
        caseta1 = par['caseta1']
        caseta2 = par['caseta2']
        diff_mins = par['diferencia_minutos']
        
        # Agregar primera caseta del par
        registros_para_imprimir.append({
            'tracto': caseta1['tracto'],
            'fecha': caseta1['fecha'],
            'hora': caseta1['hora'],
            'importe': caseta1['importe'],
            'nombre': caseta1['nombre'],
            'clase': caseta1['clase'],
            'id_combinado': caseta1['id_caseta'],
            'tarjeta': caseta1['tarjeta'],
            'status': f'Par #{len(registros_para_imprimir)//2 + 1}'
        })
        
        # Agregar segunda caseta del par
        registros_para_imprimir.append({
            'tracto': caseta2['tracto'],
            'fecha': caseta2['fecha'],
            'hora': caseta2['hora'],
            'importe': caseta2['importe'],
            'nombre': caseta2['nombre'],
            'clase': caseta2['clase'],
            'id_combinado': caseta2['id_caseta'],
            'tarjeta': caseta2['tarjeta'],
            'status': f'Δ {diff_mins:.1f}min'
        })
    
    # Usar la función existente para imprimir
    imprimir_lista_casetas(registros_para_imprimir)
    
    print(f"\nTotal de pares encontrados: {len(pares_encontrados)}")
    print(f"Nota: La columna 'Nota' muestra el número de par y la diferencia en minutos (Δ)\n")
    
    return pares_encontrados


def detectar_casetas_duplicadas_por_tracto(db_manager, tracto, diferencia_minutos=5):
    """
    Wrapper para buscar casetas duplicadas de un tracto específico.
    
    Parámetros:
    - db_manager: instancia de DatabaseManager
    - tracto: número de tracto a buscar
    - diferencia_minutos: umbral de minutos (default: 5)
    """
    return detectar_casetas_duplicadas_db(db_manager, diferencia_minutos, tracto_filter=tracto)


##############################Funciones de Base de Datos ###################################################################################################

def merge_single_sheet(existing_sheet, new_sheet, header_rows=1, id_col=2):
    """
    Combina una sola hoja de datos siguiendo lógica: REEMPLAZAR + MANTENER
    
    Parámetros:
    - existing_sheet: lista de filas existentes
    - new_sheet: lista de filas nuevas
    - header_rows: número de filas de encabezado
    - id_col: índice de la columna que contiene el ID
    
    Retorna: lista de filas combinadas
    """
    if not existing_sheet or len(existing_sheet) <= header_rows:
        return new_sheet if new_sheet else []
    
    if not new_sheet or len(new_sheet) <= header_rows:
        return existing_sheet
    
    # Copiar encabezado(s)
    combined = []
    if header_rows > 0:
        combined.extend(existing_sheet[:header_rows])
    
    # Construir diccionario de filas por ID
    filas_por_id = {}
    
    # Agregar filas existentes
    for row in existing_sheet[header_rows:]:
        try:
            row_id = row[id_col] if id_col < len(row) else None
            if row_id is not None and not pd.isna(row_id):
                row_id_str = str(row_id).strip()
                # Filtrar IDs que comiencen con HID y no con HIDVCHID
                if row_id_str.startswith('HID') and not row_id_str.startswith('HIDVCHID'):
                    filas_por_id[row_id_str] = row
        except Exception:
            continue
    
    # Contadores
    nuevos = 0
    actualizados = 0
    
    # Procesar filas nuevas (reemplazar o agregar)
    for row in new_sheet[header_rows:]:
        try:
            row_id = row[id_col] if id_col < len(row) else None
            if row_id is None or pd.isna(row_id):
                continue
            
            row_id_str = str(row_id).strip()
            
            # Filtrar IDs
            if not row_id_str.startswith('HID') or row_id_str.startswith('HIDVCHID'):
                continue
            
            if row_id_str in filas_por_id:
                # Reemplazar
                filas_por_id[row_id_str] = row
                actualizados += 1
            else:
                # Agregar
                filas_por_id[row_id_str] = row
                nuevos += 1
                
        except Exception:
            continue
    
    # Reconstruir lista
    combined.extend(filas_por_id.values())
    
    print(f"  Hoja: {nuevos} nuevos, {actualizados} actualizados, {len(filas_por_id)} total")
    return combined


def merge_single_sheet(existing_sheet, new_sheet, header_rows=1, id_col=2):
    """
    Combina una sola hoja de datos basándose en IDs únicos.
    
    Parámetros:
    - existing_sheet: lista de filas existentes
    - new_sheet: lista de filas nuevas
    - header_rows: número de filas de encabezado
    - id_col: índice de la columna que contiene el ID
    
    Retorna: lista de filas combinadas
    
    Ubicación: Agregar después de merge_gastos_data
    """
    
    if not existing_sheet or len(existing_sheet) <= header_rows:
        # No hay datos previos: devolver new_sheet tal cual
        return new_sheet
    
    # Copiar encabezado(s)
    combined = []
    if header_rows > 0:
        combined.extend(existing_sheet[:header_rows])
    
    # Construir diccionario de filas existentes por ID
    # Esto permite actualizar registros existentes y agregar nuevos
    existing_by_id = {}
    for row in existing_sheet[header_rows:]:
        try:
            row_id = row[id_col] if id_col < len(row) else None
            if row_id is not None and not pd.isna(row_id):
                row_id_str = str(row_id).strip()
                # Filtrar IDs que comiencen con HID y no con HIDVCHID
                if row_id_str.startswith('HID') and not row_id_str.startswith('HIDVCHID'):
                    existing_by_id[row_id_str] = row
        except Exception:
            continue
    
    # Procesar filas nuevas: actualizar existentes o agregar nuevas
    for row in new_sheet[header_rows:]:
        try:
            row_id = row[id_col] if id_col < len(row) else None
            if row_id is None or pd.isna(row_id):
                continue
            
            row_id_str = str(row_id).strip()
            
            # Filtrar IDs que comiencen con HID y no con HIDVCHID
            if not row_id_str.startswith('HID') or row_id_str.startswith('HIDVCHID'):
                continue
            
            # Actualizar o agregar
            existing_by_id[row_id_str] = row
            
        except Exception:
            continue
    
    # Reconstruir lista de filas desde el diccionario
    combined.extend(existing_by_id.values())
    
    added = len(combined) - header_rows - (len(existing_sheet) - header_rows)
    print(f"  Hoja procesada: {len(existing_by_id)} registros únicos, {max(0, added)} nuevos")
    
    return combined


##############################Funciones version 10###################################################################################################
##############################Funciones version 10###################################################################################################
##############################Funciones version 10###################################################################################################

##############################Funciones de Base de Datos###################################################################################################

class DatabaseManager:
    """Gestor de la base de datos persistente"""
    
    def __init__(self, db_path='database.json'):
        self.db_path = Path(db_path)
        self.data = {
            'viajes': {},           # {id_viaje: {datos del viaje}}
            'casetas': {},          # {id_caseta: {datos de caseta}}
            'relaciones': [],       # [{id_viaje, id_caseta, is_fuera}]
            'rutas': {},            # {id_ruta_tipo: {registro completo}}
            'metadata': {
                'ultima_actualizacion': None,
                'version': '2.0'    # Incrementar versión
            }
        }
        self.cargar_datos()
    
    def cargar_datos(self):
        """Carga los datos desde el archivo JSON si existe"""
        if self.db_path.exists():
            try:
                with open(self.db_path, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
                print(f"[DB] Base de datos cargada: {self.db_path}")
                print(f"[DB] Viajes: {len(self.data['viajes'])}, Casetas: {len(self.data['casetas'])}, Relaciones: {len(self.data['relaciones'])}")
            except Exception as e:
                print(f"[DB ERROR] No se pudo cargar la base de datos: {e}")
                print("[DB] Iniciando con base de datos vacía")
        else:
            print("[DB] No existe base de datos previa. Creando nueva...")
    
    def guardar_datos(self):
        """Guarda los datos en el archivo JSON"""
        try:
            self.data['metadata']['ultima_actualizacion'] = datetime.now().isoformat()
            with open(self.db_path, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
            print(f"[DB] Datos guardados exitosamente en {self.db_path}")
            return True
        except Exception as e:
            print(f"[DB ERROR] No se pudieron guardar los datos: {e}")
            return False
    
    def agregar_viaje(self, id_viaje, datos_viaje):
        """
        Agrega o actualiza un viaje en la base de datos.
        Lógica: REEMPLAZAR + MANTENER
        - Si el viaje existe: se reemplaza completamente con los nuevos datos
        - Si no existe: se agrega como nuevo
        
        Parámetros:
        - id_viaje: ID único del viaje
        - datos_viaje: dict con todos los datos del viaje
        """
        if id_viaje in self.data['viajes']:
            # REEMPLAZAR: actualizar completamente con nuevos datos
            self.data['viajes'][id_viaje] = datos_viaje
        else:
            # AGREGAR: nuevo viaje
            self.data['viajes'][id_viaje] = datos_viaje
    
    def agregar_caseta(self, id_caseta, datos_caseta):
        """
        Agrega o actualiza una caseta en la base de datos.
        Lógica: REEMPLAZAR + MANTENER
        - Si la caseta existe: se reemplaza completamente con los nuevos datos
        - Si no existe: se agrega como nueva
        
        Parámetros:
        - id_caseta: ID único de la caseta (ID_COMBINADO)
        - datos_caseta: dict con todos los datos de la caseta
        """
        if id_caseta in self.data['casetas']:
            # REEMPLAZAR: actualizar completamente
            self.data['casetas'][id_caseta] = datos_caseta
        else:
            # AGREGAR: nueva caseta
            self.data['casetas'][id_caseta] = datos_caseta
    
    def agregar_relacion(self, id_viaje, id_caseta, is_fuera=False):
        """
        Agrega una relación viaje-caseta.
        Evita duplicados.
        """
        relacion = {
            'id_viaje': id_viaje,
            'id_caseta': id_caseta,
            'is_fuera': is_fuera
        }
        
        # Verificar si ya existe esta relación
        existe = any(
            r['id_viaje'] == id_viaje and r['id_caseta'] == id_caseta
            for r in self.data['relaciones']
        )
        
        if not existe:
            self.data['relaciones'].append(relacion)
    
    def enriquecer_viajes_desde_cierre(self, dataayc, ayc_header_rows=4, col_id_viaje=1):
        """
        Enriquece los viajes existentes con datos del excel de cierre.
        Lógica: REEMPLAZAR + MANTENER
        - Campos nuevos: se agregan
        - Campos existentes: se reemplazan con nuevos datos
        
        Agrega campos: clave, ensamble, frontal, dolly, trasero, id_ruta
        """
        if not dataayc or len(dataayc) <= ayc_header_rows:
            print("[DB] No hay datos de cierre para procesar")
            return 0
        
        info_by_viaje, procesados = build_info_from_dataayc(dataayc, ayc_header_rows, col_id_viaje)
        
        actualizados = 0
        rutas_creadas = 0
        nuevos = 0
        
        for id_viaje, info in info_by_viaje.items():
            if id_viaje in self.data['viajes']:
                viaje = self.data['viajes'][id_viaje]
                
                # REEMPLAZAR: actualizar todos los campos con nuevos datos
                if info.get('clave'):
                    viaje['clave'] = info['clave']
                
                if info.get('colL'):
                    viaje['ensamble'] = info['colL']
                
                if info.get('frontal'):
                    viaje['frontal'] = info['frontal']
                
                if info.get('dolly'):
                    viaje['dolly'] = info['dolly']
                
                if info.get('trasero'):
                    viaje['trasero'] = info['trasero']
                
                # Crear/actualizar ID_RUTA
                clave = info.get('clave')
                nombre_ruta = viaje.get('ruta')
                
                if clave and nombre_ruta:
                    id_ruta = crear_id_ruta(clave, nombre_ruta)
                    if id_ruta:
                        viaje['id_ruta'] = id_ruta
                        rutas_creadas += 1
                
                actualizados += 1
            else:
                # AGREGAR: viaje no existe en monitoreo, crear entrada parcial
                nuevo_viaje = {
                    'clave': info.get('clave'),
                    'ensamble': info.get('colL'),
                    'frontal': info.get('frontal'),
                    'dolly': info.get('dolly'),
                    'trasero': info.get('trasero'),
                    'advertencias': [],
                    'origen': 'cierre_solamente'
                }
                
                # Intentar crear ID_RUTA si es posible
                clave = info.get('clave')
                # No tenemos nombre_ruta aquí, así que no podemos crear id_ruta
                
                self.data['viajes'][id_viaje] = nuevo_viaje
                nuevos += 1
        
        print(f"[DB] Viajes actualizados: {actualizados}, nuevos: {nuevos}")
        print(f"[DB] ID_RUTA creados/actualizados: {rutas_creadas}")
        return actualizados + nuevos
    
    def enriquecer_viajes_desde_gastos(self, datagastos):
        """
        Enriquece los viajes existentes con datos del excel de gastos.
        Lógica: REEMPLAZAR + MANTENER
        - Si el viaje existe: reemplaza los campos gasto y comisión
        - Si no existe: crea entrada parcial
        
        datagastos debe ser un dict con hojas 'GASTOS' y 'CIERRE 35'
        Opcionalmente puede incluir 'MODIFICACIONES_COMISION'
        Agrega campos: gasto, comision, comision_ajustada (si aplica)
        """
        actualizados = 0
        nuevos = 0
        
        # Procesar hoja GASTOS
        if 'GASTOS' in datagastos:
            hoja_gastos = datagastos['GASTOS']
            for row in hoja_gastos[1:]:  # Saltar header
                try:
                    id_viaje = row[2] if len(row) > 2 else None
                    gasto = row[5] if len(row) > 5 else None
                    
                    if id_viaje is None or pd.isna(id_viaje):
                        continue
                    
                    id_viaje = str(id_viaje).strip()
                    
                    if not id_viaje.startswith('HID') or id_viaje.startswith('HIDVCHID'):
                        continue
                    
                    if id_viaje in self.data['viajes']:
                        # REEMPLAZAR: actualizar gasto
                        if gasto is not None and not pd.isna(gasto):
                            self.data['viajes'][id_viaje]['gasto'] = float(gasto)
                            actualizados += 1
                    else:
                        # AGREGAR: crear entrada parcial
                        if gasto is not None and not pd.isna(gasto):
                            self.data['viajes'][id_viaje] = {
                                'gasto': float(gasto),
                                'advertencias': [],
                                'origen': 'gastos_gastos'
                            }
                            nuevos += 1
                except Exception:
                    continue
        
        # NUEVO: Obtener tabla de modificaciones de comisiones
        modificaciones_comision = datagastos.get('MODIFICACIONES_COMISION', {})
        
        # Procesar hoja CIERRE 35
        if 'CIERRE 35' in datagastos:
            hoja_cierre35 = datagastos['CIERRE 35']
            for row in hoja_cierre35[1:]:  # Saltar header
                try:
                    id_viaje = row[2] if len(row) > 2 else None
                    comision = row[9] if len(row) > 9 else None
                    
                    if id_viaje is None or pd.isna(id_viaje):
                        continue
                    
                    id_viaje = str(id_viaje).strip()
                    
                    if not id_viaje.startswith('HID') or id_viaje.startswith('HIDVCHID'):
                        continue
                    
                    # NUEVO: Buscar si hay modificación de comisión para este viaje
                    comision_ajustada = modificaciones_comision.get(id_viaje)
                    
                    if id_viaje in self.data['viajes']:
                        # REEMPLAZAR: actualizar comisión
                        if comision is not None and not pd.isna(comision):
                            self.data['viajes'][id_viaje]['comision'] = float(comision)
                            
                            # NUEVO: Si hay comisión ajustada, guardarla también
                            if comision_ajustada is not None:
                                self.data['viajes'][id_viaje]['comision_ajustada'] = float(comision_ajustada)
                            
                            actualizados += 1
                    else:
                        # AGREGAR: crear entrada parcial
                        if comision is not None and not pd.isna(comision):
                            viaje_data = {
                                'comision': float(comision),
                                'advertencias': [],
                                'origen': 'gastos_comm'
                            }
                            
                            # NUEVO: Si hay comisión ajustada, incluirla
                            if comision_ajustada is not None:
                                viaje_data['comision_ajustada'] = float(comision_ajustada)
                            
                            self.data['viajes'][id_viaje] = viaje_data
                            nuevos += 1
                except Exception:
                    continue
        
        print(f"[DB] Viajes actualizados con gastos: {actualizados}, nuevos: {nuevos}")
        if modificaciones_comision:
            print(f"[DB] Viajes con comisión ajustada: {len(modificaciones_comision)}")
        
        return actualizados + nuevos
        
    
    def obtener_viaje(self, id_viaje):
        """Obtiene los datos de un viaje"""
        return self.data['viajes'].get(id_viaje)
    
    def obtener_caseta(self, id_caseta):
        """Obtiene los datos de una caseta"""
        return self.data['casetas'].get(id_caseta)
    
    def obtener_casetas_por_viaje(self, id_viaje, solo_fuera=None):
        """
        Obtiene las casetas asociadas a un viaje.
        solo_fuera: None (todas), True (solo fuera), False (solo dentro)
        """
        casetas = []
        for rel in self.data['relaciones']:
            if rel['id_viaje'] == id_viaje:
                if solo_fuera is None or rel['is_fuera'] == solo_fuera:
                    caseta_data = self.data['casetas'].get(rel['id_caseta'])
                    if caseta_data:
                        caseta_completa = caseta_data.copy()
                        caseta_completa['id_caseta'] = rel['id_caseta']
                        caseta_completa['is_fuera'] = rel['is_fuera']
                        casetas.append(caseta_completa)
        return casetas
    
    def obtener_todos_los_viajes(self):
        """Obtiene todos los viajes con estadísticas de casetas"""
        viajes_completos = []
        
        for id_viaje, datos_viaje in self.data['viajes'].items():
            # Contar casetas dentro y fuera
            casetas_dentro = sum(1 for r in self.data['relaciones'] 
                                if r['id_viaje'] == id_viaje and not r['is_fuera'])
            casetas_fuera = sum(1 for r in self.data['relaciones'] 
                               if r['id_viaje'] == id_viaje and r['is_fuera'])
            
            # Calcular importes
            importe_dentro = 0.0
            importe_fuera = 0.0
            for r in self.data['relaciones']:
                if r['id_viaje'] == id_viaje:
                    caseta = self.data['casetas'].get(r['id_caseta'])
                    if caseta and caseta.get('importe'):
                        try:
                            imp = float(caseta['importe'])
                            if r['is_fuera']:
                                importe_fuera += imp
                            else:
                                importe_dentro += imp
                        except:
                            pass
            
            viaje_completo = datos_viaje.copy()
            viaje_completo['id_viaje'] = id_viaje
            viaje_completo['num_casetas'] = casetas_dentro
            viaje_completo['num_casetas_fuera'] = casetas_fuera
            viaje_completo['importe_total'] = importe_dentro
            viaje_completo['importe_fuera'] = importe_fuera
            
            viajes_completos.append(viaje_completo)
        
        return viajes_completos
    
    def enriquecer_viajes_desde_adicionales(self, dataadicional, header_rows=1):
        """
        Enriquece los viajes con información de gastos adicionales.
        
        Agrega al viaje:
        - gastos_adicionales_clasificacion: lista de clasificaciones
        - gastos_adicionales_motivo: lista de motivos
        - gastos_adicionales_importe: lista de importes
        - gastos_adicionales_total: suma total de importes
        
        Parámetros:
        - dataadicional: datos del excel de gastos adicionales
        - header_rows: filas de encabezado
        """
        if not dataadicional or len(dataadicional) <= header_rows:
            print("[DB] No hay datos de gastos adicionales para enriquecer")
            return
        
        col_id_viaje = 3      # Columna D
        col_clasificacion = 15 # Columna P
        col_motivo = 16       # Columna Q
        col_importe = 10      # Columna K
        
        # Agrupar gastos adicionales por ID de viaje
        gastos_por_viaje = {}
        
        for row in dataadicional[header_rows:]:
            if len(row) <= max(col_id_viaje, col_clasificacion, col_motivo, col_importe):
                continue
            
            id_viaje = row[col_id_viaje]
            if pd.isna(id_viaje):
                continue
            
            id_viaje = str(id_viaje).strip()
            
            # Verificar que el viaje existe en la BD
            if id_viaje not in self.data['viajes']:
                continue
            
            # Extraer datos
            clasificacion = str(row[col_clasificacion]) if not pd.isna(row[col_clasificacion]) else ""
            motivo = str(row[col_motivo]) if not pd.isna(row[col_motivo]) else ""
            
            try:
                importe = float(row[col_importe]) if not pd.isna(row[col_importe]) else 0.0
            except:
                importe = 0.0
            
            # Agrupar por viaje
            if id_viaje not in gastos_por_viaje:
                gastos_por_viaje[id_viaje] = {
                    'clasificaciones': [],
                    'motivos': [],
                    'importes': []
                }
            
            gastos_por_viaje[id_viaje]['clasificaciones'].append(clasificacion)
            gastos_por_viaje[id_viaje]['motivos'].append(motivo)
            gastos_por_viaje[id_viaje]['importes'].append(importe)
        
        # Enriquecer viajes
        viajes_enriquecidos = 0
        
        for id_viaje, gastos in gastos_por_viaje.items():
            viaje = self.data['viajes'][id_viaje]
            
            # Agregar listas de gastos adicionales
            viaje['gastos_adicionales_clasificacion'] = gastos['clasificaciones']
            viaje['gastos_adicionales_motivo'] = gastos['motivos']
            viaje['gastos_adicionales_importe'] = gastos['importes']
            
            # Calcular total
            viaje['gastos_adicionales_total'] = sum(gastos['importes'])
            
            viajes_enriquecidos += 1
        
        print(f"[DB] Viajes enriquecidos con gastos adicionales: {viajes_enriquecidos}")
        
    # ------------------------------Funciones de las relaciones de rutas------------------------------------------
    
    def cargar_rutas(self, tabla_rutas):
        """
        Carga la tabla de rutas en la base de datos haciendo merge.
        Actualiza registros existentes y agrega nuevos.
        
        Parámetros:
        - tabla_rutas: dict {clave_registro: {datos}}
        """
        if not tabla_rutas:
            print("[DB] No hay rutas para cargar")
            return 0
        
        nuevos = 0
        actualizados = 0
        
        # Merge: actualizar existentes y agregar nuevos
        for clave, datos in tabla_rutas.items():
            if clave in self.data['rutas']:
                # Actualizar existente
                self.data['rutas'][clave] = datos
                actualizados += 1
            else:
                # Agregar nuevo
                self.data['rutas'][clave] = datos
                nuevos += 1
        
        print(f"[DB] Rutas cargadas: {nuevos} nuevas, {actualizados} actualizadas, {len(self.data['rutas'])} total")
        
        # IMPORTANTE: Guardar inmediatamente después de cargar
        self.guardar_datos()
        
        return nuevos + actualizados
    
    def obtener_ruta(self, id_ruta, tipo_viaje=None):
        """
        Obtiene registro(s) de ruta por ID_RUTA.
        
        Parámetros:
        - id_ruta: ID de la ruta
        - tipo_viaje: si se especifica, busca registro específico
        
        Retorna: dict o lista de dicts
        """
        if tipo_viaje:
            clave = f"{id_ruta}-{tipo_viaje}"
            return self.data['rutas'].get(clave)
        else:
            # Retornar todos los registros que empiecen con id_ruta
            return [v for k, v in self.data['rutas'].items() if k.startswith(id_ruta)]
    
    def obtener_todas_rutas(self):
        """Retorna todas las rutas"""
        return list(self.data['rutas'].values())

def detectar_tipo_excel_multi_hoja(excel_path):
    """
    Detecta si un excel es de tipo 'gastos' verificando nombres de hojas.
    Retorna: 'gastos' o None
    """
    try:
        xls = pd.ExcelFile(excel_path)
        sheet_names = [name.upper() for name in xls.sheet_names]
        all_sheet_names = xls.sheet_names
        
        # Verificar si tiene hojas GASTOS y CIERRE 35
        if 'GASTOS' in sheet_names and 'CIERRE 35' in sheet_names:
            return 'gastos'
        # Verificar si tiene hoja iniciando con BD y esta es de gastos adicionales
        for sheet_name in all_sheet_names:
            if sheet_name.upper().startswith('BD'):
                # Leer la hoja
                df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, engine='openpyxl')
                data = df.values.tolist()
                
                # Verificar celda D1 (índice [0][3])
                if len(data) > 0 and len(data[0]) > 3:
                    celda_d1 = str(data[0][3]).strip().upper()
                    
                    if celda_d1 == "CLAVE DEL VIAJE":
                        print(f"[DETECCIÓN] Gastos adicionales detectado (hoja: {sheet_name})")
                        return 'gastos_adicionales'
    except Exception:
        pass
    
    return None

def construir_base_datos_desde_excels(datamon, datacas, dataayc, datagastos,
                                       mon_header_rows=5, cas_header_rows=1,
                                       db_manager=None, is_update=False,
                                       dataadicional=None):
    """
    Construye la base de datos a partir de los excels cargados.
    Reemplaza la función crear_tabla_relacional_viaje_caseta.
    
    Parámetros:
    - is_update: Si True, permite que datamon y datacas sean None/vacíos (modo actualización)
                 Si False, datamon y datacas son obligatorios (modo creación inicial)
    
    Retorna: DatabaseManager con los datos cargados
    """
    if db_manager is None:
        db_manager = DatabaseManager()
    
    print("\n--- Construyendo base de datos desde excels ---")
    
    # NUEVA VALIDACIÓN: En modo creación inicial, verificar datos obligatorios
    if not is_update:
        if not datamon or len(datamon) <= mon_header_rows:
            print("ERROR: Se requiere excel de monitoreo para crear la base de datos inicial")
            return db_manager
        if not datacas or len(datacas) <= cas_header_rows:
            print("ERROR: Se requiere excel de casetas para crear la base de datos inicial")
            return db_manager
    
    # NUEVA VALIDACIÓN: Convertir listas vacías a None para facilitar verificaciones
    if datamon is not None and len(datamon) <= mon_header_rows:
        datamon = None
    if datacas is not None and len(datacas) <= cas_header_rows:
        datacas = None
    
    # Índices de columnas para monitoreo
    mon_id_viaje_col = 0
    mon_id_tracto_col = 8
    mon_salida_col = 14
    mon_llegada_col = 16
    mon_ruta_col = 3
    mon_operador_col = 6
    
    # Índices de columnas para casetas
    cas_tracto_col = 8
    cas_fecha_col = 9
    cas_hora_col = 10
    cas_importe_col = 14
    cas_nombre_col = 12
    cas_clase_col = 11
    cas_tarjeta_col = 5
    
    # NUEVA VALIDACIÓN: Solo buscar ID_COMBINADO si hay datos de casetas
    cas_id_col = None
    if datacas is not None:
        cas_header = datacas[0]
        try:
            cas_id_col = cas_header.index('ID_COMBINADO')
        except ValueError:
            print("ERROR: No se encontró columna 'ID_COMBINADO' en casetas")
            return db_manager
    
    # PASO 1: Cargar todos los viajes desde monitoreo (SOLO SI HAY DATOS)
    if datamon is not None:
        print("[DB] Cargando viajes desde monitoreo...")
        viajes_procesados = 0
        viajes_nuevos = 0
        
        total_viajes = max(0, len(datamon) - mon_header_rows)
        progreso_interval = max(1, total_viajes // 100)
        
        for idx, viaje_row in enumerate(datamon[mon_header_rows:], start=1):
            if idx % progreso_interval == 0 or idx == 1 or idx == total_viajes:
                print(f"[PROGRESO] Cargando viajes: {idx} de {total_viajes}", end='\r', flush=True)
            
            id_viaje = viaje_row[mon_id_viaje_col]
            if pd.isna(id_viaje):
                continue
            
            id_viaje = str(id_viaje).strip()
            tracto_mon = extraer_numero_tracto(viaje_row[mon_id_tracto_col])
            fecha_salida = parse_datetime_monitoreo(viaje_row[mon_salida_col])
            fecha_llegada = parse_datetime_monitoreo(viaje_row[mon_llegada_col])
            ruta = viaje_row[mon_ruta_col] if len(viaje_row) > mon_ruta_col and not pd.isna(viaje_row[mon_ruta_col]) else None
            operador = viaje_row[mon_operador_col] if len(viaje_row) > mon_operador_col and not pd.isna(viaje_row[mon_operador_col]) else None
            
            if fecha_salida is None or fecha_llegada is None:
                continue
            
            # Inicializar advertencias
            advertencias = []
            
            # Verificar si es irreal (duración < 30 minutos)
            try:
                duracion = fecha_llegada - fecha_salida
                if pd.isna(duracion) or duracion < timedelta(minutes=30):
                    advertencias.append('IRREAL')
            except:
                advertencias.append('IRREAL')
            
            # Convertir fechas a string ISO para JSON
            datos_viaje = {
                'fecha_inicio': fecha_salida.isoformat() if fecha_salida else None,
                'fecha_final': fecha_llegada.isoformat() if fecha_llegada else None,
                'ruta': str(ruta) if ruta else None,
                'operador': str(operador) if operador else None,
                'tracto': str(tracto_mon) if tracto_mon else None,
                'advertencias': advertencias,  
                'origen': 'monitoreo'
            }
            
            if id_viaje not in db_manager.data['viajes']:
                viajes_nuevos += 1
            
            db_manager.agregar_viaje(id_viaje, datos_viaje)
            viajes_procesados += 1
        
        print(f"\n[DB] Viajes procesados: {viajes_procesados}, nuevos: {viajes_nuevos}")
    else:
        print("[DB] No hay datos de monitoreo para procesar")
    
    # PASO 2: Cargar todas las casetas (SOLO SI HAY DATOS)
    if datacas is not None and cas_id_col is not None:
        print("[DB] Cargando casetas...")
        casetas_procesadas = 0
        casetas_nuevas = 0
        
        for cas_row in datacas[cas_header_rows:]:
            id_caseta = cas_row[cas_id_col]
            if pd.isna(id_caseta):
                continue
            
            id_caseta = str(id_caseta).strip()
            
            tracto = cas_row[cas_tracto_col] if cas_tracto_col < len(cas_row) else None
            fecha = cas_row[cas_fecha_col] if cas_fecha_col < len(cas_row) else None
            hora = cas_row[cas_hora_col] if cas_hora_col < len(cas_row) else None
            importe = cas_row[cas_importe_col] if cas_importe_col < len(cas_row) else None
            nombre = cas_row[cas_nombre_col] if cas_nombre_col < len(cas_row) else None
            clase = cas_row[cas_clase_col] if cas_clase_col < len(cas_row) else None
            tarjeta = cas_row[cas_tarjeta_col] if cas_tarjeta_col < len(cas_row) else None
            
            # Formatear fecha y hora para almacenamiento
            fecha_str = ''
            hora_str = ''
            try:
                if fecha is not None and not pd.isna(fecha):
                    fecha_str = str(fecha).split()[0]
            except:
                fecha_str = str(fecha) if fecha else ''
            
            try:
                if hora is not None and not pd.isna(hora):
                    hora_str = str(hora)
            except:
                hora_str = str(hora) if hora else ''
            
            datos_caseta = {
                'tracto': str(tracto) if tracto and not pd.isna(tracto) else None,
                'fecha': fecha_str,
                'hora': hora_str,
                'importe': float(importe) if importe and not pd.isna(importe) else 0.0,
                'nombre': str(nombre) if nombre and not pd.isna(nombre) else None,
                'clase': str(clase) if clase and not pd.isna(clase) else None,
                'tarjeta': str(tarjeta) if tarjeta and not pd.isna(tarjeta) else None
            }
            
            if id_caseta not in db_manager.data['casetas']:
                casetas_nuevas += 1
            
            db_manager.agregar_caseta(id_caseta, datos_caseta)
            casetas_procesadas += 1
        
        print(f"[DB] Casetas procesadas: {casetas_procesadas}, nuevas: {casetas_nuevas}")
    else:
        print("[DB] No hay datos de casetas para procesar")
    
    # PASO 3: Crear relaciones viaje-caseta OPTIMIZADO
    if len(db_manager.data['viajes']) > 0 and len(db_manager.data['casetas']) > 0:
        print("[DB] Procesando relaciones viaje-caseta (optimizado)...")
        import time
        tiempo_inicio = time.time()
        
        # Construir índices de casetas
        indices = IndicesCasetas()
        indices.construir_desde_db(db_manager.data['casetas'])
        
        # Sistema de tracking de cambios
        tracking = TrackingCambios()
        tracking.cargar_estado()
        
        # Detectar qué viajes necesitan procesarse
        viajes_a_procesar = tracking.detectar_cambios(db_manager.data['viajes'])
        
        if is_update and viajes_a_procesar:
            print(f"[DB] Modo incremental: procesando {len(viajes_a_procesar)} viajes modificados/nuevos")
        else:
            print(f"[DB] Modo completo: procesando {len(db_manager.data['viajes'])} viajes")
            viajes_a_procesar = set(db_manager.data['viajes'].keys())
        
        # Limpiar relaciones de viajes que se van a reprocesar
        if is_update and viajes_a_procesar:
            relaciones_nuevas = []
            for rel in db_manager.data['relaciones']:
                if rel['id_viaje'] not in viajes_a_procesar:
                    # Mantener relaciones de viajes no modificados
                    relaciones_nuevas.append(rel)
            db_manager.data['relaciones'] = relaciones_nuevas
            print(f"[DB] Relaciones preservadas: {len(relaciones_nuevas)}")
        else:
            # Modo completo: limpiar todas las relaciones
            db_manager.data['relaciones'] = []
        
        # Procesar viajes
        relaciones_creadas = 0
        viajes_sin_casetas = 0
        
        total_procesar = len(viajes_a_procesar)
        procesados = 0
        
        for id_viaje in viajes_a_procesar:
            procesados += 1
            if procesados % max(1, total_procesar // 20) == 0 or procesados == 1:
                print(f"[PROGRESO] {procesados}/{total_procesar} viajes", end='\r', flush=True)
            
            datos_viaje = db_manager.data['viajes'][id_viaje]
            tracto = datos_viaje.get('tracto')
            
            if not tracto:
                continue
            
            try:
                fecha_inicio = pd.to_datetime(datos_viaje['fecha_inicio'])
                fecha_final = pd.to_datetime(datos_viaje['fecha_final'])
            except:
                continue
            
            # BÚSQUEDA OPTIMIZADA usando índices
            ids_casetas = indices.buscar_por_rango(tracto, fecha_inicio, fecha_final)
            
            if ids_casetas:
                for id_caseta in ids_casetas:
                    db_manager.data['relaciones'].append({
                        'id_viaje': id_viaje,
                        'id_caseta': id_caseta,
                        'is_fuera': False
                    })
                    relaciones_creadas += 1
            else:
                viajes_sin_casetas += 1
        
        # Guardar estado del tracking
        tracking.guardar_estado()
        
        tiempo_total = time.time() - tiempo_inicio
        print(f"\n[DB] Relaciones creadas: {relaciones_creadas}")
        print(f"[DB] Total relaciones: {len(db_manager.data['relaciones'])}")
        print(f"[DB] Viajes sin casetas: {viajes_sin_casetas}")
        print(f"[DB] Tiempo: {tiempo_total:.2f}s")
    else:
        print("[DB] No hay viajes/casetas para crear relaciones")
    
    # PASO 4: Agregar casetas FUERA de viaje (OPTIMIZADO)
    if datacas is not None and len(db_manager.data['viajes']) > 0:
        casetas_fuera = agregar_casetas_fuera_db(db_manager, datacas, cas_header_rows)
    else:
        print("[DB] Saltando casetas fuera de viaje (no hay datos)")
    
    # PASO 5: Enriquecer con datos de cierre
    if dataayc and dataayc != 0 and len(dataayc) > 0:
        print("[DB] Enriqueciendo viajes con datos de cierre...")
        db_manager.enriquecer_viajes_desde_cierre(dataayc)
    
    # PASO 6: Enriquecer con datos de gastos
    if datagastos and datagastos != 0:
        print("[DB] Enriqueciendo viajes con datos de gastos...")
        db_manager.enriquecer_viajes_desde_gastos(datagastos)
        
    # NUEVO PASO 6.3: Enriquecer con gastos adicionales
    if dataadicional and len(dataadicional) > 1:
        print("[DB] Enriqueciendo viajes con gastos adicionales...")
        db_manager.enriquecer_viajes_desde_adicionales(dataadicional)
    
        # PASO 6.5: Verificar comisiones vs rutas
    # Solo si hay rutas cargadas en la base de datos
    if len(db_manager.data.get('rutas', {})) > 0:
        verificar_comisiones_vs_rutas(db_manager)

    # PASO 7: Guardar la base de datos
    print("[DB] Guardando base de datos...")
    db_manager.guardar_datos()
    
    print("\n--- Base de datos construida exitosamente ---")
    return db_manager


def agregar_casetas_fuera_db(db_manager, datacas, cas_header_rows=1):
    """
    Busca casetas entre viajes consecutivos del mismo tracto (optimizado).
    Usa índices para búsqueda rápida.
    """
    if not datacas:
        return 0
    
    import time
    tiempo_inicio = time.time()
    
    print("[DB] Procesando casetas fuera de viaje (optimizado)...")
    
    # Construir índices si no existen
    indices = IndicesCasetas()
    indices.construir_desde_db(db_manager.data['casetas'])
    
    # Agrupar viajes por tracto y ordenar por fecha
    viajes_por_tracto = {}
    
    for id_viaje, datos_viaje in db_manager.data['viajes'].items():
        tracto = datos_viaje.get('tracto')
        if not tracto:
            continue
        
        try:
            fecha_inicio = pd.to_datetime(datos_viaje['fecha_inicio'])
            fecha_final = pd.to_datetime(datos_viaje['fecha_final'])
        except:
            continue
        
        if tracto not in viajes_por_tracto:
            viajes_por_tracto[tracto] = []
        
        viajes_por_tracto[tracto].append({
            'id_viaje': id_viaje,
            'fecha_inicio': fecha_inicio,
            'fecha_final': fecha_final
        })
    
    # Ordenar viajes de cada tracto por fecha
    for tracto in viajes_por_tracto:
        viajes_por_tracto[tracto].sort(key=lambda x: x['fecha_inicio'])
    
    casetas_fuera_agregadas = 0
    
    # Set de relaciones existentes para evitar duplicados
    relaciones_existentes = set()
    for rel in db_manager.data['relaciones']:
        relaciones_existentes.add((rel['id_viaje'], rel['id_caseta']))
    
    # Procesar cada tracto
    for tracto, viajes in viajes_por_tracto.items():
        # Buscar huecos entre viajes consecutivos
        for i in range(len(viajes) - 1):
            viaje_actual = viajes[i]
            viaje_siguiente = viajes[i + 1]
            
            fecha_final_actual = viaje_actual['fecha_final']
            fecha_inicio_siguiente = viaje_siguiente['fecha_inicio']
            
            # Solo procesar si hay un hueco temporal
            if fecha_final_actual >= fecha_inicio_siguiente:
                continue
            
            # BÚSQUEDA OPTIMIZADA de casetas en el hueco
            ids_casetas = indices.buscar_por_rango(
                tracto, 
                fecha_final_actual, 
                fecha_inicio_siguiente
            )
            
            for id_caseta in ids_casetas:
                # Verificar que no esté ya relacionada
                clave = (viaje_actual['id_viaje'], id_caseta)
                
                if clave not in relaciones_existentes:
                    db_manager.data['relaciones'].append({
                        'id_viaje': viaje_actual['id_viaje'],
                        'id_caseta': id_caseta,
                        'is_fuera': True
                    })
                    relaciones_existentes.add(clave)
                    casetas_fuera_agregadas += 1
    
    tiempo_total = time.time() - tiempo_inicio
    print(f"[DB] Casetas fuera agregadas: {casetas_fuera_agregadas}")
    print(f"[DB] Tiempo: {tiempo_total:.2f}s")
    
    return casetas_fuera_agregadas

##############################Funciones version 10 Filtros###################################################################################################
##############################Funciones version 10 Filtros###################################################################################################
##############################Funciones version 10 Filtros###################################################################################################


def aplicar_filtros_db(db_manager, filtros, single_use=False):
    """
    Aplica filtros sobre los viajes en la base de datos.
    
    Parámetros:
    - db_manager: instancia de DatabaseManager
    - filtros: dict con claves 'id', 'ruta', 'op', 'fecha', 'track'
    - single_use: si True, solo busca por ID y no modifica filtros persistentes
    
    Retorna: lista de viajes filtrados con estadísticas completas
    """
    
    # Obtener todos los viajes desde la base de datos
    todos_los_viajes = []
    for id_viaje, datos_viaje in db_manager.data['viajes'].items():
        viaje = datos_viaje.copy()
        viaje['id_viaje'] = id_viaje
        todos_los_viajes.append(viaje)
    
    # Filtrado por ID en modo single_use
    if 'id' in filtros and filtros['id'] is not None:
        id_buscar = str(filtros['id']).strip()
        encontrados = [v for v in todos_los_viajes if v['id_viaje'] == id_buscar]
        
        if single_use:
            if not encontrados:
                print(f"\nNo se encontró el viaje con ID '{id_buscar}'.\n")
                return []
            
            # Enriquecer con estadísticas de casetas
            viajes_enriquecidos = enriquecer_viajes_con_casetas(db_manager, encontrados)
            imprimir_filtros(viajes_enriquecidos)
            return viajes_enriquecidos
        else:
            todos_los_viajes = encontrados
    
    # Aplicar filtros acumulativos
    def cumple_filtros(viaje):
        # Filtro operador
        if 'op' in filtros and filtros['op'] is not None:
            op_f = str(filtros['op']).strip().lower()
            val = (viaje.get('operador') or '')
            if op_f not in str(val).lower():
                return False
        
        # Filtro tracto
        if 'track' in filtros and filtros['track'] is not None:
            track_f = str(filtros['track']).strip()
            val = (viaje.get('tracto') or '')
            if track_f not in str(val):
                return False
        
        # Filtro fecha (buscar en fecha_inicio o fecha_final)
        if 'fecha' in filtros and filtros['fecha'] is not None:
            fecha_f = filtros['fecha'].strip()
            fi = viaje.get('fecha_inicio')
            ff = viaje.get('fecha_final')
            if not (fecha_coincide_con_filtro(fi, fecha_f) or 
                    fecha_coincide_con_filtro(ff, fecha_f)):
                return False
        
        # Filtro ruta
        if 'ruta' in filtros and filtros['ruta'] is not None:
            ruta_f = str(filtros['ruta']).strip().lower()
            val_r = (viaje.get('ruta') or '')
            if ruta_f not in str(val_r).lower():
                return False
                
        # Filtro estado
        if 'estado' in filtros and filtros['estado'] is not None:
            estado_f = str(filtros['estado']).strip().upper()
            advertencias = viaje.get('advertencias', [])
            
            # Convertir advertencias a mayúsculas para comparación
            advertencias_upper = [a.upper() for a in advertencias]
            
            # Verificar si el estado buscado está en las advertencias
            if estado_f not in advertencias_upper:
                return False
        
        return True
    
    filtrados = [v for v in todos_los_viajes if cumple_filtros(v)]
    
    if not filtrados:
        print("\nNo se encontraron viajes que cumplan los filtros especificados.\n")
        return filtrados
    
    # Enriquecer con estadísticas de casetas
    viajes_enriquecidos = enriquecer_viajes_con_casetas(db_manager, filtrados)
    
    # Imprimir resultados
    imprimir_filtros(viajes_enriquecidos)
    
    return viajes_enriquecidos


def enriquecer_viajes_con_casetas(db_manager, viajes):
    """
    Enriquece una lista de viajes con estadísticas de casetas.
    
    Para cada viaje:
    1. Busca sus relaciones en la tabla relacional
    2. Obtiene los IDs de casetas
    3. Busca los datos de esas casetas
    4. Calcula estadísticas (número de casetas, importes)
    
    Parámetros:
    - db_manager: instancia de DatabaseManager
    - viajes: lista de dicts con datos de viajes
    
    Retorna: lista de viajes enriquecidos con campos adicionales
    """
    viajes_enriquecidos = []
    
    total_viajes = len(viajes)
    for idx, viaje in enumerate(viajes, start=1):
        # Mostrar progreso solo si hay muchos viajes
        if total_viajes > 100 and (idx % max(1, total_viajes // 10) == 0 or idx == 1 or idx == total_viajes):
            print(f"[PROGRESO] Enriqueciendo viajes: {idx} de {total_viajes}", end='\r', flush=True)
        
        id_viaje = viaje['id_viaje']
        
        # Inicializar contadores
        num_casetas_dentro = 0
        num_casetas_fuera = 0
        importe_dentro = 0.0
        importe_fuera = 0.0
        
        # Buscar relaciones de este viaje
        for relacion in db_manager.data['relaciones']:
            if relacion['id_viaje'] != id_viaje:
                continue
            
            id_caseta = relacion['id_caseta']
            is_fuera = relacion['is_fuera']
            
            # Obtener datos de la caseta
            caseta = db_manager.data['casetas'].get(id_caseta)
            if not caseta:
                continue
            
            # Obtener importe
            try:
                importe = float(caseta.get('importe', 0))
            except:
                importe = 0.0
            
            # Acumular según si está dentro o fuera del viaje
            if is_fuera:
                num_casetas_fuera += 1
                importe_fuera += importe
            else:
                num_casetas_dentro += 1
                importe_dentro += importe
        
        # Crear viaje enriquecido
        viaje_enriquecido = viaje.copy()
        viaje_enriquecido['num_casetas'] = num_casetas_dentro
        viaje_enriquecido['num_casetas_fuera'] = num_casetas_fuera
        viaje_enriquecido['importe_total'] = importe_dentro
        viaje_enriquecido['importe_fuera'] = importe_fuera
        
        # Normalizar nombres de campos para compatibilidad con imprimir_filtros
        # 'ensamble' es el nombre del campo que antes era 'colL'
        # Los campos frontal, dolly, trasero ya están con sus nombres correctos
        if 'is_irreal' in viaje_enriquecido:
            viaje_enriquecido['irreal'] = viaje_enriquecido['is_irreal']
        
        viajes_enriquecidos.append(viaje_enriquecido)
    
    if total_viajes > 100:
        print()  # Nueva línea después del progreso
    
    return viajes_enriquecidos


def listar_viajes_irreales_db(db_manager):
    """
    Lista todos los viajes marcados como irreales desde la base de datos.
    """
    # Obtener todos los viajes con advertencia IRREAL
    todos_los_viajes = []
    for id_viaje, datos_viaje in db_manager.data['viajes'].items():
        advertencias = datos_viaje.get('advertencias', [])
        
        if 'IRREAL' in advertencias:
            viaje = datos_viaje.copy()
            viaje['id_viaje'] = id_viaje
            todos_los_viajes.append(viaje)
    
    if not todos_los_viajes:
        print("\nNo se encontraron viajes irreales.\n")
        return []
    
    # Enriquecer con estadísticas de casetas
    viajes_enriquecidos = enriquecer_viajes_con_casetas(db_manager, todos_los_viajes)
    
    # Imprimir usando la función estándar
    print("\n=== VIAJES IRREALES ===")
    imprimir_filtros(viajes_enriquecidos)
    
    return viajes_enriquecidos
    
def listar_viajes_error_comision_db(db_manager):
    """
    Lista todos los viajes con error de comisión.
    """
    # Obtener todos los viajes con advertencia ERR. COM
    todos_los_viajes = []
    for id_viaje, datos_viaje in db_manager.data['viajes'].items():
        advertencias = datos_viaje.get('advertencias', [])
        
        if 'ERR. COM' in advertencias:
            viaje = datos_viaje.copy()
            viaje['id_viaje'] = id_viaje
            todos_los_viajes.append(viaje)
    
    if not todos_los_viajes:
        print("\nNo se encontraron viajes con error de comisión.\n")
        return []
    
    # Enriquecer con estadísticas de casetas
    viajes_enriquecidos = enriquecer_viajes_con_casetas(db_manager, todos_los_viajes)
    
    # Imprimir usando la función estándar
    print("\n=== VIAJES CON ERROR DE COMISIÓN ===")
    imprimir_filtros(viajes_enriquecidos)
    
    return viajes_enriquecidos

##############################Funciones version 9###################################################################################################
##############################Funciones version 9###################################################################################################
##############################Funciones version 9###################################################################################################

def normalizar_fechas_casetas(datacas, cas_fecha_col=9, cas_hora_col=10, header_rows=1):
    """
    Normaliza las columnas de fecha (J, índice cas_fecha_col) y hora (K, índice cas_hora_col)
    en datacas. Para filas donde la hora es NaN, intenta extraer la hora desde la columna fecha
    si ésta contiene un datetime completo (p. ej. '2025-09-27 01:54:50').
    - datacas: lista de filas (header incluido en índice 0)
    - header_rows: número de filas de encabezado (por defecto 1)
    Modifica datacas in-place y devuelve (datacas, n_modificados).
    """
    if not datacas or len(datacas) <= header_rows:
        return datacas, 0

    n_modificados = 0
    for i in range(header_rows, len(datacas)):
        row = datacas[i]
        # Asegurar longitud suficiente
        max_idx = max(cas_fecha_col, cas_hora_col)
        if len(row) <= max_idx:
            # rellenar con None si la fila es corta
            row.extend([None] * (max_idx - len(row) + 1))

        fecha_val = row[cas_fecha_col]
        hora_val = row[cas_hora_col]

        # Si hora ya existe y no es NaN, nada que hacer
        if not pd.isna(hora_val):
            continue

        # Si hora es NaN pero fecha contiene datetime con hora, extraer
        if fecha_val is None or pd.isna(fecha_val):
            continue

        # Intentar parsear fecha_val como datetime
        try:
            dt = pd.to_datetime(fecha_val, infer_datetime_format=True, errors='coerce')
        except Exception:
            dt = pd.NaT

        if pd.isna(dt):
            # Si no se pudo parsear, intentar si la celda es string con espacio separador
            s = str(fecha_val).strip()
            if ' ' in s:
                parts = s.split()
                if len(parts) >= 2:
                    date_part = parts[0]
                    time_part = parts[1]
                    # intentar parsear date_part y time_part
                    try:
                        dt_date = pd.to_datetime(date_part, errors='coerce')
                        dt_time = pd.to_datetime(time_part, format='%H:%M:%S', errors='coerce')
                        if not pd.isna(dt_date) and not pd.isna(dt_time):
                            # asignar fecha y hora
                            datacas[i][cas_fecha_col] = dt_date.strftime('%Y-%m-%d 00:00:00')
                            datacas[i][cas_hora_col] = dt_time.strftime('%H:%M:%S')
                            n_modificados += 1
                            continue
                    except Exception:
                        pass
            # no se pudo extraer
            continue

        # Si dt es válido, extraer fecha y hora
        try:
            date_str = dt.strftime('%Y-%m-%d 00:00:00')  # mantener formato similar al resto (fecha con 00:00:00)
            time_str = dt.strftime('%H:%M:%S')
            datacas[i][cas_fecha_col] = date_str
            datacas[i][cas_hora_col] = time_str
            n_modificados += 1
        except Exception:
            # si falla, no modificar
            continue

    print(f"[NORMALIZAR] Filas de casetas modificadas (fecha/hora extraídas): {n_modificados}")
    return datacas, n_modificados

def build_info_from_dataayc(dataayc, ayc_header_rows=4, col_id_viaje=1, col_L=11, col_M=12, col_H=7):
    """
    Construye un mapa { id_viaje: {'colL': valL, 'frontal':..., 'dolly':..., 'trasero':...} }
    - ayc_header_rows: 1-based row index of header (user said header row 4 -> pass 4)
      Internamente se usa como 0-based start = ayc_header_rows (fila siguiente al header).
    - col_id_viaje, col_L, col_M: índices 0-based de las columnas B, L, M por defecto.
    Retorna (info_by_viaje, conteo) donde conteo es número de registros procesados.
    """
    info_by_viaje = {}
    if not dataayc or len(dataayc) <= ayc_header_rows:
        print("[INFO] dataayc vacío o sin filas de datos según ayc_header_rows.")
        return info_by_viaje, 0

    start_idx = ayc_header_rows  # header en fila 4 (1-based) -> start en índice 4 (0-based)
    procesados = 0
    for row in dataayc[start_idx:]:
        try:
            id_v = row[col_id_viaje] if col_id_viaje < len(row) else None
        except Exception:
            id_v = None
        if id_v is None or pd.isna(id_v):
            continue
        id_v_str = str(id_v).strip()
        # Obtener columna L y M
        valL = row[col_L] if col_L < len(row) else None
        valM = row[col_M] if col_M < len(row) else None
        clave = row[col_H] if col_H < len(row) else None

        frontal, dolly, trasero = parse_ensamble_field(valM)

        # Guardar (si ya existe, no sobrescribimos salvo que quieras actualizar)
        if id_v_str not in info_by_viaje:
            info_by_viaje[id_v_str] = {
                'clave': str(clave).strip() if clave and not pd.isna(clave) else None,
                'colL': valL if not pd.isna(valL) else None,
                'frontal': frontal or None,
                'dolly': dolly or None,
                'trasero': trasero or None
            }
        else:
            # Si ya existe, preferimos no sobrescribir; pero si hay campos vacíos, rellenar
            cur = info_by_viaje[id_v_str]
            if not cur.get('clave') and clave and not pd.isna(clave):
                cur['clave'] = str(clave).strip()
            if not cur.get('colL') and valL and not pd.isna(valL):
                cur['colL'] = valL
            if not cur.get('frontal') and frontal:
                cur['frontal'] = frontal
            if not cur.get('dolly') and dolly:
                cur['dolly'] = dolly
            if not cur.get('trasero') and trasero:
                cur['trasero'] = trasero

        procesados += 1

    print(f"[INFO] build_info_from_dataayc: registros procesados en dataayc = {procesados}, viajes únicos = {len(info_by_viaje)}")
    return info_by_viaje, procesados

def parse_ensamble_field(valor_m):
    """
    Extrae Frontal, Dolly, Trasero desde la cadena de la columna M.
    Escenarios esperados:
      - "FRONTAL: PD.001-A."
      - "FRONTAL: TG.349-ADOLLY: DL.553-ATRASERO: TG.350."
    Devuelve tupla (frontal, dolly, trasero) con strings o '' si no existen.
    """
    if valor_m is None or pd.isna(valor_m):
        return '', '', ''

    s = str(valor_m).strip()
    # Normalizar quitando puntos finales redundantes y espacios extra
    s = s.rstrip('.').strip()

    # Buscar patrones con etiquetas explícitas
    frontal = ''
    dolly = ''
    trasero = ''

    # Usar regex para capturar XX.000 (dos letras, punto, tres dígitos y opcional sufijo)
    token_re = re.compile(r'([A-Z]{2}\.\d{3}[A-Z]?)', re.IGNORECASE)

    # Caso completo con etiquetas
    if 'DOLLY' in s.upper() or 'TRASERO' in s.upper():
        # Intentar extraer por etiquetas
        # FRONTAL: ... - ADOLLY: ... - ATRASERO: ...
        parts = re.split(r'\s*-\s*', s)
        for p in parts:
            p_up = p.upper()
            if 'FRONTAL' in p_up:
                m = token_re.search(p)
                frontal = m.group(1) if m else p.split(':')[-1].strip()
            elif 'DOLLY' in p_up:
                m = token_re.search(p)
                dolly = m.group(1) if m else p.split(':')[-1].strip()
            elif 'TRASERO' in p_up:
                m = token_re.search(p)
                trasero = m.group(1) if m else p.split(':')[-1].strip()
    else:
        # Caso simple: solo FRONTAL
        # Buscar token en la cadena
        m = token_re.search(s)
        if m:
            frontal = m.group(1)
        else:
            # fallback: tomar lo que venga después de ':' si existe
            if ':' in s:
                frontal = s.split(':', 1)[-1].strip()
            else:
                frontal = s

    # Normalizar a mayúsculas y sin espacios extremos
    frontal = frontal.strip()
    dolly = dolly.strip()
    trasero = trasero.strip()

    return frontal, dolly, trasero

def extraer_numero_de_tr(tr_string):
    """Extrae ### de 'TR.###-A' -> '###' o devuelve None si no coincide."""
    try:
        s = str(tr_string)
        m = re.search(r'TR\.(\d+)-', s)
        if m:
            return m.group(1)
    except Exception:
        pass
    return None

def normalizar_ct_token(token):
    """Normaliza tokens tipo 'CT-03' -> devuelve 'CT-3' y el número entero 3."""
    if token is None:
        return None, None
    s = str(token).strip().upper()
    m = re.match(r'^CT-0*(\d+)$', s)
    if m:
        num = int(m.group(1))
        return f"CT-{num}", num
    # aceptar también formatos sin guion o variantes si las hubiera
    m2 = re.match(r'^CT(\d+)$', s)
    if m2:
        num = int(m2.group(1))
        return f"CT-{num}", num
    return None, None

def mapear_y_reemplazar_ct_por_tracto(datacas, dataayc, 
                                      cas_tracto_col=8, ayc_tracto_col=8, ayc_tr_col=13):
    """
    Busca CT-## en datacas[col I] y los reemplaza por el número de tracto obtenido desde dataayc[col N].
    Imprime resumen de mapeos y reemplazos.
    Retorna (datacas_modificada, mapa_ct_a_tracto)
    """
    if not datacas or not dataayc:
        print("No hay datos suficientes en datacas o dataayc para realizar el mapeo.")
        return datacas, {}

    # 1) Obtener CT únicos de datacas (columna cas_tracto_col)
    ct_unicos = set()
    for row in datacas[1:]:  # saltar header si existe
        try:
            val = row[cas_tracto_col] if cas_tracto_col < len(row) else None
        except Exception:
            val = None
        if val is None or pd.isna(val):
            continue
        token = str(val).strip().upper()
        if token.startswith('CT-') or re.match(r'^CT\d+$', token):
            ct_norm, num = normalizar_ct_token(token)
            if ct_norm:
                ct_unicos.add(ct_norm)

    if not ct_unicos:
        print("No se encontraron tokens CT-## en datacas.")
        return datacas, {}

    print(f"[CT MAPPING] Tokens CT únicos detectados en datacas: {len(ct_unicos)} -> {sorted(ct_unicos)}")

    # 2) Construir mapa CT -> tracto (string) buscando en dataayc
    mapa = {}  # CT token -> tracto string (ej. '311') o None si no encontrado
    # Preprocesar dataayc para acelerar búsquedas: lista de (valor_colI, valor_colN)
    ayc_indexed = []
    for row in dataayc[1:]:
        try:
            key = row[ayc_tracto_col] if ayc_tracto_col < len(row) else None
            val_tr = row[ayc_tr_col] if ayc_tr_col < len(row) else None
        except Exception:
            key = None
            val_tr = None
        ayc_indexed.append((str(key).strip() if key is not None else '', val_tr))

    for ct in sorted(ct_unicos):
        # extraer número del CT
        _, num = normalizar_ct_token(ct)
        if num is None:
            mapa[ct] = None
            continue

        # buscar en dataayc una fila cuya columna I termine en 'C{num}' o 'C{num}' con ceros
        encontrado_tr = None
        pattern = re.compile(rf'C0*{num}$', re.IGNORECASE)
        for key, val_tr in ayc_indexed:
            if key and pattern.search(key):
                # extraer ### de val_tr (TR.###-A)
                tracto_num = extraer_numero_de_tr(val_tr)
                if tracto_num:
                    encontrado_tr = tracto_num
                    break
        mapa[ct] = encontrado_tr

    # 3) Reemplazar en datacas todos los CT por el tracto encontrado
    total_reemplazos = 0
    reemplazos_por_ct = {}
    for i in range(1, len(datacas)):  # mantener header en índice 0
        row = datacas[i]
        if cas_tracto_col >= len(row):
            continue
        val = row[cas_tracto_col]
        if val is None or pd.isna(val):
            continue
        token = str(val).strip().upper()
        # normalizar token para comparar con mapa keys
        ct_norm, _ = normalizar_ct_token(token)
        if ct_norm and ct_norm in mapa:
            nuevo = mapa[ct_norm]
            if nuevo is not None:
                # reemplazar por string del tracto
                datacas[i][cas_tracto_col] = nuevo
                total_reemplazos += 1
                reemplazos_por_ct[ct_norm] = reemplazos_por_ct.get(ct_norm, 0) + 1
            else:
                # no hay mapping; dejar tal cual (o podrías marcarlo para revisión)
                pass

    # 4) Imprimir resumen
    mapeados = [k for k, v in mapa.items() if v is not None]
    no_mapeados = [k for k, v in mapa.items() if v is None]
    print(f"[CT MAPPING] Mapeos encontrados: {len(mapeados)}")
    for k in mapeados:
        print(f"  {k} -> {mapa[k]}")
    if no_mapeados:
        print(f"[CT MAPPING] Tokens sin mapeo ({len(no_mapeados)}): {no_mapeados}")
    print(f"[CT MAPPING] Total de reemplazos realizados en datacas: {total_reemplazos}")
    if reemplazos_por_ct:
        print("[CT MAPPING] Reemplazos por token:")
        for k, cnt in reemplazos_por_ct.items():
            print(f"  {k}: {cnt} filas")

    return datacas, mapa

##############################Funciones version 7###################################################################################################
##############################Funciones version 7###################################################################################################
##############################Funciones version 7###################################################################################################

##############################Funciones version 6###################################################################################################
##############################Funciones version 6###################################################################################################
##############################Funciones version 6###################################################################################################

def imprimir_lista_casetas(registros):
    """
    Imprime una lista de registros de caseta en formato uniforme.
    Cada registro debe ser un dict con claves:
      'tracto', 'fecha', 'hora', 'importe', 'nombre', opcional 'id_combinado'
    """
    if not registros:
        print("\nNo hay casetas para mostrar.\n")
        return

    print("\n" + "="*112)
    #print(f"{'Tracto':8} | {'Fecha':12} | {'Hora':10} | {'Importe':12} | {'Nombre Caseta':30} | {'ID_COMBINADO':15} |")
    print(f"{'Tarjeta':15} | {'Tracto':6} | {'Fecha':12} | {'Hora':10} | {'Importe':12} | {'Nombre Caseta':30} | {'Nota':10} |")
    print("-"*112)
    for r in registros:
        tracto = str(r.get('tracto') or '')[:8]
        fecha = str(r.get('fecha') or '')[:12]
        hora = str(r.get('hora') or '')[:10]
        importe = r.get('importe')
        status = r.get('status')
        tarjeta = str(r.get('tarjeta') or '')[:15]
        if importe is not None and not pd.isna(importe):
            try:
                imp_str = f"${float(importe):,.2f}"
            except Exception:
                imp_str = str(importe)
        else:
            imp_str = ''
        nombre = ('C:' + str(r.get('clase') or '') + '. ' + str(r.get('nombre') or ''))[:30]
        idc = str(r.get('id_combinado') or '')[:15]
        #print(f"{tracto:8} | {fecha:12} | {hora:10} | {imp_str:12} | {nombre:30} | {idc:15} |")
        print(f"{tarjeta:15} | {tracto:6} | {fecha:12} | {hora:10} | {imp_str:12} | {nombre:30} | {status:10} |")
    print("="*112 + "\n")
    

# -------------------------------------------------------------------------------

# ------------------ Helpers para formateo y parseo de filtro de fecha ------------------
def format_date_safe(dt):
    if dt is None or pd.isna(dt):
        return ''
    try:
        return pd.to_datetime(dt).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return str(dt)

def fecha_coincide_con_filtro(fecha_dt, filtro_fecha_str):
    """
    filtro_fecha_str puede ser 'YYYY', 'YYYY-MM' o 'YYYY-MM-DD'.
    Devuelve True si el string del filtro aparece en la representación YYYY-MM-DD de fecha_dt.
    """
    if fecha_dt is None or pd.isna(fecha_dt):
        return False
    try:
        s = pd.to_datetime(fecha_dt).strftime('%Y-%m-%d')
    except Exception:
        s = str(fecha_dt)
    return filtro_fecha_str in s
# ---------------------------------------------------------------------------------------

def imprimir_filtros(encontrados):
    """
    Imprime viajes en formato tabular.
    Compatible con la estructura de datos del DatabaseManager.
    """
    if not encontrados:
        print("\nNo hay viajes para mostrar.\n")
        return
    
    print("\n" + "="*215)
    print(f"{'ID viaje':14} | {'Fecha inicio':19} | {'Fecha final':19} | {'Ruta':10} | {'Operador':8} | {'Tracto':6} | {'Ensa':4} | {'Frontal':7} | {'Dolly':7} | {'Trasero':7} | {'#C':2} | {'Importe':10} | {'#CV':3} | {'Importe V':10} | {'Total':10} | {'Gastos':10} | {'Comision':10} | {'Estado':8} | {'origen':14} |")
    print("-"*215)
    
    for e in encontrados:
        id_v = str(e.get('id_viaje', ''))[:14]
        fi = format_date_safe(e.get('fecha_inicio'))
        ff = format_date_safe(e.get('fecha_final'))
        ruta = str(e.get('ruta') or '')[:10]
        operador = str(e.get('operador') or '')[:8]
        tracto = str(e.get('tracto') or '')[:6]
        
        # Campos de cierre (pueden venir como 'ensamble' o 'colL')
        ensamble = str(e.get('ensamble') or e.get('colL') or '')[:4]
        frontal = str(e.get('frontal') or '')[:7]
        dolly = str(e.get('dolly') or '')[:7]
        trasero = str(e.get('trasero') or '')[:7]
        
        # Estadísticas de casetas
        ncas = e.get('num_casetas', 0)
        imp = e.get('importe_total', 0.0)
        ncas_fuera = e.get('num_casetas_fuera', 0)
        imp_fuera = e.get('importe_fuera', 0.0)
        imp_total = imp_fuera + imp
        
        # Estadisticas de gastos
        gasto = e.get('gasto', 0.0)
        comision = e.get('comision', 0.0)
        origen = str(e.get('origen') or '')[:14]
        
        # Estado con sistema de advertencias
        advertencias = e.get('advertencias', [])
        if advertencias:
            # Mostrar advertencias separadas por coma
            status = ', '.join(advertencias)[:12]
        else:
            status = 'OK'
        
        # Formatear valores numéricos
        ncas_str = f"{ncas:d}" if ncas != 0 else ''
        imp_str = f"${imp:,.2f}" if imp != 0 else ''
        ncas_fuera_str = f"{ncas_fuera:d}" if ncas_fuera != 0 else ''
        imp_fuera_str = f"${imp_fuera:,.2f}" if imp_fuera != 0 else ''
        imp_total_str = f"${imp_total:,.2f}" if imp_total != 0 else ''
        gasto_str = f"${gasto:,.2f}" if gasto != 0 else ''
        comision_str = f"${comision:,.2f}" if comision != 0 else ''
        
        print(f"{id_v:14} | {fi:19} | {ff:19} | {ruta:10} | {operador:8} | {tracto:6} | {ensamble:4} | {frontal:7} | {dolly:7} | {trasero:7} | {ncas_str:2} | {imp_str:10} | {ncas_fuera_str:3} | {imp_fuera_str:10} | {imp_total_str:10} | {gasto_str:10} | {comision_str:10} | {status:8} | {origen:14} |")
    
    print("="*215 + "\n")

##############################Funciones version 5###################################################################################################
##############################Funciones version 5###################################################################################################
##############################Funciones version 5###################################################################################################

DATE_RE = re.compile(r'\d{1,2}/\d{1,2}/\d{2}')  # detecta MM/DD/AA

def parse_datetime_monitoreo(datetime_str):
    """
    Convierte 'MM/DD/AA HH:MM:SS XM' a datetime o devuelve None si no es fecha válida.
    """
    if pd.isna(datetime_str):
        return None

    s = str(datetime_str).strip()
    # Normalizar espacios internos (quita dobles espacios)
    s = ' '.join(s.split())

    # Rechazar textos que claramente no contienen fecha
    if not DATE_RE.search(s):
        return None

    # Intento con formato estricto primero
    try:
        return pd.to_datetime(s, format='%m/%d/%y %I:%M:%S %p', errors='raise')
    except Exception:
        # Intento más flexible: inferir formato, devolver NaT si no puede
        dt = pd.to_datetime(s, infer_datetime_format=True, errors='coerce')
        if pd.isna(dt):
            return None
        return dt

def parse_datetime_caseta(fecha_str, hora_str):
    """
    Combina fecha y hora de caseta en un datetime.
    fecha: 'AAAA-MM-DD 00:00:00'
    hora: 'HH:MM:SS'
    """
    if pd.isna(fecha_str) or pd.isna(hora_str):
        return None
    try:
        fecha_str = str(fecha_str).strip().split()[0]  # Tomar solo la fecha
        hora_str = str(hora_str).strip()
        datetime_str = f"{fecha_str} {hora_str}"
        dt = pd.to_datetime(datetime_str, format='%Y-%m-%d %H:%M:%S')
        return dt
    except Exception as e:
        print(f"Error parseando fecha caseta '{fecha_str}' '{hora_str}': {e}")
        return None

def extraer_numero_tracto(tracto_str):
    """
    Extrae el número de tracto de formato 'TR.###-X' -> '###'
    Ejemplo: 'TR.511-A' -> '511'
    """
    if pd.isna(tracto_str):
        return None
    try:
        tracto_str = str(tracto_str).strip()
        # Buscar patrón TR.###-X
        m = re.match(r'TR\.(\d+)-[A-Za-z]', tracto_str)
        if m:
            return m.group(1)
        return None
    except Exception as e:
        print(f"Error extrayendo tracto '{tracto_str}': {e}")
        return None

##############################Funciones de uso###################################################################################################
##############################Funciones de uso###################################################################################################
##############################Funciones de uso###################################################################################################

# Orden hardcoded para casetas
def key_func(row):
    """
    Devuelve una tupla numérica (ts_fecha_ns, segundos_hora) para ordenar.
    Si falta fecha/hora devuelve valores muy pequeños para que queden al inicio.
    """
    cas_fecha_col = 9
    cas_hora_col = 10

    def fecha_to_ns(x):
        if pd.isna(x):
            return float('-inf')
        try:
            dt = pd.to_datetime(x, infer_datetime_format=True, errors='coerce')
            if pd.isna(dt):
                return float('-inf')
            return int(dt.value)  # nanoseconds since epoch (int)
        except Exception:
            return float('-inf')

    def hora_to_seconds(x):
        if pd.isna(x):
            return float('-inf')
        try:
            # parsear hora o datetime y extraer hora:min:seg
            t = pd.to_datetime(str(x), infer_datetime_format=True, errors='coerce')
            if pd.isna(t):
                # intentar parsear solo como time string HH:MM:SS
                parts = str(x).strip().split(':')
                if len(parts) >= 3:
                    h, m, s = int(parts[0]), int(parts[1]), int(parts[2])
                    return h*3600 + m*60 + s
                return float('-inf')
            tt = t.time()
            return tt.hour*3600 + tt.minute*60 + tt.second
        except Exception:
            return float('-inf')

    # proteger índices cortos
    a = row[cas_fecha_col] if cas_fecha_col < len(row) else None
    b = row[cas_hora_col] if cas_hora_col < len(row) else None

    return (fecha_to_ns(a), hora_to_seconds(b))

# Funcion para encontrar nan en filtros
def is_nan_safe(row, idx):
    if idx >= len(row):
        return True
    return pd.isna(row[idx])
    
#Funcion para coincidencia de registros
def is_same(row1, row2):
    return row1 == row2

#Eliminar registros hidvchis de monitoreo, ordenar casetas, eliminar registros vacios
def preparar_excel(data, opc):

    # preparar datos de monitoreo
    if opc == "monitoreo":

        # Conteo inicial
        total_inicial = len(data)

        # 1) Eliminar filas que empiezan con 'HIDVCHID' en la columna A (índice 0)
        datos_filtrados = [fila for fila in data if not str(fila[0]).startswith('HIDVCHID')]
        eliminados_hid = total_inicial - len(datos_filtrados)
        print(f"Monitoreo: eliminados por HIDVCHID: {eliminados_hid} registros")

        # 2) Eliminar filas con NaN en la columna A (índice 0)
        antes_nan = len(datos_filtrados)
        datos_filtrados = [fila for fila in datos_filtrados if not is_nan_safe(fila, 0)]
        eliminados_nan = antes_nan - len(datos_filtrados)
        print(f"Monitoreo: eliminados por NAN en columna A: {eliminados_nan} registros")

        # 3) Eliminar filas que contengan la palabra 'CANCELADO' en la columna B (índice 1)
        antes_cancel = len(datos_filtrados)
        # Comparación case-insensitive; convertir a str para evitar errores con NaN u otros tipos
        datos_filtrados = [fila for fila in datos_filtrados if 'CANCELADO' not in str(fila[1]).upper()]
        eliminados_cancel = antes_cancel - len(datos_filtrados)
        print(f"Monitoreo: eliminados por 'CANCELADO' en columna B: {eliminados_cancel} registros")

        # Resumen
        total_final = len(datos_filtrados)
        print(f"Monitoreo: total inicial={total_inicial}, total final={total_final}, eliminados={total_inicial - total_final}")

        return datos_filtrados
        
    # preparar datos de casetas
    elif opc == "casetas":
        
        #Try to sort casetas por fecha y hora
    
        header = data[0]
        body = data[1:]
        
        #orden ascendente primero J luego K, funcion hardcoded a J K
        body_sorted = sorted(body, key=key_func)
        
        #reconstruir lista incluyendo header
        data = [header] + body_sorted
        
        print("Casetas ordenadas")

        
        return data
        
    # preparar datos de cerrados
    else: # config para cerrados
        
        # Conteo inicial
        total_inicial = len(data)
        
        # 1) Eliminar filas con NaN en la columna B (índice 1)
        datos_filtrados = [fila for fila in data if not is_nan_safe(fila, 1)]
        eliminados_nan = total_inicial - len(datos_filtrados)
        print(f"Cerrados: eliminados por NAN en columna B: {eliminados_nan} registros")

        # 2) Eliminar filas que contengan la palabra 'CANCELADO' en la columna A (índice 0)
        antes_cancel = len(datos_filtrados)
        # Comparación case-insensitive; convertir a str para evitar errores con NaN u otros tipos
        datos_filtrados = [fila for fila in datos_filtrados if 'CANCELADO' not in str(fila[0]).upper()]
        eliminados_cancel = antes_cancel - len(datos_filtrados)
        print(f"Cerrados: eliminados por 'CANCELADO' en columna B: {eliminados_cancel} registros")

        # Resumen
        total_final = len(datos_filtrados)
        print(f"Cerrados: total inicial={total_inicial}, total final={total_final}, eliminados={total_inicial - total_final}")

        return datos_filtrados

def preparar_excel_gastos_adicionales(data, header_rows=1):
    """
    Prepara el excel de gastos adicionales eliminando registros inválidos.
    
    Elimina filas donde:
    - Columna D (ID viaje) contiene "SIN VIAJE"
    - Columna D está vacía
    
    Parámetros:
    - data: lista de listas con los datos del excel
    - header_rows: número de filas de encabezado (default: 1)
    
    Retorna: data limpio
    """
    if not data or len(data) <= header_rows:
        return data
    
    col_id_viaje = 3  # Columna D (índice 3)
    
    # Separar header y datos
    header = data[:header_rows]
    rows = data[header_rows:]
    
    # Filtrar filas válidas
    rows_validas = []
    
    for row in rows:
        if len(row) <= col_id_viaje:
            continue
        
        id_viaje = row[col_id_viaje]
        
        # Saltar si está vacío o es "SIN VIAJE"
        if pd.isna(id_viaje):
            continue
        
        id_viaje_str = str(id_viaje).strip().upper()
        
        if not id_viaje_str or id_viaje_str == "SIN VIAJE":
            continue
        
        rows_validas.append(row)
    
    print(f"[GASTOS ADICIONALES] Filas originales: {len(rows)}, válidas: {len(rows_validas)}")
    
    return header + rows_validas

# Buscar y eliminar duplicados
def verificar_valores_unicos(data, tag, opc):
    """
    Verifica que todos los valores en la columna A sean únicos.
    Retorna una lista de valores duplicados con sus conteos.
    """
    
    if tag == 0: #excel monitoreo
        valores_columna_a = [fila[0] for fila in data]
        tag = "monitoreo"
    else: #excel apertura/cierre
        valores_columna_a = [fila[1] for fila in data]
        tag = "apertura/cierre"
    
    # Contar ocurrencias de cada valor
    conteo = {}
    for valor in valores_columna_a:
        conteo[valor] = conteo.get(valor, 0) + 1
    
    # Filtrar solo los duplicados
    duplicados = {valor: count for valor, count in conteo.items() if count > 1}
    
    if duplicados:
        print("\nValores duplicados encontrados en "+tag+":")
        for valor, count in duplicados.items():
            print(f"  '{valor}': aparece {count} veces")
            
            if opc == "procesoeliminar":
                # Eliminar duplicados de monitoreo
                if tag == "monitoreo":
                    
                    #Logica si el valor es nan, eliminar todos los nan
                    if pd.isna(valor):
                        print("     -Eliminando registros sin datos")
                        nan_rows = [r for r in data if is_nan_safe(r, 0)]
                    
                    # Logica para buscar el valor y eliminar duplicados
                    # Filtrar excel por valor en columna_a
                    valorindata = [row for row in data if row[0] == valor]
                    
                    # Separar registros que tienen nan en fechas de salida
                    valorindatawithdata = [r for r in valorindata if not is_nan_safe(r, 16)]
                    
                    # Si solo hay un registro con datos
                    if len(valorindatawithdata) == 1:
                        print("     -Solo un registro tiene datos de salida, se elimina el resto")
                        # logica de eliminar datamon - valorindatawithoutdata
                        valorindatawithoutdata = [r for r in valorindata if is_nan_safe(r, 16)]
                        data = [r for r in data if not any(is_same(r, nr) for nr in valorindatawithoutdata)]
                        
                    # Varios registros con datos
                    elif len(valorindatawithdata) > 1:
                        print("     -Varios registros tienen datos de salida, revisar excel")
                        valorQ = set(registro[16] for registro in valorindata)
                        for registro in valorQ:
                            print(registro)
                    
                    # No registros con datos
                    elif len(valorindatawithdata) == 0:
                        print("     -No hay registros con datos de salida, revisar excel")
                
         
    else:
        print("\nTodos los valores en "+tag+" son únicos.")
    
    if opc == "procesoeliminar":
        return data
    else:
        return duplicados

# Funcion para crear un nuevo ID, especificamente en excel de casetas y asi poder identificar registros unicos
def create_combined_id(data, col_a=8, col_b=9, col_c=10, header_rows=1, new_col_name='ID_COMBINADO', insert_at=None):
    """
    Combina columnas col_a y col_b en una nueva columna new_col_name.
    - data: lista de listas (hoja completa).
    - col_a, col_b: índices 0-based de las columnas a combinar (I=8, J=9, K=10).
    - header_rows: número de filas de encabezado (por defecto 1).
    - new_col_name: nombre para la columna nueva (se añade al header).
    - insert_at: índice 0-based donde insertar la nueva columna; si None, se añade al final.
    Retorna (data_modificada, new_col_index).
    """
    if not data:
        return data, None

    # Asegurar que hay al menos header_rows filas
    if len(data) < header_rows:
        return data, None

    # Si ya existe una columna con ese nombre en el header, devolver su índice
    header = data[0]
    if new_col_name in header:
        new_idx = header.index(new_col_name)
        return data, new_idx

    # Determinar posición de inserción
    if insert_at is None:
        insert_at = len(header)

    # Insertar el nombre de la nueva columna en el header
    header.insert(insert_at, new_col_name)

    # Procesar filas de datos
    for i in range(header_rows, len(data)):
        row = data[i]
        # Asegurar que la fila tenga suficientes columnas; si no, rellenar con None
        max_idx = max(col_a, col_b, col_c, insert_at)
        if len(row) <= max_idx:
            row.extend([None] * (max_idx - len(row) + 1))

        a = row[col_a]
        b = row[col_b]
        c = row[col_c]

        # Normalizar valores (evitar NaN de pandas)
        if pd.isna(a):
            a_val = ''
        else:
            a_val = str(a).strip()

        if pd.isna(b):
            b_val = ''
        else:
            b_val = str(b).strip()
            
        if pd.isna(c):
            c_val = ''
        else:
            c_val = str(c).strip()

        # Construir ID combinado; ajustar separador según prefieras
        if a_val == '' and b_val == '':
            combined = None
        elif a_val == '':
            combined = b_val
        elif b_val == '':
            combined = a_val
        else:
            combined = f"{a_val}-{b_val}-{c_val}"

        # Insertar la nueva columna en la posición deseada
        row.insert(insert_at, combined)

    return data, insert_at
 
#Funcion para combinar exceles EN BASE AL ID, recibe datos anteriores, nuevos, inicio de filas y columnas del ID
def merge_excel_data(existing, new, start_data, id_col):
    """
    Combina datos Excel siguiendo la lógica: REEMPLAZAR + MANTENER
    - Si un ID existe en ambos: se reemplaza con el nuevo
    - Si un ID solo existe en existing: se mantiene
    - Si un ID solo existe en new: se agrega
    
    Parámetros:
    - existing: datos anteriores (lista de filas)
    - new: datos nuevos (lista de filas)
    - start_data: fila donde empiezan los datos (después del header)
    - id_col: columna del ID
    
    Retorna: lista combinada
    """
    if existing == 0 or existing is None:
        return new

    # Copiar encabezado(s) de existing
    combined = []
    if start_data > 0:
        combined.extend(existing[:start_data])

    # Construir diccionario de filas por ID (para fácil reemplazo)
    filas_por_id = {}
    
    # Primero agregar todas las filas existentes
    for row in existing[start_data:]:
        try:
            rid = row[id_col] if id_col < len(row) else None
        except Exception:
            rid = None
        
        if rid is not None and not pd.isna(rid):
            rid_str = str(rid).strip()
            filas_por_id[rid_str] = row
        else:
            # Filas sin ID válido se mantienen pero no se pueden actualizar
            # Las agregamos con un ID temporal único
            import uuid
            temp_id = f"__no_id_{uuid.uuid4()}"
            filas_por_id[temp_id] = row
    
    # Ahora procesar las filas nuevas (reemplazar o agregar)
    nuevos = 0
    actualizados = 0
    
    for row in new[start_data:]:
        try:
            rid = row[id_col] if id_col < len(row) else None
        except Exception:
            rid = None
        
        if rid is None or pd.isna(rid):
            # Fila sin ID válido, agregar como nueva
            import uuid
            temp_id = f"__no_id_{uuid.uuid4()}"
            filas_por_id[temp_id] = row
            nuevos += 1
        else:
            rid_str = str(rid).strip()
            if rid_str in filas_por_id:
                # Reemplazar existente
                filas_por_id[rid_str] = row
                actualizados += 1
            else:
                # Agregar nuevo
                filas_por_id[rid_str] = row
                nuevos += 1
    
    # Reconstruir lista manteniendo orden (primero existentes, luego nuevos)
    combined.extend(filas_por_id.values())
    
    print(f"Merge: {nuevos} nuevos, {actualizados} actualizados, {len(filas_por_id)} total")
    return combined

def merge_gastos_adicionales(data_existente, data_nuevo, header_rows=1):
    """
    Combina datos de gastos adicionales (actualización inteligente).
    
    Identifica registros únicos por: ID_viaje (col D) + Clasificacion (col P)
    Si existe el mismo ID_viaje+Clasificacion, reemplaza el registro.
    Si no existe, lo agrega.
    
    Parámetros:
    - data_existente: datos actuales
    - data_nuevo: datos nuevos a incorporar
    - header_rows: filas de encabezado
    
    Retorna: datos combinados
    """
    if not data_existente or len(data_existente) <= header_rows:
        return data_nuevo
    
    if not data_nuevo or len(data_nuevo) <= header_rows:
        return data_existente
    
    col_id_viaje = 3   # Columna D
    col_clasificacion = 15  # Columna P
    
    # Separar headers y datos
    header = data_existente[:header_rows]
    rows_existentes = data_existente[header_rows:]
    rows_nuevos = data_nuevo[header_rows:]
    
    # Crear mapa de registros existentes: {(id_viaje, clasificacion): row}
    mapa_existentes = {}
    
    for row in rows_existentes:
        if len(row) <= max(col_id_viaje, col_clasificacion):
            continue
        
        id_viaje = str(row[col_id_viaje]).strip() if not pd.isna(row[col_id_viaje]) else ""
        clasificacion = str(row[col_clasificacion]).strip() if not pd.isna(row[col_clasificacion]) else ""
        
        if id_viaje and clasificacion:
            key = (id_viaje, clasificacion)
            mapa_existentes[key] = row
    
    # Procesar nuevos registros
    registros_actualizados = 0
    registros_agregados = 0
    
    for row in rows_nuevos:
        if len(row) <= max(col_id_viaje, col_clasificacion):
            continue
        
        id_viaje = str(row[col_id_viaje]).strip() if not pd.isna(row[col_id_viaje]) else ""
        clasificacion = str(row[col_clasificacion]).strip() if not pd.isna(row[col_clasificacion]) else ""
        
        if id_viaje and clasificacion:
            key = (id_viaje, clasificacion)
            
            if key in mapa_existentes:
                # Reemplazar registro existente
                mapa_existentes[key] = row
                registros_actualizados += 1
            else:
                # Agregar nuevo registro
                mapa_existentes[key] = row
                registros_agregados += 1
    
    # Reconstruir datos
    rows_finales = list(mapa_existentes.values())
    
    print(f"[MERGE GASTOS ADICIONALES] Actualizados: {registros_actualizados}, Agregados: {registros_agregados}, Total: {len(rows_finales)}")
    
    return header + rows_finales

##############################MAIN###################################################################################################
##############################MAIN###################################################################################################
##############################MAIN###################################################################################################

##############################Flujo Principal del Programa###################################################################################################

def main():
    """
    Punto de entrada principal del programa.
    Detecta si existe base de datos y actúa en consecuencia.
    """
    db_path = Path('database.json')
    db_manager = DatabaseManager(db_path)
    
    # Verificar si existe base de datos con datos
    if db_manager.data['viajes'] or db_manager.data['casetas']:
        print("\n=== Base de datos detectada ===")
        print(f"Viajes: {len(db_manager.data['viajes'])}")
        print(f"Casetas: {len(db_manager.data['casetas'])}")
        print(f"Relaciones: {len(db_manager.data['relaciones'])}")
        print("="*50 + "\n")
        
        # Ir directo al módulo de consulta
        modulo_consulta(db_manager)
    else:
        print("\n=== No se detectó base de datos ===")
        print("Iniciando carga inicial de datos...\n")
        
        # Primera carga: monitoreo y casetas OBLIGATORIOS
        db_manager = carga_inicial_datos(db_manager)
        
        # Después de cargar, ir al módulo de consulta
        modulo_consulta(db_manager)


def carga_inicial_datos(db_manager):
    """
    Carga inicial OBLIGATORIA de monitoreo y casetas.
    Otros excels son opcionales.
    
    Retorna: db_manager actualizado
    """
    print("CARGA INICIAL DE DATOS")
    print("="*50)
    print("Se requieren obligatoriamente:")
    print("  1. Excel de MONITOREO")
    print("  2. Excel de CASETAS")
    print("Opcionales:")
    print("  3. Excel de CIERRE")
    print("  4. Excel de GASTOS")
    print("  5. Excel de CLIENTE PARTICULAR")
    print("  6. Excel de TABLERO COSTO RUTA")
    print("="*50 + "\n")
    
    data = None
    datamon = None
    datacas = None
    dataayc = None
    datagastos = None
    new_gastos = None
    dataclientes = None
    datacostos = None
    
    while True:
        path = input("Ruta del archivo Excel (o 'scan dir' para procesar carpeta)> ").strip().strip('"').strip("'")
        
        if not path:
            print("No se proporcionó ruta")
            continue    
        
        path_lower = path.strip().lower()
        
        # Comando para procesar directorio
        if path_lower in ('scan dir', 'procesar directorio', 'procesar carpeta'):
            print("Escaneando directorio del script...")
            datamon, datacas, dataayc, datagastos, dataclientes, datacostos = procesar_excels_en_directorio_inicial()
            # Verificar que se cargaron los obligatorios
            if datamon is None or datacas is None:
                print("\n¡ERROR! No se detectaron archivos de monitoreo y/o casetas en el directorio.")
                print("Carga manual requerida.\n")
                continue
            break
        
        # Comando para finalizar carga
        if path_lower == 'go':
            if datamon is not None and datacas is not None:
                break
            else:
                print("\n¡FALTAN DATOS OBLIGATORIOS!")
                if datamon is None:
                    print("  - Falta: Excel de MONITOREO")
                if datacas is None:
                    print("  - Falta: Excel de CASETAS")
                print()
                continue
        
        # Cargar archivo Excel
        excel_path = Path(path)
        if not excel_path.exists():
            print(f"El archivo no existe: {excel_path}")
            continue
        
        try:
            df = pd.read_excel(excel_path, sheet_name=0, header=None, engine='openpyxl')
        except:
            try:
                df = pd.read_excel(excel_path, sheet_name=0, header=None)
            except Exception as e:
                print(f"Error al leer el archivo: {e}")
                continue
        
        # Detectar tipo de excel (primero verificar si es gastos)
        tipo_detectado = detectar_tipo_excel_multi_hoja(excel_path)
        
        if tipo_detectado == 'gastos':
            print("✓ Excel de GASTOS detectado")
            # Cargar hojas GASTOS y CIERRE 35
            try:
                hoja_gastos = pd.read_excel(excel_path, sheet_name='GASTOS', header=None).values.tolist()
                hoja_cierre35 = pd.read_excel(excel_path, sheet_name='CIERRE 35', header=None).values.tolist()
                new_gastos = {
                    'GASTOS': hoja_gastos,
                    'CIERRE 35': hoja_cierre35
                }
                # NUEVA LÓGICA: Combinar con datos existentes si los hay
                if datagastos is not None:
                    datagastos = merge_gastos_data(datagastos, new_gastos)
                    print("  Gastos actualizado\n")
                else:
                    datagastos = new_gastos
                    print("  Gastos cargado\n")
            except Exception as e:
                print(f"  Error cargando hojas de gastos: {e}\n")
            continue  # Importante: saltar el procesamiento normal

        data = df.values.tolist()
        
        tipo_rutas = detectar_excel_rutas(data)
        
        if tipo_rutas == 'cliente_particular':
            print("✓ Excel de CLIENTE PARTICULAR detectado")
            datalenght = len(data)
            data = [fila for fila in data if not is_nan_safe(fila, 0)]
            print(f"  Eliminados por NAN en columna A: {datalenght - len(data)} registros")
            
            nuevos_clientes = parsear_cliente_particular(data, header_rows=1)
            
            if dataclientes is not None:
                # Merge: mantener primeros, agregar nuevos
                for id_r, datos in nuevos_clientes.items():
                    if id_r not in dataclientes:
                        dataclientes[id_r] = datos
                print(f"  Cliente Particular actualizado: {len(dataclientes)} total\n")
            else:
                dataclientes = nuevos_clientes
                print("  Cliente Particular cargado\n")
            continue
        
        if tipo_rutas == 'tablero_costo_ruta':
            print("✓ Excel de TABLERO COSTO RUTA detectado")
            datalenght = len(data)
            data = [fila for fila in data if not is_nan_safe(fila, 0)]
            print(f"  Eliminados por NAN en columna A: {datalenght - len(data)} registros")
            
            nuevos_costos = parsear_tablero_costo_ruta(data, header_rows=1)
            
            if datacostos is not None:
                datacostos.extend(nuevos_costos)
                print(f"  Tablero Costo Ruta actualizado: {len(datacostos)} total\n")
            else:
                datacostos = nuevos_costos
                print("  Tablero Costo Ruta cargado\n")
            continue
        
        
        # Detectar tipo de excel
        tipo_detectado = detectar_tipo_excel(data)
        
        if tipo_detectado == 'casetas':
            print("✓ Excel de CASETAS detectado")
            data, n_norm = normalizar_fechas_casetas(data, cas_fecha_col=9, cas_hora_col=10, header_rows=1)
            data, _ = create_combined_id(data, col_a=8, col_b=9, col_c=10, header_rows=1, new_col_name='ID_COMBINADO')
            
            if datacas is not None:
                datacas = merge_excel_data(datacas, data, 1, datacas[0].index('ID_COMBINADO'))
                print("  Casetas actualizadas\n")
            else:
                datacas = data
                print("  Casetas cargadas\n")
        
        elif tipo_detectado == 'monitoreo':
            print("✓ Excel de MONITOREO detectado")
            datalenght = len(data)
            data = [fila for fila in data if not is_nan_safe(fila, 0)]
            print(f"  Eliminados por NAN en columna A: {datalenght - len(data)} registros")
            
            if datamon is not None:
                datamon = merge_excel_data(datamon, data, 5, 0)
                print("  Monitoreo actualizado\n")
            else:
                datamon = data
                print("  Monitoreo cargado\n")
        
        elif tipo_detectado == 'cierre':
            print("✓ Excel de CIERRE detectado")
            datalenght = len(data)
            data = [fila for fila in data if not is_nan_safe(fila, 1)]
            print(f"  Eliminados por NAN en columna B: {datalenght - len(data)} registros")
            
            if dataayc is not None:
                dataayc = merge_excel_data(dataayc, data, 4, 1)
                print("  Cierre actualizado\n")
            else:
                dataayc = data
                print("  Cierre cargado\n")
        
        else:
            print("✗ Excel no reconocido\n")
    
    # Mostrar resumen de lo cargado
    print("\n" + "="*50)
    print("RESUMEN DE CARGA")
    print("="*50)
    if datamon:
        print(f"✓ Monitoreo: {len(datamon)} filas")
    if datacas:
        print(f"✓ Casetas: {len(datacas)} filas")
    if dataayc:
        print(f"✓ Cierre: {len(dataayc)} filas")
    if datagastos:
        print(f"✓ Gastos: cargado")
    if dataclientes:                                         
        print(f"✓ Clientes: {len(dataclientes)} registros") 
    if datacostos:                                           
        print(f"✓ Costos Ruta: {len(datacostos)} registros")
    print("="*50 + "\n")

    print("="*50 + "\n")
    
    # Preparar datos
    print("Preparando datos...")
    if datamon:
        datamon = preparar_excel(datamon, "monitoreo")
        datamon = verificar_valores_unicos(datamon, 0, "procesoeliminar")
    
    if datacas:
        datacas = preparar_excel(datacas, "casetas")
        # Mapear CT por tracto si hay dataayc
        if dataayc:
            dataayc = preparar_excel(dataayc, "cerrados")
            dataayc = verificar_valores_unicos(dataayc, 1, "procesoeliminar")
            datacas, mapa_ct = mapear_y_reemplazar_ct_por_tracto(datacas, dataayc, cas_tracto_col=8, ayc_tracto_col=8, ayc_tr_col=13)
    
    # Construir base de datos
    print("\nConstruyendo base de datos...")
    db_manager = construir_base_datos_desde_excels(
        datamon, datacas, dataayc, datagastos,
        mon_header_rows=5, cas_header_rows=1,
        db_manager=db_manager,
        is_update=False
    )
    
    # Cargar datos de rutas en la base de datos
    if dataclientes or datacostos:
        print("\nConstruyendo tabla de rutas...")
        tabla_rutas = construir_tabla_rutas(
            dataclientes if dataclientes else {},
            datacostos if datacostos else []
        )
        db_manager.cargar_rutas(tabla_rutas)
    
    # Solo si hay rutas cargadas en la base de datos
    if len(db_manager.data.get('rutas', {})) > 0:
        verificar_comisiones_vs_rutas(db_manager)
        
    # Aunque no haya rutas nuevas, guardar los demás cambios    
    db_manager.guardar_datos()
    
    # IMPORTANTE: Limpiar variables para liberar memoria
    print("\nLiberando memoria...")
    del data, datamon, datacas, dataayc, datagastos, new_gastos, dataclientes, datacostos
    
    print("\n✓ Base de datos creada y guardada exitosamente\n")
    
    return db_manager


def actualizar_datos_existentes(db_manager):
    """
    Actualiza la base de datos existente con nuevos excels.
    Monitoreo y casetas YA NO son obligatorios.
    
    Retorna: db_manager actualizado
    """
    print("\nACTUALIZACIÓN DE DATOS")
    print("="*50)
    print("Carga archivos Excel para actualizar la base de datos")
    print("Todos los archivos son opcionales")
    print("Escribe 'go' cuando termines de cargar archivos")
    print("="*50 + "\n")
    
    data = None
    datamon = None
    datacas = None
    dataayc = None
    datagastos = None
    new_gastos = None
    dataclientes = None
    datacostos = None

    while True:
        path = input("Ruta del archivo Excel (o 'go' para terminar, 'scan dir' para carpeta)> ").strip().strip('"').strip("'")
        
        if not path:
            continue
        
        path_lower = path.strip().lower()
        
        # Comando para procesar directorio
        if path_lower in ('scan dir', 'procesar directorio', 'procesar carpeta'):
            print("Escaneando directorio del script...")
            datamon, datacas, dataayc, datagastos, dataclientes, datacostos = procesar_excels_en_directorio_inicial()
            if datamon is None and datacas is None and dataayc is None and dataclientes is None and datacostos is None:
                print("No se detectaron archivos Excel en el directorio")
                continue
            break
        
        # Comando para finalizar
        if path_lower == 'go':
            if datamon is None and datacas is None and dataayc is None and datagastos is None and dataclientes is None and datacostos is None:
                print("No se cargaron archivos nuevos. Cancelando actualización.\n")
                return db_manager
            break
        
        # Cargar archivo Excel
        excel_path = Path(path)
        if not excel_path.exists():
            print(f"El archivo no existe: {excel_path}")
            continue
        
        try:
            df = pd.read_excel(excel_path, sheet_name=0, header=None, engine='openpyxl')
        except:
            try:
                df = pd.read_excel(excel_path, sheet_name=0, header=None)
            except Exception as e:
                print(f"Error al leer el archivo: {e}")
                continue
        
        # Detectar tipo de excel (primero verificar si es gastos)
        tipo_detectado = detectar_tipo_excel_multi_hoja(excel_path)
        
        if tipo_detectado == 'gastos':
            print("✓ Excel de GASTOS detectado")
            # Cargar hojas GASTOS y CIERRE 35
            try:
                hoja_gastos = pd.read_excel(excel_path, sheet_name='GASTOS', header=None).values.tolist()
                hoja_cierre35 = pd.read_excel(excel_path, sheet_name='CIERRE 35', header=None).values.tolist()
                new_gastos = {
                    'GASTOS': hoja_gastos,
                    'CIERRE 35': hoja_cierre35
                }
                # NUEVA LÓGICA: Combinar con datos existentes si los hay
                if datagastos is not None:
                    datagastos = merge_gastos_data(datagastos, new_gastos)
                    print("  Gastos actualizado\n")
                else:
                    datagastos = new_gastos
                    print("  Gastos cargado\n")
            except Exception as e:
                print(f"  Error cargando hojas de gastos: {e}\n")
            continue  # Importante: saltar el procesamiento normal

        data = df.values.tolist()
        
        tipo_rutas = detectar_excel_rutas(data)
        
        if tipo_rutas == 'cliente_particular':
            print("✓ Excel de CLIENTE PARTICULAR detectado")
            datalenght = len(data)
            data = [fila for fila in data if not is_nan_safe(fila, 0)]
            print(f"  Eliminados por NAN en columna A: {datalenght - len(data)} registros")
            
            nuevos_clientes = parsear_cliente_particular(data, header_rows=1)
            
            if dataclientes is not None:
                # Merge: mantener primeros, agregar nuevos
                for id_r, datos in nuevos_clientes.items():
                    if id_r not in dataclientes:
                        dataclientes[id_r] = datos
                print(f"  Cliente Particular actualizado: {len(dataclientes)} total\n")
            else:
                dataclientes = nuevos_clientes
                print("  Cliente Particular cargado\n")
            continue
        
        if tipo_rutas == 'tablero_costo_ruta':
            print("✓ Excel de TABLERO COSTO RUTA detectado")
            datalenght = len(data)
            data = [fila for fila in data if not is_nan_safe(fila, 0)]
            print(f"  Eliminados por NAN en columna A: {datalenght - len(data)} registros")
            
            nuevos_costos = parsear_tablero_costo_ruta(data, header_rows=1)
            
            if datacostos is not None:
                datacostos.extend(nuevos_costos)
                print(f"  Tablero Costo Ruta actualizado: {len(datacostos)} total\n")
            else:
                datacostos = nuevos_costos
                print("  Tablero Costo Ruta cargado\n")
            continue
        
        tipo_detectado = detectar_tipo_excel(data)
        
        if tipo_detectado == 'casetas':
            print("✓ Excel de CASETAS detectado")
            data, n_norm = normalizar_fechas_casetas(data, cas_fecha_col=9, cas_hora_col=10, header_rows=1)
            data, _ = create_combined_id(data, col_a=8, col_b=9, col_c=10, header_rows=1, new_col_name='ID_COMBINADO')
            
            if datacas is not None:
                datacas = merge_excel_data(datacas, data, 1, datacas[0].index('ID_COMBINADO'))
                print("  Casetas actualizadas\n")
            else:
                datacas = data
                print("  Casetas cargadas\n")
        
        elif tipo_detectado == 'monitoreo':
            print("✓ Excel de MONITOREO detectado")
            datalenght = len(data)
            data = [fila for fila in data if not is_nan_safe(fila, 0)]
            print(f"  Eliminados por NAN: {datalenght - len(data)} registros")
            
            if datamon is not None:
                datamon = merge_excel_data(datamon, data, 5, 0)
                print("  Monitoreo actualizado\n")
            else:
                datamon = data
                print("  Monitoreo cargado\n")
        
        elif tipo_detectado == 'cierre':
            print("✓ Excel de CIERRE detectado")
            datalenght = len(data)
            data = [fila for fila in data if not is_nan_safe(fila, 1)]
            print(f"  Eliminados por NAN: {datalenght - len(data)} registros")
            
            if dataayc is not None:
                dataayc = merge_excel_data(dataayc, data, 4, 1)
                print("  Cierre actualizado\n")
            else:
                dataayc = data
                print("  Cierre cargado\n")
        
        else:
            print("✗ Excel no reconocido\n")
    
    # Preparar datos si existen
    if datamon:
        print("Preparando monitoreo...")
        datamon = preparar_excel(datamon, "monitoreo")
        datamon = verificar_valores_unicos(datamon, 0, "procesoeliminar")
    
    if datacas:
        print("Preparando casetas...")
        datacas = preparar_excel(datacas, "casetas")
        if dataayc:
            dataayc = preparar_excel(dataayc, "cerrados")
            dataayc = verificar_valores_unicos(dataayc, 1, "procesoeliminar")
            datacas, _ = mapear_y_reemplazar_ct_por_tracto(datacas, dataayc)
    
    # Actualizar base de datos
    print("\nActualizando base de datos...")
    db_manager = construir_base_datos_desde_excels(
        datamon if datamon else [],
        datacas if datacas else [],
        dataayc if dataayc else [],
        datagastos if datagastos else [],
        mon_header_rows=5, cas_header_rows=1,
        db_manager=db_manager,
        is_update=True
    )
    
    # Cargar datos de rutas en la base de datos
    if dataclientes or datacostos:
        print("\nConstruyendo tabla de rutas...")
        tabla_rutas = construir_tabla_rutas(
            dataclientes if dataclientes else {},
            datacostos if datacostos else []
        )
        
    # Solo si hay rutas cargadas en la base de datos
    if len(db_manager.data.get('rutas', {})) > 0:
        verificar_comisiones_vs_rutas(db_manager)
        
    # Aunque no haya rutas nuevas, guardar los demás cambios    
    db_manager.guardar_datos()
    
    # Limpiar memoria
    print("Liberando memoria...")
    del data, datamon, datacas, dataayc, datagastos, new_gastos, dataclientes, datacostos
    
    print("\n✓ Base de datos actualizada y guardada\n")
    
    return db_manager


def detectar_tipo_excel(data):
    """
    Detecta el tipo de excel basándose en su contenido.
    Retorna: 'casetas', 'monitoreo', 'cierre', o None
    """
    if not data or len(data) == 0:
        return None
    
    # Casetas: primera celda "Clave de Red"
    if data[0] and data[0][0] == "Clave de Red":
        return 'casetas'
    
    # Monitoreo: buscar "Reporte Monitoreo" en alguna fila
    if any("Reporte Monitoreo" in str(c) for row in data for c in row if c is not None):
        return 'monitoreo'
    
    # Cierre: buscar "Reporte Apertura/Cierre Viajes"
    if any("Reporte Apertura/Cierre Viajes" in str(c) for row in data for c in row if c is not None):
        return 'cierre'
    
    return None


def procesar_excels_en_directorio_inicial():
    """
    Escanea el directorio del script buscando archivos Excel.
    Retorna: (datamon, datacas, dataayc)
    """
    base_dir = Path(__file__).parent.resolve()
    exts = ('.xls', '.xlsx', '.xlsm', '.xlsb')
    archivos = [p for p in base_dir.iterdir() if p.is_file() and p.suffix.lower() in exts]
    
    if not archivos:
        print(f"No se encontraron archivos Excel en {base_dir}")
        return None, None, None
    
    print(f"\nProcesando {len(archivos)} archivos Excel en: {base_dir}\n")
    
    data = None
    datamon = None
    datacas = None
    dataayc = None
    datagastos = None
    new_gastos = None
    dataclientes = None
    datacostos = None
    
    for excel_path in archivos:
        #Verificar si es excel de gastos PRIMERO
        tipo_multi = detectar_tipo_excel_multi_hoja(excel_path)
        if tipo_multi == 'gastos':
            print(f"[SCAN] Excel de gastos: {excel_path.name}")
            try:
                hoja_gastos = pd.read_excel(excel_path, sheet_name='GASTOS', header=None).values.tolist()
                hoja_cierre35 = pd.read_excel(excel_path, sheet_name='CIERRE 35', header=None).values.tolist()
                new_gastos = {
                    'GASTOS': hoja_gastos,
                    'CIERRE 35': hoja_cierre35
                }
                # NUEVA LÓGICA: Combinar con datos existentes si los hay
                if datagastos is not None:
                    datagastos = merge_gastos_data(datagastos, new_gastos)
                    print("  -> Gastos actualizado")
                else:
                    datagastos = new_gastos
                    print("  -> Gastos cargado")
            except Exception as e:
                print(f"  -> Error: {e}")
            continue  # Saltar procesamiento normal

        try:
            print(f"[SCAN] Leyendo: {excel_path.name}")
            try:
                df = pd.read_excel(excel_path, sheet_name=0, header=None, engine='openpyxl')
            except:
                df = pd.read_excel(excel_path, sheet_name=0, header=None)
        except Exception as e:
            print(f"[ERROR] No se pudo leer {excel_path.name}: {e}")
            continue
        
        data = df.values.tolist()
        
        tipo_rutas = detectar_excel_rutas(data)
        
        if tipo_rutas == 'cliente_particular':
            print("✓ Excel de CLIENTE PARTICULAR detectado")
            datalenght = len(data)
            data = [fila for fila in data if not is_nan_safe(fila, 0)]
            print(f"  Eliminados por NAN en columna A: {datalenght - len(data)} registros")
            
            nuevos_clientes = parsear_cliente_particular(data, header_rows=1)
            
            if dataclientes is not None:
                # Merge: mantener primeros, agregar nuevos
                for id_r, datos in nuevos_clientes.items():
                    if id_r not in dataclientes:
                        dataclientes[id_r] = datos
                print(f"  Cliente Particular actualizado: {len(dataclientes)} total\n")
            else:
                dataclientes = nuevos_clientes
                print("  Cliente Particular cargado\n")
            continue
        
        if tipo_rutas == 'tablero_costo_ruta':
            print("✓ Excel de TABLERO COSTO RUTA detectado")
            datalenght = len(data)
            data = [fila for fila in data if not is_nan_safe(fila, 0)]
            print(f"  Eliminados por NAN en columna A: {datalenght - len(data)} registros")
            
            nuevos_costos = parsear_tablero_costo_ruta(data, header_rows=1)
            
            if datacostos is not None:
                datacostos.extend(nuevos_costos)
                print(f"  Tablero Costo Ruta actualizado: {len(datacostos)} total\n")
            else:
                datacostos = nuevos_costos
                print("  Tablero Costo Ruta cargado\n")
            continue
        
        tipo = detectar_tipo_excel(data)
        
        if tipo == 'casetas':
            print(f"  -> Excel de casetas: {excel_path.name}")
            data, _ = normalizar_fechas_casetas(data)
            data, _ = create_combined_id(data, col_a=8, col_b=9, col_c=10, header_rows=1, new_col_name='ID_COMBINADO')
            
            if datacas is not None:
                datacas = merge_excel_data(datacas, data, 1, datacas[0].index('ID_COMBINADO'))
            else:
                datacas = data
        
        elif tipo == 'monitoreo':
            print(f"  -> Excel de monitoreo: {excel_path.name}")
            data = [fila for fila in data if not is_nan_safe(fila, 0)]
            
            if datamon is not None:
                datamon = merge_excel_data(datamon, data, 5, 0)
            else:
                datamon = data
        
        elif tipo == 'cierre':
            print(f"  -> Excel de cierre: {excel_path.name}")
            data = [fila for fila in data if not is_nan_safe(fila, 1)]
            
            if dataayc is not None:
                dataayc = merge_excel_data(dataayc, data, 4, 1)
            else:
                dataayc = data
    
    print()
    return datamon, datacas, dataayc, datagastos, dataclientes, datacostos


def modulo_consulta(db_manager):
    """
    Módulo de consulta de datos.
    Opera exclusivamente sobre la base de datos.
    """
    print("\nMÓDULO DE CONSULTA")
    print("="*50)
    print("Filtros disponibles:")
    print("  - ruta <valor>")
    print("  - fecha <YYYY|YYYY-MM|YYYY-MM-DD>")
    print("  - op <operador>")
    print("  - track <tracto>")
    print("\nComandos:")
    print("  - filtros                    : Mostrar filtros activos")
    print("  - clear filters              : Limpiar filtros")
    print("  - casetas duplicadas [tracto]: Detectar casetas con <= 5min diferencia") 
    print("  - casetas sin relacion <tracto>")
    print("  - viajes irreales")
    print("  - viajes error comision")
    print("  - HID####VP##### (ID viaje)  : Ver detalles de viaje")
    print("  - viajes irreales")
    print("\nComandos de Rutas:")                                    
    print("  - rutas                      : Ver todas las rutas")    
    print("  - rutas stats                : Estadísticas de rutas")  
    print("  - ruta id <id>               : Buscar por ID de ruta")  
    print("  - ruta cliente <nombre>      : Buscar por cliente")     
    print("  - ruta tipo <tipo>           : Buscar por tipo viaje")  
    print("\nOtros:")     
    print("  - recargar                   : Actualizar BD con nuevos excels")
    print("  - salir                      : Cerrar programa")
    print("="*50 + "\n")
    
    filtros_relacionales = {
        'ruta': None,
        'op': None,
        'fecha': None,
        'track': None,
        'estado': None
    }
    
    while True:
        try:
            s = input("Consulta> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nSaliendo.")
            break
        
        if not s:
            continue
        
        s_lower = s.strip().lower()
        
        if s_lower in ('salir', 'exit', 'quit'):
            print("Finalizando programa.")
            break
        
        # Comando: rutas (mostrar todas)
        if s_lower == 'rutas':
            listar_todas_rutas_db(db_manager)
            continue
        
        # Comando: rutas stats
        if s_lower in ('rutas stats', 'rutas estadisticas', 'stats rutas'):
            estadisticas_rutas_db(db_manager)
            continue
        
        # Comando: ruta id <id>
        m = re.match(r'^ruta\s+id\s+(.+)$', s.strip(), re.IGNORECASE)
        if m:
            id_buscar = m.group(1).strip()
            # Limpiar comillas si existen
            if (id_buscar.startswith('"') and id_buscar.endswith('"')) or \
               (id_buscar.startswith("'") and id_buscar.endswith("'")):
                id_buscar = id_buscar[1:-1].strip()
            buscar_ruta_por_id_db(db_manager, id_buscar)
            continue
        
        # Comando: ruta cliente <nombre>
        m = re.match(r'^ruta\s+cliente\s+(.+)$', s.strip(), re.IGNORECASE)
        if m:
            cliente = m.group(1).strip()
            if (cliente.startswith('"') and cliente.endswith('"')) or \
               (cliente.startswith("'") and cliente.endswith("'")):
                cliente = cliente[1:-1].strip()
            buscar_rutas_db(db_manager, {'cliente': cliente})
            continue
        
        # Comando: ruta tipo <tipo>
        m = re.match(r'^ruta\s+tipo\s+(.+)$', s.strip(), re.IGNORECASE)
        if m:
            tipo = m.group(1).strip()
            if (tipo.startswith('"') and tipo.endswith('"')) or \
               (tipo.startswith("'") and tipo.endswith("'")):
                tipo = tipo[1:-1].strip()
            buscar_rutas_db(db_manager, {'tipo': tipo})
            continue
        
        # Comando: ruta origen <origen>
        m = re.match(r'^ruta\s+origen\s+(.+)$', s.strip(), re.IGNORECASE)
        if m:
            origen = m.group(1).strip()
            if (origen.startswith('"') and origen.endswith('"')) or \
               (origen.startswith("'") and origen.endswith("'")):
                origen = origen[1:-1].strip()
            buscar_rutas_db(db_manager, {'origen': origen})
            continue
        
        # Comando: viajes error comision
        if re.match(r'^viajes\s+error\s+comision', s.strip(), re.IGNORECASE):
            listar_viajes_error_comision_db(db_manager)
            continue
        
        # Comando: viajes irreales
        if re.match(r'^viajes\s+irreales$', s.strip(), re.IGNORECASE):
            listar_viajes_irreales_db(db_manager)
            continue
        
        # Detectar: "casetas duplicadas" o "casetas duplicadas 511"
        m = re.match(r'^casetas\s+duplicadas(?:\s+(.+))?$', s.strip(), re.IGNORECASE)
        if m:
            tracto_especificado = m.group(1)
            
            if tracto_especificado:
                # Limpiar comillas si existen
                tracto_raw = tracto_especificado.strip()
                if (tracto_raw.startswith('"') and tracto_raw.endswith('"')) or \
                   (tracto_raw.startswith("'") and tracto_raw.endswith("'")):
                    tracto_val = tracto_raw[1:-1].strip()
                else:
                    tracto_val = tracto_raw
                
                # Buscar solo en ese tracto
                detectar_casetas_duplicadas_por_tracto(db_manager, tracto_val)
            else:
                # Buscar en todos los tractos
                detectar_casetas_duplicadas_db(db_manager)
            
            continue
        
        # Detectar ID de viaje
        if re.match(r'^HID\d{4}VP\d+$', s.strip(), re.IGNORECASE):
            id_viaje = s.strip()
            filtros_temp = {'id': id_viaje}
             # Aplicar filtros para obtener el viaje
            viajes_encontrados = aplicar_filtros_db(db_manager, filtros_temp, single_use=True)
            # Mostrar casetas del viaje
            imprimir_casetas_viaje_db(db_manager, id_viaje)
            # Si el viaje tiene ID_RUTA, mostrar información de ruta
            if viajes_encontrados and len(viajes_encontrados) > 0:
                viaje = viajes_encontrados[0]
                id_ruta = viaje.get('id_ruta')
                
                if id_ruta:
                    # print(f"\n[INFO] Este viaje está relacionado con la ruta: {id_ruta}")
                    
                    # Buscar todos los registros de ruta que empiecen con este ID_RUTA
                    todas_rutas = db_manager.obtener_todas_rutas()
                    rutas_relacionadas = [
                        r for r in todas_rutas 
                        if r.get('id_ruta') == id_ruta
                    ]
                    
                    if rutas_relacionadas:
                        imprimir_rutas_relacionadas(rutas_relacionadas, id_ruta)
                    else:
                        print(f"  ⚠️  No se encontró información detallada de la ruta '{id_ruta}' en la tabla de rutas.")
                        print(f"     Es posible que falten los excel de Cliente Particular o Tablero Costo Ruta.\n")
                else:
                    print(f"\n[INFO] Este viaje no tiene ID_RUTA asociado (no fue enriquecido con datos de cierre).\n")
            
            continue
        
        # Comandos de filtros
        m = re.match(r'^(ruta|op|fecha|track|estado)\s+(.+)$', s.strip(), re.IGNORECASE)
        if m:
            clave = m.group(1).lower()
            valor_raw = m.group(2).strip()
            
            if (valor_raw.startswith('"') and valor_raw.endswith('"')) or \
               (valor_raw.startswith("'") and valor_raw.endswith("'")):
                valor = valor_raw[1:-1].strip()
            else:
                valor = valor_raw
            
            if clave == 'ruta':
                filtros_relacionales['ruta'] = valor if valor != '' else None
                print(f"Filtro 'ruta' establecido: {filtros_relacionales['ruta']}")
            elif clave == 'op':
                filtros_relacionales['op'] = valor if valor != '' else None
                print(f"Filtro 'operador' establecido: {filtros_relacionales['op']}")
            elif clave == 'fecha':
                if re.match(r'^\d{4}(-\d{2})?(-\d{2})?$', valor):
                    filtros_relacionales['fecha'] = valor
                    print(f"Filtro 'fecha' establecido: {filtros_relacionales['fecha']}")
                else:
                    print("Formato de fecha inválido. Usa YYYY, YYYY-MM o YYYY-MM-DD.")
            elif clave == 'track':
                filtros_relacionales['track'] = valor
                print(f"Filtro 'track' establecido: {filtros_relacionales['track']}")
            elif clave == 'estado':
                # Validar que sea un estado válido
                valor_upper = valor.upper()
                if valor_upper in ['IRREAL', 'ERR. COM', 'ERR.COM']:
                    # Normalizar ERR.COM a ERR. COM
                    if valor_upper == 'ERR.COM':
                        valor_upper = 'ERR. COM'
                    filtros_relacionales['estado'] = valor_upper
                    print(f"Filtro 'estado' establecido: {filtros_relacionales['estado']}")
                else:
                    print("Estado inválido. Usa: IRREAL o 'ERR. COM'")
                    continue
            
            aplicar_filtros_db(db_manager, filtros_relacionales)
            continue
        
        # Mostrar filtros
        if s_lower in ('filtros', 'show filtros', 'show filters'):
            print("\nFiltros activos:")
            for k, v in filtros_relacionales.items():
                print(f"  {k}: {v}")
            print()
            continue
        
        # Limpiar filtros
        if s_lower in ('clear filtros', 'clear filters', 'reset filtros', 'reset filters'):
            for k in filtros_relacionales.keys():
                filtros_relacionales[k] = None
            print("Filtros limpiados.")
            aplicar_filtros_db(db_manager, filtros_relacionales)
            continue
        
        # Casetas sin relación
        m = re.match(r'^casetas\s+sin\s+relacion\s+(.+)$', s.strip(), re.IGNORECASE)
        if m:
            tracto_raw = m.group(1).strip()
            if (tracto_raw.startswith('"') and tracto_raw.endswith('"')) or \
               (tracto_raw.startswith("'") and tracto_raw.endswith("'")):
                tracto_val = tracto_raw[1:-1].strip()
            else:
                tracto_val = tracto_raw
            
            if not tracto_val:
                print("Debes indicar el tracto. Ejemplo: casetas sin relacion 215")
                continue
            
            listar_casetas_no_relacionadas_db(db_manager, tracto_val)
            continue
        
        # Recargar datos
        if s_lower in ('recargar', 'reload'):
            db_manager = actualizar_datos_existentes(db_manager)
            continue
        
        print("Comando no reconocido. Escribe un filtro, comando o ID de viaje.")


# Funciones actualizadas sin dependencias de datacas

def imprimir_casetas_viaje_db(db_manager, id_viaje):
    """
    Imprime casetas de un viaje usando SOLO la base de datos.
    """
    casetas_dentro = db_manager.obtener_casetas_por_viaje(id_viaje, solo_fuera=False)
    casetas_fuera = db_manager.obtener_casetas_por_viaje(id_viaje, solo_fuera=True)
    
    if not casetas_dentro and not casetas_fuera:
        print(f"\nNo hay casetas asociadas al viaje '{id_viaje}'\n")
        return
    
    registros_formateados = []
    
    for caseta in casetas_dentro:
        registros_formateados.append({
            'tracto': caseta.get('tracto', ''),
            'fecha': caseta.get('fecha', ''),
            'hora': caseta.get('hora', ''),
            'importe': caseta.get('importe', 0),
            'nombre': caseta.get('nombre', ''),
            'clase': caseta.get('clase', ''),
            'id_combinado': caseta.get('id_caseta', ''),
            'tarjeta': caseta.get('tarjeta', ''),
            'status': 'Cargado'
        })
    
    for caseta in casetas_fuera:
        registros_formateados.append({
            'tracto': caseta.get('tracto', ''),
            'fecha': caseta.get('fecha', ''),
            'hora': caseta.get('hora', ''),
            'importe': caseta.get('importe', 0),
            'nombre': caseta.get('nombre', ''),
            'clase': caseta.get('clase', ''),
            'id_combinado': caseta.get('id_caseta', ''),
            'tarjeta': caseta.get('tarjeta', ''),
            'status': 'Descargado'
        })
    
    print(f"\nCasetas del viaje {id_viaje}:")
    imprimir_lista_casetas(registros_formateados)


def listar_casetas_no_relacionadas_db(db_manager, tracto_filter):
    """
    Muestra casetas sin relación usando SOLO la base de datos.
    """
    ids_relacionados = set(r['id_caseta'] for r in db_manager.data['relaciones'])
    
    resultados = []
    
    for id_caseta, caseta in db_manager.data['casetas'].items():
        tracto = caseta.get('tracto', '')
        
        if tracto_filter not in str(tracto):
            continue
        
        if id_caseta in ids_relacionados:
            continue
        
        resultados.append({
            'tracto': tracto,
            'fecha': caseta.get('fecha', ''),
            'hora': caseta.get('hora', ''),
            'importe': caseta.get('importe', 0),
            'nombre': caseta.get('nombre', ''),
            'clase': caseta.get('clase', ''),
            'id_combinado': id_caseta,
            'tarjeta': caseta.get('tarjeta', ''),
            'status': 'Sin relación'
        })
    
    if not resultados:
        print(f"\nNo se encontraron casetas sin relación para tracto '{tracto_filter}'.\n")
        return resultados
    
    imprimir_lista_casetas(resultados)
    return resultados


if __name__ == "__main__":
    main()