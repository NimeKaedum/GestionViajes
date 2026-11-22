#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interfaz Gráfica para el Sistema de Viajes
Versión con lógica real integrada
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import pandas as pd
from pathlib import Path
from tksheet import Sheet
from decimal import Decimal
from datetime import datetime 

# NUEVA IMPORTACIÓN
try:
    from tkcalendar import DateEntry
    TKCALENDAR_AVAILABLE = True
except ImportError:
    TKCALENDAR_AVAILABLE = False
    print("⚠️ tkcalendar no disponible. Instala con: pip install tkcalendar")

# Intentar importar TkinterDnD
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TKDND_AVAILABLE = True
except ImportError:
    TKDND_AVAILABLE = False
    print("⚠️ tkinterdnd2 no disponible. Instala con: pip install tkinterdnd2")

# ==================== IMPORTAR FUNCIONES DEL SISTEMA ====================
from core_viajes import (
    DatabaseManager,
    detectar_tipo_excel,
    detectar_tipo_excel_multi_hoja,
    detectar_excel_rutas,
    preparar_excel,
    verificar_valores_unicos,
    normalizar_fechas_casetas,
    create_combined_id,
    merge_excel_data,
    mapear_y_reemplazar_ct_por_tracto,
    construir_base_datos_desde_excels,
    parsear_cliente_particular,
    parsear_tablero_costo_ruta,
    construir_tabla_rutas,
    merge_clientes_particulares,
    merge_tablero_costos,
    detectar_viajes_dobles,
    detectar_estados_analisis,
    marcar_viajes_finales,
    merge_gastos_data,
    detectar_subtipo_monitoreo,
    detectar_subtipo_apertura_cierre,
    estandarizar_monitoreo,
    estandarizar_apertura_cierre,
    preparar_excel_gastos_adicionales,
    procesar_hojas_operadores,
    merge_gastos_adicionales
)

def format_date_safe(dt):
    """Helper para formatear fechas de manera segura"""
    if dt is None or pd.isna(dt):
        return ''
    try:
        return pd.to_datetime(dt).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return str(dt)


class AutocompleteCombobox(ttk.Combobox):
    """
    Combobox con autocompletado que filtra opciones mientras el usuario escribe.
    """
    
    def __init__(self, parent, completevalues=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        
        self._completevalues = completevalues or []
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self._dropdown_open = False
        
        # Configurar eventos
        self.bind('<KeyRelease>', self.handle_keyrelease)
        self.bind('<<ComboboxSelected>>', self.handle_selection)
        self.bind('<FocusIn>', self.handle_focus_in)
        
        # Configurar valores iniciales
        self['values'] = self._completevalues
    
    def set_completion_list(self, completion_list):
        """Actualiza la lista de valores para autocompletado"""
        self._completevalues = sorted(completion_list, key=str.lower)
        self['values'] = self._completevalues
    
    def autocomplete(self, delta=0):
        """Autocompletado: filtra valores según lo escrito"""
        if delta:  # Si se presionó flecha arriba/abajo
            self.delete(self.position, tk.END)
        else:
            self.position = len(self.get())
        
        # Obtener texto actual
        text = self.get()
        
        if not text:
            # Si está vacío, mostrar todas las opciones
            self['values'] = self._completevalues
            return
        
        # Filtrar valores que contengan el texto (case-insensitive)
        _hits = []
        for element in self._completevalues:
            if text.lower() in element.lower():
                _hits.append(element)
        
        # Actualizar lista de coincidencias
        if _hits != self._hits:
            self._hit_index = 0
            self._hits = _hits
        
        # Mostrar coincidencias
        if _hits:
            self['values'] = _hits
            
            # CORREGIDO: Abrir dropdown sin robar el foco
            if not self._dropdown_open:
                # Usar after para abrir el dropdown después de que el evento actual termine
                self.after(1, self._open_dropdown)
        else:
            self['values'] = []
    
    def _open_dropdown(self):
        """Abre el dropdown manteniendo el foco en el input"""
        try:
            # Guardar posición del cursor
            cursor_pos = self.index(tk.INSERT)
            
            # Abrir dropdown
            self.event_generate('<Button-1>')
            self.event_generate('<ButtonRelease-1>')
            
            # Restaurar foco y posición del cursor inmediatamente
            self.focus_set()
            self.icursor(cursor_pos)
            
            self._dropdown_open = True
        except:
            pass
    
    def handle_keyrelease(self, event):
        """Maneja eventos de teclado"""
        # Ignorar teclas especiales
        if event.keysym in ('BackSpace', 'Left', 'Right', 'Up', 'Down', 
                           'Home', 'End', 'Return', 'Tab', 'Escape'):
            if event.keysym in ('Return', 'Escape'):
                self._dropdown_open = False
            return
        
        # Realizar autocompletado
        self.autocomplete()
    
    def handle_selection(self, event):
        """Maneja la selección de un elemento del dropdown"""
        self.position = len(self.get())
        self._dropdown_open = False
        # Mantener foco en el combobox después de selección
        self.focus_set()
    
    def handle_focus_in(self, event):
        """Maneja cuando el combobox recibe el foco"""
        self._dropdown_open = False

class AplicacionViajes:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Gestión de Viajes")
        self.root.geometry("900x750")
        
        # Variable para saber si DnD está disponible
        self.dnd_enabled = TKDND_AVAILABLE
        
        # Variables para almacenar datos (listas de filas de Excel)
        self.datamon = None
        self.datacas = None
        self.dataayc = None
        self.datagastos = None
        self.dataclientes = None
        self.datacostos = None
        self.dataadicional = None
        
        # Database manager
        self.db_manager = None
        
        # Modo de operación: 'inicial' o 'actualizacion'
        self.modo = 'inicial'
        
        # Inicializar widgets como None
        self.console = None
        self.btn_continuar = None
        self.tree = None
        self.drop_label = None

        # Lista para trackear ventanas de detalles abiertas
        self.ventanas_detalles = []
        
        # MODIFICADO: Verificar BD y actuar automáticamente
        db_path = Path('database.json')
        if db_path.exists():
            # Hay BD → Cargar y abrir consulta directamente
            self.cargar_bd_y_abrir_consulta()
        else:
            # No hay BD → Ir a carga inicial
            self.crear_interfaz_carga()
    
    def log_inicial(self, mensaje):
        """Log antes de que exista la consola"""
        print(f"> {mensaje}")
    
    def cargar_bd_y_abrir_consulta(self):
        """Carga la BD existente y abre el módulo de consulta"""
        try:
            # Crear interfaz mínima para mostrar progreso
            frame = ttk.Frame(self.root, padding="20")
            frame.pack(expand=True)
            
            label = ttk.Label(frame, text="Cargando base de datos...", font=('Arial', 14))
            label.pack(pady=20)
            
            progress = ttk.Progressbar(frame, mode='indeterminate', length=300)
            progress.pack(pady=10)
            progress.start()
            
            self.root.update()
            
            # Cargar base de datos
            self.db_manager = DatabaseManager('database.json')
            
            progress.stop()
            
            # Abrir módulo de consulta
            self.abrir_modulo_consulta()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error cargando base de datos:\n{e}")
            # Si falla, crear interfaz de carga normal
            frame.destroy()
            self.crear_interfaz_carga()
    
    def crear_interfaz_carga(self):
        """Crea la interfaz del módulo de carga de exceles"""
        
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título según modo
        titulo_text = "CARGA INICIAL DE DATOS" if self.modo == 'inicial' else "ACTUALIZACIÓN DE DATOS"
        titulo = ttk.Label(main_frame, text=titulo_text, font=('Arial', 16, 'bold'))
        titulo.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Subtítulo informativo según modo
        if self.modo == 'inicial':
            info_text = "Se requieren: Monitoreo y Casetas (obligatorios)"
        else:
            info_text = "Modo actualización: Todos los archivos son opcionales"
        
        info_label = ttk.Label(main_frame, text=info_text, 
                              font=('Arial', 9), foreground='#666')
        info_label.grid(row=1, column=0, columnspan=2, pady=2)
        
        # Instrucciones
        instrucciones_text = "Arrastra archivos Excel aquí o usa los botones" if self.dnd_enabled else "Usa los botones para seleccionar archivos"
        instrucciones = ttk.Label(main_frame, text=instrucciones_text, font=('Arial', 10))
        instrucciones.grid(row=2, column=0, columnspan=2, pady=5)
        
        # Área de consola primero
        self.crear_consola(main_frame)
        
        # Área de drag & drop
        self.crear_zona_drop(main_frame)
        
        # Lista de archivos
        self.crear_lista_archivos(main_frame)
        
        # Botones de acción
        self.crear_botones_accion(main_frame)
        
        # Configurar grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
    
    def crear_zona_drop(self, parent):
        """Crea zona para arrastrar archivos"""
        drop_frame = ttk.LabelFrame(parent, text="Zona de Arrastre", padding="20")
        drop_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        # Texto según disponibilidad de DnD
        texto = "📁 Arrastra archivos Excel aquí\n\n(Soporta múltiples archivos)" if self.dnd_enabled else "📁 Drag & Drop no disponible\n\nUsa 'Seleccionar Archivos'"
        
        self.drop_label = tk.Label(drop_frame, 
            text=texto,
            bg="#e8f4f8" if self.dnd_enabled else "#f0f0f0",
            relief=tk.RIDGE, 
            borderwidth=2,
            font=('Arial', 11),
            height=4)
        self.drop_label.pack(fill=tk.BOTH, expand=True)
        
        # Habilitar drag & drop solo si está disponible
        if self.dnd_enabled:
            try:
                self.drop_label.drop_target_register(DND_FILES)
                self.drop_label.dnd_bind('<<Drop>>', self.archivos_arrastrados)
                self.drop_label.bind('<Enter>', lambda e: self.drop_label.config(bg="#d0e8f0"))
                self.drop_label.bind('<Leave>', lambda e: self.drop_label.config(bg="#e8f4f8"))
            except Exception as e:
                self.log(f"⚠️ Error configurando drag & drop: {e}")
                self.dnd_enabled = False
    
    def crear_lista_archivos(self, parent):
        """Crea lista de archivos cargados"""
        list_frame = ttk.LabelFrame(parent, text="Archivos Cargados", padding="10")
        list_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        columns = ('tipo', 'nombre', 'registros', 'estado')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=8)
        
        self.tree.heading('tipo', text='Tipo')
        self.tree.heading('nombre', text='Nombre del Archivo')
        self.tree.heading('registros', text='Registros')
        self.tree.heading('estado', text='Estado')
        
        self.tree.column('tipo', width=150)
        self.tree.column('nombre', width=350)
        self.tree.column('registros', width=100)
        self.tree.column('estado', width=150)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
    
    def crear_botones_accion(self, parent):
        """Crea botones de acción"""
        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=10)
        
        btn_seleccionar = ttk.Button(btn_frame, text="📂 Seleccionar Archivos", 
                                     command=self.seleccionar_archivos)
        btn_seleccionar.grid(row=0, column=0, padx=5)
        
        btn_escanear = ttk.Button(btn_frame, text="🔍 Escanear Carpeta", 
                                 command=self.escanear_directorio)
        btn_escanear.grid(row=0, column=1, padx=5)
        
        btn_limpiar = ttk.Button(btn_frame, text="🗑️ Limpiar", 
                                command=self.limpiar_archivos)
        btn_limpiar.grid(row=0, column=2, padx=5)
        
        self.btn_continuar = ttk.Button(btn_frame, text="✅ Procesar y Continuar", 
                                       command=self.procesar_archivos,
                                       state=tk.DISABLED)
        self.btn_continuar.grid(row=0, column=3, padx=5)
    
    def crear_consola(self, parent):
        """Crea área de consola/log"""
        console_frame = ttk.LabelFrame(parent, text="Consola", padding="5")
        console_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        self.console = tk.Text(console_frame, height=8, width=80, 
                              bg='#1e1e1e', fg='#00ff00', 
                              font=('Consolas', 9))
        console_scroll = ttk.Scrollbar(console_frame, orient=tk.VERTICAL, 
                                      command=self.console.yview)
        self.console.configure(yscroll=console_scroll.set)
        
        self.console.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        console_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        console_frame.columnconfigure(0, weight=1)
        console_frame.rowconfigure(0, weight=1)
        
        self.log("Sistema de Gestión de Viajes - Inicializado")
        self.log("Carga archivos Excel para comenzar...")
    
    def log(self, mensaje):
        """Agrega mensaje a la consola"""
        if self.console is None:
            print(f"> {mensaje}")
            return
        self.console.insert(tk.END, f"> {mensaje}\n")
        self.console.see(tk.END)
        self.console.update()
    
    def archivos_arrastrados(self, event):
        """Maneja archivos arrastrados"""
        try:
            archivos = self.root.tk.splitlist(event.data)
        except Exception as e:
            self.log(f"✗ Error leyendo archivos: {e}")
            return
        
        for archivo in archivos:
            archivo = archivo.strip('{}').strip()
            if archivo and os.path.exists(archivo):
                self.procesar_archivo(archivo)
    
    def seleccionar_archivos(self):
        """Abre diálogo para seleccionar archivos"""
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos Excel",
            filetypes=[
                ("Excel files", "*.xlsx *.xls *.xlsm *.xlsb"),
                ("All files", "*.*")
            ]
        )
        
        for archivo in archivos:
            self.procesar_archivo(archivo)
    
    def escanear_directorio(self):
        """Escanea el directorio actual buscando exceles"""
        self.log("Escaneando directorio actual...")
        try:
            base_dir = Path(__file__).parent.resolve()
        except:
            base_dir = Path.cwd()
        
        exts = ('.xls', '.xlsx', '.xlsm', '.xlsb')
        archivos = [p for p in base_dir.iterdir() 
                   if p.is_file() and p.suffix.lower() in exts]
        
        if not archivos:
            self.log(f"No se encontraron archivos Excel")
            messagebox.showinfo("Escaneo", "No se encontraron archivos Excel")
            return
        
        self.log(f"Encontrados {len(archivos)} archivos")
        for archivo in archivos:
            self.procesar_archivo(str(archivo))
    
    def procesar_archivo(self, ruta_archivo):
        """Procesa un archivo Excel REAL usando la lógica del sistema"""
        try:
            nombre = os.path.basename(ruta_archivo)
            self.log(f"Leyendo: {nombre}")
            
            # Leer archivo Excel
            try:
                df = pd.read_excel(ruta_archivo, sheet_name=0, header=None, engine='openpyxl')
            except:
                df = pd.read_excel(ruta_archivo, sheet_name=0, header=None)
            
            data = df.values.tolist()
            
            # DETECCIÓN DE TIPO - Primero verificar si es multi-hoja (gastos)
            tipo_multi = detectar_tipo_excel_multi_hoja(ruta_archivo)
            
            if tipo_multi == 'gastos_adicionales':
                # Procesar gastos adicionales
                self.log(f"✓ Excel de GASTOS ADICIONALES detectado")
                
                # Leer la primera hoja que inicie con "BD"
                xls = pd.ExcelFile(ruta_archivo, engine='openpyxl')
                hoja_bd = None
                
                for sheet_name in xls.sheet_names:
                    if sheet_name.upper().startswith('BD'):
                        hoja_bd = sheet_name
                        break
                
                if hoja_bd:
                    df = pd.read_excel(ruta_archivo, sheet_name=hoja_bd, header=None, engine='openpyxl')
                    data = df.values.tolist()
                    
                    # Preparar datos (eliminar "SIN VIAJE" y vacíos)
                    data = preparar_excel_gastos_adicionales(data, header_rows=1)
                    
                    if self.dataadicional is not None:
                        # Merge con datos existentes
                        self.dataadicional = merge_gastos_adicionales(
                            self.dataadicional, data, header_rows=1
                        )
                        self.log("  Gastos adicionales actualizados")
                    else:
                        self.dataadicional = data
                        self.log("  Gastos adicionales cargados")
                
                self.tree.insert('', tk.END, values=('Gastos adicionales', nombre, 
                    len(self.dataadicional) if self.dataadicional is not None else 0, '✓'))
                self.actualizar_estado_botones()
                return

            if tipo_multi == 'gastos':
    
                self.log(f"  → Tipo: Gastos (multi-hoja)")
                hoja_gastos = pd.read_excel(ruta_archivo, sheet_name='GASTOS', header=None).values.tolist()
                hoja_cierre35 = pd.read_excel(ruta_archivo, sheet_name='CIERRE 35', header=None).values.tolist()
                
                # NUEVO: Procesar hojas de operadores (hojas que inician con "O")
                self.log("  Procesando hojas de operadores...")
                modificaciones_comision = procesar_hojas_operadores(ruta_archivo)
                
                nuevos_gastos = {
                    'GASTOS': hoja_gastos,
                    'CIERRE 35': hoja_cierre35
                }
                
                # NUEVO: Añadir modificaciones de comisión si existen
                if modificaciones_comision:
                    nuevos_gastos['MODIFICACIONES_COMISION'] = modificaciones_comision
                    self.log(f"  Comisiones modificadas: {len(modificaciones_comision)}")
                
                # ACTUALIZADO: Usar merge real
                if self.datagastos is not None:
                    self.datagastos = merge_gastos_data(self.datagastos, nuevos_gastos)
                else:
                    self.datagastos = nuevos_gastos
                
                self.tree.insert('', tk.END, values=('Gastos', nombre, 
                    f"{len(hoja_gastos) + len(hoja_cierre35)}", '✓'))
                self.actualizar_estado_botones()
                return
            
            # Detectar excel de rutas
            tipo_rutas = detectar_excel_rutas(data)
            
            if tipo_rutas == 'cliente_particular':
                self.log(f"  → Tipo: Cliente Particular")
                data_limpia = [f for f in data if len(f) > 0 and not pd.isna(f[0])]
                nuevos_clientes = parsear_cliente_particular(data_limpia, header_rows=1)
                
                # ACTUALIZADO: Usar merge con lógica reemplazar + mantener
                if self.dataclientes is not None:
                    self.dataclientes = merge_clientes_particulares(self.dataclientes, nuevos_clientes)
                else:
                    self.dataclientes = nuevos_clientes
                
                self.tree.insert('', tk.END, values=('Cliente Particular', nombre, 
                    len(self.dataclientes), '✓'))
                self.actualizar_estado_botones()
                return
            
            if tipo_rutas == 'tablero_costo_ruta':
                self.log(f"  → Tipo: Tablero Costo Ruta")
                data_limpia = [f for f in data if len(f) > 0 and not pd.isna(f[0])]
                nuevos_costos = parsear_tablero_costo_ruta(data_limpia, header_rows=1)
                
                # ACTUALIZADO: Usar merge con lógica reemplazar + mantener
                if self.datacostos is not None:
                    self.datacostos = merge_tablero_costos(self.datacostos, nuevos_costos)
                else:
                    self.datacostos = nuevos_costos
                
                self.tree.insert('', tk.END, values=('Tablero Costo', nombre, 
                    len(self.datacostos), '✓'))
                self.actualizar_estado_botones()
                return
            
            # Detección normal (monitoreo, casetas, cierre)
            tipo = detectar_tipo_excel(data)
            
            # ESTANDARIZACIÓN ANTES DE PROCESAR SEGÚN SUBTIPO
            # Se aplica solo a monitoreo y cierre cuando la estructura detectada
            # corresponde a una variante nueva; así el resto del pipeline permanece igual.
            if tipo == 'monitoreo':
                subtipo = detectar_subtipo_monitoreo(data)
                if subtipo:
                    self.log(f"  → Formato monitoreo: {subtipo}")
                    data = estandarizar_monitoreo(data, subtipo)
            elif tipo == 'cierre':
                subtipo = detectar_subtipo_apertura_cierre(data)
                if subtipo:
                    self.log(f"  → Formato apertura/cierre: {subtipo}")
                    data = estandarizar_apertura_cierre(data, subtipo)
            
            if tipo == 'monitoreo':
                self.log(f"  → Tipo: Monitoreo")
                data = preparar_excel(data, "monitoreo")
                data = verificar_valores_unicos(data, 0, "procesoeliminar")
                
                # ACTUALIZADO: merge ahora usa lógica reemplazar + mantener
                if self.datamon is not None:
                    self.datamon = merge_excel_data(self.datamon, data, 5, 0)
                else:
                    self.datamon = data
                
                self.tree.insert('', tk.END, values=('Monitoreo', nombre, 
                    len(data)-5, '✓'))
                
            elif tipo == 'casetas':
                self.log(f"  → Tipo: Casetas")
                data, _ = normalizar_fechas_casetas(data)
                data, _ = create_combined_id(data, col_a=8, col_b=9, col_c=10, 
                                            header_rows=1, new_col_name='ID_COMBINADO')
                data = preparar_excel(data, "casetas")
                
                # ACTUALIZADO: merge ahora usa lógica reemplazar + mantener
                if self.datacas is not None:
                    id_col_idx = self.datacas[0].index('ID_COMBINADO')
                    self.datacas = merge_excel_data(self.datacas, data, 1, id_col_idx)
                else:
                    self.datacas = data
                
                self.tree.insert('', tk.END, values=('Casetas', nombre, 
                    len(self.datacas)-1, '✓'))
                
            elif tipo == 'cierre':
                self.log(f"  → Tipo: Cierre")
                data = preparar_excel(data, "cerrados")
                data = verificar_valores_unicos(data, 1, "procesoeliminar")
                
                # ACTUALIZADO: merge ahora usa lógica reemplazar + mantener
                if self.dataayc is not None:
                    self.dataayc = merge_excel_data(self.dataayc, data, 4, 1)
                else:
                    self.dataayc = data
                
                self.tree.insert('', tk.END, values=('Cierre', nombre, 
                    len(self.dataayc)-4, '✓'))
            else:
                self.log(f"  ✗ Tipo no reconocido")
                self.tree.insert('', tk.END, values=('???', nombre, '?', '✗'))
                return
            
            self.actualizar_estado_botones()
            
        except Exception as e:
            import traceback
            self.log(f"✗ Error: {str(e)}")
            print(traceback.format_exc())
            messagebox.showerror("Error", f"Error procesando:\n{str(e)}")
    
    def actualizar_estado_botones(self):
        """Actualiza el estado del botón continuar según el modo"""
        if self.modo == 'inicial':
            # Modo inicial: requiere monitoreo y casetas
            tiene_obligatorios = self.datamon is not None and self.datacas is not None
            
            if self.btn_continuar:
                self.btn_continuar.config(state=tk.NORMAL if tiene_obligatorios else tk.DISABLED)
            
            if tiene_obligatorios:
                self.log("✓ Archivos obligatorios listos")
        else:
            # Modo actualización: cualquier archivo es suficiente
            tiene_datos = any([
                self.datamon is not None,
                self.datacas is not None,
                self.dataayc is not None,
                self.datagastos is not None,
                self.dataclientes is not None,
                self.datacostos is not None
            ])
            
            if self.btn_continuar:
                self.btn_continuar.config(state=tk.NORMAL if tiene_datos else tk.DISABLED)
            
            if tiene_datos:
                self.log("✓ Datos listos para actualizar")
    
    def limpiar_archivos(self):
        """Limpia todos los archivos cargados"""
        if messagebox.askyesno("Confirmar", "¿Limpiar archivos cargados?"):
            self.tree.delete(*self.tree.get_children())
            self.datamon = None
            self.datacas = None
            self.dataayc = None
            self.datagastos = None
            self.dataclientes = None
            self.datacostos = None
            self.dataadicional = None
            self.btn_continuar.config(state=tk.DISABLED)
            self.log("Archivos limpiados")
    
    def procesar_archivos(self):
        """Procesa archivos y construye/actualiza base de datos"""
        try:
            self.log("="*50)
            if self.modo == 'inicial':
                self.log("INICIANDO CARGA INICIAL")
            else:
                self.log("INICIANDO ACTUALIZACIÓN")
            self.log("="*50)
            
            # Preparar casetas si hay cierre (para mapeo CT)
            if self.datacas and self.dataayc:
                self.log("Mapeando CT por tracto...")
                self.datacas, _ = mapear_y_reemplazar_ct_por_tracto(
                    self.datacas, self.dataayc, 
                    cas_tracto_col=8, ayc_tracto_col=8, ayc_tr_col=13
                )
            
            # Crear o cargar database manager
            if self.db_manager is None:
                self.log("Inicializando base de datos...")
                self.db_manager = DatabaseManager('database.json')
            else:
                self.log("Usando base de datos existente...")
            
            # Construir/actualizar base de datos
            self.log("Procesando datos...")
            self.root.update()
            
            # Determinar si es actualización según el modo
            is_update = (self.modo == 'actualizacion')
            
            self.db_manager = construir_base_datos_desde_excels(
                self.datamon if self.datamon else [],
                self.datacas if self.datacas else [],
                self.dataayc if self.dataayc else [],
                self.datagastos if self.datagastos else [],
                mon_header_rows=5,
                cas_header_rows=1,
                db_manager=self.db_manager,
                is_update=is_update,
                dataadicional=self.dataadicional if self.dataadicional else []
            )
            
            # Cargar/actualizar rutas si existen
            if self.dataclientes or self.datacostos:
                self.log("Procesando tabla de rutas...")
                tabla_rutas = construir_tabla_rutas(
                    self.dataclientes if self.dataclientes else {},
                    self.datacostos if self.datacostos else []
                )
                self.db_manager.cargar_rutas(tabla_rutas)
            
            self.log("="*50)
            if self.modo == 'inicial':
                self.log("✓ CARGA INICIAL COMPLETADA")
            else:
                self.log("✓ ACTUALIZACIÓN COMPLETADA")
            self.log("="*50)
            
            # Abrir módulo de consulta
            self.abrir_modulo_consulta()
            
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            self.log(f"✗ ERROR CRÍTICO:")
            self.log(error_msg)
            messagebox.showerror("Error", f"Error procesando:\n{str(e)}")
    
    def abrir_modulo_consulta(self):
        """Abre el módulo de consulta con filtros y vista de datos"""
        self.root.withdraw()
        
        # Crear ventana de consulta
        self.ventana_consulta = tk.Toplevel(self.root)
        self.ventana_consulta.title("Módulo de Consulta - Gestión de Viajes")
        self.ventana_consulta.geometry("1400x900")
        
        # Frame principal
        main_frame = ttk.Frame(self.ventana_consulta, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ==================== SECCIÓN DE FILTROS ====================
        self.crear_seccion_filtros(main_frame)
        
        # ==================== SECCIÓN DE RESULTADOS ====================
        self.crear_seccion_resultados(main_frame)
        
        # ==================== BARRA DE ESTADO ====================
        self.crear_barra_estado(main_frame)
        
        # ==================== BOTONES DE ACCIÓN ====================
        self.crear_botones_principales(main_frame)
        
        # Cargar datos iniciales (todos los viajes)
        #self.aplicar_filtros_consulta()
        
        # Manejar cierre de ventana
        self.ventana_consulta.protocol("WM_DELETE_WINDOW", self.cerrar_modulo_consulta_completo)
    
    def cerrar_modulo_consulta_completo(self):
        """Cierra el módulo de consulta y todas las ventanas de detalles"""
        # Cerrar todas las ventanas de detalles abiertas
        for ventana in self.ventanas_detalles[:]:  # Copiar lista para iterar
            try:
                ventana.destroy()
            except:
                pass
        
        self.ventanas_detalles.clear()
        
        # Cerrar aplicación
        self.root.quit()

    def crear_seccion_filtros(self, parent):
        """Crea la sección de filtros en la parte superior"""
        filtros_frame = ttk.LabelFrame(parent, text="Filtros de Búsqueda", padding="10")
        filtros_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Frame para organizar filtros en grid
        grid_frame = ttk.Frame(filtros_frame)
        grid_frame.pack(fill=tk.X)
        
        # Variables de filtros
        self.filtro_activo = {}
        self.filtro_valor = {}
        self.filtro_combos = {}
        self.filtro_entries = {}
        self.filtro_calendarios = {}  # NUEVO: Para almacenar DateEntry widgets
        
        # Obtener valores únicos de la BD
        viajes = self.db_manager.data['viajes']
        
        # Rutas únicas
        rutas_set = set()
        for v in viajes.values():
            ruta = v.get('ruta')
            if ruta and str(ruta).strip():
                rutas_set.add(str(ruta).strip())
        lista_rutas = sorted(list(rutas_set))
        
        # Operadores únicos
        operadores_set = set()
        for v in viajes.values():
            op = v.get('operador')
            if op and str(op).strip():
                operadores_set.add(str(op).strip())
        lista_operadores = sorted(list(operadores_set))
        
        # Tractos únicos
        tractos_set = set()
        for v in viajes.values():
            tracto = v.get('tracto')
            if tracto and str(tracto).strip():
                tractos_set.add(str(tracto).strip())
        lista_tractos = sorted(list(tractos_set))
        
        # Clientes únicos (de rutas)
        clientes_set = set()
        rutas_db = self.db_manager.obtener_todas_rutas()
        for ruta in rutas_db:
            cliente = ruta.get('cliente')
            if cliente and str(cliente).strip():
                clientes_set.add(str(cliente).strip())
        lista_clientes = sorted(list(clientes_set))
        
        # Ensambles (opciones fijas)
        lista_ensambles = ['Todos', 'SENCILLO', 'FULL']
        
        # ==================== FILA 0: Ruta y Operador ====================
        row = 0
        
        # Filtro RUTA (con autocompletado)
        self.crear_filtro_combobox(grid_frame, 'ruta', 'Ruta', lista_rutas, row, 0, 
                                width=80, autocomplete=True)
        
        # Filtro OPERADOR (con autocompletado)
        self.crear_filtro_combobox(grid_frame, 'op', 'Operador', lista_operadores, row, 2, 
                                width=70, autocomplete=True)
        
        # ==================== FILA 1: Tracto, Cliente y Ensamble ====================
        row = 1
        
        # Filtro TRACTO (con autocompletado)
        self.crear_filtro_combobox(grid_frame, 'track', 'Tracto', lista_tractos, row, 0, 
                                width=20, autocomplete=True)
        
        # Filtro CLIENTE (con autocompletado)
        self.crear_filtro_combobox(grid_frame, 'cliente', 'Cliente', lista_clientes, row, 2, 
                                width=50, autocomplete=True)
        
        
        # ==================== FILA 2: Rango de fechas CON CALENDARIO ====================
        row = 2
        
        # Filtro ENSAMBLE (sin autocompletado)
        self.crear_filtro_combobox(grid_frame, 'ensamble', 'Ensamble', lista_ensambles, row, 2, 
                                width=18, autocomplete=False)

        # Checkbox para activar filtro de fecha
        var_fecha_activo = tk.BooleanVar(value=False)
        self.filtro_activo['fecha'] = var_fecha_activo
        
        check_fecha = ttk.Checkbutton(grid_frame, text="Rango de Fechas", 
                                    variable=var_fecha_activo,
                                    command=self.toggle_filtro)
        check_fecha.grid(row=row, column=0, sticky=tk.W, padx=(0, 5), pady=5)
        
        # Frame para los dos calendarios
        fecha_frame = ttk.Frame(grid_frame)
        fecha_frame.grid(row=row, column=1, columnspan=5, sticky=tk.W, pady=5)
        
        if TKCALENDAR_AVAILABLE:
            # ===== FECHA INICIAL CON CALENDARIO =====
            ttk.Label(fecha_frame, text="Desde:").pack(side=tk.LEFT, padx=(0, 5))
            
            # DateEntry para fecha inicial
            self.date_inicio = DateEntry(fecha_frame, 
                                        width=15,
                                        background='darkblue',
                                        foreground='white',
                                        borderwidth=2,
                                        date_pattern='yyyy-mm-dd',  # Formato de fecha
                                        state='disabled')
            self.date_inicio.pack(side=tk.LEFT, padx=(0, 15))
            self.filtro_calendarios['fecha_inicio'] = self.date_inicio
            
            # ===== FECHA FINAL CON CALENDARIO =====
            ttk.Label(fecha_frame, text="Hasta:").pack(side=tk.LEFT, padx=(0, 5))
            
            # DateEntry para fecha final
            self.date_final = DateEntry(fecha_frame,
                                    width=15,
                                    background='darkblue',
                                    foreground='white',
                                    borderwidth=2,
                                    date_pattern='yyyy-mm-dd',
                                    state='disabled')
            self.date_final.pack(side=tk.LEFT)
            self.filtro_calendarios['fecha_final'] = self.date_final
            
        else:
            # FALLBACK: Si no está instalado tkcalendar, usar Entry normal
            ttk.Label(fecha_frame, text="Desde:").pack(side=tk.LEFT, padx=(0, 5))
            var_fecha_inicio = tk.StringVar()
            self.filtro_valor['fecha_inicio'] = var_fecha_inicio
            entry_fecha_inicio = ttk.Entry(fecha_frame, textvariable=var_fecha_inicio, 
                                        width=12, state=tk.DISABLED)
            entry_fecha_inicio.pack(side=tk.LEFT, padx=(0, 10))
            entry_fecha_inicio.filtro_key = 'fecha'
            self.filtro_entries['fecha_inicio'] = entry_fecha_inicio
            
            ttk.Label(fecha_frame, text="Hasta:").pack(side=tk.LEFT, padx=(0, 5))
            var_fecha_final = tk.StringVar()
            self.filtro_valor['fecha_final'] = var_fecha_final
            entry_fecha_final = ttk.Entry(fecha_frame, textvariable=var_fecha_final, 
                                        width=12, state=tk.DISABLED)
            entry_fecha_final.pack(side=tk.LEFT)
            entry_fecha_final.filtro_key = 'fecha'
            self.filtro_entries['fecha_final'] = entry_fecha_final
            
            ttk.Label(fecha_frame, text="(YYYY-MM-DD)", foreground='#666').pack(side=tk.LEFT, padx=(5, 0))
        
        # ==================== FILA 3: Filtros de Estado (checkboxes) ====================
        row = 3
        
        # Label para la sección
        ttk.Label(grid_frame, text="Estado:", font=('TkDefaultFont', 9, 'bold')).grid(
            row=row, column=0, sticky=tk.W, padx=(0, 5), pady=5
        )
        
        estado_frame = ttk.Frame(grid_frame)
        estado_frame.grid(row=row, column=1, columnspan=5, sticky=tk.W, pady=5)
        
        # Checkbox IRREAL
        var_irreal = tk.BooleanVar(value=False)
        self.filtro_valor['estado_irreal'] = var_irreal
        check_irreal = ttk.Checkbutton(estado_frame, text="IRREAL", 
                                    variable=var_irreal)
        check_irreal.pack(side=tk.LEFT, padx=(0, 15))
        
        # Checkbox ERR. COM
        var_err_com = tk.BooleanVar(value=False)
        self.filtro_valor['estado_err_com'] = var_err_com
        check_err_com = ttk.Checkbutton(estado_frame, text="ERR. COM", 
                                    variable=var_err_com)
        check_err_com.pack(side=tk.LEFT, padx=(0, 15))
        
        # Checkbox ANALIZAR 1
        var_analizar1 = tk.BooleanVar(value=False)
        self.filtro_valor['estado_analizar1'] = var_analizar1
        check_analizar1 = ttk.Checkbutton(estado_frame, text="ANALIZAR 1", 
                                        variable=var_analizar1)
        check_analizar1.pack(side=tk.LEFT, padx=(0, 15))
        
        # Checkbox ANALIZAR 2
        var_analizar2 = tk.BooleanVar(value=False)
        self.filtro_valor['estado_analizar2'] = var_analizar2
        check_analizar2 = ttk.Checkbutton(estado_frame, text="ANALIZAR 2", 
                                        variable=var_analizar2)
        check_analizar2.pack(side=tk.LEFT, padx=(0, 15))
        
        # Checkbox ANALIZAR 3
        var_analizar3 = tk.BooleanVar(value=False)
        self.filtro_valor['estado_analizar3'] = var_analizar3
        check_analizar3 = ttk.Checkbutton(estado_frame, text="ANALIZAR 3", 
                                        variable=var_analizar3)
        check_analizar3.pack(side=tk.LEFT)
        
        # ==================== BOTONES DE ACCIÓN ====================
        btn_frame = ttk.Frame(filtros_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        btn_buscar = ttk.Button(btn_frame, text="🔍 Buscar", 
                            command=self.aplicar_filtros_consulta)
        btn_buscar.pack(side=tk.LEFT, padx=5)
        
        btn_limpiar = ttk.Button(btn_frame, text="🗑️ Limpiar Filtros", 
                                command=self.limpiar_filtros_consulta)
        btn_limpiar.pack(side=tk.LEFT, padx=5)
        
        btn_todos = ttk.Button(btn_frame, text="📋 Mostrar Todos", 
                            command=self.mostrar_todos_viajes)
        btn_todos.pack(side=tk.LEFT, padx=5)
        
        # Info rápida
        info_label = ttk.Label(btn_frame, 
                            text="💡 Selecciona los filtros que desees usar",
                            foreground='#666')
        info_label.pack(side=tk.RIGHT, padx=10)

    def crear_filtro_combobox(self, parent, key, label, valores, row, col, width=18, 
                         autocomplete=False):  # NUEVO parámetro
        """Helper para crear un filtro con combobox (opcionalmente con autocompletado)"""
        # Checkbox para activar
        var_activo = tk.BooleanVar(value=False)
        self.filtro_activo[key] = var_activo
        
        check = ttk.Checkbutton(parent, text=label, 
                            variable=var_activo,
                            command=self.toggle_filtro)
        check.grid(row=row, column=col, sticky=tk.W, padx=(0, 5), pady=5)
        
        # Combobox para valor (con o sin autocompletado)
        var_valor = tk.StringVar()
        self.filtro_valor[key] = var_valor
        
        if autocomplete:
            # Usar AutocompleteCombobox para filtros con autocompletado
            combo = AutocompleteCombobox(parent, completevalues=valores,
                                        textvariable=var_valor,
                                        state='disabled', width=width)
        else:
            # Usar Combobox normal para los demás
            combo = ttk.Combobox(parent, textvariable=var_valor, values=valores,
                                state='disabled', width=width)
        
        combo.grid(row=row, column=col+1, sticky=tk.W, padx=(0, 15), pady=5)
        combo.filtro_key = key
        
        if not autocomplete:
            combo.set('Todos')  # Solo para combobox normal
        
        self.filtro_combos[key] = combo
    
    def crear_filtro_texto(self, parent, key, label, row, col):
        """Helper para crear un filtro de texto estándar"""
        # Checkbox para activar
        var_activo = tk.BooleanVar(value=False)
        self.filtro_activo[key] = var_activo
        
        check = ttk.Checkbutton(parent, text=label, 
                               variable=var_activo,
                               command=self.toggle_filtro)
        check.grid(row=row, column=col, sticky=tk.W, padx=(0, 5), pady=5)
        
        # Entry para valor
        var_valor = tk.StringVar()
        self.filtro_valor[key] = var_valor
        
        entry = ttk.Entry(parent, textvariable=var_valor, width=20, state=tk.DISABLED)
        entry.grid(row=row, column=col+1, sticky=tk.W, padx=(0, 15), pady=5)
        entry.filtro_key = key
        
        self.filtro_entries[key] = entry
    
    def crear_seccion_resultados(self, parent):
        """Crea la sección del tksheet para mostrar resultados"""
        resultados_frame = ttk.LabelFrame(parent, text="Resultados", padding="5")
        resultados_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Frame de controles de ordenamiento
        orden_frame = ttk.Frame(resultados_frame)
        orden_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(orden_frame, text="Ordenar por:").pack(side=tk.LEFT, padx=(0, 5))
        
        self.columnas_orden = [
            "ID viaje", "Fecha inicio", "Fecha final", "Ruta", "Operador", "Tracto",
            "Ensamble", "Frontal", "Dolly", "Trasero", "# Casetas", "Importe",
            "# Casetas fuera de Viaje", "Importe fuera de viaje", "Total", "Deposito",
            "Comisión", "Gastos", "Estado", "Origen"
        ]
        self.combo_orden = ttk.Combobox(orden_frame, values=self.columnas_orden, 
                                        state='readonly', width=15)
        self.combo_orden.pack(side=tk.LEFT, padx=(0, 10))
        
        btn_asc = ttk.Button(orden_frame, text="▲ Asc", width=8,
                            command=lambda: self.ordenar_desde_combo(True))
        btn_asc.pack(side=tk.LEFT, padx=2)
        
        btn_desc = ttk.Button(orden_frame, text="▼ Desc", width=8,
                             command=lambda: self.ordenar_desde_combo(False))
        btn_desc.pack(side=tk.LEFT, padx=2)
        
        # Crear Sheet
        self.sheet = Sheet(
            resultados_frame,
            headers=[],
            data=[[]],
            width=1350,
            height=500,
            show_horizontal_scrollbar=True,
            show_vertical_scrollbar=True,
            enable_edit=False
        )
        self.sheet.pack(fill=tk.BOTH, expand=True)
        
        self.sheet.enable_bindings((
            "single_select",
            "row_select",
            "column_width_resize",
            "arrowkeys",
            "right_click_popup_menu",
            "copy",
            "column_select",
        ))
        
        self.sheet.extra_bindings([
            ("cell_select", self.on_celda_seleccionada),
        ])
        
        # NUEVO: Personalizar menú contextual
        self.sheet.popup_menu_add_command(
            "📋 Ver Detalles del Viaje",
            self.abrir_detalles_desde_menu
        )
        
        # Variables para ordenamiento
        self.columna_ordenada = None
        self.orden_ascendente = True
    
    def abrir_detalles_desde_menu(self):
        """Abre detalles del viaje seleccionado desde el menú contextual"""
        selected = self.sheet.get_currently_selected()
        if selected and selected.row is not None:
            row = selected.row
            if row < len(self.datos_actuales):
                viaje = self.datos_actuales[row]
                self.abrir_ventana_detalles(viaje)

    def abrir_ventana_detalles(self, viaje):
        """
             Abre una ventana con los detalles completos de un viaje.
        Muestra: datos del viaje, datos de ruta (si existe), casetas dentro y fuera, gastos.
        """
        id_viaje = viaje.get('id_viaje', 'N/A')
        
        # Crear ventana Toplevel
        ventana = tk.Toplevel(self.ventana_consulta)
        ventana.title(f"Detalles del Viaje - {id_viaje}")
        ventana.geometry("1200x900")
        
        # Registrar ventana
        self.ventanas_detalles.append(ventana)
        
        # Al cerrar, eliminar del registro
        def al_cerrar():
            if ventana in self.ventanas_detalles:
                self.ventanas_detalles.remove(ventana)
            ventana.destroy()
        
        ventana.protocol("WM_DELETE_WINDOW", al_cerrar)
        
        # Frame principal con scroll
        main_frame = ttk.Frame(ventana, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas para scroll
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ==================== TÍTULO ====================
        titulo = ttk.Label(scrollable_frame, 
                        text=f"DETALLES DEL VIAJE: {id_viaje}",
                        font=('Arial', 16, 'bold'))
        titulo.pack(pady=(0, 20))
        
        # ==================== SECCIÓN 1: DATOS DEL VIAJE ====================
        self.crear_seccion_datos_viaje(scrollable_frame, viaje)
        
        # ==================== SECCIÓN 2: DATOS DE RUTA (si existe) ====================
        id_ruta = viaje.get('id_ruta')
        if id_ruta:
            rutas = [r for r in self.db_manager.obtener_todas_rutas() 
                    if r.get('id_ruta') == id_ruta]
            if rutas:
                self.crear_seccion_datos_ruta(scrollable_frame, rutas)
        
        # ==================== SECCIÓN 3: CASETAS DENTRO DEL VIAJE ====================
        casetas_dentro = self.db_manager.obtener_casetas_por_viaje(id_viaje, solo_fuera=False)
        self.crear_seccion_casetas(scrollable_frame, casetas_dentro, "CASETAS DENTRO DEL VIAJE")
        
        # ==================== SECCIÓN 4: CASETAS FUERA DEL VIAJE ====================
        casetas_fuera = self.db_manager.obtener_casetas_por_viaje(id_viaje, solo_fuera=True)
        self.crear_seccion_casetas(scrollable_frame, casetas_fuera, "CASETAS FUERA DEL VIAJE")
        
        # ==================== SECCIÓN 5: GASTOS (NUEVO) ====================
        self.crear_seccion_gastos_adicionales(scrollable_frame, viaje)
        
        # ==================== BOTÓN CERRAR ====================
        btn_cerrar = ttk.Button(scrollable_frame, text="❌ Cerrar", 
                            command=al_cerrar)
        btn_cerrar.pack(pady=20)
    
    def crear_seccion_datos_viaje(self, parent, viaje):
        """Crea sección de datos del viaje en tksheet"""
        frame = ttk.LabelFrame(parent, text="Datos del Viaje", padding="10")
        frame.pack(fill=tk.BOTH, pady=(0, 15))
        
        # Obtener valores
        gastos_adicionales_total = viaje.get('gastos_adicionales_total', 0.0)
        comision_original = viaje.get('comision', 0.0)
        comision_ajustada = viaje.get('comision_ajustada')
        
        # Preparar datos para sheet (formato: Campo | Valor)
        datos = [
            ["ID Viaje", viaje.get('id_viaje', '')],
            ["Fecha Inicio", format_date_safe(viaje.get('fecha_inicio'))],
            ["Fecha Final", format_date_safe(viaje.get('fecha_final'))],
            ["Ruta", viaje.get('ruta', '')],
            ["Operador", viaje.get('operador', '')],
            ["Tracto", viaje.get('tracto', '')],
            ["Ensamble", viaje.get('ensamble', '')],
            ["Frontal", viaje.get('frontal', '')],
            ["Dolly", viaje.get('dolly', '')],
            ["Trasero", viaje.get('trasero', '')],
            ["Clave", viaje.get('clave', '')],
            ["ID Ruta", viaje.get('id_ruta', '')],
            ["", ""],  # Separador
            ["# Casetas Dentro", str(viaje.get('num_casetas', 0))],
            ["Importe Casetas Dentro", f"${viaje.get('importe_total', 0):,.2f}"],
            ["# Casetas Fuera", str(viaje.get('num_casetas_fuera', 0))],
            ["Importe Casetas Fuera", f"${viaje.get('importe_fuera', 0):,.2f}"],
            ["Importe Total", f"${(viaje.get('importe_total', 0) + viaje.get('importe_fuera', 0)):,.2f}"],
            ["", ""],  # Separador
            ["Deposito", f"${viaje.get('gasto', 0):,.2f}" if viaje.get('gasto') else ''],  # MODIFICADO: "Gastos" → "Deposito"
            ["Gastos", f"${gastos_adicionales_total:,.2f}" if gastos_adicionales_total else ''],  # AÑADIDO
            ["Comisión Original", f"${comision_original:,.2f}" if comision_original else ''],  # MODIFICADO
            ["Comisión Ajustada", f"${comision_ajustada:,.2f}" if comision_ajustada else 'N/A'],  # AÑADIDO
            ["", ""],  # Separador
            ["Estado", ', '.join(viaje.get('advertencias', [])) or 'OK'],
            ["Origen Datos", viaje.get('origen', '')],
        ]
        
        sheet = Sheet(frame, headers=["Campo", "Valor"], data=datos,
                    width=1100, height=450, enable_edit=False)  # MODIFICADO: height 400 → 450 para más filas
        sheet.pack(fill=tk.BOTH, expand=True)
        
        sheet.enable_bindings(("single_select", "copy"))
        sheet.column_width(0, 200)
        sheet.column_width(1, 800)

    def crear_seccion_gastos_adicionales(self, parent, viaje):
        """Crea sección de gastos adicionales en tksheet"""
        # Obtener datos de gastos adicionales del viaje
        clasificaciones = viaje.get('gastos_adicionales_clasificacion', [])
        motivos = viaje.get('gastos_adicionales_motivo', [])
        importes = viaje.get('gastos_adicionales_importe', [])
        total = viaje.get('gastos_adicionales_total', 0.0)
        
        frame = ttk.LabelFrame(parent, text=f"GASTOS ({len(clasificaciones)})", padding="10")
        frame.pack(fill=tk.BOTH, pady=(0, 15))
        
        if not clasificaciones or len(clasificaciones) == 0:
            ttk.Label(frame, text="No hay gastos", 
                    foreground='#666').pack(pady=20)
            return
        
        # Preparar datos - combinar las tres listas
        headers = ["Clasificación", "Motivo", "Importe"]
        datos = []
        
        for i in range(len(clasificaciones)):
            clasificacion = clasificaciones[i] if i < len(clasificaciones) else ''
            motivo = motivos[i] if i < len(motivos) else ''
            importe = importes[i] if i < len(importes) else 0.0
            
            importe_str = f"${importe:,.2f}"
            
            datos.append([clasificacion, motivo, importe_str])
        
        # Calcular altura según cantidad de gastos
        altura = min(300, max(100, len(clasificaciones) * 25 + 50))
        
        sheet = Sheet(frame, headers=headers, data=datos,
                    width=1100, height=altura, enable_edit=False)
        sheet.pack(fill=tk.BOTH, expand=True)
        
        sheet.enable_bindings(("single_select", "copy", "row_select"))
        
        # Ajustar anchos
        sheet.column_width(0, 200)  # Clasificación
        sheet.column_width(1, 700)  # Motivo
        sheet.column_width(2, 150)  # Importe
        
        # Mostrar total
        ttk.Label(frame, text=f"Total: ${total:,.2f}", 
                font=('Arial', 10, 'bold')).pack(pady=(5, 0))
    
    def crear_seccion_datos_ruta(self, parent, rutas):
        """Crea sección de datos de ruta en tksheet"""
        frame = ttk.LabelFrame(parent, text="Datos de Ruta", padding="10")
        frame.pack(fill=tk.BOTH, pady=(0, 15))
        
        # Datos generales (primera ruta)
        primera = rutas[0]
        
        datos_generales = [
            ["ID Ruta", primera.get('id_ruta', '')],
            ["Clave", primera.get('clave', '')],
            ["Nombre Ruta", primera.get('nombre_ruta', '')],
            ["Cliente", primera.get('cliente', '')],
            ["Origen", primera.get('origen', '')],
            ["Entrega", primera.get('entrega', '')],
            ["Base", primera.get('base', '')],
            ["Ensamble", primera.get('ensamble', '')],
            ["Kilómetros", str(primera.get('kilometros', ''))],
        ]
        
        sheet_general = Sheet(frame, headers=["Campo", "Valor"], data=datos_generales,
                             width=1100, height=200, enable_edit=False)
        sheet_general.pack(fill=tk.BOTH, pady=(0, 10))
        sheet_general.enable_bindings(("single_select", "copy"))
        sheet_general.column_width(0, 200)
        sheet_general.column_width(1, 800)
        
        # Tipos de viaje (tabla)
        if len(rutas) > 0:
            ttk.Label(frame, text="Tipos de Viaje:", font=('Arial', 10, 'bold')).pack(pady=(5, 5))
            
            headers_tipos = ["Tipo Viaje", "Flete", "Tipo Cobro", "Pago Operador", "Tipo Cobro Op"]
            datos_tipos = []
            
            for ruta in rutas:
                tipo = ruta.get('tipo_viaje', 'Sin tipo')
                flete = f"${ruta.get('flete', 0):,.2f}" if ruta.get('flete') else ''
                tipo_cobro = ruta.get('tipo_cobro', '')
                pago_op = f"${ruta.get('pago_operador', 0):,.2f}" if ruta.get('pago_operador') else ''
                tipo_cobro_op = ruta.get('tipo_cobro_op', '')
                
                datos_tipos.append([tipo, flete, tipo_cobro, pago_op, tipo_cobro_op])
            
            sheet_tipos = Sheet(frame, headers=headers_tipos, data=datos_tipos,
                               width=1100, height=150, enable_edit=False)
            sheet_tipos.pack(fill=tk.BOTH)
            sheet_tipos.enable_bindings(("single_select", "copy"))
            
            # Ajustar anchos
            sheet_tipos.column_width(0, 150)
            sheet_tipos.column_width(1, 120)
            sheet_tipos.column_width(2, 120)
            sheet_tipos.column_width(3, 150)
            sheet_tipos.column_width(4, 150)
    
    def crear_seccion_casetas(self, parent, casetas, titulo):
        """Crea sección de casetas en tksheet"""
        frame = ttk.LabelFrame(parent, text=f"{titulo} ({len(casetas)})", padding="10")
        frame.pack(fill=tk.BOTH, pady=(0, 15))
        
        if not casetas:
            ttk.Label(frame, text="No hay casetas", 
                     foreground='#666').pack(pady=20)
            return
        
        # Preparar datos
        headers = ["Tracto", "Fecha", "Hora", "Importe", "Clase", "Nombre", "ID Caseta"]
        datos = []
        
        for caseta in casetas:
            tracto = caseta.get('tracto', '')
            fecha = caseta.get('fecha', '')
            hora = caseta.get('hora', '')
            importe = f"${caseta.get('importe', 0):,.2f}"
            clase = caseta.get('clase', '')
            nombre = caseta.get('nombre', '')
            id_caseta = caseta.get('id_caseta', '')
            
            datos.append([tracto, fecha, hora, importe, clase, nombre, id_caseta])
        
        # Calcular altura según cantidad de casetas
        altura = min(300, max(100, len(casetas) * 25 + 50))
        
        sheet = Sheet(frame, headers=headers, data=datos,
                     width=1100, height=altura, enable_edit=False)
        sheet.pack(fill=tk.BOTH, expand=True)
        
        sheet.enable_bindings(("single_select", "copy", "row_select"))
        
        # Ajustar anchos
        sheet.column_width(0, 80)   # Tracto
        sheet.column_width(1, 120)  # Fecha
        sheet.column_width(2, 100)  # Hora
        sheet.column_width(3, 100)  # Importe
        sheet.column_width(4, 80)   # Clase
        sheet.column_width(5, 300)  # Nombre
        sheet.column_width(6, 200)  # ID Caseta
        
        # Mostrar total
        total = sum(c.get('importe', 0) for c in casetas)
        ttk.Label(frame, text=f"Total: ${total:,.2f}", 
                 font=('Arial', 10, 'bold')).pack(pady=(5, 0))
    
    def ordenar_desde_combo(self, ascendente):
        """Ordena usando el combobox"""
        columna_nombre = self.combo_orden.get()
        if not columna_nombre:
            messagebox.showwarning("Ordenar", "Selecciona una columna primero")
            return
        
        # Mapeo de nombre a índice
        try:
            col_idx = self.columnas_orden.index(columna_nombre)
            # Mapear a índice real de columna en el sheet
            mapeo_indices = {
                0: 0,   # ID viaje
                1: 1,   # Fecha inicio
                2: 2,   # Fecha final
                3: 3,   # Ruta
                4: 4,   # Operador
                5: 5,   # Tracto
                6: 6,   # Ensamble
                7: 10,  # #Casetas
                8: 11,  # Importe
                9: 14,  # Total
                10: 15, # Gastos
                11: 16, # Comisión
                12: 17, # Estado
            }
            
            col_real = mapeo_indices.get(col_idx, col_idx)
            self.orden_ascendente = ascendente
            self.ordenar_por_columna(col_real)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error ordenando: {e}")
    
    def on_header_clic(self, event):
        """Callback cuando se hace clic en un header"""
        if event.column is not None:
            self.ordenar_por_columna(event.column)
    
    def on_header_doble_clic(self, event):
        """Callback cuando se hace doble clic en un header"""
        # Por si quieres agregar funcionalidad extra
        pass

    def crear_barra_estado(self, parent):
        """Crea barra de estado en la parte inferior"""
        estado_frame = ttk.Frame(parent)
        estado_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.label_estado = ttk.Label(estado_frame, text="Listo", relief=tk.SUNKEN)
        self.label_estado.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.label_stats = ttk.Label(estado_frame, text="Registros: 0", relief=tk.SUNKEN)
        self.label_stats.pack(side=tk.RIGHT, padx=(5, 0))
    
    def crear_botones_principales(self, parent):
        """Crea botones principales de acción"""
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X)
        
        btn_actualizar = ttk.Button(btn_frame, text="📥 Actualizar Datos", 
                                    command=lambda: self.abrir_modo_actualizacion(self.ventana_consulta))
        btn_actualizar.pack(side=tk.LEFT, padx=5)
        
        btn_nuevo_viaje = ttk.Button(btn_frame, text="➕ Añadir Viaje Manual", 
                                    command=self.abrir_formulario_viaje,
                                    style='Accent.TButton')
        btn_nuevo_viaje.pack(side=tk.LEFT, padx=5)
        
        btn_exportar = ttk.Button(btn_frame, text="📊 Exportar a Excel", 
                                command=self.exportar_resultados)
        btn_exportar.pack(side=tk.LEFT, padx=5)
        
        btn_stats = ttk.Button(btn_frame, text="📈 Estadísticas", 
                            command=self.mostrar_estadisticas)
        btn_stats.pack(side=tk.LEFT, padx=5)
        
        btn_rutas = ttk.Button(btn_frame, text="🗺️ Vista de Rutas", 
                            command=self.abrir_vista_rutas)
        btn_rutas.pack(side=tk.LEFT, padx=5)
        
        btn_cerrar = ttk.Button(btn_frame, text="❌ Cerrar", 
                            command=self.cerrar_modulo_consulta_completo)
        btn_cerrar.pack(side=tk.RIGHT, padx=5)
    
    def toggle_filtro(self):
        """Habilita/deshabilita controles según checkboxes"""
        # Filtros con combobox
        for key in ['ruta', 'op', 'track', 'cliente', 'ensamble']:
            if key in self.filtro_activo and key in self.filtro_combos:
                combo = self.filtro_combos[key]
                if self.filtro_activo[key].get():
                    combo.config(state='normal')
                else:
                    combo.config(state='disabled')
                    if isinstance(combo, AutocompleteCombobox):
                        combo.set('')
                        combo['values'] = combo._completevalues
                    else:
                        combo.set('Todos')
        
        # Filtro de fecha (con calendario o entry)
        if 'fecha' in self.filtro_activo:
            activo = self.filtro_activo['fecha'].get()
            
            if TKCALENDAR_AVAILABLE and hasattr(self, 'filtro_calendarios'):
                # Usar DateEntry
                estado = 'normal' if activo else 'disabled'
                
                if 'fecha_inicio' in self.filtro_calendarios:
                    self.filtro_calendarios['fecha_inicio'].config(state=estado)
                if 'fecha_final' in self.filtro_calendarios:
                    self.filtro_calendarios['fecha_final'].config(state=estado)
                
                if not activo:
                    # Resetear a fecha actual
                    if hasattr(self, 'date_inicio'):
                        self.date_inicio.set_date(datetime.now())
                    if hasattr(self, 'date_final'):
                        self.date_final.set_date(datetime.now())
            else:
                # Fallback a Entry normal
                estado = tk.NORMAL if activo else tk.DISABLED
                
                if 'fecha_inicio' in self.filtro_entries:
                    self.filtro_entries['fecha_inicio'].config(state=estado)
                if 'fecha_final' in self.filtro_entries:
                    self.filtro_entries['fecha_final'].config(state=estado)
                
                if not activo:
                    if 'fecha_inicio' in self.filtro_valor:
                        self.filtro_valor['fecha_inicio'].set('')
                    if 'fecha_final' in self.filtro_valor:
                        self.filtro_valor['fecha_final'].set('')
    
    def aplicar_filtros_consulta(self):
        """Aplica los filtros activos y muestra resultados"""
        self.actualizar_estado("Aplicando filtros...")
        
        from core_viajes import enriquecer_viajes_con_casetas
        
        # CORREGIDO: Obtener todos los viajes CON sus IDs incluidos
        viajes_filtrados = []
        for id_viaje, datos_viaje in self.db_manager.data['viajes'].items():
            viaje_con_id = {'id_viaje': id_viaje, **datos_viaje}
            viajes_filtrados.append(viaje_con_id)
        
        # NUEVA VALIDACIÓN: Verificar que los valores de autocompletado sean válidos
        filtros_autocomplete = ['ruta', 'op', 'track', 'cliente']
        
        for key in filtros_autocomplete:
            if self.filtro_activo.get(key, tk.BooleanVar()).get():
                valor = self.filtro_valor[key].get().strip()
                if valor:
                    combo = self.filtro_combos[key]
                    if isinstance(combo, AutocompleteCombobox):
                        # Verificar que el valor esté en la lista de opciones
                        if valor not in combo._completevalues:
                            messagebox.showwarning("Filtro inválido", 
                                                f"El valor '{valor}' no es válido para {key.upper()}.\n"
                                                "Selecciona un valor de la lista desplegable.")
                            self.actualizar_estado("Filtro inválido")
                            return
        
        # Filtro RUTA
        if self.filtro_activo.get('ruta', tk.BooleanVar()).get():
            ruta_buscar = self.filtro_valor['ruta'].get().strip()
            if ruta_buscar:
                viajes_filtrados = [v for v in viajes_filtrados 
                                if str(v.get('ruta', '')) == ruta_buscar]
        
        # Filtro OPERADOR
        if self.filtro_activo.get('op', tk.BooleanVar()).get():
            op_buscar = self.filtro_valor['op'].get().strip()
            if op_buscar:
                viajes_filtrados = [v for v in viajes_filtrados 
                                if str(v.get('operador', '')) == op_buscar]
        
        # Filtro TRACTO
        if self.filtro_activo.get('track', tk.BooleanVar()).get():
            track_buscar = self.filtro_valor['track'].get().strip()
            if track_buscar:
                viajes_filtrados = [v for v in viajes_filtrados 
                                if str(v.get('tracto', '')) == track_buscar]
        
        # Filtro CLIENTE
        if self.filtro_activo.get('cliente', tk.BooleanVar()).get():
            cliente_buscar = self.filtro_valor['cliente'].get().strip()
            if cliente_buscar:
                viajes_filtrados_temp = []
                rutas_db = self.db_manager.obtener_todas_rutas()
                
                ruta_cliente_map = {}
                for ruta in rutas_db:
                    ruta_cliente_map[ruta.get('id_ruta')] = ruta.get('cliente', '')
                
                for v in viajes_filtrados:
                    id_ruta = v.get('id_ruta') or v.get('ruta')
                    if id_ruta and ruta_cliente_map.get(id_ruta, '') == cliente_buscar:
                        viajes_filtrados_temp.append(v)
                
                viajes_filtrados = viajes_filtrados_temp
        
        # Filtro ENSAMBLE
        if self.filtro_activo.get('ensamble', tk.BooleanVar()).get():
            ensamble_buscar = self.filtro_valor['ensamble'].get().strip()
            if ensamble_buscar and ensamble_buscar != 'Todos':
                viajes_filtrados = [v for v in viajes_filtrados 
                                if str(v.get('ensamble', '')).upper() == ensamble_buscar]
        
        # Filtro FECHA
        if self.filtro_activo.get('fecha', tk.BooleanVar()).get():
            if TKCALENDAR_AVAILABLE and hasattr(self, 'date_inicio') and hasattr(self, 'date_final'):
                # Obtener fechas desde DateEntry
                try:
                    fecha_inicio = pd.to_datetime(self.date_inicio.get_date())
                    fecha_final = pd.to_datetime(self.date_final.get_date())
                    
                    if fecha_final < fecha_inicio:
                        messagebox.showerror("Error en fechas", 
                                        "La fecha final no puede ser anterior a la fecha de inicio")
                        self.actualizar_estado("Error en fechas")
                        return
                    
                    # Filtrar viajes cuya fecha_inicio esté en el rango
                    viajes_filtrados = [v for v in viajes_filtrados 
                                    if fecha_inicio <= pd.to_datetime(v.get('fecha_inicio', pd.Timestamp.min)) <= fecha_final]
                except Exception as e:
                    messagebox.showerror("Error en fechas", f"Error procesando fechas:\n{e}")
                    self.actualizar_estado("Error en fechas")
                    return
            else:
                # Fallback a Entry normal
                fecha_inicio_str = self.filtro_valor.get('fecha_inicio', tk.StringVar()).get().strip()
                fecha_final_str = self.filtro_valor.get('fecha_final', tk.StringVar()).get().strip()
                
                if fecha_inicio_str and fecha_final_str:
                    try:
                        fecha_inicio = pd.to_datetime(fecha_inicio_str)
                        fecha_final = pd.to_datetime(fecha_final_str)
                        
                        if fecha_final < fecha_inicio:
                            messagebox.showerror("Error en fechas", 
                                            "La fecha final no puede ser anterior a la fecha de inicio")
                            self.actualizar_estado("Error en fechas")
                            return
                        
                        viajes_filtrados = [v for v in viajes_filtrados 
                                        if fecha_inicio <= pd.to_datetime(v.get('fecha_inicio', pd.Timestamp.min)) <= fecha_final]
                    except Exception as e:
                        messagebox.showerror("Error en fechas", 
                                        f"Formato de fecha inválido.\nUsa: YYYY-MM-DD\n\nError: {e}")
                        self.actualizar_estado("Error en formato de fechas")
                        return
        
        # Filtro ESTADO (checkboxes)
        estados_buscados = []
        if self.filtro_valor.get('estado_irreal', tk.BooleanVar()).get():
            estados_buscados.append('IRREAL')
        if self.filtro_valor.get('estado_err_com', tk.BooleanVar()).get():
            estados_buscados.append('ERR. COM')
        if self.filtro_valor.get('estado_analizar1', tk.BooleanVar()).get():
            estados_buscados.append('ANALIZAR 1')
        if self.filtro_valor.get('estado_analizar2', tk.BooleanVar()).get():
            estados_buscados.append('ANALIZAR 2')
        if self.filtro_valor.get('estado_analizar3', tk.BooleanVar()).get():
            estados_buscados.append('ANALIZAR 3')
        
        if estados_buscados:
            viajes_filtrados = [v for v in viajes_filtrados 
                            if any(estado in v.get('advertencias', []) for estado in estados_buscados)]
        
        # Enriquecer con casetas
        self.actualizar_estado("Calculando estadísticas...")
        viajes_enriquecidos = enriquecer_viajes_con_casetas(self.db_manager, viajes_filtrados)
        
        # Mostrar en sheet
        self.mostrar_viajes_en_sheet(viajes_enriquecidos)
        
        self.actualizar_estado(f"Filtros aplicados - {len(viajes_enriquecidos)} viajes encontrados")
    
    def limpiar_filtros_consulta(self):
        """Limpia todos los filtros"""
        # Limpiar filtros con combobox
        for key in ['ruta', 'op', 'track', 'cliente', 'ensamble']:
            if key in self.filtro_activo:
                self.filtro_activo[key].set(False)
                if key in self.filtro_combos:
                    combo = self.filtro_combos[key]
                    combo.config(state='disabled')
                    
                    if isinstance(combo, AutocompleteCombobox):
                        combo.set('')
                        combo['values'] = combo._completevalues
                    else:
                        combo.set('Todos')
        
        # Limpiar filtro de fecha
        if 'fecha' in self.filtro_activo:
            self.filtro_activo['fecha'].set(False)
            
            if TKCALENDAR_AVAILABLE and hasattr(self, 'date_inicio') and hasattr(self, 'date_final'):
                # Resetear DateEntry
                self.date_inicio.config(state='disabled')
                self.date_final.config(state='disabled')
                self.date_inicio.set_date(datetime.now())
                self.date_final.set_date(datetime.now())
            else:
                # Fallback a Entry
                if 'fecha_inicio' in self.filtro_valor:
                    self.filtro_valor['fecha_inicio'].set('')
                if 'fecha_final' in self.filtro_valor:
                    self.filtro_valor['fecha_final'].set('')
                if 'fecha_inicio' in self.filtro_entries:
                    self.filtro_entries['fecha_inicio'].config(state=tk.DISABLED)
                if 'fecha_final' in self.filtro_entries:
                    self.filtro_entries['fecha_final'].config(state=tk.DISABLED)
        
        # Limpiar filtros de estado
        if 'estado_irreal' in self.filtro_valor:
            self.filtro_valor['estado_irreal'].set(False)
        if 'estado_err_com' in self.filtro_valor:
            self.filtro_valor['estado_err_com'].set(False)
        if 'estado_analizar1' in self.filtro_valor:
            self.filtro_valor['estado_analizar1'].set(False)
        if 'estado_analizar2' in self.filtro_valor:
            self.filtro_valor['estado_analizar2'].set(False)
        if 'estado_analizar3' in self.filtro_valor:
            self.filtro_valor['estado_analizar3'].set(False)
        
        self.actualizar_estado("Filtros limpiados")
    
    def mostrar_todos_viajes(self):
        """Muestra todos los viajes sin filtros"""
        self.limpiar_filtros_consulta()
        self.aplicar_filtros_consulta()
    
    def mostrar_viajes_en_sheet(self, viajes):
        """Muestra viajes en el tksheet"""
        if not viajes:
            self.sheet.set_sheet_data([[]])
            self.sheet.headers([])
            self.actualizar_stats(0)
            return
        
        # Definir encabezados
        headers = [
            "ID viaje", "Fecha inicio", "Fecha final", "Ruta", "Operador", "Tracto",
            "Ensamble", "Frontal", "Dolly", "Trasero", "# Casetas", "Importe", 
            "# Casetas fuera de Viaje", "Importe fuera de viaje",
            "Total", "Deposito", "Comisión", "Gastos", "Estado", "Origen"
        ]
        
        # Preparar datos
        filas = []
        for e in viajes:
            id_v = str(e.get('id_viaje', ''))
            fi = format_date_safe(e.get('fecha_inicio'))
            ff = format_date_safe(e.get('fecha_final'))
            ruta = str(e.get('ruta') or '')
            operador = str(e.get('operador') or '')
            tracto = str(e.get('tracto') or '')
            
            ensamble = str(e.get('ensamble') or e.get('colL') or '')
            frontal = str(e.get('frontal') or '')
            dolly = str(e.get('dolly') or '')
            trasero = str(e.get('trasero') or '')
            
            ncas = e.get('num_casetas', 0)
            imp = e.get('importe_total', 0.0)
            ncas_fuera = e.get('num_casetas_fuera', 0)
            imp_fuera = e.get('importe_fuera', 0.0)
            imp_total = imp_fuera + imp
            
            gasto = e.get('gasto', 0.0)  # Ahora "Deposito"
            
            # MODIFICADO: Usar comision_ajustada si existe, sino comision
            comision = e.get('comision_ajustada') or e.get('comision', 0.0)
            
            # AÑADIDO: Gastos adicionales
            gastos_adicionales_total = e.get('gastos_adicionales_total', 0.0)
            
            origen = str(e.get('origen') or '')
            
            advertencias = e.get('advertencias', [])
            status = ', '.join(advertencias) if advertencias else 'OK'
            
            # Formatear valores
            ncas_str = f"{ncas:d}" if ncas != 0 else ''
            imp_str = f"${imp:,.2f}" if imp != 0 else ''
            ncas_fuera_str = f"{ncas_fuera:d}" if ncas_fuera != 0 else ''
            imp_fuera_str = f"${imp_fuera:,.2f}" if imp_fuera != 0 else ''
            imp_total_str = f"${imp_total:,.2f}" if imp_total != 0 else ''
            gasto_str = f"${gasto:,.2f}" if gasto != 0 else ''
            comision_str = f"${comision:,.2f}" if comision != 0 else ''
            gastos_adicionales_str = f"${gastos_adicionales_total:,.2f}" if gastos_adicionales_total != 0 else ''
            
            fila = [
                id_v, fi, ff, ruta, operador, tracto,
                ensamble, frontal, dolly, trasero, ncas_str, imp_str,
                ncas_fuera_str, imp_fuera_str, imp_total_str, gasto_str,
                comision_str, gastos_adicionales_str, status, origen
            ]
            filas.append(fila)
        
        # Configurar sheet
        self.sheet.headers(headers)
        self.sheet.set_sheet_data(filas)
        
        # Ajustar anchos de columna
        col_widths = [140, 160, 160, 100, 100, 80, 60, 90, 90, 90, 
                    50, 110, 60, 110, 110, 110, 110, 110, 120, 100]
        for i, w in enumerate(col_widths):
            try:
                self.sheet.column_width(i, w)
            except:
                pass
        
        # Guardar datos originales para ordenamiento
        self.datos_actuales = viajes
        
        self.actualizar_stats(len(viajes))
        self.sheet.redraw()
    
    def ordenar_por_columna(self, col_idx):
        """Ordena los datos por la columna seleccionada"""
        if not hasattr(self, 'datos_actuales') or not self.datos_actuales:
            return

        # Mapeo de columnas a campos del viaje
        campos_mapeo = [
            'id_viaje', 'fecha_inicio', 'fecha_final', 'ruta', 'operador', 'tracto',
            'ensamble', 'frontal', 'dolly', 'trasero', 'num_casetas', 'importe_total',
            'num_casetas_fuera', 'importe_fuera', 'total_calculado', 'gasto',
            'comision_ajustada', 'gastos_adicionales_total', 'estado_texto', 'origen'
        ]

        if col_idx >= len(campos_mapeo):
            return

        campo = campos_mapeo[col_idx]

        # Ordenar datos
        try:
            if campo == 'comision_ajustada':
                self.datos_actuales.sort(
                    key=lambda x: x.get('comision_ajustada') or x.get('comision', 0),
                    reverse=not self.orden_ascendente
                )
            elif campo == 'total_calculado':
                self.datos_actuales.sort(
                    key=lambda x: (x.get('importe_total', 0) or 0) + (x.get('importe_fuera', 0) or 0),
                    reverse=not self.orden_ascendente
                )
            elif campo == 'estado_texto':
                self.datos_actuales.sort(
                    key=lambda x: ', '.join(x.get('advertencias', [])) or 'OK',
                    reverse=not self.orden_ascendente
                )
            else:
                self.datos_actuales.sort(
                    key=lambda x: x.get(campo, ''),
                    reverse=not self.orden_ascendente
                )

            # Actualizar vista
            self.mostrar_viajes_en_sheet(self.datos_actuales)

            # Indicador visual de ordenamiento
            simbolo = "▲" if self.orden_ascendente else "▼"
            self.actualizar_estado(f"Ordenado por {campo} {simbolo}")

        except Exception as e:
            print(f"Error ordenando: {e}")
    
    def on_celda_seleccionada(self, event):
        """Maneja selección de celda"""
        # Obtener fila seleccionada
        selected = self.sheet.get_currently_selected()
        if selected:
            row = selected.row
            if row is not None and row < len(self.datos_actuales):
                viaje = self.datos_actuales[row]
                self.actualizar_estado(f"Seleccionado: {viaje.get('id_viaje', 'N/A')}")
    
    
    def exportar_resultados(self):
        """Exporta resultados actuales a Excel"""
        if not hasattr(self, 'datos_actuales') or not self.datos_actuales:
            messagebox.showwarning("Exportar", "No hay datos para exportar")
            return
        
        # Diálogo para guardar archivo
        from tkinter import filedialog
        archivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"viajes_export_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if not archivo:
            return
        
        try:
            # Crear DataFrame desde los datos del sheet
            headers = self.sheet.headers()
            data = self.sheet.get_sheet_data()
            
            df = pd.DataFrame(data, columns=headers)
            df.to_excel(archivo, index=False)
            
            messagebox.showinfo("Exportar", f"Datos exportados exitosamente a:\n{archivo}")
            self.actualizar_estado(f"Exportado: {len(data)} registros")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error exportando:\n{e}")
    
    def mostrar_estadisticas(self):
        """Muestra estadísticas de la base de datos"""
        stats = self.db_manager.obtener_estadisticas_actualizacion()
        
        mensaje = f"""
ESTADÍSTICAS DE LA BASE DE DATOS

Viajes totales: {stats['total_viajes']:,}
Casetas totales: {stats['total_casetas']:,}
Relaciones: {stats['total_relaciones']:,}
Rutas: {stats['total_rutas']:,}

VIAJES:
  - Completos (Mon+Cierre+Gastos): {stats['viajes_completos']:,}
  - Con ruta asignada: {stats['viajes_con_ruta']:,}
  - Con advertencias: {stats['viajes_con_advertencias']:,}
    · Irreales: {stats['viajes_irreal']:,}
    · Error comisión: {stats['viajes_err_com']:,}

RESULTADOS ACTUALES:
  - Registros mostrados: {len(self.datos_actuales) if hasattr(self, 'datos_actuales') else 0:,}
        """
        
        messagebox.showinfo("Estadísticas", mensaje)
    
    def abrir_formulario_viaje(self):
        """Abre el formulario para añadir viajes manualmente"""
        FormularioViaje(self.ventana_consulta, self.db_manager)

    def abrir_vista_rutas(self):
        """
        Abre una ventana nueva con la vista de rutas.
        Similar a la vista de viajes pero especializada para rutas.
        """
        # Crear ventana Toplevel
        ventana_rutas = tk.Toplevel(self.ventana_consulta)
        ventana_rutas.title("Vista de Rutas")
        ventana_rutas.geometry("1400x800")
        
        # Registrar ventana
        self.ventanas_detalles.append(ventana_rutas)
        
        # Al cerrar, eliminar del registro
        def al_cerrar():
            if ventana_rutas in self.ventanas_detalles:
                self.ventanas_detalles.remove(ventana_rutas)
            ventana_rutas.destroy()
        
        ventana_rutas.protocol("WM_DELETE_WINDOW", al_cerrar)
        
        # Frame principal
        main_frame = ttk.Frame(ventana_rutas, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ==================== TÍTULO ====================
        titulo = ttk.Label(main_frame, 
                        text="VISTA DE RUTAS",
                        font=('Arial', 16, 'bold'))
        titulo.pack(pady=(0, 10))
        
        # ==================== SECCIÓN DE FILTROS ====================
        filtros_frame = ttk.LabelFrame(main_frame, text="Filtros", padding="10")
        filtros_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Frame para organizar filtros
        grid_frame = ttk.Frame(filtros_frame)
        grid_frame.pack(fill=tk.X)
        
        # Obtener lista de clientes únicos de la base de datos
        rutas = self.db_manager.obtener_todas_rutas()
        clientes_set = set()
        for ruta in rutas:
            cliente = ruta.get('cliente')
            if cliente and cliente.strip():
                clientes_set.add(cliente.strip())
        
        lista_clientes = ['Todos'] + sorted(list(clientes_set))
        
        # Label para filtro de cliente
        ttk.Label(grid_frame, text="Cliente:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5), pady=5)
        
        # Combobox de cliente
        combo_cliente = ttk.Combobox(grid_frame, values=lista_clientes, 
                                    state='readonly', width=30)
        combo_cliente.grid(row=0, column=1, sticky=tk.W, padx=(0, 20), pady=5)
        combo_cliente.current(0)  # Seleccionar "Todos" por defecto
        
        # Botones de acción de filtros
        btn_filtrar = ttk.Button(grid_frame, text="🔍 Filtrar", 
                                command=lambda: aplicar_filtro_rutas())
        btn_filtrar.grid(row=0, column=2, padx=5, pady=5)
        
        btn_limpiar = ttk.Button(grid_frame, text="🗑️ Limpiar", 
                                command=lambda: limpiar_filtro_rutas())
        btn_limpiar.grid(row=0, column=3, padx=5, pady=5)
        
        # ==================== SECCIÓN DE RESULTADOS ====================
        resultados_frame = ttk.LabelFrame(main_frame, text="Resultados", padding="5")
        resultados_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Frame de controles de ordenamiento
        orden_frame = ttk.Frame(resultados_frame)
        orden_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(orden_frame, text="Ordenar por:").pack(side=tk.LEFT, padx=(0, 5))
        
        columnas_orden = [
            "ID Ruta", "Clave", "Nombre Ruta", "Cliente", "Origen", "Entrega", 
            "Base", "Ensamble", "Kilómetros", "Tipo Viaje", "Flete", "Pago Operador"
        ]
        combo_orden = ttk.Combobox(orden_frame, values=columnas_orden, 
                                state='readonly', width=15)
        combo_orden.pack(side=tk.LEFT, padx=(0, 10))
        
        btn_asc = ttk.Button(orden_frame, text="▲ Asc", width=8,
                            command=lambda: ordenar_rutas(True))
        btn_asc.pack(side=tk.LEFT, padx=2)
        
        btn_desc = ttk.Button(orden_frame, text="▼ Desc", width=8,
                            command=lambda: ordenar_rutas(False))
        btn_desc.pack(side=tk.LEFT, padx=2)
        
        # Crear Sheet para rutas
        sheet_rutas = Sheet(
            resultados_frame,
            headers=[],
            data=[[]],
            width=1350,
            height=450,
            show_horizontal_scrollbar=True,
            show_vertical_scrollbar=True,
            enable_edit=False
        )
        sheet_rutas.pack(fill=tk.BOTH, expand=True)
        
        sheet_rutas.enable_bindings((
            "single_select",
            "row_select",
            "column_width_resize",
            "arrowkeys",
            "right_click_popup_menu",
            "copy",
            "column_select",
        ))
        
        # Variables para datos y ordenamiento
        datos_rutas_actuales = []
        columna_ordenada = None
        orden_ascendente = True
        
        # ==================== FUNCIONES AUXILIARES ====================
        
        def cargar_rutas_en_sheet(rutas_filtradas):
            """Carga las rutas en el tksheet"""
            nonlocal datos_rutas_actuales
            datos_rutas_actuales = rutas_filtradas
            
            # Definir headers
            headers = [
                "ID Ruta", "Clave", "Nombre Ruta", "Cliente", "Origen", "Entrega",
                "Base", "Ensamble", "Kilómetros", "Tipo Viaje", "Flete", 
                "Tipo Cobro", "Pago Operador", "Tipo Cobro Op", "Estado"
            ]
            
            # Preparar datos
            data = []
            for ruta in rutas_filtradas:
                fila = [
                    str(ruta.get('id_ruta', '')),
                    str(ruta.get('clave', '')),
                    str(ruta.get('nombre_ruta', '')),
                    str(ruta.get('cliente', '')),
                    str(ruta.get('origen', '')),
                    str(ruta.get('entrega', '')),
                    str(ruta.get('base', '')),
                    str(ruta.get('ensamble', '')),
                    f"{ruta.get('kilometros', 0):,.0f}" if ruta.get('kilometros') else '',
                    str(ruta.get('tipo_viaje', '')),
                    f"${ruta.get('flete', 0):,.2f}" if ruta.get('flete') else '',
                    str(ruta.get('tipo_cobro', '')),
                    f"${ruta.get('pago_operador', 0):,.2f}" if ruta.get('pago_operador') else '',
                    str(ruta.get('tipo_cobro_op', '')),
                    obtener_estado_ruta(ruta)
                ]
                data.append(fila)
            
            # Configurar sheet
            sheet_rutas.headers(headers)
            sheet_rutas.set_sheet_data(data)
            
            # Ajustar anchos de columna
            anchos = {
                0: 200,  # ID Ruta
                1: 80,   # Clave
                2: 150,  # Nombre Ruta
                3: 150,  # Cliente
                4: 120,  # Origen
                5: 120,  # Entrega
                6: 100,  # Base
                7: 100,  # Ensamble
                8: 100,  # Kilómetros
                9: 100,  # Tipo Viaje
                10: 100, # Flete
                11: 100, # Tipo Cobro
                12: 120, # Pago Operador
                13: 100, # Tipo Cobro Op
                14: 120  # Estado
            }
            
            for col, ancho in anchos.items():
                sheet_rutas.column_width(column=col, width=ancho)
            
            actualizar_contador(len(data))
        
        def obtener_estado_ruta(ruta):
            """Determina el estado de una ruta según origen_datos"""
            origen_datos = ruta.get('origen_datos', '')
            if origen_datos == 'completo':
                return 'Completo'
            elif origen_datos == 'solo_cliente':
                return 'Solo Cliente'
            elif origen_datos == 'solo_costos':
                return 'Solo Costos'
            return ''
        
        def aplicar_filtro_rutas():
            """Aplica el filtro de cliente seleccionado"""
            cliente_seleccionado = combo_cliente.get()
            
            # Obtener todas las rutas
            todas_rutas = self.db_manager.obtener_todas_rutas()
            
            if cliente_seleccionado == 'Todos':
                rutas_filtradas = todas_rutas
            else:
                rutas_filtradas = [
                    r for r in todas_rutas 
                    if r.get('cliente', '').strip() == cliente_seleccionado
                ]
            
            cargar_rutas_en_sheet(rutas_filtradas)
            actualizar_estado_label(f"Filtrado por cliente: {cliente_seleccionado}")
        
        def limpiar_filtro_rutas():
            """Limpia el filtro y muestra todas las rutas"""
            combo_cliente.current(0)  # Seleccionar "Todos"
            aplicar_filtro_rutas()
        
        def ordenar_rutas(ascendente):
            """Ordena las rutas según la columna seleccionada"""
            nonlocal datos_rutas_actuales, columna_ordenada, orden_ascendente
            
            columna_nombre = combo_orden.get()
            if not columna_nombre:
                messagebox.showwarning("Ordenar", "Selecciona una columna para ordenar")
                return
            
            # Mapeo de nombres de columna a claves de ruta
            mapeo_columnas = {
                "ID Ruta": "id_ruta",
                "Clave": "clave",
                "Nombre Ruta": "nombre_ruta",
                "Cliente": "cliente",
                "Origen": "origen",
                "Entrega": "entrega",
                "Base": "base",
                "Ensamble": "ensamble",
                "Kilómetros": "kilometros",
                "Tipo Viaje": "tipo_viaje",
                "Flete": "flete",
                "Pago Operador": "pago_operador"
            }
            
            clave_ordenar = mapeo_columnas.get(columna_nombre)
            if not clave_ordenar:
                return
            
            # Ordenar
            try:
                def get_valor_orden(ruta):
                    valor = ruta.get(clave_ordenar)
                    if valor is None:
                        return '' if isinstance(valor, str) else 0
                    return valor
                
                datos_rutas_actuales.sort(key=get_valor_orden, reverse=not ascendente)
                cargar_rutas_en_sheet(datos_rutas_actuales)
                
                direccion = "ascendente" if ascendente else "descendente"
                actualizar_estado_label(f"Ordenado por {columna_nombre} ({direccion})")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al ordenar:\n{e}")
        
        def exportar_rutas_excel():
            """Exporta las rutas actuales a Excel"""
            if not datos_rutas_actuales:
                messagebox.showwarning("Exportar", "No hay datos para exportar")
                return
            
            # Diálogo para guardar archivo
            archivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=f"rutas_export_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            
            if not archivo:
                return
            
            try:
                # Crear DataFrame desde los datos del sheet
                headers = sheet_rutas.headers()
                data = sheet_rutas.get_sheet_data()
                
                df = pd.DataFrame(data, columns=headers)
                df.to_excel(archivo, index=False)
                
                messagebox.showinfo("Exportar", f"Rutas exportadas exitosamente a:\n{archivo}")
                actualizar_estado_label(f"Exportado: {len(data)} rutas")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error exportando:\n{e}")
        
        def actualizar_estado_label(mensaje):
            """Actualiza el label de estado"""
            label_estado.config(text=mensaje)
            ventana_rutas.update_idletasks()
        
        def actualizar_contador(num):
            """Actualiza el contador de rutas"""
            label_stats.config(text=f"Rutas: {num:,}")
        
        # ==================== BOTONES PRINCIPALES ====================
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        btn_exportar = ttk.Button(btn_frame, text="📊 Exportar a Excel", 
                                command=exportar_rutas_excel)
        btn_exportar.pack(side=tk.LEFT, padx=5)
        
        btn_cerrar = ttk.Button(btn_frame, text="❌ Cerrar", 
                            command=al_cerrar)
        btn_cerrar.pack(side=tk.RIGHT, padx=5)
        
        # ==================== BARRA DE ESTADO ====================
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(5, 0))
        
        label_estado = ttk.Label(status_frame, text="Listo", relief=tk.SUNKEN, anchor=tk.W)
        label_estado.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        label_stats = ttk.Label(status_frame, text="Rutas: 0", relief=tk.SUNKEN, anchor=tk.E, width=20)
        label_stats.pack(side=tk.RIGHT)
        
        # ==================== CARGAR DATOS INICIALES ====================
        aplicar_filtro_rutas()

    def actualizar_estado(self, mensaje):
        """Actualiza mensaje en barra de estado"""
        if hasattr(self, 'label_estado'):
            self.label_estado.config(text=mensaje)
            self.ventana_consulta.update_idletasks()
    
    def actualizar_stats(self, num_registros):
        """Actualiza contador de registros"""
        if hasattr(self, 'label_stats'):
            self.label_stats.config(text=f"Registros: {num_registros:,}")

    def abrir_modo_actualizacion(self, ventana_consulta):
        """Abre el módulo de recepción en modo actualización"""
        # Cerrar ventana de consulta
        ventana_consulta.destroy()
        
        # Limpiar datos de carga previos
        self.datamon = None
        self.datacas = None
        self.dataayc = None
        self.datagastos = None
        self.dataclientes = None
        self.datacostos = None
        
        # Cambiar a modo actualización
        self.modo = 'actualizacion'
        
        # Mostrar ventana principal y recrear interfaz
        self.root.deiconify()
        
        # Limpiar ventana principal
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Recrear interfaz en modo actualización
        self.crear_interfaz_carga()
        
        self.log("="*50)
        self.log("MODO ACTUALIZACIÓN")
        self.log("="*50)
        self.log("Carga archivos Excel para actualizar la base de datos")
        self.log("Todos los archivos son opcionales")
    
    def cerrar_consulta(self, ventana):
        """Cierra consulta y vuelve a carga"""
        ventana.destroy()
        self.root.deiconify()

class FormularioViaje:
    """
    Formulario para añadir viajes manualmente.
    Permite añadir múltiples viajes en un solo proceso.
    """
    
    def __init__(self, parent, db_manager):
        self.parent = parent
        self.db_manager = db_manager
        
        # Lista de viajes temporales
        self.viajes_temporales = []
        self.indice_actual = -1
        
        # Obtener datos de la BD para dropdowns
        self.ids_ruta = self.obtener_ids_ruta()
        self.operadores = self.obtener_operadores()
        self.rutas_info = self.obtener_info_rutas()
        
        # Crear ventana
        self.ventana = tk.Toplevel(parent)
        self.ventana.title("Añadir Viajes Manualmente")
        self.ventana.geometry("700x800")
        
        # Variables del formulario
        self.vars = {}
        
        # Crear interfaz
        self.crear_interfaz()
        
        # Añadir primer viaje vacío
        self.nuevo_viaje()
    
    def obtener_ids_ruta(self):
        """Obtiene lista única de IDs de ruta"""
        ids = set()
        for ruta in self.db_manager.obtener_todas_rutas():
            id_ruta = ruta.get('id_ruta')
            if id_ruta:
                ids.add(id_ruta)
        return sorted(list(ids))
    
    def obtener_operadores(self):
        """Obtiene lista única de operadores"""
        ops = set()
        for id_viaje, datos in self.db_manager.data['viajes'].items():
            op = datos.get('operador')
            if op:
                ops.add(str(op).strip())
        return sorted(list(ops))
    
    def obtener_info_rutas(self):
        """Obtiene información de rutas agrupada por ID"""
        info = {}
        for ruta in self.db_manager.obtener_todas_rutas():
            id_ruta = ruta.get('id_ruta')
            if id_ruta:
                if id_ruta not in info:
                    info[id_ruta] = {
                        'pagos_operador': set(),
                        'tipos_viaje': []
                    }
                
                pago = ruta.get('pago_operador')
                if pago is not None:
                    info[id_ruta]['pagos_operador'].add(float(pago))
                
                tipo = ruta.get('tipo_viaje')
                if tipo:
                    info[id_ruta]['tipos_viaje'].append(tipo)
        
        # Convertir sets a listas ordenadas
        for id_ruta in info:
            info[id_ruta]['pagos_operador'] = sorted(list(info[id_ruta]['pagos_operador']))
        
        return info
    
    def crear_interfaz(self):
        """Crea la interfaz del formulario"""
        main_frame = ttk.Frame(self.ventana, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        titulo = ttk.Label(main_frame, text="Añadir Viaje Manual", 
                          font=('Arial', 14, 'bold'))
        titulo.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Frame para el formulario con scroll
        form_canvas = tk.Canvas(main_frame, height=550)
        form_scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=form_canvas.yview)
        form_frame = ttk.Frame(form_canvas)
        
        form_frame.bind(
            "<Configure>",
            lambda e: form_canvas.configure(scrollregion=form_canvas.bbox("all"))
        )
        
        form_canvas.create_window((0, 0), window=form_frame, anchor="nw")
        form_canvas.configure(yscrollcommand=form_scrollbar.set)
        
        form_canvas.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), columnspan=2)
        form_scrollbar.grid(row=1, column=2, sticky=(tk.N, tk.S))
        
        # Crear campos del formulario
        row = 0
        
        # ID Viaje
        row = self.crear_campo(form_frame, row, "ID Viaje *", "id_viaje", tk.Entry)
        
        # Fecha Inicio
        row = self.crear_campo(form_frame, row, "Fecha Inicio * (YYYY-MM-DD HH:MM:SS)", 
                              "fecha_inicio", tk.Entry)
        
        # Fecha Final
        row = self.crear_campo(form_frame, row, "Fecha Final * (YYYY-MM-DD HH:MM:SS)", 
                              "fecha_final", tk.Entry)
        
        # Ruta (Combobox con búsqueda)
        ttk.Label(form_frame, text="Ruta *").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.vars['ruta'] = tk.StringVar()
        self.combo_ruta = ttk.Combobox(form_frame, textvariable=self.vars['ruta'],
                                       values=self.ids_ruta, width=40)
        self.combo_ruta.grid(row=row, column=1, sticky=tk.W, pady=5)
        self.combo_ruta.bind('<<ComboboxSelected>>', self.on_ruta_selected)
        row += 1
        
        # Operador (Combobox editable)
        ttk.Label(form_frame, text="Operador *").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.vars['operador'] = tk.StringVar()
        self.combo_operador = ttk.Combobox(form_frame, textvariable=self.vars['operador'],
                                          values=self.operadores, width=40)
        self.combo_operador.grid(row=row, column=1, sticky=tk.W, pady=5)
        row += 1
        
        # Tracto
        row = self.crear_campo(form_frame, row, "Tracto *", "tracto", tk.Entry)
        
        # Ensamble
        row = self.crear_campo(form_frame, row, "Ensamble", "ensamble", tk.Entry)
        
        # Frontal
        row = self.crear_campo(form_frame, row, "Frontal", "frontal", tk.Entry)
        
        # Dolly
        row = self.crear_campo(form_frame, row, "Dolly", "dolly", tk.Entry)
        
        # Trasero
        row = self.crear_campo(form_frame, row, "Trasero", "trasero", tk.Entry)
        
        # Gastos
        row = self.crear_campo(form_frame, row, "Gastos", "gasto", tk.Entry)
        
        # Comisión (Combobox, inicialmente deshabilitado)
        ttk.Label(form_frame, text="Comisión *").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.vars['comision'] = tk.StringVar()
        self.combo_comision = ttk.Combobox(form_frame, textvariable=self.vars['comision'],
                                          values=[], width=40, state=tk.DISABLED)
        self.combo_comision.grid(row=row, column=1, sticky=tk.W, pady=5)
        ttk.Label(form_frame, text="(Se habilita al elegir ruta)", 
                 foreground='#666', font=('Arial', 8)).grid(row=row, column=2, sticky=tk.W)
        row += 1
        
        # Separador
        ttk.Separator(main_frame, orient='horizontal').grid(row=2, column=0, columnspan=2, 
                                                            sticky=(tk.W, tk.E), pady=15)
        
        # Frame de navegación
        nav_frame = ttk.Frame(main_frame)
        nav_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        self.label_contador = ttk.Label(nav_frame, text="Viaje 1 de 1", 
                                        font=('Arial', 10, 'bold'))
        self.label_contador.pack(pady=5)
        
        btn_nav_frame = ttk.Frame(nav_frame)
        btn_nav_frame.pack()
        
        self.btn_anterior = ttk.Button(btn_nav_frame, text="◄ Anterior", 
                                       command=self.viaje_anterior, state=tk.DISABLED)
        self.btn_anterior.pack(side=tk.LEFT, padx=5)
        
        self.btn_siguiente = ttk.Button(btn_nav_frame, text="Siguiente ►", 
                                       command=self.viaje_siguiente, state=tk.DISABLED)
        self.btn_siguiente.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_nav_frame, text="➕ Nuevo Viaje", 
                  command=self.nuevo_viaje).pack(side=tk.LEFT, padx=5)
        
        # Botones finales
        btn_final_frame = ttk.Frame(main_frame)
        btn_final_frame.grid(row=4, column=0, columnspan=2, pady=15)
        
        ttk.Button(btn_final_frame, text="✅ Guardar Todos", 
                  command=self.guardar_todos).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(btn_final_frame, text="❌ Cancelar", 
                  command=self.ventana.destroy).pack(side=tk.LEFT, padx=10)
        
        # Configurar peso de grid
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
    
    def crear_campo(self, parent, row, label, var_name, widget_class):
        """Helper para crear un campo del formulario"""
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=5)
        
        self.vars[var_name] = tk.StringVar()
        widget = widget_class(parent, textvariable=self.vars[var_name], width=42)
        widget.grid(row=row, column=1, sticky=tk.W, pady=5)
        
        return row + 1
    
    def on_ruta_selected(self, event=None):
        """Cuando se selecciona una ruta, habilitar comisión y mostrar opciones"""
        id_ruta = self.vars['ruta'].get()
        
        if id_ruta in self.rutas_info:
            pagos = self.rutas_info[id_ruta]['pagos_operador']
            
            # Habilitar y llenar combo de comisión
            self.combo_comision.config(state='readonly')
            self.combo_comision.config(values=[f"{p:,.2f}" for p in pagos])
            
            if pagos:
                # Autoseleccionar el primero
                self.vars['comision'].set(f"{pagos[0]:,.2f}")
        else:
            # Deshabilitar comisión si no hay ruta válida
            self.combo_comision.config(state=tk.DISABLED)
            self.vars['comision'].set('')
    
    def capturar_viaje_actual(self):
        """Captura los valores actuales del formulario"""
        return {
            'id_viaje': self.vars['id_viaje'].get().strip(),
            'fecha_inicio': self.vars['fecha_inicio'].get().strip(),
            'fecha_final': self.vars['fecha_final'].get().strip(),
            'ruta': self.vars['ruta'].get().strip(),
            'operador': self.vars['operador'].get().strip(),
            'tracto': self.vars['tracto'].get().strip(),
            'ensamble': self.vars['ensamble'].get().strip(),
            'frontal': self.vars['frontal'].get().strip(),
            'dolly': self.vars['dolly'].get().strip(),
            'trasero': self.vars['trasero'].get().strip(),
            'gasto': self.vars['gasto'].get().strip(),
            'comision': self.vars['comision'].get().strip(),
        }
    
    def cargar_viaje_en_formulario(self, viaje):
        """Carga un viaje en el formulario"""
        for key, value in viaje.items():
            if key in self.vars:
                self.vars[key].set(value)
        
        # Trigger ruta selection para habilitar comisión
        if viaje.get('ruta'):
            self.on_ruta_selected()
    
    def limpiar_formulario(self):
        """Limpia todos los campos"""
        for var in self.vars.values():
            var.set('')
        
        self.combo_comision.config(state=tk.DISABLED)
    
    def nuevo_viaje(self):
        """Crea un nuevo viaje vacío"""
        # Guardar viaje actual si hay datos
        if self.indice_actual >= 0:
            viaje_actual = self.capturar_viaje_actual()
            self.viajes_temporales[self.indice_actual] = viaje_actual
        
        # Añadir nuevo viaje vacío
        self.viajes_temporales.append({})
        self.indice_actual = len(self.viajes_temporales) - 1
        
        # Limpiar formulario
        self.limpiar_formulario()
        
        # Actualizar controles
        self.actualizar_navegacion()
    
    def viaje_anterior(self):
        """Navega al viaje anterior"""
        if self.indice_actual > 0:
            # Guardar actual
            self.viajes_temporales[self.indice_actual] = self.capturar_viaje_actual()
            
            # Ir al anterior
            self.indice_actual -= 1
            self.cargar_viaje_en_formulario(self.viajes_temporales[self.indice_actual])
            
            self.actualizar_navegacion()
    
    def viaje_siguiente(self):
        """Navega al viaje siguiente"""
        if self.indice_actual < len(self.viajes_temporales) - 1:
            # Guardar actual
            self.viajes_temporales[self.indice_actual] = self.capturar_viaje_actual()
            
            # Ir al siguiente
            self.indice_actual += 1
            self.cargar_viaje_en_formulario(self.viajes_temporales[self.indice_actual])
            
            self.actualizar_navegacion()
    
    def actualizar_navegacion(self):
        """Actualiza el estado de los botones de navegación"""
        total = len(self.viajes_temporales)
        actual = self.indice_actual + 1
        
        self.label_contador.config(text=f"Viaje {actual} de {total}")
        
        # Botón anterior
        self.btn_anterior.config(state=tk.NORMAL if self.indice_actual > 0 else tk.DISABLED)
        
        # Botón siguiente
        self.btn_siguiente.config(state=tk.NORMAL if self.indice_actual < total - 1 else tk.DISABLED)
    
    def validar_viaje(self, viaje):
        """Valida que un viaje tenga los datos requeridos correctos"""
        errores = []
        
        # ID viaje requerido y único
        if not viaje.get('id_viaje'):
            errores.append("ID de viaje es requerido")
        elif viaje['id_viaje'] in self.db_manager.data['viajes']:
            errores.append(f"ID de viaje '{viaje['id_viaje']}' ya existe en la base de datos")
        
        # Fechas requeridas
        if not viaje.get('fecha_inicio'):
            errores.append("Fecha de inicio es requerida")
        if not viaje.get('fecha_final'):
            errores.append("Fecha final es requerida")
        
        # Validar formato de fechas
        if viaje.get('fecha_inicio') and viaje.get('fecha_final'):
            try:
                fi = pd.to_datetime(viaje['fecha_inicio'])
                ff = pd.to_datetime(viaje['fecha_final'])
                
                if ff < fi:
                    errores.append("Fecha final no puede ser anterior a fecha de inicio")
            except:
                errores.append("Formato de fecha inválido. Usa: YYYY-MM-DD HH:MM:SS")
        
        # Ruta requerida y debe existir
        if not viaje.get('ruta'):
            errores.append("Ruta es requerida")
        elif viaje['ruta'] not in self.ids_ruta:
            errores.append(f"Ruta '{viaje['ruta']}' no existe en la base de datos")
        
        # Operador requerido
        if not viaje.get('operador'):
            errores.append("Operador es requerido")
        
        # Tracto requerido
        if not viaje.get('tracto'):
            errores.append("Tracto es requerido")
        
        # Comisión requerida y debe coincidir con pago operador
        if not viaje.get('comision'):
            errores.append("Comisión es requerida")
        elif viaje.get('ruta') in self.rutas_info:
            try:
                comision_val = float(viaje['comision'].replace(',', ''))
                pagos_validos = self.rutas_info[viaje['ruta']]['pagos_operador']
                
                if comision_val not in pagos_validos:
                    errores.append(f"Comisión debe ser uno de: {[f'${p:,.2f}' for p in pagos_validos]}")
            except:
                errores.append("Comisión debe ser un número válido")
        
        return errores
    
    def guardar_todos(self):
        """Valida y guarda todos los viajes"""
        # Guardar viaje actual
        self.viajes_temporales[self.indice_actual] = self.capturar_viaje_actual()
        
        # Validar todos los viajes
        todos_errores = []
        viajes_validos = []
        
        for i, viaje in enumerate(self.viajes_temporales, 1):
            errores = self.validar_viaje(viaje)
            if errores:
                todos_errores.append(f"\nViaje {i}:")
                todos_errores.extend([f"  - {e}" for e in errores])
            else:
                viajes_validos.append(viaje)
        
        # Si hay errores, mostrar y no guardar
        if todos_errores:
            messagebox.showerror("Errores de Validación", 
                            "Se encontraron errores:\n" + "\n".join(todos_errores))
            return
        
        # Confirmar guardado
        if not messagebox.askyesno("Confirmar", 
                                f"¿Guardar {len(viajes_validos)} viaje(s)?"):
            return
        
        # Guardar en la base de datos
        try:
            from core_viajes import crear_id_ruta
            
            for viaje in viajes_validos:
                # Crear datos del viaje
                datos_viaje = {
                    'fecha_inicio': viaje['fecha_inicio'],
                    'fecha_final': viaje['fecha_final'],
                    'ruta': viaje['ruta'],
                    'operador': viaje['operador'],
                    'tracto': viaje['tracto'],
                    'advertencias': [],
                    'origen': 'manual'
                }
                
                # Campos opcionales
                if viaje.get('ensamble'):
                    datos_viaje['ensamble'] = viaje['ensamble']
                if viaje.get('frontal'):
                    datos_viaje['frontal'] = viaje['frontal']
                if viaje.get('dolly'):
                    datos_viaje['dolly'] = viaje['dolly']
                if viaje.get('trasero'):
                    datos_viaje['trasero'] = viaje['trasero']
                if viaje.get('gasto'):
                    datos_viaje['gasto'] = float(viaje['gasto'])
                if viaje.get('comision'):
                    datos_viaje['comision'] = float(viaje['comision'].replace(',', ''))
                
                # Crear ID de ruta si es posible
                datos_viaje['id_ruta'] = viaje['ruta']
                
                # Agregar a la BD
                self.db_manager.agregar_viaje(viaje['id_viaje'], datos_viaje)
            
            # NUEVO ORDEN: Primero detectar dobles
            from core_viajes import detectar_viajes_dobles
            detectar_viajes_dobles(self.db_manager)
            
            # Luego recalcular relaciones (que ya ignora IRREAL y DOBLE)
            messagebox.showinfo("Procesando", 
                            "Viajes guardados. Calculando relaciones con casetas...")
            
            self.recalcular_relaciones_nuevos_viajes([v['id_viaje'] for v in viajes_validos])
            
            # Detectar estados de análisis
            from core_viajes import detectar_estados_analisis, marcar_viajes_finales
            detectar_estados_analisis(self.db_manager)
            marcar_viajes_finales(self.db_manager)
            
            # Finalmente guardar BD
            self.db_manager.guardar_datos()
            
            messagebox.showinfo("Éxito", 
                            f"{len(viajes_validos)} viaje(s) guardado(s) exitosamente")
            
            self.ventana.destroy()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error guardando viajes:\n{e}")
    
    def recalcular_relaciones_nuevos_viajes(self, ids_viajes):
        """Recalcula relaciones de casetas solo para los viajes nuevos"""
        from core_viajes import IndicesCasetas
        
        # Construir índices
        indices = IndicesCasetas()
        indices.construir_desde_db(self.db_manager.data['casetas'])
        
        # Filtrar solo viajes válidos (sin IRREAL ni DOBLE)
        ids_viajes_validos = []
        
        for id_viaje in ids_viajes:
            datos_viaje = self.db_manager.data['viajes'][id_viaje]
            advertencias = datos_viaje.get('advertencias', [])
            
            # Ignorar viajes IRREAL o DOBLE
            if 'IRREAL' in advertencias or 'DOBLE' in advertencias:
                continue
            
            ids_viajes_validos.append(id_viaje)
        
        # Procesar solo viajes válidos
        for id_viaje in ids_viajes_validos:
            datos_viaje = self.db_manager.data['viajes'][id_viaje]
            tracto = datos_viaje.get('tracto')
            
            if not tracto:
                continue
            
            try:
                fecha_inicio = pd.to_datetime(datos_viaje['fecha_inicio'])
                fecha_final = pd.to_datetime(datos_viaje['fecha_final'])
            except:
                continue
            
            # Buscar casetas dentro del rango
            ids_casetas = indices.buscar_por_rango(tracto, fecha_inicio, fecha_final)
            
            for id_caseta in ids_casetas:
                self.db_manager.data['relaciones'].append({
                    'id_viaje': id_viaje,
                    'id_caseta': id_caseta,
                    'is_fuera': False
                })

def main():
    # Crear ventana con o sin DnD
    if TKDND_AVAILABLE:
        try:
            root = TkinterDnD.Tk()
        except:
            root = tk.Tk()
    else:
        root = tk.Tk()
    
    try:
        app = AplicacionViajes(root)
        root.mainloop()
    except Exception as e:
        import traceback
        traceback.print_exc()
        try:
            messagebox.showerror("Error", f"Error crítico:\n{e}")
        except:
            print(f"Error crítico: {e}")


if __name__ == "__main__":
    main()