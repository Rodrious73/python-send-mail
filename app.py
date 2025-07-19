import pandas as pd
import smtplib
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import codecs
import threading
import platform
import subprocess
import tempfile
import json
import tempfile

load_dotenv()

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestor de Envío de Correos Masivos")
        self.root.geometry("1200x800")
        
        # Establecer el icono de la aplicación
        try:
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"No se pudo cargar el icono: {str(e)}")
        
        # Variables
        self.df = None
        self.plantilla_html = None
        
        # Valores iniciales predeterminados que se actualizarán con la carga de configuración
        self.email_account = os.getenv("email_account", "")
        self.password_account = os.getenv("password_account", "")
        self.name_account = os.getenv("name_account", "")
        self.asunto_default = "Información Universitaria"
        self.mensaje_default = "Este es un mensaje personalizado para recordarte sobre los trámites universitarios pendientes."
        self.ultima_seleccion = None  # Variable para guardar la última selección en la tabla
        
        # Variables de control
        self.usar_html = tk.BooleanVar(value=True)
        self.rango_inicio = tk.IntVar(value=1)
        self.rango_fin = tk.IntVar(value=50)
        
        # Configuración de columnas - nueva estructura universitaria
        self.columnas_display = ['#', 'Nombres', 'Apellidos', 'Correo', 'Facultad', 'Escuela', 'Cod.Universitario']
        self.columnas_requeridas = ['id', 'Cod.Universitario', 'Nombres', 'Apellidos', 'Facultad', 'Escuela', 'Correo']
        
        # Configuración de separadores CSV
        self.separadores_csv = {
            "Punto y coma (;)": ";",
            "Coma (,)": ",",
            "Tabulación (\\t)": "\t",
            "Pipe (|)": "|",
            "Espacio ( )": " "
        }
        
        self.setup_ui()
        self.cargar_datos()
        self.cargar_plantilla()
        
        # Una vez configurada la interfaz, cargar la configuración personalizada
        self.cargar_configuracion()
        
    def setup_ui(self):
        # Crear notebook para pestañas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Agregar evento para detectar cambios de pestaña
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        
        # Pestaña de Vista Previa
        self.frame_preview = ttk.Frame(self.notebook)
        self.notebook.add(self.frame_preview, text="Vista Previa de Correos")
        self.setup_preview_tab()
        
        # Pestaña de Gestión de Datos
        self.frame_data = ttk.Frame(self.notebook)
        self.notebook.add(self.frame_data, text="Gestión de Datos")
        self.setup_data_tab()
        
        # Pestaña de Envío
        self.frame_send = ttk.Frame(self.notebook)
        self.notebook.add(self.frame_send, text="Envío de Correos")
        self.setup_send_tab()
        
        # Pestaña de Configuración
        self.frame_config = ttk.Frame(self.notebook)
        self.notebook.add(self.frame_config, text="Configuración")
        self.setup_config_tab()
        
    def setup_preview_tab(self):
        # Frame superior para controles
        control_frame = ttk.Frame(self.frame_preview)
        control_frame.pack(fill='x', padx=5, pady=5)
        
        # Selector de rango
        ttk.Label(control_frame, text="Rango de correos:").pack(side='left', padx=5)
        ttk.Label(control_frame, text="Desde:").pack(side='left', padx=(20,5))
        ttk.Entry(control_frame, textvariable=self.rango_inicio, width=10).pack(side='left', padx=5)
        ttk.Label(control_frame, text="Hasta:").pack(side='left', padx=5)
        ttk.Entry(control_frame, textvariable=self.rango_fin, width=10).pack(side='left', padx=5)
        
        # Etiqueta para el total de registros
        self.lbl_total_registros = ttk.Label(control_frame, text="Total: 0 registros")
        self.lbl_total_registros.pack(side='left', padx=(20,5))
        
        ttk.Button(control_frame, text="Actualizar Vista", command=self.actualizar_preview).pack(side='left', padx=10)
        
        # Botón para cargar archivo Excel
        ttk.Button(control_frame, text="Cargar Excel/CSV", command=self.cargar_excel_personalizado).pack(side='left', padx=10)
        
        # Selector de vista
        ttk.Label(control_frame, text="Modo:").pack(side='right', padx=5)
        self.modo_vista = ttk.Combobox(control_frame, values=["Tabla Resumen", "Vista Detallada"], state="readonly")
        self.modo_vista.set("Tabla Resumen")
        self.modo_vista.pack(side='right', padx=5)
        self.modo_vista.bind('<<ComboboxSelected>>', lambda e: self.actualizar_preview())
        
        # Frame para la vista previa
        preview_frame = ttk.Frame(self.frame_preview)
        preview_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Treeview para mostrar datos - usando las nuevas columnas
        self.tree = ttk.Treeview(preview_frame, columns=self.columnas_display, show='headings', height=15)
        
        # Configurar columnas dinámicamente
        for col in self.columnas_display:
            self.tree.heading(col, text=col)
            if col == '#':
                self.tree.column(col, width=50, anchor='center')
            elif col == 'Nombres':
                self.tree.column(col, width=150)
            elif col == 'Apellidos':
                self.tree.column(col, width=150)
            elif col == 'Correo':
                self.tree.column(col, width=220)
            elif col == 'Facultad':
                self.tree.column(col, width=180)
            elif col == 'Escuela':
                self.tree.column(col, width=180)
            elif col == 'Cod.Universitario':
                self.tree.column(col, width=120)
        
        # Scrollbars
        scrollbar_v = ttk.Scrollbar(preview_frame, orient='vertical', command=self.tree.yview)
        scrollbar_h = ttk.Scrollbar(preview_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)
        
        # Pack treeview y scrollbars
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar_v.pack(side='right', fill='y')
        scrollbar_h.pack(side='bottom', fill='x')
        
        # Frame para vista detallada
        self.detail_frame = ttk.Frame(self.frame_preview)
        
        # Bind para selección en el tree
        self.tree.bind('<<TreeviewSelect>>', self.mostrar_detalle)
        
    def setup_data_tab(self):
        # Frame para gestión individual
        individual_frame = ttk.LabelFrame(self.frame_data, text="Agregar/Editar Estudiante Individual")
        individual_frame.pack(fill='x', padx=5, pady=5)
        
        # Primera fila de campos
        ttk.Label(individual_frame, text="Nombres:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.entry_nombres = ttk.Entry(individual_frame, width=25)
        self.entry_nombres.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(individual_frame, text="Apellidos:").grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.entry_apellidos = ttk.Entry(individual_frame, width=25)
        self.entry_apellidos.grid(row=0, column=3, padx=5, pady=5)
        
        # Segunda fila de campos
        ttk.Label(individual_frame, text="Correo:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.entry_correo = ttk.Entry(individual_frame, width=25)
        self.entry_correo.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(individual_frame, text="Cod. Universitario:").grid(row=1, column=2, sticky='w', padx=5, pady=5)
        self.entry_codigo = ttk.Entry(individual_frame, width=25)
        self.entry_codigo.grid(row=1, column=3, padx=5, pady=5)
        
        # Tercera fila de campos
        ttk.Label(individual_frame, text="Facultad:").grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.entry_facultad = ttk.Entry(individual_frame, width=25)
        self.entry_facultad.grid(row=2, column=1, padx=5, pady=5)
        
        ttk.Label(individual_frame, text="Escuela:").grid(row=2, column=2, sticky='w', padx=5, pady=5)
        self.entry_escuela = ttk.Entry(individual_frame, width=25)
        self.entry_escuela.grid(row=2, column=3, padx=5, pady=5)
        
        # Botones de acción
        ttk.Button(individual_frame, text="Agregar", command=self.agregar_estudiante).grid(row=3, column=0, pady=10)
        ttk.Button(individual_frame, text="Actualizar Seleccionado", command=self.actualizar_estudiante).grid(row=3, column=1, pady=10)
        ttk.Button(individual_frame, text="Eliminar Seleccionado", command=self.eliminar_estudiante).grid(row=3, column=2, pady=10)
        ttk.Button(individual_frame, text="Limpiar Selección", command=self.limpiar_seleccion).grid(row=3, column=3, pady=10)
        
        # Frame para importación masiva
        masiva_frame = ttk.LabelFrame(self.frame_data, text="Importación de Archivo Excel/CSV")
        masiva_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Información sobre la estructura esperada
        info_text = "Estructura del archivo esperada (Excel/CSV): id, Cod.Universitario, Nombres, Apellidos, Facultad, Escuela, Correo"
        ttk.Label(masiva_frame, text=info_text, font=('Arial', 9, 'italic')).pack(anchor='w', padx=5, pady=5)
        
        button_frame = ttk.Frame(masiva_frame)
        button_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Button(button_frame, text="Cargar Excel/CSV", command=self.cargar_excel_personalizado).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Guardar Cambios", command=self.guardar_excel).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Recargar desde Excel", command=self.cargar_datos).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Abrir Excel", command=self.abrir_excel).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Crear Ejemplos (Excel/CSV)", command=self.crear_excel_ejemplo).pack(side='left', padx=5)
        
        # Botón temporal para probar el archivo bebita
        if os.path.exists("data/prueba bebita.csv"):
            ttk.Button(button_frame, text="Cargar 'prueba bebita.csv'", command=self.cargar_prueba_bebita).pack(side='left', padx=5)
        
    def setup_send_tab(self):
        # Frame de configuración de envío
        config_frame = ttk.LabelFrame(self.frame_send, text="Configuración de Envío")
        config_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Checkbutton(config_frame, text="Usar plantilla HTML", variable=self.usar_html).pack(anchor='w', padx=5, pady=5)
        
        # Rango de envío
        rango_frame = ttk.Frame(config_frame)
        rango_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(rango_frame, text="Enviar desde el correo:").pack(side='left', padx=5)
        self.send_inicio = ttk.Entry(rango_frame, width=10)
        self.send_inicio.pack(side='left', padx=5)
        self.send_inicio.insert(0, "1")
        
        ttk.Label(rango_frame, text="hasta el correo:").pack(side='left', padx=5)
        self.send_fin = ttk.Entry(rango_frame, width=10)
        self.send_fin.pack(side='left', padx=5)
        self.send_fin.insert(0, "50")
        
        # Botones de envío
        button_frame = ttk.Frame(config_frame)
        button_frame.pack(fill='x', padx=5, pady=10)
        
        self.btn_enviar = ttk.Button(button_frame, text="Enviar Correos", command=self.enviar_correos)
        self.btn_enviar.pack(side='left', padx=5)
        
        self.btn_test = ttk.Button(button_frame, text="Envío de Prueba", command=self.envio_prueba)
        self.btn_test.pack(side='left', padx=5)
        
        # Área de log
        log_frame = ttk.LabelFrame(self.frame_send, text="Log de Envío")
        log_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=20, width=80)
        self.log_text.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(log_frame, mode='determinate')
        self.progress.pack(fill='x', padx=5, pady=5)
        
    def setup_config_tab(self):
        # Configuración de cuenta
        cuenta_frame = ttk.LabelFrame(self.frame_config, text="Configuración de Cuenta de Correo")
        cuenta_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(cuenta_frame, text="Nombre del remitente:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.config_nombre = ttk.Entry(cuenta_frame, width=40)
        self.config_nombre.grid(row=0, column=1, padx=5, pady=5)
        self.config_nombre.insert(0, self.name_account or "")
        
        ttk.Label(cuenta_frame, text="Email del remitente:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.config_email = ttk.Entry(cuenta_frame, width=40)
        self.config_email.grid(row=1, column=1, padx=5, pady=5)
        self.config_email.insert(0, self.email_account or "")
        
        ttk.Label(cuenta_frame, text="Contraseña de aplicación:").grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.config_password = ttk.Entry(cuenta_frame, width=40, show="*")
        self.config_password.grid(row=2, column=1, padx=5, pady=5)
        self.config_password.insert(0, self.password_account or "")
        
        button_frame = ttk.Frame(cuenta_frame)
        button_frame.grid(row=3, column=1, pady=10, sticky='e')
        
        ttk.Button(button_frame, text="Guardar Configuración", command=self.guardar_valores_predeterminados).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Probar Conexión", command=self.probar_conexion).pack(side='left', padx=5)
        
        # Configuración de plantillas
        plantilla_frame = ttk.LabelFrame(self.frame_config, text="Configuración de Plantillas")
        plantilla_frame.pack(fill='x', padx=5, pady=5)
        
        plantilla_top_frame = ttk.Frame(plantilla_frame)
        plantilla_top_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Button(plantilla_top_frame, text="Seleccionar Plantilla HTML", command=self.seleccionar_plantilla).pack(side='left', padx=5)
        ttk.Button(plantilla_top_frame, text="Ver Vista Previa", command=self.mostrar_vista_previa_html).pack(side='left', padx=5)
        
        self.plantilla_path = ttk.Label(plantilla_frame, text="Plantilla actual: templates/email_template.html")
        self.plantilla_path.pack(anchor='w', padx=5, pady=5)
        
        # Configuración de mensaje predeterminado
        mensaje_frame = ttk.LabelFrame(self.frame_config, text="Configuración de Mensaje Predeterminado")
        mensaje_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        ttk.Label(mensaje_frame, text="Asunto predeterminado:").pack(anchor='w', padx=5, pady=5)
        self.entry_asunto_default = ttk.Entry(mensaje_frame, width=60)
        self.entry_asunto_default.pack(fill='x', padx=5, pady=5)
        self.entry_asunto_default.insert(0, self.asunto_default)
        
        ttk.Label(mensaje_frame, text="Mensaje predeterminado:").pack(anchor='w', padx=5, pady=5)
        self.text_mensaje_default = scrolledtext.ScrolledText(mensaje_frame, height=6)
        self.text_mensaje_default.pack(fill='both', expand=True, padx=5, pady=5)
        self.text_mensaje_default.insert('1.0', self.mensaje_default)
        
        boton_frame = ttk.Frame(mensaje_frame)
        boton_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Button(boton_frame, text="Guardar Configuración", 
                  command=self.guardar_valores_predeterminados).pack(side='right', padx=5, pady=5)
        
        # Frame para vista previa de plantilla HTML
        self.preview_html_frame = ttk.LabelFrame(self.frame_config, text="Vista Previa de Plantilla HTML")
        self.preview_html_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
    def cargar_datos(self):
        """Carga los datos del archivo Excel con la nueva estructura universitaria"""
        try:
            excel_path = "data/estudiantes.xlsx"
            if os.path.exists(excel_path):
                self.df = pd.read_excel(excel_path)
                
                # Verificar si tiene la estructura universitaria
                if all(col in self.df.columns for col in self.columnas_requeridas):
                    self.log("Datos universitarios cargados exitosamente desde " + excel_path)
                else:
                    # Intentar migrar desde estructura antigua o crear nueva
                    self.migrar_o_crear_estructura_universitaria()
                
                self.actualizar_preview()
            else:
                # Crear DataFrame vacío con las columnas universitarias
                self.crear_dataframe_universitario()
                self.log("Archivo Excel no encontrado. Se creó un archivo nuevo con estructura universitaria en " + excel_path)
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar datos: {str(e)}")
            self.log(f"Error al cargar datos: {str(e)}")
    
    def migrar_o_crear_estructura_universitaria(self):
        """Migra datos antiguos o crea nueva estructura universitaria"""
        try:
            # Si tiene columnas antiguas, intentar migrar
            if 'Name' in self.df.columns and 'Email' in self.df.columns:
                # Crear DataFrame con nueva estructura, manteniendo datos existentes donde sea posible
                df_nuevo = pd.DataFrame(columns=self.columnas_requeridas)
                
                for i, row in self.df.iterrows():
                    nueva_fila = {
                        'id': i + 1,
                        'Cod.Universitario': '',
                        'Nombres': row.get('Name', '').split(' ')[0] if pd.notna(row.get('Name', '')) else '',
                        'Apellidos': ' '.join(row.get('Name', '').split(' ')[1:]) if pd.notna(row.get('Name', '')) and len(row.get('Name', '').split(' ')) > 1 else '',
                        'Facultad': '',
                        'Escuela': '',
                        'Correo': row.get('Email', '')
                    }
                    df_nuevo = pd.concat([df_nuevo, pd.DataFrame([nueva_fila])], ignore_index=True)
                
                self.df = df_nuevo
                self.log("Datos migrados a nueva estructura universitaria")
            else:
                # Crear estructura completamente nueva
                self.crear_dataframe_universitario()
        except Exception as e:
            self.log(f"Error en migración: {str(e)}")
            self.crear_dataframe_universitario()
    
    def crear_dataframe_universitario(self):
        """Crea un DataFrame vacío con la estructura universitaria"""
        self.df = pd.DataFrame(columns=self.columnas_requeridas)
        os.makedirs("data", exist_ok=True)
        self.df.to_excel("data/estudiantes.xlsx", index=False)
    
    def cargar_plantilla(self):
        """Carga la plantilla HTML"""
        try:
            ruta_plantilla = "templates/email_template.html"
            if os.path.exists(ruta_plantilla):
                with codecs.open(ruta_plantilla, 'r', encoding='utf-8') as file:
                    self.plantilla_html = file.read()
                self.log("Plantilla HTML cargada exitosamente")
            else:
                self.log("Plantilla HTML no encontrada. Se usará formato de texto plano.")
        except Exception as e:
            self.log(f"Error al cargar plantilla: {str(e)}")
    
    def cargar_configuracion(self):
        """Carga la configuración personalizada del usuario desde archivo"""
        try:
            config_file = "config/user_config.json"
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                
                # Cargar valores desde el archivo de configuración
                self.name_account = config_data.get('name_account', self.name_account)
                self.email_account = config_data.get('email_account', self.email_account)
                self.password_account = config_data.get('password_account', self.password_account)
                self.asunto_default = config_data.get('asunto_default', self.asunto_default)
                self.mensaje_default = config_data.get('mensaje_default', self.mensaje_default)
                
                print("Configuración personalizada cargada desde archivo")
                # Actualizar campos en la interfaz si existen
                if hasattr(self, 'config_nombre'):
                    self.config_nombre.delete(0, tk.END)
                    self.config_nombre.insert(0, self.name_account)
                
                if hasattr(self, 'config_email'):
                    self.config_email.delete(0, tk.END)
                    self.config_email.insert(0, self.email_account)
                
                if hasattr(self, 'config_password'):
                    self.config_password.delete(0, tk.END)
                    self.config_password.insert(0, self.password_account)
                
                if hasattr(self, 'entry_asunto_default'):
                    self.entry_asunto_default.delete(0, tk.END)
                    self.entry_asunto_default.insert(0, self.asunto_default)
                
                if hasattr(self, 'text_mensaje_default'):
                    self.text_mensaje_default.delete('1.0', tk.END)
                    self.text_mensaje_default.insert('1.0', self.mensaje_default)
                
            else:
                print("No se encontró archivo de configuración. Se utilizarán valores predeterminados.")
        except Exception as e:
            print(f"Error al cargar configuración: {str(e)}")
            # En caso de error, intentar cargar desde .env (compatibilidad)
            load_dotenv()
    
    def actualizar_preview(self):
        """Actualiza la vista previa de estudiantes"""
        if self.df is None or self.df.empty:
            # Actualizar etiqueta con total de registros
            self.lbl_total_registros.config(text="Total: 0 registros")
            return
        
        # Actualizar etiqueta con total de registros
        total_registros = len(self.df) if self.df is not None else 0
        self.lbl_total_registros.config(text=f"Total: {total_registros} registros")
        
        # Limpiar tree
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Obtener rango
        inicio = max(1, self.rango_inicio.get()) - 1
        fin = min(len(self.df), self.rango_fin.get())
        
        # Agregar datos al tree en el orden especificado
        for i in range(inicio, fin):
            row = self.df.iloc[i]
            
            # Obtener valores con valores por defecto si están vacíos
            nombres = row.get('Nombres', '') if pd.notna(row.get('Nombres', '')) else ""
            apellidos = row.get('Apellidos', '') if pd.notna(row.get('Apellidos', '')) else ""
            correo = row.get('Correo', '') if pd.notna(row.get('Correo', '')) else ""
            facultad = row.get('Facultad', '') if pd.notna(row.get('Facultad', '')) else ""
            escuela = row.get('Escuela', '') if pd.notna(row.get('Escuela', '')) else ""
            codigo = row.get('Cod.Universitario', '') if pd.notna(row.get('Cod.Universitario', '')) else ""
            
            # Truncar texto largo para la vista de tabla
            if len(facultad) > 25:
                facultad = facultad[:25] + "..."
            if len(escuela) > 25:
                escuela = escuela[:25] + "..."
            
            # Número de fila (índice + 1)
            numero_fila = str(i + 1)
            
            # Insertar en el orden especificado: #, Nombres, Apellidos, Correo, Facultad, Escuela, Cod.Universitario
            valores = [numero_fila, nombres, apellidos, correo, facultad, escuela, codigo]
            self.tree.insert('', 'end', values=valores)
    
    def mostrar_detalle(self, event):
        """Muestra el detalle del estudiante seleccionado"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = self.tree.item(selection[0])
        valores = item['values']
        
        if self.df is not None and len(valores) > 0:
            # Buscar la fila correspondiente por los valores mostrados
            # Ajuste para tener en cuenta la nueva columna "#"
            if len(valores) > 3:  # Verificamos si hay suficientes columnas (por la estructura actual)
                nombres = valores[1]  # Ajuste por la columna "#"
                apellidos = valores[2]
                correo = valores[3]
            else:
                nombres = valores[0]
                apellidos = valores[1]
                correo = valores[2]
            
            # Buscar en el DataFrame
            fila_encontrada = self.df[
                (self.df['Nombres'] == nombres) & 
                (self.df['Apellidos'] == apellidos) & 
                (self.df['Correo'] == correo)
            ]
            
            if not fila_encontrada.empty:
                row = fila_encontrada.iloc[0]
                
                # Guardar la última selección
                self.ultima_seleccion = row
                
                # Llenar campos de edición
                self.entry_nombres.delete(0, tk.END)
                self.entry_nombres.insert(0, row.get('Nombres', ''))
                
                self.entry_apellidos.delete(0, tk.END)
                self.entry_apellidos.insert(0, row.get('Apellidos', ''))
                
                self.entry_correo.delete(0, tk.END)
                self.entry_correo.insert(0, row.get('Correo', ''))
                
                self.entry_codigo.delete(0, tk.END)
                self.entry_codigo.insert(0, row.get('Cod.Universitario', ''))
                
                self.entry_facultad.delete(0, tk.END)
                self.entry_facultad.insert(0, row.get('Facultad', ''))
                
                self.entry_escuela.delete(0, tk.END)
                self.entry_escuela.insert(0, row.get('Escuela', ''))

    def agregar_estudiante(self):
        """Agrega un nuevo estudiante a la lista"""
        nombres = self.entry_nombres.get().strip()
        apellidos = self.entry_apellidos.get().strip()
        correo = self.entry_correo.get().strip()
        codigo = self.entry_codigo.get().strip()
        facultad = self.entry_facultad.get().strip()
        escuela = self.entry_escuela.get().strip()
        
        if not correo:
            messagebox.showwarning("Advertencia", "El campo Correo es obligatorio")
            return
        
        if not nombres:
            messagebox.showwarning("Advertencia", "El campo Nombres es obligatorio")
            return
        
        # Crear nuevo registro
        nuevo_id = len(self.df) + 1 if not self.df.empty else 1
        new_row = pd.DataFrame([{
            'id': nuevo_id,
            'Cod.Universitario': codigo,
            'Nombres': nombres,
            'Apellidos': apellidos,
            'Facultad': facultad,
            'Escuela': escuela,
            'Correo': correo
        }])
        
        self.df = pd.concat([self.df, new_row], ignore_index=True)
        self.actualizar_preview()
        self.limpiar_campos_estudiante()
        
        # Guardar automáticamente en Excel
        self.guardar_excel_silencioso()
        self.log(f"Estudiante agregado y guardado: {nombres} {apellidos}")

    def actualizar_estudiante(self):
        """Actualiza el estudiante seleccionado"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Advertencia", "Selecciona un estudiante para actualizar")
            return
        
        item = self.tree.item(selection[0])
        valores = item['values']
        
        if valores:
            # Buscar la fila correspondiente
            nombres_original = valores[0]
            apellidos_original = valores[1]
            correo_original = valores[2]
            
            fila_index = self.df[
                (self.df['Nombres'] == nombres_original) & 
                (self.df['Apellidos'] == apellidos_original) & 
                (self.df['Correo'] == correo_original)
            ].index
            
            if not fila_index.empty:
                index = fila_index[0]
                
                # Actualizar los datos
                self.df.at[index, 'Nombres'] = self.entry_nombres.get().strip()
                self.df.at[index, 'Apellidos'] = self.entry_apellidos.get().strip()
                self.df.at[index, 'Correo'] = self.entry_correo.get().strip()
                self.df.at[index, 'Cod.Universitario'] = self.entry_codigo.get().strip()
                self.df.at[index, 'Facultad'] = self.entry_facultad.get().strip()
                self.df.at[index, 'Escuela'] = self.entry_escuela.get().strip()
                
                self.actualizar_preview()
                
                # Guardar automáticamente en Excel
                self.guardar_excel_silencioso()
                self.log(f"Estudiante actualizado y guardado en posición {index + 1}")

    def eliminar_estudiante(self):
        """Elimina el estudiante seleccionado"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Advertencia", "Selecciona un estudiante para eliminar")
            return
        
        if messagebox.askyesno("Confirmar", "¿Estás seguro de eliminar este estudiante?"):
            item = self.tree.item(selection[0])
            valores = item['values']
            
            if valores:
                # Buscar la fila correspondiente
                nombres = valores[0]
                apellidos = valores[1]
                correo = valores[2]
                
                fila_index = self.df[
                    (self.df['Nombres'] == nombres) & 
                    (self.df['Apellidos'] == apellidos) & 
                    (self.df['Correo'] == correo)
                ].index
                
                if not fila_index.empty:
                    index = fila_index[0]
                    self.df = self.df.drop(index).reset_index(drop=True)
                    
                    # Reajustar los IDs
                    self.df['id'] = range(1, len(self.df) + 1)
                    
                    self.actualizar_preview()
                    self.limpiar_campos_estudiante()
                    
                    # Guardar automáticamente en Excel
                    self.guardar_excel_silencioso()
                    self.log(f"Estudiante eliminado: {nombres} {apellidos}")

    def cargar_excel_personalizado(self):
        """Permite al usuario cargar un archivo Excel o CSV personalizado"""
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo Excel o CSV",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("Excel files", "*.xls"), 
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        if filename:
            try:
                # Determinar el tipo de archivo y cargarlo apropiadamente
                if filename.lower().endswith('.csv'):
                    # Intentar diferentes codificaciones y separadores para archivos CSV
                    encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
                    separators = [',', ';', '\t', '|']
                    df_temp = None
                    
                    # Leer primeras líneas para detectar el separador
                    with open(filename, 'r', errors='ignore') as f:
                        primera_linea = f.readline().strip()
                    
                    # Intentar detectar el separador automáticamente
                    detected_sep = None
                    for sep in separators:
                        if sep in primera_linea:
                            # Contar ocurrencias de cada separador
                            if primera_linea.count(sep) > 2:  # Si hay múltiples ocurrencias, probablemente es el separador
                                detected_sep = sep
                                self.log(f"Separador detectado automáticamente: '{sep}'")
                                break
                    
                    # Si no se detecta, usar coma por defecto
                    if detected_sep is None:
                        detected_sep = ','
                    
                    # Probar diferentes codificaciones con el separador detectado
                    for encoding in encodings:
                        try:
                            df_temp = pd.read_csv(filename, encoding=encoding, sep=detected_sep)
                            self.log(f"Archivo CSV cargado con codificación: {encoding} y separador: '{detected_sep}'")
                            break
                        except UnicodeDecodeError:
                            continue
                        except Exception as e:
                            self.log(f"Error con codificación {encoding}: {str(e)}")
                            continue
                    
                    # Si el separador detectado falla, probar otros separadores
                    if df_temp is None:
                        for sep in separators:
                            if sep == detected_sep:
                                continue  # Ya probamos este
                            for encoding in encodings:
                                try:
                                    df_temp = pd.read_csv(filename, encoding=encoding, sep=sep)
                                    self.log(f"Archivo CSV cargado con separador alternativo: '{sep}' y codificación: {encoding}")
                                    break
                                except:
                                    continue
                            if df_temp is not None:
                                break
                    
                    if df_temp is None:
                        # Preguntar al usuario qué separador usar
                        separador_manual = self.preguntar_separador_csv()
                        if separador_manual:
                            try:
                                df_temp = pd.read_csv(filename, sep=separador_manual, encoding='utf-8')
                                self.log(f"Archivo CSV cargado con separador manual: '{separador_manual}'")
                            except Exception as e:
                                raise Exception(f"No se pudo leer el archivo CSV con el separador '{separador_manual}': {str(e)}")
                        else:
                            raise Exception("No se pudo leer el archivo CSV con ninguna combinación de codificación y separador")
                        
                else:
                    # Cargar archivo Excel
                    df_temp = pd.read_excel(filename)
                
                # Verificar si tiene las columnas requeridas
                columnas_faltantes = [col for col in self.columnas_requeridas if col not in df_temp.columns]
                
                if columnas_faltantes:
                    # Mostrar qué columnas faltan
                    mensaje = f"El archivo no tiene las columnas requeridas:\nFaltantes: {', '.join(columnas_faltantes)}\nRequeridas: {', '.join(self.columnas_requeridas)}"
                    messagebox.showerror("Error de estructura", mensaje)
                    return
                
                # Si tiene todas las columnas, cargar los datos
                self.df = df_temp
                
                # Asegurar que el id sea numérico y único
                if 'id' in self.df.columns:
                    self.df['id'] = range(1, len(self.df) + 1)
                
                # Actualizar vista y contador
                self.actualizar_preview()
                
                # Guardar en el archivo local (siempre como Excel)
                os.makedirs("data", exist_ok=True)
                self.df.to_excel("data/estudiantes.xlsx", index=False)
                
                tipo_archivo = "CSV" if filename.lower().endswith('.csv') else "Excel"
                self.log(f"Archivo {tipo_archivo} cargado exitosamente: {filename}")
                self.log(f"Se cargaron {len(self.df)} estudiantes")
                
                # Actualizar el total de registros
                self.lbl_total_registros.config(text=f"Total: {len(self.df)} registros")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al cargar archivo: {str(e)}")
                self.log(f"Error al cargar archivo: {str(e)}")
    
    def crear_excel_ejemplo(self):
        """Crea archivos Excel y CSV de ejemplo con la estructura correcta"""
        try:
            # Datos de ejemplo
            datos_ejemplo = [
                {
                    'id': 1,
                    'Cod.Universitario': '2020123456',
                    'Nombres': 'Juan Carlos',
                    'Apellidos': 'García López',
                    'Facultad': 'Ingeniería de Sistemas',
                    'Escuela': 'Sistemas e Informática',
                    'Correo': 'juan.garcia@universidad.edu.pe'
                },
                {
                    'id': 2,
                    'Cod.Universitario': '2021654321',
                    'Nombres': 'María Elena',
                    'Apellidos': 'Rodríguez Martínez',
                    'Facultad': 'Ciencias Económicas',
                    'Escuela': 'Administración',
                    'Correo': 'maria.rodriguez@universidad.edu.pe'
                },
                {
                    'id': 3,
                    'Cod.Universitario': '2019987654',
                    'Nombres': 'Carlos Alberto',
                    'Apellidos': 'Vásquez Torres',
                    'Facultad': 'Derecho y Ciencias Políticas',
                    'Escuela': 'Derecho',
                    'Correo': 'carlos.vasquez@universidad.edu.pe'
                }
            ]
            
            df_ejemplo = pd.DataFrame(datos_ejemplo)
            
            os.makedirs("data", exist_ok=True)
            
            # Crear archivo Excel
            archivo_excel = "data/ejemplo_estudiantes.xlsx"
            df_ejemplo.to_excel(archivo_excel, index=False)
            
            # Crear archivos CSV con diferentes separadores
            archivo_csv = "data/ejemplo_estudiantes.csv"
            df_ejemplo.to_csv(archivo_csv, index=False, encoding='utf-8')
            
            # Crear CSV con punto y coma (común en países hispanos)
            archivo_csv_semicolon = "data/ejemplo_estudiantes_semicolon.csv"
            df_ejemplo.to_csv(archivo_csv_semicolon, index=False, encoding='utf-8', sep=';')
            
            messagebox.showinfo("Éxito", f"Archivos de ejemplo creados:\n- {archivo_excel}\n- {archivo_csv}\n- {archivo_csv_semicolon} (con punto y coma)")
            self.log(f"Archivos de ejemplo creados: {archivo_excel}, {archivo_csv}, {archivo_csv_semicolon}")
            
            # Preguntar qué archivo quiere abrir
            dialog = tk.Toplevel(self.root)
            dialog.title("Abrir archivo de ejemplo")
            dialog.geometry("350x200")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()
            
            # Centrar la ventana
            dialog.geometry("+%d+%d" % (
                self.root.winfo_rootx() + self.root.winfo_width() // 2 - 175,
                self.root.winfo_rooty() + self.root.winfo_height() // 2 - 100
            ))
            
            ttk.Label(dialog, text="¿Desea abrir alguno de los archivos de ejemplo?", 
                    wraplength=330).pack(pady=10, padx=10)
            
            resultado = [None]
            
            def abrir_excel():
                resultado[0] = archivo_excel
                dialog.destroy()
            
            def abrir_csv():
                resultado[0] = archivo_csv
                dialog.destroy()
                
            def abrir_csv_semicolon():
                resultado[0] = archivo_csv_semicolon
                dialog.destroy()
                
            def no_abrir():
                resultado[0] = None
                dialog.destroy()
            
            ttk.Button(dialog, text="Abrir Excel (.xlsx)", command=abrir_excel).pack(fill="x", padx=20, pady=5)
            ttk.Button(dialog, text="Abrir CSV con comas (.csv)", command=abrir_csv).pack(fill="x", padx=20, pady=5)
            ttk.Button(dialog, text="Abrir CSV con punto y coma (.csv)", command=abrir_csv_semicolon).pack(fill="x", padx=20, pady=5)
            ttk.Button(dialog, text="No abrir ninguno", command=no_abrir).pack(fill="x", padx=20, pady=5)
            
            # Esperar a que se cierre el diálogo
            self.root.wait_window(dialog)
            
            if resultado[0]:
                self.abrir_archivo_excel(resultado[0])
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al crear archivo de ejemplo: {str(e)}")
            self.log(f"Error al crear archivo de ejemplo: {str(e)}")
    
    def abrir_archivo_excel(self, ruta_archivo):
        """Abre un archivo Excel específico"""
        try:
            ruta_absoluta = os.path.abspath(ruta_archivo)
            
            if platform.system() == 'Windows':
                os.startfile(ruta_absoluta)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.run(['open', ruta_absoluta])
            else:  # Linux
                subprocess.run(['xdg-open', ruta_absoluta])
                
        except Exception as e:
            self.log(f"Error al abrir archivo: {str(e)}")

    def enviar_correos(self):
        """Envía los correos en un hilo separado"""
        if self.df is None or self.df.empty:
            messagebox.showwarning("Advertencia", "No hay correos para enviar")
            return
        
        self.btn_enviar.config(state='disabled')
        thread = threading.Thread(target=self._enviar_correos_thread)
        thread.daemon = True
        thread.start()
    
    def _enviar_correos_thread(self):
        """Hilo para envío de correos"""
        try:
            inicio = max(1, int(self.send_inicio.get())) - 1
            fin = min(len(self.df), int(self.send_fin.get()))
            
            if inicio >= fin:
                self.log("Rango de envío inválido")
                return
            
            total = fin - inicio
            self.progress['maximum'] = total
            self.progress['value'] = 0
            
            # Conectar al servidor
            server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            server.ehlo()
            server.login(self.config_email.get(), self.config_password.get())
            
            enviados = 0
            errores = 0
            
            for i in range(inicio, fin):
                try:
                    row = self.df.iloc[i]
                    self._enviar_correo_individual(server, row, i + 1)
                    enviados += 1
                    self.log(f"✅ Correo {i+1} enviado exitosamente")
                except Exception as e:
                    errores += 1
                    self.log(f"❌ Error en correo {i+1}: {str(e)}")
                
                self.progress['value'] = i - inicio + 1
                self.root.update_idletasks()
            
            server.close()
            self.log(f"Envío completado. Enviados: {enviados}, Errores: {errores}")
            
        except Exception as e:
            self.log(f"Error de conexión: {str(e)}")
        finally:
            self.btn_enviar.config(state='normal')
    
    def _enviar_correo_individual(self, server, row, index):
        """Envía un correo individual"""
        try:
            # Obtener datos con valores por defecto
            nombres = row.get('Nombres', '') if pd.notna(row.get('Nombres', '')) else "Estimado/a"
            apellidos = row.get('Apellidos', '') if pd.notna(row.get('Apellidos', '')) else "estudiante"
            nombre_completo = f"{nombres} {apellidos}".strip()
            email = row.get('Correo', '') if pd.notna(row.get('Correo', '')) else row.get('Email', '')
            facultad = row.get('Facultad', '') if pd.notna(row.get('Facultad', '')) else ""
            escuela = row.get('Escuela', '') if pd.notna(row.get('Escuela', '')) else ""
            codigo = row.get('Cod.Universitario', '') if pd.notna(row.get('Cod.Universitario', '')) else ""
            
            # Verificar que tengamos un email válido
            if not email:
                raise ValueError("Dirección de correo vacía")
            
            # Obtener el asunto y mensaje de la configuración
            # Primero intentamos usar las variables de instancia (que se actualizan al guardar la configuración)
            # Si no están disponibles, leemos directamente de los controles de la UI
            asunto_default = self.asunto_default
            mensaje_default = self.mensaje_default
            
            # Como respaldo, si las variables de instancia están vacías, usar los valores de los widgets
            if (asunto_default is None or asunto_default == "") and hasattr(self, 'entry_asunto_default'):
                asunto_default = self.entry_asunto_default.get()
                self.log("DEBUG - Usando asunto desde widget porque el valor en instancia estaba vacío")
                
            if (mensaje_default is None or mensaje_default == "") and hasattr(self, 'text_mensaje_default'):
                mensaje_default = self.text_mensaje_default.get('1.0', 'end-1c')
                self.log("DEBUG - Usando mensaje desde widget porque el valor en instancia estaba vacío")
            
            # Usar valores predeterminados configurados o los del DataFrame
            asunto = row.get('Asunto', '') if pd.notna(row.get('Asunto', '')) else asunto_default
            mensaje = row.get('Mensaje', '') if pd.notna(row.get('Mensaje', '')) else mensaje_default
            
            # Asegurarse de que tenemos un asunto y mensaje
            if not asunto:
                asunto = "Información Universitaria"  # Valor por defecto
            
            if not mensaje:
                mensaje = "Este es un mensaje personalizado para recordarte sobre los trámites universitarios pendientes."
            
            # Registrar lo que estamos enviando para diagnosticar
            self.log(f"DEBUG - Asunto: {asunto}")
            self.log(f"DEBUG - Email destino: {email}")
            self.log(f"DEBUG - Mensaje (primeros 30 caracteres): {mensaje[:30]}...")
            
            # Personalizar asunto
            if not asunto.endswith(f", {nombre_completo}!"):
                asunto = f"{asunto}, {nombre_completo}!"
            
            # Crear mensaje de correo
            if self.usar_html.get() and self.plantilla_html:
                # Mensaje multipart con HTML
                msg = MIMEMultipart('alternative')
                
                # Encabezados principales
                msg['From'] = f"{self.config_nombre.get()} <{self.config_email.get()}>"
                msg['To'] = email
                msg['Subject'] = asunto
                
                # Texto plano como fallback
                texto_plano = f"Hola, {nombre_completo}!\n\n{mensaje}\n\nAtentamente,\n{self.config_nombre.get()}"
                part1 = MIMEText(texto_plano, 'plain', 'utf-8')
                
                # Parte HTML
                try:
                    # Intentar usar la plantilla universitaria primero
                    mensaje_html = self.personalizar_plantilla_universitaria(
                        self.plantilla_html,
                        nombres,
                        apellidos,
                        facultad,
                        escuela,
                        codigo,
                        mensaje,
                        self.config_nombre.get(),
                        asunto.replace(f', {nombre_completo}!', '')
                    )
                except Exception as e:
                    self.log(f"Error al personalizar plantilla universitaria: {str(e)}, usando plantilla simple")
                    # Si falla, usar la plantilla simple
                    mensaje_html = self.personalizar_plantilla(
                        self.plantilla_html,
                        nombre_completo,
                        mensaje,
                        self.config_nombre.get(),
                        asunto.replace(f', {nombre_completo}!', '')
                    )
                
                part2 = MIMEText(mensaje_html, 'html', 'utf-8')
                
                # Añadir ambas partes
                msg.attach(part1)
                msg.attach(part2)
            else:
                # Mensaje simple en texto plano
                msg = MIMEMultipart()
                
                # Encabezados principales
                msg['From'] = f"{self.config_nombre.get()} <{self.config_email.get()}>"
                msg['To'] = email
                msg['Subject'] = asunto
                
                # Mensaje en texto plano
                message = f"Hola, {nombre_completo}!\n\n{mensaje}\n\nAtentamente,\n{self.config_nombre.get()}"
                part = MIMEText(message, 'plain', 'utf-8')
                msg.attach(part)
            
            # Enviar el correo
            sent_email = msg.as_string()
            server.sendmail(self.config_email.get(), [email], sent_email)
            
            # Loguear éxito
            self.log(f"Correo enviado exitosamente a {email}")
            
        except Exception as e:
            # Registrar el error específico
            self.log(f"Error al enviar correo a {email}: {str(e)}")
            raise  # Re-lanzar la excepción para que se maneje en el método llamante
    
    def envio_prueba(self):
        """Envía un correo de prueba al primer email de la lista"""
        if self.df is None or self.df.empty:
            messagebox.showwarning("Advertencia", "No hay correos en la lista")
            return
        
        # Verificar que los campos de configuración no estén vacíos
        if not self.config_email.get().strip():
            messagebox.showerror("Error", "El correo del remitente está vacío. Configure su cuenta en la pestaña 'Configuración'.")
            return
            
        if not self.config_password.get().strip():
            messagebox.showerror("Error", "La contraseña de aplicación está vacía. Configure su cuenta en la pestaña 'Configuración'.")
            return
        
        # Verificar que tengamos un asunto y mensaje
        asunto = self.entry_asunto_default.get().strip()
        mensaje = self.text_mensaje_default.get('1.0', 'end-1c').strip()
        
        if not asunto:
            messagebox.showwarning("Advertencia", "El asunto está vacío. Se usará un valor predeterminado.")
            
        if not mensaje:
            messagebox.showwarning("Advertencia", "El mensaje está vacío. Se usará un valor predeterminado.")
        
        self.log("Iniciando envío de prueba...")
        
        try:
            row = self.df.iloc[0]
            email_destino = row.get('Correo', '') 
            self.log(f"Enviando prueba a: {email_destino}")
            
            # Imprimir información de depuración
            self.log(f"Servidor: smtp.gmail.com:465")
            self.log(f"Cuenta: {self.config_email.get()}")
            self.log(f"Asunto: {asunto}")
            
            # Conectar al servidor
            server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            server.ehlo()
            self.log("Conexión establecida con el servidor")
            
            # Iniciar sesión
            server.login(self.config_email.get(), self.config_password.get())
            self.log("Inicio de sesión exitoso")
            
            # Enviar el correo
            self._enviar_correo_individual(server, row, 1)
            server.close()
            
            messagebox.showinfo("Éxito", f"Correo de prueba enviado exitosamente a {email_destino}")
            self.log(f"✅ Correo de prueba enviado exitosamente a {email_destino}")
        except Exception as e:
            error_detalle = str(e)
            messagebox.showerror("Error", f"Error en envío de prueba: {error_detalle}")
            self.log(f"❌ Error en envío de prueba: {error_detalle}")
            
            # Sugerencias basadas en errores comunes
            if "getaddrinfo failed" in error_detalle or "Connection refused" in error_detalle:
                self.log("Sugerencia: Revise su conexión a Internet o pruebe más tarde")
            elif "Authentication" in error_detalle or "Username and Password" in error_detalle:
                self.log("Sugerencia: Verifique su correo y contraseña de aplicación")
            elif "550" in error_detalle or "553" in error_detalle:
                self.log("Sugerencia: El servidor rechazó su correo. Verifique la dirección de destino")
            elif "530" in error_detalle:
                self.log("Sugerencia: Requiere autenticación. Asegúrese de usar una contraseña de aplicación válida")
            elif "smtplib.SMTPDataError" in error_detalle:
                self.log("Sugerencia: Problema con el contenido del mensaje. Intente simplificarlo")
    
    def probar_conexion(self):
        """Prueba la conexión con el servidor de correo"""
        try:
            server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            server.ehlo()
            server.login(self.config_email.get(), self.config_password.get())
            server.close()
            
            messagebox.showinfo("Éxito", "Conexión exitosa con el servidor de correo")
            self.log("Conexión exitosa con el servidor")
        except Exception as e:
            messagebox.showerror("Error", f"Error de conexión: {str(e)}")
            self.log(f"Error de conexión: {str(e)}")
    
    def seleccionar_plantilla(self):
        """Selecciona un archivo de plantilla HTML"""
        filename = filedialog.askopenfilename(
            title="Seleccionar plantilla HTML",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")]
        )
        if filename:
            try:
                with codecs.open(filename, 'r', encoding='utf-8') as file:
                    self.plantilla_html = file.read()
                self.plantilla_path.config(text=f"Plantilla actual: {filename}")
                self.log(f"Plantilla cargada: {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Error al cargar plantilla: {str(e)}")
                self.log(f"Error al cargar plantilla: {str(e)}")
    
    def preguntar_separador_csv(self):
        """Muestra un diálogo para que el usuario elija el separador del archivo CSV"""
        separadores = {
            "Punto y coma (;)": ";",
            "Coma (,)": ",",
            "Tabulación (\\t)": "\t",
            "Pipe (|)": "|",
            "Espacio ( )": " "
        }
        
        # Crear una ventana de diálogo personalizada
        dialog = tk.Toplevel(self.root)
        dialog.title("Seleccionar Separador CSV")
        dialog.geometry("400x250")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Centrar la ventana
        dialog.geometry("+%d+%d" % (
            self.root.winfo_rootx() + self.root.winfo_width() // 2 - 200,
            self.root.winfo_rooty() + self.root.winfo_height() // 2 - 125
        ))
        
        # Variable para almacenar la selección
        selected_sep = tk.StringVar(value=";")  # Por defecto punto y coma
        
        ttk.Label(dialog, text="No se pudo detectar automáticamente el separador del archivo CSV.", 
                 wraplength=380).pack(pady=10, padx=10)
        ttk.Label(dialog, text="Por favor, seleccione el separador utilizado en su archivo:", 
                 wraplength=380).pack(pady=5, padx=10)
        
        # Crear radio buttons para cada separador
        for i, (texto, sep) in enumerate(separadores.items()):
            ttk.Radiobutton(dialog, text=texto, value=sep, variable=selected_sep).pack(anchor="w", padx=20, pady=5)
        
        # Campo para separador personalizado
        custom_frame = ttk.Frame(dialog)
        custom_frame.pack(fill="x", padx=20, pady=5)
        custom_sep = tk.StringVar()
        ttk.Radiobutton(custom_frame, text="Otro:", value="custom", variable=selected_sep).pack(side="left")
        ttk.Entry(custom_frame, textvariable=custom_sep, width=5).pack(side="left", padx=5)
        
        # Resultado y botones
        result = [None]  # Usar lista para poder modificar desde el callback
        
        def on_ok():
            if selected_sep.get() == "custom" and custom_sep.get():
                result[0] = custom_sep.get()
            else:
                result[0] = selected_sep.get()
            dialog.destroy()
        
        def on_cancel():
            result[0] = None
            dialog.destroy()
        
        # Botones
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill="x", pady=15)
        ttk.Button(button_frame, text="Aceptar", command=on_ok).pack(side="right", padx=10)
        ttk.Button(button_frame, text="Cancelar", command=on_cancel).pack(side="right", padx=10)
        
        # Esperar a que se cierre el diálogo
        self.root.wait_window(dialog)
        
        return result[0]
    
    def log(self, mensaje):
        """Agrega un mensaje al log"""
        import pandas as pd  # Importación local para evitar problemas
        hora_actual = pd.Timestamp.now().strftime('%H:%M:%S')
        
        # Verificar si el widget log_text existe
        if hasattr(self, 'log_text') and self.log_text:
            self.log_text.insert(tk.END, f"{hora_actual} - {mensaje}\n")
            self.log_text.see(tk.END)
        else:
            # Si no existe, imprimir en la consola
            print(f"{hora_actual} - {mensaje}")
        self.root.update_idletasks()

    def guardar_excel(self):
        """Guarda los cambios en el archivo Excel universitario"""
        try:
            os.makedirs("data", exist_ok=True)
            self.df.to_excel("data/estudiantes.xlsx", index=False)
            messagebox.showinfo("Éxito", "Archivo guardado exitosamente")
            self.log("Archivo Excel guardado")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar: {str(e)}")
            self.log(f"Error al guardar: {str(e)}")
    
    def guardar_excel_silencioso(self):
        """Guarda los cambios en el archivo Excel sin mostrar mensajes de confirmación"""
        try:
            os.makedirs("data", exist_ok=True)
            self.df.to_excel("data/estudiantes.xlsx", index=False)
            # Solo loguear, no mostrar messagebox
        except Exception as e:
            self.log(f"Error al guardar automáticamente: {str(e)}")
    
    def abrir_excel(self):
        """Abre el archivo Excel con la aplicación predeterminada del sistema"""
        try:
            # Usar ruta absoluta para evitar problemas en Windows
            excel_path = os.path.abspath("data/estudiantes.xlsx")
            
            # Verificar si el archivo existe
            if not os.path.exists(excel_path):
                # Si no existe, crearlo primero
                os.makedirs(os.path.dirname(excel_path), exist_ok=True)
                if self.df is not None:
                    self.df.to_excel(excel_path, index=False)
                else:
                    # Crear un DataFrame vacío si no hay datos
                    empty_df = pd.DataFrame(columns=self.columnas_requeridas)
                    empty_df.to_excel(excel_path, index=False)
                self.log(f"Archivo Excel creado en: {excel_path}")
            
            # Abrir el archivo con la aplicación predeterminada
            if platform.system() == 'Windows':
                # En Windows, usar la ruta absoluta
                os.startfile(excel_path)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.run(['open', excel_path])
            else:  # Linux
                subprocess.run(['xdg-open', excel_path])
            
            self.log(f"Archivo Excel abierto: {excel_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al abrir Excel: {str(e)}")
            self.log(f"Error al abrir Excel: {str(e)}")
    
    @staticmethod
    def personalizar_plantilla_universitaria(plantilla, nombres, apellidos, facultad, escuela, codigo, mensaje, nombre_remitente, titulo="Información Universitaria"):
        """Personaliza la plantilla HTML con los datos universitarios del destinatario"""
        if plantilla is None:
            return None
        
        # Reemplazar las variables en la plantilla
        plantilla_personalizada = plantilla.replace("{{NOMBRES}}", nombres)
        plantilla_personalizada = plantilla_personalizada.replace("{{APELLIDOS}}", apellidos)
        plantilla_personalizada = plantilla_personalizada.replace("{{NOMBRE_COMPLETO}}", f"{nombres} {apellidos}")
        plantilla_personalizada = plantilla_personalizada.replace("{{FACULTAD}}", facultad)
        plantilla_personalizada = plantilla_personalizada.replace("{{ESCUELA}}", escuela)
        plantilla_personalizada = plantilla_personalizada.replace("{{CODIGO_UNIVERSITARIO}}", codigo)
        plantilla_personalizada = plantilla_personalizada.replace("{{MENSAJE_PRINCIPAL}}", mensaje)
        plantilla_personalizada = plantilla_personalizada.replace("{{NOMBRE_REMITENTE}}", nombre_remitente)
        plantilla_personalizada = plantilla_personalizada.replace("{{TITULO_PRINCIPAL}}", titulo)
        
        # Mantener compatibilidad con plantillas antiguas
        plantilla_personalizada = plantilla_personalizada.replace("{{NOMBRE}}", f"{nombres} {apellidos}")
        
        return plantilla_personalizada
    
    @staticmethod
    def personalizar_plantilla(plantilla, nombre, mensaje, nombre_remitente, titulo="Mensaje Personalizado"):
        """Personaliza la plantilla HTML con los datos del destinatario"""
        if plantilla is None:
            return None
    
        # Reemplazar las variables en la plantilla
        plantilla_personalizada = plantilla.replace("{{NOMBRE}}", nombre)
        plantilla_personalizada = plantilla_personalizada.replace("{{MENSAJE_PRINCIPAL}}", mensaje)
        plantilla_personalizada = plantilla_personalizada.replace("{{NOMBRE_REMITENTE}}", nombre_remitente)
        plantilla_personalizada = plantilla_personalizada.replace("{{TITULO_PRINCIPAL}}", titulo)
        
        return plantilla_personalizada

    def cargar_prueba_bebita(self):
        """Carga directamente el archivo prueba bebita.csv para testing"""
        try:
            filename = "data/prueba bebita.csv"
            
            if os.path.exists(filename):
                # Intentamos con punto y coma que es lo más probable
                df_temp = pd.read_csv(filename, sep=";", encoding='utf-8')
                self.log(f"Archivo prueba bebita.csv cargado correctamente con separador ';'")
                
                # Verificar si tiene las columnas requeridas
                columnas_faltantes = [col for col in self.columnas_requeridas if col not in df_temp.columns]
                
                if columnas_faltantes:
                    # Mostrar qué columnas faltan
                    mensaje = f"El archivo no tiene las columnas requeridas:\nFaltantes: {', '.join(columnas_faltantes)}\nRequeridas: {', '.join(self.columnas_requeridas)}"
                    messagebox.showerror("Error de estructura", mensaje)
                    return
                
                # Si tiene todas las columnas, cargar los datos
                self.df = df_temp
                
                # Asegurar que el id sea numérico y único
                if 'id' in self.df.columns:
                    self.df['id'] = range(1, len(self.df) + 1)
                
                self.actualizar_preview()
                
                # Guardar en el archivo local como Excel
                os.makedirs("data", exist_ok=True)
                self.df.to_excel("data/estudiantes.xlsx", index=False)
                
                messagebox.showinfo("Éxito", "Archivo prueba bebita.csv cargado correctamente")
            else:
                messagebox.showwarning("Archivo no encontrado", "No se encontró el archivo 'data/prueba bebita.csv'")
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar archivo: {str(e)}")
            self.log(f"Error al cargar archivo: {str(e)}")


    def limpiar_seleccion(self):
        """Limpia los campos de edición del estudiante"""
        self.limpiar_campos_estudiante()
        
        # Quitar la selección en el árbol
        for item in self.tree.selection():
            self.tree.selection_remove(item)
        
        self.log("Selección limpiada")
    
    def limpiar_campos_estudiante(self):
        """Limpia los campos de edición del estudiante"""
        self.entry_nombres.delete(0, tk.END)
        self.entry_apellidos.delete(0, tk.END)
        self.entry_correo.delete(0, tk.END)
        self.entry_codigo.delete(0, tk.END)
        self.entry_facultad.delete(0, tk.END)
        self.entry_escuela.delete(0, tk.END)

    def guardar_valores_predeterminados(self):
        """Guarda los valores predeterminados de asunto, mensaje y configuración de correo"""
        try:
            # Registro de valores anteriores para depuración
            asunto_anterior = self.asunto_default
            mensaje_anterior = self.mensaje_default
            
            # Actualizar valores en la instancia desde los widgets
            self.asunto_default = self.entry_asunto_default.get()
            self.mensaje_default = self.text_mensaje_default.get('1.0', 'end-1c')
            self.name_account = self.config_nombre.get()
            self.email_account = self.config_email.get()
            self.password_account = self.config_password.get()
            
            # Log de cambios para depuración
            if asunto_anterior != self.asunto_default:
                self.log(f"Asunto por defecto actualizado: '{asunto_anterior}' → '{self.asunto_default}'")
            
            if mensaje_anterior != self.mensaje_default:
                self.log(f"Mensaje por defecto actualizado: Primeros 30 caracteres: '{mensaje_anterior[:30]}...' → '{self.mensaje_default[:30]}...'")
            
            # Crear diccionario con los datos de configuración
            config_data = {
                'name_account': self.name_account,
                'email_account': self.email_account,
                'password_account': self.password_account,
                'asunto_default': self.asunto_default,
                'mensaje_default': self.mensaje_default
            }
            
            # Crear directorio de configuración si no existe
            os.makedirs("config", exist_ok=True)
            
            # Guardar configuración en archivo JSON
            with open("config/user_config.json", 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=4)
            
            # También actualizar el archivo .env para mantener compatibilidad
            with open(".env", 'w', encoding='utf-8') as f:
                f.write(f'name_account="{self.name_account}"\n')
                f.write(f'email_account="{self.email_account}"\n')
                f.write(f'password_account="{self.password_account}"\n')
            
            messagebox.showinfo("Éxito", "Configuración guardada correctamente")
            self.log("Configuración personalizada guardada en archivo")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar configuración: {str(e)}")
            self.log(f"Error al guardar configuración: {str(e)}")
    
    def mostrar_vista_previa_html(self):
        """Muestra una vista previa de la plantilla HTML con los datos de ejemplo"""
        if not self.plantilla_html:
            messagebox.showwarning("Advertencia", "No hay plantilla HTML cargada")
            return
        
        try:
            # Crear una ventana para mostrar la vista previa
            preview_window = tk.Toplevel(self.root)
            preview_window.title("Vista Previa de Plantilla HTML")
            preview_window.geometry("800x600")
            preview_window.transient(self.root)
            
            # Usar datos de ejemplo
            nombres_ejemplo = "Juan Carlos"
            apellidos_ejemplo = "García López"
            facultad_ejemplo = "Ingeniería de Sistemas"
            escuela_ejemplo = "Sistemas e Informática"
            codigo_ejemplo = "2020123456"
            asunto_ejemplo = self.entry_asunto_default.get()
            mensaje_ejemplo = self.text_mensaje_default.get('1.0', 'end-1c')
            nombre_remitente = self.config_nombre.get()
            
            # Personalizar plantilla
            try:
                # Intentar con la función universitaria primero
                html_personalizado = self.personalizar_plantilla_universitaria(
                    self.plantilla_html,
                    nombres_ejemplo,
                    apellidos_ejemplo,
                    facultad_ejemplo,
                    escuela_ejemplo,
                    codigo_ejemplo,
                    mensaje_ejemplo,
                    nombre_remitente,
                    asunto_ejemplo
                )
            except:
                # Si falla, usar la función simple
                html_personalizado = self.personalizar_plantilla(
                    self.plantilla_html,
                    f"{nombres_ejemplo} {apellidos_ejemplo}",
                    mensaje_ejemplo,
                    nombre_remitente,
                    asunto_ejemplo
                )
            
            # Crear un widget de texto para mostrar el código HTML
            html_frame = ttk.LabelFrame(preview_window, text="Código HTML")
            html_frame.pack(fill='both', expand=True, padx=10, pady=5)
            
            html_text = scrolledtext.ScrolledText(html_frame, wrap='word', height=10)
            html_text.pack(fill='both', expand=True, padx=5, pady=5)
            html_text.insert('1.0', html_personalizado)
            
            # Botón para abrir en el navegador
            ttk.Button(preview_window, text="Abrir en Navegador", 
                     command=lambda: self.abrir_html_en_navegador(html_personalizado)).pack(pady=10)
            
            # Centrar la ventana
            preview_window.update_idletasks()
            width = preview_window.winfo_width()
            height = preview_window.winfo_height()
            x = (preview_window.winfo_screenwidth() // 2) - (width // 2)
            y = (preview_window.winfo_screenheight() // 2) - (height // 2)
            preview_window.geometry(f"{width}x{height}+{x}+{y}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al mostrar vista previa: {str(e)}")
            self.log(f"Error al mostrar vista previa: {str(e)}")
    
    def abrir_html_en_navegador(self, html_content):
        """Abre el HTML en el navegador predeterminado"""
        try:
            # Crear un archivo temporal
            with tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode='w', encoding='utf-8') as f:
                f.write(html_content)
                temp_html_path = f.name
            
            # Abrir el archivo en el navegador
            if platform.system() == 'Windows':
                os.startfile(temp_html_path)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.run(['open', temp_html_path])
            else:  # Linux
                subprocess.run(['xdg-open', temp_html_path])
                
            self.log(f"Vista previa HTML abierta en el navegador")
        except Exception as e:
            messagebox.showerror("Error", f"Error al abrir en navegador: {str(e)}")
            self.log(f"Error al abrir en navegador: {str(e)}")


    def on_tab_changed(self, event):
        """Maneja el evento de cambio de pestaña"""
        # Obtener el índice de la pestaña seleccionada actualmente
        tab_actual = self.notebook.select()
        tab_index = self.notebook.index(tab_actual)
        
        # Verificar si estamos cambiando a la pestaña de Gestión de Datos (índice 1)
        if tab_index == 1 and self.ultima_seleccion is not None:
            # Actualizar los campos de edición con la última selección
            self.entry_nombres.delete(0, tk.END)
            self.entry_nombres.insert(0, self.ultima_seleccion.get('Nombres', ''))
            
            self.entry_apellidos.delete(0, tk.END)
            self.entry_apellidos.insert(0, self.ultima_seleccion.get('Apellidos', ''))
            
            self.entry_correo.delete(0, tk.END)
            self.entry_correo.insert(0, self.ultima_seleccion.get('Correo', ''))
            
            self.entry_codigo.delete(0, tk.END)
            self.entry_codigo.insert(0, self.ultima_seleccion.get('Cod.Universitario', ''))
            
            self.entry_facultad.delete(0, tk.END)
            self.entry_facultad.insert(0, self.ultima_seleccion.get('Facultad', ''))
            
            self.entry_escuela.delete(0, tk.END)
            self.entry_escuela.insert(0, self.ultima_seleccion.get('Escuela', ''))

if __name__ == "__main__":
    root = tk.Tk()
    
    # Configurar el icono para la barra de tareas en Windows
    try:
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
        root.iconbitmap(default=icon_path)
    except:
        pass
    
    app = EmailSenderApp(root)
    root.mainloop()