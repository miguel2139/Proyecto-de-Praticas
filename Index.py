import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, simpledialog
from PIL import Image, ImageTk
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.patches as mpatches

# ===================== CONFIGURACIÓN DE COLORES Y ESTILO INSTITUCIONAL =====================
COLOR_FONDO_APP = "#1A1A2E"      # Azul oscuro/negro de fondo
COLOR_AZUL_RNEC = "#0F3460"      # Azul institucional oficial
COLOR_DORADO = "#E94560"         # Dorado/Rosado de realce corporativo
COLOR_VERDE_EXITO = "#2ECC71"    # Verde brillante para Con Autenticación
COLOR_TEXTO_PASIVO = "#E74C3C"   # Rojo para Sin Autenticación
COLOR_BLANCO = "#FFFFFF"
COLOR_GRIS_TEXTO = "#E0E0E0"

# Nuevos tonos para las filas alternadas de las tablas (Aesthetic)
COLOR_FILA_PAR = "#16162A"
COLOR_FILA_IMPAR = "#1F1F3D"
COLOR_FILA_SELECCION = "#264D73"

FONT_TITULO = ("Segoe UI", 16, "bold")
FONT_SUBTITULO = ("Segoe UI", 12, "bold")
FONT_NORMAL = ("Segoe UI", 10)

# ===================== CLASE PARA BOTONES REDONDEADOS PREMIUM =====================
class RoundedButton(tk.Canvas):
    def __init__(self, parent, text, color, command=None, width=320, height=42):
        super().__init__(parent, width=width, height=height, bg=parent['bg'], highlightthickness=0)
        self.command = command
        self.color = color
        
        radius = 18
        self.rect = self.create_rounded_rect(0, 0, width, height, radius, fill=color)
        self.label = self.create_text(width/2, height/2, text=text, fill="white", font=("Segoe UI", 10, "bold"))
        
        self.bind("<Button-1>", lambda e: self.on_click())
        self.bind("<Enter>", lambda e: self.itemconfig(self.rect, fill=self.lighten_color(color)))
        self.bind("<Leave>", lambda e: self.itemconfig(self.rect, fill=color))

    def create_rounded_rect(self, x1, y1, x2, y2, r, **kwargs):
        points = [x1+r, y1, x1+r, y1, x2-r, y1, x2-r, y1, x2, y1, x2, y1+r, x2, y1+r, x2, y2-r, x2, y2-r, x2, y2, x2-r, y2, x2-r, y2, x1+r, y2, x1+r, y2, x1, y2, x1, y2-r, x1, y2-r, x1, y1+r, x1, y1+r, x1, y1]
        return self.create_polygon(points, **kwargs, smooth=True)

    def on_click(self):
        if self.command: self.command()

    def lighten_color(self, hex_color):
        if hex_color == COLOR_AZUL_RNEC: return "#164B85"
        if hex_color == COLOR_VERDE_EXITO: return "#40D47E"
        if hex_color == COLOR_DORADO: return "#FF5773"
        return hex_color

# ===================== LÓGICA DEL SISTEMA Y UNIFICACIÓN =====================
df_global = None
rutas_base = []
ruta2 = None

def seleccionar_base():
    global rutas_base
    rutas = filedialog.askopenfilenames(title="Seleccionar Bases Oficiales DNI", filetypes=[("Archivos de Excel", "*.xlsx *.xls *.xlsm")])
    if rutas:
        rutas_base = list(rutas)
        lbl_base.config(text=f"✓ {len(rutas_base)} Archivos Base Seleccionados", fg="#2ECC71")

def seleccionar_nombres():
    global ruta2
    ruta = filedialog.askopenfilename(title="Seleccionar Maestro de Operadores", filetypes=[("Archivos de Excel", "*.xlsx *.xls *.xlsm")])
    if ruta:
        ruta2 = ruta
        lbl_nombres.config(text=f"✓ Archivo Operadores Cargado", fg="#2ECC71")

def unificar():
    global df_global
    if not rutas_base or not ruta2:
        messagebox.showwarning("Atención al Usuario", "Por favor, seleccione las bases y el archivo de nombres de operadores.")
        return
    try:
        lista = []
        for r in rutas_base:
            df_temp = pd.read_excel(r)
            df_temp.columns = [c.upper().strip() for c in df_temp.columns]
            lista.append(df_temp)
            
        d1 = pd.concat(lista, ignore_index=True)
        d2 = pd.read_excel(ruta2)
        d2.columns = [c.upper().strip() for c in d2.columns]
        
        df_global = pd.concat([d1, d2], ignore_index=True)
        
        for c in ['SIN AUTENTICACION', 'CON AUTENTICACION', 'CUMPLIMIENTO']:
            if c in df_global.columns: 
                df_global[c] = pd.to_numeric(df_global[c], errors='coerce').fillna(0)
                
        if 'NOMBRE' in df_global.columns and 'CCOPERADOR' in df_global.columns:
            mapa = df_global.dropna(subset=['NOMBRE']).drop_duplicates('CCOPERADOR').set_index('CCOPERADOR')['NOMBRE'].to_dict()
            df_global['NOMBRE'] = df_global['NOMBRE'].fillna(df_global['CCOPERADOR'].map(mapa)).fillna('SIN NOMBRE IDENTIFICADO')

        columnas_filtros = ['DEPARTAMENTO', 'MUNICIPIO']
        for col in columnas_filtros:
            if col in df_global.columns:
                 df_global[col] = df_global[col].fillna('INFORMACIÓN NO DISPONIBLE').astype(str).str.upper().str.strip()

        lbl_estado.config(text=f"✓ PROCESAMIENTO EXITOSO: {len(df_global)} REGISTROS", fg=COLOR_VERDE_EXITO)
        btn_descargar.config(state='normal')
        
        deptos = ['TODOS LOS DEPARTAMENTOS'] + sorted(df_global['DEPARTAMENTO'].unique().tolist())
        combo_depto['values'] = deptos
        combo_depto.set('TODOS LOS DEPARTAMENTOS')
        combo_muni['values'] = ['SELECCIONE DEPARTAMENTO']
        combo_muni.set('SELECCIONE DEPARTAMENTO')
        
        combo_depto_sin['values'] = deptos
        combo_depto_sin.set('TODOS LOS DEPARTAMENTOS')
        combo_muni_sin['values'] = ['TODOS LOS MUNICIPIOS']
        combo_muni_sin.set('TODOS LOS MUNICIPIOS')
        
        actualizar_tabla_sin_autenticacion()
        
        messagebox.showinfo("Proceso Completado", "Los datos oficiales se han integrado y estructurado de forma correcta.")
    except Exception as e:
        messagebox.showerror("Error de Procesamiento", f"Ocurrió un problema estructural al procesar las fuentes:\n{e}")

def descargar_archivo():
    global df_global
    if df_global is None: return
    ruta_guardar = filedialog.asksaveasfilename(
        title="Guardar Informe Unificado",
        defaultextension=".xlsx",
        filetypes=[("Libro de Excel", "*.xlsx")]
    )
    if ruta_guardar:
        df_global.to_excel(ruta_guardar, index=False)
        messagebox.showinfo("Informe Descargado", "El reporte consolidado se guardó con éxito.")

def actualizar_municipios(event=None):
    if df_global is None: return
    depto = combo_depto.get()
    if depto != 'TODOS LOS DEPARTAMENTOS':
        munis = ['TODOS LOS MUNICIPIOS'] + sorted(df_global[df_global['DEPARTAMENTO'] == depto]['MUNICIPIO'].unique().tolist())
        combo_muni['values'] = munis
        combo_muni.set('TODOS LOS MUNICIPIOS')
    else:
        combo_muni['values'] = ['SELECCIONE DEPARTAMENTO']
        combo_muni.set('SELECCIONE DEPARTAMENTO')

def actualizar_municipios_sin(event=None):
    if df_global is None: return
    depto = combo_depto_sin.get()
    if depto != 'TODOS LOS DEPARTAMENTOS':
        munis = ['TODOS LOS MUNICIPIOS'] + sorted(df_global[df_global['DEPARTAMENTO'] == depto]['MUNICIPIO'].unique().tolist())
        combo_muni_sin['values'] = munis
        combo_muni_sin.set('TODOS LOS MUNICIPIOS')
    else:
        combo_muni_sin['values'] = ['TODOS LOS MUNICIPIOS']
        combo_muni_sin.set('TODOS LOS MUNICIPIOS')
    actualizar_tabla_sin_autenticacion()

# ===================== GENERACIÓN DEL GRÁFICO OPTIMIZADO =====================
def generar_grafico():
    if df_global is None:
        messagebox.warning("Atención", "Debe unificar y cargar los datos antes de consultar el Dashboard.")
        return

    df_f = df_global.copy()
    if combo_depto.get() != 'TODOS LOS DEPARTAMENTOS':
        df_f = df_f[df_f['DEPARTAMENTO'] == combo_depto.get()]
    if combo_muni.get() != 'TODOS LOS MUNICIPIOS' and combo_muni.get() != 'SELECCIONE DEPARTAMENTO':
        df_f = df_f[df_f['MUNICIPIO'] == combo_muni.get()]

    if df_f.empty:
        messagebox.showinfo("Sin Resultados", "No hay registros bajo los criterios seleccionados.")
        return

    resumen = df_f.groupby(['NOMBRE', 'CCOPERADOR']).agg({
        'CON AUTENTICACION': 'sum',
        'SIN AUTENTICACION': 'sum',
        'CUMPLIMIENTO': 'mean'
    }).sort_values('CON AUTENTICACION', ascending=True).head(10)

    etiquetas_y = [f"{idx[0]} - {int(idx[1]) if pd.notna(idx[1]) else ''}" for idx in resumen.index]

    for w in frame_grafico.winfo_children(): w.destroy()
    
    fig, ax = plt.subplots(figsize=(10.5, 4.8), dpi=100)
    fig.patch.set_facecolor('#1A1A2E')
    ax.set_facecolor('#1A1A2E')

    ax.barh(etiquetas_y, resumen['CON AUTENTICACION'], color=COLOR_VERDE_EXITO)
    ax.barh(etiquetas_y, resumen['SIN AUTENTICACION'], left=resumen['CON AUTENTICACION'], color=COLOR_TEXTO_PASIVO)

    limite_ancho_grafica = resumen['CON AUTENTICACION'].max() + resumen['SIN AUTENTICACION'].max()

    for i, row in enumerate(resumen.itertuples()):
        con_aut = int(row._1)  
        sin_aut = int(row._2)  
        total = con_aut + sin_aut
        cumplimiento = int(row.CUMPLIMIENTO * 100)
        incumplimiento = 100 - cumplimiento
        
        if con_aut > 0:
            ax.text(con_aut / 2, i, f"{con_aut}", ha='center', va='center', color='black', fontweight='bold', fontsize=8)
            
        if sin_aut > 0:
            ax.text(con_aut + (sin_aut / 2), i, f"{sin_aut}", ha='center', va='center', color='white', fontweight='bold', fontsize=8)

        texto_porcentajes = f"  Cump: {cumplimiento}% \n  Incump: {incumplimiento}%"
        ax.text(total, i, texto_porcentajes, ha='left', va='center', color='#F1C40F', fontweight='bold', fontsize=8)

    ax.set_title(f"TOP 10 OPERADORES - {combo_depto.get()} / {combo_muni.get()}\nDirección Nacional de Identificación", fontsize=11, fontweight='bold', color='white', pad=12)
    ax.set_xlabel("Volumen Total de Trámites Realizados", color='white', fontsize=9)
    
    ax.set_xlim(0, limite_ancho_grafica * 1.35)
    
    ax.tick_params(colors='white', axis='both', labelsize=8)
    ax.grid(axis='x', alpha=0.10, color='white')
    
    cump_patch = mpatches.Patch(color=COLOR_VERDE_EXITO, label='Con Autenticación')
    incump_patch = mpatches.Patch(color=COLOR_TEXTO_PASIVO, label='Sin Autenticación')
    ax.legend(handles=[cump_patch, incump_patch], facecolor='#161623', edgecolor="#444", loc='lower right', fontsize=8, labelcolor='white')
    
    plt.tight_layout()
    canvas = FigureCanvasTkAgg(fig, master=frame_grafico)
    canvas.draw()
    canvas.get_tk_widget().pack(fill='both', expand=True)


# ===================== LÓGICA DE LA PESTAÑA (SIN AUTENTICACIÓN) =====================
def actualizar_tabla_sin_autenticacion(event=None):
    global df_global
    if df_global is None: return
    
    for row in tabla_sin_aut.get_children():
        tabla_sin_aut.delete(row)
        
    df_sin_aut = df_global[df_global['SIN AUTENTICACION'] > 0].copy()
    
    depto_sel = combo_depto_sin.get()
    if depto_sel != 'TODOS LOS DEPARTAMENTOS':
        df_sin_aut = df_sin_aut[df_sin_aut['DEPARTAMENTO'] == depto_sel]
        
    muni_sel = combo_muni_sin.get()
    if muni_sel != 'TODOS LOS MUNICIPIOS' and muni_sel != 'SELECCIONE DEPARTAMENTO':
        df_sin_aut = df_sin_aut[df_sin_aut['MUNICIPIO'] == muni_sel]

    if df_sin_aut.empty:
        return

    resumen_sin = df_sin_aut.groupby(['CCOPERADOR', 'NOMBRE', 'MUNICIPIO'], as_index=False)['SIN AUTENTICACION'].sum()
    resumen_sin = resumen_sin.sort_values(by='SIN AUTENTICACION', ascending=False)
    
    # Renderizado aplicando los tags intercalados de color cebra ('par' / 'impar')
    for i, (indice, fila) in enumerate(resumen_sin.iterrows()):
        cc_val = fila['CCOPERADOR']
        cedula_str = str(int(cc_val)) if pd.notna(cc_val) and isinstance(cc_val, (int, float)) else str(cc_val)
        
        cant_sin_aut = int(fila['SIN AUTENTICACION'])
        nombre_val = fila['NOMBRE']
        muni_val = fila['MUNICIPIO']
        
        tag_fila = 'fila_par' if i % 2 == 0 else 'fila_impar'
        tabla_sin_aut.insert("", "end", values=(cedula_str, nombre_val, muni_val, cant_sin_aut), tags=(tag_fila,))
    
    tabla_sin_aut.update_idletasks()

def editar_nombre_operador(event):
    global df_global
    if df_global is None: return
    
    seleccion = tabla_sin_aut.selection()
    if not seleccion: return
    
    item = seleccion[0]
    valores = tabla_sin_aut.item(item, "values")
    
    cedula_seleccionada = valores[0]
    nombre_actual = valores[1]
    
    try:
        cedula_busqueda = float(cedula_seleccionada) if '.' in cedula_seleccionada else int(cedula_seleccionada)
    except ValueError:
        cedula_busqueda = cedula_seleccionada

    nuevo_nombre = simpledialog.askstring("Corrección de Datos", 
                                         f"Asignar Identificación para la CC: {cedula_seleccionada}\nValor actual: {nombre_actual}\n\nEscriba el nombre completo del operador:",
                                         initialvalue=nombre_actual if nombre_actual not in ["SIN NOMBRE IDENTIFICADO", "S.N."] else "")
    
    if nuevo_nombre and nuevo_nombre.strip() != "":
        nuevo_nombre = nuevo_nombre.strip().upper()
        
        df_global.loc[df_global['CCOPERADOR'] == cedula_busqueda, 'NOMBRE'] = nuevo_nombre
        actualizar_tabla_sin_autenticacion()
        messagebox.showinfo("Registro Actualizado", f"Cambios grabados en memoria.\nOperador CC {cedula_seleccionada} -> {nuevo_nombre}")


# ===================== INTERFAZ GRÁFICA MAESTRA (AESTHETIC) =====================
root = tk.Tk()
root.title("CUMPLIMIENTO MEMORANDO RDRCI DNI 062")
root.geometry("1120x760")
root.configure(bg=COLOR_FONDO_APP)

# --- CABECERA OFICIAL ---
header = tk.Frame(root, bg=COLOR_AZUL_RNEC, height=105)
header.pack(fill='x')
header.pack_propagate(False)

try:
    img_path = os.path.join(os.path.dirname(__file__), "Escudo.png")
    pil_img = Image.open(img_path).resize((105, 85), Image.LANCZOS)
    img_logo = ImageTk.PhotoImage(pil_img)
    lbl_l = tk.Label(header, image=img_logo, bg=COLOR_AZUL_RNEC)
    lbl_l.image = img_logo
    lbl_l.pack(side='left', padx=25)
except:
    tk.Label(header, text="REGISTRADURÍA", fg=COLOR_BLANCO, bg=COLOR_AZUL_RNEC, font=("Segoe UI", 12, "bold")).pack(side='left', padx=25)

tk.Label(header, text="SISTEMA DE CONTROL VISUAL - INFORME DNI\nREGISTRADURÍA NACIONAL DEL ESTADO CIVIL", fg="white", bg=COLOR_AZUL_RNEC, font=FONT_TITULO, justify="left").pack(side='left', padx=5)

# --- PANEL CENTRAL (TARJETA MODERNA) ---
main_card = tk.Frame(root, bg=COLOR_AZUL_RNEC, highlightthickness=0)
main_card.place(relx=0.5, rely=0.55, anchor='center', width=1040, height=580)

# ===================== CONFIGURACIÓN DE ESTILOS AVANZADOS (TREEVIEW) =====================
style = ttk.Style()
style.theme_use('default')

style.configure("TNotebook", background=COLOR_AZUL_RNEC, borderwidth=0)
style.configure("TNotebook.Tab", background="#161623", padding=[22, 6], font=("Segoe UI", 9, "bold"), foreground="white")
style.map("TNotebook.Tab", background=[("selected", "#0F3460")], foreground=[("selected", "white")])

# 1. Ajustes del Treeview (Fondo de la tabla y remoción de bordes por defecto)
style.configure("Treeview", 
                background=COLOR_FONDO_APP, 
                fieldbackground=COLOR_FONDO_APP, 
                rowheight=30, 
                font=("Segoe UI", 10),
                borderwidth=0)

# 2. Diseño de Cabeceras Estilo Registraduría (Fondo azul institucional, letras blancas nítidas)
style.configure("Treeview.Heading", 
                background=COLOR_AZUL_RNEC, 
                foreground=COLOR_BLANCO, 
                font=("Segoe UI", 10, "bold"),
                relief="flat",
                padding=6)

# Mapa para mantener los colores fijos incluso si el usuario pasa el mouse por encima
style.map("Treeview.Heading",
          background=[('active', COLOR_AZUL_RNEC)],
          foreground=[('active', COLOR_BLANCO)])

style.map("Treeview", 
          background=[('selected', COLOR_FILA_SELECCION)], 
          foreground=[('selected', COLOR_BLANCO)])

nb = ttk.Notebook(main_card)
nb.pack(fill='both', expand=True)

# --- PESTAÑA 1: ADQUISICIÓN DE DATOS ---
tab1 = tk.Frame(nb, bg=COLOR_FONDO_APP)
nb.add(tab1, text="📁 Carga de Datos")

tk.Label(tab1, text="Gestión de Archivos de Cumplimiento", font=FONT_SUBTITULO, bg=COLOR_FONDO_APP, fg="white").pack(pady=20)

RoundedButton(tab1, "📄 Seleccionar Archivos BASE", COLOR_AZUL_RNEC, seleccionar_base).pack(pady=8)
lbl_base = tk.Label(tab1, text="Ninguno", bg=COLOR_FONDO_APP, font=FONT_NORMAL, fg=COLOR_GRIS_TEXTO)
lbl_base.pack()

RoundedButton(tab1, "📋 Seleccionar Archivo NOMBRES", COLOR_AZUL_RNEC, seleccionar_nombres).pack(pady=8)
lbl_nombres = tk.Label(tab1, text="Ninguno", bg=COLOR_FONDO_APP, font=FONT_NORMAL, fg=COLOR_GRIS_TEXTO)
lbl_nombres.pack()

tk.Frame(tab1, height=15, bg=COLOR_FONDO_APP).pack()

RoundedButton(tab1, "⚡ UNIFICAR DATOS", COLOR_VERDE_EXITO, unificar, width=360).pack(pady=15)

lbl_estado = tk.Label(tab1, text="Estado: Consola preparada para recibir archivos.", bg=COLOR_FONDO_APP, font=("Segoe UI", 9, "italic"), fg="#F1C40F")
lbl_estado.pack()

btn_descargar_frame = tk.Frame(tab1, bg=COLOR_FONDO_APP)
btn_descargar_frame.pack(side='bottom', pady=15)
btn_descargar = tk.Button(btn_descargar_frame, text="📥 DESCARGAR EXCEL", command=descargar_archivo, bg="#B8860B", fg="white", font=("Segoe UI", 10, "bold"), relief='flat', padx=25, pady=8, state='disabled')
btn_descargar.pack()

# --- PESTAÑA 2: DASHBOARD DINÁMICO ---
tab2 = tk.Frame(nb, bg=COLOR_FONDO_APP)
nb.add(tab2, text="📊 Dashboard")

filtros_bar = tk.Frame(tab2, bg="#161623", height=55)
filtros_bar.pack(fill='x', padx=15, pady=12)

tk.Label(filtros_bar, text="📍 Departamento:", bg="#161623", font=("Segoe UI", 9, "bold"), fg="white").pack(side='left', padx=15)
combo_depto = ttk.Combobox(filtros_bar, state='readonly', width=26, font=FONT_NORMAL)
combo_depto.pack(side='left', padx=5)
combo_depto.bind('<<ComboboxSelected>>', actualizar_municipios)

tk.Label(filtros_bar, text="🗺️ Municipio:", bg="#161623", font=("Segoe UI", 9, "bold"), fg="white").pack(side='left', padx=15)
combo_muni = ttk.Combobox(filtros_bar, state='readonly', width=26, font=FONT_NORMAL)
combo_muni.pack(side='left', padx=5)

btn_view = tk.Button(filtros_bar, text="📊 GENERAR GRÁFICO", command=generar_grafico, bg=COLOR_VERDE_EXITO, fg="white", font=("Segoe UI", 10, "bold"), padx=18, relief='flat', cursor='hand2')
btn_view.pack(side='right', padx=15, pady=10)

frame_grafico = tk.Frame(tab2, bg='#1A1A2E')
frame_grafico.pack(fill='both', expand=True, padx=15, pady=5)


# --- PESTAÑA 3: DETALLE SIN AUTENTICACIÓN (DISEÑO INSTITUCIONAL RNEC) ---
tab3 = tk.Frame(nb, bg=COLOR_FONDO_APP)
nb.add(tab3, text="⚠️ Sin Autenticación")

filtros_bar_sin = tk.Frame(tab3, bg="#161623", height=55)
filtros_bar_sin.pack(fill='x', padx=15, pady=10)

tk.Label(filtros_bar_sin, text="📍 Filtrar Departamento:", bg="#161623", font=("Segoe UI", 9, "bold"), fg="white").pack(side='left', padx=15)
combo_depto_sin = ttk.Combobox(filtros_bar_sin, state='readonly', width=24, font=FONT_NORMAL)
combo_depto_sin.pack(side='left', padx=5)
combo_depto_sin.bind('<<ComboboxSelected>>', actualizar_municipios_sin)

tk.Label(filtros_bar_sin, text="🗺️ Municipio:", bg="#161623", font=("Segoe UI", 9, "bold"), fg="white").pack(side='left', padx=15)
combo_muni_sin = ttk.Combobox(filtros_bar_sin, state='readonly', width=24, font=FONT_NORMAL)
combo_muni_sin.pack(side='left', padx=5)
combo_muni_sin.bind('<<ComboboxSelected>>', actualizar_tabla_sin_autenticacion)

tk.Label(filtros_bar_sin, text="📝 Doble clic fila para corregir S.N.", bg="#161623", font=("Segoe UI", 8, "italic"), fg="#F1C40F").pack(side='right', padx=15)

frame_tabla = tk.Frame(tab3, bg=COLOR_FONDO_APP)
frame_tabla.pack(fill='both', expand=True, padx=15, pady=5)

scroll_y = ttk.Scrollbar(frame_tabla, orient="vertical")
scroll_y.pack(side='right', fill='y')

columnas = ("cedula", "nombre", "municipio", "cantidad")
tabla_sin_aut = ttk.Treeview(frame_tabla, columns=columnas, show="headings", yscrollcommand=scroll_y.set)
scroll_y.config(command=tabla_sin_aut.yview)

# CONFIGURACIÓN VISUAL DEL EFECTO CEBRA (Legibilidad Óptima)
tabla_sin_aut.tag_configure('fila_par', background=COLOR_FILA_PAR, foreground=COLOR_BLANCO)
tabla_sin_aut.tag_configure('fila_impar', background=COLOR_FILA_IMPAR, foreground=COLOR_BLANCO)

tabla_sin_aut.heading("cedula", text="CÉDULA OPERADOR")
tabla_sin_aut.heading("nombre", text="NOMBRE DEL OPERADOR")
tabla_sin_aut.heading("municipio", text="MUNICIPIO")
tabla_sin_aut.heading("cantidad", text="CANT. SIN AUTENTICAR")

tabla_sin_aut.column("cedula", width=150, anchor="center")
tabla_sin_aut.column("nombre", width=350, anchor="w")
tabla_sin_aut.column("municipio", width=250, anchor="w")
tabla_sin_aut.column("cantidad", width=150, anchor="center")

tabla_sin_aut.pack(fill='both', expand=True)
tabla_sin_aut.bind("<Double-1>", editar_nombre_operador)

# --- PIE DE PÁGINA CORPORATIVO ---
footer = tk.Frame(root, bg=COLOR_AZUL_RNEC, height=28)
footer.pack(side='bottom', fill='x')
tk.Label(footer, text="Registraduría Nacional del Estado Civil - Dirección Nacional de Identificación • Vigilado por la alta dirección 2026", fg="white", bg=COLOR_AZUL_RNEC, font=("Segoe UI", 8)).pack(pady=4)

root.mainloop()