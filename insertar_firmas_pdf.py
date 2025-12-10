#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Aplicación para insertar firmas en PDFs - VERSIÓN MEJORADA
Con mejor manejo de errores y logging detallado
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import fitz  # PyMuPDF
from PIL import Image
import os
import re
import sys
import traceback
from pathlib import Path
import unicodedata

def normalizar_texto(texto):
    """Quita acentos, convierte a mayúsculas y limpia espacios"""
    # Quitar acentos
    texto_sin_acentos = ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )
    # Convertir a mayúsculas y limpiar espacios
    return texto_sin_acentos.upper().strip()

def limpiar_nombre_para_busqueda(texto):
    """Limpia el texto quitando 'Fdo:', acentos y espacios extra"""
    # Quitar "Fdo:" y variantes
    texto = re.sub(r'^Fdo\.?\s*:\s*', '', texto, flags=re.IGNORECASE)
    texto = re.sub(r'^Firmado\s*:\s*', '', texto, flags=re.IGNORECASE)
    # Normalizar
    return normalizar_texto(texto)

class InsertadorFirmasPDF:
    def __init__(self, root):
        self.root = root
        self.root.title("Insertar Firmas en PDFs - Plan de Formación")
        self.root.geometry("900x700")
        
        self.pdfs_seleccionados = []
        self.excel_path = None
        self.directorio_salida = None
        self.firmas_cache = {}
        
        self.crear_interfaz()
    
    def crear_interfaz(self):
        """Crea la interfaz gráfica"""
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(2, weight=1)
        
        # Título
        titulo = ttk.Label(main_frame, text="Insertar Firmas en PDFs", 
                          font=('Arial', 18, 'bold'))
        titulo.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # SECCIÓN 1: PDFs
        frame_pdf = ttk.LabelFrame(main_frame, text="1. Archivos PDF", padding="10")
        frame_pdf.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(frame_pdf, text="📄 Añadir PDFs", 
                  command=self.seleccionar_pdfs, width=20).grid(row=0, column=0, padx=5)
        
        self.label_pdfs = ttk.Label(frame_pdf, text="No hay PDFs seleccionados", 
                                    foreground="gray")
        self.label_pdfs.grid(row=0, column=1, sticky=tk.W, padx=10)
        
        # SECCIÓN 2: Excel
        frame_excel = ttk.LabelFrame(main_frame, text="2. Excel con Firmas", padding="10")
        frame_excel.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(frame_excel, text="📊 Seleccionar Excel", 
                  command=self.seleccionar_excel, width=20).grid(row=0, column=0, padx=5)
        
        self.label_excel = ttk.Label(frame_excel, text="No hay Excel seleccionado", 
                                     foreground="gray")
        self.label_excel.grid(row=0, column=1, sticky=tk.W, padx=10)
        
        # SECCIÓN 3: Directorio
        frame_salida = ttk.LabelFrame(main_frame, text="3. Carpeta de Salida (Opcional)", padding="10")
        frame_salida.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(frame_salida, text="📁 Seleccionar Carpeta", 
                  command=self.seleccionar_directorio, width=20).grid(row=0, column=0, padx=5)
        
        self.label_directorio = ttk.Label(frame_salida, 
                                          text="Misma carpeta que PDFs originales", 
                                          foreground="gray")
        self.label_directorio.grid(row=0, column=1, sticky=tk.W, padx=10)
        
        # CONFIGURACIÓN
        frame_config = ttk.LabelFrame(main_frame, text="Configuración", padding="10")
        frame_config.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(frame_config, text="Ancho (px):").grid(row=0, column=0)
        self.ancho_firma = tk.IntVar(value=120)
        ttk.Spinbox(frame_config, from_=50, to=300, textvariable=self.ancho_firma, 
                   width=10).grid(row=0, column=1, padx=5)
        
        ttk.Label(frame_config, text="Alto (px):").grid(row=0, column=2, padx=(20, 0))
        self.alto_firma = tk.IntVar(value=43)  # Valor óptimo
        ttk.Spinbox(frame_config, from_=30, to=200, textvariable=self.alto_firma, 
                   width=10).grid(row=0, column=3, padx=5)
        
        ttk.Label(frame_config, text="Margen (px):").grid(row=0, column=4, padx=(20, 0))
        self.margen_superior = tk.IntVar(value=0)  # Sin margen
        ttk.Spinbox(frame_config, from_=-50, to=50, textvariable=self.margen_superior, 
                   width=10).grid(row=0, column=5, padx=5)
        
        # BOTÓN PROCESAR
        self.btn_procesar = ttk.Button(main_frame, text="▶ Procesar PDFs", 
                                       command=self.procesar_pdfs, 
                                       state='disabled')
        self.btn_procesar.grid(row=5, column=0, columnspan=3, pady=20)
        
        # LOG
        frame_log = ttk.LabelFrame(main_frame, text="Log de Operaciones", padding="10")
        frame_log.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        self.log_text = scrolledtext.ScrolledText(frame_log, width=100, height=18, 
                                                  font=('Courier', 9))
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        frame_log.columnconfigure(0, weight=1)
        frame_log.rowconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        # Barra de progreso
        self.progreso = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progreso.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        self.log("✓ Aplicación iniciada")
        self.log("ℹ Seleccione PDFs y Excel para comenzar")
    
    def log(self, mensaje, tipo="info"):
        """Log con timestamp"""
        timestamp = pd.Timestamp.now().strftime("%H:%M:%S")
        prefijos = {"info": "", "success": "✓ ", "warning": "⚠ ", "error": "✗ "}
        prefijo = prefijos.get(tipo, "")
        self.log_text.insert(tk.END, f"[{timestamp}] {prefijo}{mensaje}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def seleccionar_pdfs(self):
        archivos = filedialog.askopenfilenames(
            title="Seleccionar PDFs",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")]
        )
        if archivos:
            self.pdfs_seleccionados = list(archivos)
            self.label_pdfs.config(
                text=f"✓ {len(self.pdfs_seleccionados)} archivo(s)",
                foreground="green"
            )
            self.log(f"PDFs seleccionados: {len(self.pdfs_seleccionados)}", "success")
            for pdf in self.pdfs_seleccionados:
                self.log(f"  • {os.path.basename(pdf)}")
            self.verificar_estado_procesar()
    
    def seleccionar_excel(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar Excel",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]
        )
        if archivo:
            self.excel_path = archivo
            self.label_excel.config(text=f"✓ {os.path.basename(archivo)}", foreground="green")
            self.log(f"Excel: {os.path.basename(archivo)}", "success")
            self.verificar_estado_procesar()
    
    def seleccionar_directorio(self):
        directorio = filedialog.askdirectory(title="Carpeta de salida")
        if directorio:
            self.directorio_salida = directorio
            self.label_directorio.config(text=f"✓ {os.path.basename(directorio)}", foreground="green")
            self.log(f"Salida: {directorio}", "success")
    
    def verificar_estado_procesar(self):
        if self.pdfs_seleccionados and self.excel_path:
            self.btn_procesar.config(state='normal')
        else:
            self.btn_procesar.config(state='disabled')
    
    def leer_excel_firmas(self):
        """Lee Excel y crea diccionarios de búsqueda"""
        try:
            self.log("\n" + "="*70)
            self.log("Leyendo Excel...")
            
            df = pd.read_excel(self.excel_path)
            self.log(f"✓ {len(df)} personas encontradas", "success")
            self.log(f"Columnas: {', '.join(df.columns.tolist())}")
            
            firmas_por_telefono = {}
            firmas_por_nombre = {}
            
            self.log("\nPersonas en el Excel:")
            
            for idx, row in df.iterrows():
                nombre = str(row['nombre']).strip() if pd.notna(row['nombre']) else ""
                apellido1 = str(row['apellido1']).strip() if pd.notna(row['apellido1']) else ""
                apellido2 = str(row['apellido2']).strip() if pd.notna(row['apellido2']) else ""
                
                # Crear versión normalizada (sin acentos, en mayúsculas)
                nombre_completo = normalizar_texto(f"{nombre} {apellido1} {apellido2}")
                nombre_sin_segundo = normalizar_texto(f"{nombre} {apellido1}")
                
                telefono = None
                if pd.notna(row['telefono']):
                    telefono_str = str(int(row['telefono'])) if isinstance(row['telefono'], float) else str(row['telefono'])
                    telefono = re.sub(r'\D', '', telefono_str)
                
                ruta = str(row['ruta']).strip() if pd.notna(row['ruta']) else None
                
                if not ruta:
                    self.log(f"  ⚠ {nombre_completo}: Sin ruta", "warning")
                    continue
                
                # Verificar si la ruta existe
                ruta_existe = False
                ruta_final = ruta
                
                if os.path.exists(ruta):
                    ruta_existe = True
                    ruta_final = ruta
                else:
                    # Intentar ruta relativa
                    ruta_relativa = os.path.join(os.path.dirname(self.excel_path), ruta)
                    if os.path.exists(ruta_relativa):
                        ruta_existe = True
                        ruta_final = ruta_relativa
                
                if telefono and len(telefono) >= 9:
                    firmas_por_telefono[telefono] = ruta_final
                    status = "✓" if ruta_existe else "✗"
                    self.log(f"  {status} {nombre_completo} (Tel: {telefono})")
                else:
                    status = "✓" if ruta_existe else "✗"
                    self.log(f"  {status} {nombre_completo} (Sin teléfono)")
                
                if not ruta_existe:
                    self.log(f"     ⚠ ARCHIVO NO ENCONTRADO: {ruta[:60]}...", "warning")
                
                firmas_por_nombre[nombre_completo] = ruta_final
                if nombre_sin_segundo != nombre_completo:
                    firmas_por_nombre[nombre_sin_segundo] = ruta_final
            
            self.log(f"\n✓ Con teléfono: {len(firmas_por_telefono)}", "success")
            self.log(f"✓ Total registradas: {len(set(firmas_por_nombre.values()))}", "success")
            
            return {
                'por_telefono': firmas_por_telefono,
                'por_nombre': firmas_por_nombre
            }
        
        except Exception as e:
            self.log(f"✗ ERROR Excel: {str(e)}", "error")
            self.log(traceback.format_exc())
            messagebox.showerror("Error", f"No se pudo leer el Excel:\n\n{str(e)}")
            return None
    
    def buscar_campos_por_texto(self, page, firmas_dict):
        """
        Busca textos en el PDF que coincidan con nombres/teléfonos del Excel
        Sin depender de resaltados
        """
        campos = {
            'alumno': {'texto': None, 'rect': None},
            'profesor': {'texto': None, 'rect': None}
        }
        
        try:
            # Obtener todo el texto de la página con sus posiciones
            bloques_texto = page.get_text("dict")["blocks"]
            
            self.log(f"    Analizando {len(bloques_texto)} bloques de texto...")
            
            for bloque in bloques_texto:
                if bloque.get("type") == 0:  # Tipo 0 = bloque de texto
                    for linea in bloque.get("lines", []):
                        for span in linea.get("spans", []):
                            texto_original = span["text"].strip()
                            if not texto_original:
                                continue
                            
                            # Obtener rectángulo del texto
                            rect = fitz.Rect(span["bbox"])
                            
                            # Limpiar y normalizar el texto
                            texto_normalizado = limpiar_nombre_para_busqueda(texto_original)
                            
                            # Buscar si es un TELÉFONO de alumno
                            numeros = ''.join(re.findall(r'\d+', texto_original))
                            if len(numeros) >= 9:
                                if numeros in firmas_dict['por_telefono']:
                                    # Determinar posición (derecha = alumno)
                                    pos_horizontal = "DERECHA" if rect.x0 > page.rect.width / 2 else "IZQUIERDA"
                                    if pos_horizontal == "DERECHA":
                                        campos['alumno']['texto'] = numeros
                                        campos['alumno']['rect'] = rect
                                        self.log(f"    📱 ALUMNO encontrado (teléfono): {numeros} en {pos_horizontal}", "success")
                                        continue
                            
                            # Buscar si es un NOMBRE (alumno o profesor)
                            if len(texto_normalizado) > 5:  # Ignorar textos muy cortos
                                # Buscar en el diccionario de nombres
                                for nombre_excel in firmas_dict['por_nombre'].keys():
                                    # Coincidencia exacta o parcial
                                    if texto_normalizado == nombre_excel or \
                                       nombre_excel in texto_normalizado or \
                                       texto_normalizado in nombre_excel:
                                        
                                        # Determinar posición (derecha = alumno, izquierda = profesor)
                                        pos_horizontal = "DERECHA" if rect.x0 > page.rect.width / 2 else "IZQUIERDA"
                                        
                                        if pos_horizontal == "DERECHA" and not campos['alumno']['texto']:
                                            campos['alumno']['texto'] = texto_normalizado
                                            campos['alumno']['rect'] = rect
                                            self.log(f"    👤 ALUMNO encontrado: '{texto_original}' → '{texto_normalizado}' en {pos_horizontal}", "success")
                                        elif pos_horizontal == "IZQUIERDA" and not campos['profesor']['texto']:
                                            campos['profesor']['texto'] = texto_normalizado
                                            campos['profesor']['rect'] = rect
                                            self.log(f"    👨‍🏫 PROFESOR encontrado: '{texto_original}' → '{texto_normalizado}' en {pos_horizontal}", "success")
                                        break
        
        except Exception as e:
            self.log(f"    ✗ Error buscando textos: {str(e)}", "error")
            self.log(traceback.format_exc())
        
        return campos
    
    def cargar_firma(self, ruta):
        """Carga firma desde archivo"""
        try:
            if ruta in self.firmas_cache:
                self.log(f"      ↻ Usando caché")
                return self.firmas_cache[ruta]
            
            if not os.path.exists(ruta):
                self.log(f"      ✗ Archivo no existe: {ruta}", "error")
                return None
            
            self.log(f"      ↓ Cargando: {os.path.basename(ruta)}")
            firma_img = Image.open(ruta)
            
            if firma_img.mode != 'RGB':
                firma_img = firma_img.convert('RGB')
            
            self.firmas_cache[ruta] = firma_img
            self.log(f"      ✓ Cargada ({firma_img.size[0]}x{firma_img.size[1]})", "success")
            return firma_img
        
        except Exception as e:
            self.log(f"      ✗ Error cargando: {str(e)}", "error")
            return None
    
    def insertar_firma_en_pdf(self, page, firma_img, rect, tipo='alumno'):
        """Inserta firma ENCIMA del texto resaltado"""
        try:
            import tempfile
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
            temp_path = temp_file.name
            temp_file.close()
            
            firma_img.save(temp_path, "PNG")
            
            ancho = self.ancho_firma.get()
            alto = self.alto_firma.get()
            margen = self.margen_superior.get()
            
            # IMPORTANTE: Centrar horizontalmente sobre el texto
            x0 = rect.x0 + (rect.width - ancho) / 2
            x1 = x0 + ancho
            
            # CRÍTICO: En PyMuPDF, las coordenadas Y crecen hacia ABAJO
            # rect.y0 = parte SUPERIOR del rectángulo (valor menor)
            # rect.y1 = parte INFERIOR del rectángulo (valor mayor)
            # Para poner la firma ENCIMA, usamos rect.y0 (parte superior)
            
            # La firma va ENCIMA, así que:
            # y1 (parte inferior de la firma) = parte superior del texto - margen
            y1 = rect.y0 - margen
            # y0 (parte superior de la firma) = y1 - altura de la firma
            y0 = y1 - alto
            
            firma_rect = fitz.Rect(x0, y0, x1, y1)
            
            self.log(f"      📍 Posición texto: y0={int(rect.y0)}, y1={int(rect.y1)}")
            self.log(f"      📍 Posición firma: y0={int(y0)}, y1={int(y1)} (ENCIMA del texto)")
            
            page.insert_image(firma_rect, filename=temp_path, overlay=True)
            
            os.unlink(temp_path)
            
            self.log(f"      ✓ Firma insertada correctamente", "success")
            return True
        
        except Exception as e:
            self.log(f"      ✗ Error insertando: {str(e)}", "error")
            return False
    
    def buscar_firma(self, identificador, firmas_dict):
        """Busca firma por identificador con logging detallado"""
        self.log(f"    🔎 Buscando: '{identificador}'")
        
        # Por teléfono
        if identificador.isdigit() and len(identificador) >= 9:
            if identificador in firmas_dict['por_telefono']:
                self.log(f"    ✓ Encontrado por teléfono", "success")
                return firmas_dict['por_telefono'][identificador]
            else:
                self.log(f"    ⚠ Teléfono no está en el Excel")
        
        # Por nombre exacto
        if identificador in firmas_dict['por_nombre']:
            self.log(f"    ✓ Coincidencia exacta en Excel", "success")
            return firmas_dict['por_nombre'][identificador]
        
        # Búsqueda parcial
        self.log(f"    ℹ Buscando coincidencia parcial...")
        self.log(f"    ℹ Nombres en Excel: {list(firmas_dict['por_nombre'].keys())[:5]}...")
        
        for nombre in firmas_dict['por_nombre'].keys():
            # Coincidencia si uno contiene al otro
            if identificador in nombre or nombre in identificador:
                self.log(f"    ✓ Coincidencia parcial: '{nombre}'", "success")
                return firmas_dict['por_nombre'][nombre]
            
            # Por palabras (nombre + apellido)
            partes_id = identificador.split()
            partes_nom = nombre.split()
            if len(partes_id) >= 2 and len(partes_nom) >= 2:
                if partes_id[0] == partes_nom[0] and partes_id[1] == partes_nom[1]:
                    self.log(f"    ✓ Coincidencia por nombre+apellido: '{nombre}'", "success")
                    return firmas_dict['por_nombre'][nombre]
        
        self.log(f"    ✗ NO se encontró ninguna coincidencia", "error")
        return None
    
    def procesar_pdf(self, pdf_path, firmas_dict):
        """Procesa un PDF"""
        self.log(f"\n{'='*70}")
        self.log(f"📄 {os.path.basename(pdf_path)}")
        self.log(f"{'='*70}")
        
        doc = None
        try:
            doc = fitz.open(pdf_path)
            self.log(f"✓ PDF abierto: {len(doc)} página(s)")
            modificado = False
            
            # Procesar solo desde la página 2 en adelante (saltar encabezado)
            # Si el PDF tiene solo 1 página, procesarla de todas formas
            pagina_inicio = 1 if len(doc) > 1 else 0
            
            for num_pag in range(pagina_inicio, len(doc)):
                page = doc[num_pag]
                self.log(f"\n  Página {num_pag + 1}:")
                
                campos = self.buscar_campos_por_texto(page, firmas_dict)
                
                # ALUMNO
                if campos['alumno']['texto'] and campos['alumno']['rect']:
                    identificador = campos['alumno']['texto']
                    self.log(f"  🔍 ===== PROCESANDO ALUMNO =====")
                    self.log(f"  Identificador detectado: '{identificador}'")
                    self.log(f"  Tipo: {'TELÉFONO' if identificador.isdigit() else 'NOMBRE'}")
                    
                    ruta_firma = self.buscar_firma(identificador, firmas_dict)
                    
                    if ruta_firma:
                        self.log(f"    ✓ Ruta de firma obtenida: {ruta_firma}", "success")
                        self.log(f"    ✓ Verificando si existe el archivo...")
                        
                        if os.path.exists(ruta_firma):
                            self.log(f"    ✓ Archivo existe", "success")
                        else:
                            self.log(f"    ✗ ARCHIVO NO EXISTE: {ruta_firma}", "error")
                        
                        firma_img = self.cargar_firma(ruta_firma)
                        
                        if firma_img:
                            self.log(f"    ✓ Imagen cargada correctamente", "success")
                            resultado = self.insertar_firma_en_pdf(page, firma_img, campos['alumno']['rect'], 'alumno')
                            if resultado:
                                modificado = True
                                self.log(f"    ✓✓✓ FIRMA DE ALUMNO INSERTADA", "success")
                            else:
                                self.log(f"    ✗✗✗ ERROR AL INSERTAR FIRMA", "error")
                        else:
                            self.log(f"    ✗ No se pudo cargar la imagen", "error")
                    else:
                        self.log(f"    ✗ NO SE ENCONTRÓ RUTA DE FIRMA", "error")
                        self.log(f"    ✗ Identificador buscado: '{identificador}'")
                        self.log(f"    ✗ Diccionario de nombres disponibles:")
                        for nombre_excel in list(firmas_dict['por_nombre'].keys()):
                            self.log(f"      - '{nombre_excel}'")
                else:
                    self.log(f"  ⚠ No se detectó campo amarillo (alumno)", "warning")
                    if not campos['alumno']['texto']:
                        self.log(f"     Razón: No hay texto detectado")
                    if not campos['alumno']['rect']:
                        self.log(f"     Razón: No hay rectángulo detectado")
                
                # PROFESOR
                if campos['profesor']['texto'] and campos['profesor']['rect']:
                    nombre_prof = campos['profesor']['texto']
                    self.log(f"  🔍 ===== PROCESANDO PROFESOR =====")
                    self.log(f"  Nombre detectado: '{nombre_prof}'")
                    
                    ruta_firma = self.buscar_firma(nombre_prof, firmas_dict)
                    
                    if ruta_firma:
                        self.log(f"    ✓ Ruta de firma obtenida: {ruta_firma}", "success")
                        
                        firma_img = self.cargar_firma(ruta_firma)
                        
                        if firma_img:
                            self.log(f"    ✓ Imagen cargada correctamente", "success")
                            resultado = self.insertar_firma_en_pdf(page, firma_img, campos['profesor']['rect'], 'profesor')
                            if resultado:
                                modificado = True
                                self.log(f"    ✓✓✓ FIRMA DE PROFESOR INSERTADA", "success")
                            else:
                                self.log(f"    ✗✗✗ ERROR AL INSERTAR FIRMA", "error")
                        else:
                            self.log(f"    ✗ No se pudo cargar la imagen", "error")
                    else:
                        self.log(f"    ✗ NO SE ENCONTRÓ RUTA DE FIRMA", "error")
                else:
                    self.log(f"  ⚠ No se detectó campo rojo/naranja (profesor)", "warning")
            
            if modificado:
                # Guardar
                if self.directorio_salida:
                    nombre = os.path.basename(pdf_path)
                    nombre_sin_ext = os.path.splitext(nombre)[0]
                    ruta_salida = os.path.join(self.directorio_salida, f"{nombre_sin_ext}_firmado.pdf")
                else:
                    directorio = os.path.dirname(pdf_path)
                    nombre_sin_ext = os.path.splitext(os.path.basename(pdf_path))[0]
                    ruta_salida = os.path.join(directorio, f"{nombre_sin_ext}_firmado.pdf")
                
                self.log(f"\n  💾 Guardando en: {ruta_salida}")
                doc.save(ruta_salida, garbage=4, deflate=True, clean=True)
                doc.close()
                
                if os.path.exists(ruta_salida):
                    self.log(f"  ✓✓✓ GUARDADO EXITOSAMENTE", "success")
                    return True, ruta_salida
                else:
                    self.log(f"  ✗✗✗ ERROR: Archivo no se creó", "error")
                    return False, None
            else:
                doc.close()
                self.log(f"\n  ⚠ Sin modificaciones (no se encontraron campos o firmas)", "warning")
                return False, None
        
        except Exception as e:
            self.log(f"\n  ✗✗✗ ERROR: {str(e)}", "error")
            self.log(traceback.format_exc())
            if doc:
                doc.close()
            return False, None
    
    def procesar_pdfs(self):
        """Procesa todos los PDFs"""
        self.log("\n\n" + "="*70)
        self.log("🚀 INICIANDO PROCESAMIENTO")
        self.log("="*70 + "\n")
        
        self.btn_procesar.config(state='disabled')
        self.progreso.start()
        
        try:
            firmas_dict = self.leer_excel_firmas()
            if not firmas_dict:
                return
            
            exitosos = 0
            fallidos = 0
            archivos_generados = []
            
            for i, pdf_path in enumerate(self.pdfs_seleccionados, 1):
                self.log(f"\n📊 Progreso: {i}/{len(self.pdfs_seleccionados)}")
                exito, ruta_salida = self.procesar_pdf(pdf_path, firmas_dict)
                
                if exito:
                    exitosos += 1
                    archivos_generados.append(ruta_salida)
                else:
                    fallidos += 1
            
            self.log("\n" + "="*70)
            self.log("📊 RESUMEN")
            self.log("="*70)
            self.log(f"✓ Exitosos: {exitosos}")
            self.log(f"✗ Fallidos: {fallidos}")
            
            if archivos_generados:
                self.log(f"\n📂 Archivos generados:")
                for archivo in archivos_generados:
                    self.log(f"  • {archivo}")
            
            self.log("\n" + "="*70)
            if exitosos > 0:
                self.log("✓✓✓ COMPLETADO", "success")
            else:
                self.log("⚠⚠⚠ NINGÚN PDF SE PROCESÓ CORRECTAMENTE", "warning")
            self.log("="*70)
            
            if exitosos > 0:
                messagebox.showinfo("Completado", 
                                   f"✓ Exitosos: {exitosos}\n✗ Fallidos: {fallidos}")
            else:
                messagebox.showwarning("Atención", 
                                      "No se procesó ningún PDF correctamente.\n" +
                                      "Revise el log para más detalles.")
        
        finally:
            self.btn_procesar.config(state='normal')
            self.progreso.stop()

def main():
    root = tk.Tk()
    style = ttk.Style()
    style.theme_use('clam')
    app = InsertadorFirmasPDF(root)
    root.mainloop()

if __name__ == "__main__":
    main()
