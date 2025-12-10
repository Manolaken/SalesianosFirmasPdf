#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de prueba para verificar la lectura del Excel LibroFirmas.xlsx
Uso: python test_excel.py
"""

import pandas as pd
import os
import re
from tkinter import filedialog, Tk

def verificar_excel(excel_path):
    """Verifica la estructura y contenido del Excel"""
    
    print("="*70)
    print("VERIFICACIÓN DEL EXCEL LibroFirmas.xlsx")
    print("="*70)
    print(f"\nArchivo: {os.path.basename(excel_path)}")
    print(f"Ubicación: {os.path.dirname(excel_path)}")
    
    try:
        # Leer Excel
        df = pd.read_excel(excel_path)
        
        print(f"\n✓ Excel cargado correctamente")
        print(f"  Filas: {len(df)}")
        print(f"  Columnas: {len(df.columns)}")
        
        # Verificar columnas esperadas
        columnas_esperadas = ['id', 'nombre', 'apellido1', 'apellido2', 'telefono', 'ruta']
        columnas_presentes = df.columns.tolist()
        
        print(f"\n{'='*70}")
        print("VERIFICACIÓN DE COLUMNAS")
        print("="*70)
        
        for col in columnas_esperadas:
            if col in columnas_presentes:
                print(f"  ✓ Columna '{col}' encontrada")
            else:
                print(f"  ✗ Columna '{col}' NO encontrada - REQUERIDA")
        
        # Procesar personas
        print(f"\n{'='*70}")
        print("PERSONAS EN EL EXCEL")
        print("="*70)
        
        firmas_por_telefono = {}
        firmas_por_nombre = {}
        
        for idx, row in df.iterrows():
            nombre = str(row['nombre']).strip().upper() if pd.notna(row['nombre']) else ""
            apellido1 = str(row['apellido1']).strip().upper() if pd.notna(row['apellido1']) else ""
            apellido2 = str(row['apellido2']).strip().upper() if pd.notna(row['apellido2']) else ""
            
            nombre_completo = f"{nombre} {apellido1} {apellido2}".strip()
            
            telefono = None
            if pd.notna(row['telefono']):
                telefono_str = str(int(row['telefono'])) if isinstance(row['telefono'], float) else str(row['telefono'])
                telefono = re.sub(r'\D', '', telefono_str)
            
            ruta = str(row['ruta']).strip() if pd.notna(row['ruta']) else None
            
            # Verificar si existe el archivo
            existe = "✗ NO" if ruta else "- Sin ruta"
            if ruta and os.path.exists(ruta):
                existe = "✓ SÍ"
            elif ruta:
                ruta_rel = os.path.join(os.path.dirname(excel_path), ruta)
                if os.path.exists(ruta_rel):
                    existe = "✓ SÍ (relativa)"
            
            print(f"\n{idx + 1}. {nombre_completo}")
            print(f"   Tel: {telefono if telefono else 'Sin teléfono'}")
            print(f"   Ruta: {ruta[:50]}{'...' if ruta and len(ruta) > 50 else ''}")
            print(f"   Archivo existe: {existe}")
            
            if telefono:
                firmas_por_telefono[telefono] = ruta
            if nombre_completo:
                firmas_por_nombre[nombre_completo] = ruta
        
        print(f"\n{'='*70}")
        print("RESUMEN")
        print("="*70)
        print(f"✓ Total personas: {len(df)}")
        print(f"✓ Con teléfono: {len(firmas_por_telefono)}")
        print(f"✓ Todas las personas: {len(firmas_por_nombre)}")
        
        return True
        
    except Exception as e:
        print(f"\n✗ ERROR: {str(e)}")
        return False

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    
    excel_path = filedialog.askopenfilename(
        title="Seleccionar LibroFirmas.xlsx",
        filetypes=[("Excel", "*.xlsx *.xls")]
    )
    
    if excel_path:
        verificar_excel(excel_path)
    
    input("\nEnter para salir...")
