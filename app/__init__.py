import os
import sqlite3
import pandas as pd


# ──────────────────────────────────────────────
# Conexión
# ──────────────────────────────────────────────

def connect_db():
    """Conecta a la base de datos, creando el archivo si no existe."""
    db_path = os.path.join("data", "transferencias", "UPC.db")
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    return sqlite3.connect(db_path)


# ──────────────────────────────────────────────
# UPC
# ──────────────────────────────────────────────

def guardar_upcs(lista_upcs: list):
    """
    Reemplaza todos los UPCs en la DB.
    lista_upcs: [(upc, codigo_producto, descripcion), ...]
    """
    conn = connect_db()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS UPC (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                upc TEXT NOT NULL,
                codigo_producto TEXT NOT NULL,
                descripcion TEXT NOT NULL
            )
        """)
        cursor.execute("DELETE FROM UPC")
        cursor.executemany(
            "INSERT INTO UPC (upc, codigo_producto, descripcion) VALUES (?, ?, ?)",
            lista_upcs,
        )
        conn.commit()
    finally:
        conn.close()


def buscar_producto_por_upc(upc: str):
    """
    Busca un producto por UPC (acepta prefijos).
    Retorna (codigo_producto, descripcion) o None si no existe.
    """
    conn = connect_db()
    try:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT codigo_producto, descripcion FROM UPC WHERE upc LIKE ?",
            (upc + "%",),
        )
        return cursor.fetchone()
    finally:
        conn.close()


# ──────────────────────────────────────────────
# Transferencias
# ──────────────────────────────────────────────

def obtener_cantidad_transferencia(codigo_producto: str):
    """
    Retorna la cantidad en transferencia de un producto, o None si no está.
    """
    conn = connect_db()
    try:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT cantidad FROM Transferencias WHERE codigo_producto = ?",
            (codigo_producto,),
        )
        resultado = cursor.fetchone()
        return resultado[0] if resultado else None
    finally:
        conn.close()


def guardar_transferencias(lista: list):
    """
    Inserta o actualiza productos en la tabla Transferencias.
    lista: [(codigo_producto, descripcion, cantidad), ...]
    """
    conn = connect_db()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS Transferencias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_producto TEXT NOT NULL,
                descripcion TEXT NOT NULL,
                cantidad INTEGER NOT NULL
            )
        """)
        for codigo, desc, cant in lista:
            cursor.execute(
                "SELECT cantidad FROM Transferencias WHERE codigo_producto = ? AND descripcion = ?",
                (codigo, desc),
            )
            existente = cursor.fetchone()
            if existente:
                cursor.execute(
                    "UPDATE Transferencias SET cantidad = ? WHERE codigo_producto = ? AND descripcion = ?",
                    (existente[0] + cant, codigo, desc),
                )
            else:
                cursor.execute(
                    "INSERT INTO Transferencias (codigo_producto, descripcion, cantidad) VALUES (?, ?, ?)",
                    (codigo, desc, cant),
                )
        conn.commit()
    finally:
        conn.close()


def obtener_todas_transferencias() -> list:
    """
    Retorna todos los productos en la tabla Transferencias.
    [(codigo_producto, descripcion, cantidad), ...]
    """
    conn = connect_db()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT codigo_producto, descripcion, cantidad FROM Transferencias")
        return cursor.fetchall()
    finally:
        conn.close()


def limpiar_transferencias():
    """Elimina y recrea la tabla de Transferencias."""
    conn = connect_db()
    try:
        cursor = conn.cursor()
        cursor.execute("DROP TABLE IF EXISTS Transferencias")
        cursor.execute("""
            CREATE TABLE Transferencias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_producto TEXT NOT NULL,
                descripcion TEXT NOT NULL,
                cantidad INTEGER NOT NULL
            )
        """)
        conn.commit()
    finally:
        conn.close()


def eliminar_transferencias_en_db():
    """Borra todos los registros de Transferencias sin eliminar la tabla."""
    conn = connect_db()
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM Transferencias")
        conn.commit()
    finally:
        conn.close()


# ──────────────────────────────────────────────
# Lectura de Excel (solo parseo, sin UI)
# ──────────────────────────────────────────────

def leer_upcs_desde_excel(ruta: str):
    """
    Lee el archivo Excel de UPCs.
    Retorna [(upc, codigo_producto, descripcion), ...] o None si hay error.
    """
    df = pd.read_excel(ruta, header=None, engine="openpyxl")
    if df.shape[1] != 3:
        return None
    return [(str(row[2]), str(row[0]), str(row[1])) for _, row in df.iterrows()]


def leer_transferencias_desde_excel(ruta: str):
    """
    Lee y consolida el archivo Excel de transferencias.
    Retorna un DataFrame con columnas [Producto, Descripcion, Cantidad] o None si hay error.
    """
    df = pd.read_excel(ruta, engine="openpyxl")
    columnas = {"Producto", "Descripción", "Cantidad"}
    if not columnas.issubset(df.columns):
        return None
    return df.groupby(["Producto", "Descripción"], as_index=False).agg({"Cantidad": "sum"})