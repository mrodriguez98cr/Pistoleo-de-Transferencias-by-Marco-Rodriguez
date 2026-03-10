from app.models.transferencia_model import (
    buscar_producto_por_upc,
    obtener_cantidad_transferencia,
    guardar_upcs,
    guardar_transferencias,
    leer_upcs_desde_excel,
    leer_transferencias_desde_excel,
    limpiar_transferencias,
    eliminar_transferencias_en_db,
)


class PistoleoController:
    """
    Contiene toda la lógica de negocio del sistema.
    No sabe nada de tkinter — solo recibe datos y devuelve resultados.
    """

    def __init__(self):
        # Estado interno del sistema
        self.productos_por_transferencia: dict = {}  # nombre -> {codigo: {descripcion, cantidad}}
        self.pistoleo_por_producto: dict = {}        # codigo -> cantidad pistoleada

        # Contadores globales
        self.total_transferidas: float = 0
        self.total_pistoleadas: float = 0
        self.total_fuera: int = 0

        limpiar_transferencias()

    # ──────────────────────────────────────────────
    # UPC
    # ──────────────────────────────────────────────

    def cargar_upc(self, ruta: str) -> tuple[bool, str]:
        """
        Carga un archivo Excel de UPCs en la DB.
        Retorna (éxito: bool, mensaje: str)
        """
        datos = leer_upcs_desde_excel(ruta)
        if datos is None:
            return False, "El archivo debe tener exactamente 3 columnas."
        guardar_upcs(datos)
        return True, "UPC cargado exitosamente."

    def escanear_upc(self, upc: str) -> dict:
        """
        Procesa un UPC escaneado.
        Retorna un dict con el resultado del escaneo:
        {
            "encontrado": bool,
            "en_transferencia": bool,
            "codigo": str,
            "descripcion": str,
            "cantidad_transferencia": int,
            "cantidad_pistoleada": int,
            "mensaje": str,
        }
        """
        if not upc:
            return {"encontrado": False, "mensaje": "UPC vacío."}

        resultado_upc = buscar_producto_por_upc(upc)
        if not resultado_upc:
            return {"encontrado": False, "mensaje": "UPC INVÁLIDO."}

        codigo, descripcion = resultado_upc
        cantidad_transferencia = obtener_cantidad_transferencia(codigo)

        # Incrementar pistoleos
        self.pistoleo_por_producto[codigo] = self.pistoleo_por_producto.get(codigo, 0) + 1
        cantidad_pistoleada = self.pistoleo_por_producto[codigo]
        if cantidad_transferencia is not None and cantidad_pistoleada > cantidad_transferencia:
            return {
                "encontrado": True,
                "en_transferencia": True,
                "sobrante": True,
                "codigo": codigo,
                "descripcion": descripcion,
                "cantidad_transferencia": cantidad_transferencia,
                "cantidad_pistoleada": cantidad_pistoleada,
                "mensaje": f"⚠ Artículo de más: {codigo} lleva {cantidad_pistoleada} de {cantidad_transferencia}.",
            }

        self.total_pistoleadas += 1

        if cantidad_transferencia is not None:
            return {
                "encontrado": True,
                "en_transferencia": True,
                "codigo": codigo,
                "descripcion": descripcion,
                "cantidad_transferencia": cantidad_transferencia,
                "cantidad_pistoleada": cantidad_pistoleada,
                "mensaje": "",
            }
        else:
            self.total_fuera += 1
            return {
                "encontrado": True,
                "en_transferencia": False,
                "codigo": codigo,
                "descripcion": descripcion,
                "cantidad_transferencia": 0,
                "cantidad_pistoleada": cantidad_pistoleada,
                "mensaje": "El UPC no se encontró en la Transferencia.",
            }

    # ──────────────────────────────────────────────
    # Transferencias
    # ──────────────────────────────────────────────

    def cargar_transferencia(self, ruta: str, nombre: str) -> tuple:
        df = leer_transferencias_desde_excel(ruta)
        if df is None:
            return False, "El archivo debe tener las columnas 'Producto', 'Descripción' y 'Cantidad'.", [], nombre

        productos = {
            str(row["Producto"]).strip(): {
                "descripcion": str(row["Descripción"]).strip(),
                "cantidad": int(row["Cantidad"]),
            }
            for _, row in df.iterrows()
        }

        # Generar nombre único si ya existe
        nombre_unico = nombre
        contador = 1
        while nombre_unico in self.productos_por_transferencia:
            nombre_unico = f"{nombre} ({contador})"
            contador += 1

        
        self.productos_por_transferencia[nombre_unico] = productos
        
        
        self._sincronizar_db()
        lista = self._construir_lista_treeview()
        return True, "Transferencia cargada exitosamente.", lista, nombre_unico

    def eliminar_transferencia(self, nombre: str) -> list:
        """
        Elimina una transferencia del estado y recalcula.
        Retorna la lista actualizada para el Treeview.
        """
        self.productos_por_transferencia.pop(nombre, None)
        self._sincronizar_db()
        return self._construir_lista_treeview()

    def _sincronizar_db(self):
        """Recalcula el acumulado de todas las transferencias y lo guarda en la DB."""
        eliminar_transferencias_en_db()
        acumulado = self._acumular_transferencias()
        lista_db = [(cod, info["descripcion"], info["cantidad"]) for cod, info in acumulado.items()]
        guardar_transferencias(lista_db)
        self.total_transferidas = sum(info["cantidad"] for info in acumulado.values())

    def _acumular_transferencias(self) -> dict:
        """Combina todas las transferencias activas en un solo dict."""
        acumulado = {}
        for productos in self.productos_por_transferencia.values():
            for codigo, info in productos.items():
                if codigo not in acumulado:
                    acumulado[codigo] = {"descripcion": info["descripcion"], "cantidad": info["cantidad"]}
                else:
                    acumulado[codigo]["cantidad"] += info["cantidad"]
        return acumulado

    def _construir_lista_treeview(self) -> list:
        """
        Construye la lista de filas para el Treeview, combinando transferencias y pistoleos.
        Retorna [(codigo, descripcion, cantidad_transferencia, cantidad_pistoleada), ...]
        """
        acumulado = self._acumular_transferencias()
        filas = []

        for codigo, info in acumulado.items():
            pistoleada = self.pistoleo_por_producto.get(codigo, 0)
            filas.append((codigo, info["descripcion"], info["cantidad"], pistoleada))

        # Productos pistoleados que ya no están en ninguna transferencia
        codigos_en_transferencia = set(acumulado.keys())
        for codigo, pistoleada in self.pistoleo_por_producto.items():
            if codigo not in codigos_en_transferencia and pistoleada > 0:
                filas.append((codigo, "Producto pistoleado", 0, pistoleada))

        return filas

    # ──────────────────────────────────────────────
    # Contadores
    # ──────────────────────────────────────────────

    def get_contadores(self) -> dict:
        """Retorna el estado actual de todos los contadores."""
        diferencia = self.total_transferidas - self.total_pistoleadas
        return {
            "transferidas": self.total_transferidas,
            "pistoleadas": self.total_pistoleadas,
            "fuera": self.total_fuera,
            "diferencia": diferencia,
        }

    # ──────────────────────────────────────────────
    # Limpiar todo
    # ──────────────────────────────────────────────

    def limpiar_todo(self):
        """Resetea el estado completo del sistema."""
        self.productos_por_transferencia.clear()
        self.pistoleo_por_producto.clear()
        self.total_transferidas = 0
        self.total_pistoleadas = 0
        self.total_fuera = 0
        limpiar_transferencias()