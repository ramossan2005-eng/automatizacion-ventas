"""
Genera archivos de ventas de ejemplo "desordenados", como los que manda
un cliente real: columnas con nombres distintos, espacios de más,
precios como texto, filas duplicadas y datos faltantes.

Uso:  python3 generar_datos.py
Crea la carpeta datos/ con tres archivos de Excel.
"""

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

random.seed(42)  # Para que siempre salgan los mismos datos

PRODUCTOS = {
    "Audífonos": 180_000,
    "Cargador portátil": 95_000,
    "Reloj inteligente": 420_000,
    "Teclado mecánico": 310_000,
    "Cámara de seguridad": 250_000,
}
CIUDADES = ["Bogotá", "Medellín", "Cali", "Barranquilla", "Bucaramanga"]
VENDEDORES = ["Ana", "Carlos", "Laura", "Andrés"]


def ensuciar_texto(texto):
    """Simula errores de digitación: mayúsculas raras y espacios de más."""
    opcion = random.random()
    if opcion < 0.15:
        return texto.upper()
    if opcion < 0.30:
        return texto.lower()
    if opcion < 0.45:
        return f"  {texto} "
    return texto


def crear_ventas(mes, n_filas):
    filas = []
    inicio = date(2026, mes, 1)
    for _ in range(n_filas):
        producto = random.choice(list(PRODUCTOS))
        cantidad = random.randint(1, 6)
        precio = PRODUCTOS[producto]
        filas.append({
            "fecha": inicio + timedelta(days=random.randint(0, 27)),
            "producto": ensuciar_texto(producto),
            "ciudad": ensuciar_texto(random.choice(CIUDADES)),
            "vendedor": random.choice(VENDEDORES),
            "cantidad": cantidad,
            "precio": precio,
        })
    df = pd.DataFrame(filas)

    # Algunos datos faltantes
    for _ in range(3):
        df.loc[random.randrange(len(df)), "cantidad"] = None
    df.loc[random.randrange(len(df)), "producto"] = None

    # Filas duplicadas (alguien pegó dos veces)
    df = pd.concat([df, df.sample(4, random_state=mes)], ignore_index=True)
    return df


def main():
    carpeta = Path("datos")
    carpeta.mkdir(exist_ok=True)

    # Enero: formato "normal"
    enero = crear_ventas(1, 60)
    enero.to_excel(carpeta / "ventas_enero.xlsx", index=False)

    # Febrero: otra persona usó nombres de columna distintos
    febrero = crear_ventas(2, 55).rename(columns={
        "fecha": "Fecha de venta",
        "producto": "PRODUCTO",
        "ciudad": "Ciudad ",
        "vendedor": "Vendedor",
        "cantidad": "Cant.",
        "precio": "Precio unitario",
    })
    febrero.to_excel(carpeta / "ventas_febrero.xlsx", index=False)

    # Marzo: precios escritos como texto con "$" y puntos de miles
    marzo = crear_ventas(3, 65)
    marzo["precio"] = marzo["precio"].apply(lambda p: f"${p:,.0f}".replace(",", "."))
    marzo = marzo.rename(columns={"precio": "Valor"})
    marzo.to_excel(carpeta / "ventas_marzo.xlsx", index=False)

    print("Listo: se crearon 3 archivos en la carpeta datos/")


if __name__ == "__main__":
    main()
