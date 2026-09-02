# -*- coding: utf-8 -*-
"""
Comparar precios competencia — Kbas Office
Pega URLs de categorías de tiendas competidoras (con la etiqueta que
elijas, p.ej. "Capazos") y compara precio medio/mínimo/máximo entre
ellas. Detecta Shopify automáticamente; para otras plataformas usa
un método genérico de lectura de la página.
"""

import io
import re
import time

import pandas as pd
import requests
import streamlit as st

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; KbasPriceCheck/1.0)"}
PATRON_PRECIO = re.compile(r"(\d{1,4}[.,]\d{2})\s?€")


def check_password():
    """Pide contraseña antes de mostrar nada. Se guarda en Secrets."""
    def password_entered():
        if st.session_state.get("password") == st.secrets.get("dashboard_password"):
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct"):
        return True

    st.title("Comparar precios de la competencia")
    st.text_input("Contraseña", type="password", on_change=password_entered, key="password")
    if st.session_state.get("password_correct") is False:
        st.error("Contraseña incorrecta.")
    return False


def limpiar_precio(s):
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def extraer_shopify(url_coleccion):
    """Si la URL es de una colección Shopify, devuelve sus productos
    con nombre y precio. Si no es Shopify o falla, devuelve None."""
    m = re.search(r"(https?://[^/]+)/collections/([^/?]+)", url_coleccion)
    if not m:
        return None
    dominio, handle = m.group(1), m.group(2)
    productos = []
    for pagina in range(1, 6):  # hasta 5 páginas = 1250 productos
        try:
            r = requests.get(
                f"{dominio}/collections/{handle}/products.json",
                params={"limit": 250, "page": pagina},
                headers=HEADERS, timeout=15,
            )
            data = r.json()
        except Exception:
            return None if pagina == 1 else productos
        items = data.get("products", [])
        if not items:
            break
        for p in items:
            precios = []
            for v in p.get("variants", []):
                if v.get("price"):
                    try:
                        precios.append(float(v["price"]))  # Shopify: "69.50" -> 69.5, directo
                    except (ValueError, TypeError):
                        pass
            if precios:
                productos.append({"nombre": p.get("title"), "precio": min(precios)})
        time.sleep(0.5)
    return productos if productos else None


def extraer_generico(url_pagina):
    """Método de respaldo para tiendas que no son Shopify: junta todo el
    texto visible de la página y, por cada precio encontrado (NN,NN €),
    toma el texto inmediatamente anterior como nombre aproximado del
    producto. Menos preciso que Shopify, solo una página (sin paginación),
    pero funciona razonablemente en catálogos normales."""
    try:
        r = requests.get(url_pagina, headers=HEADERS, timeout=15)
    except Exception as e:
        raise ValueError(f"No he podido acceder a la página: {e}")

    from html.parser import HTMLParser

    class Extractor(HTMLParser):
        def __init__(self):
            super().__init__()
            self.partes = []

        def handle_data(self, data):
            data = data.strip()
            if data:
                self.partes.append(data)

    parser = Extractor()
    parser.feed(r.text)
    texto_completo = " § ".join(parser.partes)  # separador para no pegar palabras de tags distintos

    productos = []
    for m in PATRON_PRECIO.finditer(texto_completo):
        precio = limpiar_precio(m.group(1))
        antes = texto_completo[max(0, m.start() - 80):m.start()]
        # nombre = último fragmento entre separadores § antes del precio
        fragmentos = [f.strip() for f in antes.split("§") if f.strip()]
        nombre = fragmentos[-1] if fragmentos else ""
        if precio and nombre and len(nombre) > 3 and not nombre.replace(",", "").replace(".", "").isdigit():
            productos.append({"nombre": nombre[:70], "precio": precio})
    return productos


def analizar_url(url, etiqueta_manual=None):
    productos = extraer_shopify(url)
    origen = "Shopify"
    if productos is None:
        productos = extraer_generico(url)
        origen = "Genérico (puede ser menos preciso)"
    return productos, origen


# ---------- Interfaz ----------
st.set_page_config(page_title="Comparar precios competencia · Kbas Office", page_icon="⚖️", layout="wide")

if not check_password():
    st.stop()

st.title("Comparar precios de la competencia")
st.caption(
    "Pega URLs de páginas de categoría de competidores (una por línea, formato: "
    "Etiqueta | Competidor | URL) y compara precio medio, mínimo y máximo."
)

with st.expander("ℹ️ Cómo escribir las líneas y qué esperar"):
    st.markdown("""
Cada línea: `Etiqueta | Competidor | URL de la categoría`

Ejemplo:
```
Capazos | Abbacino | https://abbacino.es/collections/capazos
Capazos | Matties Bags | https://mattiesbags.com/collections/capazos
Bandoleras | Abbacino | https://abbacino.es/collections/bandoleras
```

Usa la **misma etiqueta** para categorías que quieras comparar entre sí,
aunque cada tienda las llame distinto (p.ej. "Bolsos de playa" vs "Capazos"
puedes etiquetarlas igual como "Capazos" si son equivalentes).

Las tiendas Shopify (Abbacino, Matties Bags) dan datos precisos y completos.
Otras plataformas (Simó Sastre, Lola Casademunt...) usan un método más
aproximado — revisa los resultados antes de sacar conclusiones fuertes.
    """)

texto = st.text_area(
    "URLs a analizar",
    height=150,
    placeholder="Capazos | Abbacino | https://abbacino.es/collections/capazos\nCapazos | Matties Bags | https://mattiesbags.com/collections/capazos",
)

st.divider()
st.subheader("Vuestros propios precios (opcional, para comparar)")
excel_volum = st.file_uploader("Excel de Volum/Kbas (con columnas de referencia y precio)", type=["xlsx", "xlsm"])
etiquetas_propias = {}
if excel_volum:
    try:
        df_propio = pd.read_excel(excel_volum)
        col_precio_propio = st.selectbox("Columna de precio", df_propio.columns)
        col_desc_propio = st.selectbox("Columna de descripción (para buscar por palabra clave)", df_propio.columns)
    except Exception as e:
        st.error(f"No puedo leer el Excel: {e}")
        df_propio = None

if st.button("Comparar precios", type="primary"):
    lineas = [l.strip() for l in texto.strip().split("\n") if l.strip()]
    if not lineas:
        st.error("Pega al menos una línea con Etiqueta | Competidor | URL.")
        st.stop()

    resultados = []
    barra = st.progress(0.0)
    for i, linea in enumerate(lineas):
        partes = [p.strip() for p in linea.split("|")]
        if len(partes) != 3:
            st.warning(f"Línea ignorada (formato incorrecto): {linea}")
            continue
        etiqueta, competidor, url = partes
        try:
            productos, origen = analizar_url(url)
        except ValueError as e:
            st.warning(f"{competidor} ({etiqueta}): {e}")
            continue

        if not productos:
            st.warning(f"{competidor} ({etiqueta}): no encontré productos con precio en esa página.")
            continue

        precios = [p["precio"] for p in productos]
        resultados.append({
            "Etiqueta": etiqueta, "Competidor": competidor,
            "Nº productos": len(precios),
            "Precio medio": round(sum(precios) / len(precios), 2),
            "Precio mín": min(precios), "Precio máx": max(precios),
            "Método": origen,
        })
        barra.progress((i + 1) / len(lineas))
    barra.empty()

    # Añadir la propia comparación de Volum/Kbas, si se ha subido Excel
    if excel_volum and 'df_propio' in dir():
        etiquetas_usadas = sorted(set(r["Etiqueta"] for r in resultados))
        st.caption("Para cada etiqueta, escribe la palabra clave que identifica esos productos en vuestro Excel:")
        for etq in etiquetas_usadas:
            palabra = st.text_input(f"Palabra clave para '{etq}' en vuestra descripción", key=f"kw_{etq}")
            if palabra:
                filtro = df_propio[col_desc_propio].astype(str).str.contains(palabra, case=False, na=False)
                precios_propios = pd.to_numeric(df_propio.loc[filtro, col_precio_propio], errors="coerce").dropna()
                if len(precios_propios):
                    resultados.append({
                        "Etiqueta": etq, "Competidor": "VOLUM/KBAS (vosotros)",
                        "Nº productos": len(precios_propios),
                        "Precio medio": round(precios_propios.mean(), 2),
                        "Precio mín": round(precios_propios.min(), 2),
                        "Precio máx": round(precios_propios.max(), 2),
                        "Método": "Vuestro Excel",
                    })

    if resultados:
        df_resultado = pd.DataFrame(resultados)
        st.success(f"Comparativa lista: {len(df_resultado)} filas.")
        st.dataframe(df_resultado, use_container_width=True, hide_index=True)

        salida = io.BytesIO()
        df_resultado.to_excel(salida, index=False)
        salida.seek(0)
        st.download_button(
            "⬇ Descargar comparativa.xlsx", data=salida,
            file_name="comparativa_precios_competencia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error("No he podido sacar datos de ninguna de las URLs.")
