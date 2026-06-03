"""
Paso 1: Extractor de datos desde PDF de cotización.
Lee el PDF y muestra todo lo que extrae, para verificación antes de llenar el Excel.
"""

import re
import pdfplumber


def normalizar(texto: str) -> str:
    """Corrige doble-encoding (mojibake) de caracteres especiales en PDFs."""
    if not texto:
        return texto
    # Detecta si el texto viene con doble UTF-8 y lo corrige
    try:
        corregido = texto.encode("latin-1").decode("utf-8")
        return corregido
    except (UnicodeEncodeError, UnicodeDecodeError):
        return texto


def extraer_cotizacion(ruta_pdf: str) -> dict:
    datos = {
        "nro_cotizacion": "",
        "fecha_cotizacion": "",
        "razon_social": "",
        "ruc": "",
        "direccion": "",
        "contacto": "",
        "email_contacto": "",
        "telefono_contacto": "",
        "responsable_cotizacion": "",
        "n_muestras": "",
        "informacion_adicional": "",
        "matriz": "",
        "parametros": [],   # lista de {"nombre": str, "cantidad": int}
    }

    with pdfplumber.open(ruta_pdf) as pdf:
        # --- Página 1: cabecera y datos del cliente ---
        texto_p1 = pdf.pages[0].extract_text() or ""
        lineas_p1 = [l.strip() for l in texto_p1.splitlines() if l.strip()]

        # Etiquetas que marcan el FIN de la razón social (inicio de otro campo)
        LABEL_SIGUIENTE = re.compile(
            r"^(Direcci.n|Contacto|E-?mail|Documento|Realizado|Informaci.n|"
            r"N.\s*Register|RUC|Tel.fono|Unidad|FR-|PACIFIC|www\.)",
            re.IGNORECASE,
        )

        for i, linea in enumerate(lineas_p1):
            if re.search(r"COTIZACI.N N", linea, re.IGNORECASE):
                # Acepta "2026-0103-3", "2026-0410" y "2025-11656" (2do segmento de 3-6 dígitos)
                m = re.search(r"(\d{4}-\d{3,6}(?:-\d+)?)", linea)
                if m:
                    datos["nro_cotizacion"] = m.group(1)
                m2 = re.search(r"Fecha[:\s]+(\d{4}-\d{2}-\d{2})", linea)
                if m2:
                    datos["fecha_cotizacion"] = m2.group(1)

            # IMPORTANTE: anclar al inicio con ":" y solo la PRIMERA vez,
            # para no confundir con notas legales que mencionan "razón social".
            elif re.match(r"\s*Raz.n Social\s*:", linea, re.IGNORECASE) and not datos["razon_social"]:
                m = re.search(r"Raz.n Social\s*:\s*(.+?)(?:N.\s*Register|RUC|$)", linea, re.IGNORECASE)
                if m:
                    datos["razon_social"] = m.group(1).strip().rstrip("/").strip()
                    datos["_razon_social_idx"] = i
                m2 = re.search(r"RUC[:\s]+(\d+)", linea)
                if m2:
                    datos["ruc"] = m2.group(1)

            elif re.search(r"^Direcci.n[:\s]", linea, re.IGNORECASE):
                datos["direccion"] = linea.split(":", 1)[-1].strip()

            elif re.search(r"^Contacto[:\s]", linea, re.IGNORECASE):
                datos["contacto"] = linea.split(":", 1)[-1].strip()

            elif re.search(r"^E-mail[:\s]", linea, re.IGNORECASE) and not datos["email_contacto"]:
                partes = linea.split()
                for p in partes:
                    if "@" in p:
                        datos["email_contacto"] = p
                        break
                m = re.search(r"Tel.fono[:\s]+([\+\d\s]+)", linea)
                if m:
                    datos["telefono_contacto"] = m.group(1).strip()

            elif re.search(r"Realizado por", linea, re.IGNORECASE):
                m = re.search(r"Realizado por[:\s]+(.+?)(?:Unidad|$)", linea, re.IGNORECASE)
                if m:
                    datos["responsable_cotizacion"] = m.group(1).strip()

        # La razón social puede continuar en las líneas siguientes (ej: "ANONIMA
        # CERRADA", "S.A.C."). Capturar todo hasta llegar a otra etiqueta (Dirección, etc.)
        idx = datos.pop("_razon_social_idx", None)
        if idx is not None and datos["razon_social"]:
            extras = []
            for j in range(idx + 1, len(lineas_p1)):
                sig = lineas_p1[j].strip()
                if not sig:
                    continue
                if LABEL_SIGUIENTE.match(sig):
                    break
                extras.append(sig)
                if len(extras) >= 2:   # máximo 2 líneas de continuación, por seguridad
                    break
            if extras:
                datos["razon_social"] = (datos["razon_social"] + " " + " ".join(extras)).strip()

        # Dirección puede ocupar múltiples líneas — tomamos las siguientes a "Dirección:"
        capturando_dir = False
        dir_lineas = []
        for linea in lineas_p1:
            if re.search(r"^Direcci.n[:\s]", linea, re.IGNORECASE):
                capturando_dir = True
                dir_lineas.append(linea.split(":", 1)[-1].strip())
            elif capturando_dir:
                if re.search(r"^(Contacto|E-mail|Documento|Realizado|Informaci)", linea, re.IGNORECASE):
                    capturando_dir = False
                else:
                    dir_lineas.append(linea)
        if dir_lineas:
            datos["direccion"] = " ".join(dir_lineas).strip()

        # Información adicional (vertientes, productos terminados)
        # Estructura en el PDF: "Información X MUESTRAS" en una línea, "Adicional: ..." en la siguiente
        cap_info = False
        info_lineas = []
        n_muestras = ""
        for linea in lineas_p1:
            if re.search(r"^Informaci.n", linea, re.IGNORECASE):
                # Extraer cantidad de muestras de esta línea
                m = re.search(r"(\d+)\s+MUESTRAS", linea, re.IGNORECASE)
                if m:
                    n_muestras = m.group(1)
                cap_info = True
            elif cap_info and re.search(r"^Adicional[:\s]", linea, re.IGNORECASE):
                resto = re.sub(r"^Adicional[:\s]*", "", linea, flags=re.IGNORECASE).strip()
                if resto:
                    info_lineas.append(resto)
            elif cap_info:
                if re.search(r"^(Se deber|ENSAYOS|Tiempo|FR-|PACIFIC)", linea, re.IGNORECASE):
                    cap_info = False
                elif re.search(r"^(Raz.n|Direcci|Contacto|E-mail|Documento|Realizado)", linea, re.IGNORECASE):
                    cap_info = False
                elif info_lineas:  # solo agregar si ya capturamos algo
                    info_lineas.append(linea)
        # Si no se encontró "X MUESTRAS" en el encabezado, inferir del valor más común en Cant.
        datos["n_muestras"] = n_muestras

        # Solo guardar info adicional si contiene datos de muestreo (no condiciones de pago)
        info_texto = "\n".join(info_lineas).strip()
        keywords_muestreo = ["vertiente", "puntos de muestreo", "punto de muestreo",
                             "estación de muestreo", "producto terminado", "n° muestras",
                             "nro. muestras", "pozo", "efluente", "canal de"]
        if any(kw in info_texto.lower() for kw in keywords_muestreo):
            datos["informacion_adicional"] = info_texto
        else:
            datos["informacion_adicional"] = ""

        # Matriz: buscar encabezado de sección justo antes de la tabla de ensayos
        for linea in lineas_p1:
            if re.search(r"^AGUA|^AIRE|^SUELO|^SEDIMENTO", linea, re.IGNORECASE):
                if not re.search(r"(mineral|residual|potable|subterranea|superficial|mar|salobre)", linea, re.IGNORECASE):
                    datos["matriz"] = linea.strip()
                    break
                else:
                    datos["matriz"] = linea.strip()
                    break

        # --- Todas las páginas: extraer parámetros desde tablas ---
        parametros_vistos = set()
        for i, pagina in enumerate(pdf.pages):
            tablas = pagina.extract_tables()
            for tabla in tablas:
                if not tabla:
                    continue
                # Buscar filas con datos de ensayo (columnas: Prueba, LCM, Unidad, Precio, Cant, ...)
                for fila in tabla:
                    if not fila or len(fila) < 5:
                        continue
                    nombre_raw = fila[0] or ""
                    cant_raw = fila[4] if len(fila) > 4 else ""

                    # Saltar encabezados
                    if re.search(r"prueba a realizar|ensayo|LCM", nombre_raw, re.IGNORECASE):
                        continue
                    if not nombre_raw.strip():
                        continue

                    # Extraer solo la primera línea del nombre (antes de "Metodología")
                    nombre = nombre_raw.split("\n")[0].strip()
                    nombre = re.sub(r"\s+", " ", nombre)
                    nombre = normalizar(nombre)
                    if not nombre or nombre in parametros_vistos:
                        continue

                    # Extraer cantidad
                    try:
                        cantidad = int(str(cant_raw).strip().split("\n")[0])
                    except (ValueError, AttributeError):
                        cantidad = 0

                    parametros_vistos.add(nombre)
                    datos["parametros"].append({"nombre": nombre, "cantidad": cantidad})

    # Si n_muestras sigue vacío, inferir del valor más frecuente en la columna Cant.
    if not datos["n_muestras"] and datos["parametros"]:
        cantidades = [p["cantidad"] for p in datos["parametros"] if p["cantidad"] > 0]
        if cantidades:
            datos["n_muestras"] = str(max(set(cantidades), key=cantidades.count))

    return datos


def mostrar_datos(datos: dict):
    print("=" * 60)
    print("DATOS EXTRAÍDOS DE LA COTIZACIÓN")
    print("=" * 60)
    print(f"  N° Cotización      : {datos['nro_cotizacion']}")
    print(f"  Fecha cotización   : {datos['fecha_cotizacion']}")
    print(f"  Razón Social       : {datos['razon_social']}")
    print(f"  RUC                : {datos['ruc']}")
    print(f"  Dirección          : {datos['direccion']}")
    print(f"  Contacto           : {datos['contacto']}")
    print(f"  Email contacto     : {datos['email_contacto']}")
    print(f"  Teléfono contacto  : {datos['telefono_contacto']}")
    print(f"  Resp. Cotización   : {datos['responsable_cotizacion']}")
    print(f"  Matriz             : {datos['matriz']}")
    print(f"  N° Muestras        : {datos['n_muestras']}")
    print(f"  Info adicional     :\n    {datos['informacion_adicional'].replace(chr(10), chr(10)+'    ')}")
    print(f"\n  PARÁMETROS ({len(datos['parametros'])} encontrados):")
    for i, p in enumerate(datos["parametros"], 1):
        print(f"    {i:>3}. [{p['cantidad']:>2} pts] {p['nombre']}")
    print("=" * 60)


if __name__ == "__main__":
    import sys
    ruta = sys.argv[1] if len(sys.argv) > 1 else "cotizacion 2026-0103.pdf"
    datos = extraer_cotizacion(ruta)
    mostrar_datos(datos)
