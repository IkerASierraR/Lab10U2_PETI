"""AR07 · Capturas de las salidas de consola y vistas previas de anexos — Municipalidad Distrital de Alto Selva (Taller 10, SI-886).

Convierte cada salida registrada en docs/evidencias/S10/salidas/*.txt en una
imagen PNG con estilo de terminal (docs/evidencias/S10/capturas/), para que el
informe muestre la ejecución y enlace el archivo de texto original en GitHub.

Uso (desde la raíz del repositorio):
    python 05_arquitectura/AR07_capturas.py
    
"""


import textwrap
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

BASE = Path(__file__).resolve().parent
S10 = BASE.parent / "docs" / "evidencias" / "S10"
SALIDAS = S10 / "salidas"
CAPTURAS = S10 / "capturas"
CAPTURAS.mkdir(parents=True, exist_ok=True)

PROMPT = "$ "
CAPTURAS_A_GENERAR = [
    ("salida_AR02_estructura.txt", "python 05_arquitectura/AR02_estructura.py", "captura_AR02_estructura.png"),
    ("salida_AR05_modelo.txt", "python 05_arquitectura/AR05_modelo_archi.py", "captura_AR05_modelo.png"),
    ("salida_AR03_brechas.txt", "python 05_arquitectura/AR03_brechas.py", "captura_AR03_brechas.png"),
    ("salida_AR04_validacion_inicial.txt", "python 05_arquitectura/AR04_valida.py   # primera ejecución", "captura_AR04_validacion_inicial.png"),
    ("salida_AR04_validacion.txt", "python 05_arquitectura/AR04_valida.py   # después de corregir", "captura_AR04_validacion.png"),
    ("salida_git_log.txt", "git log --oneline --decorate", "captura_git_log.png"),
]


def color(linea):
    if "[FALLA]" in linea or "[ALERTA]" in linea:
        return "#FF6B6B"
    if "[OK]" in linea or "todas las comprobaciones pasaron" in linea:
        return "#7CD67C"
    if linea.startswith("=") or linea.isupper():
        return "#6CB6FF"
    if "->" in linea or "→" in linea:
        return "#F5D76E"
    return "#D4D4D4"


def captura(txt, comando, png):
    lineas = (SALIDAS / txt).read_text(encoding="utf-8").rstrip().splitlines()
    lineas = [PROMPT + comando] + lineas
    ancho = max(len(l) for l in lineas)
    alto = len(lineas)
    fig = plt.figure(figsize=(max(8, ancho * 0.072), 0.6 + alto * 0.165))
    fig.patch.set_facecolor("#1E1E1E")
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_facecolor("#1E1E1E")
    ax.axis("off")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, alto + 2)
    ax.add_patch(plt.Rectangle((0, alto + 1), 1, 1, color="#333333"))
    for k, c in enumerate(("#FF5F56", "#FFBD2E", "#27C93F")):
        ax.text(0.012 + k * 0.014, alto + 1.5, "●", ha="center", va="center", color=c, fontsize=10)
    ax.text(0.5, alto + 1.5, "Terminal · " + txt, ha="center", va="center", color="#BBBBBB", fontsize=8,
            family="monospace")
    for n, l in enumerate(lineas):
        ax.text(0.01, alto - n + 0.3, l, ha="left", va="center", fontsize=8.2, family="monospace",
                color="#4EC9B0" if n == 0 else color(l))
    fig.savefig(CAPTURAS / png, dpi=150, facecolor=fig.get_facecolor())
    plt.close(fig)
    print(f"  docs/evidencias/S10/capturas/{png}")


def vista_previa(xlsx, hoja, columnas, anchos, png, titulo, filas=None):
    """Imagen de la hoja de un anexo .xlsx, para insertarla en el informe."""
    df = pd.read_excel(S10 / xlsx, sheet_name=hoja, header=3)[columnas].dropna(how="all")
    if filas:
        df = df.head(filas)
    celdas = [[textwrap.fill(str(v), max(int(a * 1.25), 3), break_long_words=False) for v, a in zip(fila, anchos)] for fila in df.itertuples(index=False)]
    alto = sum(max(c.count("\n") + 1 for c in fila) for fila in celdas)
    fig = plt.figure(figsize=(10, 1.0 + alto * 0.21))
    ax = fig.add_axes([0.02, 0.01, 0.955, 0.9])
    ax.axis("off")
    fig.suptitle(titulo, fontsize=12, fontweight="bold", color="#1F2D5C", x=0.01, ha="left")
    t = ax.table(cellText=celdas, colLabels=[textwrap.shorten(c, int(a * 1.5) + 3, placeholder="…") for c, a in zip(columnas, anchos)],
                 loc="upper left", cellLoc="left",
                 colWidths=[a / sum(anchos) for a in anchos])
    t.auto_set_font_size(False)
    t.set_fontsize(8)
    for (r, c), cel in t.get_celld().items():
        n = 1 if r == 0 else max(v.count("\n") + 1 for v in celdas[r - 1])
        cel.set_height(0.95 / (alto + 1) * max(n, 1))
        if r == 0:
            cel.set_facecolor("#1F2D5C")
            cel.get_text().set_color("white")
            cel.get_text().set_fontweight("bold")
    fig.savefig(CAPTURAS / png, dpi=150)
    plt.close(fig)
    print(f"  docs/evidencias/S10/capturas/{png}")


def vista_previa_pdf(pdf, png):
    try:
        import pymupdf
    except ImportError:
        print("  [AVISO] pymupdf no instalado: sin vista previa del anexo F")
        return
    doc = pymupdf.open(S10 / pdf)
    doc[0].get_pixmap(dpi=110).save(CAPTURAS / png)
    print(f"  docs/evidencias/S10/capturas/{png}")


if __name__ == "__main__":
    print("AR07 · Capturas generadas:")
    for txt, comando, png in CAPTURAS_A_GENERAR:
        if (SALIDAS / txt).exists():
            captura(txt, comando, png)
        else:
            print(f"  [AVISO] falta {txt}")
    print("AR07 · Vistas previas de los anexos:")
    vista_previa("anexo_C_estructura_informacion.xlsx", "Entidades",
                 ["Entidad", "N.º de registros por sistema", "¿Fuente autoritativa?", "Dueño del dato (cargo)", "Discrepancia observada", "Clasificación", "¿Contiene datos personales?"],
                 [11, 15, 9, 13, 30, 10, 11], "preview_anexo_C.png", "Anexo C · anexo_C_estructura_informacion.xlsx · hoja Entidades (vista previa)")
    vista_previa("anexo_E_analisis_brechas.xlsx", "Brechas", ["id", "capa", "brecha", "tipo", "esfuerzo_sem", "dependencias", "paquete"],
                 [4, 8, 36, 9, 5, 12, 5], "preview_anexo_E.png", "Anexo E · anexo_E_analisis_brechas.xlsx · hoja Brechas (vista previa)")
    vista_previa_pdf("anexo_F_seccion_5.pdf", "preview_anexo_F.png")
