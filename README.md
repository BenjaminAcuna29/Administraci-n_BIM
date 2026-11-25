# Dashboard de presupuesto de construcción

## Ejecución en Windows

Desde esta carpeta, abre PowerShell y ejecuta:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
streamlit run app.py
```

Si `py` no está disponible, instala Python 3.10 o superior y repite desde la creación del entorno virtual. También puedes iniciar la aplicación directamente con `python -m streamlit run app.py`.

La aplicación abre el libro `.xlsx` presente junto a `app.py` si no se carga otro archivo en el panel lateral. Los informes Excel y HTML se descargan desde la pestaña **Exportar**.

## Criterios de análisis

- Detecta hojas de presupuesto, catálogos y APU a partir de sus encabezados y códigos.
- Vincula partidas con APU por código exacto; no completa vínculos por semejanza.
- Calcula la demanda estimada como `cantidad de partida × rendimiento APU × factor de pérdida`, de acuerdo con las fórmulas del libro (por ejemplo, `H13 = G13 * F13 * D13`).
- Mantiene los totales separados por unidad, y exige ingresar una superficie total construida validada para calcular el costo por m².
- Cuando no detecta un total inequívoco de gastos generales, solicita que el usuario ingrese el monto validado.
- Muestra alertas de fórmulas, datos faltantes, duplicidad de códigos y diferencias entre presupuesto, APU y catálogos.

Los resultados son una herramienta de revisión y deben contrastarse con el libro fuente antes de usarlos para compras, pagos o decisiones contractuales.
