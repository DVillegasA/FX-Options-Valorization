# Valorización automática de Opciones FX (USD/CLP)

Reemplaza el proceso manual diario (notebook de extracción + pegado de CSV en el
maestro Excel + altas/vencimientos a mano) por un proceso en Python. La
valorización **coincide al centavo** con el modelo VBA original (validado sobre
las 64 operaciones del 2026-07-14: error máximo 0,000000 CLP en MtM, forwards,
vols y strikes).

## Estructura

```
fx-options-valorizacion/
├── fx_options_valuation.py   # Motor: curvas, superficie de vol, Black-Scholes FX, MtM
├── load_market.py            # Lee 'Grid Vol Bloomberg Creasys.xlsx' (reemplaza el notebook)
├── dateparse.py              # Parser de fechas robusto (incl. formato español)
├── daily_process.py          # Orquestador diario (único punto de entrada)
├── requirements.txt          # Dependencias de producción (openpyxl, Pillow)
├── requirements-dev.txt      # Dependencias de desarrollo (+ pytest)
├── assets/                   # logo.png para las hojas de salida (opcional)
├── tests/                    # pruebas unitarias y de regresión (pytest)
│   └── data/                 # datos de mercado de prueba (sí se versionan)
├── .github/workflows/        # CI: corre las pruebas en cada push
├── data/                     # (ignorado por git) los 2 adjuntos del correo
├── salidas/                  # (ignorado por git) resultados generados
└── estado/                   # (ignorado por git) foto de la cartera vigente
                               #   (seguimiento de cambios entre corridas)
```

Solo se ejecuta `daily_process.py`; los otros módulos los usa por dentro.

## Requisitos

- Python 3.9 o superior.
- Dependencias externas: `openpyxl` y `Pillow` (para el logo); el resto es librería estándar.

## Instalación (entorno virtual)

Desde la carpeta del proyecto:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1        # PowerShell  (CMD: .\.venv\Scripts\activate.bat)
pip install -r requirements.txt        # solo produccion
pip install -r requirements-dev.txt    # produccion + pruebas (pytest)
```

## Uso diario

1. Guarda los 2 adjuntos del correo en `data/`.
2. Con el entorno activo, corre:

```powershell
python daily_process.py "data\Grid Vol Bloomberg Creasys.xlsx" "data\Operaciones Diarias Opciones y Cartera Vigente al 2026_07_14.xlsx" 2026-07-14 salidas
```

(La fecha y la carpeta de salida son opcionales.)

Genera en `salidas/`:
- `MtM_YYYYMMDD.xlsx` — cartera vigente valorizada.
- `MtM_YYYYMMDD_vencimientos.xlsx` — vencimientos + payoff con fixing USDOBS.
- `resumen_YYYYMMDD.txt` — conteos tipo correo (altas / vencimientos / vigentes).

## Uso como módulo (para automatizar desde el correo)

```python
from daily_process import procesar

res = procesar(grid_path="data/Grid Vol ....xlsx",
               ops_path="data/Operaciones ....xlsx",
               fecha_proceso=None,       # o date(2026,7,14)
               out_dir="salidas")
print(res["resumen"])                    # cuerpo del correo
```

## Lógica de cartera

Cada día reconstruye la cartera vigente desde `Operaciones Diarias` con la regla
`inicio ≤ fecha_proceso < vencimiento` (auto-consistente respecto a la valorización,
sin estado acumulado para eso).
- **Altas** = `inicio == fecha_proceso`
- **Vencimientos** = `vencimiento == fecha_proceso` (payoff con fixing USDOBS)

## Seguimiento de cambios en operaciones vigentes

Al cerrar cada corrida se guarda `estado/estado_cartera_YYYYMMDD.json` (carpeta
propia, separada de `salidas/` para que sobreviva aunque esta se limpie o
archive periódicamente): una foto de los campos clave (contraparte, monto,
strike, vencimiento, tipo, modalidad, entrega, lado, moneda, par) de la
cartera vigente de esa fecha — **un archivo por fecha de proceso**, no uno
solo que se va sobrescribiendo.

La base de comparación de una corrida es siempre la foto de la **última fecha
de proceso estrictamente anterior** a la que se está procesando (nunca la de
la misma fecha, aunque ya exista). Esto es importante si una misma fecha se
reprocesa (por error o corrección de los inputs): la segunda corrida se sigue
comparando contra el día anterior real, no contra el intento previo de esa
misma fecha — así no se generan avisos falsos por corregir un dato, ni se
pierde silenciosamente un aviso real si la operación ya había quedado
reflejada en la foto de la primera corrida.

- Si algún campo de una operación ya vigente cambia de un día a otro (ej. cambio
  de contraparte), se avisa en el resumen: `Folio ...: cambio de <campo> de
  '<antes>' a '<ahora>'.` La planilla de MtM ya queda actualizada con el valor
  nuevo, porque siempre se reconstruye desde `Operaciones Diarias`.
- Si una operación que debía seguir vigente (vencimiento posterior a la fecha de
  proceso) desaparece del archivo de operaciones sin llegar a su vencimiento, se
  interpreta como **liquidación anticipada** y se avisa: `Folio ...: liquidada
  anticipadamente (...).`
- La primera vez que se corre (o si no hay ninguna foto de una fecha anterior)
  no hay base con la cual comparar; no se genera ningún aviso ese día, solo se
  crea la foto de esa fecha para las corridas siguientes.

## Validación

- Los 64 MtM coinciden con el Excel (Δ máx = 0,000000 CLP).
- Forwards, vols y strikes de la superficie coinciden exactamente.
- Los payoffs de vencimientos coinciden con el archivo `_vencimientos` de referencia.

## Pruebas

```powershell
pytest
```

- **Unitarias** (datos sintéticos): motor de valorización (paridad put-call, signos,
  interpolación, superficie), fx por calce (1-1, N-1, N-M), clasificación de cartera,
  fechas y completado de curva. Corren en cualquier máquina.
- **Regresión de mercado**: contra el grid real del 2026-07-14 (`tests/data/`),
  con valores esperados validados al centavo contra el maestro Excel.
- **Regresión de pipeline**: requiere el archivo de cartera del 07-14 en `data/`;
  se salta automáticamente si no está (no se versiona por confidencialidad).

CI: GitHub Actions (`.github/workflows/tests.yml`) corre la suite en cada push a
`main` y en cada pull request.

## Nota de seguridad

`data/`, `salidas/` y `estado/` están en `.gitignore`: los archivos de trabajo
diario, los resultados y la foto de la cartera **no se suben** al repositorio.
Los archivos de **cartera** (contrapartes, posiciones) son confidenciales y nunca
deben versionarse; los datos de **mercado** no son sensibles y hay copias de
prueba en `tests/data/`. Usa un repositorio **privado**.