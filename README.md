# SIGES Automation

Automatizacion del flujo mensual de **SALUD** en SIGES:

1. Extrae datos desde Excel y actualiza `reporte_siges_salud.xlsx`.
2. Genera el archivo `Detalle de horas SALUD MMYYYY.xlsx`.
3. Construye borrador de correo en Gmail con resumen y adjunto.
4. Actualiza una hoja de Google Sheets con el detalle consolidado.

## Estructura del proyecto

- `0. key_processes_extration.py`: refresca `indicadores.xlsx`, arma hojas `general` y `detallexproyecto` en `reporte_siges_salud.xlsx`.
- `1. hours_detailed.py`: consulta `Copia Reportes Siges v2.2 3.xlsm` y exporta `Detalle de horas SALUD MMYYYY.xlsx`.
- `2. email_salud.py`: lee `reporte_siges_salud.xlsx`, genera HTML y crea borrador en Gmail con adjunto.
- `3. gsheets_update_siges_salud.py`: sube/actualiza datos en Google Sheets desde `reporte_siges_salud.xlsx`.
- `4. execution.py`: ejecuta todo el pipeline en orden.

## Requisitos

- Windows con Microsoft Excel instalado (automatizacion COM).
- Python 3.12.
- Dependencias Python en `requirements.txt`.

## Configuracion inicial

1. Crear y activar entorno virtual:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

2. Instalar dependencias:

```powershell
python -m pip install -r requirements.txt
```

3. Archivos necesarios en la raiz del proyecto:

- `indicadores.xlsx`
- `Copia Reportes Siges v2.2 3.xlsm`
- `reporte_siges_salud.xlsx` (plantilla destino)
- `client_secret_*.json` (OAuth Gmail)
- `token.json` (se crea/actualiza al autenticar Gmail)
- `pmo-*.json` (Service Account para Google Sheets)

## Ejecucion

Ejecutar todo el pipeline:

```powershell
python "4. execution.py"
```

Ejecutar por pasos (opcional):

```powershell
python "0. key_processes_extration.py"
python "1. hours_detailed.py"
python "2. email_salud.py"
python "3. gsheets_update_siges_salud.py"
```

## Salidas generadas

- `reporte_siges_salud.xlsx` actualizado.
- `Detalle de horas SALUD MMYYYY.xlsx`.
- `reporte_siges_salud_email.html`.
- Borrador en Gmail con adjunto del detalle.
- Hoja de Google Sheets actualizada (`siges_salud`).

## Problemas comunes

- Error COM de Excel (`no hay sesion activa`):
  - Ejecutar desde una sesion interactiva de Windows (usuario logueado).
  - Evitar ejecutar como servicio sin sesion de escritorio.
- Error `ModuleNotFoundError`:
  - Verificar que estas usando el `venv` correcto de este proyecto.
- Error de credenciales Google:
  - Validar `client_secret_*.json`, `token.json` y `pmo-*.json`.

## Seguridad

No subir credenciales sensibles al repositorio:

- `token.json`
- `client_secret_*.json`
- `pmo-*.json`

