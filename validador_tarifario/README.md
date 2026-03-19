# Validador Tarifario

Aplicación de escritorio en Python para preparar la primera etapa del proceso de validación de tarifarios: crear un Excel final con una única hoja agrupadora, uniendo todas sus pestañas y agregando la columna **Pestaña origen**.

## Qué hace

- Permite seleccionar un archivo Excel (`.xlsx` o `.xlsm`).
- Permite seleccionar la carpeta de salida.
- Lee todas las hojas del archivo, excepto `AGRUPADOR`.
- Detecta de forma automática la fila más probable de encabezado en cada hoja.
- Une la información en una única hoja nueva llamada `AGRUPADOR` (o el nombre que indiques).
- Agrega la columna `Pestaña origen` al inicio.
- Valida la columna `CUPS` contra las bases `SOAT UVB 2026.xlsx`, `SOAT 2025.xlsx` y `TablaReferencia_CUPS__1.csv`.
- Valida además la combinación `CÓDIGO HABILITACIÓN (12 DÍGITOS)` + `COD SERVICIO` contra la base `SERVICIOS REPS LIMPIOS.xlsx`.
- Calcula la analítica inicial de `TARIFA FOMAG` con las columnas `DIFERENCIA VS SOAT 2025` y `DIFERENCIA VS SOAT UVB 2026`, usando la fórmula `SOAT - TARIFA FOMAG`.
- Envía a una hoja `ERRORES` las filas con inconsistencias, indicando si el `CUPS` es inválido, si el código de habilitación no existe en REPS o si la sede existe pero no tiene habilitado el servicio reportado.
- Genera un nuevo archivo en la carpeta de salida, sin modificar el original.
- El archivo de salida conserva la hoja agrupada solicitada y, si aplica, una hoja adicional con errores.

## Estructura

```text
validador_tarifario/
├── README.md
├── requirements.txt
├── run.py
├── build_exe.bat
└── src/
    └── validador_tarifario/
        ├── __init__.py
        ├── app.py
        ├── ui/
        │   └── main_window.py
        ├── services/
        │   └── agrupador_service.py
        └── utils/
            └── paths.py
```

## Requisitos

- Python 3.11 o superior recomendado
- Windows, Linux o macOS

## Instalación

En la carpeta del proyecto:

```bash
pip install -r requirements.txt
```

## Ejecución

```bash
python run.py
```

## Cómo usar

1. Haz clic en **Examinar** y selecciona el archivo Excel.
2. Haz clic en **Elegir** y selecciona la carpeta de salida.
3. Opcionalmente cambia el nombre de la hoja agrupadora.
4. Presiona **Procesar archivo**.
5. Revisa el log en pantalla.

## Generar EXE en Windows

Primero instala PyInstaller:

```bash
pip install pyinstaller
```

Luego ejecuta:

```bash
build_exe.bat
```

El ejecutable quedará en la carpeta `dist`.

## Notas

- Soporta `.xlsx` y `.xlsm`.
- En archivos `.xlsm`, intenta conservar la estructura del libro al generar el archivo con una sola hoja final.
- La lógica está separada por módulos para que luego puedas agregar más validaciones.
- Las bases de referencia para la validación CUPS se leen desde la carpeta `base de datos/` ubicada en la raíz del repositorio.
