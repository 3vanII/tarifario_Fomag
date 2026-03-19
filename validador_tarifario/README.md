# Validador Tarifario

Aplicación de escritorio en Python para preparar la primera etapa del proceso de validación de tarifarios: crear un Excel final con una única hoja agrupadora, uniendo todas sus pestañas y agregando la columna **Pestaña origen**.

## Qué hace

- Permite seleccionar un archivo Excel (`.xlsx` o `.xlsm`).
- Permite seleccionar la carpeta de salida.
- Lee todas las hojas del archivo, excepto `AGRUPADOR`.
- Detecta de forma automática la fila más probable de encabezado en cada hoja.
- Une la información en una única hoja nueva llamada `AGRUPADOR` (o el nombre que indiques).
- Agrega la columna `Pestaña origen` al inicio.
- Genera un nuevo archivo en la carpeta de salida, sin modificar el original.
- El archivo de salida conserva únicamente la hoja agrupada que solicitaste.

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
