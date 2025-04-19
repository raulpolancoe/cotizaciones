# Cotizador Web con PDF (Render-ready)

Esta app permite generar cotizaciones en PDF con plantilla personalizada desde cualquier navegador.

## ¿Qué incluye?

- Interfaz web para ingresar nombre y productos
- Generación de PDF con nombre personalizado
- Compatible con Render.com

## Instrucciones

1. Agrega el archivo `Lista_Precios.xlsx` con tu tabla de productos en la raíz del proyecto.

2. Instala dependencias:

```
pip install -r requirements.txt
```

3. Ejecuta la aplicación:

```
python app_weasy_backend.py
```

4. Abre en el navegador:

```
http://localhost:5000
```

## Despliegue en Render

- Subir este repositorio a GitHub
- Crear nuevo Web Service en [https://render.com](https://render.com)
- Usa `python app_weasy_backend.py` como comando de inicio
