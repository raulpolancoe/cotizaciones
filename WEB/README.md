# Generador de documentos Santafé de los Guaduales

Aplicación web para generar cotizaciones y cuentas de cobro en PDF con la identidad visual de la reserva.

## Funciones

- Selección inicial entre cotización y cuenta de cobro.
- Productos cargados desde `Lista_Precios.xlsx`.
- Artículos manuales y valores negativos para descuentos.
- Cuentas de cobro con número, cliente, NIT, concepto y fecha del servicio.
- Uno o varios abonos opcionales, totalizados automáticamente con el saldo.
- PDF con logo, firma y colores institucionales; texto de la cuenta de cobro en negro.

## Ejecución

```bash
pip install -r requirements.txt
python app_weasy_backend.py
```

Abre `http://localhost:5000` en el navegador.

WeasyPrint necesita sus bibliotecas nativas (GTK/Pango) cuando se ejecuta en Windows. En Render/Linux deben incluirse los paquetes del sistema requeridos por WeasyPrint.
