"""Punto de entrada estable para despliegues desde la raíz del repositorio."""

from WEB.app_weasy_backend import app


if __name__ == "__main__":
    import os

    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
