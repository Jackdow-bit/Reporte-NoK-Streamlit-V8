# Reporte NoK Streamlit V8

## Ejecución
```bash
pip install -r requirements.txt
streamlit run Reporte_NoK_Streamlit_V8.py
```

La aplicación acepta PDFs individuales, uno o varios ZIP y combinaciones PDF + ZIP.

El límite de carga de Streamlit está configurado en 4096 MB en `.streamlit/config.toml`. Si el proveedor de hosting tiene un límite propio (proxy, balanceador o plataforma), ese límite también debe aumentarse.
