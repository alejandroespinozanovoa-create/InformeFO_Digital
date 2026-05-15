# Informe Gerencial Fibra · Atento

Dashboard gerencial de operaciones Canal Online — Fibra.

## 🚀 Estructura del repositorio

```
├── index.html              # Dashboard web principal
├── data.js                 # Datos BBDD (auto-generado desde Excel)
├── convert.py              # Script de conversión Excel → JS
├── INFORME FIBRA FINAL.xlsx # Fuente de datos
├── .github/
│   └── workflows/
│       └── update-data.yml # Auto-actualización al subir el Excel
└── README.md
```

## 🔄 Actualización automática

Cada vez que subas una nueva versión del archivo Excel al repositorio, GitHub Actions ejecutará `convert.py` automáticamente y actualizará `data.js`. El dashboard se refrescará con los nuevos datos sin intervención manual.

**Flujo:**
1. Sube el nuevo `.xlsx` al repositorio
2. GitHub Actions ejecuta `convert.py`
3. Se genera el nuevo `data.js` y se hace commit automático
4. GitHub Pages sirve el dashboard actualizado

## 🌐 GitHub Pages

1. Ve a **Settings → Pages**
2. Selecciona **Branch: main** y carpeta **/ (root)**
3. Tu URL quedará: `https://<usuario>.github.io/<repositorio>/`

## 📥 Descargas disponibles desde el dashboard

- **Excel**: Botón "Excel" en la barra superior
- **PDF**: Botón "PDF" — captura visual de la pestaña activa
