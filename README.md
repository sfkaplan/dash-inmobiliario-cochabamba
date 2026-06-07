# Dashboard Inmobiliario — React + Vercel

App de análisis de mercado inmobiliario construida en React, con gráficos interactivos y módulos financieros.

## Stack
- **React 18** (Create React App)
- **Recharts** — gráficos
- **SheetJS (xlsx)** — lectura de Excel directo desde GitHub
- **Vercel** — deploy gratuito

---



## Estructura
```
inmobiliaria-app/
├── public/
│   └── index.html
├── src/
│   ├── index.js
│   └── App.jsx       ← toda la lógica y UI
├── package.json
├── vercel.json
└── README.md
```

## Datos
Los archivos Excel se sirven desde `public/data`:
- `departamentos_cochabamba.xlsx`
- `casas_cochabamba.xlsx`

El objeto `URLS` en `src/App.jsx` apunta a esos archivos locales para que Vercel los publique junto con la app.
