# Deploy en Railway — paso a paso

## 1. Crear cuenta gratuita
- Ir a https://railway.app
- Registrarse con Google o GitHub

## 2. Subir el código a GitHub (una sola vez)
- Crear un repo nuevo en github.com (puede ser privado)
- Subir la carpeta `gornitz_app` completa

## 3. Deploy en Railway
- En Railway: New Project → Deploy from GitHub repo
- Seleccionar el repo
- Railway detecta automáticamente Flask y hace el deploy

## 4. URL pública
- Railway te da una URL tipo: https://gornitz-app-xxxx.up.railway.app
- Compartís esa URL con quien opere en el lab

## Archivos de la app
```
gornitz_app/
├── app.py            ← servidor Flask
├── derivaciones.py   ← lógica de derivaciones y mappings
├── requirements.txt  ← dependencias
├── Procfile          ← comando de inicio
├── railway.json      ← config Railway
└── templates/
    └── index.html    ← interfaz web
```

## Actualizar mappings
Cuando aparezca una determinación nueva, editás `derivaciones.py`
(diccionarios DERIVAR y NOMBRE_ANALISIS) y hacés push al repo.
Railway hace el redeploy automático.
