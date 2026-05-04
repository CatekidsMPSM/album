# CateKids Álbum 🎖️

## Archivos
- `index.html` → Sube a GitHub (activa GitHub Pages)  
- `Code.gs`    → Pega en Google Apps Script

## Pasos

### 1. Google Apps Script
1. Abre [script.google.com](https://script.google.com) → Nuevo proyecto
2. Pega el contenido de `Code.gs`
3. Cambia `DRIVE_FOLDER_ID` por el ID de tu carpeta de Google Drive
4. **Implementar → Nueva implementación → Web App**
   - Ejecutar como: **Yo**
   - Acceso: **Cualquier persona**
5. Copia la URL que termina en `/exec`

### 2. index.html
1. Abre `index.html` y busca `WEBAPP_URL`
2. Pega la URL del paso anterior
3. (Opcional) Si quieres comentarios IA en las insignias, pon tu clave de Anthropic en `ANTHROPIC_API_KEY`

### 3. GitHub Pages
1. Sube `index.html` a tu repositorio de GitHub
2. Settings → Pages → Branch: main → Carpeta: root
3. ¡Listo! Tu app estará en `https://TU_USUARIO.github.io/TU_REPO/`

## Hojas de Google Sheets necesarias

**Hoja `DATOS`**
```
A: Código | B: Nombre | C: Grupo | D: Catequista
```

**Hoja `Respuestas de formulario 1`**
```
A: Timestamp | B: Email | C: Código | D: Nombre | E: Tarjeta# | F: URL imagen
```

La hoja `INSIGNIAS_CK` se crea automáticamente.
