# 🚀 DESPLEGAR EN RAILWAY.APP

## PASO 1: Preparar los archivos

Necesitas 4 archivos en una carpeta:

```
Mi_Aplicacion/
├── app.py
├── requirements.txt
├── index.html
└── templates/
    └── index.html (copia del archivo index.html aquí)
```

**IMPORTANTE:** La carpeta `templates/` debe estar en el mismo nivel que `app.py`

---

## PASO 2: Crear carpeta templates

1. Crea una carpeta llamada `templates`
2. Copia el archivo `index.html` dentro de esa carpeta

```
templates/
└── index.html
```

---

## PASO 3: Subir a GitHub

1. Ve a https://github.com/new
2. Crea un nuevo repositorio (ej: `generador-reportes`)
3. Selecciona:
   - ✓ Public (para que Railway pueda acceder)
   - ✗ NO inicialices con README
4. Click en "Create repository"

### En tu computadora:

1. Abre CMD/Terminal en la carpeta de tu aplicación
2. Ejecuta estos comandos:

```bash
git init
git add .
git commit -m "Generador de reportes lotes"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/generador-reportes.git
git push -u origin main
```

(Reemplaza `TU_USUARIO` con tu usuario de GitHub)

---

## PASO 4: Conectar Railway

1. Ve a https://railway.app
2. Click en "Login" (o Sign Up si no tienes cuenta)
3. Login con GitHub
4. Click en "+ New Project"
5. Selecciona "Deploy from GitHub repo"
6. Autoriza a Railway a acceder a GitHub
7. Selecciona el repositorio `generador-reportes`
8. ¡Railway empieza a desplegar automáticamente!

---

## PASO 5: Obtener el link

Después de ~2-3 minutos:

1. En Railway, ve a la pestaña "Deployments"
2. Espera a que vea "✓ Success"
3. En la parte superior, verás un link como:
   ```
   https://generador-reportes-production.up.railway.app
   ```
4. ¡Haz click y listo! Tu app está en vivo

---

## COMPARTIR CON TU COLEGA

Simplemente comparte el link:
```
https://generador-reportes-production.up.railway.app
```

Tu colega abre en el navegador y:
1. Sube el archivo Excel
2. Pone Pto. de venta
3. Pone Z
4. ¡Descarga el reporte!

---

## SOLUCIÓN DE PROBLEMAS

### "Build failed" en Railway

**Solución:** Verifica que:
- ✓ El archivo `requirements.txt` existe y tiene las librerías correctas
- ✓ El archivo `app.py` está en la raíz
- ✓ La carpeta `templates/` tiene el archivo `index.html`
- ✓ El repositorio GitHub es público

### "Connection refused"

**Solución:** Espera 2-3 minutos a que el deploy termine

### El formulario se ve roto

**Solución:** 
- Recarga la página (Ctrl+F5)
- Borra el caché del navegador

### "File not found" en el navegador

**Solución:** Verifica en Railway:
1. Build status = "Success" ✓
2. Deploy status = "Success" ✓
3. Reinicia el deploy

---

## ACTUALIZAR LA APP

Si realizas cambios:

1. Copia los archivos nuevos a la carpeta
2. En CMD/Terminal:
   ```bash
   git add .
   git commit -m "Descripción del cambio"
   git push
   ```
3. Railway detectará los cambios y redesplegará automáticamente

---

## CONFIGURACIÓN ADICIONAL (Opcional)

Si Railway te pide variables de entorno:

En Railway Dashboard → Environment:
- No necesitas agregar nada
- Railway automáticamente detecta Flask

---

## LINK FINAL PARA COMPARTIR

Una vez desplegado, tu link será algo como:

```
https://generador-reportes-production.up.railway.app
```

**Guarda este link y compártelo con tu colega cajero** ✅

---

## PREGUNTAS FRECUENTES

**P: ¿Mi colega necesita cuenta en GitHub?**
R: No, solo necesita el link en el navegador.

**P: ¿Hay costo?**
R: Tienes $5 USD/mes gratis. Esta app usa menos.

**P: ¿Puedo cambiar el nombre del link?**
R: Sí, en Railway → Settings → Domains

**P: ¿Qué pasa si se acaban los $5 USD?**
R: La app se detiene. Puedes seguir usando la versión local o pagar.

---

**¡Listo! Ahora tienes una app web profesional en vivo! 🎉**
