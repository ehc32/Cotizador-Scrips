# Generador de Cotizaciones con Conversión a PDF

Esta aplicación genera documentos de cotización en formato Word y los convierte automáticamente a PDF antes de enviarlos al usuario.

## 🚀 Funcionalidades

- ✅ Genera documentos Word con plantillas personalizadas
- ✅ Convierte automáticamente a PDF
- ✅ Guarda datos en Google Sheets
- ✅ Manejo robusto de errores
- ✅ Limpieza automática de archivos temporales

## 📋 Proceso de Conversión a PDF

La aplicación intenta convertir el documento a PDF usando **múltiples métodos** en este orden:

1. **LibreOffice** (método principal)
2. **Unoconv** (alternativa)
3. **Pandoc** (tercera opción)
4. **LibreOffice con ruta absoluta** (respaldo)

Si **ningún método funciona**, la aplicación envía el documento Word como respaldo.

## 🛠️ Herramientas de Conversión

### LibreOffice (Recomendado)
```bash
soffice --headless --convert-to pdf documento.docx
```

### Unoconv (Alternativa)
```bash
unoconv -f pdf documento.docx
```

### Pandoc (Tercera opción)
```bash
pandoc documento.docx -o documento.pdf
```

## 📦 Instalación

### Dependencias del Sistema
```bash
apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    unoconv \
    pandoc
```

### Dependencias de Python
```bash
pip install -r requirements.txt
```

## 🔧 Configuración

### Variables de Entorno
```bash
GOOGLE_CREDENTIALS=tu_json_de_credenciales
ID_DRIVE=tu_id_de_google_sheets
PORT=5000
```

## 🚀 Despliegue

### Render.com
La aplicación está configurada para desplegarse automáticamente en Render.com con todas las dependencias necesarias.

### Docker
```bash
docker build -t cotizaciones-app .
docker run -p 5000:5000 cotizaciones-app
```

## 📝 Uso

1. Envía una petición POST a `/generar-word`
2. La aplicación genera el documento Word
3. Intenta convertir a PDF automáticamente
4. Envía el PDF (o Word como respaldo)
5. Limpia archivos temporales

## 🔍 Verificación

Para verificar qué herramientas están disponibles:
```bash
python check_tools.py
```

## 📊 Logs

La aplicación muestra logs detallados del proceso de conversión:
- ✅ Éxito en conversión
- ❌ Errores de conversión
- 📁 Archivos creados/eliminados

## 🆘 Solución de Problemas

### Error 500 en conversión
1. Verifica que LibreOffice esté instalado
2. Revisa los logs de la aplicación
3. Ejecuta `python check_tools.py`

### No se genera PDF
- La aplicación automáticamente envía el documento Word como respaldo
- Revisa los logs para identificar el problema específico

## 📞 Soporte

Si tienes problemas con la conversión a PDF, la aplicación siempre enviará el documento Word como respaldo para asegurar que el usuario reciba su documento. 