#!/bin/bash

# Instalar dependencias del sistema si no están instaladas
if ! command -v soffice &> /dev/null; then
    echo "Instalando LibreOffice..."
    apt-get update && apt-get install -y libreoffice libreoffice-writer
fi

if ! command -v unoconv &> /dev/null; then
    echo "Instalando unoconv..."
    apt-get update && apt-get install -y unoconv
fi

if ! command -v pandoc &> /dev/null; then
    echo "Instalando pandoc..."
    apt-get update && apt-get install -y pandoc
fi

# Ejecutar la aplicación
python main.py 