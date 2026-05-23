#!/bin/bash
# Script de instalación - ejecutar en tu Mac
# Primero clona o descarga esta carpeta, luego corre: bash instalar-en-mac.sh

echo "→ Instalando dependencias..."
npm install

echo "→ Agregando plataforma iOS..."
npx cap add ios

echo "→ Sincronizando archivos..."
npx cap sync ios

echo "→ Abriendo Xcode..."
npx cap open ios

echo ""
echo "✓ Listo. En Xcode:"
echo "  1. Conecta tu iPhone con el cable"
echo "  2. Selecciona tu iPhone en el menú de dispositivos (arriba)"
echo "  3. Presiona el botón ▶ para instalar"
