#!/bin/bash
# Скрипт для установки системных зависимостей Qt

echo "Установка системных зависимостей для Qt..."
sudo apt-get update
sudo apt-get install -y \
    libxcb-cursor0 \
    libxcb-xinerama0 \
    libxcb-xfixes0 \
    libxcb-render-util0 \
    libxcb-icccm4 \
    libxcb-image0 \
    libxcb-keysyms1 \
    libxcb-randr0 \
    libxcb-render0 \
    libxcb-shape0 \
    libxcb-sync1 \
    libxcb-xkb1 \
    libxkbcommon-x11-0

echo "Зависимости установлены!"



