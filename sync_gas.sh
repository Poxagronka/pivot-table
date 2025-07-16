#!/bin/bash

# Функция для синхронизации GAS проектов (локальные файлы → GAS + GitHub)
sync_gas_project() {
    local project_name=$1
    local project_path=$2
    
    echo "🔄 Синхронизация GAS проекта: $project_name"
    
    if [ ! -d "$project_path" ]; then
        echo "❌ Папка проекта не найдена: $project_path"
        return 1
    fi
    
    cd "$project_path"
    
    # Проверяем, есть ли .clasp.json
    if [ ! -f ".clasp.json" ]; then
        echo "❌ Файл .clasp.json не найден. Проект не связан с GAS"
        echo "💡 Выполните: clasp clone <script_id>"
        return 1
    fi
    
    echo "📝 Проверяем локальные изменения..."
    
    # Добавляем файлы в git
    git add .
    
    # Сначала всегда пушим в GAS (независимо от git статуса)
    echo "📤 Отправляем все файлы в Google Apps Script..."
    if clasp push --force; then
        echo "✅ GAS обновлен успешно!"
    else
        echo "❌ Ошибка при отправке в GAS"
        echo "💡 Попробуйте выполнить вручную:"
        echo "   cd $project_path"
        echo "   clasp push --force"
        return 1
    fi
    
    # Теперь работаем с git
    if [[ `git status --porcelain` ]]; then
        echo "✅ Найдены локальные изменения в проекте: $project_name"
        echo ""
        
        # Показываем статус изменений
        echo "📋 Izmenenные файлы:"
        git status --short
        echo ""
        
        # Запрашиваем commit message
        echo "💬 Введите сообщение для коммита (Enter для авто-сообщения):"
        read -r commit_message
        
        # Если сообщение пустое, используем дефолтное
        if [ -z "$commit_message" ]; then
            commit_message="Local changes: $(date '+%Y-%m-%d %H:%M')"
        fi
        
        echo "📝 Создаем коммит с сообщением: \"$commit_message\""
        git commit -m "$commit_message"
        
        echo "🔄 Получаем последние изменения из GitHub..."
        if ! git pull --rebase origin main; then
            echo "⚠️  Возможны конфликты при merge. Проверьте вручную:"
            echo "   cd $project_path"
            echo "   git status"
            echo "   Разрешите конфликты и выполните: git push"
            return 1
        fi
        
        echo "📤 Отправляем изменения на GitHub..."
        
        # Проверяем, есть ли upstream branch
        if ! git rev-parse --abbrev-ref --symbolic-full-name @{u} > /dev/null 2>&1; then
            echo "🔗 Настраиваем upstream branch..."
            git push --set-upstream origin main
        else
            git push
        fi
        
        if [ $? -eq 0 ]; then
            echo "✅ GitHub sync успешно завершен!"
            echo "🎉 $project_name - синхронизация завершена успешно!"
            echo "   ✅ GAS: обновлен"
            echo "   ✅ GitHub: обновлен"
        else
            echo "❌ $project_name - ошибка при отправке на GitHub"
            echo "💡 Попробуйте выполнить вручную:"
            echo "   cd $project_path"
            echo "   git push"
            return 1
        fi
    else
        echo "📄 $project_name - нет локальных изменений в git"
        echo "🎉 $project_name - синхронизация завершена!"
        echo "   ✅ GAS: обновлен"
        echo "   ℹ️  GitHub: без изменений"
    fi
    
    # Спрашиваем, нужно ли открыть GAS в браузере
    echo "🌐 Открыть GAS проект в браузере? (y/n):"
    read -r open_gas
    if [ "$open_gas" = "y" ] || [ "$open_gas" = "Y" ]; then
        clasp open
    fi
    
    echo ""
    return 0
}

# Функция для принудительного GAS пуша
force_gas_push() {
    echo "🚀 Принудительная синхронизация GAS проектов..."
    echo "=========================================="
    
    # Принудительный пуш UA Management Optimized
    echo "📤 Принудительный пуш UA Management Optimized..."
    if [ -d ~/UA-management-Optimized- ]; then
        cd ~/UA-management-Optimized-
        if [ -f ".clasp.json" ]; then
            clasp push --force
            echo "✅ UA Management Optimized принудительно обновлен в GAS"
        else
            echo "❌ .clasp.json не найден в UA Management Optimized"
        fi
    else
        echo "❌ Папка UA Management Optimized не найдена"
    fi
    
    echo ""
    
    # Принудительный пуш Pivot Table
    echo "📤 Принудительный пуш Pivot Table..."
    if [ -d ~/pivot-table-gas ]; then
        cd ~/pivot-table-gas
        if [ -f ".clasp.json" ]; then
            clasp push --force
            echo "✅ Pivot Table принудительно обновлен в GAS"
        else
            echo "❌ .clasp.json не найден в Pivot Table"
        fi
    else
        echo "❌ Папка Pivot Table не найдена"
    fi
    
    echo "=========================================="
    echo "🏁 Принудительная синхронизация завершена!"
}

# Основной скрипт
echo "🚀 Система синхронизации проектов (локальные файлы → GAS/GitHub)"
echo "=================================================================="

# Проверяем наличие clasp
if ! command -v clasp &> /dev/null; then
    echo "⚠️  Clasp не установлен!"
    echo "🔧 Установить сейчас? (y/n):"
    read -r install_clasp
    
    if [ "$install_clasp" = "y" ] || [ "$install_clasp" = "Y" ]; then
        echo "📦 Устанавливаем clasp..."
        npm install -g @google/clasp
        echo "✅ Clasp установлен!"
        echo "🔐 Теперь авторизуйтесь: clasp login"
    fi
    echo ""
fi

# Показываем меню
echo "🎛️  Выберите действие:"
echo "1) Синхронизация GAS проектов"
echo "2) Принудительный GAS пуш (без git)"
echo "3) Выход"
echo ""
echo "Введите номер (1-3):"
read -r choice

case $choice in
    1) 
        echo "🚀 Синхронизация GAS проектов..."
        echo "=========================================="
        
        sync_gas_project "UA Management Optimized" ~/UA-management-Optimized-
        echo "=========================================="
        
        sync_gas_project "Pivot Table" ~/pivot-table-gas
        echo "=========================================="
        
        echo "🏁 GAS проекты обработаны!"
        ;;
    2) 
        force_gas_push
        ;;
    3) 
        exit 0
        ;;
    *) 
        echo "❌ Неверный выбор"
        ;;
esac