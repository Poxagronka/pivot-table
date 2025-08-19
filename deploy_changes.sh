#!/bin/bash

# Автоматический деплой изменений на GAS и GitHub
# Получаем сводку изменений как аргумент

CHANGE_SUMMARY="$1"

# Если сводка не передана, используем дефолтное сообщение
if [ -z "$CHANGE_SUMMARY" ]; then
    CHANGE_SUMMARY="Code updates via Claude Code"
fi

echo "🚀 Автоматический деплой изменений..."
echo "📝 Изменения: $CHANGE_SUMMARY"

# 1. Пуш на GAS сервер
echo "📤 Отправка на Google Apps Script..."
if clasp push --force; then
    echo "✅ GAS: Успешно обновлено"
else
    echo "❌ GAS: Ошибка при обновлении"
    exit 1
fi

# 2. Проверяем есть ли изменения для коммита
if git diff-index --quiet HEAD --; then
    echo "ℹ️  Git: Нет изменений для коммита"
else
    echo "📝 Git: Найдены изменения, создаем коммит..."
    
    # Добавляем все изменения
    git add -A
    
    # Создаем коммит с описанием изменений
    git commit -m "$CHANGE_SUMMARY

🤖 Generated with Claude Code

Co-Authored-By: Claude <noreply@anthropic.com>"
    
    # Пушим на GitHub
    echo "📤 Отправка на GitHub..."
    if git push origin main; then
        echo "✅ GitHub: Изменения отправлены"
    else
        echo "❌ GitHub: Ошибка при отправке"
        exit 1
    fi
fi

echo "🎉 Автоматический деплой завершен успешно!"