#!/bin/bash

# Умный автоматический деплой с анализом изменений

echo "🔍 Анализ изменений..."

# Проверяем есть ли изменения
if git diff-index --quiet HEAD --; then
    echo "ℹ️  Нет изменений для деплоя"
    exit 0
fi

# Анализируем какие файлы изменились
CHANGED_FILES=$(git diff --name-only)
CHANGED_JS_FILES=$(git diff --name-only | grep "\.js$" || true)

# Создаем предложение для коммита на основе изменений
SUGGESTED_MSG="Code updates:"

# Проверяем типы изменений
if echo "$CHANGED_FILES" | grep -q "06_Analytics.js"; then
    SUGGESTED_MSG="$SUGGESTED_MSG Analytics improvements,"
fi

if echo "$CHANGED_FILES" | grep -q "01_Config.js"; then
    SUGGESTED_MSG="$SUGGESTED_MSG Configuration updates,"
fi

if echo "$CHANGED_FILES" | grep -q "15_TableBuilder.js"; then
    SUGGESTED_MSG="$SUGGESTED_MSG Table builder enhancements,"
fi

if echo "$CHANGED_FILES" | grep -q "16_RowGrouping.js"; then
    SUGGESTED_MSG="$SUGGESTED_MSG Row grouping improvements,"
fi

if echo "$CHANGED_FILES" | grep -q "05_ApiClient.js"; then
    SUGGESTED_MSG="$SUGGESTED_MSG API client updates,"
fi

# Убираем последнюю запятую
SUGGESTED_MSG=$(echo "$SUGGESTED_MSG" | sed 's/,$//')

# Добавляем количество измененных файлов
FILE_COUNT=$(echo "$CHANGED_FILES" | wc -l)
SUGGESTED_MSG="$SUGGESTED_MSG (${FILE_COUNT} files)"

echo "💡 Предложенный коммит: $SUGGESTED_MSG"
echo "📁 Измененные файлы: $(echo $CHANGED_FILES | tr '\n' ' ')"
echo ""
echo "💬 Введите сообщение коммита (Enter = использовать предложенное):"
read -r USER_INPUT

if [ -n "$USER_INPUT" ]; then
    COMMIT_MSG="$USER_INPUT"
else
    COMMIT_MSG="$SUGGESTED_MSG"
fi

echo "📝 Коммит: $COMMIT_MSG"

# Деплой на GAS
echo "📤 Отправка на Google Apps Script..."
if clasp push --force; then
    echo "✅ GAS: Успешно обновлено"
else
    echo "❌ GAS: Ошибка при обновлении"
    exit 1
fi

# Коммит и пуш в Git
git add -A

git commit -m "$COMMIT_MSG

🤖 Generated with Claude Code

Co-Authored-By: Claude <noreply@anthropic.com>"

echo "📤 Отправка на GitHub..."
if git push origin main; then
    echo "✅ GitHub: Изменения отправлены"
else
    echo "❌ GitHub: Ошибка при отправке"
    exit 1
fi

echo "🎉 Умный деплой завершен!"
echo "🔗 Изменения синхронизированы с GAS и GitHub"