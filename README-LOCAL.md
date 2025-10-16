# Б&Р Ассистент — PRO (WPF .NET 8)

Функции:
- Чат с GPT (OpenAI API)
- Голос: озвучка (TTS), распознавание команд (STT)
- Кнопки: Счёт, Накладная, Акт — генерация DOCX в `%AppData%\BRDesktopAssistant\Documents`
- Трей-иконка, автозапуск Windows через реестр, глобальный хоткей **Ctrl+Alt+B** («Позвать Бир»)
- Иконка приложения

## Сборка локально
1. Установите .NET 8 SDK
2. Создайте переменную `OPENAI_API_KEY`
3. В папке проекта:
```
dotnet restore
dotnet run
```
Публикация в один exe:
```
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

## Голос
- TTS использует установленный голос Windows; если есть русский — выберется автоматически.
- STT включает простые голосовые команды: «создай счёт / накладную / акт», «отправить», «позвать бир».

## Артефакты
Готовые файлы лежат в `%AppData%\BRDesktopAssistant\Documents`.
