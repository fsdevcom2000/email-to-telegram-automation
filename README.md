# Email-to-Telegram Automation (Windows)

## Overview

This project automates the processing of emails from a specific sender in Microsoft Outlook. It saves PDF attachments locally and sends them to a Telegram chat. If Outlook is not running, a fallback notification is sent to Telegram instead.

## Features

- Monitors new unread emails from a specified sender.
    
- Saves PDF attachments locally with unique filenames to avoid overwrites.
    
- Sends PDF attachments to Telegram via Bot API.
    
- Sends fallback notification if Outlook is not running.
    
- Centralized logging of all actions and errors.
    

## Requirements

- Windows 10 / 11
    
- Microsoft Outlook installed and configured
    
- PowerShell 5.x or higher
    
- Internet access for Telegram API
    

## Installation

1. Download files.

2. Place `auto.ps1` and `run.bat` in `C:\Temp` (or adjust paths in the batch script).
    
3. Update the configuration in `auto.ps1`:
    
    - `$senderEmail` – the email address to monitor
        
    - `$botToken` – your Telegram bot token
        
    - `$chatId` – the Telegram chat ID
        
4. Update `run.bat` if needed (paths to scripts and log file).
    

## Usage

1. Run the batch script manually or schedule it using Windows Task Scheduler:
    

`C:\Temp\run.bat`

2. Logs are written to `C:\Temp\auto_log.txt`.
    
3. If Outlook is not running, a Telegram notification is sent instead of processing emails.
    

## How it works

- `run.bat` checks if Outlook is running.
    
- If Outlook is running → `auto.ps1` processes unread emails and sends PDFs.
    
- If Outlook is not running → `auto.ps1 -Fallback` sends a Telegram notification.
    
- All actions are logged in a centralized log file.
    

## Notes

- PDF files are saved in the `$savePath` directory.
    
- Ensure your Telegram bot has permission to send messages to the chat.
    
- PowerShell must allow script execution (`Bypass` or `RemoteSigned`).
    

---

# Автоматизация Email → Telegram (Windows)

## Обзор

Проект автоматически обрабатывает письма от указанного отправителя в Microsoft Outlook, сохраняет PDF-вложения локально и отправляет их в Telegram. Если Outlook не запущен, вместо обработки отправляется уведомление в Telegram.

## Возможности

- Отслеживание новых непрочитанных писем от указанного отправителя.
    
- Сохранение PDF-вложений локально с уникальными именами, чтобы не перезаписывать файлы.
    
- Отправка PDF-вложений в Telegram через Bot API.
    
- Отправка уведомления в Telegram, если Outlook не запущен.
    
- Централизованное логирование всех действий и ошибок.
    

## Требования

- Windows 10 / 11
    
- Установленный и настроенный Microsoft Outlook
    
- PowerShell 5.x или выше
    
- Доступ в интернет для Telegram API
    

## Установка

1. Скачайте все файлы проекта:
    
2. Поместите `auto.ps1` и `run.bat` в `C:\Temp` (или измените пути в batch-скрипте).
    
3. Настройте конфигурацию в `auto.ps1`:
    
    - `$senderEmail` – адрес электронной почты для мониторинга
        
    - `$botToken` – токен Telegram бота
        
    - `$chatId` – ID чата Telegram
        
4. При необходимости обновите `run.bat` (пути к скриптам и лог-файл).
    

## Использование

1. Запустите batch-скрипт вручную или через Планировщик заданий Windows:
    

`C:\Temp\run.bat`

2. Логи сохраняются в `C:\Temp\auto_log.txt`.
    
3. Если Outlook не запущен, отправляется уведомление в Telegram вместо обработки писем.
    

## Как это работает

- `run.bat` проверяет, запущен ли Outlook.
    
- Если Outlook запущен → `auto.ps1` обрабатывает новые письма и отправляет PDF.
    
- Если Outlook не запущен → `auto.ps1 -Fallback` отправляет уведомление в Telegram.
    
- Все действия фиксируются в централизованном лог-файле.
    

## Примечания

- PDF-файлы сохраняются в директории `$savePath`.
    
- Убедитесь, что ваш Telegram бот может отправлять сообщения в указанный чат.
    
- Политика выполнения PowerShell должна разрешать запуск скриптов (`Bypass` или `RemoteSigned`).
