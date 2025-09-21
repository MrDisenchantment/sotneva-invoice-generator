# Используем официальный Node.js образ
FROM node:18

# Устанавливаем рабочую директорию внутри контейнера
WORKDIR /app

# Копируем только package.json и package-lock.json для установки зависимостей
COPY package*.json ./

# Устанавливаем зависимости проекта
RUN npm install --production

# Устанавливаем LibreOffice для конвертации документов
RUN apt-get update && \
    apt-get install -y libreoffice && \
    apt-get clean

# Копируем все остальные файлы
COPY . .

# Открываем порт 3000
EXPOSE 3000

# Запускаем сервер
CMD ["npm", "start"]
