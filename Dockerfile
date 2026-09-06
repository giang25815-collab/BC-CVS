FROM python:3.10-slim

WORKDIR /app

# Cai dat dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy toan bo code vao container
COPY . .

# Chay bot telegram
CMD ["python", "telegram_bot.py"]
