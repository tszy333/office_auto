FROM python:3.12-slim

LABEL maintainer="tszy33"
LABEL description="办公自动化工具 - Excel收集表批量导入 + Word模板导出"

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# 数据目录
RUN mkdir -p /data

EXPOSE 5000

ENV CONFIG_FILE=/data/config.ini
ENV DATA_DIR=/data
ENV PORT=5000

CMD ["gunicorn", "-w", "2", "-b", "0.0.0.0:5000", "--timeout", "120", "app:app"]
