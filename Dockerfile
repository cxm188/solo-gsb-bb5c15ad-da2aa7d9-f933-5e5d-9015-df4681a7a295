FROM python:3.12-slim

# uvicorn[standard] 需要 gcc 构建少量可选依赖，slim 镜像里先装再清
RUN apt-get update \
    && apt-get install -y --no-install-recommends curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app
COPY static ./static
COPY examples ./examples

# 运行期参数（也可在 compose 中覆盖）
ENV MAX_UPLOAD_MB=20 \
    MAX_AFFECTED=20000

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD curl -fsS http://localhost:8000/api/workbooks || exit 1

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
