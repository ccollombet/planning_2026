FROM python:3.12-slim

ENV DEBIAN_FRONTEND=noninteractive \
    PIP_NO_CACHE_DIR=1 \
    TZ=Europe/Paris

# --------------------------------------------------------------------------
# 1) Dépendances système de base
# --------------------------------------------------------------------------
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl gnupg2 apt-transport-https ca-certificates \
    build-essential unixodbc unixodbc-dev \
    libgssapi-krb5-2 krb5-user locales \
 && rm -rf /var/lib/apt/lists/*

# --------------------------------------------------------------------------
# 2) Locale FR (optionnel)
# --------------------------------------------------------------------------
RUN sed -i 's/# fr_FR.UTF-8 UTF-8/fr_FR.UTF-8 UTF-8/' /etc/locale.gen && locale-gen
ENV LANG=fr_FR.UTF-8 LC_ALL=fr_FR.UTF-8

# --------------------------------------------------------------------------
# 3) Dépôt Microsoft + ODBC 18 + openssl1.1
# --------------------------------------------------------------------------
RUN set -eux; \
    curl -fsSL https://packages.microsoft.com/keys/microsoft.asc \
      | gpg --dearmor -o /usr/share/keyrings/microsoft-prod.gpg; \
    echo "deb [arch=amd64 signed-by=/usr/share/keyrings/microsoft-prod.gpg] https://packages.microsoft.com/debian/12/prod bookworm main" \
      > /etc/apt/sources.list.d/microsoft-prod.list; \
    apt-get update; \
    ACCEPT_EULA=Y apt-get install -y msodbcsql18 mssql-tools18 openssl1.1; \
    rm -rf /var/lib/apt/lists/*

# --------------------------------------------------------------------------
# 4) Alias « ODBC Driver 17 » → Driver 18 (on AJOUTE, on n’écrase pas)
# --------------------------------------------------------------------------
RUN set -eux; \
    DRIVER_PATH="$(ls /opt/microsoft/msodbcsql18/lib64/libmsodbcsql-*.so | head -n1)"; \
    { \
      echo ""; \
      echo "[ODBC Driver 17 for SQL Server]"; \
      echo "Description=Alias vers ODBC Driver 18"; \
      echo "Driver=${DRIVER_PATH}"; \
      echo "UsageCount=1"; \
    } >> /etc/odbcinst.ini

# --------------------------------------------------------------------------
# 5) Environnement Python + app
# --------------------------------------------------------------------------
WORKDIR /app
COPY requirements.txt .
RUN pip install --upgrade pip && pip install -r requirements.txt

COPY app.py ./app.py
COPY .streamlit ./ .streamlit/

EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
