# Orimi - Django Shop Management System

- Python 3.12+
- Docker 8 Docker Compose
- PostgreSQL 16
- Nginx 


### 1. Start

```bash
    git clone <repository-url>
    cd orimi
```

### 2. Create env

```bash
    make env
```


### 3. Write env


```bash
SECRET_KEY=your-secret-key-here
ALLOWED_HOSTS=localhost 127.0.0.1 your-domain.com
DEBUG=TRUE

# Database
DB_ENGINE=d
DB_NAME=
DB_USER=
DB_PASSWORD=
DB_HOST=
DB_PORT=


S3_ACCESS_KEY_ID=
S3_SECRET_ACCESS_KEY=y
S3_ENDPOINT_URL=
S3_BUCKET_NAME=
IMAGE_URL=
```

### 4. RUN

```bash
    make run
```

