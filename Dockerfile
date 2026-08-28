FROM python:3.9
WORKDIR /app
COPY . /app
RUN apt-get update && apt-get install -y \
    libpq-dev \
    gcc \
    build-essential \
&& rm -rf /var/lib/apt/lists/*
RUN pip install -r requirements.txt
EXPOSE 5000
ENV FLASK_APP=project.py
ENV FLASK_ENV=development
CMD ["flask", "run", "--host=0.0.0.0"]