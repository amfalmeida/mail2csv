FROM python:3.12-slim

WORKDIR /app

# No extra dependencies needed — uses only Python stdlib
COPY mail2csv.py .

# Output directory (mount a host folder here to retrieve the CSV)
RUN mkdir /output

ENV OUTPUT=/output/mail2csv.csv

ENTRYPOINT ["python", "mail2csv.py"]
