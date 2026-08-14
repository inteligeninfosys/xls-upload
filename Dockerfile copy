# Use Node base image
FROM node:18-slim

RUN apt-get update && apt-get install -y \
    libaio1 unzip curl build-essential \
 && rm -rf /var/lib/apt/lists/*

# Copy Oracle Instant Client zip to container
COPY instantclient-basiclite-linux.x64-23.8.0.25.04.zip /tmp/ic.zip

RUN mkdir -p /opt/oracle \
 && unzip /tmp/ic.zip -d /opt/oracle \
 && ln -s /opt/oracle/instantclient_21_13 /opt/oracle/instantclient \
 && rm /tmp/ic.zip

ENV LD_LIBRARY_PATH=/opt/oracle/instantclient
ENV OCI_LIB_DIR=/opt/oracle/instantclient
ENV OCI_INC_DIR=/opt/oracle/instantclient/sdk/include
ENV PATH=$PATH:/opt/oracle/instantclient

WORKDIR /app
COPY package*.json ./
RUN npm install
COPY . .
EXPOSE 3000
CMD ["node", "index.js"]

# docker build -t inteligeninfosys/xls-upload:1.3 .
