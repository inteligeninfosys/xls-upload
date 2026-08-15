FROM node:24-slim

ENV NODE_ENV=production

WORKDIR /app

# Install dependencies first to maximize build cache reuse
COPY package.json ./

RUN npm install --omit=dev \
    && npm cache clean --force

# Copy source
COPY . .

# Run as the non-root user already provided by the Node image
RUN chown -R node:node /app

USER node

EXPOSE 3000

CMD ["node", "index.js"]

# docker build -t inteligeninfosys/xls-upload:1.5 .
