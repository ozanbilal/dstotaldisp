FROM node:20-slim
WORKDIR /app
COPY web-ui/ ./
RUN npm install --production
EXPOSE 8080
ENV NODE_ENV=production
CMD ["node", "server.js"]
