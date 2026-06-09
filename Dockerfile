FROM node:20-slim
WORKDIR /app
COPY web-ui/ ./
COPY disp_core.py ./disp_core.py
RUN npm install --production
EXPOSE 8080
ENV NODE_ENV=production
CMD ["node", "server.js"]
