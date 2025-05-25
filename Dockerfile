# 使用官方Node.js镜像作为基础镜像
FROM node:18-alpine  
# 选择合适的Node.js版本

# 创建应用工作目录
WORKDIR /usr/src/app

# 复制package.json和package-lock.json（如果有的话），以便安装依赖
COPY package*.json ./

# 安装生产环境依赖，避免重复安装开发依赖
RUN npm ci --only=production

# 将应用程序文件复制到镜像中
COPY . .

# 暴露应用运行的端口
EXPOSE 3000
# 根据你的应用调整端口号

# 设置启动命令（例如：npm run start 或 node server.js）
CMD ["node", "app.js"]