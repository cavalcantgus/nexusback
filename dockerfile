# Use uma imagem base do Node.js (ou da sua linguagem/framework)
FROM node:18

# Crie e defina o diretório de trabalho dentro do container
WORKDIR /app

# Copie os arquivos do projeto para o diretório de trabalho no container
COPY . .

# Instale as dependências do projeto
RUN npm install

# Exponha a porta em que a aplicação roda
EXPOSE 3000

# Comando para rodar a aplicação
CMD ["npm", "start"]
