
const fastify = require('fastify')();
const fastifyCors = require('@fastify/cors');
const multipart = require('@fastify/multipart');
const fastifyStatic = require('@fastify/static');
const path = require('path');
const fs = require('fs');
const util = require('util');
const { pipeline } = require('stream');
const { processExcelFile } = require('./services/index.js');
const pump = util.promisify(pipeline);
const searchFile = require('./services/removeFile.js');

fastify.register(fastifyStatic, {
    root: path.join(__dirname, '/uploads'),
});

// Configurações do CORS
fastify.register(fastifyCors, {
    origin: '*', // Permitir todas as origens.
});

// Registrar o plugin multipart
fastify.register(multipart);

// Função para remover arquivos com tentativas
const unlinkWithRetry = (filePath, retries = 5, delay = 1000) => {
    return new Promise((resolve, reject) => {
        const attemptUnlink = (retriesLeft) => {
            fs.unlink(filePath, (err) => {
                if (err) {
                    if (err.code === 'EBUSY' && retriesLeft > 0) {
                        console.log(`Tentando remover ${filePath} novamente... (${retriesLeft} tentativas restantes)`);
                        setTimeout(() => attemptUnlink(retriesLeft - 1), delay); // Tenta novamente após um delay
                    } else {
                        reject(err); // Se for um erro diferente de EBUSY, rejeita a promise
                    }
                } else {
                    resolve(); // Sucesso na remoção
                }
            });
        };

        attemptUnlink(retries);
    });
};

// Rota para upload de arquivos
fastify.post('/upload', async function (req, rep) {
    let inputFilePath;
    let outputFilePath;
    const directoryToSearch = '/app/downloads';
    let foundFiles = [];

    try {
        const data = await req.file();
        const fields = data.fields;
        const type = fields.report.value;

        if (!data) {
            return rep.status(400).send({ error: "No file provided" });
        }

        const uploadDir = './uploads';
        if (!fs.existsSync(uploadDir)) {
            fs.mkdirSync(uploadDir);
        }

        inputFilePath = path.join(uploadDir, data.filename);
        await pump(data.file, fs.createWriteStream(inputFilePath));

        const outputFileName = `formatted_${data.filename.replace(/\.[^/.]+$/, "")}.xlsx`;
        outputFilePath = path.join(uploadDir, outputFileName);

        // Processar o arquivo Excel
        await processExcelFile(inputFilePath, outputFilePath, type);
        
        // Chamar a função de busca
        foundFiles = searchFile(directoryToSearch, data.filename);

        // Enviar o arquivo formatado como resposta
        const response = await rep.sendFile(outputFileName, path.join(__dirname, 'uploads'));

        // Chamada da função de remoção fora do POST
        setImmediate(async () => {
            if (foundFiles.length > 0) {
                for (const filePath of foundFiles) {
                    try {
                        await unlinkWithRetry(filePath); // Usa a função com retry para remover o arquivo
                        console.log(`Arquivo removido: ${filePath}`);
                    } catch (err) {
                        console.error(`Erro ao remover ${filePath}:`, err.message);
                    }
                }
            } else {
                console.log('Arquivo não encontrado para remoção.');
            }
        });

        return response; // Retorna a resposta
    } catch (error) {
        console.error(error); // Log do erro
        rep.status(500).send({ error: 'Failed to upload file' });
    }
});

// Iniciar o servidor
fastify.listen({ port: 3000, host: '0.0.0.0' }, (err, address) => {
    if (err) {
        console.error(err);
        process.exit(1);
    }
    console.log(`Server listening on ${address}`);
});
