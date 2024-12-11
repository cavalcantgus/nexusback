// import { processExcelFile } from '../services/index.js';
// import fs from 'fs';
// import util from 'util';
// import { pipeline } from 'stream';
// import path from 'path';
// import { fileURLToPath } from 'url';

// const __filename = fileURLToPath(import.meta.url); // get the resolved path to the file
// const __dirname = path.dirname(__filename); // get the name of the directory
// const pump = util.promisify(pipeline);

// async function formatSheetHandler(req, rep) {
//     try {
//         const data = await req.file()
//         const fields = data.fields
//         const type = fields.report.value
//         if(!data) {
//           return rep.status(400).send({ error: "No file provided"})
//         }

//         const uploadDir = './uploads';
//         if(!fs.existsSync(uploadDir)) {
//           fs.mkdirSync(uploadDir);
//         }

//         const inputFilePath = path.join(uploadDir, data.filename);
//         await pump(data.file, fs.createWriteStream(inputFilePath));

//         const outputFileName = `formatted_${data.filename.replace(/\.[^/.]+$/, "")}.xlsx`;
//         const outputFilePath = path.join(uploadDir, outputFileName);
        
//         // Processar o arquivo Excel
//         await processExcelFile(inputFilePath, outputFilePath, type);
//         return rep.sendFile(outputFileName, path.join(__dirname, 'uploads'))
//       } catch (error) {
//         rep.status(500).send({ error: 'Failed to upload file' });
//       }
// }

// export default formatSheetHandler;