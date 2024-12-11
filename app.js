import fastify from 'fastify';
import path from "path"
import fastifyStatic from '@fastify/static'
import { fileURLToPath } from 'url';

const app = fastify({
  logger: true
});

const __filename = fileURLToPath(import.meta.url); // get the resolved path to the file
const __dirname = path.dirname(__filename); // get the name of the directory

app.register(fastifyStatic, {
	root: path.join(__dirname, '/uploads'),
})


app.get('/upload', async (request, reply) => {
	reply.sendFile("TipoA.xlsx", path.join(__dirname, '/uploads'))
})

app.listen({port: 3000}, (err, address) => {
  if(err) {
      app.log(err)
      console.log('Erro ocasionado')
      process.exit(1)
  }
  console.log('Server listening on port 3000')
})
