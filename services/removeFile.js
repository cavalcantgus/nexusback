const fs = require('fs');
const path = require('path');

// Função para buscar e retornar caminhos de arquivos
function searchFile(directory, fileName) {
  let results = []; // Array para armazenar os caminhos dos arquivos encontrados

  // Lê o conteúdo do diretório
  try {
    const files = fs.readdirSync(directory);

    // Itera sobre os arquivos e diretórios
    for (const file of files) {
      const fullPath = path.join(directory, file);

      // Verifica se o item é um diretório
      if (fs.statSync(fullPath).isDirectory()) {
        // Se for um diretório, faz a busca recursivamente
        const foundFiles = searchFile(fullPath, fileName);
        results = results.concat(foundFiles); // Adiciona os arquivos encontrados ao array de resultados
      } else {
        // Se for um arquivo, verifica se o nome bate com o arquivo que procuramos
        if (file === fileName) {
          results.push(fullPath); // Adiciona o caminho completo do arquivo encontrado
        }
      }
    }
  } catch (error) {
    console.error(`Erro ao ler o diretório ${directory}:`, error.message);
  }

  return results; // Retorna todos os caminhos encontrados
}

module.exports = searchFile