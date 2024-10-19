const AWS = require('aws-sdk');
const SES = new AWS.SES();
const xlsx = require('xlsx');
const fs = require('fs');

exports.handler = async (event) => {
  const from = 'PMO <pmo@fiveacts.com.br>';
  const pmoEmail = 'pmo@fiveacts.com.br';

  // Verificar se o evento contém o arquivo base64
  try {
    const body = JSON.parse(event.body);
    const base64File = body.file;  // O arquivo .xlsx em base64

    // Verificar se o arquivo foi enviado
    if (!base64File) {
      return {
        statusCode: 400,
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Access-Control-Allow-Headers': 'Content-Type',
          'Access-Control-Allow-Methods': 'POST,OPTIONS',
        },
        body: JSON.stringify({ message: 'Nenhum arquivo fornecido.' }),
      };
    }

    // Caminho temporário onde o arquivo será salvo
    const filePath = '/tmp/input.xlsx';

    // Escrever o arquivo base64 no diretório temporário
    fs.writeFileSync(filePath, Buffer.from(base64File, 'base64'));

    // Carregar o arquivo .xlsx
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    // Converter a planilha para JSON
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    // Regex para identificar URLs no texto
    const urlRegex = /(https?:\/\/[^\s]+)/g;

    // Iterar pelas linhas da planilha, começando pela segunda linha (ignorando o cabeçalho)
    for (let i = 1; i < data.length; i++) {
      const email = data[i][5];  // Coluna F (índice 5) contém o email
      let mensagem = data[i][8];  // Coluna H (índice 7) contém a mensagem

      // Verificar se o email e a mensagem estão presentes
      if (email && mensagem) {
        // Procurar pela URL no corpo da mensagem
        const urls = mensagem.match(urlRegex);
        let url = '';
        
        // Se houver uma URL, capture a primeira
        if (urls && urls.length > 0) {
          url = urls[0];
          // Substituir a URL pelo botão "Responder Pesquisa"
          mensagem = mensagem.replace(url, `
            <a href="${url}" style="display: inline-block; padding: 10px 20px; font-size: 16px; color: white; background-color: #4CAF50; text-decoration: none; border-radius: 5px;">Responder Pesquisa</a>
          `);
        }

        const htmlBody = `
          <html>
            <body>
              <p>${mensagem.replace(/\n/g, '<br>')}</p>
            </body>
          </html>`;

        const params = {
          Destination: {
            ToAddresses: [email],
            CcAddresses: [pmoEmail],  // Adicionar o PMO em cópia
          },
          Message: {
            Body: {
              Html: {
                Charset: 'UTF-8',
                Data: htmlBody,  // Enviar o HTML com o botão
              },
            },
            Subject: {
              Charset: 'UTF-8',
              Data: 'NPS - Five Acts',  // Assunto do e-mail
            },
          },
          Source: from,
        };

        // Enviar o e-mail usando SES
        try {
          await SES.sendEmail(params).promise();
          console.log(`E-mail enviado com sucesso para: ${email}`);
        } catch (error) {
          console.error(`Falha ao enviar e-mail para: ${email}, erro: ${error.message}`);
        }
      }
    }

    // Resposta de sucesso
    return {
      statusCode: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',  // Habilitar CORS
        'Access-Control-Allow-Headers': 'Content-Type',
        'Access-Control-Allow-Methods': 'POST,OPTIONS',
      },
      body: JSON.stringify({ message: 'Envio de e-mails concluído' }),
    };

  } catch (error) {
    console.error('Erro no processamento:', error);

    return {
      statusCode: 500,
      headers: {
        'Access-Control-Allow-Origin': '*',  // Habilitar CORS mesmo em caso de erro
        'Access-Control-Allow-Headers': 'Content-Type',
        'Access-Control-Allow-Methods': 'POST,OPTIONS',
      },
      body: JSON.stringify({ error: 'Erro ao processar o arquivo ou enviar e-mails.' }),
    };
  }
};
