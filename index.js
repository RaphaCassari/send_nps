const AWS = require('aws-sdk');
const SES = new AWS.SES();
const xlsx = require('xlsx');
const fs = require('fs');

exports.handler = async (event) => {
  // Verificar se a requisição é uma requisição de preflight (OPTIONS)
  if (event.httpMethod === 'OPTIONS') {
    return {
      statusCode: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Access-Control-Allow-Methods': 'POST,OPTIONS',
      },
      body: JSON.stringify({ message: 'Preflight check successful' }),
    };
  }

  // Habilitar CORS na resposta inicial
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST,OPTIONS',
  };

  try {
    const body = JSON.parse(event.body);
    const base64File = body.file;  // O arquivo .xlsx em base64

    if (!base64File) {
      return {
        statusCode: 400,
        headers,  // Inclui os cabeçalhos CORS
        body: JSON.stringify({ message: 'Nenhum arquivo fornecido.' }),
      };
    }

    const filePath = '/tmp/input.xlsx';
    fs.writeFileSync(filePath, Buffer.from(base64File, 'base64'));

    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    const urlRegex = /(https?:\/\/[^\s]+)/g;

    for (let i = 1; i < data.length; i++) {
        const email = data[i][5];  // Coluna F (índice 5) contém o email
        let mensagem = data[i][8];  // Coluna H (índice 7) contém a mensagem

        if (email && mensagem) {
            const urls = mensagem.match(urlRegex);
            let url = '';

            if (urls && urls.length > 0) {
            url = urls[0];
            mensagem = mensagem.replace(url, `
                <a href="${url}" style="display: inline-block; padding: 10px 20px; font-size: 16px; color: white; background-color: #582BD9; text-decoration: none; border-radius: 5px;">Responder Pesquisa</a>
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
                //BccAddresses: ['pmo@fiveacts.com.br'],  // Adicionar o PMO como cópia oculta (BCC)
            },
            Message: {
                Body: {
                Html: {
                    Charset: 'UTF-8',
                    Data: "htmlBody",
                },
                },
                Subject: {
                Charset: 'UTF-8',
                Data: 'NPS - Five Acts',
                },
            },
            Source: 'PMO <pmo@fiveacts.com.br>',
            };

            try {
            await SES.sendEmail(params).promise();
            console.log(`E-mail enviado com sucesso para: ${email}`);
            } catch (error) {
            console.error(`Falha ao enviar e-mail para: ${email}, erro: ${error.message}`);
            }
        }
    }

    return {
      statusCode: 200,
      headers,  // Inclui os cabeçalhos CORS
      body: JSON.stringify({ message: 'Envio de e-mails concluído' }),
    };

  } catch (error) {
    console.error('Erro no processamento:', error);

    return {
      statusCode: 500,
      headers,  // Inclui os cabeçalhos CORS
      body: JSON.stringify({ error: 'Erro ao processar o arquivo ou enviar e-mails.' }),
    };
  }
};
