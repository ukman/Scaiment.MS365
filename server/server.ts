import express, { Request, Response } from 'express';
import axios from 'axios';
import dotenv from 'dotenv';
import cors from 'cors';
import https from 'https';
import { getHttpsServerOptions } from 'office-addin-dev-certs';
import { ConfidentialClientApplication } from '@azure/msal-node';
import * as XLSX from 'xlsx';
import { AzureOpenAI } from 'openai';
import { Attachment, FileAttachment, Message } from '@microsoft/microsoft-graph-types';
import { Client } from '@microsoft/microsoft-graph-client';

dotenv.config();


const msalConfig = {
    auth: {
      clientId: process.env.CLIENT_ID!,
      authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
      clientSecret: process.env.CLIENT_SECRET!,
    },
  };
const cca = new ConfidentialClientApplication(msalConfig);


interface ChatCompletionRequest {
    messages: Array<{ role: 'system' | 'user' | 'assistant'; content: string }>;
    model: string;
    max_tokens?: number;
    temperature?: number;
}

interface ChatCompletionResponse {
    choices: Array<{ message: { content: string } }>;
    id: string;
    model: string;
    usage: { prompt_tokens: number; completion_tokens: number; total_tokens: number };
}

async function startServer() {

    const app = express();
    app.use(cors({ origin: ['https://localhost:3000'] })); // Разрешить ваш frontend
    app.use(express.json());


    app.post('/api/openai/chat', handleOpenAIRequest);
    // app.get('/api/openai/chat', handleOpenAIRequest);

    app.post('/emails', async (req: Request, res: Response) => {
        const reqData = await req.body;
        console.log("Email attachments request = ", reqData);
        const result = await getOutlookMessageById(reqData.accessToken, reqData.messageId);
        console.log("Email attachments result = ", result);

        res.json(result);
    });

    app.get('/health', async (_req: Request, res: Response) => {
        res.json({status:"ok"});
    });

    // OAuth Callback
    app.get('/auth/callback', async (req, res) => {
        console.log("Auth callback started");
        const tokenRequest = {
            code: req.query.code as string,
            scopes: ['Mail.Read', 'Mail.Send', 'User.Read'],
            redirectUri: 'https://localhost:3001/auth/callback',
        };
        console.log("Token request created " + tokenRequest.code);
        try {
            const response = await cca.acquireTokenByCode(tokenRequest);
            const { accessToken } = response!;
            console.log("accessToken = " + accessToken);
            // Redirect с token в fragment (безопасно)
            // res.redirect(`https://localhost:3000/dialog-close.html#access_token=${accessToken}&refresh_token=${refreshToken || ''}`);
            res.redirect(`https://localhost:3000/dialog-close.html#access_token=${accessToken}&refresh_token=$ {refreshToken || ''}`);
            // Опционально: Сохраните refreshToken в session для будущего refresh
        } catch (error) {
            console.error(error);
            res.status(500).send('Auth error');
        }
    });

    // Получение сертификатов
    const httpsOptions = await getHttpsServerOptions();

    const port = PORT || 3001;

    const server = https.createServer(httpsOptions, app);
    server.listen(port, () => {
        console.log(`Backend proxy running on https://localhost:${port}`);
    });

    app.listen(port, () => {
    console.log(`Backend proxy running on port ${port}`);
    });
}

const { AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_API_KEY, AZURE_OPENAI_DEPLOYMENT_NAME, AZURE_OPENAI_API_VERSION, PORT } = process.env;

async function handleOpenAIRequest(req: Request, res: Response) {
    console.log("handleOpenAIRequest ")
    const requestBody: ChatCompletionRequest = req.body;
    /*
    const requestBody: ChatCompletionRequest = {
        messages:[{role: "user", content: "Назови первые 10 простых чисел"}],
        model:"gpt-4.1",
        temperature:0.1,
        max_tokens:100
    };
    */

    // Настройка клиента Azure OpenAI
    const client = new AzureOpenAI({
        apiKey: process.env.AZURE_OPENAI_API_KEY, // Ваш API ключ
        endpoint: process.env.AZURE_OPENAI_ENDPOINT, // https://your-resource.openai.azure.com
        apiVersion: process.env.AZURE_OPENAI_API_VERSION // '2024-02-01', // Версия API
    });

    try {
        console.log("AI Input : " + JSON.stringify(requestBody, null, 2));
        const response = await client.chat.completions.create(requestBody);
        /*
        ({
          model: 'gpt-4', // Имя вашего развернутого деплоймента в Azure
          messages: [
            {
              role: 'system',
              content: 'Ты полезный ассистент, который отвечает на русском языке.'
            },
            {
              role: 'user',
              content: 'Сколько будет пять плюс 10 и все в квадрате?'
            }
          ],
          max_tokens: 500,
          temperature: 0.3,
        });
        */
       console.log("AI Output : " + JSON.stringify(response, null, 2));
       res.json(response);

       
       try {
        console.log("response.choices[0]?.message?.content = ");
        console.log(response.choices[0]?.message?.content);
        const content = JSON.parse(response.choices[0]?.message?.content);

            console.log(JSON.stringify(content, null, 2));
       }catch(e) {
            console.error("Cannot parse content \n" + e);
       }
    
        // TypeScript автоматически определит типы
        // const assistantMessage = response.choices[0]?.message?.content;
        // return response; // assistantMessage || 'Ответ не получен';
        
    } catch (error) {
        console.error('Ошибка при обращении к Azure OpenAI:', error);
        throw error;
    }

    /*
    try {
        const response = await axios.post<ChatCompletionResponse>(
            `${AZURE_OPENAI_ENDPOINT}/openai/deployments/${AZURE_OPENAI_DEPLOYMENT_NAME}/chat/completions?api-version=${AZURE_OPENAI_API_VERSION}`,
            requestBody,
            {
                headers: {
                    'Content-Type': 'application/json',
                    'api-key': AZURE_OPENAI_API_KEY,
                },
            }
        );
        res.json(response.data);
    } catch (error: any) {
        console.error('Proxy error:', error.message);
        res.status(error.response?.status || 500).json({ error: 'Failed to call Azure OpenAI API' });
    }
    */
}

export interface ExcelFile {
    fileName : string;
    csv : string;
}

async function getOutlookMessageById(
    accessToken: string, 
    messageId: string, 
    userId: string = 'me'
): Promise<ExcelFile[]> {
    
    // 1. Инициализация клиента Microsoft Graph
    const client = Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });

    try {


        // 2. Формирование запроса к Graph API
        // Запрос к: /users/{userId}/messages/{messageId}
        const message = await client
            .api(`/${userId}/messages/${messageId}`)
            // Используем .get() для выполнения GET-запроса
            .query({ '$expand': 'attachments' })
            .get();

        console.log(`Письмо с ID ${messageId} успешно получено.`);
        
        
        // 3. Проверка и параллельная загрузка содержимого вложений
        if (message.attachments && message.attachments.length > 0) {
            console.log(`Найдено ${message.attachments.length} вложений. Загрузка содержимого...`);
            
            // Создаем массив промисов для параллельного выполнения запросов
            const fetchContentPromises = message.attachments.map(async (att: Attachment) => {
                
                // Только FileAttachment содержит двоичные данные. 
                // ItemAttachment (другое письмо или контакт) или ReferenceAttachment (ссылка на OneDrive) обрабатываются иначе.
                if (att['@odata.type'] !== '#microsoft.graph.fileAttachment') {
                    console.warn(`Вложение ID ${att.id} имеет тип ${att['@odata.type']} (не FileAttachment). Пропускаем загрузку содержимого.`);
                    return att; // Возвращаем метаданные без загрузки
                }

                // URL для получения полного объекта вложения, который включает contentBytes
                const attachmentPath = `/users/${userId}/messages/${messageId}/attachments/${att.id}`;
                
                try {
                    // Делаем отдельный запрос, чтобы получить поле contentBytes
                    const fullAttachment = await client.api(attachmentPath).get();
                    
                    // Возвращаем полный объект (теперь с contentBytes: string (Base64))
                    return fullAttachment as FileAttachment; 
                } catch (fetchError) {
                    console.error(`Ошибка при получении содержимого вложения ID ${att.id}.`, fetchError);
                    return att; // В случае ошибки возвращаем только метаданные
                }
            });

            // Ждем завершения всех запросов на загрузку содержимого
            const fullAttachments = await Promise.all(fetchContentPromises);

            const res = await Promise.all(fullAttachments.filter(a => a.name.endsWith(".xlsx")).map(a => convertExcelToCSV(a)));
            
            // 4. Обновляем объект сообщения полными данными вложений
            // message.attachments = fullAttachments as Attachment[];
            return res;
        }



        return [];

    } catch (error) {
        console.error(`Ошибка при чтении письма с ID ${messageId}:`, error);
        // Выбрасываем ошибку для обработки вызывающим кодом
        throw new Error(`Не удалось получить письмо: ${error instanceof Error ? error.message : String(error)}`);
    }
}

/**
 * Преобразует Excel файл из FileAttachment в CSV текст (строку).
 * Предполагается, что файл - это Excel (.xlsx или .xls).
 * Конвертирует только первый лист.
 * 
 * @param attachment - FileAttachment из Graph API (attachment из email)
 * @returns Promise<string> - CSV текст
 */
async function convertExcelToCSV(attachment: FileAttachment): Promise<ExcelFile> {
    // Получаем содержимое attachment как base64
    const base64Content = attachment.contentBytes;
    
    if (!base64Content) {
      throw new Error('Attachment content is empty');
    }
  
    // Декодируем base64 в Buffer (удаляем префикс data: если есть, но в Graph API это чистый base64)
    const buffer = Buffer.from(base64Content, 'base64');
  
    // Парсим Excel файл
    const workbook = XLSX.read(buffer, { type: 'buffer' });
  
    if (workbook.SheetNames.length === 0) {
      throw new Error('No sheets found in the Excel file');
    }
  
    // Берем первый лист
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
  
    // Конвертируем в CSV
    const csv = XLSX.utils.sheet_to_csv(worksheet);
  
    return {fileName : attachment.name, csv};
  }

startServer().catch((error) => {
    console.error('Failed to start server:', error);
});