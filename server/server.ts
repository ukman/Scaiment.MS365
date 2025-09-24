import express, { Request, Response } from 'express';
import axios from 'axios';
import dotenv from 'dotenv';
import cors from 'cors';
import https from 'https';
import { getHttpsServerOptions } from 'office-addin-dev-certs';
import { ConfidentialClientApplication } from '@azure/msal-node';
import * as XLSX from 'xlsx';
import { AzureOpenAI } from 'openai';

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

    app.get('/health', async (_req: Request, res: Response) => {
        res.json({status:"ok"});
    });

    // Callback
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

startServer().catch((error) => {
    console.error('Failed to start server:', error);
});