import { Configuration, PublicClientApplication } from "@azure/msal-browser";
import { User, Message } from '@microsoft/microsoft-graph-types';
import { excelLog } from "../util/Logs";
import { ChatCompletion } from "openai/resources/chat/completions/completions";
import { Requisition } from "../util/data/DBSchema";
import { notStrictEqual } from "assert";

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

export type NoteLevel = "info" | "warning" | "error";

export interface ContentResponseNote {
    level : NoteLevel;
    text : string;
}

export interface ContentResponse {
    requisition : Requisition;
    notes : ContentResponseNote[];
    replyMessage : string;
}

export class AIService {

    private constructor () {
    }

    private static instance : AIService;

    public static getInstance() : AIService {
        if(!this.instance) {
            
            this.instance = new AIService();
            // await this.instance.validateSession();
        }
        return this.instance;
    }

    public async callAI(messageToAI : string) : Promise<ChatCompletion> {
        const requestBody: ChatCompletionRequest = {
            messages:[{role: "user", content: messageToAI}],
            model:"gpt-4.1",
            temperature:0.1,
            max_tokens:3000
        };
        (requestBody as any).response_format= {"type": "json_object"};
    
        try {
//            const response = await fetch("https://localhost:3000/taskpane.html", {
            excelLog("Start AI request....");
            const response = await fetch("https://localhost:3001/api/openai/chat", {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(requestBody),
            });  
            await excelLog("AI request is frinished");
            await excelLog("AI request is frinished status = " + response.status);
            const json = await response.json();
            await excelLog("Response status = " + response.status + " text = " + json);      
            return json;
        }catch(e) {
            await excelLog("Error = " + e);  
            throw e;    
        }

    }
}