import { Configuration, PublicClientApplication } from "@azure/msal-browser";
import { User, Message } from '@microsoft/microsoft-graph-types';
import { excelLog } from "../util/Logs";


export interface ExcelFile {
    fileName : string;
    csv : string;
}

export class MSGraphService {

    private static readonly ACCESS_TOKEN_KEY = "__accessToken";
    private static readonly REFRESH_TOKEN_KEY = "__refreshToken";

    private accessToken : string;
    private signedIn : boolean = false;
    private currentUser : User;

    private constructor () {
        this.accessToken = localStorage.getItem(MSGraphService.ACCESS_TOKEN_KEY);
        this.validateSession();
    }

    private static instance : MSGraphService;

    public static getInstance() : MSGraphService {
        if(!this.instance) {
            this.instance = new MSGraphService();
            // await this.instance.validateSession();
        }
        return this.instance;

    }

// Функция auth (добавьте кнопку в return)
    public async authenticate() {

        const msalConfig : Configuration = {
            auth: {
                clientId: 'dbfefe9a-a7d6-45ce-8eee-2c3df73efe50',
                authority: 'https://login.microsoftonline.com/8f719ff3-dda5-4884-bd32-692ccf5f0c54',
                redirectUri: 'https://localhost:3000/dialog-close.html', // Финальный redirect
            },
        };

        const msalInstance = new PublicClientApplication(msalConfig); // Глобально или в context

        try {
            await msalInstance.initialize(); // Если не инициализировано
            Office.context.ui.displayDialogAsync(
                'https://localhost:3000/dialog-start.html', // В вашем домене
                    { height: 60, width: 30 }, 
                (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error(result.error.message);
                    return;
                }
                const dialog = result.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
                    dialog.close();
                    const message = JSON.parse(arg.message);
                    if (message.status === 'success') {
                        const { accessToken, refreshToken } = message;
                        // Сохраните в localStorage или state
                        this.accessToken = accessToken;
                        localStorage.setItem(MSGraphService.ACCESS_TOKEN_KEY, accessToken);
                        localStorage.setItem(MSGraphService.REFRESH_TOKEN_KEY, refreshToken); // Для refresh
                        excelLog(`refreshToken = ${refreshToken}`);
                    } else {
                        console.error('Auth failed:', message.error);
                    }
                });
                }
            );
        } catch (error) {
            console.error(error);
        }

    };

    public isAuthenticated() : boolean {
        return typeof(this.currentUser) !== "undefined";
    }

    public async validateSession() : Promise<boolean> {
        try {
        // await excelLog("Start validate session");
        this.accessToken = localStorage.getItem(MSGraphService.ACCESS_TOKEN_KEY);
        if(typeof(this.currentUser) == 'undefined') {
            try {
                await this.refreshCurrentUser();
            } catch(e) {
                await excelLog("Error validating session" + e);
                console.error("Error validating session", e);
            }
        }
        return this.isAuthenticated();
        } catch(e) {
            await excelLog("Error " + e);
            throw e;
        }
    }

    private async refreshCurrentUser() {
        // const response = await fetch('https://graph.microsoft.com/v1.0/me/messages', {

        this.currentUser = undefined;

        const response = await fetch('https://graph.microsoft.com/v1.0/me', {
            headers: { 
                Authorization: `Bearer ${this.accessToken}` 
            },
        });

        if (!response.ok) {
            throw new Error(`Graph API error: ${response.status} ${response.statusText}`);
        }    
        this.currentUser = await response.json();    
    }

    public async getMessages(fromEmails? : string[]) : Promise<Message[]> {
        // ?$filter=(from/emailAddress/address eq 'a@domain.com' or from/emailAddress/address eq 'b@domain.com')

        let filterString = "";
        if(fromEmails) {
            const cond = fromEmails.filter(s => s.length > 0).map(email => `(from/emailAddress/address eq '${ email }')`).join(" or ");
            filterString = `$filter=(${ cond } and (receivedDateTime ge 2025-08-22T00:00:00Z))&`
        }
        excelLog("filterString = " + filterString);
        // filterString = '';

        // const url = `https://graph.microsoft.com/v1.0/me/messages?${filterString}$top=125&$orderby=receivedDateTime desc`;
        const url = `https://graph.microsoft.com/v1.0/me/messages?${filterString}$top=125`; 
        excelLog(`URL = ${url}`);
        const response = await fetch(url, {
            headers: {
                Authorization: `Bearer ${this.accessToken}`, 
                'Content-Type': 'application/json',
                'Prefer': 'IdType="ImmutableId"'
            }
          });
        if (!response.ok) {
            const text = await response.text();
            excelLog(`Graph API error: ${response.status} ${response.statusText} ${text}`); 
            throw new Error(`Graph API error: ${response.status} ${response.statusText} ${text}`);
        }
        const res = await response.json();
        return res.value as Message[];
    }

    public getCurrentUser() : User {
        return this.currentUser;
    }

    public async getMailAttachments(emailId : string) : Promise<ExcelFile[]> {
        const response = await fetch("https://localhost:3001/emails", {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                accessToken: this.accessToken, 
                messageId: emailId
            }),
        });  
        if(!response.ok) {
            throw new Error("Error getting attachments " + response.text);
        }
        const result = await response.json();
        excelLog("attachments = ", result);
        return result;
    }
}