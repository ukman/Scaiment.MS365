import * as React from "react";
import { useState } from 'react';
import { PublicClientApplication } from '@azure/msal-browser';
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react'; // Если используете MSAL React
import { PrimaryButton } from "@fluentui/react";
import { excelLog } from "../../util/Logs";
import { MSGraphService } from "../../services/MSGraphService";
import { PersonService } from "../../services/PersonService";

// MSAL config (в env или const)
const msalConfig = {
  auth: {
    clientId: 'dbfefe9a-a7d6-45ce-8eee-2c3df73efe50',
    authority: 'https://login.microsoftonline.com/8f719ff3-dda5-4884-bd32-692ccf5f0c54',
    redirectUri: 'https://localhost:3000/dialog-close.html', // Финальный redirect
  },
};
const msalInstance = new PublicClientApplication(msalConfig); // Глобально или в context

const authenticate = async () => {
    (await MSGraphService.getInstance()).authenticate();
}

// Функция auth (добавьте кнопку в return)
const authenticate2 = async () => {
    try {
      await msalInstance.initialize(); // Если не инициализировано
      excelLog("Initialized ");
      Office.context.ui.displayDialogAsync(
        'https://localhost:3000/dialog-start.html', // В вашем домене
        { height: 60, width: 30 }, 
        (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            excelLog("Error 1");
            excelLog("" + result.error.message);
            console.error(result.error.message);
            return;
          }
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
            dialog.close();
            excelLog("Message received : " + arg.message);
            const message = JSON.parse(arg.message);
            if (message.status === 'success') {
              const { accessToken, refreshToken } = message;
              // Сохраните в localStorage или state
              localStorage.setItem('accessToken', accessToken);
              localStorage.setItem('refreshToken', refreshToken); // Для refresh
              callGraphAPI(accessToken); // Используйте для Outlook
            } else {
              console.error('Auth failed:', message.error);
              
              // excelLog("Error : " + message.error);
            }
          });
        }
      );
    } catch (error) {
      console.error(error);
      // excelLog("Error : " + error);
    }
  };

  async function callGraphAPI(token: string) {
    // const response = await fetch('https://graph.microsoft.com/v1.0/me/messages', {
    const response = await fetch('https://graph.microsoft.com/v1.0/me', {
            headers: { Authorization: `Bearer ${token}` },
    });
    const emails = await response.json();
    excelLog("Email = " + emails);
    excelLog("Email = " + JSON.stringify(emails));
    // Или отправка: POST /me/sendMail
    // Для Outlook: Используйте Graph для чтения/отправки email
  }

export interface AuthButtonProps {
    title: string;
}
  
const AuthButton: React.FC<AuthButtonProps> = (props: AuthButtonProps) => {
    const { title } = props;
    const [ data, setData] = useState<string>('---');

    const initAsync = async () => {
        await MSGraphService.getInstance().validateSession();
        const user = MSGraphService.getInstance().getCurrentUser();
        if(user) {
            setData(user.displayName + " - " + user.mail + " - " + user.userPrincipalName);
        } else {
            setData("Not authenticated");
        }
    }
    initAsync(); 

    const getCurrentUser = async() => {
        setData("Fetching ");
        try {
            const isAuthenticated = MSGraphService.getInstance().isAuthenticated();
            if(isAuthenticated) {
                const user = MSGraphService.getInstance().getCurrentUser();
                if(user) {
                    setData(user.displayName + " - " + user.mail + " - " + user.userPrincipalName);
                } else {
                    setData("User is undefined 22");
                }
            } else {
                await MSGraphService.getInstance().validateSession();
                setData("Not authenticated " + localStorage.getItem("__accessToken"));
            }          


        }catch(e) {
            setData("Error : " + e);
            excelLog("Error " + e);
        }
    }
    
    
    return (
        <div>
            Test {title}
            <PrimaryButton onClick={authenticate}>Authenticate</PrimaryButton>
            <PrimaryButton onClick={getCurrentUser}>Get Current User</PrimaryButton>
            <pre><code>{data}</code></pre>
        </div>
    );
}

export default AuthButton;