import * as React from 'react';
import { MsalProvider, useMsal } from '@azure/msal-react';
import { PublicClientApplication } from '@azure/msal-browser';

const pca = new PublicClientApplication({
  auth: {
    clientId: 'dbfefe9a-a7d6-45ce-8eee-2c3df73efe50', // Замените на ваш client ID из Entra ID
    authority: 'https://login.microsoftonline.com/8f719ff3-dda5-4884-bd32-692ccf5f0c54', // Замените на ваш tenant ID или используйте 'common' для multi-tenant
    redirectUri: 'http://localhost:3000' // Должно совпадать с redirect URI в Entra ID
  },
  cache: { cacheLocation: 'localStorage' }
});

export const AuthProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  return <MsalProvider instance={pca}>{children}</MsalProvider>;
};

export const useAuth = () => {
  const { instance } = useMsal();

  const getToken = async () => {
    try {
      // Проверяем наличие OfficeRuntime для SSO
      if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.auth) {
        const token = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
        return token;
      } else {
        console.warn('OfficeRuntime.auth not available, falling back to MSAL');
      }

      // Fallback to MSAL login
      const account = instance.getActiveAccount();
      if (!account) {
        const loginResponse = await instance.loginPopup({
          scopes: ['openid', 'profile', 'email', 'Mail.Read', 'Mail.ReadWrite', 'Files.ReadWrite']
        });
        instance.setActiveAccount(loginResponse.account);
      }

      const response = await instance.acquireTokenSilent({
        scopes: ['Mail.Read', 'Mail.ReadWrite', 'Files.ReadWrite'],
        account: instance.getActiveAccount()!
      });
      return response.accessToken;
    } catch (error) {
      console.error('Error getting token:', error);
      throw error;
    }
  };

  return { getToken };
};