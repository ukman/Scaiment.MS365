import React, { useState, useEffect } from 'react';
import { PublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser';
import { Card, Button, Alert, Spinner, Badge, Row, Col } from 'react-bootstrap';
import { Icon } from '@fluentui/react';

// Интерфейсы для типизации
interface EmailData {
  id: string;
  subject: string;
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  } | null;
  receivedDateTime: string;
  isRead: boolean;
  importance: 'low' | 'normal' | 'high';
  bodyPreview: string;
  hasAttachments: boolean;
}

interface OutlookEmailsProps {
  clientId: string; // Ваш Azure App Registration Client ID
}

// Конфигурация MSAL для Office Add-ins
const msalConfig = {
  auth: {
    clientId: '', // Будет установлен через props
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: 'https://localhost:3000/taskpane.html' // Конкретная страница Add-in
  },
  cache: {
    cacheLocation: 'localStorage', // Изменено на localStorage для Add-ins
    storeAuthStateInCookie: true, // Включено для лучшей совместимости
  },
  system: {
    allowNativeBroker: false, // Отключаем native broker для Add-ins
    windowHashTimeout: 60000,
    iframeHashTimeout: 6000,
    loadFrameTimeout: 0,
    loggerOptions: {
      logLevel: 3, // Error level для отладки
      loggerCallback: (level: any, message: string, _containsPii: boolean) => {
        console.log(`MSAL [${level}]: ${message}`);
      }
    }
  }
};

const loginRequest = {
  scopes: ['https://graph.microsoft.com/Mail.Read'],
  prompt: 'select_account' // Принудительно показываем выбор аккаунта
};

const OutlookEmails: React.FC<OutlookEmailsProps> = ({ clientId }) => {
  const [emails, setEmails] = useState<EmailData[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>('');
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [msalInstance, setMsalInstance] = useState<PublicClientApplication | null>(null);

  // Инициализация MSAL
  useEffect(() => {
    const initializeMsal = async () => {
      const config = {
        ...msalConfig,
        auth: {
          ...msalConfig.auth,
          clientId: clientId
        }
      };
      
      try {
        const instance = new PublicClientApplication(config);
        await instance.initialize();
        
        // Обработка redirect после авторизации
        console.log('Checking for redirect response...');
        const response = await instance.handleRedirectPromise();
        console.log('Redirect response:', response);
        
        if (response !== null) {
          console.log('Successful redirect auth, setting active account');
          instance.setActiveAccount(response.account);
          setIsAuthenticated(true);
          // Автоматически загружаем письма после успешной авторизации
          setTimeout(() => fetchEmailsAfterAuth(instance), 1000);
        } else {
          // Проверяем, есть ли активный аккаунт
          const accounts = instance.getAllAccounts();
          if (accounts.length > 0) {
            instance.setActiveAccount(accounts[0]);
            setIsAuthenticated(true);
          }
        }
        
        setMsalInstance(instance);
      } catch (error: any) {
        console.error('MSAL initialization error:', error);
        setError(`Ошибка инициализации: ${error.message}`);
      }
    };

    if (clientId) {
      initializeMsal();
    }
  }, [clientId]);

  // Функция для получения токена доступа
  const getAccessToken = async (instance?: PublicClientApplication): Promise<string> => {
    const msalInst = instance || msalInstance;
    if (!msalInst) {
      throw new Error('MSAL не инициализирован');
    }

    try {
      const account = msalInst.getActiveAccount();
      if (!account) {
        throw new Error('Нет активного аккаунта');
      }

      const silentRequest = {
        ...loginRequest,
        account: account
      };

      const response = await msalInst.acquireTokenSilent(silentRequest);
      return response.accessToken;
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        // Используем redirect вместо popup для Add-ins
        await msalInst.acquireTokenRedirect(loginRequest);
        throw new Error('Перенаправление на авторизацию...');
      }
      throw error;
    }
  };

  // Функция для входа в систему (используем redirect)
  const handleLogin = async () => {
    if (!msalInstance) return;

    try {
      setLoading(true);
      setError('');
      
      // Используем loginRedirect для Office Add-ins
      await msalInstance.loginRedirect(loginRequest);
    } catch (error: any) {
      setError(`Ошибка входа: ${error.message}`);
      setLoading(false);
    }
  };

  // Функция для получения писем после авторизации
  const fetchEmailsAfterAuth = async (instance: PublicClientApplication) => {
    try {
      setLoading(true);
      setError('');
      
      const accessToken = await getAccessToken(instance);
      
      const response = await fetch('https://graph.microsoft.com/v1.0/me/messages?$top=5&$orderby=receivedDateTime desc', {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const data = await response.json();
      setEmails(data.value || []);
    } catch (error: any) {
      if (error.message !== 'Перенаправление на авторизацию...') {
        setError(`Ошибка загрузки писем: ${error.message}`);
      }
    } finally {
      setLoading(false);
    }
  };

  // Функция для получения писем
  const fetchEmails = async () => {
    if (!isAuthenticated || !msalInstance) return;
    await fetchEmailsAfterAuth(msalInstance);
  };

  // Функция выхода
  const handleLogout = async () => {
    if (!msalInstance) return;
    
    try {
      const account = msalInstance.getActiveAccount();
      if (account) {
        await msalInstance.logoutRedirect({
          account: account
        });
      }
    } catch (error: any) {
      console.error('Logout error:', error);
    }
  };

  // Функция для форматирования даты
  const formatDate = (dateString: string): string => {
    const date = new Date(dateString);
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    const diffDays = Math.floor(diffHours / 24);

    if (diffHours < 1) return 'Только что';
    if (diffHours < 24) return `${diffHours} ч. назад`;
    if (diffDays === 1) return 'Вчера';
    if (diffDays < 7) return `${diffDays} дн. назад`;
    
    return date.toLocaleDateString('ru-RU', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric'
    });
  };

  // Получение иконки для важности письма
  const getImportanceIcon = (importance: string) => {
    switch (importance) {
      case 'high':
        return <Icon iconName="Important" style={{ color: '#d13438', marginRight: '4px' }} />;
      case 'low':
        return <Icon iconName="Down" style={{ color: '#107c10', marginRight: '4px' }} />;
      default:
        return null;
    }
  };

  // Обрезка текста предпросмотра
  const truncateText = (text: string, maxLength: number = 100): string => {
    return text.length > maxLength ? text.substring(0, maxLength) + '...' : text;
  };

  return (
    <div className="outlook-emails-container">
      {/* Временный код для отладки - удалите после настройки */}
      <Alert variant="info" className="mb-3">
        <strong>Текущий URL:</strong> {window.location.origin}<br/>
        <strong>Полный URL:</strong> {window.location.href}
      </Alert>
      
      <style>{`
        .outlook-emails-container {
          padding: 20px;
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          max-width: 100%;
          overflow-x: hidden;
        }
        
        .email-card {
          margin-bottom: 12px;
          border: 1px solid #e1e5e9;
          transition: all 0.2s ease;
          cursor: pointer;
        }
        
        .email-card:hover {
          border-color: #0078d4;
          box-shadow: 0 2px 8px rgba(0, 120, 212, 0.1);
        }
        
        .email-card.unread {
          border-left: 4px solid #0078d4;
          background-color: #f8f9fa;
        }
        
        .email-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 8px;
        }
        
        .email-subject {
          font-weight: 600;
          color: #323130;
          margin: 0;
          font-size: 14px;
        }
        
        .email-from {
          color: #605e5c;
          font-size: 13px;
          margin: 0;
        }
        
        .email-date {
          color: #8a8886;
          font-size: 12px;
        }
        
        .email-preview {
          color: #605e5c;
          font-size: 12px;
          margin-top: 8px;
          line-height: 1.3;
        }
        
        .email-badges {
          display: flex;
          gap: 4px;
          margin-top: 8px;
        }
        
        .refresh-button {
          margin-bottom: 16px;
        }
        
        .login-container {
          text-align: center;
          padding: 40px 20px;
        }
        
        .login-icon {
          font-size: 48px;
          color: #0078d4;
          margin-bottom: 16px;
        }
        
        .auth-info {
          background-color: #fff4ce;
          border: 1px solid #ffb900;
          border-radius: 4px;
          padding: 12px;
          margin-bottom: 16px;
          font-size: 13px;
        }
      `}</style>

      <div className="d-flex justify-content-between align-items-center mb-3">
        <h5 className="mb-0">
          <Icon iconName="Mail" style={{ marginRight: '8px', color: '#0078d4' }} />
          Последние письма
        </h5>
        {isAuthenticated && (
          <div className="d-flex gap-2">
            <Button 
              variant="outline-primary" 
              size="sm" 
              onClick={fetchEmails}
              disabled={loading}
              className="refresh-button"
            >
              {loading ? (
                <>
                  <Spinner animation="border" size="sm" className="me-2" />
                  Загрузка...
                </>
              ) : (
                <>
                  <Icon iconName="Refresh" style={{ marginRight: '4px' }} />
                  Обновить
                </>
              )}
            </Button>
            <Button 
              variant="outline-secondary" 
              size="sm" 
              onClick={handleLogout}
            >
              <Icon iconName="SignOut" style={{ marginRight: '4px' }} />
              Выйти
            </Button>
          </div>
        )}
      </div>

      {error && (
        <Alert variant="danger" className="mb-3">
          <Icon iconName="ErrorBadge" style={{ marginRight: '8px' }} />
          {error}
        </Alert>
      )}

      {!isAuthenticated ? (
        <div className="login-container">
          <div className="auth-info">
            <Icon iconName="Info" style={{ marginRight: '8px', color: '#ffb900' }} />
            После нажатия "Войти" вы будете перенаправлены на Microsoft Login, 
            а затем вернетесь на страницу taskpane.html для завершения авторизации.
          </div>
          <div className="login-icon">
            <Icon iconName="Signin" />
          </div>
          <h6 className="mb-3">Войдите в Microsoft 365</h6>
          <p className="text-muted mb-4">
            Для просмотра писем необходимо войти в ваш аккаунт Microsoft 365
          </p>
          <Button 
            variant="primary" 
            onClick={handleLogin}
            disabled={loading}
          >
            {loading ? (
              <>
                <Spinner animation="border" size="sm" className="me-2" />
                Выполняется вход...
              </>
            ) : (
              <>
                <Icon iconName="Signin" style={{ marginRight: '8px' }} />
                Войти
              </>
            )}
          </Button>
        </div>
      ) : (
        <>
          {loading && emails.length === 0 ? (
            <div className="text-center py-4">
              <Spinner animation="border" className="mb-3" />
              <p className="text-muted">Загрузка писем...</p>
            </div>
          ) : emails.length > 0 ? (
            <div>
              {emails.map((email) => (
                <Card key={email.id} className={`email-card ${!email.isRead ? 'unread' : ''}`}>
                  <Card.Body className="p-3">
                    <div className="email-header">
                      <div className="flex-grow-1">
                        <h6 className="email-subject">
                          {getImportanceIcon(email.importance)}
                          {email.subject || '(Без темы)'}
                          {!email.isRead && (
                            <Badge bg="primary" className="ms-2" style={{ fontSize: '10px' }}>
                              Новое
                            </Badge>
                          )}
                        </h6>
                        <p className="email-from">
                          От: {email.from?.emailAddress?.name || email.from?.emailAddress?.address || 'Неизвестный отправитель'}
                        </p>
                      </div>
                      <div className="email-date">
                        {formatDate(email.receivedDateTime)}
                      </div>
                    </div>
                    
                    {email.bodyPreview && (
                      <div className="email-preview">
                        {truncateText(email.bodyPreview)}
                      </div>
                    )}
                    
                    <div className="email-badges">
                      {email.hasAttachments && (
                        <Badge bg="secondary" style={{ fontSize: '10px' }}>
                          <Icon iconName="Attach" style={{ marginRight: '2px' }} />
                          Вложения
                        </Badge>
                      )}
                    </div>
                  </Card.Body>
                </Card>
              ))}
            </div>
          ) : (
            <div className="text-center py-4">
              <Icon iconName="Mail" style={{ fontSize: '48px', color: '#8a8886' }} />
              <p className="text-muted mt-3">Письма не найдены</p>
              <Button variant="outline-primary" onClick={fetchEmails}>
                Попробовать снова
              </Button>
            </div>
          )}
        </>
      )}
    </div>
  );
};

export default OutlookEmails;