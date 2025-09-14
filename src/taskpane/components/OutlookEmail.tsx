import React, { useState, useEffect } from 'react';
import { Card, Button, Alert, Spinner, Badge } from 'react-bootstrap';
import { Icon } from '@fluentui/react';
// import * as Office from '@microsoft/office-js';

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
  // Temporarily keep clientId for compatibility with App.tsx; remove later
  clientId?: string;
}

const OutlookEmails: React.FC<OutlookEmailsProps> = () => {
  const [emails, setEmails] = useState<EmailData[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string>('');

  // Автоматическая загрузка при монтировании
  useEffect(() => {
    fetchEmails();
  }, []);

  // Функция для получения токена и писем
  const fetchEmails = async () => {
    setLoading(true);
    setError('');

    try {
      // Получаем access token через Office SSO
      const bootstrapToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      });

      // Запрос к Graph API
      const response = await fetch('https://graph.microsoft.com/v1.0/me/messages?$top=5&$orderby=receivedDateTime desc', {
        headers: {
          Authorization: `Bearer ${bootstrapToken}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${await response.text()}`);
      }

      const data = await response.json();
      setEmails(data.value || []);
    } catch (err: any) {
      console.error('SSO error:', err);
      setError(`Ошибка: ${err.message || 'Не удалось получить доступ к Outlook. Проверьте настройки в Azure.'}`);
    } finally {
      setLoading(false);
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
      year: 'numeric',
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
      `}</style>

      <div className="d-flex justify-content-between align-items-center mb-3">
        <h5 className="mb-0">
          <Icon iconName="Mail" style={{ marginRight: '8px', color: '#0078d4' }} />
          Последние письма
        </h5>
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
      </div>

      {error && (
        <Alert variant="danger" className="mb-3">
          <Icon iconName="ErrorBadge" style={{ marginRight: '8px' }} />
          {error}
        </Alert>
      )}

      {loading ? (
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
                  <div className="email-date">{formatDate(email.receivedDateTime)}</div>
                </div>

                {email.bodyPreview && <div className="email-preview">{truncateText(email.bodyPreview)}</div>}

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
    </div>
  );
};

export default OutlookEmails;