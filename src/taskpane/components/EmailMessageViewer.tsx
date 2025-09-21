import React from 'react';
import { Message } from '@microsoft/microsoft-graph-types';
import { Stack, Text, Separator, Link } from '@fluentui/react'; // Assuming Fluent UI for Office-like styling

interface EmailMessageViewerProps {
  message: Message;
}

const EmailMessageViewer: React.FC<EmailMessageViewerProps> = ({ message }) => {
  const fromEmail = message.from?.emailAddress?.address || 'Unknown';
  const fromName = message.from?.emailAddress?.name || 'Unknown';
  const toRecipients = message.toRecipients?.map(recipient => recipient.emailAddress?.address).join(', ') || 'None';
  const subject = message.subject || 'No Subject';
  const sentDateTime = message.sentDateTime ? new Date(message.sentDateTime).toLocaleString() : 'Unknown';
  const bodyContent = message.body?.content || '';
  const bodyType = message.body?.contentType || 'text';

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: 10, maxWidth: 800 } }}>
      <Stack horizontal horizontalAlign="space-between">
        <Text variant="mediumPlus" block><strong>From:</strong> {fromName} ({fromEmail})</Text>
        <Text variant="medium" block>{sentDateTime}</Text>
      </Stack>
      <Text variant="medium"><strong>To:</strong> {toRecipients}</Text>
      <Separator />
      <Text variant="large" block><strong>Subject:</strong> {subject}</Text>
      <Separator />
      <Stack>
        <Text variant="medium"><strong>Body:</strong></Text>
        {bodyType === 'html' ? (
          <div dangerouslySetInnerHTML={{ __html: bodyContent }} />
        ) : (
          <Text block>{bodyContent}</Text>
        )}
      </Stack>
      {message.hasAttachments && (
        <Stack>
          <Separator />
          <Text variant="medium"><strong>Attachments:</strong> This message has attachments (implement fetching if needed).</Text>
        </Stack>
      )}
      {message.webLink && (
        <Link href={message.webLink} target="_blank">Open in Outlook</Link>
      )}
    </Stack>
  );
};

export default EmailMessageViewer;