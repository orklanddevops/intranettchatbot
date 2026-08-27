import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { AadTokenProvider } from '@microsoft/sp-http';

import defaultBotImageUrl from './assets/kommune_karlsen.svg';

interface IChatWithAiApplicationCustomizerProperties {
  imageUrl?: string;
  chatbotApiBaseUrl?: string;
  iframeUrl?: string;
  chatbotResource?: string;
  tokenRefreshMinutes?: number;
}

interface IChatMessage {
  id: string;
  role: string;
  content: string;
  date: string;
}

interface IChatResponseMessage {
  role: string;
  content?: string;
  context?: string;
}

interface IChatResponse {
  id?: string;
  choices?: Array<{
    messages?: IChatResponseMessage[];
  }>;
  error?: string | { message?: string };
}

const DEFAULT_CHATBOT_API_BASE_URL = 'https://intranettchatbot.orkland.kommune.no';
const DEFAULT_CHATBOT_RESOURCE = 'api://47cbcbfe-6efd-4113-b089-0dcb7c7b33bc';

export default class ChatWithAiApplicationCustomizer extends BaseApplicationCustomizer<IChatWithAiApplicationCustomizerProperties> {
  private panelContainer: HTMLDivElement;
  private tokenRefreshTimer: number | undefined;

  public onInit(): Promise<void> {
    const imageUrl = this.properties.imageUrl || defaultBotImageUrl;
    const chatbotApiBaseUrl = (
      this.properties.chatbotApiBaseUrl ||
      this.properties.iframeUrl ||
      DEFAULT_CHATBOT_API_BASE_URL
    ).replace(/\/$/, '');
    const configuredChatbotResource = this.properties.chatbotResource;
    const chatbotResource =
      configuredChatbotResource && configuredChatbotResource !== chatbotApiBaseUrl
        ? configuredChatbotResource
        : DEFAULT_CHATBOT_RESOURCE;
    const tokenRefreshMinutes = this.properties.tokenRefreshMinutes || 45;
    const messages: IChatMessage[] = [];

    const imageContainer = document.createElement('div');
    imageContainer.style.position = 'fixed';
    imageContainer.style.bottom = '20px';
    imageContainer.style.right = '20px';
    imageContainer.style.zIndex = '1000';
    imageContainer.style.cursor = 'pointer';
    imageContainer.style.width = '70px';
    imageContainer.style.height = '70px';

    const image = document.createElement('img');
    image.src = imageUrl;
    image.style.width = '60px';
    image.style.height = '60px';
    image.style.borderRadius = '50%';
    image.style.boxShadow = '0 4px 8px rgba(0,0,0,0.2)';
    image.alt = 'Bilde av Chat Bot som åpner Intranettchatbot';
    image.title = 'Åpne Intranett Chatbot';

    const launcherCloseButton = document.createElement('button');
    launcherCloseButton.innerText = '×';
    launcherCloseButton.type = 'button';
    launcherCloseButton.title = 'Skjul chatbot';
    launcherCloseButton.setAttribute('aria-label', 'Skjul chatbot');
    launcherCloseButton.style.position = 'absolute';
    launcherCloseButton.style.top = '0';
    launcherCloseButton.style.right = '0';
    launcherCloseButton.style.width = '22px';
    launcherCloseButton.style.height = '22px';
    launcherCloseButton.style.border = '1px solid rgba(0, 0, 0, 0.18)';
    launcherCloseButton.style.borderRadius = '50%';
    launcherCloseButton.style.background = '#ffffff';
    launcherCloseButton.style.boxShadow = '0 2px 6px rgba(0,0,0,0.2)';
    launcherCloseButton.style.color = '#323130';
    launcherCloseButton.style.cursor = 'pointer';
    launcherCloseButton.style.fontSize = '16px';
    launcherCloseButton.style.lineHeight = '18px';
    launcherCloseButton.style.padding = '0';

    imageContainer.appendChild(image);
    imageContainer.appendChild(launcherCloseButton);
    document.body.appendChild(imageContainer);

    this.panelContainer = document.createElement('div');
    this.panelContainer.style.position = 'fixed';
    this.panelContainer.style.top = '0';
    this.panelContainer.style.right = '0';
    this.panelContainer.style.width = '420px';
    this.panelContainer.style.height = '100%';
    this.panelContainer.style.backgroundColor = '#ffffff';
    this.panelContainer.style.boxShadow = '-2px 0 8px rgba(0,0,0,0.22)';
    this.panelContainer.style.zIndex = '1001';
    this.panelContainer.style.display = 'none';
    this.panelContainer.style.flexDirection = 'column';
    this.panelContainer.style.maxWidth = '100vw';
    this.panelContainer.style.fontFamily = '"Segoe UI", Arial, sans-serif';

    const panelHeader = document.createElement('div');
    panelHeader.style.display = 'flex';
    panelHeader.style.alignItems = 'center';
    panelHeader.style.justifyContent = 'space-between';
    panelHeader.style.minHeight = '52px';
    panelHeader.style.borderBottom = '1px solid #edebe9';
    panelHeader.style.padding = '0 8px 0 16px';

    const titleText = document.createElement('div');
    titleText.innerText = 'Intranett Chatbot';
    titleText.style.fontSize = '16px';
    titleText.style.fontWeight = '600';
    titleText.style.color = '#201f1e';

    const closeButton = document.createElement('button');
    closeButton.innerText = '×';
    closeButton.type = 'button';
    closeButton.title = 'Lukk chat';
    closeButton.setAttribute('aria-label', 'Lukk chat');
    closeButton.style.width = '40px';
    closeButton.style.height = '40px';
    closeButton.style.fontSize = '24px';
    closeButton.style.border = 'none';
    closeButton.style.background = 'none';
    closeButton.style.color = '#323130';
    closeButton.style.cursor = 'pointer';

    const statusText = document.createElement('div');
    statusText.style.minHeight = '18px';
    statusText.style.padding = '8px 16px';
    statusText.style.fontSize = '12px';
    statusText.style.color = '#605e5c';
    statusText.innerText = 'Klar';

    const messagesContainer = document.createElement('div');
    messagesContainer.style.flex = '1';
    messagesContainer.style.minHeight = '0';
    messagesContainer.style.overflowY = 'auto';
    messagesContainer.style.padding = '12px 16px';
    messagesContainer.style.background = '#faf9f8';

    const emptyState = document.createElement('div');
    emptyState.innerText = 'Hva kan jeg hjelpe deg med?';
    emptyState.style.color = '#605e5c';
    emptyState.style.fontSize = '14px';
    emptyState.style.marginTop = '12px';
    messagesContainer.appendChild(emptyState);

    const composer = document.createElement('form');
    composer.style.display = 'flex';
    composer.style.gap = '8px';
    composer.style.padding = '12px';
    composer.style.borderTop = '1px solid #edebe9';
    composer.style.background = '#ffffff';

    const input = document.createElement('textarea');
    input.placeholder = 'Skriv en melding';
    input.rows = 2;
    input.style.flex = '1';
    input.style.resize = 'none';
    input.style.border = '1px solid #c8c6c4';
    input.style.borderRadius = '4px';
    input.style.fontFamily = 'inherit';
    input.style.fontSize = '14px';
    input.style.padding = '8px';
    input.style.minHeight = '40px';
    input.style.maxHeight = '120px';

    const sendButton = document.createElement('button');
    sendButton.type = 'submit';
    sendButton.innerText = 'Send';
    sendButton.style.border = 'none';
    sendButton.style.borderRadius = '4px';
    sendButton.style.background = '#0078d4';
    sendButton.style.color = '#ffffff';
    sendButton.style.cursor = 'pointer';
    sendButton.style.fontWeight = '600';
    sendButton.style.padding = '0 14px';

    let isSending = false;

    const createMessageId = (): string => {
      return `${new Date().getTime()}-${Math.random().toString(16).slice(2)}`;
    };

    const appendMessage = (role: string, content: string): HTMLDivElement => {
      if (emptyState.parentElement) {
        emptyState.remove();
      }

      const wrapper = document.createElement('div');
      wrapper.style.display = 'flex';
      wrapper.style.justifyContent = role === 'user' ? 'flex-end' : 'flex-start';
      wrapper.style.margin = '8px 0';

      const bubble = document.createElement('div');
      bubble.innerText = content;
      bubble.style.maxWidth = '86%';
      bubble.style.padding = '9px 11px';
      bubble.style.borderRadius = '6px';
      bubble.style.whiteSpace = 'pre-wrap';
      bubble.style.overflowWrap = 'break-word';
      bubble.style.fontSize = '14px';
      bubble.style.lineHeight = '1.4';
      bubble.style.background = role === 'user' ? '#0078d4' : '#ffffff';
      bubble.style.color = role === 'user' ? '#ffffff' : '#201f1e';
      bubble.style.border = role === 'user' ? 'none' : '1px solid #edebe9';

      wrapper.appendChild(bubble);
      messagesContainer.appendChild(wrapper);
      messagesContainer.scrollTop = messagesContainer.scrollHeight;
      return bubble;
    };

    const setError = (message: string): void => {
      statusText.innerText = message;
      statusText.style.color = '#a4262c';
    };

    const setStatus = (message: string): void => {
      statusText.innerText = message;
      statusText.style.color = '#605e5c';
    };

    const setSendingState = (sending: boolean): void => {
      isSending = sending;
      input.disabled = sending;
      sendButton.disabled = sending;
      sendButton.style.opacity = sending ? '0.65' : '1';
    };

    const getAccessToken = async (): Promise<string> => {
      setStatus('Henter pålogging...');
      const tokenProvider: AadTokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      const accessToken = await tokenProvider.getToken(chatbotResource);
      setStatus('Klar');
      return accessToken;
    };

    const processResponseMessage = (
      responseMessage: IChatResponseMessage,
      assistantMessage: IChatMessage,
      assistantBubble: HTMLDivElement
    ): void => {
      if (responseMessage.role === 'tool') {
        messages.push({
          id: createMessageId(),
          role: 'tool',
          content: responseMessage.content || responseMessage.context || '',
          date: new Date().toISOString()
        });
        return;
      }

      if (responseMessage.role !== 'assistant') {
        return;
      }

      assistantMessage.content += responseMessage.content || '';
      assistantBubble.innerText = assistantMessage.content || '...';
      messagesContainer.scrollTop = messagesContainer.scrollHeight;
    };

    const processChatResponse = (
      payload: IChatResponse,
      assistantMessage: IChatMessage,
      assistantBubble: HTMLDivElement
    ): void => {
      if (payload.error) {
        if (typeof payload.error === 'string') {
          throw new Error(payload.error);
        }
        throw new Error(payload.error.message || 'Ukjent feil fra chatboten.');
      }

      if (!payload.choices || payload.choices.length === 0) {
        return;
      }

      const responseMessages = payload.choices[0].messages || [];
      responseMessages.forEach((responseMessage: IChatResponseMessage) => {
        processResponseMessage(responseMessage, assistantMessage, assistantBubble);
      });
    };

    const sendMessage = async (): Promise<void> => {
      const question = input.value.trim();
      if (!question || isSending) {
        return;
      }

      setSendingState(true);
      input.value = '';
      setStatus('Sender...');

      const userMessage: IChatMessage = {
        id: createMessageId(),
        role: 'user',
        content: question,
        date: new Date().toISOString()
      };
      messages.push(userMessage);
      appendMessage('user', question);

      const assistantMessage: IChatMessage = {
        id: createMessageId(),
        role: 'assistant',
        content: '',
        date: new Date().toISOString()
      };
      const assistantBubble = appendMessage('assistant', '...');

      try {
        const token = await getAccessToken();
        const response = await fetch(`${chatbotApiBaseUrl}/conversation`, {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            messages: messages.filter((message: IChatMessage) => message.role !== 'error')
          })
        });

        if (!response.ok) {
          throw new Error(`Chatboten svarte med ${response.status}.`);
        }

        if (response.body) {
          const reader = response.body.getReader();
          const decoder = new TextDecoder('utf-8');
          let runningText = '';
          let readResult = await reader.read();

          while (!readResult.done) {
            const chunk = decoder.decode(readResult.value);
            const objects = chunk.split('\n');
            objects.forEach((obj: string) => {
              if (obj === '' || obj === '{}') {
                return;
              }

              try {
                runningText += obj;
                processChatResponse(JSON.parse(runningText) as IChatResponse, assistantMessage, assistantBubble);
                runningText = '';
              } catch (error) {
                if (!(error instanceof SyntaxError)) {
                  throw error;
                }
              }
            });

            readResult = await reader.read();
          }
        } else {
          processChatResponse(await response.json() as IChatResponse, assistantMessage, assistantBubble);
        }

        if (!assistantMessage.content) {
          assistantMessage.content = 'Jeg fikk ikke noe svar fra chatboten.';
          assistantBubble.innerText = assistantMessage.content;
        }

        messages.push(assistantMessage);
        setStatus('Klar');
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : 'Ukjent feil.';
        assistantBubble.innerText = `En feil oppstod. ${errorMessage}`;
        messages.push({
          id: createMessageId(),
          role: 'error',
          content: assistantBubble.innerText,
          date: new Date().toISOString()
        });
        setError('Kunne ikke kontakte chatboten.');
        console.error('Could not send chatbot message', error);
      } finally {
        setSendingState(false);
        input.focus();
      }
    };

    closeButton.onclick = () => {
      this.panelContainer.style.display = 'none';
    };

    launcherCloseButton.onclick = (event: MouseEvent) => {
      event.stopPropagation();
      this.panelContainer.style.display = 'none';
      imageContainer.style.display = 'none';
    };

    imageContainer.onclick = () => {
      this.panelContainer.style.display = 'flex';
      input.focus();
      getAccessToken().catch((error: unknown) => {
        setError('Kunne ikke hente pålogging til chatboten.');
        console.error('Could not acquire chatbot token', error);
      });
    };

    composer.onsubmit = (event: Event) => {
      event.preventDefault();
      sendMessage().catch((error: unknown) => {
        setError('Kunne ikke sende melding.');
        console.error('Could not submit chatbot message', error);
      });
    };

    input.onkeydown = (event: KeyboardEvent) => {
      if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        sendMessage().catch((error: unknown) => {
          setError('Kunne ikke sende melding.');
          console.error('Could not submit chatbot message', error);
        });
      }
    };

    this.tokenRefreshTimer = window.setInterval(() => {
      if (this.panelContainer.style.display !== 'none') {
        getAccessToken().catch((error: unknown) => {
          setError('Kunne ikke fornye pålogging til chatboten.');
          console.error('Could not refresh chatbot token', error);
        });
      }
    }, tokenRefreshMinutes * 60 * 1000);

    panelHeader.appendChild(titleText);
    panelHeader.appendChild(closeButton);
    composer.appendChild(input);
    composer.appendChild(sendButton);
    this.panelContainer.appendChild(panelHeader);
    this.panelContainer.appendChild(statusText);
    this.panelContainer.appendChild(messagesContainer);
    this.panelContainer.appendChild(composer);
    document.body.appendChild(this.panelContainer);

    return Promise.resolve();
  }

  public onDispose(): void {
    if (this.tokenRefreshTimer !== undefined) {
      window.clearInterval(this.tokenRefreshTimer);
    }
  }
}
