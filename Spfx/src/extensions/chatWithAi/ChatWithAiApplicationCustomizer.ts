import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { AadTokenProvider } from '@microsoft/sp-http';

import defaultBotImageUrl from './assets/kommune_karlsen.svg';
import sendIconUrl from './assets/Send.svg';

interface IChatWithAiApplicationCustomizerProperties {
  imageUrl?: string;
  chatbotApiBaseUrl?: string;
  iframeUrl?: string;
  chatbotResource?: string;
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
  private panelContainer: HTMLDivElement | undefined;
  private launcherContainer: HTMLDivElement | undefined;

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
    const messages: IChatMessage[] = [];

    const launcherContainer = document.createElement('div');
    launcherContainer.style.position = 'fixed';
    launcherContainer.style.bottom = '20px';
    launcherContainer.style.right = '20px';
    launcherContainer.style.zIndex = '1000';
    launcherContainer.style.cursor = 'pointer';
    launcherContainer.style.width = '70px';
    launcherContainer.style.height = '70px';

    const launcherImage = document.createElement('img');
    launcherImage.src = imageUrl;
    launcherImage.style.width = '60px';
    launcherImage.style.height = '60px';
    launcherImage.style.borderRadius = '50%';
    launcherImage.style.boxShadow = '0 4px 8px rgba(0,0,0,0.2)';
    launcherImage.alt = 'Bilde av Chat Bot som åpner Intranettchatbot';
    launcherImage.title = 'Åpne Intranett Chatbot';

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

    launcherContainer.appendChild(launcherImage);
    launcherContainer.appendChild(launcherCloseButton);
    document.body.appendChild(launcherContainer);
    this.launcherContainer = launcherContainer;

    const panelContainer = document.createElement('div');
    panelContainer.style.position = 'fixed';
    panelContainer.style.top = '0';
    panelContainer.style.right = '0';
    panelContainer.style.width = 'min(520px, 48vw)';
    panelContainer.style.height = '100%';
    panelContainer.style.maxWidth = '100vw';
    panelContainer.style.background = 'radial-gradient(108.78% 108.78% at 50.02% 19.78%, #ffffff 57.29%, #eef6fe 100%)';
    panelContainer.style.boxShadow = '-2px 0 8px rgba(0,0,0,0.22)';
    panelContainer.style.zIndex = '1001';
    panelContainer.style.display = 'none';
    panelContainer.style.flexDirection = 'column';
    panelContainer.style.fontFamily = '"Segoe UI", Arial, sans-serif';
    panelContainer.style.overflow = 'hidden';

    const closeButton = document.createElement('button');
    closeButton.innerText = '×';
    closeButton.type = 'button';
    closeButton.title = 'Lukk chat';
    closeButton.setAttribute('aria-label', 'Lukk chat');
    closeButton.style.position = 'absolute';
    closeButton.style.top = '8px';
    closeButton.style.right = '10px';
    closeButton.style.zIndex = '2';
    closeButton.style.width = '36px';
    closeButton.style.height = '36px';
    closeButton.style.fontSize = '24px';
    closeButton.style.lineHeight = '28px';
    closeButton.style.border = 'none';
    closeButton.style.borderRadius = '4px';
    closeButton.style.background = 'rgba(255,255,255,0.85)';
    closeButton.style.color = '#323130';
    closeButton.style.cursor = 'pointer';
    closeButton.style.padding = '0';

    const messagesContainer = document.createElement('div');
    messagesContainer.style.flex = '1';
    messagesContainer.style.minHeight = '0';
    messagesContainer.style.overflowY = 'auto';
    messagesContainer.style.overflowX = 'hidden';
    messagesContainer.style.padding = '56px 24px 0 24px';
    messagesContainer.style.display = 'flex';
    messagesContainer.style.flexDirection = 'column';

    const emptyState = document.createElement('div');
    emptyState.style.flex = '1';
    emptyState.style.display = 'flex';
    emptyState.style.justifyContent = 'center';
    emptyState.style.alignItems = 'center';

    const emptyIcon = document.createElement('img');
    emptyIcon.src = imageUrl;
    emptyIcon.alt = '';
    emptyIcon.setAttribute('aria-hidden', 'true');
    emptyIcon.style.height = '150px';
    emptyIcon.style.width = 'auto';
    emptyState.appendChild(emptyIcon);
    messagesContainer.appendChild(emptyState);

    const composerRegion = document.createElement('div');
    composerRegion.style.position = 'relative';
    composerRegion.style.flex = '0 0 auto';
    composerRegion.style.minHeight = '112px';
    composerRegion.style.padding = '12px 24px 24px 12px';

    const composer = document.createElement('form');
    composer.style.position = 'absolute';
    composer.style.left = '5%';
    composer.style.right = '2.5%';
    composer.style.bottom = '20px';
    composer.style.minHeight = '60px';
    composer.style.maxHeight = '200px';
    composer.style.background = '#ffffff';
    composer.style.boxShadow = '0px 8px 16px rgba(0, 0, 0, 0.14), 0px 0px 2px rgba(0, 0, 0, 0.12)';
    composer.style.borderRadius = '0px 8px 8px 0px';
    composer.style.overflow = 'hidden';
    composer.style.display = 'flex';
    composer.style.alignItems = 'center';
    composer.style.padding = '0 42px 4px 8px';

    const input = document.createElement('textarea');
    input.placeholder = 'Skriv inn et nytt spørsmål...';
    input.rows = 1;
    input.style.width = '100%';
    input.style.minHeight = '28px';
    input.style.maxHeight = '150px';
    input.style.border = 'none';
    input.style.outline = 'none';
    input.style.background = 'transparent';
    input.style.resize = 'none';
    input.style.overflowY = 'hidden';
    input.style.fontFamily = 'inherit';
    input.style.fontSize = '14px';
    input.style.lineHeight = '24px';
    input.style.color = '#242424';
    input.style.margin = '5px';

    const sendButton = document.createElement('button');
    sendButton.type = 'submit';
    sendButton.title = 'Send';
    sendButton.setAttribute('aria-label', 'Send melding');
    sendButton.style.position = 'absolute';
    sendButton.style.bottom = '15px';
    sendButton.style.right = '5px';
    sendButton.style.width = '32px';
    sendButton.style.height = '32px';
    sendButton.style.border = 'none';
    sendButton.style.background = 'transparent';
    sendButton.style.cursor = 'pointer';
    sendButton.style.padding = '4px';

    const sendIcon = document.createElement('img');
    sendIcon.src = sendIconUrl;
    sendIcon.alt = '';
    sendIcon.setAttribute('aria-hidden', 'true');
    sendIcon.style.width = '24px';
    sendIcon.style.height = '23px';
    sendButton.appendChild(sendIcon);

    const inputBottomBorder = document.createElement('div');
    inputBottomBorder.style.position = 'absolute';
    inputBottomBorder.style.width = '100%';
    inputBottomBorder.style.height = '4px';
    inputBottomBorder.style.left = '0';
    inputBottomBorder.style.bottom = '0';
    inputBottomBorder.style.background = 'radial-gradient(106.04% 106.06% at 100.1% 90.19%, #18864a 33.63%, #8dddd8 100%)';
    inputBottomBorder.style.borderBottomRightRadius = '8px';

    composer.appendChild(input);
    composer.appendChild(sendButton);
    composer.appendChild(inputBottomBorder);
    composerRegion.appendChild(composer);

    panelContainer.appendChild(closeButton);
    panelContainer.appendChild(messagesContainer);
    panelContainer.appendChild(composerRegion);
    document.body.appendChild(panelContainer);
    this.panelContainer = panelContainer;

    let isSending = false;

    const createMessageId = (): string => {
      return `${new Date().getTime()}-${Math.random().toString(16).slice(2)}`;
    };

    const scrollToBottom = (): void => {
      messagesContainer.scrollTop = messagesContainer.scrollHeight;
    };

    const removeEmptyState = (): void => {
      if (emptyState.parentElement) {
        emptyState.remove();
      }
    };

    const appendUserMessage = (content: string): void => {
      removeEmptyState();

      const wrapper = document.createElement('div');
      wrapper.style.display = 'flex';
      wrapper.style.justifyContent = 'flex-end';
      wrapper.style.marginBottom = '12px';

      const bubble = document.createElement('div');
      bubble.innerText = content;
      bubble.style.position = 'relative';
      bubble.style.display = 'flex';
      bubble.style.padding = '20px';
      bubble.style.background = '#e9f5ee';
      bubble.style.borderRadius = '8px';
      bubble.style.boxShadow = '0px 2px 4px rgba(0, 0, 0, 0.14), 0px 0px 2px rgba(0, 0, 0, 0.12)';
      bubble.style.fontWeight = '400';
      bubble.style.fontSize = '14px';
      bubble.style.lineHeight = '22px';
      bubble.style.color = '#242424';
      bubble.style.whiteSpace = 'pre-wrap';
      bubble.style.wordWrap = 'break-word';
      bubble.style.maxWidth = '80%';

      wrapper.appendChild(bubble);
      messagesContainer.appendChild(wrapper);
      scrollToBottom();
    };

    const appendAssistantMessage = (content: string): HTMLDivElement => {
      removeEmptyState();

      const wrapper = document.createElement('div');
      wrapper.style.marginBottom = '12px';
      wrapper.style.maxWidth = '80%';
      wrapper.style.display = 'flex';

      const answerContainer = document.createElement('div');
      answerContainer.style.display = 'flex';
      answerContainer.style.flexDirection = 'column';
      answerContainer.style.alignItems = 'flex-start';
      answerContainer.style.padding = '8px';
      answerContainer.style.gap = '5px';
      answerContainer.style.background = '#ffffff';
      answerContainer.style.boxShadow = '0px 1px 2px rgba(0, 0, 0, 0.14), 0px 0px 2px rgba(0, 0, 0, 0.12)';
      answerContainer.style.borderRadius = '5px';

      const answerText = document.createElement('div');
      answerText.innerText = content;
      answerText.style.fontWeight = '400';
      answerText.style.fontSize = '14px';
      answerText.style.lineHeight = '20px';
      answerText.style.color = '#323130';
      answerText.style.margin = '11px';
      answerText.style.whiteSpace = 'pre-wrap';
      answerText.style.wordWrap = 'break-word';
      answerText.style.overflowX = 'auto';

      const footer = document.createElement('div');
      footer.style.display = 'flex';
      footer.style.width = '100%';
      footer.style.justifyContent = 'flex-end';

      const disclaimer = document.createElement('span');
      disclaimer.innerText = 'Innhold generert av AI kan være feil';
      disclaimer.style.fontWeight = '400';
      disclaimer.style.fontSize = '12px';
      disclaimer.style.lineHeight = '16px';
      disclaimer.style.color = '#707070';

      footer.appendChild(disclaimer);
      answerContainer.appendChild(answerText);
      answerContainer.appendChild(footer);
      wrapper.appendChild(answerContainer);
      messagesContainer.appendChild(wrapper);
      scrollToBottom();
      return answerText;
    };

    const setSendingState = (sending: boolean): void => {
      isSending = sending;
      input.disabled = sending;
      sendButton.disabled = sending;
      sendButton.style.opacity = sending ? '0.65' : '1';
      sendButton.style.cursor = sending ? 'default' : 'pointer';
    };

    const resizeInput = (): void => {
      input.style.height = 'auto';
      const nextHeight = Math.min(Math.max(input.scrollHeight, 28), 150);
      input.style.height = `${nextHeight}px`;
      input.style.overflowY = nextHeight >= 150 ? 'auto' : 'hidden';
    };

    const getAccessToken = async (): Promise<string> => {
      const tokenProvider: AadTokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      return tokenProvider.getToken(chatbotResource);
    };

    const formatErrorMessage = (error: unknown): string => {
      if (error instanceof Error && error.message) {
        return `En feil oppstod. ${error.message}`;
      }

      return 'En feil oppstod. Vennligst prøv igjen. Hvis problemet vedvarer, kontakt nettstedets administrator.';
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
      assistantBubble.innerText = assistantMessage.content || 'Genererer svar...';
      scrollToBottom();
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
      resizeInput();

      const userMessage: IChatMessage = {
        id: createMessageId(),
        role: 'user',
        content: question,
        date: new Date().toISOString()
      };
      messages.push(userMessage);
      appendUserMessage(question);

      const assistantMessage: IChatMessage = {
        id: createMessageId(),
        role: 'assistant',
        content: '',
        date: new Date().toISOString()
      };
      const assistantBubble = appendAssistantMessage('Genererer svar...');

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
            const chunk = readResult.value ? decoder.decode(readResult.value) : '';
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
      } catch (error) {
        assistantMessage.content = formatErrorMessage(error);
        assistantBubble.innerText = assistantMessage.content;
        console.error('Could not send chatbot message', error);
      } finally {
        messages.push(assistantMessage);
        setSendingState(false);
        input.focus();
      }
    };

    closeButton.onclick = () => {
      panelContainer.style.display = 'none';
    };

    launcherCloseButton.onclick = (event: MouseEvent) => {
      event.stopPropagation();
      panelContainer.style.display = 'none';
      launcherContainer.style.display = 'none';
    };

    launcherContainer.onclick = () => {
      panelContainer.style.display = 'flex';
      input.focus();
    };

    composer.onsubmit = (event: Event) => {
      event.preventDefault();
      sendMessage().catch((error: unknown) => {
        console.error('Could not submit chatbot message', error);
      });
    };

    input.oninput = () => {
      resizeInput();
    };

    input.onkeydown = (event: KeyboardEvent) => {
      if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        sendMessage().catch((error: unknown) => {
          console.error('Could not submit chatbot message', error);
        });
      }
    };

    return Promise.resolve();
  }

  public onDispose(): void {
    if (this.panelContainer && this.panelContainer.parentElement) {
      this.panelContainer.remove();
    }

    if (this.launcherContainer && this.launcherContainer.parentElement) {
      this.launcherContainer.remove();
    }
  }
}
