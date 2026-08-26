import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { AadTokenProvider } from '@microsoft/sp-http';

import defaultBotImageUrl from './assets/kommune_karlsen.svg';

interface IChatWithAiApplicationCustomizerProperties {
  imageUrl?: string;
  iframeUrl?: string;
  chatbotResource?: string;
  tokenRefreshMinutes?: number;
}

const AUTH_READY_MESSAGE = 'orkland-chatbot-auth-ready';
const AUTH_TOKEN_MESSAGE = 'orkland-chatbot-auth-token';
const DEFAULT_IFRAME_URL = 'https://intranettchatbot.orkland.kommune.no';
const DEFAULT_CHATBOT_RESOURCE = 'api://47cbcbfe-6efd-4113-b089-0dcb7c7b33bc';

export default class ChatWithAiApplicationCustomizer extends BaseApplicationCustomizer<IChatWithAiApplicationCustomizerProperties> {
  private panelContainer: HTMLDivElement;
  private tokenRefreshTimer: number | undefined;

  public onInit(): Promise<void> {
    const imageUrl = this.properties.imageUrl || defaultBotImageUrl;
    const iframeUrl = this.properties.iframeUrl || DEFAULT_IFRAME_URL;
    const configuredChatbotResource = this.properties.chatbotResource;
    const chatbotResource =
      configuredChatbotResource && configuredChatbotResource !== iframeUrl
        ? configuredChatbotResource
        : DEFAULT_CHATBOT_RESOURCE;
    const iframeOrigin = new URL(iframeUrl).origin;
    const tokenRefreshMinutes = this.properties.tokenRefreshMinutes || 45;

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
    this.panelContainer.style.width = '400px';
    this.panelContainer.style.height = '100%';
    this.panelContainer.style.backgroundColor = '#fff';
    this.panelContainer.style.boxShadow = '-2px 0 8px rgba(0,0,0,0.2)';
    this.panelContainer.style.zIndex = '1001';
    this.panelContainer.style.display = 'none';
    this.panelContainer.style.flexDirection = 'column';
    this.panelContainer.style.maxWidth = '100vw';

    const panelHeader = document.createElement('div');
    panelHeader.style.display = 'flex';
    panelHeader.style.alignItems = 'center';
    panelHeader.style.justifyContent = 'flex-end';
    panelHeader.style.minHeight = '48px';
    panelHeader.style.borderBottom = '1px solid #edebe9';

    const closeButton = document.createElement('button');
    closeButton.innerText = '×';
    closeButton.type = 'button';
    closeButton.title = 'Lukk chat';
    closeButton.setAttribute('aria-label', 'Lukk chat');
    closeButton.style.width = '40px';
    closeButton.style.height = '40px';
    closeButton.style.margin = '4px 8px';
    closeButton.style.fontSize = '24px';
    closeButton.style.border = 'none';
    closeButton.style.background = 'none';
    closeButton.style.color = '#323130';
    closeButton.style.cursor = 'pointer';

    const statusText = document.createElement('div');
    statusText.style.padding = '0 12px 10px 12px';
    statusText.style.fontSize = '13px';
    statusText.style.color = '#605e5c';
    statusText.innerText = 'Henter pålogging...';

    const iframe = document.createElement('iframe');
    iframe.style.width = '100%';
    iframe.style.height = 'auto';
    iframe.style.flex = '1';
    iframe.style.minHeight = '0';
    iframe.style.border = 'none';
    iframe.title = 'Intranett Chatbot';

    const sendTokenToIframe = async (): Promise<void> => {
      if (!iframe.contentWindow) {
        return;
      }

      try {
        const tokenProvider: AadTokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
        const accessToken = await tokenProvider.getToken(chatbotResource);
        iframe.contentWindow.postMessage(
          {
            type: AUTH_TOKEN_MESSAGE,
            accessToken
          },
          iframeOrigin
        );
        statusText.innerText = '';
      } catch (error) {
        statusText.innerText = 'Kunne ikke hente pålogging til chatboten. Kontakt administrator.';
        console.error('Could not acquire chatbot token', error);
      }
    };

    const sendToken = (): void => {
      sendTokenToIframe().catch((error) => {
        statusText.innerText = 'Kunne ikke hente pålogging til chatboten. Kontakt administrator.';
        console.error('Could not send chatbot token', error);
      });
    };

    const handleMessage = (event: MessageEvent): void => {
      if (event.origin !== iframeOrigin || event.data?.type !== AUTH_READY_MESSAGE) {
        return;
      }

      sendToken();
    };

    closeButton.onclick = () => {
      this.panelContainer.style.display = 'none';
    };

    launcherCloseButton.onclick = (event: MouseEvent) => {
      event.stopPropagation();
      this.panelContainer.style.display = 'none';
      imageContainer.style.display = 'none';
    };

    iframe.onload = () => {
      sendToken();
    };

    window.addEventListener('message', handleMessage);

    panelHeader.appendChild(closeButton);
    this.panelContainer.appendChild(panelHeader);
    this.panelContainer.appendChild(statusText);
    this.panelContainer.appendChild(iframe);
    document.body.appendChild(this.panelContainer);

    imageContainer.onclick = () => {
      this.panelContainer.style.display = 'flex';
      if (!iframe.src) {
        iframe.src = iframeUrl;
      } else {
        sendToken();
      }
    };

    this.tokenRefreshTimer = window.setInterval(() => {
      if (this.panelContainer.style.display !== 'none') {
        sendToken();
      }
    }, tokenRefreshMinutes * 60 * 1000);

    return Promise.resolve();
  }

  public onDispose(): void {
    if (this.tokenRefreshTimer !== undefined) {
      window.clearInterval(this.tokenRefreshTimer);
    }
  }
}
