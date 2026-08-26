import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { AadTokenProvider } from '@microsoft/sp-http';

interface IChatWithAiApplicationCustomizerProperties {
  imageUrl?: string;
  iframeUrl?: string;
  chatbotResource?: string;
  tokenRefreshMinutes?: number;
}

const AUTH_READY_MESSAGE = 'orkland-chatbot-auth-ready';
const AUTH_TOKEN_MESSAGE = 'orkland-chatbot-auth-token';

export default class ChatWithAiApplicationCustomizer extends BaseApplicationCustomizer<IChatWithAiApplicationCustomizerProperties> {
  private panelContainer: HTMLDivElement;
  private tokenRefreshTimer: number | undefined;

  public onInit(): Promise<void> {
    const imageUrl =
      this.properties.imageUrl ||
      'https://prokomresources.prokomcdn.no/client-grunt/bot/design_variations/kari-orkland.svg?';
    const iframeUrl = this.properties.iframeUrl || 'https://intranettchatbot.orkland.kommune.no';
    const chatbotResource = this.properties.chatbotResource || iframeUrl;
    const iframeOrigin = new URL(iframeUrl).origin;
    const tokenRefreshMinutes = this.properties.tokenRefreshMinutes || 45;

    const imageContainer = document.createElement('div');
    imageContainer.style.position = 'fixed';
    imageContainer.style.bottom = '20px';
    imageContainer.style.right = '20px';
    imageContainer.style.zIndex = '1000';
    imageContainer.style.cursor = 'pointer';

    const image = document.createElement('img');
    image.src = imageUrl;
    image.style.width = '60px';
    image.style.height = '60px';
    image.style.borderRadius = '50%';
    image.style.boxShadow = '0 4px 8px rgba(0,0,0,0.2)';
    image.alt = 'Bilde av Chat Bot som åpner Intranettchatbot';
    image.title = 'Åpne Intranett Chatbot';

    imageContainer.appendChild(image);
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

    const closeButton = document.createElement('button');
    closeButton.innerText = '×';
    closeButton.style.alignSelf = 'flex-end';
    closeButton.style.margin = '10px';
    closeButton.style.fontSize = '24px';
    closeButton.style.border = 'none';
    closeButton.style.background = 'none';
    closeButton.style.cursor = 'pointer';

    const statusText = document.createElement('div');
    statusText.style.padding = '0 12px 10px 12px';
    statusText.style.fontSize = '13px';
    statusText.style.color = '#605e5c';
    statusText.innerText = 'Henter pålogging...';

    const iframe = document.createElement('iframe');
    iframe.style.width = '100%';
    iframe.style.height = 'calc(100% - 74px)';
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

    iframe.onload = () => {
      sendToken();
    };

    window.addEventListener('message', handleMessage);

    this.panelContainer.appendChild(closeButton);
    this.panelContainer.appendChild(statusText);
    this.panelContainer.appendChild(iframe);
    document.body.appendChild(this.panelContainer);

    image.onclick = () => {
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
