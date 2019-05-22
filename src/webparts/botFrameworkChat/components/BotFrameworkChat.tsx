import * as React from 'react';
import { css } from 'office-ui-fabric-react';
import { TextField } from 'office-ui-fabric-react';
import styles from './BotFrameworkChat.module.scss';
import { IBotFrameworkChatProps } from './IBotFrameworkChatProps';
declare function require(path: string) : any;

export default class BotFrameworkChat extends React.Component<IBotFrameworkChatProps, {}> {

  private pollInterval = 1000;
  private directLineClient;
  private conversationId;
  private directLineClientSwagger;
  private messagesHtml;
  private currentMessageText;
  private sendAsUserName;

  // webpart lifecycle 和 event listner 结合
  // 为 输入框绑定 keyDown 和 keyUp event
  public render(): JSX.Element {
    return (
      <div className={styles.botFrameworkChat}>
        <div className={styles.container}>
          <div className={css('ms-Grid-row ms-font-xl', styles.chatHeader)} style={{ backgroundColor: '#' + this.props.titleBarBackgroundColor }} >
            {this.props.title}
          </div>
          <div className={css('ms-Grid-row', styles.messagesRow)} >
            <div className='ms-Grid-col ms-u-sm12' ref='messageHistoryDiv' dangerouslySetInnerHTML={{ __html: this.getMessagesHtml() }}>
            </div>
          </div>
          <div className={css('ms-Grid-row')}>
            <TextField
              id='MessageBox'
              onKeyUp={(e) => this.tbKeyUp(e)}
              onKeyDown={(e) => this.tbKeyDown(e)}
              value={this.currentMessageText}
              placeholder={this.props.placeholderText}
              className={css('ms-fontSize-m', styles.messageBox)}
            />
          </div>
        </div>
      </div>
    );
  }

  // 调用 Graph API 把 converastion 最新的消息 pull下来
  public componentDidUpdate(prevProps: IBotFrameworkChatProps, prevState: {}, prevContext: any): void {
    if (this.props.directLineSecret !== prevProps.directLineSecret) {
      if (this.props.directLineSecret) {
        var Swagger = require('swagger-client');
        var directLineSpec = require('./directline-swagger.json');

        this.directLineClientSwagger = new Swagger(
          {
            spec: directLineSpec,
            usePromise: true,
          }).then((client) => {
            client.clientAuthorizations.add('AuthorizationBotConnector', new Swagger.ApiKeyAuthorization('Authorization', 'BotConnector ' + this.props.directLineSecret, 'header'));
            console.log('DirectLine client generated');
            return client;
          }).catch((err) =>
            console.error('Error initializing DirectLine client', err));

        this.directLineClientSwagger.then((client) => {
          client.Conversations.Conversations_NewConversation()
            .then((response) => response.obj.conversationId)
            .then((conversationId) => {

              this.conversationId = conversationId;
              this.pollMessages(client, conversationId);
              this.directLineClient = client;
            });
        });

        this.sendAsUserName = this.props.context.pageContext.user.loginName;

        this.printMessage = this.printMessage.bind(this);
      }
    }
  }

  // 用户按下 Enter 键之后，调用 Graph API 把用户输入的消息发送出去
  public tbKeyDown(e) {
    if (e.keyCode === 13) {
      var messageToSend = this.currentMessageText;

      this.currentMessageText = '';

      this.setState({
        message: '',
      });

      if (!this.messagesHtml) {
        this.messagesHtml = '';
      }

      this.messagesHtml = this.messagesHtml + ' <span class="' + styles.message
        + ' ' + styles.fromUser + '  ms-fontSize-mPlus" style="background-color:#' + this.props.userMessagesBackgroundColor
        + '; color:#' + this.props.userMessagesForegroundColor + '">' + e.target.value + '</span> ';

      this.directLineClient.Conversations.Conversations_PostMessage(
        {
          conversationId: this.conversationId,
          message: {
            from: this.sendAsUserName,
            text: messageToSend
          }
        }).catch((err) => console.error('Error sending message:', err));
    }
  }
  
  // 用户释放 Enter 键之后，把整个消息对话框往上拉，形成消息流
  public tbKeyUp(e) {
    this.currentMessageText = e.target.value;
    this.forceMessagesContainerScroll();
  }

  protected pollMessages(client, conversationId) {
    console.log('Starting polling message for conversationId: ' + conversationId);
    var watermark = null;
    setInterval(() => {
      client.Conversations.Conversations_GetMessages({ conversationId: conversationId, watermark: watermark })
        .then((response) => {
          watermark = response.obj.watermark;
          return response.obj.messages;
        })
        .then((messages) => this.printMessages(messages));
    }, this.pollInterval);
  }

  protected printMessages(messages) {
    if (messages && messages.length) {
      messages = messages.filter((m) => m.from !== this.sendAsUserName);
      if (messages.length) {
        messages.forEach(this.printMessage);
      }
    }
  }

  protected getMessagesHtml() {
    return this.messagesHtml;
  }

  protected printMessage(message) {
    if (message.text) {
      this.setState({
        message: this.currentMessageText,
      });

      if (!this.messagesHtml) {
        this.messagesHtml = '';
      }

      this.messagesHtml = this.messagesHtml + ' <span class="' + styles.message + ' '
        + styles.fromBot + ' ms-fontSize-m" style="background-color:#' + this.props.botMessagesBackgroundColor
        + '; color:#' + this.props.botMessagesForegroundColor + '">' + message.text + '</span> ';
      this.forceUpdate();

      this.forceMessagesContainerScroll();
    }
  }

  protected forceMessagesContainerScroll() {
    var messagesRowClass = '.' + styles.messagesRow;
    var messagesDivElement = document.querySelector(messagesRowClass);
    messagesDivElement.scrollTop = messagesDivElement.scrollHeight;
  }

}
