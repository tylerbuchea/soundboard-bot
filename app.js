const {
  BotFrameworkAdapter,
  MemoryStorage,
  ConversationState,
  Message,
} = require('botbuilder');
const restify = require('restify');

// Create server
let server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
  console.log(`${server.name} listening to ${server.url}`);
});

// Create adapter
const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD
});

// Add conversation state middleware
const conversationState = new ConversationState(new MemoryStorage());
adapter.use(conversationState);

// Listen for incoming requests 
server.post('/api/messages', (req, res) => {
  // Route received request to adapter for processing
  adapter.processActivity(req, res, (context) => {
    if (context.activity.type === 'message') {
      const state = conversationState.get(context);
      const count = state.count === undefined ? state.count = 0 : ++state.count;
      const msg = new Message(context).addAttachment({
        contentUrl: 'http://www.pachd.com/sfx/traffic-8.mp3',
        contentType: 'audio/mpeg',
        name: 'My video clip'
      });
      // return context.sendActivity(`${count}: You said "${context.activity.text}"`);
      return context.send(msg);
    } else {
      return context.sendActivity(`[${context.activity.type} event detected]`);
    }
  });
});