const { CloudAdapter, ConfigurationBotFrameworkAuthentication, TeamsInfo } = require('botbuilder');
const restify = require('restify');
const dotenv = require('dotenv');

// Load environment variables from .env file
dotenv.config();

console.log('=== BOT STARTING ===');
console.log('App ID:', process.env.MicrosoftAppId);
console.log('Environment variables loaded:', {
    PORT: process.env.PORT,
    MicrosoftAppId: process.env.MicrosoftAppId ? 'Set' : 'Not set',
    MicrosoftAppPassword: process.env.MicrosoftAppPassword ? 'Set' : 'Not set'
});

// Create HTTP server
const server = restify.createServer({
    name: 'teams-bot'
});

// Add body parser
server.use(restify.plugins.bodyParser());

// Create the auth configuration
const botFrameworkAuth = new ConfigurationBotFrameworkAuthentication({
    MicrosoftAppId: process.env.MicrosoftAppId,
    MicrosoftAppPassword: process.env.MicrosoftAppPassword,
    MicrosoftAppType: "MultiTenant"
});

// Create adapter with the auth configuration
const adapter = new CloudAdapter(botFrameworkAuth);

// Error handler
adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError] unhandled error: ${error}`);
    await context.sendActivity('Oops! Something went wrong. Please try again later.');
};

// Main bot logic
async function handleBotLogic(context) {
    console.log('\n=== BOT LOGIC START ===');
    
    if (context.activity.type === 'message') {
        const text = context.activity.text.replace(/<at>.*<\/at>/g, '').toLowerCase().trim();
        console.log('Cleaned message:', text);
        
        if (text === 'standup') {
            console.log('Standup command received');
            
            // For personal chat, just use the sender
            if (context.activity.conversation.conversationType === 'personal') {
                console.log('Personal chat detected, using sender as participant');
                const participant = {
                    id: context.activity.from.id,
                    name: context.activity.from.name
                };
                
                let message = "🎲 Today's standup order:\n\n1. " + participant.name;
                console.log('Sending message:', message);
                await context.sendActivity(message);
                return;
            }

            try {
                // For group chats and meetings
                console.log('Getting conversation members...');
                const members = await TeamsInfo.getPagedMembers(context);
                console.log('Got members:', members);
                
                if (!members || members.members.length === 0) {
                    await context.sendActivity("No participants found.");
                    return;
                }

                const shuffledParticipants = shuffleArray(members.members);
                let message = "🎲 Today's standup order:\n\n";
                shuffledParticipants.forEach((participant, index) => {
                    message += `${index + 1}. ${participant.name}\n`;
                });
                
                console.log('Sending message:', message);
                await context.sendActivity(message);
            } catch (error) {
                console.error('Error getting members:', error);
                await context.sendActivity("Sorry, I couldn't get the participant list. Please try again.");
            }
        }
    }
    console.log('=== BOT LOGIC END ===');
}

// Handle incoming requests using restify's preferred middleware pattern
server.post('/api/messages', async function(req, res) { // Make it async but only two params
    console.log('Received message:', req.body);
    try {
        await adapter.process(req, res, async (context) => {
            await handleBotLogic(context);
        });
    } catch (err) {
        console.error('Error processing request:', err);
        if (!res.headersSent) {
            res.status(500);
            res.end();
        }
    }
});

// Fisher-Yates shuffle algorithm
function shuffleArray(array) {
    const shuffled = [...array];
    for (let i = shuffled.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
    }
    return shuffled;
}

// Start server
server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n=== SERVER STARTED ===`);
    console.log(`${server.name} listening to ${server.url}`);
    console.log('Bot is ready to receive messages');
});