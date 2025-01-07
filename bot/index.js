const { CloudAdapter, ConfigurationBotFrameworkAuthentication, TeamsInfo } = require('botbuilder');
const restify = require('restify');
const dotenv = require('dotenv');

// Load environment variables
dotenv.config();

// Logger utility to standardize log format and levels
const logger = {
    info: (message, data) => {
        console.log(`[INFO] ${message}`, data ? data : '');
    },
    error: (message, error) => {
        console.error(`[ERROR] ${message}`, error);
    },
    debug: (message, data) => {
        if (process.env.DEBUG) {
            console.log(`[DEBUG] ${message}`, data ? data : '');
        }
    }
};

// Bot configuration
const BOT_CONFIG = {
    name: 'Standup Randomizer',
    port: process.env.PORT || 3978,
    auth: {
        MicrosoftAppId: process.env.MicrosoftAppId,
        MicrosoftAppPassword: process.env.MicrosoftAppPassword,
        MicrosoftAppType: "MultiTenant"
    }
};

// Initialize bot services
const server = restify.createServer({ name: BOT_CONFIG.name });
server.use(restify.plugins.bodyParser());

const botFrameworkAuth = new ConfigurationBotFrameworkAuthentication(BOT_CONFIG.auth);
const adapter = new CloudAdapter(botFrameworkAuth);

// Error handler
adapter.onTurnError = async (context, error) => {
    logger.error('Bot error:', error);
    await context.sendActivity('Sorry, something went wrong. Please try again.');
};

class StandupBot {
    async getParticipants(context) {
        if (context.activity.conversation.conversationType === 'personal') {
            return [{
                id: context.activity.from.id,
                name: context.activity.from.name
            }];
        }
        
        const members = await TeamsInfo.getPagedMembers(context);
        return members.members;
    }

    shuffleArray(array) {
        const shuffled = [...array];
        for (let i = shuffled.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
        }
        return shuffled;
    }

    formatStandupOrder(participants) {
        return "🎲 Today's standup order:\n\n" + 
            participants.map((p, i) => `${i + 1}. ${p.name}`).join('\n');
    }

    async handleCommand(context) {
        try {
            const participants = await this.getParticipants(context);
            
            if (!participants || participants.length === 0) {
                await context.sendActivity("No participants found.");
                return;
            }

            const shuffledParticipants = this.shuffleArray(participants);
            const message = this.formatStandupOrder(shuffledParticipants);
            
            logger.debug('Sending standup order', { participantCount: participants.length });
            await context.sendActivity(message);
        } catch (error) {
            logger.error('Error handling standup command:', error);
            await context.sendActivity("Sorry, I couldn't get the participant list. Please try again.");
        }
    }

    async handleMessage(context) {
        const text = context.activity.text.replace(/<at>.*<\/at>/g, '').toLowerCase().trim();
        
        if (text === 'standup') {
            logger.info('Standup command received', {
                conversationType: context.activity.conversation.conversationType,
                userId: context.activity.from.id
            });
            await this.handleCommand(context);
        }
    }
}

const bot = new StandupBot();

// Message handler
server.post('/api/messages', async (req, res) => {
    try {
        await adapter.process(req, res, async (context) => {
            if (context.activity.type === 'message') {
                await bot.handleMessage(context);
            }
        });
    } catch (err) {
        logger.error('Error processing request:', err);
        if (!res.headersSent) {
            res.status(500).end();
        }
    }
});

// Start server
server.listen(BOT_CONFIG.port, () => {
    logger.info(`${BOT_CONFIG.name} listening on port ${BOT_CONFIG.port}`);
});