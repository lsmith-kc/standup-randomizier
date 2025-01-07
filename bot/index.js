const { BotFrameworkAdapter, TeamsInfo } = require('botbuilder');
const restify = require('restify');
const dotenv = require('dotenv');

// Load environment variables from .env file
dotenv.config();

// Create adapter
const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

// Error handler
adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError] unhandled error: ${error}`);
    await context.sendActivity('Oops! Something went wrong. Please try again later.');
};

// Create HTTP server
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${server.name} listening to ${server.url}`);
});

// Listen for incoming requests
server.post('/api/messages', async (req, res) => {
    await adapter.processActivity(req, res, async (context) => {
        if (context.activity.type === 'message') {
            const text = context.activity.text.toLowerCase().trim();
            
            if (text === '!standup') {
                try {
                    // Get meeting participants
                    const meeting = await TeamsInfo.getMeetingParticipants(context);
                    
                    if (!meeting || meeting.length === 0) {
                        await context.sendActivity("No participants found in the meeting.");
                        return;
                    }

                    // Shuffle participants
                    const shuffledParticipants = shuffleArray(meeting);
                    
                    // Create presenter order message
                    let message = "🎲 Today's standup order:\n\n";
                    shuffledParticipants.forEach((participant, index) => {
                        message += `${index + 1}. ${participant.user.name}\n`;
                    });
                    
                    await context.sendActivity(message);
                } catch (error) {
                    console.error('Error getting meeting participants:', error);
                    await context.sendActivity("Sorry, I couldn't get the meeting participants. Make sure I'm added to the meeting.");
                }
            }
        }
    });
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