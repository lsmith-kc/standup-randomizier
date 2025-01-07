# Teams Standup Bot Testing Guide

## Prerequisites

- Node.js installed via Homebrew: `brew install node`
- Git for version control
- Access to Microsoft Teams
- Ability to upload custom apps to Teams

## Initial Setup

### 1. Clone and Install Dependencies

```bash
# Clone the repository
git clone https://github.com/lsmith-kc/standup-randomizier.git
cd standup-randomizer

# Install bot dependencies
cd bot
npm install
```

### 2. Install LocalTunnel

```bash
# Install localtunnel globally
npm install -g localtunnel
```

## Configure the Bot

### 1. Create Bot Registration

1. Go to <https://dev.teams.microsoft.com/bots>
2. Click "New Bot" or use existing bot
3. Save the Bot ID and password

### 2. Set Up Environment Variables

Create a `.env` file in the `bot` directory:

```text
MicrosoftAppId=your-bot-id-here
MicrosoftAppPassword=your-bot-password-here
PORT=3978
```

Note: add `DEBUG=true` to enable verbose logging.

## Running the Bot Locally

### 1. Start the Bot

```bash
# Make sure you're in the bot directory
cd bot
node index.js
```

You should see: `[server-name] listening to [url]`

### 2. Expose Local Server

In a new terminal:

```bash
lt --port 3978
```

Save the URL provided by LocalTunnel (e.g., <https://something.loca.lt>)

## Update Teams App Package

### 1. Update Manifest

Edit `appPackage/manifest.json`:

1. Replace `{{BOT_ID}}` with your Bot ID (in two places)
2. Replace `{{DOMAIN}}` with your LocalTunnel domain (without https://)
3. Update any other placeholder values

### 2. Create App Package

1. Ensure you have the required files in `appPackage/`:
   - manifest.json
   - color.png (192x192)
   - outline.png (32x32)

2. Zip these three files together:

```bash
cd appPackage
zip -r ../bot.zip *
```

## Upload to Teams

### 1. Upload Custom App

1. Open Teams
2. Click "Apps" in the sidebar
3. Click "Upload a custom app"
4. Select your `bot.zip` file

### 2. Test the Bot

1. Start a Teams meeting
2. Add the bot to the meeting
3. Type `!standup` in the chat
4. Bot should respond with randomized participant order

## Troubleshooting

### Bot Not Responding

1. Check if bot is running locally (`node index.js`)
2. Verify LocalTunnel is running and URL is accessible
3. Confirm Bot ID and password in `.env` are correct
4. Check Teams app manifest has correct bot ID and domain

### LocalTunnel Issues

1. If connection fails, try restarting LocalTunnel
2. If URL changes, update manifest and repackage app
3. Consider deploying to proper hosting for stability

### Teams Upload Issues

1. Verify manifest.json format is correct
2. Ensure icons are correct size and format
3. Check if you have permissions to upload custom apps

## Redeployment

When making changes:

1. Stop the bot (`Ctrl+C`)
2. Make your changes
3. Restart bot (`node index.js`)
4. If LocalTunnel URL changed:
   - Update manifest
   - Repackage app
   - Upload new package to Teams

## Production Deployment

For stable production use:

1. Deploy bot to proper hosting (Azure, Heroku, etc.)
2. Update manifest with permanent domain
3. Package and distribute app through proper Teams channels
