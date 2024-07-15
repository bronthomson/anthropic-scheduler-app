# Creating Your Own Anthropic API Scheduler App on Heroku

This guide will walk you through creating your own application that uses the Anthropic API to generate responses and schedules tasks using the Heroku Scheduler add-on.

## Prerequisites

- [Heroku account](https://www.heroku.com/) Login with Heroku credentials from Bitwarden
- [Heroku CLI installed](https://devcenter.heroku.com/articles/heroku-cli)
- [Git installed](https://git-scm.com/book/en/v2/Getting-Started-Installing-Git)
- [Python 3.7+ installed](https://www.python.org/downloads/)
- [Anthropic API key](https://www.anthropic.com/)

## Steps

### 1. Install Heroku CLI

If you haven't already, install the Heroku CLI:

```sh
# On macOS
brew tap heroku/brew && brew install heroku

# On Windows
choco install heroku-cli

# On Linux (Ubuntu)
curl https://cli-assets.heroku.com/install-ubuntu.sh | sh
```

### 2.  Login with Heroku credentials from Bitwarden

### 3.  Create HEROKU project
We have an example in that repo


```
mkdir anthropic-scheduler-app
cd anthropic-scheduler-app
git init
```

###  4. Create the Application Script
We have an example in that repo

### 5. Create a requirements.txt File
Create a requirements.txt file with the following content to specify the project dependencies:
`anthropic`

### 6. Create a Procfile
Create a Procfile to tell Heroku how to run your app:
```
web: python main.py
```

### 7. Create and Deploy to Heroku
7.1. Log in to Heroku
Log in to your Heroku account:

```
heroku login
```
Then proceed with a link, you should be loged in the browser already

7.2. Create a Heroku App
Create a new Heroku app:

```
heroku create
```

7.3. Set Your Anthropic API Key
Replace your_api_key with your actual API key:

```
heroku config:set ANTHROPIC_API_KEY=your_api_key
```

7.4. Add, Commit, and Push Your Code to Heroku
Add your files to Git, commit, and push to Heroku:

```
git add .
git commit -m "Initial commit"
git push heroku main
```
### 8. Add the Heroku Scheduler
Add the Heroku Scheduler add-on to your app:

```
heroku addons:create scheduler:standard
```

### 9. Configure the Scheduler:
In the Heroku UI, you will find "Configure Add-ons" in the Overview section:

Select Heroku Scheduler.
Press the "Add job" button.
Specify the schedule and run command, for example: python main.py.

### 10. Verify
Check your Heroku app logs to verify that the scheduled tasks are running correctly:

```
heroku logs --tail
```

