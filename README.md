# TelegramExpensesPyBot-Heroku-
This is alter version of another bot from https://github.com/stanislaw0910/TelegramExpensesBot.
I've made some changes in code to work with Heroku:

1. There is no .env file. All environment variables are storing in Heroku config

2. Gunicorn is used as WSGI HTTP Server

3. Routes rewritten for using webhooks with Heroku



