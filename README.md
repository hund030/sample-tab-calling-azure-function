# Tabs

Tabs are Teams-aware webpages embedded in Microsoft Teams. Personal tabs are scoped to a single user. They can be pinned to the left navigation bar for easy access.

## Prerequisites
-  [NodeJS](https://nodejs.org/en/)

-  [M365 developer account](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/prepare-your-o365-tenant) or access to a Teams account with the appropriate permissions to install an app.

## Build and Run

In the project directory, execute:

`npm install`

`npm start`

## Deploy to Teams

1) Create a teams app project with VS Code Teams extension, and copy the ./publish/Development.env to this repository.

1) Deploy this project to your new created frotend environment with VS Code Teams extension.

1) Upload the `Development.zip` from the *.publish* folder to Teams.

1) Configure CORS for Azure Function. Learn more [here](https://www.c-sharpcorner.com/article/handling-cors-in-azure-function/)
