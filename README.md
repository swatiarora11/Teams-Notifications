# Teams-Notification Overview and Deployment Guide

**Teams-Notification** solution is for communication with members of one or multiple O365 groups within your company tenant. This solution enables you to broadcast a notification and also bring user attention to a particular notification by sending reminders. This solution enables you to provide richer experience to users and better engage with them by keeping them up to date with your company communication and various changes regarding apps and workflows used by them. 

Use this solution for scenarios like updates about townhalls, announcements, sending content updates of Microsoft 365 Learning Pathways, such as new trainings, URLs to learn more about Microsoft Outlook, SharePoint and Teams.

## Key Features

1. **Custom Notification:** Send company wide communication through 1:1 custom messages/ adaptive cards for a more engaging experience with your users.
1. **Reminder for Notification:** Send reminders for earlier notifications at desired frequency i.e. daily, weekly, monthly through Microsoft Teams Activity Feed.
1. **Target Audience:** Choose your target audience i.e. members of one or multiple O365 groups/ teams.

## Deployment Guide
### Prerequisites 
To begin, you will need:
* An Azure account that has an active subscription. [Create an account for free.](https://azure.microsoft.com/free/?WT.mc_id=A261C142F)
* The Azure account must have permission to manage applications in Azure Active Directory (Azure AD). Any of the following Azure AD roles include the required permissions:
    * Application administrator
    * Application developer
    * Cloud application administrator
* An Office 365 account that has an active subscription of Exchange Online, Sharepoint Online and Microsoft Teams
* Office 365 account(s) with administrative rights for the following workloads -
    * Exchange Online
    * Sharepoint Online
    * Microsoft Teams
* Teams-Notification zip package. Link to [Teams-Notification package](https://github.com/swatiarora11/QuizApp/blob/master/Deployment/QuizApp.zip)

### Step 1: Register Azure AD Application
Registering your application establishes a trust relationship between your app and the Microsoft identity platform. The trust is unidirectional: your app trusts the Microsoft identity platform, and not the other way around.

Follow these steps to create the app registration:
1. Sign in to the [Azure portal.](https://portal.azure.com/)
2. Search for and select Azure Active Directory.
3. Under **Manage**, select **App registrations** > **New registration**.
    * Enter a display **Name** for your application for e.g. **TeamsNotifications**. 
    * Specify who can use the application. Select "Accounts in this organizational directory only" option for **Supported account types**.
    * Don't enter anything for **Redirect URI (optional)**.
    * Select **Register** to complete the initial app registration.
<p> <img src="images/aad_app_register.png" />

4. When registration finishes, the Azure portal displays the app registration's **Overview** pane. You see the **Application (client) ID**. Also called the **client ID**, this value uniquely identifies your application in the Microsoft identity platform. Verify that the **Supported account types** is set to "Single organization". Copy the **Application (client) ID**; we will need it later. 

<p> <img src="images/aad_app_overview.png" />

5. On the side rail in the **Manage** section, navigate to the **Certificates & secrets** section. In the Client secrets section, click on **+ New client secret**. Add a description for the secret, and choose when the secret will expire. Click **Add**.

<p> <img src="images/aad_app_secret.png" />

6. Once the **Client secret** is created, copy its value; At this point you should have values of **Client secret**, **Directory (tenant) ID**, **Application (client) Id**. We will need these values later. 
7. Under **API Permissions** > **Add a permission** > **Select an API** >  **Commonly used Microsoft APIs**, select “Microsoft Graph” and give following permissions,
    * Under **Delegated permissions**,
        * Chat.Read
        * Chat.ReadBasic
        * Chat.ReadWrite
        * TeamsActivity.Read
        * TeamsActivity.Send
        * User.Read
        * UserActivity.ReadWrite.Created
    * Under **Application Permissions**,
        * Directory.Read.All
        * Directory.ReadWrite.All
        * TeamsActivity.Read.All
        * TeamsAcitvity.Send
        * User.Read.All
        * User.ReadWrite.All
    * Click on **Add Permissions** to save your changes.

<p> <img src="images/aad_app_api_perm1.png" />
<p> <img src="images/aad_app_api_perm2.png" />

If you are logged in as Tenant Administrator, click on “Grant admin consent”, else inform your Tenant Administrator to do the same.

### Step 2: Setup SharePoint Online

### Step 3: Setup Microsoft Teams

### Step 4: Setup Powershell

### Step 5: Test Your First Notification

### Step 6: Test Your First Reminder
