# AutoCoach

## Intro
AutoCoach is a Client Relationship Manager(CRM) developed in Google Apps Script, integrating Google Sheets and Gmail. It is a Google Sheets Add-On that operates on a trigger system to add clients and send them templated emails based on incoming emails of a certain format.

This tool brings the power of a personalized autoresponder to a organized spreadsheet, making it very comfortable to use, unlike other CRM systems that require weeks to get used to.

## How It Works
The tool requires little setup before it can start running. There are 3 Steps.

### Step 1: Creating Your Trigger Labels
Before you can start using the tool, you need to do some prep. First of all, you should be recieving automated emails of some sort containing client data. The tool will pick up these fields: Name, Email, Password, Phone. These fields can be used later to fill in variable tags in your templated emails. In order to make this type of email visible as a trigger however, you must create a Gmail filter that will apply a label to all emails with the same subject. The label you apply is what will be your trigger. This can be done in the Gmail settings.

### Step 2: Creating Your Canned Emails
The next step is to create your canned emails that you will be sending out. A canned email does not necesarrily have to be created using the canned emails feature in Gmail, it just has to be an email that is saved as a draft. In your email you will be able to place variable tags that will be filled in with the client's information before it is sent. acceptable tags are:

```Text
${firstname}
${lastname}
${email}
${password}
${phone}
```
Using these tags you will be able to create the appearance of a personal response and provide the client with any information they might need.

### Step 3: Configuring Your Utility
Once you have completed the preperation, you can now begin to actually set up your system. To add the sheets add on to your sheet, go to the add ons tab and saerch for a new add on. [Keep in mind this can only be added from domains on which the add on was published.] Search for AutoCoach and add it to your sheet. Once youve added it, an addon menu should appear.
First select the Reset option, and follow the prompts in order to intialize the properties (only done once after installation). After this is done, you can go to the Automated Email settings to begin customizing your email sequence.
