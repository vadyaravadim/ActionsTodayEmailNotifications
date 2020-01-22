# ActionsTodayEmailNotifications

The project links to the Azure Cloud Service and Microsoft Dynamics 365. This app notifies the user of current activities scheduled for today by Email. 

The app starts at 10: 00 and 17: 00, is logged in to dynamics 365, goes through the entire active user base in CRM, searches for active events, and compiles an Email message in the format: 
Subject: subject, action type: action type, start date: start date, action link: action link
  
Messages are sent through the standard features of the CRM system.
This template for example send email to dynamics 365
